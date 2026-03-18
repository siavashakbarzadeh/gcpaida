"""
CGP (Cartesian Genetic Programming) - City Relationship Discovery (2021)
=========================================================================
Uses CGP to discover which Italian cities have direct influence on
each other's COVID-19 infection rates. For each target city, CGP
evolves mathematical expressions that predict its daily infections
from other cities' data. Active inputs in the evolved program
reveal which cities are truly linked.

Output: 2021/cgp_results/ folder with network graphs, heatmaps, and summaries.
"""

import os
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from collections import defaultdict
import warnings
warnings.filterwarnings('ignore')

# ==============================================================================
# CONFIGURATION
# ==============================================================================

DATA_PATH = r"c:\Users\Utente\Documents\GitHub\gcpaida\2021\merged_cities_2021.csv"
OUTPUT_DIR = r"c:\Users\Utente\Documents\GitHub\gcpaida\2021\cgp_results"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# CGP Parameters
N_ROWS = 3
N_COLS = 8
N_OUTPUTS = 1
N_GENERATIONS = 300
LAMBDA = 4           # children per generation
MUTATION_RATE = 0.08
TOP_K_CANDIDATES = 15  # top correlated cities as CGP inputs per target
N_LAGS = 3             # use lags 0, 1, 2 days

# Connection threshold: if CGP prediction R^2 > this, connection is significant
R2_THRESHOLD = 0.3

# ==============================================================================
# CGP IMPLEMENTATION
# ==============================================================================

class CartesianGeneticProgramming:
    """
    Cartesian Genetic Programming (CGP) implementation.

    Nodes are arranged in a grid (rows x cols). Each node has:
      - A function gene (which operation to perform)
      - Two input genes (connections to previous nodes or inputs)

    Evolution uses (1+lambda) strategy with point mutation.
    """

    # Available mathematical functions
    FUNCTIONS = [
        ('add',  lambda a, b: a + b),
        ('sub',  lambda a, b: a - b),
        ('mul',  lambda a, b: a * b),
        ('div',  lambda a, b: np.where(np.abs(b) > 1e-6, a / b, 0.0)),
        ('max',  lambda a, b: np.maximum(a, b)),
        ('min',  lambda a, b: np.minimum(a, b)),
        ('abs_diff', lambda a, b: np.abs(a - b)),
        ('avg',  lambda a, b: (a + b) / 2.0),
    ]

    def __init__(self, n_inputs, n_outputs=1, n_rows=3, n_cols=8):
        self.n_inputs = n_inputs
        self.n_outputs = n_outputs
        self.n_rows = n_rows
        self.n_cols = n_cols
        self.n_functions = len(self.FUNCTIONS)
        self.n_nodes = n_rows * n_cols
        self.genome_length = self.n_nodes * 3 + self.n_outputs

    def random_genome(self):
        genome = np.zeros(self.genome_length, dtype=int)
        for node_idx in range(self.n_nodes):
            col = node_idx // self.n_rows
            gene_start = node_idx * 3
            genome[gene_start] = np.random.randint(self.n_functions)
            max_conn = self.n_inputs + col * self.n_rows
            genome[gene_start + 1] = np.random.randint(max(1, max_conn))
            genome[gene_start + 2] = np.random.randint(max(1, max_conn))
        out_start = self.n_nodes * 3
        for i in range(self.n_outputs):
            genome[out_start + i] = np.random.randint(self.n_inputs + self.n_nodes)
        return genome

    def get_active_nodes(self, genome):
        """Trace active nodes from outputs backwards."""
        active = set()
        out_start = self.n_nodes * 3
        queue = []
        for i in range(self.n_outputs):
            idx = genome[out_start + i]
            if idx >= self.n_inputs:
                node = idx - self.n_inputs
                active.add(node)
                queue.append(node)
        while queue:
            node_idx = queue.pop()
            gs = node_idx * 3
            for inp in [genome[gs + 1], genome[gs + 2]]:
                if inp >= self.n_inputs and (inp - self.n_inputs) not in active:
                    active.add(inp - self.n_inputs)
                    queue.append(inp - self.n_inputs)
        return active

    def get_active_inputs(self, genome):
        """Get which original inputs are used in the active computation graph."""
        active_nodes = self.get_active_nodes(genome)
        used_inputs = set()
        out_start = self.n_nodes * 3
        for i in range(self.n_outputs):
            idx = genome[out_start + i]
            if idx < self.n_inputs:
                used_inputs.add(idx)
        for node_idx in active_nodes:
            gs = node_idx * 3
            for inp in [genome[gs + 1], genome[gs + 2]]:
                if inp < self.n_inputs:
                    used_inputs.add(inp)
        return used_inputs

    def evaluate(self, genome, X):
        """Evaluate the CGP program on input data X (n_samples x n_inputs)."""
        n_samples = X.shape[0]
        values = np.zeros((self.n_inputs + self.n_nodes, n_samples))
        for i in range(self.n_inputs):
            values[i] = X[:, i]
        for node_idx in range(self.n_nodes):
            gs = node_idx * 3
            func_id = genome[gs] % self.n_functions
            in1 = genome[gs + 1]
            in2 = genome[gs + 2]
            _, func = self.FUNCTIONS[func_id]
            result = func(values[in1], values[in2])
            result = np.clip(result, -1e8, 1e8)
            result = np.nan_to_num(result, nan=0.0, posinf=1e8, neginf=-1e8)
            values[self.n_inputs + node_idx] = result
        out_start = self.n_nodes * 3
        outputs = np.zeros((n_samples, self.n_outputs))
        for i in range(self.n_outputs):
            outputs[:, i] = values[genome[out_start + i]]
        return outputs

    def fitness(self, genome, X, y):
        """Compute MSE fitness (lower is better)."""
        pred = self.evaluate(genome, X)[:, 0]
        return np.mean((pred - y) ** 2)

    def mutate(self, genome, rate=0.08):
        new = genome.copy()
        out_start = self.n_nodes * 3
        for i in range(len(new)):
            if np.random.random() < rate:
                if i >= out_start:
                    new[i] = np.random.randint(self.n_inputs + self.n_nodes)
                else:
                    node_idx = i // 3
                    gene_type = i % 3
                    col = node_idx // self.n_rows
                    if gene_type == 0:
                        new[i] = np.random.randint(self.n_functions)
                    else:
                        max_conn = self.n_inputs + col * self.n_rows
                        new[i] = np.random.randint(max(1, max_conn))
        return new

    def evolve(self, X, y, n_generations=300, lam=4, mutation_rate=0.08):
        """Evolve using (1+lambda) strategy."""
        parent = self.random_genome()
        parent_fit = self.fitness(parent, X, y)
        history = [parent_fit]

        for gen in range(n_generations):
            children = [self.mutate(parent, mutation_rate) for _ in range(lam)]
            fits = [self.fitness(c, X, y) for c in children]
            best_idx = np.argmin(fits)
            if fits[best_idx] <= parent_fit:
                parent = children[best_idx]
                parent_fit = fits[best_idx]
            history.append(parent_fit)

        return parent, parent_fit, history

    def compute_r2(self, genome, X, y):
        """Compute R-squared score."""
        pred = self.evaluate(genome, X)[:, 0]
        ss_res = np.sum((y - pred) ** 2)
        ss_tot = np.sum((y - np.mean(y)) ** 2)
        if ss_tot < 1e-10:
            return 0.0
        return 1.0 - ss_res / ss_tot


# ==============================================================================
# DATA PREPARATION
# ==============================================================================

def prepare_data(filepath):
    """Load data and compute daily new infections for each city."""
    df = pd.read_csv(filepath)
    df['data'] = pd.to_datetime(df['data'])
    df.sort_values('data', inplace=True)
    df.set_index('data', inplace=True)

    # 2021 format: infected_CityName
    city_cols = [c for c in df.columns if c.startswith('infected_')]
    city_names = [c.replace('infected_', '').lower().replace(' ', '_') for c in city_cols]

    # Compute daily new infections
    daily_new = pd.DataFrame(index=df.index)
    for col, city in zip(city_cols, city_names):
        cumulative = df[col].values.astype(float)
        cumulative = np.nan_to_num(cumulative, nan=0)
        diff = np.diff(cumulative, prepend=cumulative[0])
        daily_new[city] = np.maximum(diff, 0)

    return daily_new, city_names


def get_correlation_matrix(daily_new, city_names):
    """Compute correlation matrix between all cities."""
    return daily_new[city_names].corr()


def build_lagged_inputs(daily_new, target_city, candidate_cities, n_lags=3):
    """
    Build input matrix with lagged values from candidate cities.
    Returns X (n_samples x n_features), y (n_samples), feature_names.
    """
    features = []
    feature_names = []

    for city in candidate_cities:
        for lag in range(n_lags):
            col = daily_new[city].shift(lag).values
            features.append(col)
            feature_names.append(f"{city}_lag{lag}")

    X = np.column_stack(features)
    y = daily_new[target_city].values

    # Remove NaN rows (from lagging)
    valid = ~np.isnan(X).any(axis=1) & ~np.isnan(y)
    X = X[valid]
    y = y[valid]

    # Normalize
    X_mean = np.mean(X, axis=0, keepdims=True)
    X_std = np.std(X, axis=0, keepdims=True) + 1e-10
    X_norm = (X - X_mean) / X_std

    y_mean = np.mean(y)
    y_std = np.std(y) + 1e-10
    y_norm = (y - y_mean) / y_std

    return X_norm, y_norm, feature_names, valid


# ==============================================================================
# VISUALIZATION
# ==============================================================================

def plot_network_graph(connections, city_names, output_dir, top_n=50):
    """Plot network graph of city connections using spring layout."""
    try:
        import networkx as nx
        has_nx = True
    except ImportError:
        has_nx = False

    if not has_nx:
        print("  networkx not available, skipping network graph")
        return

    # Build graph
    G = nx.Graph()
    for city in city_names:
        G.add_node(city)

    # Add edges with weights
    edge_data = []
    for (c1, c2), weight in connections.items():
        if weight > R2_THRESHOLD:
            edge_data.append((c1, c2, weight))

    # Sort and take top_n connections
    edge_data.sort(key=lambda x: x[2], reverse=True)
    edge_data = edge_data[:top_n * 3]

    for c1, c2, w in edge_data:
        G.add_edge(c1, c2, weight=w)

    # Remove isolated nodes for cleaner visualization
    isolated = list(nx.isolates(G))
    G.remove_nodes_from(isolated)

    if len(G.edges()) == 0:
        print("  No significant connections found for network graph")
        return

    # Layout
    pos = nx.spring_layout(G, k=2, iterations=50, seed=42)

    # Node properties
    degrees = dict(G.degree())
    node_sizes = [max(degrees.get(n, 0) * 150, 100) for n in G.nodes()]
    node_colors = [degrees.get(n, 0) for n in G.nodes()]

    # Edge properties
    edge_weights = [G[u][v]['weight'] for u, v in G.edges()]
    max_w = max(edge_weights) if edge_weights else 1
    edge_widths = [w / max_w * 4 for w in edge_weights]
    edge_colors = edge_weights

    fig, ax = plt.subplots(figsize=(24, 20))

    # Draw
    nx.draw_networkx_edges(G, pos, ax=ax, edge_color=edge_colors,
                           edge_cmap=plt.cm.Reds, width=edge_widths, alpha=0.6)
    nodes = nx.draw_networkx_nodes(G, pos, ax=ax, node_size=node_sizes,
                                    node_color=node_colors, cmap=plt.cm.YlOrRd,
                                    edgecolors='black', linewidths=0.5)
    nx.draw_networkx_labels(G, pos, ax=ax, font_size=6, font_weight='bold')

    plt.colorbar(nodes, ax=ax, label='Number of Connections', shrink=0.5)
    ax.set_title('CGP-Discovered City Infection Network (2021)\n'
                 '(Edges = cities whose infections influence each other)',
                 fontsize=16, fontweight='bold')
    ax.axis('off')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'network_graph.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: network_graph.png")

    # Also save degree centrality
    centrality = nx.degree_centrality(G)
    betweenness = nx.betweenness_centrality(G)
    return G, centrality, betweenness


def plot_connection_heatmap(connection_matrix, city_names, output_dir):
    """Full heatmap of CGP-discovered connection strengths."""
    fig, ax = plt.subplots(figsize=(30, 28))
    im = ax.imshow(connection_matrix, cmap='YlOrRd', vmin=0, vmax=1, aspect='auto')
    ax.set_xticks(range(len(city_names)))
    ax.set_yticks(range(len(city_names)))
    ax.set_xticklabels([c.title() for c in city_names], rotation=90, fontsize=4)
    ax.set_yticklabels([c.title() for c in city_names], fontsize=4)
    ax.set_title('CGP Connection Strength Heatmap (R² of prediction) - 2021', fontsize=14, fontweight='bold')
    plt.colorbar(im, ax=ax, label='Connection Strength (R²)', shrink=0.5)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'connection_heatmap.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: connection_heatmap.png")


def plot_top_connections(connections, output_dir, top_n=40):
    """Bar chart of top N strongest connections."""
    sorted_conns = sorted(connections.items(), key=lambda x: x[1], reverse=True)[:top_n]

    labels = [f"{c1.title()} <-> {c2.title()}" for (c1, c2), _ in sorted_conns]
    values = [v for _, v in sorted_conns]

    fig, ax = plt.subplots(figsize=(14, 12))
    colors = plt.cm.RdYlGn_r(np.linspace(0.2, 0.8, len(values)))
    ax.barh(range(len(labels)), values, color=colors)
    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels, fontsize=8)
    ax.set_xlabel('Connection Strength (R²)', fontsize=12)
    ax.set_title(f'Top {top_n} Strongest City Connections (CGP) - 2021', fontsize=14, fontweight='bold')
    ax.axvline(x=R2_THRESHOLD, color='red', ls='--', label=f'Threshold R²={R2_THRESHOLD}')
    ax.legend()
    ax.grid(True, alpha=0.3, axis='x')
    ax.invert_yaxis()
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'top_connections.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: top_connections.png")


def plot_hub_cities(city_connection_count, output_dir):
    """Bar chart of most connected cities (hubs)."""
    sorted_cities = sorted(city_connection_count.items(), key=lambda x: x[1], reverse=True)
    cities = [c.title() for c, _ in sorted_cities]
    counts = [v for _, v in sorted_cities]

    fig, ax = plt.subplots(figsize=(16, 20))
    colors = plt.cm.viridis(np.linspace(0.2, 0.8, len(cities)))
    ax.barh(range(len(cities)), counts, color=colors)
    ax.set_yticks(range(len(cities)))
    ax.set_yticklabels(cities, fontsize=6)
    ax.set_xlabel('Number of Direct Connections', fontsize=12)
    ax.set_title('City Hub Analysis: Number of CGP-Discovered Connections (2021)',
                 fontsize=14, fontweight='bold')
    ax.grid(True, alpha=0.3, axis='x')
    ax.invert_yaxis()
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'hub_cities.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: hub_cities.png")


def plot_cgp_fitness_examples(fitness_histories, output_dir, n_examples=6):
    """Plot CGP fitness convergence for a few example cities."""
    example_cities = list(fitness_histories.keys())[:n_examples]
    fig, axes = plt.subplots(2, 3, figsize=(18, 10))
    for idx, city in enumerate(example_cities):
        ax = axes[idx // 3][idx % 3]
        ax.plot(fitness_histories[city], 'b-', lw=1)
        ax.set_title(f'{city.title()}', fontsize=11)
        ax.set_xlabel('Generation')
        ax.set_ylabel('Fitness (MSE)')
        ax.grid(True, alpha=0.3)
        ax.set_yscale('log')
    fig.suptitle('CGP Evolution Fitness Convergence (2021)', fontsize=14, fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'cgp_fitness_convergence.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: cgp_fitness_convergence.png")


def plot_city_influence_detail(target, active_cities, r2_scores, output_dir):
    """Plot which cities influence a specific target city."""
    if not active_cities:
        return
    cities = [c.title() for c in active_cities]
    scores = [r2_scores.get(c, 0) for c in active_cities]

    fig, ax = plt.subplots(figsize=(10, max(4, len(cities) * 0.4)))
    colors = plt.cm.RdYlGn(np.array(scores))
    ax.barh(range(len(cities)), scores, color=colors)
    ax.set_yticks(range(len(cities)))
    ax.set_yticklabels(cities, fontsize=9)
    ax.set_xlabel('Influence Strength (R²)')
    ax.set_title(f'Cities Influencing {target.title()} (CGP) - 2021', fontsize=13, fontweight='bold')
    ax.grid(True, alpha=0.3, axis='x')
    ax.invert_yaxis()
    plt.tight_layout()
    safe = target.replace("'", "").replace(" ", "_")
    plt.savefig(os.path.join(output_dir, f'influence_{safe}.png'), dpi=120, bbox_inches='tight')
    plt.close()


def plot_grouped_connections(connections, city_names, output_dir):
    """Show connections grouped by geographic proximity."""
    # Group cities by significant connections
    significant = {k: v for k, v in connections.items() if v > R2_THRESHOLD}

    if not significant:
        print("  No significant connections to group")
        return

    # Count connections per city
    conn_count = defaultdict(int)
    conn_strength = defaultdict(float)
    for (c1, c2), v in significant.items():
        conn_count[c1] += 1
        conn_count[c2] += 1
        conn_strength[c1] += v
        conn_strength[c2] += v

    # Sort by connection count
    sorted_cities = sorted(conn_count.keys(), key=lambda c: conn_count[c], reverse=True)[:30]

    # Build sub-matrix
    n = len(sorted_cities)
    matrix = np.zeros((n, n))
    for i, c1 in enumerate(sorted_cities):
        for j, c2 in enumerate(sorted_cities):
            key1 = (c1, c2)
            key2 = (c2, c1)
            if key1 in connections:
                matrix[i, j] = connections[key1]
            elif key2 in connections:
                matrix[i, j] = connections[key2]

    fig, ax = plt.subplots(figsize=(14, 12))
    im = ax.imshow(matrix, cmap='YlOrRd', vmin=0, vmax=1)
    ax.set_xticks(range(n))
    ax.set_yticks(range(n))
    ax.set_xticklabels([c.title() for c in sorted_cities], rotation=45, ha='right', fontsize=7)
    ax.set_yticklabels([c.title() for c in sorted_cities], fontsize=7)

    # Add values
    for i in range(n):
        for j in range(n):
            if matrix[i, j] > 0.1:
                ax.text(j, i, f'{matrix[i, j]:.2f}', ha='center', va='center', fontsize=5)

    ax.set_title('Top 30 Most Connected Cities - Connection Matrix (CGP R²) - 2021',
                 fontsize=13, fontweight='bold')
    plt.colorbar(im, ax=ax, label='R²', shrink=0.7)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'top30_connection_matrix.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: top30_connection_matrix.png")


# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

def main():
    print("=" * 70)
    print("   CGP - Cartesian Genetic Programming (2021)")
    print("   City Relationship Discovery (107 Cities)")
    print("=" * 70)

    # 1. Load data
    print("\n[1/6] Loading data and computing daily infections...")
    daily_new, city_names = prepare_data(DATA_PATH)
    n_cities = len(city_names)
    print(f"  {n_cities} cities, {len(daily_new)} days")

    # 2. Compute correlations (for candidate selection)
    print("\n[2/6] Computing correlation matrix for candidate selection...")
    corr_matrix = get_correlation_matrix(daily_new, city_names)

    # 3. Run CGP for each city
    print(f"\n[3/6] Running CGP for each city (this takes a while)...")
    print(f"  CGP params: rows={N_ROWS}, cols={N_COLS}, generations={N_GENERATIONS}")
    print(f"  Strategy: (1+{LAMBDA}), mutation_rate={MUTATION_RATE}")
    print(f"  Candidate inputs per city: top {TOP_K_CANDIDATES}, lags: {N_LAGS}")
    print()

    # Store results
    all_connections = {}      # (city_a, city_b) -> R² strength
    city_active_inputs = {}   # city -> list of influencing cities
    fitness_histories = {}
    city_r2_details = {}      # city -> {input_city: R²}

    for idx, target in enumerate(city_names):
        # Select top-k correlated cities as candidates (exclude self)
        corrs = corr_matrix[target].drop(target, errors='ignore').abs()
        top_candidates = corrs.nlargest(TOP_K_CANDIDATES).index.tolist()

        # Build lagged input matrix
        X, y, feat_names, valid_mask = build_lagged_inputs(
            daily_new, target, top_candidates, n_lags=N_LAGS
        )

        if len(y) < 30 or np.std(y) < 1e-10:
            print(f"  [{idx+1:3d}/{n_cities}] {target.title():30s} - SKIPPED (insufficient data)")
            continue

        # Initialize and evolve CGP
        n_inputs = X.shape[1]
        cgp = CartesianGeneticProgramming(n_inputs, N_OUTPUTS, N_ROWS, N_COLS)
        best_genome, best_fitness, history = cgp.evolve(
            X, y, N_GENERATIONS, LAMBDA, MUTATION_RATE
        )

        # Analyze which inputs are active
        active_input_indices = cgp.get_active_inputs(best_genome)
        r2 = cgp.compute_r2(best_genome, X, y)

        # Map input indices back to city names
        active_cities = set()
        for inp_idx in active_input_indices:
            if inp_idx < len(feat_names):
                city_from_feat = feat_names[inp_idx].rsplit('_lag', 1)[0]
                active_cities.add(city_from_feat)

        city_active_inputs[target] = list(active_cities)
        fitness_histories[target] = history
        city_r2_details[target] = {}

        # Store connections with R² weight
        for source_city in active_cities:
            pair = tuple(sorted([target, source_city]))
            current = all_connections.get(pair, 0)
            all_connections[pair] = max(current, r2)
            city_r2_details[target][source_city] = r2

        n_active = len(active_cities)
        status = "LINKED" if n_active > 0 and r2 > R2_THRESHOLD else "weak"
        print(f"  [{idx+1:3d}/{n_cities}] {target.title():30s}  "
              f"R2={r2:.3f}  Active inputs: {n_active:2d}  [{status}]"
              f"  -> {', '.join([c.title() for c in list(active_cities)[:5]])}"
              f"{'...' if n_active > 5 else ''}")

    # 4. Generate visualizations
    print(f"\n[4/6] Generating visualizations...")

    # Connection heatmap
    conn_matrix = np.zeros((n_cities, n_cities))
    for i, c1 in enumerate(city_names):
        for j, c2 in enumerate(city_names):
            pair = tuple(sorted([c1, c2]))
            if pair in all_connections:
                conn_matrix[i, j] = all_connections[pair]
    plot_connection_heatmap(conn_matrix, city_names, OUTPUT_DIR)

    # Top connections bar chart
    plot_top_connections(all_connections, OUTPUT_DIR, top_n=40)

    # Hub cities
    city_conn_count = defaultdict(int)
    for (c1, c2), v in all_connections.items():
        if v > R2_THRESHOLD:
            city_conn_count[c1] += 1
            city_conn_count[c2] += 1
    # Include cities with 0 connections
    for c in city_names:
        if c not in city_conn_count:
            city_conn_count[c] = 0
    plot_hub_cities(city_conn_count, OUTPUT_DIR)

    # Fitness convergence examples
    plot_cgp_fitness_examples(fitness_histories, OUTPUT_DIR)

    # Top 30 connection matrix
    plot_grouped_connections(all_connections, city_names, OUTPUT_DIR)

    # Network graph
    print("\n[5/6] Building network graph...")
    graph_result = plot_network_graph(all_connections, city_names, OUTPUT_DIR)

    # Individual influence plots for top 10 most connected cities
    top_hubs = sorted(city_conn_count.items(), key=lambda x: x[1], reverse=True)[:10]
    for city, count in top_hubs:
        if city in city_active_inputs and city_active_inputs[city]:
            plot_city_influence_detail(
                city, city_active_inputs[city],
                city_r2_details.get(city, {}), OUTPUT_DIR
            )

    # 6. Save summary data
    print(f"\n[6/6] Saving summary data...")

    # Connections CSV
    conn_rows = []
    for (c1, c2), strength in sorted(all_connections.items(), key=lambda x: x[1], reverse=True):
        conn_rows.append({
            'city_1': c1.title(),
            'city_2': c2.title(),
            'connection_strength_R2': round(strength, 4),
            'is_significant': strength > R2_THRESHOLD,
        })
    conn_df = pd.DataFrame(conn_rows)
    conn_df.to_csv(os.path.join(OUTPUT_DIR, 'all_connections.csv'), index=False)
    print(f"  Saved: all_connections.csv ({len(conn_df)} pairs)")

    # City summary CSV
    summary_rows = []
    for city in city_names:
        n_connections = city_conn_count.get(city, 0)
        linked_cities = city_active_inputs.get(city, [])
        summary_rows.append({
            'city': city.title(),
            'n_connections': n_connections,
            'linked_cities': '; '.join([c.title() for c in linked_cities]),
            'is_hub': n_connections >= 5,
        })
    summary_df = pd.DataFrame(summary_rows)
    summary_df.sort_values('n_connections', ascending=False, inplace=True)
    summary_df.to_csv(os.path.join(OUTPUT_DIR, 'city_summary.csv'), index=False)
    print(f"  Saved: city_summary.csv")

    # Connection matrix CSV
    conn_mat_df = pd.DataFrame(conn_matrix, index=city_names, columns=city_names)
    conn_mat_df.to_csv(os.path.join(OUTPUT_DIR, 'connection_matrix.csv'))
    print(f"  Saved: connection_matrix.csv")

    # Print summary
    n_significant = sum(1 for v in all_connections.values() if v > R2_THRESHOLD)
    n_total = len(all_connections)

    print("\n" + "=" * 70)
    print("   CGP RESULTS SUMMARY (2021)")
    print("=" * 70)
    print(f"  Total city pairs analyzed: {n_total}")
    print(f"  Significant connections (R2>{R2_THRESHOLD}): {n_significant}")
    print(f"  Cities with connections: {sum(1 for v in city_conn_count.values() if v > 0)}/{n_cities}")

    print(f"\n  TOP 10 MOST CONNECTED CITIES (Hubs):")
    for city, count in top_hubs:
        print(f"    {city.title():25s}  {count} connections")

    print(f"\n  TOP 10 STRONGEST CONNECTIONS:")
    for (c1, c2), v in sorted(all_connections.items(), key=lambda x: x[1], reverse=True)[:10]:
        print(f"    {c1.title():20s} <-> {c2.title():20s}  R2={v:.4f}")

    print(f"\n  Output directory: {OUTPUT_DIR}")
    print("=" * 70)


if __name__ == '__main__':
    main()
