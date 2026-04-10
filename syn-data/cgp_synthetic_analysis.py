"""
CGP Analysis for Synthetic Data - City Connection Discovery
=============================================================
Uses Cartesian Genetic Programming to discover inter-city infection
relationships from synthetic SIR data. Validates CGP by comparing
discovered connections against known ground truth.

MATCHES REAL DATA APPROACH (2020v2):
    - Addition-only function set (no subtraction/differencing)
    - Correlation matrix
    - Network graph
    - CGP weights heatmap
    - Fitness convergence
    - Optimal lag sweep
    - Three spillover percentage tests (1%, 5%, 10%)
    - Precision/Recall/F1 comparison with ground truth

Output: cgp_results/ folder with all visualizations and summaries.
"""

import os
import sys
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SIR_RESULTS_DIR = os.path.join(BASE_DIR, 'sir_results')
OUTPUT_DIR = os.path.join(BASE_DIR, 'cgp_results')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# CGP Parameters (same as 2020v2 real data)
N_ROWS = 3
N_COLS = 8
N_OUTPUTS = 1
N_GENERATIONS = 500
LAMBDA = 4           # children per generation
MUTATION_RATE = 0.10
LEVELS_BACK = N_COLS  # Full connectivity
N_CGP_RUNS = 3       # multiple runs for robustness

# Connection threshold
R2_THRESHOLD = 0.9

# Lag sweep range
LAG_SWEEP = [1, 2, 3, 5, 7, 10, 14, 21]

# Spillover percentages to test
SPILLOVER_PERCENTAGES = [1, 5, 10]

# City list (must match SIR generation)
CITY_NAMES = ['milano', 'roma', 'napoli', 'torino', 'palermo',
              'genova', 'bologna', 'firenze', 'bari', 'catania']

# Ground truth connections (for validation)
GROUND_TRUTH = [
    ('milano', 'torino'), ('milano', 'genova'), ('milano', 'bologna'),
    ('roma', 'napoli'), ('roma', 'firenze'),
    ('napoli', 'bari'), ('napoli', 'palermo'),
    ('bologna', 'firenze'),
    ('catania', 'palermo'),
    ('bari', 'catania'),
]

# Make a set of sorted pairs for comparison
GROUND_TRUTH_SET = set(tuple(sorted(pair)) for pair in GROUND_TRUTH)


# ==============================================================================
# CGP IMPLEMENTATION (Addition-Only Functions - same as real data)
# ==============================================================================

class CartesianGeneticProgramming:
    """
    Cartesian Genetic Programming (CGP) implementation.

    Architecture:
      - Nodes arranged in a grid (n_rows x n_cols)
      - Each node: function gene + two input connection genes
      - Levels-back parameter controls connectivity
      - Evolution: (1+lambda) strategy with point mutation

    Function Set (Addition-Only - NO subtraction/differencing):
      - add:          a + b
      - max:          max(a, b)
      - min:          min(a, b)
      - avg:          (a + b) / 2
      - weighted_add: 0.7*a + 0.3*b
    """

    # Addition-only function set (same as 2020v2 real data analysis)
    FUNCTIONS = [
        ('add',          lambda a, b: a + b),
        ('max',          lambda a, b: np.maximum(a, b)),
        ('min',          lambda a, b: np.minimum(a, b)),
        ('avg',          lambda a, b: (a + b) / 2.0),
        ('weighted_add', lambda a, b: 0.7 * a + 0.3 * b),
    ]

    def __init__(self, n_inputs, n_outputs=1, n_rows=3, n_cols=8, levels_back=None):
        self.n_inputs = n_inputs
        self.n_outputs = n_outputs
        self.n_rows = n_rows
        self.n_cols = n_cols
        self.levels_back = levels_back if levels_back is not None else n_cols
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
            genome[out_start + i] = np.random.randint(
                self.n_inputs + self.n_nodes)
        return genome

    def get_active_nodes(self, genome):
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

    def get_active_functions(self, genome):
        """Get the list of functions used in the active computation graph."""
        active_nodes = self.get_active_nodes(genome)
        used_functions = []
        for node_idx in sorted(active_nodes):
            gs = node_idx * 3
            func_id = genome[gs] % self.n_functions
            func_name, _ = self.FUNCTIONS[func_id]
            used_functions.append((node_idx, func_name))
        return used_functions

    def evaluate(self, genome, X):
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
        pred = self.evaluate(genome, X)[:, 0]
        return np.mean((pred - y) ** 2)

    def mutate(self, genome, rate=0.10):
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

    def evolve(self, X, y, n_generations=500, lam=4, mutation_rate=0.10):
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
        pred = self.evaluate(genome, X)[:, 0]
        ss_res = np.sum((y - pred) ** 2)
        ss_tot = np.sum((y - np.mean(y)) ** 2)
        if ss_tot < 1e-10:
            return 0.0
        return 1.0 - ss_res / ss_tot


# ==============================================================================
# DATA PREPARATION
# ==============================================================================

def load_synthetic_data(filepath):
    """Load synthetic CSV and compute daily new infections."""
    df = pd.read_csv(filepath)
    df['data'] = pd.to_datetime(df['data'])
    df.sort_values('data', inplace=True)
    df.set_index('data', inplace=True)

    daily_new = pd.DataFrame(index=df.index)
    for city in CITY_NAMES:
        col = f"{city}_infected"
        if col in df.columns:
            cumulative = df[col].values.astype(float)
            diff = np.diff(cumulative, prepend=cumulative[0])
            daily_new[city] = np.maximum(diff, 0)

    return daily_new


def build_lagged_inputs(daily_new, target_city, candidate_cities, n_lags=7):
    """
    Build input matrix with lagged values from candidate cities.
    Uses Z-score normalization (same as real data analysis).
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

    # Remove NaN rows
    valid = ~np.isnan(X).any(axis=1) & ~np.isnan(y)
    X = X[valid]
    y = y[valid]

    # Z-score normalization
    X_mean = np.mean(X, axis=0, keepdims=True)
    X_std = np.std(X, axis=0, keepdims=True) + 1e-10
    X_norm = (X - X_mean) / X_std

    y_mean = np.mean(y)
    y_std = np.std(y) + 1e-10
    y_norm = (y - y_mean) / y_std

    return X_norm, y_norm, feature_names, valid


# ==============================================================================
# CGP RUNNER (same as real data)
# ==============================================================================

def run_cgp_for_city(X, y, n_runs=3):
    """Run CGP multiple times and return the best result."""
    best_genome = None
    best_r2 = -np.inf
    best_history = None

    n_inputs = X.shape[1]
    for run in range(n_runs):
        cgp = CartesianGeneticProgramming(
            n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
        genome, fitness, history = cgp.evolve(
            X, y, N_GENERATIONS, LAMBDA, MUTATION_RATE)
        r2 = cgp.compute_r2(genome, X, y)
        if r2 > best_r2:
            best_r2 = r2
            best_genome = genome
            best_history = history

    return best_genome, best_r2, best_history


# ==============================================================================
# OPTIMAL LAG FINDING
# ==============================================================================

def find_optimal_lag(daily_new, lag_range=None):
    """
    Run CGP with different lag values and find the one with highest avg R2.
    """
    if lag_range is None:
        lag_range = LAG_SWEEP

    print("  Lag Sweep:")
    print(f"  Testing lags: {lag_range}")

    lag_results = []

    for n_lag in lag_range:
        r2_scores = []

        for target in CITY_NAMES:
            candidates = [c for c in CITY_NAMES if c != target]

            X, y, feat_names, _ = build_lagged_inputs(
                daily_new, target, candidates, n_lags=n_lag)

            if len(y) < 30 or np.std(y) < 1e-10:
                continue

            best_genome, best_r2, _ = run_cgp_for_city(X, y, n_runs=1)
            r2_scores.append(best_r2)

        avg_r2 = np.mean(r2_scores) if r2_scores else 0
        max_r2 = np.max(r2_scores) if r2_scores else 0
        n_above_threshold = sum(1 for r in r2_scores if r > R2_THRESHOLD)

        lag_results.append({
            'lag': n_lag,
            'avg_r2': avg_r2,
            'max_r2': max_r2,
            'n_above_threshold': n_above_threshold,
            'n_cities': len(r2_scores),
        })

        print(f"    Lag={n_lag:2d}: avg R2={avg_r2:.4f}, max R2={max_r2:.4f}, "
              f"above {R2_THRESHOLD}: {n_above_threshold}/{len(r2_scores)}")

    # Find optimal lag
    best_lag_row = max(lag_results, key=lambda x: x['avg_r2'])
    optimal_lag = best_lag_row['lag']
    print(f"\n  * Optimal Lag = {optimal_lag} (avg R2 = {best_lag_row['avg_r2']:.4f})")

    return optimal_lag, lag_results


# ==============================================================================
# FULL CGP ANALYSIS (matches real data approach)
# ==============================================================================

def run_cgp_analysis(daily_new, n_lags, label=""):
    """
    Run full CGP analysis for all city pairs with given lag.
    Returns results dict matching the 2020v2 format.
    """
    print(f"\n  Running CGP analysis with lag={n_lags}... {label}")

    cgp_results = {}

    for idx, target in enumerate(CITY_NAMES):
        candidates = [c for c in CITY_NAMES if c != target]

        X, y, feat_names, _ = build_lagged_inputs(
            daily_new, target, candidates, n_lags=n_lags)

        if len(y) < 30 or np.std(y) < 1e-10:
            print(f"    [{idx+1:2d}/{len(CITY_NAMES)}] {target.title():12s} "
                  f"- SKIPPED")
            continue

        # Run CGP (best of N runs)
        best_genome, best_r2, best_history = run_cgp_for_city(
            X, y, n_runs=N_CGP_RUNS)

        # Get active inputs -> cities
        n_inputs = X.shape[1]
        cgp_eval = CartesianGeneticProgramming(
            n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
        active_input_indices = cgp_eval.get_active_inputs(best_genome)
        active_functions = cgp_eval.get_active_functions(best_genome)

        active_cities = set()
        for inp_idx in active_input_indices:
            if inp_idx < len(feat_names):
                city_from_feat = feat_names[inp_idx].rsplit('_lag', 1)[0]
                active_cities.add(city_from_feat)

        n_active = len(active_cities)
        status = "SIGNIFICANT" if best_r2 > R2_THRESHOLD else "weak"
        print(f"    [{idx+1:2d}/{len(CITY_NAMES)}] {target.title():12s}  "
              f"R2={best_r2:.4f}  Active: {n_active:2d}  [{status}]  "
              f"-> {', '.join([c.title() for c in list(active_cities)[:5]])}")

        cgp_results[target] = {
            'genome': best_genome,
            'r2': best_r2,
            'history': best_history,
            'active_cities': list(active_cities),
            'active_functions': active_functions,
            'feat_names': feat_names,
            'n_active': n_active,
        }

    # Build connections dict
    all_connections = {}
    for target, result in cgp_results.items():
        for source in result['active_cities']:
            pair = tuple(sorted([target, source]))
            current = all_connections.get(pair, 0)
            all_connections[pair] = max(current, result['r2'])

    return cgp_results, all_connections


# ==============================================================================
# GROUND TRUTH COMPARISON
# ==============================================================================

def compare_with_ground_truth(connections, label=""):
    """
    Compare CGP-discovered connections with ground truth.
    Returns precision, recall, F1.
    """
    # Discovered connections (above threshold)
    discovered = set()
    for (c1, c2), r2 in connections.items():
        if r2 > R2_THRESHOLD:
            discovered.add(tuple(sorted([c1, c2])))

    # True positives, false positives, false negatives
    tp = discovered & GROUND_TRUTH_SET
    fp = discovered - GROUND_TRUTH_SET
    fn = GROUND_TRUTH_SET - discovered

    precision = len(tp) / len(discovered) if discovered else 0
    recall = len(tp) / len(GROUND_TRUTH_SET) if GROUND_TRUTH_SET else 0
    f1 = (2 * precision * recall / (precision + recall)
          if (precision + recall) > 0 else 0)

    result = {
        'label': label,
        'n_discovered': len(discovered),
        'n_ground_truth': len(GROUND_TRUTH_SET),
        'true_positives': len(tp),
        'false_positives': len(fp),
        'false_negatives': len(fn),
        'precision': precision,
        'recall': recall,
        'f1': f1,
        'tp_connections': tp,
        'fp_connections': fp,
        'fn_connections': fn,
    }

    print(f"\n  Ground Truth Comparison {label}:")
    print(f"    Discovered: {len(discovered)}, "
          f"Ground Truth: {len(GROUND_TRUTH_SET)}")
    print(f"    TP={len(tp)}, FP={len(fp)}, FN={len(fn)}")
    print(f"    Precision={precision:.3f}, Recall={recall:.3f}, F1={f1:.3f}")

    if tp:
        print(f"    [OK] Correctly found: {', '.join(f'{a.title()}-{b.title()}' for a, b in tp)}")
    if fn:
        print(f"    [MISS] Missed: {', '.join(f'{a.title()}-{b.title()}' for a, b in fn)}")

    return result


# ==============================================================================
# VISUALIZATION 1: CORRELATION MATRIX
# ==============================================================================

def plot_correlation_matrix(daily_new, output_dir, title_suffix=""):
    """
    Plot Pearson correlation matrix between all city pairs.
    Same as real data analysis.
    """
    # Compute correlation matrix
    corr_matrix = daily_new[CITY_NAMES].corr()

    fig, ax = plt.subplots(figsize=(12, 10))
    im = ax.imshow(corr_matrix.values, cmap='RdYlBu_r', vmin=-1, vmax=1,
                    aspect='auto')

    n = len(CITY_NAMES)
    ax.set_xticks(range(n))
    ax.set_yticks(range(n))
    ax.set_xticklabels([c.title() for c in CITY_NAMES], rotation=45,
                        ha='right', fontsize=10)
    ax.set_yticklabels([c.title() for c in CITY_NAMES], fontsize=10)

    # Annotate values
    for i in range(n):
        for j in range(n):
            val = corr_matrix.values[i, j]
            color = 'white' if abs(val) > 0.6 else 'black'
            ax.text(j, i, f'{val:.2f}', ha='center', va='center',
                    fontsize=9, color=color, fontweight='bold')

    plt.colorbar(im, ax=ax, label='Pearson Correlation', shrink=0.8)
    ax.set_title(f'Correlation Matrix - Daily New Infections{title_suffix}',
                 fontsize=14, fontweight='bold')
    plt.tight_layout()

    safe_suffix = title_suffix.replace('%', 'pct').replace(' ', '_').replace('(', '').replace(')', '')
    plt.savefig(os.path.join(output_dir,
                f'correlation_matrix{safe_suffix}.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: correlation_matrix{safe_suffix}.png")

    # Save correlation CSV
    corr_matrix.to_csv(os.path.join(output_dir,
                       f'correlation_matrix{safe_suffix}.csv'))
    print(f"  Saved: correlation_matrix{safe_suffix}.csv")

    return corr_matrix


# ==============================================================================
# VISUALIZATION 2: LAG SWEEP PLOT
# ==============================================================================

def plot_lag_sweep(lag_results, output_dir, pct_label=""):
    """Plot lag sweep results."""
    lags = [r['lag'] for r in lag_results]
    avg_r2s = [r['avg_r2'] for r in lag_results]
    max_r2s = [r['max_r2'] for r in lag_results]
    n_above = [r['n_above_threshold'] for r in lag_results]

    fig, axes = plt.subplots(1, 3, figsize=(18, 5))
    fig.suptitle(f'Lag Sweep Results {pct_label}\nAddition-Only Functions',
                 fontsize=14, fontweight='bold')

    # Avg R2
    ax = axes[0]
    ax.plot(lags, avg_r2s, 'bo-', lw=2, markersize=8)
    best_idx = np.argmax(avg_r2s)
    ax.plot(lags[best_idx], avg_r2s[best_idx], 'r*', markersize=20,
            label=f'Best: lag={lags[best_idx]}')
    ax.set_xlabel('Lag (days)', fontsize=12)
    ax.set_ylabel('Average R2', fontsize=12)
    ax.set_title('Average R2 vs Lag')
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=11)

    # Max R2
    ax = axes[1]
    ax.plot(lags, max_r2s, 'go-', lw=2, markersize=8)
    ax.axhline(y=R2_THRESHOLD, color='red', ls='--',
               label=f'Threshold={R2_THRESHOLD}')
    ax.set_xlabel('Lag (days)', fontsize=12)
    ax.set_ylabel('Max R2', fontsize=12)
    ax.set_title('Max R2 vs Lag')
    ax.grid(True, alpha=0.3)
    ax.legend(fontsize=11)

    # N above threshold
    ax = axes[2]
    ax.bar(lags, n_above, color='#3498db', alpha=0.8)
    ax.set_xlabel('Lag (days)', fontsize=12)
    ax.set_ylabel(f'Cities with R2 > {R2_THRESHOLD}', fontsize=12)
    ax.set_title('Number of Cities Above Threshold')
    ax.grid(True, alpha=0.3, axis='y')

    plt.tight_layout()
    safe_label = pct_label.replace('%', 'pct').replace(' ', '_')
    plt.savefig(os.path.join(output_dir,
                f'lag_sweep_{safe_label}.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: lag_sweep_{safe_label}.png")


# ==============================================================================
# VISUALIZATION 3: CONNECTION HEATMAP (same as real data)
# ==============================================================================

def plot_connection_heatmap(connections, city_names, output_dir,
                             title_suffix=""):
    """Heatmap of CGP-discovered connection strengths (same as real data)."""
    n = len(city_names)
    matrix = np.zeros((n, n))
    for i, c1 in enumerate(city_names):
        for j, c2 in enumerate(city_names):
            pair = tuple(sorted([c1, c2]))
            if pair in connections:
                matrix[i, j] = connections[pair]

    fig, ax = plt.subplots(figsize=(12, 10))
    im = ax.imshow(matrix, cmap='YlOrRd', vmin=0, vmax=1, aspect='auto')

    ax.set_xticks(range(n))
    ax.set_yticks(range(n))
    ax.set_xticklabels([c.title() for c in city_names], rotation=45,
                        ha='right', fontsize=10)
    ax.set_yticklabels([c.title() for c in city_names], fontsize=10)

    # Annotate values
    for i in range(n):
        for j in range(n):
            if matrix[i, j] > 0.01:
                color = 'white' if matrix[i, j] > 0.6 else 'black'
                ax.text(j, i, f'{matrix[i, j]:.2f}', ha='center',
                        va='center', fontsize=8, color=color,
                        fontweight='bold')

    plt.colorbar(im, ax=ax, label='Connection Strength (R2)', shrink=0.8)
    ax.set_title(f'CGP Connection Strength Heatmap (R2){title_suffix}\n'
                 f'Addition-Only Functions',
                 fontsize=13, fontweight='bold')
    plt.tight_layout()

    safe_suffix = title_suffix.replace('%', 'pct').replace(' ', '_').replace('(', '').replace(')', '')
    plt.savefig(os.path.join(output_dir,
                f'connection_heatmap{safe_suffix}.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: connection_heatmap{safe_suffix}.png")


# ==============================================================================
# VISUALIZATION 4: NETWORK GRAPH (same as real data)
# ==============================================================================

def plot_network_graph(connections, city_names, output_dir, gt_comparison=None,
                        title_suffix=""):
    """Plot network graph with ground truth highlighting (same as real data)."""
    try:
        import networkx as nx
    except ImportError:
        print("  networkx not available, skipping network graph")
        return

    G = nx.Graph()
    for city in city_names:
        G.add_node(city)

    for (c1, c2), r2 in connections.items():
        if r2 > R2_THRESHOLD:
            pair = tuple(sorted([c1, c2]))
            is_gt = pair in GROUND_TRUTH_SET
            G.add_edge(c1, c2, weight=r2, is_ground_truth=is_gt)

    if len(G.edges()) == 0:
        print("  No significant connections for network graph")
        fig, ax = plt.subplots(figsize=(12, 10))
        pos = nx.spring_layout(G, k=2, seed=42)
        nx.draw_networkx_nodes(G, pos, ax=ax, node_size=500,
                                node_color='lightgray',
                                edgecolors='black', linewidths=1)
        labels = {n: n.title() for n in G.nodes()}
        nx.draw_networkx_labels(G, pos, labels=labels, ax=ax,
                                 font_size=8, font_weight='bold')
        ax.set_title(f'CGP Network (R2 > {R2_THRESHOLD}){title_suffix}\n'
                     'No connections meet the threshold',
                     fontsize=14, fontweight='bold')
        ax.axis('off')
        plt.tight_layout()
        safe_suffix = title_suffix.replace('%', 'pct').replace(' ', '_').replace('(', '').replace(')', '')
        plt.savefig(os.path.join(output_dir,
                    f'network_graph{safe_suffix}.png'),
                    dpi=150, bbox_inches='tight')
        plt.close()
        print(f"  Saved: network_graph{safe_suffix}.png (no edges)")
        return

    pos = nx.spring_layout(G, k=2.5, iterations=100, seed=42)

    fig, ax = plt.subplots(figsize=(14, 11))

    # Separate edges by type
    gt_edges = [(u, v) for u, v, d in G.edges(data=True)
                if d.get('is_ground_truth', False)]
    fp_edges = [(u, v) for u, v, d in G.edges(data=True)
                if not d.get('is_ground_truth', False)]

    # Draw edges
    if gt_edges:
        nx.draw_networkx_edges(G, pos, edgelist=gt_edges, ax=ax,
                                edge_color='#2ecc71', width=3, alpha=0.8)
    if fp_edges:
        nx.draw_networkx_edges(G, pos, edgelist=fp_edges, ax=ax,
                                edge_color='#e74c3c', width=2, alpha=0.5,
                                style='dashed')

    # Draw nodes - size by degree
    degrees = dict(G.degree())
    node_sizes = [max(degrees.get(n, 0) * 200, 500) for n in G.nodes()]
    nx.draw_networkx_nodes(G, pos, ax=ax, node_size=node_sizes,
                            node_color='#3498db', edgecolors='#2c3e50',
                            linewidths=2, alpha=0.9)

    labels = {n: n.title() for n in G.nodes()}
    nx.draw_networkx_labels(G, pos, labels=labels, ax=ax, font_size=10,
                             font_weight='bold')

    # Edge labels (R2)
    edge_labels = {(u, v): f"{d['weight']:.2f}"
                   for u, v, d in G.edges(data=True)}
    nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, ax=ax,
                                  font_size=8)

    # Legend
    from matplotlib.lines import Line2D
    legend_elements = [
        Line2D([0], [0], color='#2ecc71', lw=3,
               label='True Positive (correct)'),
        Line2D([0], [0], color='#e74c3c', lw=2, ls='--',
               label='False Positive (extra)'),
    ]
    ax.legend(handles=legend_elements, loc='upper left', fontsize=11)

    ax.set_title(f'CGP-Discovered Network (R2 > {R2_THRESHOLD}){title_suffix}\n'
                 f'Addition-Only Functions',
                 fontsize=14, fontweight='bold')
    ax.axis('off')
    plt.tight_layout()

    safe_suffix = title_suffix.replace('%', 'pct').replace(' ', '_').replace('(', '').replace(')', '')
    plt.savefig(os.path.join(output_dir,
                f'network_graph{safe_suffix}.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: network_graph{safe_suffix}.png")


# ==============================================================================
# VISUALIZATION 5: CGP WEIGHTS HEATMAP (same as real data)
# ==============================================================================

def plot_cgp_weights(cgp_results, city_names, output_dir, title_suffix=""):
    """
    Plot the CGP-discovered weights (influence matrix).
    Shows which cities influence which - same as real data cgp_weights_heatmap.
    """
    n = len(city_names)
    weight_matrix = np.zeros((n, n))

    for i, target in enumerate(city_names):
        if target in cgp_results:
            for source in cgp_results[target]['active_cities']:
                if source in city_names:
                    j = city_names.index(source)
                    weight_matrix[i, j] = cgp_results[target]['r2']

    fig, ax = plt.subplots(figsize=(12, 10))
    im = ax.imshow(weight_matrix, cmap='Blues', vmin=0, vmax=1, aspect='auto')
    ax.set_xticks(range(n))
    ax.set_yticks(range(n))
    ax.set_xticklabels([c.title() for c in city_names], rotation=45,
                        ha='right', fontsize=10)
    ax.set_yticklabels([c.title() for c in city_names], fontsize=10)

    for i in range(n):
        for j in range(n):
            if weight_matrix[i, j] > 0.01:
                color = 'white' if weight_matrix[i, j] > 0.6 else 'black'
                ax.text(j, i, f'{weight_matrix[i, j]:.2f}', ha='center',
                        va='center', fontsize=8, color=color,
                        fontweight='bold')

    ax.set_xlabel('Source City (Influencer)', fontsize=12)
    ax.set_ylabel('Target City (Influenced)', fontsize=12)
    ax.set_title(f'CGP-Discovered Influence Weights{title_suffix}\n'
                 f'Which cities influence which? (R2 metric)',
                 fontsize=14, fontweight='bold')
    plt.colorbar(im, ax=ax, label='Influence Weight (R2)', shrink=0.8)
    plt.tight_layout()

    safe_suffix = title_suffix.replace('%', 'pct').replace(' ', '_').replace('(', '').replace(')', '')
    plt.savefig(os.path.join(output_dir,
                f'cgp_weights_heatmap{safe_suffix}.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: cgp_weights_heatmap{safe_suffix}.png")


# ==============================================================================
# VISUALIZATION 6: FITNESS CONVERGENCE (same as real data)
# ==============================================================================

def plot_fitness_convergence(cgp_results, output_dir, title_suffix=""):
    """Plot CGP fitness convergence for all cities (same as real data)."""
    cities_with_history = [c for c in cgp_results
                           if cgp_results[c].get('history')]
    n = len(cities_with_history)
    if n == 0:
        return

    cols = 5
    rows = (n + cols - 1) // cols
    fig, axes = plt.subplots(rows, cols, figsize=(25, 5 * rows))
    fig.suptitle(f'CGP Fitness Convergence - 10 Cities{title_suffix}\n'
                 f'Addition-Only Functions',
                 fontsize=16, fontweight='bold')

    if rows == 1:
        axes = [axes]
    axes_flat = np.array(axes).flatten()

    for idx, city in enumerate(cities_with_history):
        ax = axes_flat[idx]
        ax.plot(cgp_results[city]['history'], 'b-', lw=1)
        ax.set_title(f'{city.title()}\nR2={cgp_results[city]["r2"]:.4f}',
                     fontsize=10, fontweight='bold')
        ax.set_xlabel('Generation')
        ax.set_ylabel('Fitness (MSE)')
        ax.grid(True, alpha=0.3)
        try:
            ax.set_yscale('log')
        except:
            pass

    for idx in range(n, len(axes_flat)):
        axes_flat[idx].set_visible(False)

    plt.tight_layout()
    safe_suffix = title_suffix.replace('%', 'pct').replace(' ', '_').replace('(', '').replace(')', '')
    plt.savefig(os.path.join(output_dir,
                f'cgp_fitness_convergence{safe_suffix}.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: cgp_fitness_convergence{safe_suffix}.png")


# ==============================================================================
# VISUALIZATION 7: PRECISION/RECALL COMPARISON
# ==============================================================================

def plot_precision_recall_comparison(gt_results, output_dir):
    """Bar chart comparing precision/recall/F1 across spillover percentages."""
    labels = [r['label'] for r in gt_results]
    precision = [r['precision'] for r in gt_results]
    recall = [r['recall'] for r in gt_results]
    f1 = [r['f1'] for r in gt_results]

    x = np.arange(len(labels))
    width = 0.25

    fig, ax = plt.subplots(figsize=(12, 7))
    bars1 = ax.bar(x - width, precision, width, label='Precision',
                    color='#3498db', alpha=0.85)
    bars2 = ax.bar(x, recall, width, label='Recall',
                    color='#2ecc71', alpha=0.85)
    bars3 = ax.bar(x + width, f1, width, label='F1 Score',
                    color='#e74c3c', alpha=0.85)

    # Add value labels
    for bars in [bars1, bars2, bars3]:
        for bar in bars:
            h = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., h + 0.02,
                    f'{h:.2f}', ha='center', va='bottom', fontsize=10,
                    fontweight='bold')

    ax.set_xlabel('Spillover Percentage', fontsize=12)
    ax.set_ylabel('Score', fontsize=12)
    ax.set_title('CGP Validation: Ground Truth Recovery\n'
                 'Precision / Recall / F1 by Spillover % (Addition-Only)',
                 fontsize=14, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=11)
    ax.legend(fontsize=11)
    ax.set_ylim(0, 1.15)
    ax.grid(True, alpha=0.3, axis='y')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'precision_recall_comparison.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: precision_recall_comparison.png")


def plot_discovered_vs_gt(gt_results, output_dir):
    """Visual comparison of discovered vs ground truth connections."""
    fig, axes = plt.subplots(1, 3, figsize=(24, 8))

    for idx, result in enumerate(gt_results):
        ax = axes[idx]
        label = result['label']

        categories = ['True Positive\n(Correct)', 'False Positive\n(Extra)',
                       'False Negative\n(Missed)']
        values = [result['true_positives'], result['false_positives'],
                  result['false_negatives']]
        colors = ['#2ecc71', '#e74c3c', '#f39c12']

        bars = ax.bar(categories, values, color=colors, alpha=0.85,
                       edgecolor='black', linewidth=1.5)

        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width()/2., bar.get_height() + 0.2,
                    str(val), ha='center', va='bottom', fontsize=14,
                    fontweight='bold')

        ax.set_title(f'{label}\nP={result["precision"]:.2f}, '
                     f'R={result["recall"]:.2f}, '
                     f'F1={result["f1"]:.2f}',
                     fontsize=12, fontweight='bold')
        ax.set_ylabel('Count')
        ax.grid(True, alpha=0.3, axis='y')
        ax.set_ylim(0, max(values) + 2)

    fig.suptitle('CGP Connection Discovery: Ground Truth Comparison\n'
                 'Addition-Only Functions',
                 fontsize=16, fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'discovered_vs_ground_truth.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: discovered_vs_ground_truth.png")


# ==============================================================================
# SAVE RESULTS
# ==============================================================================

def save_results(connections, cgp_results, gt_comparison, pct_label, output_dir):
    """Save connection data to CSV."""
    safe = pct_label.replace('%', 'pct').replace(' ', '_')

    # Connections CSV
    conn_rows = []
    for (c1, c2), r2 in sorted(connections.items(),
                                 key=lambda x: x[1], reverse=True):
        pair = tuple(sorted([c1, c2]))
        conn_rows.append({
            'city_1': c1.title(),
            'city_2': c2.title(),
            'connection_strength_R2': round(r2, 4),
            'is_significant': r2 > R2_THRESHOLD,
            'is_ground_truth': pair in GROUND_TRUTH_SET,
            'classification': (
                'TP' if r2 > R2_THRESHOLD and pair in GROUND_TRUTH_SET else
                'FP' if r2 > R2_THRESHOLD and pair not in GROUND_TRUTH_SET else
                'FN' if r2 <= R2_THRESHOLD and pair in GROUND_TRUTH_SET else
                'TN'
            ),
        })
    conn_df = pd.DataFrame(conn_rows)
    conn_df.to_csv(os.path.join(output_dir, f'connections_{safe}.csv'),
                    index=False)
    print(f"  Saved: connections_{safe}.csv")

    # City summary CSV
    summary_rows = []
    for city in CITY_NAMES:
        res = cgp_results.get(city, {})
        linked = res.get('active_cities', [])
        r2 = res.get('r2', 0)
        funcs = res.get('active_functions', [])
        summary_rows.append({
            'city': city.title(),
            'r2_score': round(r2, 4),
            'n_connections': len(linked),
            'linked_cities': '; '.join([c.title() for c in linked]),
            'active_functions': '; '.join([str(f[1]) if isinstance(f, (tuple, list)) else str(f) for f in funcs]),
        })
    summary_df = pd.DataFrame(summary_rows)
    summary_df.sort_values('r2_score', ascending=False, inplace=True)
    summary_df.to_csv(os.path.join(output_dir, f'city_summary_{safe}.csv'),
                        index=False)
    print(f"  Saved: city_summary_{safe}.csv")


# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

def main():
    print("=" * 70)
    print("   CGP Analysis - Synthetic Data")
    print("   Addition-Only Functions (same as real data)")
    print(f"   Functions: {[f[0] for f in CartesianGeneticProgramming.FUNCTIONS]}")
    print(f"   R2 Threshold = {R2_THRESHOLD}")
    print(f"   Ground Truth Connections: {len(GROUND_TRUTH)}")
    print("=" * 70)

    all_gt_results = []
    all_optimal_lags = {}
    all_lag_results = {}

    for pct in SPILLOVER_PERCENTAGES:
        pct_label = f"{pct}%"
        csv_path = os.path.join(SIR_RESULTS_DIR,
                                f'synthetic_data_{pct}pct.csv')

        print(f"\n{'='*70}")
        print(f"  ANALYZING: {pct_label} Spillover")
        print(f"{'='*70}")

        if not os.path.exists(csv_path):
            print(f"  ERROR: {csv_path} not found! Run syn_data_sir_model.py first.")
            continue

        # Load data
        print(f"\n[1/7] Loading synthetic data ({pct_label})...")
        daily_new = load_synthetic_data(csv_path)
        print(f"  Loaded: {len(daily_new)} days x {len(daily_new.columns)} cities")

        # Correlation matrix
        print(f"\n[2/7] Computing correlation matrix ({pct_label})...")
        suffix = f" ({pct_label} spillover)"
        corr_matrix = plot_correlation_matrix(daily_new, OUTPUT_DIR,
                                              title_suffix=suffix)

        # Find optimal lag
        print(f"\n[3/7] Finding optimal lag ({pct_label})...")
        optimal_lag, lag_results = find_optimal_lag(daily_new)
        all_optimal_lags[pct_label] = optimal_lag
        all_lag_results[pct_label] = lag_results

        # Plot lag sweep
        plot_lag_sweep(lag_results, OUTPUT_DIR,
                       pct_label=f"({pct_label} spillover)")

        # Run full CGP with optimal lag
        print(f"\n[4/7] Running full CGP analysis with lag={optimal_lag}...")
        cgp_results, all_connections = run_cgp_analysis(
            daily_new, optimal_lag, label=f"({pct_label})")

        # Compare with ground truth
        print(f"\n[5/7] Comparing with ground truth...")
        gt_comparison = compare_with_ground_truth(
            all_connections, label=pct_label)
        all_gt_results.append(gt_comparison)

        # Generate ALL visualizations (same as real data)
        print(f"\n[6/7] Generating visualizations ({pct_label})...")

        # Connection heatmap
        plot_connection_heatmap(all_connections, CITY_NAMES,
                                 OUTPUT_DIR, title_suffix=suffix)
        # Network graph
        plot_network_graph(all_connections, CITY_NAMES,
                            OUTPUT_DIR, gt_comparison, title_suffix=suffix)
        # CGP weights heatmap
        plot_cgp_weights(cgp_results, CITY_NAMES, OUTPUT_DIR,
                          title_suffix=suffix)
        # Fitness convergence
        plot_fitness_convergence(cgp_results, OUTPUT_DIR,
                                  title_suffix=suffix)

        # Save CSVs
        print(f"\n[7/7] Saving data ({pct_label})...")
        save_results(all_connections, cgp_results,
                      gt_comparison, pct_label, OUTPUT_DIR)

    # Cross-percentage comparison
    print(f"\n{'='*70}")
    print("  CROSS-PERCENTAGE COMPARISON")
    print(f"{'='*70}")

    if all_gt_results:
        plot_precision_recall_comparison(all_gt_results, OUTPUT_DIR)
        plot_discovered_vs_gt(all_gt_results, OUTPUT_DIR)

    # Save lag sweep summary
    lag_summary_rows = []
    for pct_label, lag_results in all_lag_results.items():
        for r in lag_results:
            lag_summary_rows.append({
                'spillover': pct_label,
                'lag': r['lag'],
                'avg_r2': round(r['avg_r2'], 4),
                'max_r2': round(r['max_r2'], 4),
                'n_above_threshold': r['n_above_threshold'],
            })
    if lag_summary_rows:
        lag_df = pd.DataFrame(lag_summary_rows)
        lag_df.to_csv(os.path.join(OUTPUT_DIR, 'lag_sweep_all.csv'),
                       index=False)
        print("  Saved: lag_sweep_all.csv")

    # Save GT comparison summary
    gt_summary_rows = []
    for r in all_gt_results:
        gt_summary_rows.append({
            'spillover': r['label'],
            'discovered': r['n_discovered'],
            'ground_truth': r['n_ground_truth'],
            'true_positives': r['true_positives'],
            'false_positives': r['false_positives'],
            'false_negatives': r['false_negatives'],
            'precision': round(r['precision'], 4),
            'recall': round(r['recall'], 4),
            'f1': round(r['f1'], 4),
            'optimal_lag': all_optimal_lags.get(r['label'], 'N/A'),
        })
    if gt_summary_rows:
        gt_df = pd.DataFrame(gt_summary_rows)
        gt_df.to_csv(os.path.join(OUTPUT_DIR, 'ground_truth_comparison.csv'),
                       index=False)
        print("  Saved: ground_truth_comparison.csv")

    # Final summary
    print("\n" + "=" * 70)
    print("   CGP ANALYSIS SUMMARY (Addition-Only Functions)")
    print("=" * 70)
    print(f"  Functions: {[f[0] for f in CartesianGeneticProgramming.FUNCTIONS]}")
    for r in all_gt_results:
        print(f"\n  {r['label']} Spillover:")
        print(f"    Optimal Lag: {all_optimal_lags.get(r['label'], 'N/A')}")
        print(f"    Discovered: {r['n_discovered']} connections")
        print(f"    Precision: {r['precision']:.3f}")
        print(f"    Recall:    {r['recall']:.3f}")
        print(f"    F1 Score:  {r['f1']:.3f}")

    print(f"\n  Output directory: {OUTPUT_DIR}")
    print("=" * 70)


if __name__ == '__main__':
    main()
