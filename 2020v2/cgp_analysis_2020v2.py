"""
CGP (Cartesian Genetic Programming) - City Relationship Discovery (2020v2)
============================================================================
Enhanced CGP analysis for 15 selected Italian cities with:
  1. Addition-only function set (no subtraction/differencing)
  2. R² threshold = 0.9 for significant connections
  3. Lag sweep loop (3 to 100) with R² vs lag plot
  4. Pre/post lockdown behavior comparison (lockdown: March 9, 2020)
  5. Levels-back parameter explanation capability

Cities:
    North: Milano, Bergamo, Brescia, Monza e della Brianza, Como
    Center: Roma, Firenze, Perugia, Latina, Frosinone
    South: Napoli, Caserta, Salerno, Bari, Taranto

Output: 2020v2/cgp_results/
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

DATA_PATH = r"c:\Users\Utente\Documents\GitHub\gcpaida\2020\merged_cities_2020.csv"
OUTPUT_DIR = r"c:\Users\Utente\Documents\GitHub\gcpaida\2020v2\cgp_results"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# City groups
CITY_GROUPS = {
    'North': ['milano', 'bergamo', 'brescia', 'monza_e_della_brianza', 'como'],
    'Center': ['roma', 'firenze', 'perugia', 'latina', 'frosinone'],
    'South': ['napoli', 'caserta', 'salerno', 'bari', 'taranto'],
}

ALL_CITIES = []
for group_cities in CITY_GROUPS.values():
    ALL_CITIES.extend(group_cities)

CITY_REGION_MAP = {}
for region, cities in CITY_GROUPS.items():
    for city in cities:
        CITY_REGION_MAP[city] = region

REGION_COLORS = {
    'North': '#2196F3',
    'Center': '#4CAF50',
    'South': '#FF9800',
}

# CGP Parameters
N_ROWS = 3
N_COLS = 8
N_OUTPUTS = 1
N_GENERATIONS = 500          # Increased for better convergence with R²>0.9
LAMBDA = 4
MUTATION_RATE = 0.10
LEVELS_BACK = N_COLS         # Full connectivity (explained in report)
DEFAULT_N_LAGS = 7           # Default lag for main analysis

# R² threshold for significant connections (raised to 0.9)
R2_THRESHOLD = 0.9

# Lag sweep values for R² vs lag analysis
LAG_SWEEP_VALUES = [3, 7, 10, 15, 20, 25, 30, 40, 50, 60, 70, 80, 90, 100]

# Lockdown date: March 9, 2020 (Italy national lockdown)
LOCKDOWN_DATE = pd.Timestamp('2020-03-09')


# ==============================================================================
# CGP IMPLEMENTATION (Addition-Only Functions)
# ==============================================================================

class CartesianGeneticProgramming:
    """
    Cartesian Genetic Programming (CGP) implementation.

    Architecture:
      - Nodes arranged in a grid (n_rows × n_cols)
      - Each node: function gene + two input connection genes
      - Levels-back parameter controls how far back nodes can connect
      - Evolution: (1+λ) strategy with point mutation

    Function Set (Addition-Only):
      - add:          a + b
      - max:          max(a, b)
      - min:          min(a, b)
      - avg:          (a + b) / 2
      - weighted_add: 0.7*a + 0.3*b

    No subtraction, division, or differencing operations are used.
    """

    # Addition-only function set
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

    def _max_connection(self, col):
        """
        Compute maximum connection index for a node in a given column.
        With levels_back = L, a node in column c can connect to:
          - All inputs (indices 0 to n_inputs-1)
          - Nodes in columns max(0, c - L) to c - 1
        """
        if self.levels_back >= col:
            # Can connect to all previous columns and inputs
            return self.n_inputs + col * self.n_rows
        else:
            # Limited connectivity: only last levels_back columns
            first_allowed_col = col - self.levels_back
            return self.n_inputs + first_allowed_col * self.n_rows + (col - first_allowed_col) * self.n_rows

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

def prepare_data(filepath, selected_cities=None):
    """Load data and compute daily new infections for selected cities."""
    df = pd.read_csv(filepath)
    df['data'] = pd.to_datetime(df['data'])
    df.sort_values('data', inplace=True)
    df.set_index('data', inplace=True)

    # 2020 format: infected_CityName
    city_cols = [c for c in df.columns if c.startswith('infected_')]

    daily_new = pd.DataFrame(index=df.index)
    city_names = []

    for col in city_cols:
        city_key = col.replace('infected_', '').lower().replace(' ', '_')

        # Filter to selected cities if provided
        if selected_cities is not None and city_key not in selected_cities:
            continue

        cumulative = df[col].values.astype(float)
        cumulative = np.nan_to_num(cumulative, nan=0)
        diff = np.diff(cumulative, prepend=cumulative[0])
        daily_new[city_key] = np.maximum(diff, 0)
        city_names.append(city_key)

    return daily_new, city_names


def build_lagged_inputs(daily_new, target_city, candidate_cities, n_lags=7):
    """
    Build input matrix with lagged values from candidate cities.
    Uses Z-score normalization for fair comparison across cities.

    Z-score normalization rationale:
      - Different cities have vastly different infection scales
      - Without normalization, CGP would bias toward high-magnitude cities
      - z = (x - mean) / std centers data at 0 with unit variance
      - Ensures all cities contribute equally regardless of absolute numbers
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

    # Z-score normalization
    X_mean = np.mean(X, axis=0, keepdims=True)
    X_std = np.std(X, axis=0, keepdims=True) + 1e-10
    X_norm = (X - X_mean) / X_std

    y_mean = np.mean(y)
    y_std = np.std(y) + 1e-10
    y_norm = (y - y_mean) / y_std

    return X_norm, y_norm, feature_names, valid


def build_self_lagged_inputs(daily_new, target_city, n_lags=7):
    """
    Build input matrix using ONLY the city's own previous days.
    Used for post-lockdown analysis where inter-city mobility is restricted.
    """
    features = []
    feature_names = []

    for lag in range(1, n_lags + 1):  # Start from lag 1 (previous day)
        col = daily_new[target_city].shift(lag).values
        features.append(col)
        feature_names.append(f"{target_city}_lag{lag}")

    X = np.column_stack(features)
    y = daily_new[target_city].values

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
# MAIN CGP ANALYSIS
# ==============================================================================

def run_cgp_for_city(cgp_class, X, y, n_runs=3):
    """Run CGP multiple times and return the best result."""
    best_genome = None
    best_r2 = -np.inf
    best_history = None

    n_inputs = X.shape[1]
    for run in range(n_runs):
        cgp = cgp_class(n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
        genome, fitness, history = cgp.evolve(X, y, N_GENERATIONS, LAMBDA, MUTATION_RATE)
        r2 = cgp.compute_r2(genome, X, y)
        if r2 > best_r2:
            best_r2 = r2
            best_genome = genome
            best_history = history

    return best_genome, best_r2, best_history


def analyze_inter_city(daily_new, city_names, n_lags=7):
    """
    Run CGP to analyze inter-city influence.
    Each city's infections are predicted from all other cities' lagged data.
    """
    results = {}

    for idx, target in enumerate(city_names):
        candidates = [c for c in city_names if c != target]

        X, y, feat_names, valid_mask = build_lagged_inputs(
            daily_new, target, candidates, n_lags=n_lags
        )

        if len(y) < 30 or np.std(y) < 1e-10:
            print(f"  [{idx+1:2d}/{len(city_names)}] {target.title():30s} - SKIPPED")
            continue

        n_inputs = X.shape[1]
        cgp = CartesianGeneticProgramming(n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
        best_genome, best_r2, history = run_cgp_for_city(
            CartesianGeneticProgramming, X, y, n_runs=3
        )

        # Get active inputs -> cities
        cgp_eval = CartesianGeneticProgramming(n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
        active_input_indices = cgp_eval.get_active_inputs(best_genome)
        active_functions = cgp_eval.get_active_functions(best_genome)

        active_cities = set()
        for inp_idx in active_input_indices:
            if inp_idx < len(feat_names):
                city_from_feat = feat_names[inp_idx].rsplit('_lag', 1)[0]
                active_cities.add(city_from_feat)

        status = "SIGNIFICANT" if best_r2 > R2_THRESHOLD else "weak"
        print(f"  [{idx+1:2d}/{len(city_names)}] {target.title():30s}  "
              f"R²={best_r2:.4f}  Active: {len(active_cities):2d}  [{status}]"
              f"  -> {', '.join([c.title() for c in list(active_cities)[:4]])}")

        results[target] = {
            'genome': best_genome,
            'r2': best_r2,
            'history': history,
            'active_cities': list(active_cities),
            'active_functions': active_functions,
            'feat_names': feat_names,
            'n_active': len(active_cities),
        }

    return results


# ==============================================================================
# LAG SWEEP ANALYSIS
# ==============================================================================

def run_lag_sweep(daily_new, city_names, lag_values):
    """
    Run CGP with different lag values and record R² for each.
    This shows how prediction quality changes with temporal lookback.
    """
    print(f"\n  Running lag sweep: {lag_values}")
    lag_r2_results = {city: {} for city in city_names}

    for lag in lag_values:
        print(f"\n  === Lag = {lag} ===")
        for target in city_names:
            candidates = [c for c in city_names if c != target]

            X, y, feat_names, valid_mask = build_lagged_inputs(
                daily_new, target, candidates, n_lags=lag
            )

            if len(y) < 30 or np.std(y) < 1e-10:
                lag_r2_results[target][lag] = 0.0
                continue

            # Quick CGP run (fewer generations for sweep)
            n_inputs = X.shape[1]
            cgp = CartesianGeneticProgramming(n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
            genome, _, _ = cgp.evolve(X, y, n_generations=200, lam=LAMBDA, mutation_rate=MUTATION_RATE)
            r2 = cgp.compute_r2(genome, X, y)
            lag_r2_results[target][lag] = r2

            print(f"    {target.title():25s}  lag={lag:3d}  R²={r2:.4f}")

    return lag_r2_results


def plot_lag_vs_r2(lag_r2_results, lag_values, output_dir):
    """Plot R² vs lag for each city."""
    fig, axes = plt.subplots(3, 5, figsize=(25, 15))
    fig.suptitle('R² vs Lag Analysis (CGP Addition-Only)\n'
                 'How prediction quality changes with temporal lookback',
                 fontsize=16, fontweight='bold')

    all_cities = list(lag_r2_results.keys())
    for idx, city in enumerate(all_cities):
        row = idx // 5
        col = idx % 5
        ax = axes[row, col]

        r2_values = [lag_r2_results[city].get(lag, 0) for lag in lag_values]
        region = CITY_REGION_MAP.get(city, 'Unknown')
        color = REGION_COLORS.get(region, 'gray')

        ax.plot(lag_values, r2_values, 'o-', color=color, lw=2, markersize=6)
        ax.axhline(y=R2_THRESHOLD, color='red', ls='--', alpha=0.7, label=f'R²={R2_THRESHOLD}')
        ax.set_title(f'{city.title()} ({region})', fontsize=10, fontweight='bold')
        ax.set_xlabel('Lag (days)')
        ax.set_ylabel('R²')
        ax.set_ylim(-0.1, 1.05)
        ax.grid(True, alpha=0.3)
        ax.legend(fontsize=7)

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'lag_vs_r2_plot.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: lag_vs_r2_plot.png")

    # Also create a combined plot
    fig, ax = plt.subplots(figsize=(14, 8))
    for city in all_cities:
        r2_values = [lag_r2_results[city].get(lag, 0) for lag in lag_values]
        region = CITY_REGION_MAP.get(city, 'Unknown')
        color = REGION_COLORS.get(region, 'gray')
        ax.plot(lag_values, r2_values, 'o-', color=color, lw=1.5, alpha=0.7,
                markersize=5, label=city.title())

    ax.axhline(y=R2_THRESHOLD, color='red', ls='--', lw=2, label=f'Threshold R²={R2_THRESHOLD}')
    ax.set_xlabel('Lag (days)', fontsize=13)
    ax.set_ylabel('R²', fontsize=13)
    ax.set_title('R² vs Lag - All 15 Cities Combined\n'
                 '(How prediction quality changes with temporal lookback)',
                 fontsize=14, fontweight='bold')
    ax.legend(fontsize=7, ncol=3, loc='lower right')
    ax.grid(True, alpha=0.3)
    ax.set_ylim(-0.1, 1.05)

    patches = [mpatches.Patch(color=c, label=r) for r, c in REGION_COLORS.items()]
    ax.legend(handles=patches + [plt.Line2D([], [], color='red', ls='--', label=f'R²={R2_THRESHOLD}')],
              loc='upper right', fontsize=10)

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'lag_vs_r2_combined.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: lag_vs_r2_combined.png")


# ==============================================================================
# PRE/POST LOCKDOWN ANALYSIS
# ==============================================================================

def run_lockdown_analysis(daily_new, city_names, lockdown_date, n_lags=7):
    """
    Compare CGP results across three lockdown periods.

    Period 1 - Pre-lockdown (before March 9): free mobility between cities
               Uses lag=2 (only ~14 days available, must maximize usable samples)
    Period 2 - During strict lockdown (March 9 - May 18): severe restrictions
               Inter-city influence should be minimal; self-lag should suffice
    Period 3 - After easing (May 18+): gradual reopening
               Inter-city influence should partially re-emerge
    """
    lockdown_start = lockdown_date
    lockdown_end = pd.Timestamp('2020-05-18')  # Phase 2 reopening

    pre_lockdown = daily_new[daily_new.index < lockdown_start]
    during_lockdown = daily_new[(daily_new.index >= lockdown_start) & (daily_new.index < lockdown_end)]
    after_easing = daily_new[daily_new.index >= lockdown_end]

    print(f"\n  Lockdown start: {lockdown_start.strftime('%Y-%m-%d')}")
    print(f"  Lockdown end (Phase 2): {lockdown_end.strftime('%Y-%m-%d')}")
    print(f"  Pre-lockdown days: {len(pre_lockdown)}")
    print(f"  During lockdown days: {len(during_lockdown)}")
    print(f"  After easing days: {len(after_easing)}")

    results = {'pre': {}, 'during_inter': {}, 'during_self': {},
               'after_inter': {}, 'after_self': {}}

    pre_lag = 2           # Small lag for short pre-lockdown period
    min_samples_pre = 5   # Lower threshold for pre-lockdown

    # --- Pre-lockdown: inter-city (lag=2) ---
    print(f"\n  === Pre-Lockdown inter-city (lag={pre_lag}, min_samples={min_samples_pre}) ===")
    for target in city_names:
        candidates = [c for c in city_names if c != target]
        X, y, feat_names, _ = build_lagged_inputs(
            pre_lockdown, target, candidates, n_lags=pre_lag
        )
        if len(y) < min_samples_pre or np.std(y) < 1e-10:
            results['pre'][target] = {'r2': 0.0, 'active_cities': []}
            print(f"    {target.title():25s}  SKIPPED (samples={len(y)})")
            continue

        n_inputs = X.shape[1]
        cgp = CartesianGeneticProgramming(n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
        genome, _, _ = cgp.evolve(X, y, n_generations=300, lam=LAMBDA, mutation_rate=MUTATION_RATE)
        r2 = cgp.compute_r2(genome, X, y)

        active_indices = cgp.get_active_inputs(genome)
        active_cities = set()
        for idx in active_indices:
            if idx < len(feat_names):
                active_cities.add(feat_names[idx].rsplit('_lag', 1)[0])

        results['pre'][target] = {'r2': r2, 'active_cities': list(active_cities)}
        print(f"    {target.title():25s}  R2={r2:.4f}  Active: {len(active_cities)}  (samples={len(y)})")

    # --- During lockdown: inter-city ---
    print(f"\n  === During Lockdown inter-city (lag={n_lags}) ===")
    for target in city_names:
        candidates = [c for c in city_names if c != target]
        X, y, feat_names, _ = build_lagged_inputs(
            during_lockdown, target, candidates, n_lags=n_lags
        )
        if len(y) < 15 or np.std(y) < 1e-10:
            results['during_inter'][target] = {'r2': 0.0, 'active_cities': []}
            continue

        n_inputs = X.shape[1]
        cgp = CartesianGeneticProgramming(n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
        genome, _, _ = cgp.evolve(X, y, n_generations=300, lam=LAMBDA, mutation_rate=MUTATION_RATE)
        r2 = cgp.compute_r2(genome, X, y)

        active_indices = cgp.get_active_inputs(genome)
        active_cities = set()
        for idx in active_indices:
            if idx < len(feat_names):
                active_cities.add(feat_names[idx].rsplit('_lag', 1)[0])

        results['during_inter'][target] = {'r2': r2, 'active_cities': list(active_cities)}
        print(f"    {target.title():25s}  R2={r2:.4f}  Active: {len(active_cities)}")

    # --- During lockdown: self-lag ---
    print(f"\n  === During Lockdown self-lag only ===")
    for target in city_names:
        X, y, _, _ = build_self_lagged_inputs(during_lockdown, target, n_lags=n_lags)
        if len(y) < 15 or np.std(y) < 1e-10:
            results['during_self'][target] = {'r2': 0.0}
            continue
        n_inputs = X.shape[1]
        cgp = CartesianGeneticProgramming(n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
        genome, _, _ = cgp.evolve(X, y, n_generations=300, lam=LAMBDA, mutation_rate=MUTATION_RATE)
        results['during_self'][target] = {'r2': cgp.compute_r2(genome, X, y)}
        print(f"    {target.title():25s}  R2={results['during_self'][target]['r2']:.4f}")

    # --- After easing: inter-city ---
    print(f"\n  === After Easing inter-city (lag={n_lags}) ===")
    for target in city_names:
        candidates = [c for c in city_names if c != target]
        X, y, feat_names, _ = build_lagged_inputs(
            after_easing, target, candidates, n_lags=n_lags
        )
        if len(y) < 15 or np.std(y) < 1e-10:
            results['after_inter'][target] = {'r2': 0.0, 'active_cities': []}
            continue

        n_inputs = X.shape[1]
        cgp = CartesianGeneticProgramming(n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
        genome, _, _ = cgp.evolve(X, y, n_generations=300, lam=LAMBDA, mutation_rate=MUTATION_RATE)
        r2 = cgp.compute_r2(genome, X, y)

        active_indices = cgp.get_active_inputs(genome)
        active_cities = set()
        for idx in active_indices:
            if idx < len(feat_names):
                active_cities.add(feat_names[idx].rsplit('_lag', 1)[0])

        results['after_inter'][target] = {'r2': r2, 'active_cities': list(active_cities)}
        print(f"    {target.title():25s}  R2={r2:.4f}  Active: {len(active_cities)}")

    # --- After easing: self-lag ---
    print(f"\n  === After Easing self-lag only ===")
    for target in city_names:
        X, y, _, _ = build_self_lagged_inputs(after_easing, target, n_lags=n_lags)
        if len(y) < 15 or np.std(y) < 1e-10:
            results['after_self'][target] = {'r2': 0.0}
            continue
        n_inputs = X.shape[1]
        cgp = CartesianGeneticProgramming(n_inputs, N_OUTPUTS, N_ROWS, N_COLS, LEVELS_BACK)
        genome, _, _ = cgp.evolve(X, y, n_generations=300, lam=LAMBDA, mutation_rate=MUTATION_RATE)
        results['after_self'][target] = {'r2': cgp.compute_r2(genome, X, y)}
        print(f"    {target.title():25s}  R2={results['after_self'][target]['r2']:.4f}")

    # For backward compatibility, also populate post_inter and post_self
    # (used by other functions) as the "after_easing" values
    results['post_inter'] = results['after_inter']
    results['post_self'] = results['after_self']

    return results


def plot_lockdown_comparison(lockdown_results, city_names, output_dir):
    """Plot lockdown period R2 comparison with 5 bars per city."""

    pre_r2 = [lockdown_results['pre'].get(c, {}).get('r2', 0) for c in city_names]
    during_inter_r2 = [lockdown_results.get('during_inter', {}).get(c, {}).get('r2', 0) for c in city_names]
    during_self_r2 = [lockdown_results.get('during_self', {}).get(c, {}).get('r2', 0) for c in city_names]
    after_inter_r2 = [lockdown_results.get('after_inter', {}).get(c, {}).get('r2', 0) for c in city_names]
    after_self_r2 = [lockdown_results.get('after_self', {}).get(c, {}).get('r2', 0) for c in city_names]

    x = np.arange(len(city_names))
    width = 0.15

    fig, ax = plt.subplots(figsize=(22, 9))
    ax.bar(x - 2*width, pre_r2, width, label='Pre-Lockdown (Inter-city)',
           color='#2196F3', alpha=0.85)
    ax.bar(x - width, during_inter_r2, width, label='During Lockdown (Inter-city)',
           color='#9C27B0', alpha=0.85)
    ax.bar(x, during_self_r2, width, label='During Lockdown (Self-lag)',
           color='#E91E63', alpha=0.85)
    ax.bar(x + width, after_inter_r2, width, label='After Easing (Inter-city)',
           color='#FF9800', alpha=0.85)
    ax.bar(x + 2*width, after_self_r2, width, label='After Easing (Self-lag)',
           color='#4CAF50', alpha=0.85)

    ax.axhline(y=R2_THRESHOLD, color='red', ls='--', lw=2, label=f'R2 threshold = {R2_THRESHOLD}')

    ax.set_xlabel('City', fontsize=12)
    ax.set_ylabel('R2', fontsize=12)
    ax.set_title('Lockdown Period Comparison (3 Periods)\n'
                 'Pre-lockdown (before Mar 9) | During lockdown (Mar 9 - May 18) | After easing (May 18+)\n'
                 'Blue = Pre-lockdown free mobility | Purple/Pink = Strict lockdown | Orange/Green = Reopening',
                 fontsize=13, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels([c.title() for c in city_names], rotation=45, ha='right', fontsize=9)
    ax.legend(fontsize=9, loc='upper left')
    ax.grid(True, alpha=0.3, axis='y')
    ax.set_ylim(0, 1.15)

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'pre_post_lockdown_comparison.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: pre_post_lockdown_comparison.png")

    # Per-region comparison
    fig, axes = plt.subplots(1, 3, figsize=(24, 8))
    for ax_idx, (region, cities) in enumerate(CITY_GROUPS.items()):
        ax = axes[ax_idx]
        x_reg = np.arange(len(cities))

        pre_v = [lockdown_results['pre'].get(c, {}).get('r2', 0) for c in cities]
        dur_i = [lockdown_results.get('during_inter', {}).get(c, {}).get('r2', 0) for c in cities]
        dur_s = [lockdown_results.get('during_self', {}).get(c, {}).get('r2', 0) for c in cities]
        aft_i = [lockdown_results.get('after_inter', {}).get(c, {}).get('r2', 0) for c in cities]
        aft_s = [lockdown_results.get('after_self', {}).get(c, {}).get('r2', 0) for c in cities]

        ax.bar(x_reg - 2*width, pre_v, width, label='Pre (Inter)', color='#2196F3', alpha=0.85)
        ax.bar(x_reg - width, dur_i, width, label='Lock (Inter)', color='#9C27B0', alpha=0.85)
        ax.bar(x_reg, dur_s, width, label='Lock (Self)', color='#E91E63', alpha=0.85)
        ax.bar(x_reg + width, aft_i, width, label='Ease (Inter)', color='#FF9800', alpha=0.85)
        ax.bar(x_reg + 2*width, aft_s, width, label='Ease (Self)', color='#4CAF50', alpha=0.85)

        ax.axhline(y=R2_THRESHOLD, color='red', ls='--', lw=1.5)
        ax.set_title(f'{region} Italy', fontsize=13, fontweight='bold')
        ax.set_xticks(x_reg)
        ax.set_xticklabels([c.title() for c in cities], rotation=45, ha='right', fontsize=8)
        ax.set_ylabel('R2')
        ax.set_ylim(0, 1.15)
        ax.legend(fontsize=7)
        ax.grid(True, alpha=0.3, axis='y')

    plt.suptitle('Lockdown Analysis by Region (3 Periods)', fontsize=15, fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'lockdown_by_region.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: lockdown_by_region.png")


def plot_lockdown_active_cities(lockdown_results, city_names, output_dir):
    """Plot number of active (influencing) cities before and after lockdown."""
    pre_active = [len(lockdown_results['pre'].get(c, {}).get('active_cities', []))
                  for c in city_names]
    post_active = [len(lockdown_results['post_inter'].get(c, {}).get('active_cities', []))
                   for c in city_names]

    x = np.arange(len(city_names))
    width = 0.35

    fig, ax = plt.subplots(figsize=(16, 7))
    ax.bar(x - width/2, pre_active, width, label='Pre-Lockdown', color='#2196F3', alpha=0.8)
    ax.bar(x + width/2, post_active, width, label='Post-Lockdown', color='#FF9800', alpha=0.8)

    ax.set_xlabel('City', fontsize=12)
    ax.set_ylabel('Number of Influencing Cities', fontsize=12)
    ax.set_title('Number of Active Input Cities (CGP)\n'
                 'Pre vs Post Lockdown - How many cities influence each target?',
                 fontsize=14, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels([c.title() for c in city_names], rotation=45, ha='right', fontsize=9)
    ax.legend(fontsize=11)
    ax.grid(True, alpha=0.3, axis='y')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'lockdown_active_cities.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: lockdown_active_cities.png")


# ==============================================================================
# OTHER VISUALIZATION
# ==============================================================================

def plot_connection_heatmap(connections, city_names, output_dir):
    """Heatmap of CGP-discovered connection strengths."""
    n = len(city_names)
    matrix = np.zeros((n, n))
    for i, c1 in enumerate(city_names):
        for j, c2 in enumerate(city_names):
            pair = tuple(sorted([c1, c2]))
            if pair in connections:
                matrix[i, j] = connections[pair]

    fig, ax = plt.subplots(figsize=(14, 12))
    im = ax.imshow(matrix, cmap='YlOrRd', vmin=0, vmax=1, aspect='auto')
    ax.set_xticks(range(n))
    ax.set_yticks(range(n))
    ax.set_xticklabels([c.title() for c in city_names], rotation=45, ha='right', fontsize=9)
    ax.set_yticklabels([c.title() for c in city_names], fontsize=9)

    # Add values
    for i in range(n):
        for j in range(n):
            if matrix[i, j] > 0.1:
                ax.text(j, i, f'{matrix[i, j]:.2f}', ha='center', va='center', fontsize=7)

    ax.set_title('CGP Connection Strength Heatmap (R²)\n15 Selected Cities (2020v2) - Addition-Only Functions',
                 fontsize=13, fontweight='bold')
    plt.colorbar(im, ax=ax, label='Connection Strength (R²)', shrink=0.7)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'connection_heatmap.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: connection_heatmap.png")


def plot_network_graph(connections, city_names, output_dir):
    """Plot network graph of city connections."""
    try:
        import networkx as nx
    except ImportError:
        print("  networkx not available, skipping network graph")
        return

    G = nx.Graph()
    for city in city_names:
        region = CITY_REGION_MAP.get(city, 'Unknown')
        G.add_node(city, region=region)

    for (c1, c2), weight in connections.items():
        if weight > R2_THRESHOLD:
            G.add_edge(c1, c2, weight=weight)

    isolated = list(nx.isolates(G))
    # Don't remove isolated - show them grayed out

    if len(G.edges()) == 0:
        print("  No significant connections (R²>0.9) for network graph")
        # Still save an empty graph to show the threshold is strict
        fig, ax = plt.subplots(figsize=(12, 10))
        pos = nx.spring_layout(G, k=2, seed=42)
        nx.draw_networkx_nodes(G, pos, ax=ax, node_size=500, node_color='lightgray',
                               edgecolors='black', linewidths=1)
        nx.draw_networkx_labels(G, pos, ax=ax, font_size=8, font_weight='bold')
        ax.set_title(f'CGP Network (R² > {R2_THRESHOLD})\nNo connections meet the strict threshold',
                     fontsize=14, fontweight='bold')
        ax.axis('off')
        plt.tight_layout()
        plt.savefig(os.path.join(output_dir, 'network_graph.png'), dpi=150, bbox_inches='tight')
        plt.close()
        print("  Saved: network_graph.png (no significant edges)")
        return

    pos = nx.spring_layout(G, k=2, iterations=50, seed=42)

    # Color by region
    node_colors = [REGION_COLORS.get(CITY_REGION_MAP.get(n, ''), 'gray') for n in G.nodes()]
    degrees = dict(G.degree())
    node_sizes = [max(degrees.get(n, 0) * 200, 300) for n in G.nodes()]

    edge_weights = [G[u][v]['weight'] for u, v in G.edges()]
    max_w = max(edge_weights) if edge_weights else 1
    edge_widths = [w / max_w * 4 for w in edge_weights]

    fig, ax = plt.subplots(figsize=(16, 14))
    nx.draw_networkx_edges(G, pos, ax=ax, edge_color=edge_weights,
                           edge_cmap=plt.cm.Reds, width=edge_widths, alpha=0.6)
    nx.draw_networkx_nodes(G, pos, ax=ax, node_size=node_sizes,
                           node_color=node_colors, edgecolors='black', linewidths=1)
    nx.draw_networkx_labels(G, pos, ax=ax, font_size=8, font_weight='bold')

    patches = [mpatches.Patch(color=c, label=r) for r, c in REGION_COLORS.items()]
    ax.legend(handles=patches, loc='upper left', fontsize=11, title='Region')

    ax.set_title(f'CGP-Discovered City Network (R² > {R2_THRESHOLD})\n'
                 f'15 Selected Cities - Addition-Only Functions',
                 fontsize=14, fontweight='bold')
    ax.axis('off')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'network_graph.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: network_graph.png")


def plot_cgp_weights(cgp_results, city_names, output_dir):
    """Plot the CGP-discovered weights (number of times each city appears as active input)."""
    # Build weight matrix: how often city_j influences city_i
    n = len(city_names)
    weight_matrix = np.zeros((n, n))

    for i, target in enumerate(city_names):
        if target in cgp_results:
            for source in cgp_results[target]['active_cities']:
                if source in city_names:
                    j = city_names.index(source)
                    weight_matrix[i, j] = cgp_results[target]['r2']

    fig, ax = plt.subplots(figsize=(14, 12))
    im = ax.imshow(weight_matrix, cmap='Blues', vmin=0, vmax=1, aspect='auto')
    ax.set_xticks(range(n))
    ax.set_yticks(range(n))
    ax.set_xticklabels([c.title() for c in city_names], rotation=45, ha='right', fontsize=9)
    ax.set_yticklabels([c.title() for c in city_names], fontsize=9)

    for i in range(n):
        for j in range(n):
            if weight_matrix[i, j] > 0.01:
                ax.text(j, i, f'{weight_matrix[i, j]:.2f}', ha='center', va='center', fontsize=7)

    ax.set_xlabel('Source City (Influencer)', fontsize=12)
    ax.set_ylabel('Target City (Influenced)', fontsize=12)
    ax.set_title('CGP-Discovered Influence Weights\n'
                 'Which cities influence which? (R² metric)',
                 fontsize=14, fontweight='bold')
    plt.colorbar(im, ax=ax, label='Influence Weight (R²)', shrink=0.7)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'cgp_weights_heatmap.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: cgp_weights_heatmap.png")


def plot_fitness_convergence(cgp_results, output_dir):
    """Plot CGP fitness convergence for all cities."""
    cities_with_history = [c for c in cgp_results if cgp_results[c].get('history')]
    n = len(cities_with_history)
    if n == 0:
        return

    cols = 5
    rows = (n + cols - 1) // cols
    fig, axes = plt.subplots(rows, cols, figsize=(25, 5 * rows))
    fig.suptitle('CGP Fitness Convergence - 15 Cities (Addition-Only)', fontsize=16, fontweight='bold')
    axes_flat = np.array(axes).flatten()

    for idx, city in enumerate(cities_with_history):
        ax = axes_flat[idx]
        region = CITY_REGION_MAP.get(city, 'Unknown')
        color = REGION_COLORS.get(region, 'gray')
        ax.plot(cgp_results[city]['history'], color=color, lw=1)
        ax.set_title(f'{city.title()} ({region})\nR²={cgp_results[city]["r2"]:.4f}', fontsize=9)
        ax.set_xlabel('Generation')
        ax.set_ylabel('MSE')
        ax.grid(True, alpha=0.3)
        try:
            ax.set_yscale('log')
        except:
            pass

    for idx in range(n, len(axes_flat)):
        axes_flat[idx].set_visible(False)

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'cgp_fitness_convergence.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: cgp_fitness_convergence.png")


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    print("=" * 70)
    print("   CGP Analysis - 15 Selected Italian Cities (2020v2)")
    print("   Addition-Only Functions | R² threshold = 0.9")
    print("   Lag Sweep | Pre/Post Lockdown Analysis")
    print("=" * 70)
    print(f"   North: {', '.join([c.title() for c in CITY_GROUPS['North']])}")
    print(f"   Center: {', '.join([c.title() for c in CITY_GROUPS['Center']])}")
    print(f"   South: {', '.join([c.title() for c in CITY_GROUPS['South']])}")
    print(f"   Lockdown date: {LOCKDOWN_DATE.strftime('%Y-%m-%d')}")
    print(f"   Functions: {[f[0] for f in CartesianGeneticProgramming.FUNCTIONS]}")
    print("=" * 70)

    # 1. Load data
    print("\n[1/7] Loading data...")
    daily_new, city_names = prepare_data(DATA_PATH, set(ALL_CITIES))
    # Reorder to match our defined order
    city_names = [c for c in ALL_CITIES if c in city_names]
    print(f"  Loaded {len(city_names)} cities, {len(daily_new)} days")
    print(f"  Date range: {daily_new.index[0]} to {daily_new.index[-1]}")

    # 2. Main CGP analysis (inter-city, default lag)
    print(f"\n[2/7] Running main CGP analysis (lag={DEFAULT_N_LAGS})...")
    cgp_results = analyze_inter_city(daily_new, city_names, n_lags=DEFAULT_N_LAGS)

    # Build connections dict
    all_connections = {}
    for target, result in cgp_results.items():
        for source in result['active_cities']:
            pair = tuple(sorted([target, source]))
            current = all_connections.get(pair, 0)
            all_connections[pair] = max(current, result['r2'])

    # 3. Lag sweep analysis
    print(f"\n[3/7] Running lag sweep analysis...")
    lag_r2_results = run_lag_sweep(daily_new, city_names, LAG_SWEEP_VALUES)
    plot_lag_vs_r2(lag_r2_results, LAG_SWEEP_VALUES, OUTPUT_DIR)

    # 4. Pre/Post lockdown analysis
    print(f"\n[4/7] Running pre/post lockdown analysis...")
    lockdown_results = run_lockdown_analysis(daily_new, city_names, LOCKDOWN_DATE, n_lags=DEFAULT_N_LAGS)

    # 5. Generate visualizations
    print(f"\n[5/7] Generating visualizations...")
    plot_connection_heatmap(all_connections, city_names, OUTPUT_DIR)
    plot_network_graph(all_connections, city_names, OUTPUT_DIR)
    plot_cgp_weights(cgp_results, city_names, OUTPUT_DIR)
    plot_fitness_convergence(cgp_results, OUTPUT_DIR)

    # 6. Lockdown visualizations
    print(f"\n[6/7] Generating lockdown comparison plots...")
    plot_lockdown_comparison(lockdown_results, city_names, OUTPUT_DIR)
    plot_lockdown_active_cities(lockdown_results, city_names, OUTPUT_DIR)

    # 7. Save summary data
    print(f"\n[7/7] Saving summary data...")

    # Connections CSV
    conn_rows = []
    for (c1, c2), strength in sorted(all_connections.items(), key=lambda x: x[1], reverse=True):
        conn_rows.append({
            'city_1': c1.title(),
            'city_2': c2.title(),
            'region_1': CITY_REGION_MAP.get(c1, ''),
            'region_2': CITY_REGION_MAP.get(c2, ''),
            'connection_strength_R2': round(strength, 4),
            'is_significant': strength > R2_THRESHOLD,
        })
    conn_df = pd.DataFrame(conn_rows)
    conn_df.to_csv(os.path.join(OUTPUT_DIR, 'all_connections.csv'), index=False)
    print(f"  Saved: all_connections.csv ({len(conn_df)} pairs)")

    # City summary CSV
    summary_rows = []
    for city in city_names:
        r2 = cgp_results.get(city, {}).get('r2', 0)
        active = cgp_results.get(city, {}).get('active_cities', [])
        funcs = cgp_results.get(city, {}).get('active_functions', [])
        summary_rows.append({
            'city': city.title(),
            'region': CITY_REGION_MAP.get(city, ''),
            'cgp_r2': round(r2, 4),
            'n_active_inputs': len(active),
            'active_cities': '; '.join([c.title() for c in active]),
            'active_functions': '; '.join([f[1] for f in funcs]),
            'is_significant': r2 > R2_THRESHOLD,
        })
    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_csv(os.path.join(OUTPUT_DIR, 'city_summary.csv'), index=False)
    print(f"  Saved: city_summary.csv")

    # Lag sweep CSV
    lag_rows = []
    for city in city_names:
        row = {'city': city.title(), 'region': CITY_REGION_MAP.get(city, '')}
        for lag in LAG_SWEEP_VALUES:
            row[f'lag_{lag}'] = round(lag_r2_results.get(city, {}).get(lag, 0), 4)
        lag_rows.append(row)
    lag_df = pd.DataFrame(lag_rows)
    lag_df.to_csv(os.path.join(OUTPUT_DIR, 'lag_sweep_results.csv'), index=False)
    print(f"  Saved: lag_sweep_results.csv")

    # Lockdown comparison CSV
    lockdown_rows = []
    for city in city_names:
        lockdown_rows.append({
            'city': city.title(),
            'region': CITY_REGION_MAP.get(city, ''),
            'pre_lockdown_r2': round(lockdown_results['pre'].get(city, {}).get('r2', 0), 4),
            'during_lockdown_inter_r2': round(lockdown_results.get('during_inter', {}).get(city, {}).get('r2', 0), 4),
            'during_lockdown_self_r2': round(lockdown_results.get('during_self', {}).get(city, {}).get('r2', 0), 4),
            'after_easing_inter_r2': round(lockdown_results.get('after_inter', {}).get(city, {}).get('r2', 0), 4),
            'after_easing_self_r2': round(lockdown_results.get('after_self', {}).get(city, {}).get('r2', 0), 4),
            'pre_n_active_cities': len(lockdown_results['pre'].get(city, {}).get('active_cities', [])),
            'during_n_active_cities': len(lockdown_results.get('during_inter', {}).get(city, {}).get('active_cities', [])),
            'after_n_active_cities': len(lockdown_results.get('after_inter', {}).get(city, {}).get('active_cities', [])),
        })
    lockdown_df = pd.DataFrame(lockdown_rows)
    lockdown_df.to_csv(os.path.join(OUTPUT_DIR, 'lockdown_comparison.csv'), index=False)
    print(f"  Saved: lockdown_comparison.csv")

    # Print summary
    n_significant = sum(1 for v in all_connections.values() if v > R2_THRESHOLD)
    n_total = len(all_connections)

    print("\n" + "=" * 70)
    print("   CGP RESULTS SUMMARY (2020v2)")
    print("=" * 70)
    print(f"  Function set: {[f[0] for f in CartesianGeneticProgramming.FUNCTIONS]}")
    print(f"  R2 threshold: {R2_THRESHOLD}")
    print(f"  Levels back: {LEVELS_BACK}")
    print(f"  Total pairs analyzed: {n_total}")
    print(f"  Significant connections (R2>{R2_THRESHOLD}): {n_significant}")

    print(f"\n  CGP Architecture:")
    print(f"    Rows: {N_ROWS}, Cols: {N_COLS}, Total nodes: {N_ROWS * N_COLS}")
    print(f"    Generations: {N_GENERATIONS}, lambda: {LAMBDA}, Mutation: {MUTATION_RATE}")
    print(f"    Levels back: {LEVELS_BACK} (node in column c can connect to")
    print(f"                 any node in columns max(0, c-{LEVELS_BACK}) to c-1)")

    print(f"\n  LOCKDOWN ANALYSIS:")
    print(f"    Pre-lockdown: before {LOCKDOWN_DATE.strftime('%Y-%m-%d')} (lag=2)")
    print(f"    During lockdown: {LOCKDOWN_DATE.strftime('%Y-%m-%d')} to 2020-05-18")
    print(f"    After easing: 2020-05-18+")
    print(f"  {'City':25s}  {'Pre':>6s}  {'Dur(I)':>6s}  {'Dur(S)':>6s}  {'Aft(I)':>6s}  {'Aft(S)':>6s}")
    print(f"  {'-'*67}")
    for city in city_names:
        pre = lockdown_results['pre'].get(city, {}).get('r2', 0)
        dur_i = lockdown_results.get('during_inter', {}).get(city, {}).get('r2', 0)
        dur_s = lockdown_results.get('during_self', {}).get(city, {}).get('r2', 0)
        aft_i = lockdown_results.get('after_inter', {}).get(city, {}).get('r2', 0)
        aft_s = lockdown_results.get('after_self', {}).get(city, {}).get('r2', 0)
        print(f"    {city.title():25s}  {pre:6.3f}  {dur_i:6.3f}  {dur_s:6.3f}  {aft_i:6.3f}  {aft_s:6.3f}")

    print(f"\n  Output: {OUTPUT_DIR}")
    print("=" * 70)


if __name__ == '__main__':
    main()
