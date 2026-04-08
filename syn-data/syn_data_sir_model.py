"""
Synthetic SIR Model - Data Generation for 10 Italian Cities
=============================================================
Generates synthetic COVID-19 infection data using the SIR compartmental
model with KNOWN inter-city connections (Ground Truth).

Parameters:
    - gamma (γ) = 0.14 (recovery rate, ~7 days recovery period)
    - beta (β)  = 0.35 (infection rate)
    - R0 = β/γ = 2.5

The purpose is to validate that CGP can correctly discover these
known connections from the generated data.

Three spillover percentages are tested: 1%, 5%, 10%
"""

import os
import sys
import numpy as np
import pandas as pd
from scipy.integrate import odeint
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# ==============================================================================
# CONFIGURATION
# ==============================================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, 'sir_results')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# SIR Parameters (as specified)
GAMMA = 0.14   # Recovery rate (~7 day recovery period)
BETA = 0.35    # Infection rate
R0 = BETA / GAMMA  # Basic reproduction number = 2.5
RECOVERY_PERIOD = int(1 / GAMMA)  # ~7 days

# Simulation parameters
N_DAYS = 365
START_DATE = datetime(2022, 1, 1)

# 10 Italian Cities with known populations (ISTAT data)
CITIES = {
    'milano':  {'population': 3_250_000, 'name_fa': 'میلان',    'region': 'Lombardia'},
    'roma':    {'population': 4_350_000, 'name_fa': 'رم',       'region': 'Lazio'},
    'napoli':  {'population': 3_100_000, 'name_fa': 'ناپل',     'region': 'Campania'},
    'torino':  {'population': 2_250_000, 'name_fa': 'تورین',    'region': 'Piemonte'},
    'palermo': {'population': 1_270_000, 'name_fa': 'پالرمو',   'region': 'Sicilia'},
    'genova':  {'population':   850_000, 'name_fa': 'جنوا',     'region': 'Liguria'},
    'bologna': {'population': 1_015_000, 'name_fa': 'بولونیا',  'region': 'Emilia-Romagna'},
    'firenze': {'population': 1_010_000, 'name_fa': 'فلورانس',  'region': 'Toscana'},
    'bari':    {'population': 1_260_000, 'name_fa': 'باری',     'region': 'Puglia'},
    'catania': {'population': 1_115_000, 'name_fa': 'کاتانیا',  'region': 'Sicilia'},
}

CITY_NAMES = list(CITIES.keys())

# GROUND TRUTH: Known connections between cities
# Format: (source_city, target_city, strength, lag_days)
# strength: 'strong' or 'medium' (affects spillover multiplier)
GROUND_TRUTH_CONNECTIONS = [
    ('milano',  'torino',  'strong', 3),   # Geographic proximity (NW Italy)
    ('milano',  'genova',  'strong', 4),   # Lombardia-Liguria corridor
    ('milano',  'bologna', 'medium', 5),   # Northern hub connections
    ('roma',    'napoli',  'strong', 3),   # Central-South corridor
    ('roma',    'firenze', 'medium', 4),   # Central Italy
    ('napoli',  'bari',    'strong', 4),   # Southern Italy
    ('napoli',  'palermo', 'medium', 5),   # South-Sicily sea route
    ('bologna', 'firenze', 'strong', 3),   # Emilia-Toscana border
    ('catania', 'palermo', 'strong', 2),   # Within Sicily
    ('bari',    'catania', 'medium', 5),   # Southeast-Sicily corridor
]

# Spillover percentages to test
SPILLOVER_PERCENTAGES = [0.01, 0.05, 0.10]  # 1%, 5%, 10%

# Strength multipliers
STRENGTH_MULTIPLIER = {
    'strong': 1.0,
    'medium': 0.6,
}

# Per-city SIR variation (slightly different initial conditions and beta noise)
CITY_VARIATION = {
    'milano':  {'I0_factor': 1.2, 'beta_noise': 0.02, 'start_day': 0},
    'roma':    {'I0_factor': 1.0, 'beta_noise': 0.01, 'start_day': 5},
    'napoli':  {'I0_factor': 0.8, 'beta_noise': 0.03, 'start_day': 8},
    'torino':  {'I0_factor': 0.9, 'beta_noise': 0.02, 'start_day': 3},
    'palermo': {'I0_factor': 0.6, 'beta_noise': 0.04, 'start_day': 12},
    'genova':  {'I0_factor': 0.7, 'beta_noise': 0.02, 'start_day': 4},
    'bologna': {'I0_factor': 0.8, 'beta_noise': 0.01, 'start_day': 6},
    'firenze': {'I0_factor': 0.7, 'beta_noise': 0.02, 'start_day': 7},
    'bari':    {'I0_factor': 0.5, 'beta_noise': 0.03, 'start_day': 10},
    'catania': {'I0_factor': 0.4, 'beta_noise': 0.04, 'start_day': 14},
}


# ==============================================================================
# SIR MODEL
# ==============================================================================

def sir_ode(y, t, beta, gamma, N):
    """SIR differential equations."""
    S, I, R = y
    dSdt = -beta * S * I / N
    dIdt = beta * S * I / N - gamma * I
    dRdt = gamma * I
    return [dSdt, dIdt, dRdt]


def simulate_sir_city(city, N_days=365, seed=None):
    """
    Simulate SIR model for a single city (no cross-city influence).
    Returns daily S, I, R, and daily new infections.
    """
    if seed is not None:
        np.random.seed(seed)

    N = CITIES[city]['population']
    var = CITY_VARIATION[city]

    # Initial conditions
    I0 = max(int(N * 0.0001 * var['I0_factor']), 10)  # Initial infected
    R0_init = 0
    S0 = N - I0 - R0_init

    # Time-varying beta with seasonal wave and noise
    t = np.arange(N_days)
    # Base seasonal modulation: winter peak, summer dip
    seasonal = 1.0 + 0.3 * np.cos(2 * np.pi * (t - 30) / 365)
    # Random daily noise
    noise = 1.0 + var['beta_noise'] * np.random.randn(N_days)
    noise = np.clip(noise, 0.8, 1.2)

    beta_t = BETA * seasonal * noise

    # Simulate day by day with time-varying beta
    S_arr = np.zeros(N_days)
    I_arr = np.zeros(N_days)
    R_arr = np.zeros(N_days)
    new_infections = np.zeros(N_days)

    S, I, R_val = float(S0), float(I0), float(R0_init)

    for day in range(N_days):
        S_arr[day] = S
        I_arr[day] = I
        R_arr[day] = R_val

        beta_day = beta_t[day]
        new_inf = beta_day * S * I / N
        new_rec = GAMMA * I

        new_infections[day] = max(new_inf, 0)

        S = max(S - new_inf, 0)
        I = max(I + new_inf - new_rec, 0)
        R_val = R_val + new_rec

    return {
        'S': S_arr,
        'I': I_arr,
        'R': R_arr,
        'new_infections': new_infections,
        'beta_t': beta_t,
        'cumulative': np.cumsum(new_infections),
    }


def generate_synthetic_data(spillover_pct, seed=42):
    """
    Generate synthetic SIR data with inter-city spillover effects.

    spillover_pct: fraction of source city's daily infections that
                   spill over to target city (with time lag).
    """
    np.random.seed(seed)

    # Step 1: Generate independent SIR for each city
    city_data = {}
    for city in CITY_NAMES:
        city_seed = seed + hash(city) % 1000
        city_data[city] = simulate_sir_city(city, N_DAYS, seed=city_seed)

    # Step 2: Apply spillover effects based on ground truth connections
    # We modify the new_infections to include cross-city effects
    final_new_infections = {}
    for city in CITY_NAMES:
        final_new_infections[city] = city_data[city]['new_infections'].copy()

    # Apply spillover: for each connection, source city's infections
    # contribute to target city's infections with a lag
    for source, target, strength, lag in GROUND_TRUTH_CONNECTIONS:
        multiplier = STRENGTH_MULTIPLIER[strength]
        spillover = spillover_pct * multiplier

        source_infections = city_data[source]['new_infections']

        for day in range(lag, N_DAYS):
            # Spillover from source (lagged) to target
            spill_amount = source_infections[day - lag] * spillover
            final_new_infections[target][day] += spill_amount

        # Also add reverse effect (smaller) for bidirectional influence
        target_infections = city_data[target]['new_infections']
        reverse_spillover = spillover * 0.3  # Reverse effect is weaker
        for day in range(lag, N_DAYS):
            spill_amount = target_infections[day - lag] * reverse_spillover
            final_new_infections[source][day] += spill_amount

    # Step 3: Add observation noise
    for city in CITY_NAMES:
        noise = np.random.normal(0, 0.02 * np.mean(final_new_infections[city]),
                                  N_DAYS)
        final_new_infections[city] = np.maximum(
            final_new_infections[city] + noise, 0
        )

    # Step 4: Compute cumulative infections
    final_cumulative = {}
    for city in CITY_NAMES:
        final_cumulative[city] = np.cumsum(final_new_infections[city])

    # Build dates
    dates = [START_DATE + timedelta(days=i) for i in range(N_DAYS)]

    return {
        'dates': dates,
        'city_data': city_data,
        'new_infections': final_new_infections,
        'cumulative': final_cumulative,
    }


def save_synthetic_csv(data, spillover_pct, output_dir):
    """Save synthetic data to CSV in the same format as real data."""
    pct_label = f"{int(spillover_pct * 100)}pct"
    df = pd.DataFrame({'data': data['dates']})

    for city in CITY_NAMES:
        df[f'{city}_infected'] = data['cumulative'][city]

    filepath = os.path.join(output_dir, f'synthetic_data_{pct_label}.csv')
    df.to_csv(filepath, index=False)
    print(f"  Saved: {filepath}")
    return filepath


# ==============================================================================
# VISUALIZATION
# ==============================================================================

def plot_sir_compartments(city, data, output_dir):
    """Plot S, I, R compartments for a single city."""
    dates = data['dates']
    cd = data['city_data'][city]

    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle(f'SIR Model - {city.title()} (Synthetic Data)\n'
                 f'beta={BETA}, gamma={GAMMA}, R0={R0:.1f}, '
                 f'N={CITIES[city]["population"]:,}',
                 fontsize=14, fontweight='bold')

    # 1. SIR Compartments
    ax = axes[0, 0]
    ax.plot(dates, cd['S'], 'g-', lw=2, label='S (Susceptible)')
    ax.plot(dates, cd['I'], 'r-', lw=2, label='I (Infected)')
    ax.plot(dates, cd['R'], 'b-', lw=2, label='R (Recovered)')
    ax.set_title('SIR Compartments')
    ax.set_ylabel('Population')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))

    # 2. Daily New Infections
    ax = axes[0, 1]
    new_inf = data['new_infections'][city]
    ax.bar(dates, new_inf, color='salmon', alpha=0.6, width=1.0,
           label='Daily New Cases')
    rolling = pd.Series(new_inf).rolling(7, min_periods=1).mean()
    ax.plot(dates, rolling, 'r-', lw=2, label='7-day Average')
    ax.set_title('Daily New Infections')
    ax.set_ylabel('New Cases')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))

    # 3. Cumulative Infections
    ax = axes[1, 0]
    ax.plot(dates, data['cumulative'][city], 'b-', lw=2,
            label='Cumulative Infections')
    ax.set_title('Cumulative Infections')
    ax.set_ylabel('Total Cases')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))

    # 4. Beta(t)
    ax = axes[1, 1]
    ax.plot(dates, cd['beta_t'], 'purple', lw=1.5, alpha=0.7,
            label='beta(t) time-varying')
    ax.axhline(y=BETA, color='red', ls='--', lw=2, label=f'Base beta = {BETA}')
    ax.axhline(y=GAMMA, color='green', ls='--', lw=1.5,
               label=f'gamma = {GAMMA}')
    ax.set_title('Infection Rate beta(t)')
    ax.set_ylabel('Beta')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, f'sir_{city}.png'),
                dpi=150, bbox_inches='tight')
    plt.close()


def plot_all_cities_comparison(data, output_dir):
    """Compare all 10 cities in a single plot."""
    dates = data['dates']
    colors = plt.cm.tab10(np.linspace(0, 1, 10))

    fig, axes = plt.subplots(2, 2, figsize=(18, 14))
    fig.suptitle(f'Synthetic SIR Model - All 10 Cities Comparison\n'
                 f'beta={BETA}, gamma={GAMMA}, R0={R0:.1f}',
                 fontsize=16, fontweight='bold')

    # 1. All cumulative curves
    ax = axes[0, 0]
    for i, city in enumerate(CITY_NAMES):
        ax.plot(dates, data['cumulative'][city], color=colors[i],
                lw=2, label=city.title())
    ax.set_title('Cumulative Infections')
    ax.set_ylabel('Total Cases')
    ax.legend(fontsize=8, ncol=2)
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    # 2. Daily new per 100k
    ax = axes[0, 1]
    for i, city in enumerate(CITY_NAMES):
        N = CITIES[city]['population']
        daily_norm = data['new_infections'][city] / N * 100000
        daily_smooth = pd.Series(daily_norm).rolling(7, min_periods=1).mean()
        ax.plot(dates, daily_smooth, color=colors[i], lw=2,
                label=city.title())
    ax.set_title('Daily New Cases per 100k (7-day avg)')
    ax.set_ylabel('Cases per 100k')
    ax.legend(fontsize=8, ncol=2)
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    # 3. Beta(t) for all cities
    ax = axes[1, 0]
    for i, city in enumerate(CITY_NAMES):
        beta_t = data['city_data'][city]['beta_t']
        ax.plot(dates, beta_t, color=colors[i], lw=1.5,
                label=city.title(), alpha=0.8)
    ax.axhline(y=GAMMA, color='black', ls='--', lw=1,
               label=f'gamma={GAMMA}')
    ax.set_title('Time-varying beta(t)')
    ax.set_ylabel('Beta')
    ax.legend(fontsize=7, ncol=2)
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    # 4. R0(t) = beta/gamma
    ax = axes[1, 1]
    for i, city in enumerate(CITY_NAMES):
        R0_t = data['city_data'][city]['beta_t'] / GAMMA
        R0_smooth = pd.Series(R0_t).rolling(14, min_periods=1).mean()
        ax.plot(dates, R0_smooth, color=colors[i], lw=2,
                label=city.title())
    ax.axhline(y=1.0, color='black', ls='--', lw=2,
               label='R0 = 1 (threshold)')
    ax.set_title('Reproduction Number R0(t) = beta/gamma')
    ax.set_ylabel('R0')
    ax.legend(fontsize=7, ncol=2)
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'comparison_all_cities.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: comparison_all_cities.png")


def plot_ground_truth_network(output_dir):
    """Visualize the ground truth connection network."""
    try:
        import networkx as nx
    except ImportError:
        print("  networkx not available, skipping ground truth network plot")
        return

    G = nx.DiGraph()
    for city in CITY_NAMES:
        G.add_node(city, population=CITIES[city]['population'])

    for source, target, strength, lag in GROUND_TRUTH_CONNECTIONS:
        G.add_edge(source, target, strength=strength, lag=lag,
                    weight=1.0 if strength == 'strong' else 0.6)

    pos = nx.spring_layout(G, k=2.5, iterations=100, seed=42)

    fig, ax = plt.subplots(figsize=(14, 11))

    # Node sizes by population
    pops = [CITIES[n]['population'] for n in G.nodes()]
    max_pop = max(pops)
    node_sizes = [p / max_pop * 3000 + 500 for p in pops]

    # Draw edges
    strong_edges = [(u, v) for u, v, d in G.edges(data=True)
                     if d['strength'] == 'strong']
    medium_edges = [(u, v) for u, v, d in G.edges(data=True)
                     if d['strength'] == 'medium']

    nx.draw_networkx_edges(G, pos, edgelist=strong_edges, ax=ax,
                            edge_color='#e74c3c', width=3,
                            arrows=True, arrowsize=20,
                            connectionstyle='arc3,rad=0.1',
                            alpha=0.8)
    nx.draw_networkx_edges(G, pos, edgelist=medium_edges, ax=ax,
                            edge_color='#f39c12', width=2,
                            arrows=True, arrowsize=15,
                            connectionstyle='arc3,rad=0.1',
                            alpha=0.6, style='dashed')

    # Draw nodes
    nx.draw_networkx_nodes(G, pos, ax=ax, node_size=node_sizes,
                            node_color='#3498db', edgecolors='#2c3e50',
                            linewidths=2, alpha=0.9)

    # Labels
    labels = {city: f"{city.title()}\n({CITIES[city]['population']:,})"
              for city in CITY_NAMES}
    nx.draw_networkx_labels(G, pos, labels=labels, ax=ax, font_size=8,
                             font_weight='bold')

    # Edge labels (lag days)
    edge_labels = {(u, v): f"lag={d['lag']}d"
                   for u, v, d in G.edges(data=True)}
    nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels, ax=ax,
                                  font_size=7, font_color='#2c3e50')

    # Legend
    from matplotlib.lines import Line2D
    legend_elements = [
        Line2D([0], [0], color='#e74c3c', lw=3, label='Strong connection'),
        Line2D([0], [0], color='#f39c12', lw=2, ls='--',
               label='Medium connection'),
    ]
    ax.legend(handles=legend_elements, loc='upper left', fontsize=11)

    ax.set_title('Ground Truth: Known City Connections\n'
                 '(These are the connections CGP should discover)',
                 fontsize=14, fontweight='bold')
    ax.axis('off')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'ground_truth_network.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: ground_truth_network.png")


def plot_spillover_comparison(all_data, output_dir):
    """Compare infection curves across different spillover percentages."""
    fig, axes = plt.subplots(2, 5, figsize=(25, 10))
    fig.suptitle('Effect of Spillover Percentage on City Infections\n'
                 '(1% vs 5% vs 10%)', fontsize=16, fontweight='bold')

    colors = {'1%': '#3498db', '5%': '#e74c3c', '10%': '#2ecc71'}

    for idx, city in enumerate(CITY_NAMES):
        row = idx // 5
        col = idx % 5
        ax = axes[row, col]

        for pct_label, data in all_data.items():
            daily = data['new_infections'][city]
            smooth = pd.Series(daily).rolling(7, min_periods=1).mean()
            ax.plot(smooth.values, color=colors[pct_label], lw=1.5,
                    label=pct_label, alpha=0.8)

        ax.set_title(f'{city.title()}', fontsize=10, fontweight='bold')
        ax.grid(True, alpha=0.3)
        if idx == 0:
            ax.legend(fontsize=8)
        ax.set_xlabel('Days')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'spillover_comparison.png'),
                dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: spillover_comparison.png")


# ==============================================================================
# MAIN EXECUTION
# ==============================================================================

def main():
    print("=" * 70)
    print("   Synthetic SIR Model - Data Generation")
    print(f"   beta = {BETA}, gamma = {GAMMA}, R0 = {R0:.1f}")
    print(f"   Cities: {len(CITY_NAMES)}")
    print(f"   Simulation: {N_DAYS} days from {START_DATE.strftime('%Y-%m-%d')}")
    print("=" * 70)

    # 1. Generate data for each spillover percentage
    all_data = {}
    for pct in SPILLOVER_PERCENTAGES:
        pct_label = f"{int(pct * 100)}%"
        print(f"\n[1/4] Generating synthetic data with {pct_label} spillover...")
        data = generate_synthetic_data(pct, seed=42)
        all_data[pct_label] = data

        # Save CSV
        save_synthetic_csv(data, pct, OUTPUT_DIR)

    # Use the 5% data as the "default" for individual plots
    default_data = all_data['5%']

    # 2. Plot individual city SIR
    print("\n[2/4] Generating individual city plots...")
    for city in CITY_NAMES:
        plot_sir_compartments(city, default_data, OUTPUT_DIR)
        print(f"  Saved: sir_{city}.png")

    # 3. Comparison plots
    print("\n[3/4] Generating comparison plots...")
    plot_all_cities_comparison(default_data, OUTPUT_DIR)
    plot_ground_truth_network(OUTPUT_DIR)
    plot_spillover_comparison(all_data, OUTPUT_DIR)

    # 4. Save summary
    print("\n[4/4] Saving summary data...")

    # Summary CSV
    summary_rows = []
    for city in CITY_NAMES:
        N = CITIES[city]['population']
        total_cases = default_data['cumulative'][city][-1]
        mean_beta = np.mean(default_data['city_data'][city]['beta_t'])
        mean_R0_val = mean_beta / GAMMA
        summary_rows.append({
            'city': city.title(),
            'region': CITIES[city]['region'],
            'population': N,
            'total_cases': int(total_cases),
            'infection_rate_pct': round(total_cases / N * 100, 2),
            'mean_beta': round(mean_beta, 4),
            'mean_R0': round(mean_R0_val, 2),
        })

    summary_df = pd.DataFrame(summary_rows)
    summary_df.sort_values('total_cases', ascending=False, inplace=True)
    summary_df.to_csv(os.path.join(OUTPUT_DIR, 'sir_summary.csv'), index=False)
    print("  Saved: sir_summary.csv")

    # Ground truth CSV
    gt_rows = []
    for source, target, strength, lag in GROUND_TRUTH_CONNECTIONS:
        gt_rows.append({
            'source': source.title(),
            'target': target.title(),
            'strength': strength,
            'lag_days': lag,
        })
    gt_df = pd.DataFrame(gt_rows)
    gt_df.to_csv(os.path.join(OUTPUT_DIR, 'ground_truth_connections.csv'),
                  index=False)
    print("  Saved: ground_truth_connections.csv")

    # Print summary
    print("\n" + "=" * 70)
    print("   SIR MODEL SUMMARY")
    print("=" * 70)
    print(f"  {'City':<12} {'Pop.':>12} {'Total Cases':>14} {'Rate(%)':>8} "
          f"{'Mean B':>8} {'R0':>6}")
    print(f"  {'-' * 64}")
    for _, row in summary_df.iterrows():
        print(f"  {row['city']:<12} {row['population']:>12,} "
              f"{row['total_cases']:>14,} {row['infection_rate_pct']:>8.2f} "
              f"{row['mean_beta']:>8.4f} {row['mean_R0']:>6.2f}")

    print(f"\n  Ground Truth Connections: {len(GROUND_TRUTH_CONNECTIONS)}")
    for src, tgt, s, lag in GROUND_TRUTH_CONNECTIONS:
        print(f"    {src.title():12} -> {tgt.title():12} [{s:6}] lag={lag}d")

    print(f"\n  Output directory: {OUTPUT_DIR}")
    print("=" * 70)


if __name__ == '__main__':
    main()
