"""
SIR Model for 15 Selected Italian Cities - COVID-19 Data (2020)
=================================================================
Applies SIR compartmental model to 15 selected cities from three
geographic clusters (North, Center, South Italy).

Parameters:
    - gamma (γ) = 0.07 (recovery rate, ~14 days recovery period)
    - beta (β) = time-varying infection rate (30-day piecewise segments)

Cities:
    North (Lombardy):  Milano, Bergamo, Brescia, Monza e della Brianza, Como
    Center (Lazio/Toscana): Roma, Firenze, Perugia, Latina, Frosinone
    South (Campania/Puglia): Napoli, Caserta, Salerno, Bari, Taranto

Output: 2020v2/sir_results/
"""

import os
import numpy as np
import pandas as pd
from scipy.optimize import differential_evolution
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from itertools import combinations
import warnings
warnings.filterwarnings('ignore')

# ==============================================================================
# CONFIGURATION
# ==============================================================================

GAMMA = 0.07
RECOVERY_PERIOD = int(1 / GAMMA)
SEGMENT_DAYS = 30

DATA_PATH = r"c:\Users\Utente\Documents\GitHub\gcpaida\2020\merged_cities_2020.csv"
OUTPUT_DIR = r"c:\Users\Utente\Documents\GitHub\gcpaida\2020v2\sir_results"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 15 Selected cities grouped by region
CITY_GROUPS = {
    'North': ['milano', 'bergamo', 'brescia', 'monza_e_della_brianza', 'como'],
    'Center': ['roma', 'firenze', 'perugia', 'latina', 'frosinone'],
    'South': ['napoli', 'caserta', 'salerno', 'bari', 'taranto'],
}

ALL_CITIES = []
for group_cities in CITY_GROUPS.values():
    ALL_CITIES.extend(group_cities)

# Province populations (ISTAT 2020 approximate)
PROVINCE_POPULATIONS = {
    'milano': 3_250_000,
    'bergamo': 1_115_000,
    'brescia': 1_265_000,
    'monza_e_della_brianza': 875_000,
    'como': 600_000,
    'roma': 4_350_000,
    'firenze': 1_010_000,
    'perugia': 655_000,
    'latina': 575_000,
    'frosinone': 475_000,
    'napoli': 3_100_000,
    'caserta': 925_000,
    'salerno': 1_080_000,
    'bari': 1_250_000,
    'taranto': 560_000,
}

# Colors per region
REGION_COLORS = {
    'North': '#2196F3',    # Blue
    'Center': '#4CAF50',   # Green
    'South': '#FF9800',    # Orange
}

CITY_REGION_MAP = {}
for region, cities in CITY_GROUPS.items():
    for city in cities:
        CITY_REGION_MAP[city] = region


# ==============================================================================
# SIR MODEL FUNCTIONS
# ==============================================================================

def estimate_beta_from_data(cumulative, N, gamma, smooth_window=14):
    daily_new = np.diff(cumulative, prepend=cumulative[0])
    daily_new = np.maximum(daily_new, 0)
    recovery_period = int(1 / gamma)
    active_I = np.zeros(len(daily_new))
    for i in range(len(daily_new)):
        start = max(0, i - recovery_period + 1)
        active_I[i] = np.sum(daily_new[start:i + 1])
    active_I = np.maximum(active_I, 1)
    S = np.maximum(N - cumulative, 1)
    beta_t = (daily_new * N) / (S * active_I)
    beta_t = np.clip(beta_t, 0, 5)
    beta_smooth = pd.Series(beta_t).rolling(
        window=smooth_window, min_periods=1, center=True
    ).mean().values
    return beta_smooth, daily_new, active_I, S


def fit_sir_piecewise(cumulative, N, gamma, segment_days=30):
    n_days = len(cumulative)
    n_segments = max(1, (n_days + segment_days - 1) // segment_days)
    daily_new = np.diff(cumulative, prepend=cumulative[0])
    daily_new = np.maximum(daily_new, 0)
    I0 = max(np.mean(daily_new[:7]) * RECOVERY_PERIOD, 1)
    R0 = max(cumulative[0] - I0, 0)
    S0 = N - I0 - R0

    def objective(betas):
        S, I, R = S0, I0, R0
        predicted = np.zeros(n_days)
        for day in range(n_days):
            seg_idx = min(day // segment_days, len(betas) - 1)
            beta = abs(betas[seg_idx])
            dS = -beta * S * I / N
            dI = beta * S * I / N - gamma * I
            dR = gamma * I
            S = max(S + dS, 0)
            I = max(I + dI, 0)
            R = max(R + dR, 0)
            predicted[day] = N - S
        return np.mean(((predicted - cumulative) / max(cumulative[-1], 1)) ** 2)

    bounds = [(0.01, 2.0)] * n_segments
    result = differential_evolution(objective, bounds, seed=42, maxiter=300, tol=1e-8, polish=True)

    betas = result.x
    S, I, R = S0, I0, R0
    S_arr, I_arr, R_arr = [], [], []
    for day in range(n_days):
        S_arr.append(S)
        I_arr.append(I)
        R_arr.append(R)
        seg_idx = min(day // segment_days, len(betas) - 1)
        beta = abs(betas[seg_idx])
        dS = -beta * S * I / N
        dI = beta * S * I / N - gamma * I
        dR = gamma * I
        S = max(S + dS, 0)
        I = max(I + dI, 0)
        R = max(R + dR, 0)

    return {
        'betas': betas, 'S': np.array(S_arr), 'I': np.array(I_arr),
        'R': np.array(R_arr), 'predicted_cumulative': N - np.array(S_arr),
        'error': result.fun, 'n_segments': n_segments, 'segment_days': segment_days,
    }


# ==============================================================================
# CROSS-CITY ANALYSIS
# ==============================================================================

def compute_cross_correlation(series1, series2, max_lag=30):
    s1 = (series1 - np.mean(series1)) / (np.std(series1) + 1e-10)
    s2 = (series2 - np.mean(series2)) / (np.std(series2) + 1e-10)
    n = len(s1)
    correlations = []
    lags = range(-max_lag, max_lag + 1)
    for lag in lags:
        if lag >= 0:
            c = np.mean(s1[lag:] * s2[:n - lag]) if n - lag > 0 else 0
        else:
            c = np.mean(s1[:n + lag] * s2[-lag:]) if n + lag > 0 else 0
        correlations.append(c)
    return list(lags), correlations


# ==============================================================================
# PLOTTING
# ==============================================================================

def plot_city(city, dates, cumulative, sir_result, beta_t, daily_new, output_dir):
    region = CITY_REGION_MAP.get(city, 'Unknown')
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle(f'SIR Model - {city.title()} ({region}) (2020)', fontsize=16, fontweight='bold')

    ax = axes[0, 0]
    ax.plot(dates, cumulative, 'b-', lw=2, label='Actual')
    ax.plot(dates, sir_result['predicted_cumulative'], 'r--', lw=2, label='SIR Predicted')
    ax.set_title('Cumulative Infections')
    ax.set_ylabel('Cases')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))

    ax = axes[0, 1]
    ax.plot(dates, sir_result['S'], 'g-', lw=2, label='S')
    ax.plot(dates, sir_result['I'], 'r-', lw=2, label='I')
    ax.plot(dates, sir_result['R'], 'b-', lw=2, label='R')
    ax.set_title('SIR Compartments')
    ax.set_ylabel('Population')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))

    ax = axes[1, 0]
    ax.bar(dates, daily_new, color='salmon', alpha=0.6, width=1)
    ax.plot(dates, pd.Series(daily_new).rolling(7, min_periods=1).mean(), 'r-', lw=2, label='7-day avg')
    ax.set_title('Daily New Infections')
    ax.set_ylabel('New Cases')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))

    ax = axes[1, 1]
    ax.plot(dates, beta_t, 'purple', lw=1.5, alpha=0.7, label='Beta(t)')
    seg = sir_result['segment_days']
    for i, b in enumerate(sir_result['betas']):
        s = i * seg
        e = min((i + 1) * seg, len(dates))
        if s < len(dates):
            ax.hlines(b, dates[s], dates[min(e - 1, len(dates) - 1)],
                      colors='red', linewidths=3, label='Fitted' if i == 0 else '')
    ax.axhline(y=GAMMA, color='green', ls='--', label=f'Gamma={GAMMA}')
    ax.set_title('Beta(t) - Infection Rate')
    ax.set_ylabel('Beta')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))

    plt.tight_layout()
    safe_name = city.replace("'", "").replace(" ", "_")
    plt.savefig(os.path.join(output_dir, f'sir_{safe_name}.png'), dpi=120, bbox_inches='tight')
    plt.close()


def plot_regional_comparison(all_results, dates, output_dir):
    """Compare cities grouped by region."""
    fig, axes = plt.subplots(3, 2, figsize=(20, 22))
    fig.suptitle('SIR Model - Regional Comparison (2020)\n15 Selected Cities', fontsize=18, fontweight='bold')

    region_names = list(CITY_GROUPS.keys())
    for row_idx, region in enumerate(region_names):
        cities = CITY_GROUPS[region]
        color = REGION_COLORS[region]
        city_colors = plt.cm.Set2(np.linspace(0, 1, len(cities)))

        # Cumulative
        ax = axes[row_idx, 0]
        for i, city in enumerate(cities):
            if city in all_results:
                ax.plot(dates, all_results[city]['cumulative'], color=city_colors[i],
                        lw=2, label=city.title())
        ax.set_title(f'{region} Italy - Cumulative Infections', fontsize=13, fontweight='bold')
        ax.legend(fontsize=8)
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

        # Daily new per 100k
        ax = axes[row_idx, 1]
        for i, city in enumerate(cities):
            if city in all_results:
                N = PROVINCE_POPULATIONS[city]
                daily_norm = all_results[city]['daily_new'] / N * 100000
                ax.plot(dates, pd.Series(daily_norm).rolling(7, min_periods=1).mean(),
                        color=city_colors[i], lw=2, label=city.title())
        ax.set_title(f'{region} Italy - Daily per 100k (7-day avg)', fontsize=13, fontweight='bold')
        ax.legend(fontsize=8)
        ax.grid(True, alpha=0.3)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'regional_comparison.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: regional_comparison.png")


def plot_all_cities_comparison(all_results, dates, output_dir):
    """Compare all 15 cities in a single figure."""
    fig, axes = plt.subplots(2, 2, figsize=(20, 14))
    fig.suptitle('SIR Model - All 15 Selected Cities (2020)', fontsize=16, fontweight='bold')

    for city in all_results:
        region = CITY_REGION_MAP.get(city, 'Unknown')
        color = REGION_COLORS.get(region, 'gray')
        alpha = 0.8

        # Cumulative
        axes[0, 0].plot(dates, all_results[city]['cumulative'],
                        color=color, lw=1.5, alpha=alpha, label=city.title())

        # Daily per 100k
        N = PROVINCE_POPULATIONS.get(city, 300_000)
        daily_norm = all_results[city]['daily_new'] / N * 100000
        axes[0, 1].plot(dates, pd.Series(daily_norm).rolling(7, min_periods=1).mean(),
                        color=color, lw=1.5, alpha=alpha, label=city.title())

        # Beta
        axes[1, 0].plot(dates, all_results[city]['beta_t'],
                        color=color, lw=1, alpha=0.6, label=city.title())

        # R0
        R0 = pd.Series(all_results[city]['beta_t'] / GAMMA).rolling(14, min_periods=1).mean()
        axes[1, 1].plot(dates, R0, color=color, lw=1.5, alpha=alpha, label=city.title())

    axes[0, 0].set_title('Cumulative Infections')
    axes[0, 0].legend(fontsize=6, ncol=3)
    axes[0, 0].grid(True, alpha=0.3)

    axes[0, 1].set_title('Daily New Cases per 100k (7-day avg)')
    axes[0, 1].legend(fontsize=6, ncol=3)
    axes[0, 1].grid(True, alpha=0.3)

    axes[1, 0].set_title('Beta(t)')
    axes[1, 0].axhline(y=GAMMA, color='black', ls='--', lw=1, label=f'Gamma={GAMMA}')
    axes[1, 0].legend(fontsize=6, ncol=3)
    axes[1, 0].grid(True, alpha=0.3)

    axes[1, 1].set_title('R0(t) = Beta/Gamma')
    axes[1, 1].axhline(y=1, color='black', ls='--', lw=2, label='R0=1')
    axes[1, 1].legend(fontsize=6, ncol=3)
    axes[1, 1].grid(True, alpha=0.3)

    for ax in axes.flatten():
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    # Add legend patches for regions
    import matplotlib.patches as mpatches
    patches = [mpatches.Patch(color=c, label=r) for r, c in REGION_COLORS.items()]
    fig.legend(handles=patches, loc='upper right', fontsize=10, title='Region')

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'all_cities_comparison.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: all_cities_comparison.png")


def plot_cross_correlations(cross_results, output_dir):
    """Plot cross-correlation between city pairs within each region."""
    n_pairs = len(cross_results)
    if n_pairs == 0:
        return
    cols = 3
    rows = (n_pairs + cols - 1) // cols
    fig, axes = plt.subplots(rows, cols, figsize=(18, 4 * rows))
    fig.suptitle('Cross-Correlation Between Cities (Daily New Cases) - 2020', fontsize=14, fontweight='bold')

    if rows == 1:
        axes = [axes] if cols == 1 else axes
    axes_flat = np.array(axes).flatten()

    for idx, ((city1, city2), data) in enumerate(cross_results.items()):
        if idx >= len(axes_flat):
            break
        ax = axes_flat[idx]
        ax.bar(data['lags'], data['correlations'], color='steelblue', alpha=0.7)
        ax.axvline(x=data['best_lag'], color='red', linestyle='--',
                   label=f'Best lag = {data["best_lag"]} days')
        ax.set_title(f'{city1.title()} vs {city2.title()}\n(max corr = {data["max_correlation"]:.3f})')
        ax.set_xlabel('Lag (days)')
        ax.set_ylabel('Correlation')
        ax.legend(fontsize=8)
        ax.grid(True, alpha=0.3)

    for idx in range(n_pairs, len(axes_flat)):
        axes_flat[idx].set_visible(False)

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'cross_correlations.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: cross_correlations.png")


def plot_correlation_matrix(all_results, cities, dates, output_dir):
    """Correlation matrix between all 15 cities."""
    daily_matrix = pd.DataFrame()
    for city in cities:
        if city in all_results:
            daily_matrix[city.title()] = all_results[city]['daily_new']

    corr_matrix = daily_matrix.corr()

    fig, ax = plt.subplots(figsize=(12, 10))
    im = ax.imshow(corr_matrix.values, cmap='RdYlBu_r', vmin=0, vmax=1)

    ax.set_xticks(range(len(corr_matrix.columns)))
    ax.set_yticks(range(len(corr_matrix.columns)))
    ax.set_xticklabels(corr_matrix.columns, rotation=45, ha='right', fontsize=9)
    ax.set_yticklabels(corr_matrix.columns, fontsize=9)

    for i in range(len(corr_matrix)):
        for j in range(len(corr_matrix)):
            ax.text(j, i, f'{corr_matrix.values[i, j]:.3f}',
                    ha='center', va='center', fontsize=7, fontweight='bold')

    plt.colorbar(im, ax=ax, label='Correlation')
    ax.set_title('Correlation Matrix - 15 Selected Cities (2020)', fontsize=14, fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'correlation_matrix.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Saved: correlation_matrix.png")


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    print("=" * 70)
    print("   SIR Model - 15 Selected Italian Cities (2020v2)")
    print(f"   Gamma = {GAMMA}, Recovery = {RECOVERY_PERIOD} days")
    print("   North: " + ", ".join([c.title() for c in CITY_GROUPS['North']]))
    print("   Center: " + ", ".join([c.title() for c in CITY_GROUPS['Center']]))
    print("   South: " + ", ".join([c.title() for c in CITY_GROUPS['South']]))
    print(f"   Output: {OUTPUT_DIR}")
    print("=" * 70)

    # Load data
    print("\n[1/5] Loading merged data...")
    df = pd.read_csv(DATA_PATH)
    df['data'] = pd.to_datetime(df['data'])
    df.sort_values('data', inplace=True)
    df.set_index('data', inplace=True)
    dates = df.index

    # Identify columns for our 15 cities
    city_cols = [c for c in df.columns if c.startswith('infected_')]
    col_name_map = {}
    for col in city_cols:
        city_key = col.replace('infected_', '').lower().replace(' ', '_')
        col_name_map[city_key] = col

    print(f"  Found {len(city_cols)} total city columns in data")

    # Process each of the 15 selected cities
    print(f"\n[2/5] Fitting SIR model for 15 selected cities...")
    all_results = {}
    summary_rows = []

    for city in ALL_CITIES:
        col = col_name_map.get(city, None)
        if col is None:
            # Try alternative name matching
            for key, val in col_name_map.items():
                if city.replace('_', '') in key.replace('_', ''):
                    col = val
                    break
        if col is None:
            print(f"  WARNING: Column for '{city}' not found, skipping")
            continue

        N = PROVINCE_POPULATIONS.get(city, 300_000)
        cumulative = df[col].values.astype(float)
        cumulative = np.nan_to_num(cumulative, nan=0)

        if cumulative[-1] == 0:
            print(f"  {city.title():30s} - SKIPPED (no data)")
            continue

        region = CITY_REGION_MAP[city]

        # Estimate beta
        beta_t, daily_new, active_I, S = estimate_beta_from_data(cumulative, N, GAMMA)

        # Fit SIR
        sir_fit = fit_sir_piecewise(cumulative, N, GAMMA, SEGMENT_DAYS)

        all_results[city] = {
            'cumulative': cumulative,
            'daily_new': daily_new,
            'beta_t': beta_t,
            'sir_fit': sir_fit,
            'N': N,
        }

        mean_beta = np.mean(sir_fit['betas'])
        R0 = mean_beta / GAMMA

        summary_rows.append({
            'city': city.title(),
            'region': region,
            'population': N,
            'total_cases': cumulative[-1],
            'infection_rate_pct': cumulative[-1] / N * 100,
            'mean_beta': round(mean_beta, 4),
            'mean_R0': round(R0, 2),
            'fit_error': round(sir_fit['error'], 8),
        })

        print(f"  [{region:6s}] {city.title():30s}  N={N:>10,}  "
              f"Cases={cumulative[-1]:>10,.0f}  Beta={mean_beta:.4f}  R0={R0:.2f}")

        # Individual city plot
        plot_city(city, dates, cumulative, sir_fit, beta_t, daily_new, OUTPUT_DIR)

    # Save summary CSV
    print(f"\n[3/5] Saving summary...")
    summary_df = pd.DataFrame(summary_rows)
    summary_df.sort_values('total_cases', ascending=False, inplace=True)
    summary_df.to_csv(os.path.join(OUTPUT_DIR, 'sir_summary_15cities.csv'), index=False)
    print(f"  Summary CSV saved ({len(summary_df)} cities)")

    # Generate comparison plots
    print(f"\n[4/5] Generating comparison plots...")
    plot_all_cities_comparison(all_results, dates, OUTPUT_DIR)
    plot_regional_comparison(all_results, dates, OUTPUT_DIR)
    plot_correlation_matrix(all_results, ALL_CITIES, dates, OUTPUT_DIR)

    # Cross-city analysis (within regions)
    print(f"\n[5/5] Analyzing cross-city relationships...")
    cross_results = {}
    for region, cities in CITY_GROUPS.items():
        for c1, c2 in combinations(cities, 2):
            if c1 in all_results and c2 in all_results:
                lags, corr = compute_cross_correlation(
                    all_results[c1]['daily_new'], all_results[c2]['daily_new'], max_lag=30
                )
                best_lag = lags[np.argmax(corr)]
                max_corr = max(corr)
                cross_results[(c1, c2)] = {
                    'lags': lags, 'correlations': corr,
                    'best_lag': best_lag, 'max_correlation': max_corr,
                    'region': region,
                }
    plot_cross_correlations(cross_results, OUTPUT_DIR)

    # Print summary
    print("\n" + "=" * 70)
    print("   SIR MODEL SUMMARY - 15 SELECTED CITIES (2020v2)")
    print("=" * 70)
    print(f"  {'City':<25} {'Region':<8} {'Population':>12} {'Cases':>12} {'R0':>6}")
    print(f"  {'-' * 65}")
    for _, row in summary_df.iterrows():
        print(f"  {row['city']:<25} {row['region']:<8} {row['population']:>12,} "
              f"{row['total_cases']:>12,.0f} {row['mean_R0']:>6.2f}")

    print(f"\n  Cross-correlations (within regions):")
    for (c1, c2), data in cross_results.items():
        print(f"    {c1.title():20s} <-> {c2.title():20s}: "
              f"lag = {data['best_lag']:+3d} days, corr = {data['max_correlation']:.3f}")

    print(f"\n  Output: {OUTPUT_DIR}")
    print("=" * 70)


if __name__ == '__main__':
    main()
