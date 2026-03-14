"""
SIR Model Implementation for Italian COVID-19 Data (2022)
=========================================================
This script implements a SIR (Susceptible-Infected-Recovered) compartmental
model for 5 Italian cities using daily infection data.

Parameters:
    - gamma (γ) = 0.07 (recovery rate, ~14 days recovery period)
    - beta (β) = time-varying infection rate (fitted per city)
    - N = total population of each province

Cities: Milano, Roma, Brescia, Torino, Napoli
"""

import os
import numpy as np
import pandas as pd
from scipy.integrate import odeint
from scipy.optimize import minimize, differential_evolution
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from itertools import combinations
import warnings
warnings.filterwarnings('ignore')

# ==============================================================================
# 1. CONFIGURATION
# ==============================================================================

GAMMA = 0.07  # Recovery rate (as specified by professor)
RECOVERY_PERIOD = int(1 / GAMMA)  # ~14 days

# Province populations (ISTAT data, approximate)
POPULATIONS = {
    'milano': 3_250_000,
    'roma': 4_350_000,
    'brescia': 1_265_000,
    'torino': 2_250_000,
    'napoli': 3_100_000,
}

CITIES = list(POPULATIONS.keys())

# Data path
DATA_PATH = r"c:\Users\Utente\Documents\GitHub\gcpaida\2022_city_csv\merged_cities_2022.csv"
OUTPUT_DIR = r"c:\Users\Utente\Documents\GitHub\gcpaida\sir_results"

# Create output directory
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==============================================================================
# 2. DATA LOADING & PREPROCESSING
# ==============================================================================

def load_and_preprocess(filepath, cities):
    """Load merged CSV and compute daily new infections for each city."""
    df = pd.read_csv(filepath)
    df['data'] = pd.to_datetime(df['data'])
    df.sort_values('data', inplace=True)
    df.set_index('data', inplace=True)

    city_data = {}
    for city in cities:
        col = f"{city}_infected"
        if col in df.columns:
            cumulative = df[col].values.astype(float)
            # Daily new infections = difference of cumulative
            daily_new = np.diff(cumulative, prepend=cumulative[0])
            daily_new = np.maximum(daily_new, 0)  # Remove negative values

            city_data[city] = {
                'dates': df.index,
                'cumulative': cumulative,
                'daily_new': daily_new,
            }
        else:
            print(f"Warning: Column '{col}' not found in data!")

    return df, city_data


# ==============================================================================
# 3. SIR MODEL
# ==============================================================================

def sir_ode(y, t, beta, gamma, N):
    """SIR differential equations."""
    S, I, R = y
    dSdt = -beta * S * I / N
    dIdt = beta * S * I / N - gamma * I
    dRdt = gamma * I
    return [dSdt, dIdt, dRdt]


def simulate_sir(S0, I0, R0, beta, gamma, N, n_days):
    """Simulate SIR model for n_days with constant beta."""
    y0 = [S0, I0, R0]
    t = np.arange(n_days)
    solution = odeint(sir_ode, y0, t, args=(beta, gamma, N))
    S, I, R = solution[:, 0], solution[:, 1], solution[:, 2]
    return S, I, R


# ==============================================================================
# 4. TIME-VARYING BETA ESTIMATION (Direct from Data)
# ==============================================================================

def estimate_beta_from_data(cumulative, N, gamma, smooth_window=7):
    """
    Estimate time-varying beta directly from observed data.

    Using: beta(t) = (new_cases(t) * N) / (S(t) * I(t))
    where:
        S(t) = N - cumulative(t)
        I(t) = estimated active infections (sum of new cases in last 1/gamma days)
    """
    daily_new = np.diff(cumulative, prepend=cumulative[0])
    daily_new = np.maximum(daily_new, 0)

    # Estimate active infections: sum of new cases in last recovery_period days
    recovery_period = int(1 / gamma)
    active_I = np.zeros(len(daily_new))
    for i in range(len(daily_new)):
        start = max(0, i - recovery_period + 1)
        active_I[i] = np.sum(daily_new[start:i + 1])
    active_I = np.maximum(active_I, 1)  # Avoid division by zero

    # Susceptible population
    S = N - cumulative
    S = np.maximum(S, 1)

    # Estimate beta
    beta_t = (daily_new * N) / (S * active_I)
    beta_t = np.clip(beta_t, 0, 5)  # Clip extreme values

    # Smooth using rolling mean
    beta_smooth = pd.Series(beta_t).rolling(
        window=smooth_window, min_periods=1, center=True
    ).mean().values

    return beta_smooth, daily_new, active_I, S


# ==============================================================================
# 5. PIECEWISE SIR FITTING (Optimize beta per segment)
# ==============================================================================

def fit_sir_piecewise(cumulative, N, gamma, segment_days=30):
    """
    Fit SIR model with piecewise constant beta.
    Beta is constant within each segment but changes between segments.
    """
    n_days = len(cumulative)
    n_segments = max(1, (n_days + segment_days - 1) // segment_days)

    daily_new = np.diff(cumulative, prepend=cumulative[0])
    daily_new = np.maximum(daily_new, 0)

    # Initial conditions
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

        # Normalized MSE
        error = np.mean(((predicted - cumulative) / max(cumulative[-1], 1)) ** 2)
        return error

    # Optimize using differential evolution (more robust)
    bounds = [(0.01, 2.0)] * n_segments
    result = differential_evolution(objective, bounds, seed=42, maxiter=500,
                                     tol=1e-10, polish=True)

    # Run final simulation with best parameters
    betas = result.x
    S, I, R = S0, I0, R0
    S_arr, I_arr, R_arr, C_arr = [], [], [], []

    for day in range(n_days):
        seg_idx = min(day // segment_days, len(betas) - 1)
        beta = abs(betas[seg_idx])

        S_arr.append(S)
        I_arr.append(I)
        R_arr.append(R)
        C_arr.append(N - S)

        dS = -beta * S * I / N
        dI = beta * S * I / N - gamma * I
        dR = gamma * I

        S = max(S + dS, 0)
        I = max(I + dI, 0)
        R = max(R + dR, 0)

    return {
        'betas': betas,
        'S': np.array(S_arr),
        'I': np.array(I_arr),
        'R': np.array(R_arr),
        'predicted_cumulative': np.array(C_arr),
        'success': result.success,
        'error': result.fun,
        'n_segments': n_segments,
        'segment_days': segment_days,
    }


# ==============================================================================
# 6. CROSS-CITY ANALYSIS
# ==============================================================================

def compute_cross_correlation(series1, series2, max_lag=30):
    """Compute cross-correlation between two time series with lags."""
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


def analyze_city_relationships(city_data, cities):
    """Analyze relationships and lags between cities."""
    results = {}
    for city1, city2 in combinations(cities, 2):
        if city1 in city_data and city2 in city_data:
            lags, corr = compute_cross_correlation(
                city_data[city1]['daily_new'],
                city_data[city2]['daily_new'],
                max_lag=30
            )
            best_lag = lags[np.argmax(corr)]
            max_corr = max(corr)
            results[(city1, city2)] = {
                'lags': lags,
                'correlations': corr,
                'best_lag': best_lag,
                'max_correlation': max_corr,
            }
    return results


# ==============================================================================
# 7. VISUALIZATION
# ==============================================================================

def plot_sir_results(city, dates, cumulative, sir_result, beta_t, daily_new, output_dir):
    """Plot SIR model results for a single city."""
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle(f'SIR Model - {city.title()} (2022)', fontsize=16, fontweight='bold')

    # 1. Cumulative: Actual vs Predicted
    ax = axes[0, 0]
    ax.plot(dates, cumulative, 'b-', linewidth=2, label='Actual (totale_casi)')
    ax.plot(dates, sir_result['predicted_cumulative'], 'r--', linewidth=2, label='SIR Predicted')
    ax.set_title('Cumulative Infections: Actual vs SIR Model')
    ax.set_ylabel('Cumulative Cases')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator())

    # 2. SIR Compartments
    ax = axes[0, 1]
    ax.plot(dates, sir_result['S'], 'g-', linewidth=2, label='S (Susceptible)')
    ax.plot(dates, sir_result['I'], 'r-', linewidth=2, label='I (Infected)')
    ax.plot(dates, sir_result['R'], 'b-', linewidth=2, label='R (Recovered)')
    ax.set_title('SIR Compartments')
    ax.set_ylabel('Population')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator())

    # 3. Daily New Cases
    ax = axes[1, 0]
    ax.bar(dates, daily_new, color='salmon', alpha=0.6, label='Daily New Cases')
    ax.plot(dates, pd.Series(daily_new).rolling(7, min_periods=1).mean(),
            'r-', linewidth=2, label='7-day Average')
    ax.set_title('Daily New Infections')
    ax.set_ylabel('New Cases')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator())

    # 4. Time-varying Beta
    ax = axes[1, 1]
    ax.plot(dates, beta_t, 'purple', linewidth=1.5, alpha=0.7, label='Beta(t) estimated')
    # Overlay piecewise beta
    seg_days = sir_result['segment_days']
    for i, b in enumerate(sir_result['betas']):
        start = i * seg_days
        end = min((i + 1) * seg_days, len(dates))
        if start < len(dates):
            ax.hlines(b, dates[start], dates[min(end - 1, len(dates) - 1)],
                      colors='red', linewidths=3, label='Fitted Beta' if i == 0 else '')
    ax.set_title('Infection Rate Beta(t)')
    ax.set_ylabel('Beta')
    ax.axhline(y=GAMMA, color='green', linestyle='--', label=f'Gamma = {GAMMA}')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))
    ax.xaxis.set_major_locator(mdates.MonthLocator())

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, f'sir_{city}.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Plot saved: sir_{city}.png")


def plot_all_cities_comparison(city_data, results, dates, output_dir):
    """Compare all cities in a single plot."""
    fig, axes = plt.subplots(2, 2, figsize=(18, 14))
    fig.suptitle('SIR Model Comparison - All Cities (2022)', fontsize=16, fontweight='bold')

    colors = ['#e74c3c', '#3498db', '#2ecc71', '#f39c12', '#9b59b6']

    # 1. All cumulative curves
    ax = axes[0, 0]
    for i, city in enumerate(results.keys()):
        ax.plot(dates, city_data[city]['cumulative'], color=colors[i],
                linewidth=2, label=city.title())
    ax.set_title('Cumulative Infections - All Cities')
    ax.set_ylabel('Cumulative Cases')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    # 2. All daily new cases (normalized per 100k population)
    ax = axes[0, 1]
    for i, city in enumerate(results.keys()):
        N = POPULATIONS[city]
        daily_norm = city_data[city]['daily_new'] / N * 100000
        daily_smooth = pd.Series(daily_norm).rolling(7, min_periods=1).mean()
        ax.plot(dates, daily_smooth, color=colors[i], linewidth=2, label=city.title())
    ax.set_title('Daily New Cases per 100k (7-day avg)')
    ax.set_ylabel('Cases per 100k')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    # 3. Fitted Beta for all cities
    ax = axes[1, 0]
    for i, city in enumerate(results.keys()):
        beta_t = results[city]['beta_t']
        ax.plot(dates, beta_t, color=colors[i], linewidth=1.5, label=city.title(), alpha=0.8)
    ax.axhline(y=GAMMA, color='black', linestyle='--', linewidth=1, label=f'Gamma={GAMMA}')
    ax.set_title('Time-varying Beta(t) - All Cities')
    ax.set_ylabel('Beta')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    # 4. R0(t) = beta/gamma for all cities
    ax = axes[1, 1]
    for i, city in enumerate(results.keys()):
        R0_t = results[city]['beta_t'] / GAMMA
        R0_smooth = pd.Series(R0_t).rolling(14, min_periods=1).mean()
        ax.plot(dates, R0_smooth, color=colors[i], linewidth=2, label=city.title())
    ax.axhline(y=1.0, color='black', linestyle='--', linewidth=2, label='R0 = 1 (threshold)')
    ax.set_title('Reproduction Number R0(t) = Beta/Gamma')
    ax.set_ylabel('R0')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'sir_all_cities_comparison.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Plot saved: sir_all_cities_comparison.png")


def plot_cross_correlations(cross_results, output_dir):
    """Plot cross-correlation results."""
    n_pairs = len(cross_results)
    cols = 3
    rows = (n_pairs + cols - 1) // cols
    fig, axes = plt.subplots(rows, cols, figsize=(18, 4 * rows))
    fig.suptitle('Cross-Correlation Between Cities (Daily New Cases)', fontsize=14, fontweight='bold')

    if rows == 1:
        axes = [axes] if cols == 1 else axes
    axes_flat = np.array(axes).flatten()

    for idx, ((city1, city2), data) in enumerate(cross_results.items()):
        ax = axes_flat[idx]
        ax.bar(data['lags'], data['correlations'], color='steelblue', alpha=0.7)
        ax.axvline(x=data['best_lag'], color='red', linestyle='--',
                   label=f'Best lag = {data["best_lag"]} days')
        ax.set_title(f'{city1.title()} vs {city2.title()}\n(max corr = {data["max_correlation"]:.3f})')
        ax.set_xlabel('Lag (days)')
        ax.set_ylabel('Correlation')
        ax.legend(fontsize=8)
        ax.grid(True, alpha=0.3)

    # Hide unused axes
    for idx in range(n_pairs, len(axes_flat)):
        axes_flat[idx].set_visible(False)

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'cross_correlations.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Plot saved: cross_correlations.png")


def plot_correlation_matrix(city_data, cities, output_dir):
    """Plot correlation matrix between cities."""
    # Build matrix of daily new cases
    daily_matrix = pd.DataFrame()
    for city in cities:
        if city in city_data:
            daily_matrix[city.title()] = city_data[city]['daily_new']

    corr_matrix = daily_matrix.corr()

    fig, ax = plt.subplots(figsize=(8, 7))
    im = ax.imshow(corr_matrix.values, cmap='RdYlBu_r', vmin=0, vmax=1)

    # Labels
    ax.set_xticks(range(len(corr_matrix.columns)))
    ax.set_yticks(range(len(corr_matrix.columns)))
    ax.set_xticklabels(corr_matrix.columns, rotation=45, ha='right')
    ax.set_yticklabels(corr_matrix.columns)

    # Annotate
    for i in range(len(corr_matrix)):
        for j in range(len(corr_matrix)):
            ax.text(j, i, f'{corr_matrix.values[i, j]:.3f}',
                    ha='center', va='center', fontsize=11, fontweight='bold')

    plt.colorbar(im, ax=ax, label='Correlation')
    ax.set_title('Correlation Matrix - Daily New Infections', fontsize=14, fontweight='bold')
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'correlation_matrix.png'), dpi=150, bbox_inches='tight')
    plt.close()
    print("  Plot saved: correlation_matrix.png")


# ==============================================================================
# 8. MAIN EXECUTION
# ==============================================================================

def main():
    print("=" * 70)
    print("   SIR Model for Italian COVID-19 Data (2022)")
    print(f"   Gamma (recovery rate) = {GAMMA}")
    print(f"   Recovery period = {RECOVERY_PERIOD} days")
    print(f"   Cities: {', '.join([c.title() for c in CITIES])}")
    print("=" * 70)

    # Load data
    print("\n[1/5] Loading and preprocessing data...")
    df, city_data = load_and_preprocess(DATA_PATH, CITIES)
    dates = city_data[CITIES[0]]['dates']
    print(f"  Data loaded: {len(dates)} days, {len(city_data)} cities")

    # Estimate beta and fit SIR for each city
    print("\n[2/5] Estimating beta(t) and fitting SIR model for each city...")
    results = {}
    for city in CITIES:
        if city not in city_data:
            continue
        print(f"\n  >> {city.title()} (N = {POPULATIONS[city]:,})")

        cumulative = city_data[city]['cumulative']

        # Direct beta estimation
        beta_t, daily_new, active_I, S = estimate_beta_from_data(
            cumulative, POPULATIONS[city], GAMMA, smooth_window=14
        )

        # Piecewise SIR fitting (monthly segments)
        print(f"     Fitting piecewise SIR model (30-day segments)...")
        sir_fit = fit_sir_piecewise(cumulative, POPULATIONS[city], GAMMA, segment_days=30)

        results[city] = {
            'beta_t': beta_t,
            'daily_new': daily_new,
            'active_I': active_I,
            'S': S,
            'sir_fit': sir_fit,
        }

        # Print results
        print(f"     Fitted betas per segment: {np.round(sir_fit['betas'], 4)}")
        print(f"     Mean beta: {np.mean(sir_fit['betas']):.4f}")
        print(f"     Mean R0 = beta/gamma: {np.mean(sir_fit['betas']) / GAMMA:.2f}")
        print(f"     Fit error (norm. MSE): {sir_fit['error']:.6f}")
        print(f"     Total infections: {cumulative[-1]:,.0f}")

    # Generate plots for each city
    print("\n[3/5] Generating individual city plots...")
    for city in results:
        plot_sir_results(
            city, dates, city_data[city]['cumulative'],
            results[city]['sir_fit'], results[city]['beta_t'],
            results[city]['daily_new'], OUTPUT_DIR
        )

    # Comparison plot
    print("\n[4/5] Generating city comparison plots...")
    plot_all_cities_comparison(city_data, results, dates, OUTPUT_DIR)
    plot_correlation_matrix(city_data, list(results.keys()), OUTPUT_DIR)

    # Cross-city analysis
    print("\n[5/5] Analyzing cross-city relationships (lags)...")
    cross_results = analyze_city_relationships(city_data, list(results.keys()))
    plot_cross_correlations(cross_results, OUTPUT_DIR)

    # Print cross-correlation summary
    print("\n" + "=" * 70)
    print("   CROSS-CITY RELATIONSHIP SUMMARY")
    print("=" * 70)
    for (c1, c2), data in cross_results.items():
        independent = "INDEPENDENT" if abs(data['best_lag']) > 7 else "CORRELATED"
        print(f"  {c1.title():10s} <-> {c2.title():10s}: "
              f"lag = {data['best_lag']:+3d} days, "
              f"corr = {data['max_correlation']:.3f}  [{independent}]")

    # Summary table
    print("\n" + "=" * 70)
    print("   SIR MODEL SUMMARY TABLE")
    print("=" * 70)
    print(f"  {'City':<12} {'Population':>12} {'Total Cases':>14} {'Mean Beta':>10} {'Mean R0':>8}")
    print(f"  {'-' * 60}")
    for city in results:
        N = POPULATIONS[city]
        total = city_data[city]['cumulative'][-1]
        mean_beta = np.mean(results[city]['sir_fit']['betas'])
        R0 = mean_beta / GAMMA
        print(f"  {city.title():<12} {N:>12,} {total:>14,.0f} {mean_beta:>10.4f} {R0:>8.2f}")

    print(f"\n  All results saved to: {OUTPUT_DIR}")
    print("=" * 70)

    # Save summary to CSV
    summary_rows = []
    for city in results:
        summary_rows.append({
            'city': city.title(),
            'population': POPULATIONS[city],
            'total_cases_2022': city_data[city]['cumulative'][-1],
            'mean_beta': np.mean(results[city]['sir_fit']['betas']),
            'mean_R0': np.mean(results[city]['sir_fit']['betas']) / GAMMA,
            'fit_error': results[city]['sir_fit']['error'],
        })
    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_csv(os.path.join(OUTPUT_DIR, 'sir_summary.csv'), index=False)
    print("  Summary saved: sir_summary.csv")


if __name__ == '__main__':
    main()
