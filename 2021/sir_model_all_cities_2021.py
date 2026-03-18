"""
SIR Model for ALL 107 Italian Provinces - COVID-19 Data (2021)
===============================================================
Reads merged_cities_2021.csv and applies SIR compartmental model
to all 107 provinces. Results saved to sir-result-2021/ folder.

Parameters:
    - gamma = 0.07 (recovery rate)
    - beta = time-varying, fitted per city (30-day piecewise segments)
    - N = province population (ISTAT approximate data)
"""

import os
import numpy as np
import pandas as pd
from scipy.optimize import differential_evolution
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend for speed
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import warnings
warnings.filterwarnings('ignore')

# ==============================================================================
# CONFIGURATION
# ==============================================================================

GAMMA = 0.07
RECOVERY_PERIOD = int(1 / GAMMA)
SEGMENT_DAYS = 30

DATA_PATH = r"c:\Users\Utente\Documents\GitHub\gcpaida\2021\merged_cities_2021.csv"
OUTPUT_DIR = r"c:\Users\Utente\Documents\GitHub\gcpaida\2021\sir-result-2021"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Italian province populations (ISTAT 2021 approximate)
PROVINCE_POPULATIONS = {
    'agrigento': 415_000, 'alessandria': 410_000, 'ancona': 470_000,
    'aosta': 124_000, 'arezzo': 340_000, 'ascoli_piceno': 205_000,
    'asti': 210_000, 'avellino': 400_000, 'bari': 1_250_000,
    'barletta-andria-trani': 390_000, 'belluno': 200_000, 'benevento': 270_000,
    'bergamo': 1_115_000, 'biella': 175_000, 'bologna': 1_020_000,
    'bolzano': 535_000, 'brescia': 1_265_000, 'brindisi': 385_000,
    'cagliari': 430_000, 'caltanissetta': 255_000, 'campobasso': 215_000,
    'caserta': 925_000, 'catania': 1_080_000, 'catanzaro': 350_000,
    'chieti': 380_000, 'como': 600_000, 'cosenza': 700_000,
    'cremona': 355_000, 'crotone': 170_000, 'cuneo': 585_000,
    'enna': 160_000, 'fermo': 170_000, 'ferrara': 345_000,
    'firenze': 1_010_000, 'foggia': 610_000, 'forli-cesena': 395_000,
    "forl\u00ec-cesena": 395_000, 'frosinone': 475_000, 'genova': 830_000,
    'gorizia': 140_000, 'grosseto': 220_000, 'imperia': 210_000,
    'isernia': 83_000, "l'aquila": 295_000, 'la_spezia': 218_000,
    'latina': 575_000, 'lecce': 790_000, 'lecco': 340_000,
    'livorno': 335_000, 'lodi': 230_000, 'lucca': 390_000,
    'macerata': 310_000, 'mantova': 410_000, 'massa_carrara': 195_000,
    'matera': 195_000, 'messina': 620_000, 'milano': 3_250_000,
    'modena': 710_000, 'monza_e_della_brianza': 875_000, 'napoli': 3_100_000,
    'novara': 365_000, 'nuoro': 205_000, 'oristano': 155_000,
    'padova': 940_000, 'palermo': 1_250_000, 'parma': 455_000,
    'pavia': 545_000, 'perugia': 655_000, 'pesaro_e_urbino': 360_000,
    'pescara': 320_000, 'piacenza': 285_000, 'pisa': 420_000,
    'pistoia': 295_000, 'pordenone': 310_000, 'potenza': 360_000,
    'prato': 260_000, 'ragusa': 320_000, 'ravenna': 390_000,
    'reggio_di_calabria': 530_000, "reggio_nell'emilia": 530_000,
    'rieti': 155_000, 'rimini': 340_000, 'roma': 4_350_000,
    'rovigo': 235_000, 'salerno': 1_080_000, 'sassari': 490_000,
    'savona': 275_000, 'siena': 265_000, 'siracusa': 390_000,
    'sondrio': 180_000, 'sud_sardegna': 340_000, 'taranto': 560_000,
    'teramo': 305_000, 'terni': 225_000, 'torino': 2_250_000,
    'trapani': 425_000, 'trento': 545_000, 'treviso': 890_000,
    'trieste': 235_000, 'udine': 530_000, 'varese': 890_000,
    'venezia': 855_000, 'verbano-cusio-ossola': 155_000, 'vercelli': 170_000,
    'verona': 930_000, 'vibo_valentia': 155_000, 'vicenza': 860_000,
    'viterbo': 310_000,
}

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
    beta_smooth = pd.Series(beta_t).rolling(window=smooth_window, min_periods=1, center=True).mean().values
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
# PLOTTING
# ==============================================================================

def plot_city(city, dates, cumulative, sir_result, beta_t, daily_new, output_dir):
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    fig.suptitle(f'SIR Model - {city.title()} (2021)', fontsize=16, fontweight='bold')

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


def plot_top_cities_comparison(all_results, dates, output_dir, top_n=10):
    """Compare top N cities by total cases."""
    sorted_cities = sorted(all_results.keys(),
                           key=lambda c: all_results[c]['cumulative'][-1], reverse=True)
    top = sorted_cities[:top_n]
    colors = plt.cm.tab10(np.linspace(0, 1, top_n))

    fig, axes = plt.subplots(2, 2, figsize=(20, 14))
    fig.suptitle(f'SIR Model - Top {top_n} Cities by Total Cases (2021)', fontsize=16, fontweight='bold')

    ax = axes[0, 0]
    for i, city in enumerate(top):
        ax.plot(dates, all_results[city]['cumulative'], color=colors[i], lw=2, label=city.title())
    ax.set_title('Cumulative Infections')
    ax.legend(fontsize=8, ncol=2)
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    ax = axes[0, 1]
    for i, city in enumerate(top):
        N = all_results[city]['N']
        daily_norm = all_results[city]['daily_new'] / N * 100000
        ax.plot(dates, pd.Series(daily_norm).rolling(7, min_periods=1).mean(),
                color=colors[i], lw=2, label=city.title())
    ax.set_title('Daily New Cases per 100k (7-day avg)')
    ax.legend(fontsize=8, ncol=2)
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    ax = axes[1, 0]
    for i, city in enumerate(top):
        ax.plot(dates, all_results[city]['beta_t'], color=colors[i], lw=1.5, alpha=0.8, label=city.title())
    ax.axhline(y=GAMMA, color='black', ls='--', lw=1, label=f'Gamma={GAMMA}')
    ax.set_title('Beta(t) - All Cities')
    ax.legend(fontsize=8, ncol=2)
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    ax = axes[1, 1]
    for i, city in enumerate(top):
        R0 = pd.Series(all_results[city]['beta_t'] / GAMMA).rolling(14, min_periods=1).mean()
        ax.plot(dates, R0, color=colors[i], lw=2, label=city.title())
    ax.axhline(y=1, color='black', ls='--', lw=2, label='R0=1')
    ax.set_title('R0(t) = Beta/Gamma')
    ax.legend(fontsize=8, ncol=2)
    ax.grid(True, alpha=0.3)
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b'))

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'comparison_top10.png'), dpi=150, bbox_inches='tight')
    plt.close()


def plot_all_beta_heatmap(all_results, dates, output_dir):
    """Heatmap of beta(t) for all cities."""
    cities = sorted(all_results.keys())
    beta_matrix = np.array([all_results[c]['beta_t'] for c in cities])

    fig, ax = plt.subplots(figsize=(20, 30))
    im = ax.imshow(beta_matrix, aspect='auto', cmap='YlOrRd', vmin=0, vmax=0.3)
    ax.set_yticks(range(len(cities)))
    ax.set_yticklabels([c.title() for c in cities], fontsize=6)
    # x-axis: months
    month_ticks = [0]
    month_labels = ['Jan']
    months = ['Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    for i, m in enumerate(months):
        day = (i + 1) * 30
        if day < len(dates):
            month_ticks.append(day)
            month_labels.append(m)
    ax.set_xticks(month_ticks)
    ax.set_xticklabels(month_labels)
    ax.set_title('Beta(t) Heatmap - All 107 Cities (2021)', fontsize=14, fontweight='bold')
    ax.set_xlabel('Month')
    plt.colorbar(im, ax=ax, label='Beta', shrink=0.5)
    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'beta_heatmap_all_cities.png'), dpi=150, bbox_inches='tight')
    plt.close()


def plot_R0_distribution(all_results, output_dir):
    """Distribution of mean R0 across all cities."""
    cities = sorted(all_results.keys())
    mean_R0 = [np.mean(all_results[c]['sir_fit']['betas']) / GAMMA for c in cities]

    fig, axes = plt.subplots(1, 2, figsize=(18, 8))

    ax = axes[0]
    ax.barh(range(len(cities)), mean_R0, color='steelblue', alpha=0.7)
    ax.set_yticks(range(len(cities)))
    ax.set_yticklabels([c.title() for c in cities], fontsize=5)
    ax.axvline(x=1, color='red', ls='--', lw=2, label='R0=1')
    ax.set_xlabel('Mean R0')
    ax.set_title('Mean R0 per City', fontsize=14, fontweight='bold')
    ax.legend()
    ax.grid(True, alpha=0.3, axis='x')

    ax = axes[1]
    ax.hist(mean_R0, bins=20, color='steelblue', alpha=0.7, edgecolor='black')
    ax.axvline(x=np.mean(mean_R0), color='red', ls='--', lw=2,
               label=f'Mean = {np.mean(mean_R0):.2f}')
    ax.set_xlabel('Mean R0')
    ax.set_ylabel('Number of Cities')
    ax.set_title('Distribution of Mean R0', fontsize=14, fontweight='bold')
    ax.legend()
    ax.grid(True, alpha=0.3)

    plt.tight_layout()
    plt.savefig(os.path.join(output_dir, 'R0_distribution.png'), dpi=150, bbox_inches='tight')
    plt.close()


# ==============================================================================
# MAIN
# ==============================================================================

def main():
    print("=" * 70)
    print("   SIR Model - ALL 107 Italian Provinces (2021)")
    print(f"   Gamma = {GAMMA}, Recovery = {RECOVERY_PERIOD} days")
    print(f"   Output: {OUTPUT_DIR}")
    print("=" * 70)

    # Load data
    print("\n[1/4] Loading merged data...")
    df = pd.read_csv(DATA_PATH)
    df['data'] = pd.to_datetime(df['data'])
    df.sort_values('data', inplace=True)
    df.set_index('data', inplace=True)
    dates = df.index

    # Identify all city columns (2021 format: infected_CityName)
    city_cols = [c for c in df.columns if c.startswith('infected_')]
    city_names = [c.replace('infected_', '').lower().replace(' ', '_') for c in city_cols]
    print(f"  Found {len(city_names)} cities")

    # Process each city
    print(f"\n[2/4] Fitting SIR model for each city...")
    all_results = {}
    summary_rows = []

    for idx, (col, city) in enumerate(zip(city_cols, city_names)):
        # Find population
        N = PROVINCE_POPULATIONS.get(city, None)
        if N is None:
            # Try without accents or with variations
            for key in PROVINCE_POPULATIONS:
                if key.replace('-', '_').replace("'", '') == city.replace('-', '_').replace("'", ''):
                    N = PROVINCE_POPULATIONS[key]
                    break
            if N is None:
                N = 300_000  # Default fallback

        cumulative = df[col].values.astype(float)
        if np.isnan(cumulative).all() or cumulative[-1] == 0:
            print(f"  [{idx+1:3d}/107] {city.title():30s} - SKIPPED (no data)")
            continue

        cumulative = np.nan_to_num(cumulative, nan=0)

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
            'population': N,
            'total_cases': cumulative[-1],
            'infection_rate_pct': cumulative[-1] / N * 100,
            'mean_beta': round(mean_beta, 4),
            'mean_R0': round(R0, 2),
            'fit_error': round(sir_fit['error'], 8),
        })

        print(f"  [{idx+1:3d}/107] {city.title():30s}  N={N:>10,}  "
              f"Cases={cumulative[-1]:>10,.0f}  Beta={mean_beta:.4f}  R0={R0:.2f}")

        # Individual city plot
        plot_city(city, dates, cumulative, sir_fit, beta_t, daily_new, OUTPUT_DIR)

    # Save summary CSV
    print(f"\n[3/4] Saving summary...")
    summary_df = pd.DataFrame(summary_rows)
    summary_df.sort_values('total_cases', ascending=False, inplace=True)
    summary_df.to_csv(os.path.join(OUTPUT_DIR, 'sir_summary_all_cities.csv'), index=False)
    print(f"  Summary CSV saved ({len(summary_df)} cities)")

    # Save beta time series for all cities
    beta_df = pd.DataFrame({'data': dates})
    for city in all_results:
        safe = city.replace("'", "").replace(" ", "_")
        beta_df[f'{safe}_beta'] = all_results[city]['beta_t']
    beta_df.to_csv(os.path.join(OUTPUT_DIR, 'beta_timeseries_all_cities.csv'), index=False)
    print("  Beta timeseries CSV saved")

    # Generate comparison plots
    print(f"\n[4/4] Generating comparison plots...")
    plot_top_cities_comparison(all_results, dates, OUTPUT_DIR, top_n=10)
    print("  Top 10 comparison saved")

    plot_all_beta_heatmap(all_results, dates, OUTPUT_DIR)
    print("  Beta heatmap saved")

    plot_R0_distribution(all_results, OUTPUT_DIR)
    print("  R0 distribution saved")

    # Final summary
    print("\n" + "=" * 70)
    print(f"   DONE! {len(all_results)} cities processed")
    print(f"   Output directory: {OUTPUT_DIR}")
    print(f"   Files generated:")
    print(f"     - {len(all_results)} individual city SIR plots (sir_*.png)")
    print(f"     - comparison_top10.png")
    print(f"     - beta_heatmap_all_cities.png")
    print(f"     - R0_distribution.png")
    print(f"     - sir_summary_all_cities.csv")
    print(f"     - beta_timeseries_all_cities.csv")
    print("=" * 70)

    # Print top 10 and bottom 10
    print("\n  TOP 10 by Total Cases:")
    for _, row in summary_df.head(10).iterrows():
        print(f"    {row['city']:25s}  Cases={row['total_cases']:>12,.0f}  "
              f"R0={row['mean_R0']:.2f}  Inf.Rate={row['infection_rate_pct']:.1f}%")

    print("\n  BOTTOM 10 by Total Cases:")
    for _, row in summary_df.tail(10).iterrows():
        print(f"    {row['city']:25s}  Cases={row['total_cases']:>12,.0f}  "
              f"R0={row['mean_R0']:.2f}  Inf.Rate={row['infection_rate_pct']:.1f}%")


if __name__ == '__main__':
    main()
