"""
CGP Analysis with Lag=15 - Comparison Study (2020v2)
====================================================
Runs the same CGP analysis as cgp_analysis_2020v2.py but with DEFAULT_N_LAGS=15
instead of 7 to study the effect of increased temporal lookback.

Output: 2020v2/cgp_results_lag15/
"""

import sys
import os

# Add the 2020v2 directory to path so we can import the original module
sys.path.insert(0, r"c:\Users\Utente\Documents\GitHub\gcpaida\2020v2")

# Import everything from the original CGP module
from cgp_analysis_2020v2 import (
    CartesianGeneticProgramming, prepare_data, build_lagged_inputs,
    build_self_lagged_inputs, run_cgp_for_city, analyze_inter_city,
    run_lag_sweep, plot_lag_vs_r2, run_lockdown_analysis,
    plot_lockdown_comparison, plot_lockdown_active_cities,
    plot_connection_heatmap, plot_network_graph, plot_cgp_weights,
    plot_fitness_convergence,
    CITY_GROUPS, ALL_CITIES, CITY_REGION_MAP, REGION_COLORS,
    N_ROWS, N_COLS, N_OUTPUTS, N_GENERATIONS, LAMBDA, MUTATION_RATE,
    LEVELS_BACK, R2_THRESHOLD, LAG_SWEEP_VALUES, LOCKDOWN_DATE, DATA_PATH
)

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

# ==============================================================================
# CONFIGURATION - OVERRIDE LAG TO 15
# ==============================================================================
NEW_DEFAULT_LAG = 15
OUTPUT_DIR_LAG15 = r"c:\Users\Utente\Documents\GitHub\gcpaida\2020v2\cgp_results_lag15"
os.makedirs(OUTPUT_DIR_LAG15, exist_ok=True)

# Also load lag=7 results for comparison
LAG7_RESULTS_DIR = r"c:\Users\Utente\Documents\GitHub\gcpaida\2020v2\cgp_results"


def main():
    print("=" * 70)
    print("   CGP Analysis - LAG=15 Comparison Study")
    print("   Same 15 cities, same addition-only functions")
    print(f"   DEFAULT LAG: {NEW_DEFAULT_LAG} (original was 7)")
    print("=" * 70)

    # 1. Load data
    print("\n[1/7] Loading data...")
    daily_new, city_names = prepare_data(DATA_PATH, set(ALL_CITIES))
    city_names = [c for c in ALL_CITIES if c in city_names]
    print(f"  Loaded {len(city_names)} cities, {len(daily_new)} days")

    # 2. Main CGP analysis with lag=15
    print(f"\n[2/7] Running main CGP analysis (lag={NEW_DEFAULT_LAG})...")
    cgp_results = analyze_inter_city(daily_new, city_names, n_lags=NEW_DEFAULT_LAG)

    # Build connections dict
    all_connections = {}
    for target, result in cgp_results.items():
        for source in result['active_cities']:
            pair = tuple(sorted([target, source]))
            current = all_connections.get(pair, 0)
            all_connections[pair] = max(current, result['r2'])

    # 3. Lag sweep (same as before for completeness)
    print(f"\n[3/7] Running lag sweep analysis...")
    lag_r2_results = run_lag_sweep(daily_new, city_names, LAG_SWEEP_VALUES)
    plot_lag_vs_r2(lag_r2_results, LAG_SWEEP_VALUES, OUTPUT_DIR_LAG15)

    # 4. Lockdown analysis with lag=15
    print(f"\n[4/7] Running pre/post lockdown analysis (lag={NEW_DEFAULT_LAG})...")
    lockdown_results = run_lockdown_analysis(
        daily_new, city_names, LOCKDOWN_DATE, n_lags=NEW_DEFAULT_LAG
    )

    # 5. Visualizations
    print(f"\n[5/7] Generating visualizations...")
    plot_connection_heatmap(all_connections, city_names, OUTPUT_DIR_LAG15)
    plot_network_graph(all_connections, city_names, OUTPUT_DIR_LAG15)
    plot_cgp_weights(cgp_results, city_names, OUTPUT_DIR_LAG15)
    plot_fitness_convergence(cgp_results, OUTPUT_DIR_LAG15)

    # 6. Lockdown plots
    print(f"\n[6/7] Generating lockdown comparison plots...")
    plot_lockdown_comparison(lockdown_results, city_names, OUTPUT_DIR_LAG15)
    plot_lockdown_active_cities(lockdown_results, city_names, OUTPUT_DIR_LAG15)

    # 7. Save CSVs and comparison
    print(f"\n[7/7] Saving summary data...")

    # Connections CSV
    conn_rows = []
    for (c1, c2), strength in sorted(all_connections.items(), key=lambda x: x[1], reverse=True):
        conn_rows.append({
            'city_1': c1.title(), 'city_2': c2.title(),
            'region_1': CITY_REGION_MAP.get(c1, ''),
            'region_2': CITY_REGION_MAP.get(c2, ''),
            'connection_strength_R2': round(strength, 4),
            'is_significant': strength > R2_THRESHOLD,
        })
    conn_df = pd.DataFrame(conn_rows)
    conn_df.to_csv(os.path.join(OUTPUT_DIR_LAG15, 'all_connections.csv'), index=False)
    print(f"  Saved: all_connections.csv ({len(conn_df)} pairs)")

    # City summary CSV
    summary_rows = []
    for city in city_names:
        r2 = cgp_results.get(city, {}).get('r2', 0)
        active = cgp_results.get(city, {}).get('active_cities', [])
        funcs = cgp_results.get(city, {}).get('active_functions', [])
        summary_rows.append({
            'city': city.title(), 'region': CITY_REGION_MAP.get(city, ''),
            'cgp_r2': round(r2, 4), 'n_active_inputs': len(active),
            'active_cities': ', '.join([c.title() for c in active]),
            'active_functions': ', '.join([str(f[0]) if isinstance(f, (tuple, list)) else str(f) for f in funcs]),
        })
    summary_df = pd.DataFrame(summary_rows)
    summary_df.to_csv(os.path.join(OUTPUT_DIR_LAG15, 'city_summary.csv'), index=False)
    print(f"  Saved: city_summary.csv")

    # Lag sweep CSV
    lag_rows = []
    for city in city_names:
        row = {'city': city.title(), 'region': CITY_REGION_MAP.get(city, '')}
        for lag in LAG_SWEEP_VALUES:
            row[f'lag_{lag}'] = round(lag_r2_results.get(city, {}).get(lag, 0), 4)
        lag_rows.append(row)
    lag_df = pd.DataFrame(lag_rows)
    lag_df.to_csv(os.path.join(OUTPUT_DIR_LAG15, 'lag_sweep_results.csv'), index=False)
    print(f"  Saved: lag_sweep_results.csv")

    # Lockdown CSV
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
    lockdown_df.to_csv(os.path.join(OUTPUT_DIR_LAG15, 'lockdown_comparison.csv'), index=False)
    print(f"  Saved: lockdown_comparison.csv")

    # === COMPARISON PLOT: lag=7 vs lag=15 ===
    print("\n  Generating lag=7 vs lag=15 comparison plots...")

    # Load lag=7 results
    lag7_summary_path = os.path.join(LAG7_RESULTS_DIR, 'city_summary.csv')
    if os.path.exists(lag7_summary_path):
        lag7_df = pd.read_csv(lag7_summary_path)
        lag15_df = summary_df.copy()

        # Comparison bar chart
        fig, axes = plt.subplots(1, 2, figsize=(22, 8))

        # R2 comparison
        ax = axes[0]
        x = np.arange(len(city_names))
        width = 0.35
        r2_7 = [lag7_df[lag7_df['city'] == c.title()]['cgp_r2'].values[0]
                 if c.title() in lag7_df['city'].values else 0 for c in city_names]
        r2_15 = [lag15_df[lag15_df['city'] == c.title()]['cgp_r2'].values[0]
                  if c.title() in lag15_df['city'].values else 0 for c in city_names]

        colors_7 = ['#2196F3'] * len(city_names)
        colors_15 = ['#FF9800'] * len(city_names)

        ax.bar(x - width/2, r2_7, width, label='Lag=7', color='#2196F3', alpha=0.85)
        ax.bar(x + width/2, r2_15, width, label='Lag=15', color='#FF9800', alpha=0.85)
        ax.axhline(y=R2_THRESHOLD, color='red', ls='--', lw=2, label=f'R2={R2_THRESHOLD}')
        ax.set_xlabel('City', fontsize=12)
        ax.set_ylabel('R2', fontsize=12)
        ax.set_title('R2 Comparison: Lag=7 vs Lag=15', fontsize=14, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels([c.title() for c in city_names], rotation=45, ha='right', fontsize=8)
        ax.legend(fontsize=11)
        ax.grid(True, alpha=0.3, axis='y')
        ax.set_ylim(0, 1.15)

        # Active inputs comparison
        ax = axes[1]
        active_7 = [lag7_df[lag7_df['city'] == c.title()]['n_active_inputs'].values[0]
                     if c.title() in lag7_df['city'].values else 0 for c in city_names]
        active_15 = [lag15_df[lag15_df['city'] == c.title()]['n_active_inputs'].values[0]
                      if c.title() in lag15_df['city'].values else 0 for c in city_names]

        ax.bar(x - width/2, active_7, width, label='Lag=7', color='#2196F3', alpha=0.85)
        ax.bar(x + width/2, active_15, width, label='Lag=15', color='#FF9800', alpha=0.85)
        ax.set_xlabel('City', fontsize=12)
        ax.set_ylabel('Active Input Cities', fontsize=12)
        ax.set_title('Active Inputs: Lag=7 vs Lag=15', fontsize=14, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels([c.title() for c in city_names], rotation=45, ha='right', fontsize=8)
        ax.legend(fontsize=11)
        ax.grid(True, alpha=0.3, axis='y')

        plt.suptitle('CGP Lag Parameter Study: 7 vs 15 Days Lookback',
                     fontsize=16, fontweight='bold', y=1.02)
        plt.tight_layout()
        plt.savefig(os.path.join(OUTPUT_DIR_LAG15, 'lag7_vs_lag15_comparison.png'),
                    dpi=150, bbox_inches='tight')
        plt.close()
        print("  Saved: lag7_vs_lag15_comparison.png")

        # Detailed delta table
        comparison_rows = []
        for city in city_names:
            r7_val = 0
            a7_val = 0
            r15_val = 0
            a15_val = 0
            if city.title() in lag7_df['city'].values:
                r7_val = lag7_df[lag7_df['city'] == city.title()]['cgp_r2'].values[0]
                a7_val = lag7_df[lag7_df['city'] == city.title()]['n_active_inputs'].values[0]
            if city.title() in lag15_df['city'].values:
                r15_val = lag15_df[lag15_df['city'] == city.title()]['cgp_r2'].values[0]
                a15_val = lag15_df[lag15_df['city'] == city.title()]['n_active_inputs'].values[0]

            comparison_rows.append({
                'city': city.title(),
                'region': CITY_REGION_MAP.get(city, ''),
                'r2_lag7': round(r7_val, 4),
                'r2_lag15': round(r15_val, 4),
                'r2_delta': round(r15_val - r7_val, 4),
                'r2_improved': 'Yes' if r15_val > r7_val else 'No',
                'active_lag7': int(a7_val),
                'active_lag15': int(a15_val),
            })

        comp_df = pd.DataFrame(comparison_rows)
        comp_df.to_csv(os.path.join(OUTPUT_DIR_LAG15, 'lag7_vs_lag15_comparison.csv'), index=False)
        print("  Saved: lag7_vs_lag15_comparison.csv")

    # Lockdown comparison: lag7 vs lag15
    lag7_lockdown_path = os.path.join(LAG7_RESULTS_DIR, 'lockdown_comparison.csv')
    if os.path.exists(lag7_lockdown_path):
        ld7 = pd.read_csv(lag7_lockdown_path)
        ld15 = lockdown_df.copy()

        fig, axes = plt.subplots(1, 3, figsize=(24, 8))

        # Pre-lockdown comparison
        ax = axes[0]
        x = np.arange(len(city_names))
        pre7 = [ld7[ld7['city'] == c.title()]['pre_lockdown_r2'].values[0]
                if c.title() in ld7['city'].values else 0 for c in city_names]
        pre15 = [ld15[ld15['city'] == c.title()]['pre_lockdown_r2'].values[0]
                 if c.title() in ld15['city'].values else 0 for c in city_names]
        ax.bar(x - width/2, pre7, width, label='Lag=7', color='#2196F3', alpha=0.85)
        ax.bar(x + width/2, pre15, width, label='Lag=15', color='#FF9800', alpha=0.85)
        ax.set_title('Pre-Lockdown R2', fontsize=13, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels([c.title() for c in city_names], rotation=45, ha='right', fontsize=7)
        ax.legend(); ax.grid(True, alpha=0.3, axis='y'); ax.set_ylim(0, 1.15)

        # During lockdown
        ax = axes[1]
        dur7 = [ld7[ld7['city'] == c.title()].get('during_lockdown_inter_r2', pd.Series([0])).values[0]
                if c.title() in ld7['city'].values else 0 for c in city_names]
        dur15 = [ld15[ld15['city'] == c.title()]['during_lockdown_inter_r2'].values[0]
                 if c.title() in ld15['city'].values else 0 for c in city_names]
        ax.bar(x - width/2, dur7, width, label='Lag=7', color='#2196F3', alpha=0.85)
        ax.bar(x + width/2, dur15, width, label='Lag=15', color='#FF9800', alpha=0.85)
        ax.set_title('During Lockdown R2 (Inter-city)', fontsize=13, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels([c.title() for c in city_names], rotation=45, ha='right', fontsize=7)
        ax.legend(); ax.grid(True, alpha=0.3, axis='y'); ax.set_ylim(0, 1.15)

        # After easing
        ax = axes[2]
        aft7 = [ld7[ld7['city'] == c.title()].get('after_easing_inter_r2', pd.Series([0])).values[0]
                if c.title() in ld7['city'].values else 0 for c in city_names]
        aft15 = [ld15[ld15['city'] == c.title()]['after_easing_inter_r2'].values[0]
                 if c.title() in ld15['city'].values else 0 for c in city_names]
        ax.bar(x - width/2, aft7, width, label='Lag=7', color='#2196F3', alpha=0.85)
        ax.bar(x + width/2, aft15, width, label='Lag=15', color='#FF9800', alpha=0.85)
        ax.set_title('After Easing R2 (Inter-city)', fontsize=13, fontweight='bold')
        ax.set_xticks(x)
        ax.set_xticklabels([c.title() for c in city_names], rotation=45, ha='right', fontsize=7)
        ax.legend(); ax.grid(True, alpha=0.3, axis='y'); ax.set_ylim(0, 1.15)

        plt.suptitle('Lockdown Analysis: Lag=7 vs Lag=15', fontsize=16, fontweight='bold')
        plt.tight_layout()
        plt.savefig(os.path.join(OUTPUT_DIR_LAG15, 'lockdown_lag7_vs_lag15.png'),
                    dpi=150, bbox_inches='tight')
        plt.close()
        print("  Saved: lockdown_lag7_vs_lag15.png")

    # Print summary
    n_sig = sum(1 for v in all_connections.values() if v > R2_THRESHOLD)
    print("\n" + "=" * 70)
    print("   CGP LAG=15 RESULTS SUMMARY")
    print("=" * 70)
    print(f"  Default Lag: {NEW_DEFAULT_LAG} (original: 7)")
    print(f"  R2 threshold: {R2_THRESHOLD}")
    print(f"  Total pairs analyzed: {len(all_connections)}")
    print(f"  Significant connections (R2>{R2_THRESHOLD}): {n_sig}")
    print(f"\n  Per-city R2 (lag={NEW_DEFAULT_LAG}):")
    for city in city_names:
        r2 = cgp_results.get(city, {}).get('r2', 0)
        n_act = cgp_results.get(city, {}).get('n_active', 0)
        status = "SIGNIFICANT" if r2 > R2_THRESHOLD else "weak"
        print(f"    {city.title():25s}  R2={r2:.4f}  Active={n_act}  [{status}]")

    print(f"\n  LOCKDOWN (lag={NEW_DEFAULT_LAG}):")
    print(f"  {'City':25s}  {'Pre':>6s}  {'Dur(I)':>6s}  {'Dur(S)':>6s}  {'Aft(I)':>6s}  {'Aft(S)':>6s}")
    print(f"  {'-'*67}")
    for city in city_names:
        pre = lockdown_results['pre'].get(city, {}).get('r2', 0)
        dur_i = lockdown_results.get('during_inter', {}).get(city, {}).get('r2', 0)
        dur_s = lockdown_results.get('during_self', {}).get(city, {}).get('r2', 0)
        aft_i = lockdown_results.get('after_inter', {}).get(city, {}).get('r2', 0)
        aft_s = lockdown_results.get('after_self', {}).get(city, {}).get('r2', 0)
        print(f"    {city.title():25s}  {pre:6.3f}  {dur_i:6.3f}  {dur_s:6.3f}  {aft_i:6.3f}  {aft_s:6.3f}")

    print(f"\n  Output: {OUTPUT_DIR_LAG15}")
    print("=" * 70)


if __name__ == '__main__':
    main()
