"""
Excel export utilities.
"""

import pandas as pd


def export_to_excel(filename, long_df, annual_summary_df, cost_comparison_df):
    """
    Export analysis results to Excel file.

    Parameters
    ----------
    filename : str
        Output filename
    long_df : pd.DataFrame
        Long-format analysis data
    annual_summary_df : pd.DataFrame
        Annual summary data
    cost_comparison_df : pd.DataFrame
        Cost comparison data

    Returns
    -------
    None
    """
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Primary output: Long-format data for pivot analysis
        long_df.to_excel(writer, sheet_name='PPA_Analysis_Data', index=False)

        # Annual summary
        annual_summary_df.to_excel(writer, sheet_name='Annual_Summary', index=False)

        # Cost analysis
        cost_comparison_df.to_excel(writer, sheet_name='Cost_Analysis', index=False)


def print_analysis_summary(config, analysis_df, results_summary, optimal_ppa, optimal_cost,
                          results_ess=None, optimal_ess_ppa=None, optimal_ess_cost=None,
                          ess_capacity_kwh=None, peak_solar_mw=None):
    """
    Print comprehensive analysis summary to console.

    Parameters
    ----------
    config : dict
        Configuration dictionary
    analysis_df : pd.DataFrame
        Analysis DataFrame
    results_summary : list
        Results without ESS
    optimal_ppa : int
        Optimal PPA percentage without ESS
    optimal_cost : float
        Optimal cost without ESS
    results_ess : list, optional
        Results with ESS
    optimal_ess_ppa : int, optional
        Optimal PPA percentage with ESS
    optimal_ess_cost : float, optional
        Optimal cost with ESS
    ess_capacity_kwh : float, optional
        ESS capacity in kWh
    peak_solar_mw : float, optional
        Peak solar in MW

    Returns
    -------
    None
    """
    # Basic analysis info
    print(f"\n=== ANALYSIS SUMMARY ===")
    print(f"Analysis period: {config['start_date']} to {config['end_date']} ({len(analysis_df)/24:.1f} days)")
    print(f"Load capacity: {config['load_capacity_mw']} MW")
    print(f"Total load: {(analysis_df['load_mw'].sum() * 1000):,.0f} kWh")
    print(f"Average Grid rate: {analysis_df['grid_rate'].mean():.1f} KRW/kWh")
    print(f"Grid rate range: {analysis_df['grid_rate'].min():.1f} - {analysis_df['grid_rate'].max():.1f} KRW/kWh")
    print(f"PPA price: {config['ppa_price']} KRW/kWh")
    print(f"Average savings per kWh: {analysis_df['grid_rate'].mean() - config['ppa_price']:.1f} KRW/kWh")

    # Optimal results
    print(f"\n=== OPTIMAL RESULTS ===")
    print(f"No ESS  - Optimal: {optimal_ppa}% PPA, Cost: {optimal_cost:,.0f} KRW")

    if results_ess is not None:
        print(f"With ESS - Optimal: {optimal_ess_ppa}% PPA, Cost: {optimal_ess_cost:,.0f} KRW")
        print(f"ESS Capacity: {ess_capacity_kwh:.1f} kWh ({config['ess_capacity']*100:.0f}% of peak solar {peak_solar_mw:.1f} MW)")
        print(f"ESS benefit: {optimal_cost - optimal_ess_cost:,.0f} KRW savings")
    else:
        print(f"With ESS - Analysis skipped (ESS disabled)")

    # Configuration summary
    print(f"\n=== CONFIGURATION ===")
    print(f"ESS included: {config['ess_include']}")
    print(f"Resell enabled: {config['ppa_resell']}, Resell rate: {config['ppa_resellrate']}")
    print(f"Minimum take: {config['ppa_mintake']*100:.0f}%")
    print(f"PPA coverage range: {config['ppa_range_start']}% to {config['ppa_range_end']}% (step: {config['ppa_range_step']}%)")


def print_results_table(results_summary, title="PPA Coverage Analysis"):
    """
    Print results table to console.

    Parameters
    ----------
    results_summary : list
        List of result dictionaries
    title : str, optional
        Table title

    Returns
    -------
    None
    """
    print(f"\n=== {title} ===")
    print("PPA% | Total Cost/kWh | PPA Cost/kWh | Grid Energy/kWh | Contract/kWh | ESS Cost/kWh")
    print("-" * 85)

    for result in results_summary:
        print(f"{result['ppa_percent']:3d}% | "
              f"{result['total_cost_per_kwh']:14.2f} | "
              f"{result['ppa_cost_per_kwh']:12.2f} | "
              f"{result['grid_energy_cost_per_kwh']:15.2f} | "
              f"{result['grid_demand_cost_per_kwh']:12.2f} | "
              f"{result['ess_cost_per_kwh']:12.2f}")


def print_peak_analysis(peak_analysis, ppa_price):
    """
    Print peak hour analysis.

    Parameters
    ----------
    peak_analysis : dict
        Peak analysis results
    ppa_price : float
        PPA price

    Returns
    -------
    None
    """
    print(f"\n=== PEAK HOUR ANALYSIS ===")
    print(f"Peak hours {peak_analysis['peak_hours']}: Avg Grid rate = {peak_analysis['peak_avg_rate']:.1f} KRW/kWh")
    print(f"Off-peak hours {peak_analysis['offpeak_hours']}: Avg Grid rate = {peak_analysis['offpeak_avg_rate']:.1f} KRW/kWh")
    print(f"Peak hour savings with PPA: {peak_analysis['peak_avg_rate'] - ppa_price:.1f} KRW/kWh")
