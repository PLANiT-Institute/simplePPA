"""
SimplePPA - Modular Power Purchase Agreement Analysis Tool

This is the main entry point for running PPA analysis.
All configuration is done through the config dictionary or config file.
"""

import libs.KEPCOutils as kepco
from libs.config import get_default_config, validate_config, print_config
from libs.data_processor import (
    load_pattern_data,
    create_analysis_dataframe,
    generate_scenario_columns,
    create_long_format_dataframe
)
from libs.analyzer import (
    run_scenario_analysis,
    run_ess_analysis,
    find_optimal_scenario,
    create_annual_summary,
    create_cost_comparison,
    analyze_peak_hours
)
from libs.exporter import (
    export_to_excel,
    print_analysis_summary,
    print_results_table,
    print_peak_analysis
)


def main():
    """
    Main analysis function.
    """
    # =====================================================================
    # CONFIGURATION
    # =====================================================================
    # Load configuration (you can modify these values or load from file)
    config = get_default_config()

    # Override any parameters here if needed
    # Example: config['load_capacity_mw'] = 5000
    # Example: config['ppa_range_start'] = 0
    # Example: config['ppa_range_end'] = 150
    # Example: config['ppa_range_step'] = 5

    # Or load from a config file:
    # from libs.config import load_config_from_file
    # config = load_config_from_file('my_config.json')

    # Validate configuration
    is_valid, errors = validate_config(config)
    if not is_valid:
        print("Configuration errors:")
        for error in errors:
            print(f"  - {error}")
        return

    # Print configuration
    print_config(config)

    # =====================================================================
    # DATA LOADING
    # =====================================================================
    print("\n=== LOADING DATA ===")

    # Load pattern data
    print(f"Loading pattern data from {config['pattern_file']}...")
    load_df, solar_df, ppa_df = load_pattern_data(config['pattern_file'])

    # Process Grid data
    print(f"Loading KEPCO data from {config['kepco_file']}...")
    grid_df, contract_fee = kepco.process_kepco_data(
        config['kepco_file'],
        config['kepco_year'],
        config['kepco_tariff']
    )

    print("Data loaded successfully!")

    # =====================================================================
    # SCENARIO ANALYSIS (No ESS)
    # =====================================================================
    print("\n=== RUNNING SCENARIO ANALYSIS (No ESS) ===")

    # Run analysis across all PPA coverage scenarios
    results_summary = run_scenario_analysis(
        load_df, ppa_df, grid_df, contract_fee, config,
        verbose=config['verbose']
    )

    # Print results table
    print_results_table(results_summary, "PPA Coverage Analysis (No ESS)")

    # Find optimal scenario
    optimal_ppa, optimal_cost = find_optimal_scenario(results_summary)
    print(f"\nOptimal PPA coverage (No ESS): {optimal_ppa}% with total cost: {optimal_cost:,.0f} KRW")

    # =====================================================================
    # ESS ANALYSIS (if enabled)
    # =====================================================================
    results_ess = None
    optimal_ess_ppa = None
    optimal_ess_cost = None
    ess_capacity_kwh = None
    peak_solar_mw = None

    if config['ess_include']:
        print(f"\n=== RUNNING ESS ANALYSIS ===")

        results_ess, ess_capacity_kwh, peak_solar_mw = run_ess_analysis(
            load_df, ppa_df, grid_df, contract_fee, config,
            verbose=config['verbose']
        )

        # Print results table
        print_results_table(
            results_ess,
            f"PPA Coverage Analysis (With ESS: {config['ess_capacity']*100:.0f}% solar capacity)"
        )

        # Find optimal scenario with ESS
        optimal_ess_ppa, optimal_ess_cost = find_optimal_scenario(results_ess)
        print(f"\nOptimal PPA coverage (With ESS): {optimal_ess_ppa}% with total cost: {optimal_ess_cost:,.0f} KRW")
        print(f"ESS Capacity: {ess_capacity_kwh:.1f} kWh ({config['ess_capacity']*100:.0f}% of peak solar {peak_solar_mw:.1f} MW)")
        print(f"ESS benefit: {optimal_cost - optimal_ess_cost:,.0f} KRW savings")

    # =====================================================================
    # DETAILED ANALYSIS & EXPORT
    # =====================================================================
    if config['export_long_format']:
        print(f"\n=== GENERATING DETAILED ANALYSIS ===")
        print(f"Analysis period: {config['start_date']} to {config['end_date']}")

        # Create analysis DataFrame
        analysis_df = create_analysis_dataframe(
            grid_df, load_df, ppa_df,
            config['start_date'], config['end_date'],
            config['load_capacity_mw']
        )
        print(f"Data filtered to {len(analysis_df):,} hours ({len(analysis_df)/24:.1f} days)")

        # Generate scenario columns
        print("Calculating hourly patterns for all PPA scenarios...")
        analysis_df = generate_scenario_columns(
            analysis_df,
            config['ppa_range_start'],
            config['ppa_range_end'],
            config['ppa_range_step'],
            config['load_capacity_mw'],
            config['ppa_price'],
            config['ppa_mintake'],
            config['ppa_resell'],
            config['ppa_resellrate']
        )
        num_scenarios = len(range(config['ppa_range_start'], config['ppa_range_end']+1, config['ppa_range_step']))
        print(f"Generated hourly patterns for {num_scenarios} PPA scenarios")

        # Create long-format DataFrame
        print("Creating long-format DataFrame for pivot analysis...")
        long_df = create_long_format_dataframe(
            analysis_df,
            config['ppa_range_start'],
            config['ppa_range_end'],
            config['ppa_range_step']
        )
        print(f"Long-format data contains {len(long_df):,} rows covering all scenarios")

        # Create annual summary
        annual_summary_df = create_annual_summary(
            analysis_df,
            config['ppa_range_start'],
            config['ppa_range_end'],
            config['ppa_range_step']
        )

        # Create cost comparison
        cost_comparison_df = create_cost_comparison(results_summary)

        # Export to Excel
        print(f"\nExporting results to {config['output_file']}...")
        export_to_excel(
            config['output_file'],
            long_df,
            annual_summary_df,
            cost_comparison_df
        )
        print(f"Analysis saved to: {config['output_file']}")
        print(f"Sheets: PPA_Analysis_Data (long-format), Annual_Summary, Cost_Analysis")

        # Peak hour analysis
        peak_analysis = analyze_peak_hours(analysis_df)
        print_peak_analysis(peak_analysis, config['ppa_price'])

        # Print final summary
        print_analysis_summary(
            config, analysis_df, results_summary,
            optimal_ppa, optimal_cost,
            results_ess, optimal_ess_ppa, optimal_ess_cost,
            ess_capacity_kwh, peak_solar_mw
        )

    else:
        # Just print summary without detailed export
        print(f"\n=== SUMMARY ===")
        print(f"ESS included: {config['ess_include']}")
        print(f"Resell enabled: {config['ppa_resell']}, Resell rate: {config['ppa_resellrate']}")
        print(f"No ESS  - Optimal: {optimal_ppa}% PPA, Cost: {optimal_cost:,.0f} KRW")
        if config['ess_include']:
            print(f"With ESS - Optimal: {optimal_ess_ppa}% PPA, Cost: {optimal_ess_cost:,.0f} KRW")
            print(f"ESS benefit: {optimal_cost - optimal_ess_cost:,.0f} KRW savings")

    print("\n=== ANALYSIS COMPLETE ===")


if __name__ == "__main__":
    main()
