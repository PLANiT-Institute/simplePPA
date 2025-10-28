"""
Analysis and reporting utilities.
"""

import pandas as pd


def run_scenario_analysis(load_df, ppa_df, grid_df, emission_df, contract_fee, config, verbose=False):
    """
    Run PPA analysis across multiple coverage scenarios.

    Parameters
    ----------
    load_df : pd.DataFrame
        Load pattern data
    ppa_df : pd.DataFrame
        PPA generation pattern data
    grid_df : pd.DataFrame
        Grid rate data
    emission_df : pd.DataFrame
        Emission factor data (kgCO2e/kWh)
    contract_fee : float
        Grid contract fee
    config : dict
        Configuration dictionary with all parameters
    verbose : bool, optional
        Print detailed statistics

    Returns
    -------
    list
        List of result dictionaries for each scenario
    """
    from .calculator import calculate_ppa_cost

    results_summary = []
    total_electricity_kwh = config['load_capacity_mw'] * 1000 * load_df['load'].sum()

    ppa_range = range(config['ppa_range_start'],
                     config['ppa_range_end'] + 1,
                     config['ppa_range_step'])

    for ppa_percent in ppa_range:
        ppa_coverage = ppa_percent / 100

        # Calculate costs and emissions without ESS (ESS can be added later)
        total_cost, ppa_cost, grid_cost, grid_demand_cost, ess_cost, total_emissions = calculate_ppa_cost(
            load_df, ppa_df, grid_df, emission_df,
            load_capacity_mw=config['load_capacity_mw'],
            ppa_coverage=ppa_coverage,
            contract_fee=contract_fee,
            ppa_price=config['ppa_price'],
            ppa_mintake=config['ppa_mintake'],
            ppa_resell=config['ppa_resell'],
            ppa_resellrate=config['ppa_resellrate'],
            ess_price=config['ess_price'],
            ppa_emission_factor=config.get('ppa_emission_factor', 0.0),
            ess_capacity=0,  # No ESS for base analysis
            verbose=verbose
        )

        # Calculate carbon cost
        carbon_price = config.get('carbon_price', 0.0)
        carbon_cost = (total_emissions / 1000) * carbon_price  # Convert kg to tonnes
        total_cost_with_carbon = total_cost + carbon_cost

        results_summary.append({
            'ppa_percent': ppa_percent,
            'total_cost': total_cost,
            'ppa_cost': ppa_cost,
            'grid_cost': grid_cost,
            'grid_demand_cost': grid_demand_cost,
            'ess_cost': ess_cost,
            'carbon_cost': carbon_cost,
            'total_cost_with_carbon': total_cost_with_carbon,
            'total_electricity_kwh': total_electricity_kwh,
            'total_cost_per_kwh': total_cost / total_electricity_kwh,
            'ppa_cost_per_kwh': ppa_cost / total_electricity_kwh,
            'grid_energy_cost_per_kwh': (grid_cost - grid_demand_cost) / total_electricity_kwh,
            'grid_demand_cost_per_kwh': grid_demand_cost / total_electricity_kwh,
            'ess_cost_per_kwh': ess_cost / total_electricity_kwh,
            'carbon_cost_per_kwh': carbon_cost / total_electricity_kwh,
            'total_cost_with_carbon_per_kwh': total_cost_with_carbon / total_electricity_kwh,
            'total_emissions': total_emissions,
            'emissions_per_kwh': total_emissions / total_electricity_kwh
        })

    return results_summary


def run_ess_analysis(load_df, ppa_df, grid_df, emission_df, contract_fee, config, verbose=False):
    """
    Run PPA analysis with ESS across multiple coverage scenarios.

    Parameters
    ----------
    load_df : pd.DataFrame
        Load pattern data
    ppa_df : pd.DataFrame
        PPA generation pattern data
    grid_df : pd.DataFrame
        Grid rate data
    emission_df : pd.DataFrame
        Emission factor data (kgCO2e/kWh)
    contract_fee : float
        Grid contract fee
    config : dict
        Configuration dictionary with all parameters
    verbose : bool, optional
        Print detailed statistics

    Returns
    -------
    tuple
        (results_ess, ess_capacity_kwh, peak_solar_mw)
    """
    from .calculator import calculate_ppa_cost

    results_ess = []
    total_electricity_kwh = config['load_capacity_mw'] * 1000 * load_df['load'].sum()

    # Calculate ESS capacity based on solar peak generation
    peak_solar_mw = ppa_df['generation'].max() * config['load_capacity_mw']
    ess_capacity_kwh = peak_solar_mw * config['ess_capacity'] * 1000

    ppa_range = range(config['ppa_range_start'],
                     config['ppa_range_end'] + 1,
                     config['ppa_range_step'])

    for ppa_percent in ppa_range:
        ppa_coverage = ppa_percent / 100

        total_cost, ppa_cost, grid_cost, grid_demand_cost, ess_cost, total_emissions = calculate_ppa_cost(
            load_df, ppa_df, grid_df, emission_df,
            load_capacity_mw=config['load_capacity_mw'],
            ppa_coverage=ppa_coverage,
            contract_fee=contract_fee,
            ppa_price=config['ppa_price'],
            ppa_mintake=config['ppa_mintake'],
            ppa_resell=config['ppa_resell'],
            ppa_resellrate=config['ppa_resellrate'],
            ess_price=config['ess_price'],
            ppa_emission_factor=config.get('ppa_emission_factor', 0.0),
            ess_capacity=ess_capacity_kwh,
            verbose=verbose
        )

        # Calculate carbon cost
        carbon_price = config.get('carbon_price', 0.0)
        carbon_cost = (total_emissions / 1000) * carbon_price  # Convert kg to tonnes
        total_cost_with_carbon = total_cost + carbon_cost

        results_ess.append({
            'ppa_percent': ppa_percent,
            'total_cost': total_cost,
            'ppa_cost': ppa_cost,
            'grid_cost': grid_cost,
            'grid_demand_cost': grid_demand_cost,
            'ess_cost': ess_cost,
            'carbon_cost': carbon_cost,
            'total_cost_with_carbon': total_cost_with_carbon,
            'total_electricity_kwh': total_electricity_kwh,
            'total_cost_per_kwh': total_cost / total_electricity_kwh,
            'ppa_cost_per_kwh': ppa_cost / total_electricity_kwh,
            'grid_energy_cost_per_kwh': (grid_cost - grid_demand_cost) / total_electricity_kwh,
            'grid_demand_cost_per_kwh': grid_demand_cost / total_electricity_kwh,
            'ess_cost_per_kwh': ess_cost / total_electricity_kwh,
            'carbon_cost_per_kwh': carbon_cost / total_electricity_kwh,
            'total_cost_with_carbon_per_kwh': total_cost_with_carbon / total_electricity_kwh,
            'total_emissions': total_emissions,
            'emissions_per_kwh': total_emissions / total_electricity_kwh
        })

    return results_ess, ess_capacity_kwh, peak_solar_mw


def find_optimal_scenario(results, include_carbon=True):
    """
    Find the optimal PPA coverage from results.

    Parameters
    ----------
    results : list
        List of result dictionaries
    include_carbon : bool, optional
        If True and carbon costs exist, optimize based on total_cost_with_carbon.
        Otherwise optimize based on total_cost (electricity only).

    Returns
    -------
    tuple
        (optimal_ppa_percent, optimal_cost)
    """
    # Check if carbon pricing is enabled
    has_carbon = results[0].get('carbon_cost', 0) > 0 if results else False

    if include_carbon and has_carbon:
        # Optimize based on total cost including carbon
        optimal_idx = min(range(len(results)), key=lambda i: results[i]['total_cost_with_carbon'])
        optimal_ppa = results[optimal_idx]['ppa_percent']
        optimal_cost = results[optimal_idx]['total_cost_with_carbon']
    else:
        # Optimize based on electricity cost only
        optimal_idx = min(range(len(results)), key=lambda i: results[i]['total_cost'])
        optimal_ppa = results[optimal_idx]['ppa_percent']
        optimal_cost = results[optimal_idx]['total_cost']

    return optimal_ppa, optimal_cost


def create_annual_summary(analysis_df, ppa_range_start, ppa_range_end, ppa_range_step):
    """
    Create annual summary DataFrame for each PPA scenario.

    Parameters
    ----------
    analysis_df : pd.DataFrame
        Analysis DataFrame with scenario columns
    ppa_range_start : int
        Starting PPA coverage percentage
    ppa_range_end : int
        Ending PPA coverage percentage (inclusive)
    ppa_range_step : int
        Step size for PPA coverage

    Returns
    -------
    pd.DataFrame
        Annual summary DataFrame
    """
    annual_summary = []
    total_annual_demand_kwh = (analysis_df['load_mw'].sum() * 1000)

    for ppa_percent in range(ppa_range_start, ppa_range_end + 1, ppa_range_step):
        annual_ppa_cost = analysis_df[f'ppa_cost_{ppa_percent}pct'].sum()
        annual_grid_cost = analysis_df[f'grid_cost_{ppa_percent}pct'].sum()
        annual_total_cost = analysis_df[f'total_cost_{ppa_percent}pct'].sum()

        annual_summary.append({
            'PPA_Coverage (%)': ppa_percent,
            'Annual_PPA_Gen (kWh)': analysis_df[f'ppa_gen_{ppa_percent}pct'].sum(),
            'Annual_Mandatory_PPA (kWh)': analysis_df[f'mandatory_ppa_{ppa_percent}pct'].sum(),
            'Annual_Optional_PPA (kWh)': analysis_df[f'optional_ppa_{ppa_percent}pct'].sum(),
            'Annual_PPA_Purchase (kWh)': analysis_df[f'ppa_purchase_{ppa_percent}pct'].sum(),
            'Annual_PPA_Cost (KRW)': annual_ppa_cost,
            'Annual_Grid_Purchase (kWh)': analysis_df[f'grid_purchase_{ppa_percent}pct'].sum(),
            'Annual_Grid_Cost (KRW)': annual_grid_cost,
            'Annual_PPA_Excess (kWh)': analysis_df[f'ppa_excess_{ppa_percent}pct'].sum(),
            'Annual_Resell_Revenue (KRW)': analysis_df[f'resell_revenue_{ppa_percent}pct'].sum(),
            'Annual_Total_Cost (KRW)': annual_total_cost,
            'Load_Coverage (%)': (analysis_df[f'ppa_purchase_{ppa_percent}pct'].sum() / total_annual_demand_kwh * 100),
            'PPA_Cost (KRW/kWh)': annual_ppa_cost / total_annual_demand_kwh,
            'Grid_Cost (KRW/kWh)': annual_grid_cost / total_annual_demand_kwh,
            'Total_Cost (KRW/kWh)': annual_total_cost / total_annual_demand_kwh
        })

    return pd.DataFrame(annual_summary)


def create_cost_comparison(results_summary):
    """
    Create cost comparison DataFrame from results.

    Parameters
    ----------
    results_summary : list
        List of result dictionaries

    Returns
    -------
    pd.DataFrame
        Cost comparison DataFrame
    """
    cost_comparison = pd.DataFrame({
        'PPA_Coverage (%)': [result['ppa_percent'] for result in results_summary],
        'Total_Cost (KRW)': [result['total_cost'] for result in results_summary],
        'PPA_Cost (KRW)': [result['ppa_cost'] for result in results_summary],
        'Grid_Energy_Cost (KRW)': [result['grid_cost'] - result['grid_demand_cost'] for result in results_summary],
        'Contract_Fee (KRW)': [result['grid_demand_cost'] for result in results_summary],
        'ESS_Cost (KRW)': [result['ess_cost'] for result in results_summary],
        'Total_Cost (KRW/kWh)': [result['total_cost_per_kwh'] for result in results_summary],
        'PPA_Cost (KRW/kWh)': [result['ppa_cost_per_kwh'] for result in results_summary],
        'Grid_Energy_Cost (KRW/kWh)': [result['grid_energy_cost_per_kwh'] for result in results_summary],
        'Contract_Fee (KRW/kWh)': [result['grid_demand_cost_per_kwh'] for result in results_summary],
        'ESS_Cost (KRW/kWh)': [result['ess_cost_per_kwh'] for result in results_summary]
    })

    return cost_comparison


def analyze_peak_hours(analysis_df):
    """
    Analyze peak vs off-peak hours based on grid rates.

    Parameters
    ----------
    analysis_df : pd.DataFrame
        Analysis DataFrame

    Returns
    -------
    dict
        Dictionary with peak hour analysis results
    """
    hourly_avg_rates = analysis_df.groupby('hour')['grid_rate'].mean().sort_values(ascending=False)
    peak_hours = hourly_avg_rates.head(12).index.tolist()
    offpeak_hours = hourly_avg_rates.tail(12).index.tolist()

    peak_data = analysis_df[analysis_df['hour'].isin(peak_hours)]
    offpeak_data = analysis_df[analysis_df['hour'].isin(offpeak_hours)]

    return {
        'peak_hours': sorted(peak_hours),
        'offpeak_hours': sorted(offpeak_hours),
        'peak_avg_rate': peak_data['grid_rate'].mean(),
        'offpeak_avg_rate': offpeak_data['grid_rate'].mean()
    }
