"""
Data loading and processing utilities.
"""

import pandas as pd
import numpy as np


def load_pattern_data(filepath):
    """
    Load load and solar generation patterns from Excel file.

    Parameters
    ----------
    filepath : str
        Path to pattern Excel file

    Returns
    -------
    tuple
        (load_df, solar_df, ppa_df, emission_df) - DataFrames with normalized patterns and emission factors
    """
    pattern_df = pd.read_excel(filepath, index_col=0)

    # Save load and solar to separate DataFrames
    load_df = pattern_df[['load']]
    solar_df = pattern_df[['solar']]

    # Create ppa_df from solar_df
    ppa_df = solar_df.copy()
    ppa_df.columns = ['generation']

    # Load emission data if available (kgCO2e/kWh for grid electricity)
    if 'emission' in pattern_df.columns:
        emission_df = pattern_df[['emission']]
    else:
        # Default to zero if emission data not available
        emission_df = pd.DataFrame({'emission': [0.0] * len(pattern_df)}, index=pattern_df.index)

    return load_df, solar_df, ppa_df, emission_df


def create_analysis_dataframe(grid_df, load_df, ppa_df, emission_df, start_date, end_date, load_capacity_mw):
    """
    Create base analysis DataFrame with filtered time range.

    Parameters
    ----------
    grid_df : pd.DataFrame
        Grid data with datetime index
    load_df : pd.DataFrame
        Load pattern data
    ppa_df : pd.DataFrame
        PPA generation pattern data
    emission_df : pd.DataFrame
        Emission factor data (kgCO2e/kWh)
    start_date : str
        Start date (YYYY-MM-DD)
    end_date : str
        End date (YYYY-MM-DD)
    load_capacity_mw : float
        Load capacity in MW

    Returns
    -------
    pd.DataFrame
        Analysis DataFrame with datetime, load, grid_rate, solar_generation, emission_factor
    """
    # Create base analysis DataFrame
    analysis_df = pd.DataFrame({
        'datetime': grid_df.index,
        'load': load_df['load'].values,
        'grid_rate': grid_df['rate'].values,
        'solar_generation': ppa_df['generation'].values,
        'emission_factor': emission_df['emission'].values
    })

    # Filter data to user-defined timeframe
    mask = (analysis_df['datetime'] >= start_date) & (analysis_df['datetime'] <= end_date)
    analysis_df = analysis_df[mask].copy()

    # Add time components for analysis
    analysis_df['hour'] = analysis_df['datetime'].dt.hour
    analysis_df['month'] = analysis_df['datetime'].dt.month
    analysis_df['day_of_year'] = analysis_df['datetime'].dt.dayofyear

    # Scale load and solar data to MW
    analysis_df['load_mw'] = analysis_df['load'] * load_capacity_mw
    analysis_df['solar_generation_mw'] = analysis_df['solar_generation'] * load_capacity_mw

    return analysis_df


def generate_scenario_columns(analysis_df, ppa_range_start, ppa_range_end, ppa_range_step,
                              load_capacity_mw, ppa_price, ppa_mintake, ppa_resell, ppa_resellrate):
    """
    Generate hourly data columns for all PPA scenarios.

    Parameters
    ----------
    analysis_df : pd.DataFrame
        Base analysis DataFrame
    ppa_range_start : int
        Starting PPA coverage percentage
    ppa_range_end : int
        Ending PPA coverage percentage (inclusive)
    ppa_range_step : int
        Step size for PPA coverage
    load_capacity_mw : float
        Load capacity in MW
    ppa_price : float
        PPA price
    ppa_mintake : float
        Minimum take percentage
    ppa_resell : bool
        Resale allowed flag
    ppa_resellrate : float
        Resale rate

    Returns
    -------
    pd.DataFrame
        Analysis DataFrame with scenario columns added
    """
    scenario_columns = {}

    for ppa_percent in range(ppa_range_start, ppa_range_end + 1, ppa_range_step):
        ppa_coverage = ppa_percent / 100

        # Calculate scaled PPA generation for this scenario
        ppa_generation_mw = analysis_df['solar_generation'] * load_capacity_mw * ppa_coverage

        # Convert to kWh
        load_kwh = analysis_df['load_mw'] * 1000
        ppa_generation_kwh = ppa_generation_mw * 1000

        # Calculate mandatory PPA purchase
        mandatory_ppa = ppa_generation_kwh * ppa_mintake

        # Calculate optional PPA purchase
        optional_ppa_available = ppa_generation_kwh - mandatory_ppa
        remaining_load_after_mandatory = (load_kwh - mandatory_ppa).clip(lower=0)

        optional_ppa = pd.Series(0.0, index=analysis_df.index, dtype=float)
        if ppa_mintake < 1:
            cheaper_mask = ppa_price < analysis_df['grid_rate']
            need_energy_mask = remaining_load_after_mandatory > 0
            can_buy_mask = cheaper_mask & need_energy_mask

            optional_ppa[can_buy_mask] = np.minimum(
                optional_ppa_available[can_buy_mask],
                remaining_load_after_mandatory[can_buy_mask]
            )

        # Total PPA purchase
        ppa_purchase = mandatory_ppa + optional_ppa

        # Calculate PPA costs
        ppa_cost = ppa_purchase * ppa_price

        # Calculate remaining load after PPA
        remaining_load = load_kwh - ppa_purchase

        # Calculate excess PPA
        ppa_excess = (-remaining_load).clip(lower=0)

        # Calculate Grid purchases
        grid_purchase = remaining_load.clip(lower=0)

        # Calculate Grid costs
        grid_cost = grid_purchase * analysis_df['grid_rate']

        # Calculate resell revenue if enabled
        if ppa_resell:
            resell_revenue = ppa_excess * ppa_price * ppa_resellrate
            ppa_cost -= resell_revenue
        else:
            resell_revenue = pd.Series(0, index=analysis_df.index)

        # Calculate total hourly cost
        total_cost = ppa_cost + grid_cost

        # Store all columns for this scenario
        scenario_columns[f'ppa_gen_{ppa_percent}pct'] = ppa_generation_kwh
        scenario_columns[f'mandatory_ppa_{ppa_percent}pct'] = mandatory_ppa
        scenario_columns[f'optional_ppa_{ppa_percent}pct'] = optional_ppa
        scenario_columns[f'ppa_purchase_{ppa_percent}pct'] = ppa_purchase
        scenario_columns[f'ppa_cost_{ppa_percent}pct'] = ppa_cost
        scenario_columns[f'ppa_excess_{ppa_percent}pct'] = ppa_excess
        scenario_columns[f'grid_purchase_{ppa_percent}pct'] = grid_purchase
        scenario_columns[f'grid_cost_{ppa_percent}pct'] = grid_cost
        scenario_columns[f'resell_revenue_{ppa_percent}pct'] = resell_revenue
        scenario_columns[f'total_cost_{ppa_percent}pct'] = total_cost

    # Add all scenario columns at once
    scenario_df = pd.DataFrame(scenario_columns, index=analysis_df.index)
    analysis_df = pd.concat([analysis_df, scenario_df], axis=1)

    return analysis_df


def create_long_format_dataframe(analysis_df, ppa_range_start, ppa_range_end, ppa_range_step):
    """
    Create long-format DataFrame for pivot table analysis.

    Parameters
    ----------
    analysis_df : pd.DataFrame
        Analysis DataFrame with all scenario columns
    ppa_range_start : int
        Starting PPA coverage percentage
    ppa_range_end : int
        Ending PPA coverage percentage (inclusive)
    ppa_range_step : int
        Step size for PPA coverage

    Returns
    -------
    pd.DataFrame
        Long-format DataFrame with one row per hour per scenario
    """
    long_data = []

    for ppa_percent in range(ppa_range_start, ppa_range_end + 1, ppa_range_step):
        for hour in range(len(analysis_df)):
            # Base record
            base_record = {
                'datetime': analysis_df['datetime'].iloc[hour],
                'year': analysis_df['datetime'].iloc[hour].year,
                'month': analysis_df['datetime'].iloc[hour].month,
                'day': analysis_df['datetime'].iloc[hour].day,
                'hour': analysis_df['hour'].iloc[hour],
                'day_of_year': analysis_df['day_of_year'].iloc[hour],
                'ppa_scenario': f'PPA{ppa_percent}',
                'ppa_coverage_pct': ppa_percent,
                'ppa_coverage_factor': ppa_percent / 100,
            }

            # Energy values
            energy_values = {
                'Load_Demand_kWh': analysis_df['load_mw'].iloc[hour] * 1000,
                'Solar_Base_Generation_kWh': analysis_df['solar_generation_mw'].iloc[hour] * 1000,
                'PPA_Generation_kWh': analysis_df[f'ppa_gen_{ppa_percent}pct'].iloc[hour],
                'PPA_Mandatory_kWh': analysis_df[f'mandatory_ppa_{ppa_percent}pct'].iloc[hour],
                'PPA_Optional_kWh': analysis_df[f'optional_ppa_{ppa_percent}pct'].iloc[hour],
                'PPA_Purchase_kWh': analysis_df[f'ppa_purchase_{ppa_percent}pct'].iloc[hour],
                'PPA_Excess_kWh': analysis_df[f'ppa_excess_{ppa_percent}pct'].iloc[hour],
                'Grid_Purchase_kWh': analysis_df[f'grid_purchase_{ppa_percent}pct'].iloc[hour],
            }

            # Cost values
            cost_values = {
                'Grid_Rate_KRW_per_kWh': analysis_df['grid_rate'].iloc[hour],
                'PPA_Cost_KRW': analysis_df[f'ppa_cost_{ppa_percent}pct'].iloc[hour],
                'Grid_Cost_KRW': analysis_df[f'grid_cost_{ppa_percent}pct'].iloc[hour],
                'Resell_Revenue_KRW': analysis_df[f'resell_revenue_{ppa_percent}pct'].iloc[hour],
                'Total_Cost_KRW': analysis_df[f'total_cost_{ppa_percent}pct'].iloc[hour],
            }

            # Combine all data
            record = {**base_record, **energy_values, **cost_values}
            long_data.append(record)

    # Create long DataFrame
    long_df = pd.DataFrame(long_data)

    # Add calculated fields
    long_df['Load_Coverage_by_PPA_pct'] = (long_df['PPA_Purchase_kWh'] / long_df['Load_Demand_kWh'] * 100).fillna(0)
    long_df['Excess_to_Generation_pct'] = (long_df['PPA_Excess_kWh'] / long_df['PPA_Generation_kWh'] * 100).fillna(0)
    long_df['Cost_Savings_vs_Grid_Only'] = (long_df['Load_Demand_kWh'] * long_df['Grid_Rate_KRW_per_kWh']) - long_df['Total_Cost_KRW']
    long_df['Cost_per_kWh_KRW'] = (long_df['Total_Cost_KRW'] / long_df['Load_Demand_kWh']).fillna(0)

    return long_df
