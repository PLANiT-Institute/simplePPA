import pandas as pd
import numpy as np
import libs.KEPCOutils as _kepco

# Import pattern file
pattern_df = pd.read_excel('data/pattern.xlsx', index_col=0)

# Save load and solar to separate DataFrames
load_df = pattern_df[['load']]
solar_df = pattern_df[['solar']]

# Process Grid data
grid_df, contract_fee = _kepco.process_kepco_data("data/KEPCO.xlsx", 2024, "HV_C_III")

# Analysis timeframe
start_date = '2024-01-01'  # Start date for analysis (YYYY-MM-DD)
end_date = '2024-12-31'    # End date for analysis (YYYY-MM-DD)

# Load capacity setting
load_capacity_mw = 3000  # MW - Maximum load capacity (e.g., 100 = 100MW peak load)

ppa_price = 170        # KRW/kWh
ppa_mintake = 1        # %, minimum required purchase (1=100%, 0.5=50%). If <100%, company can choose to buy more if profitable
ppa_resell = False      # buyer is allowed to resell (True), or not (False)
ppa_resellrate = 0.9   # buyer can resell the remained amount at the purchased price * resellrate

ess_include = False     # Include ESS in analysis (True/False)
ess_capacity = 0.5     # ESS capacity as percentage of solar peak generation (0.5 = 50% of peak solar)
ess_price = 0.5        # ESS to the ppa_price (ess_price * ppa_price) is the price of energy discharged from ess
ess_hours = 6          # ESS capacity (hours to shift) - only used if ess_capacity is based on hours

# Create ppa_df from solar_df (assuming PPA generation follows solar pattern)
ppa_df = solar_df.copy()
ppa_df.columns = ['generation']

def calculate_ppa_cost(load_df, ppa_df, grid_df, load_capacity_mw, ppa_coverage, contract_fee, ess_capacity=0):
    """
    Calculate total power cost for given load capacity and PPA coverage
    
    Parameters:
    - load_df: DataFrame with hourly load demand (normalized 0-1)
    - ppa_df: DataFrame with hourly solar generation pattern (normalized 0-1)
    - grid_df: DataFrame with hourly Grid rates (KRW/kWh)
    - load_capacity_mw: Maximum load capacity in MW (e.g., 100 = 100MW peak load)
    - ppa_coverage: PPA sizing as percentage of peak load (0-2.0, where 1.0 = 100% of peak load)
    - contract_fee: Grid contract fee (KRW/kW) applied to peak Grid demand
    - ess_capacity: ESS capacity in kWh (0 = no ESS)
    
    Returns:
    - total_cost: Total annual cost in KRW
    - ppa_cost: PPA cost component
    - grid_cost: Grid energy cost component
    - grid_demand_cost: Grid demand charge component
    - ess_cost: ESS cost component
    """
    
    load = load_df['load']
    grid_rate = grid_df['rate']
    solar_generation = ppa_df['generation']
    
    # Scale load by the specified capacity
    # If load_capacity_mw = 100, then when normalized load = 1.0, actual load = 100MW
    load_mw = load * load_capacity_mw
    
    # Scale PPA generation based on coverage percentage of the load capacity
    # ppa_coverage = 1.0 means PPA solar farm peak capacity equals peak load capacity
    ppa_generation_mw = solar_generation * load_capacity_mw * ppa_coverage
    
    # Convert MW to kWh for cost calculations (MW * 1 hour = MWh = 1000 kWh)
    load_kwh = load_mw * 1000
    ppa_generation_kwh = ppa_generation_mw * 1000
    
    # Initialize costs
    ppa_cost = 0
    grid_energy_cost = 0
    ess_cost = 0
    
    # Initialize ESS state
    ess_storage = 0
    max_ess_capacity = ess_capacity
    
    # Track peak Grid demand for contract fee calculation
    peak_grid_demand_kw = 0
    
    results = []
    total_ppa_purchased = 0
    total_ppa_excess = 0
    total_ppa_resold = 0
    total_load_met_by_ppa = 0
    
    for hour in range(len(load_kwh)):
        hour_load = load_kwh.iloc[hour]
        hour_grid_rate = grid_rate.iloc[hour]
        hour_ppa_gen = ppa_generation_kwh.iloc[hour]
        
        # Mandatory PPA purchase (minimum take) - MUST buy this amount regardless of need
        mandatory_ppa = hour_ppa_gen * ppa_mintake
        
        # Optional PPA purchase (if mintake < 100%) - buy more if profitable
        optional_ppa_available = hour_ppa_gen - mandatory_ppa
        optional_ppa_purchased = 0
        
        if optional_ppa_available > 0:
            # Calculate remaining load after mandatory purchase
            remaining_load_after_mandatory = max(0, hour_load - mandatory_ppa)
            
            # Only buy optional PPA if it's cheaper than Grid and we need the energy
            if ppa_price < hour_grid_rate and remaining_load_after_mandatory > 0:
                optional_ppa_purchased = min(optional_ppa_available, remaining_load_after_mandatory)
        
        # Total PPA purchased this hour
        total_ppa_this_hour = mandatory_ppa + optional_ppa_purchased
        ppa_cost += total_ppa_this_hour * ppa_price
        total_ppa_purchased += total_ppa_this_hour
        
        # Track how much load is actually met by PPA
        ppa_used_for_load = min(total_ppa_this_hour, hour_load)
        total_load_met_by_ppa += ppa_used_for_load
        
        # After all PPA purchases, determine if we have excess or deficit
        remaining_load = hour_load - total_ppa_this_hour
        
        if remaining_load <= 0:
            # We have excess PPA energy - must handle it
            excess_ppa = abs(remaining_load)
            total_ppa_excess += excess_ppa
            
            # Try to store excess in ESS
            if ess_storage < max_ess_capacity:
                ess_charge = min(excess_ppa, max_ess_capacity - ess_storage)
                ess_storage += ess_charge
                excess_ppa -= ess_charge
                # ESS charging doesn't add cost - already paid for PPA energy
            
            # Try to resell remaining excess
            if excess_ppa > 0 and ppa_resell:
                resell_revenue = excess_ppa * ppa_price * ppa_resellrate
                ppa_cost -= resell_revenue  # Reduce cost through resell revenue
                total_ppa_resold += excess_ppa
                excess_ppa = 0
            
            # Any remaining excess is pure cost (already paid for in mandatory purchase)
            remaining_load = 0
            
        else:
            # We need more energy beyond PPA
            # First try to use ESS
            if ess_storage > 0:
                ess_discharge = min(remaining_load, ess_storage)
                ess_storage -= ess_discharge
                ess_cost += ess_discharge * ppa_price * ess_price
                remaining_load -= ess_discharge
            
            # Buy any remaining load from Grid
            if remaining_load > 0:
                grid_energy_cost += remaining_load * hour_grid_rate
                # Track peak Grid demand for contract fee
                # remaining_load is in kWh for 1 hour, which equals kW of demand
                grid_demand_kw = remaining_load
                peak_grid_demand_kw = max(peak_grid_demand_kw, grid_demand_kw)
        
        results.append({
            'hour': hour,
            'load': hour_load,
            'solar_gen': solar_generation.iloc[hour],
            'ppa_gen': hour_ppa_gen,
            'mandatory_ppa': mandatory_ppa,
            'optional_ppa': optional_ppa_purchased,
            'total_ppa_purchased': total_ppa_this_hour,
            'ess_storage': ess_storage,
            'remaining_load': max(0, remaining_load)
        })
    
    # Calculate Grid demand charge
    grid_demand_cost = peak_grid_demand_kw * contract_fee
    
    # Calculate total costs
    grid_total_cost = grid_energy_cost + grid_demand_cost
    total_cost = ppa_cost + grid_total_cost + ess_cost
    
    # Calculate actual coverage achieved
    actual_coverage = total_load_met_by_ppa / load_kwh.sum() if load_kwh.sum() > 0 else 0
    annual_ppa_gen = ppa_generation_kwh.sum()
    
    # Print some statistics for debugging
    if ppa_coverage > 0:
        ppa_capacity = load_capacity_mw * ppa_coverage
        print(f"  PPA Stats: Load={load_capacity_mw:.1f}MW, PPA={ppa_capacity:.1f}MW ({ppa_coverage*100:.0f}%), Gen={annual_ppa_gen:,.0f} kWh, Load coverage={actual_coverage*100:.1f}%, Purchased={total_ppa_purchased:,.0f} kWh, Excess={total_ppa_excess:,.0f} kWh, Resold={total_ppa_resold:,.0f} kWh")
        print(f"  Grid Stats: Peak demand={peak_grid_demand_kw:,.0f} kW, Energy cost={grid_energy_cost:,.0f} KRW, Demand cost={grid_demand_cost:,.0f} KRW")
    
    return total_cost, ppa_cost, grid_total_cost, grid_demand_cost, ess_cost

# Analysis loop for different PPA coverage percentages (0-200% of peak load)
print("=== PPA Coverage Analysis (No ESS) ===")
print("PPA% | Total Cost/kWh | PPA Cost/kWh | Grid Energy/kWh | Contract/kWh | ESS Cost/kWh")
print("-" * 85)

results_summary = []

for ppa_percent in range(0, 201, 10):
    ppa_coverage = ppa_percent / 100
    
    # Calculate costs without ESS
    total_cost, ppa_cost, grid_cost, grid_demand_cost, ess_cost = calculate_ppa_cost(
        load_df, ppa_df, grid_df, load_capacity_mw, ppa_coverage, contract_fee, ess_capacity=0
    )
    
    # Calculate total electricity supplied (PPA + Grid)
    total_electricity_kwh = load_capacity_mw * 1000 * load_df['load'].sum()
    
    results_summary.append({
        'ppa_percent': ppa_percent,
        'total_cost': total_cost,
        'ppa_cost': ppa_cost,
        'grid_cost': grid_cost,
        'grid_demand_cost': grid_demand_cost,
        'ess_cost': ess_cost,
        'total_electricity_kwh': total_electricity_kwh
    })
    
    # Calculate per-kWh costs
    total_cost_per_kwh = total_cost / total_electricity_kwh
    ppa_cost_per_kwh = ppa_cost / total_electricity_kwh
    grid_energy_cost = grid_cost - grid_demand_cost
    grid_energy_cost_per_kwh = grid_energy_cost / total_electricity_kwh
    grid_demand_cost_per_kwh = grid_demand_cost / total_electricity_kwh
    ess_cost_per_kwh = ess_cost / total_electricity_kwh
    
    print(f"{ppa_percent:3d}% | {total_cost_per_kwh:12.2f} | {ppa_cost_per_kwh:11.2f} | {grid_energy_cost_per_kwh:14.2f} | {grid_demand_cost_per_kwh:14.2f} | {ess_cost_per_kwh:11.2f}")

# Find optimal PPA coverage without ESS
optimal_idx = min(range(len(results_summary)), key=lambda i: results_summary[i]['total_cost'])
optimal_ppa = results_summary[optimal_idx]['ppa_percent']
optimal_cost = results_summary[optimal_idx]['total_cost']

print(f"\nOptimal PPA coverage (No ESS): {optimal_ppa}% with total cost: {optimal_cost:,.0f} KRW")

# ESS Analysis (only if enabled)
if ess_include:
    print(f"\n=== PPA Coverage Analysis (With ESS: {ess_capacity*100:.0f}% solar capacity) ===")
    print("PPA% | Total Cost/kWh | PPA Cost/kWh | Grid Energy/kWh | Contract/kWh | ESS Cost/kWh")
    print("-" * 85)

    results_ess = []
    # Calculate ESS capacity based on solar peak generation (scaled to MW)
    peak_solar_mw = ppa_df['generation'].max() * load_capacity_mw
    ess_capacity_kwh = peak_solar_mw * ess_capacity * 1000  # Convert MW to kWh

    for ppa_percent in range(0, 201, 10):
        ppa_coverage = ppa_percent / 100
        
        # Calculate costs with ESS
        total_cost, ppa_cost, grid_cost, grid_demand_cost, ess_cost = calculate_ppa_cost(
            load_df, ppa_df, grid_df, load_capacity_mw, ppa_coverage, contract_fee, ess_capacity=ess_capacity_kwh
        )
        
        results_ess.append({
            'ppa_percent': ppa_percent,
            'total_cost': total_cost,
            'ppa_cost': ppa_cost,
            'grid_cost': grid_cost,
            'grid_demand_cost': grid_demand_cost,
            'ess_cost': ess_cost
        })
        
        # Calculate per-kWh costs
        total_cost_per_kwh_ess = total_cost / total_electricity_kwh
        ppa_cost_per_kwh_ess = ppa_cost / total_electricity_kwh
        grid_energy_cost_ess = grid_cost - grid_demand_cost
        grid_energy_cost_per_kwh_ess = grid_energy_cost_ess / total_electricity_kwh
        grid_demand_cost_per_kwh_ess = grid_demand_cost / total_electricity_kwh
        ess_cost_per_kwh_ess = ess_cost / total_electricity_kwh
        
        print(f"{ppa_percent:3d}% | {total_cost_per_kwh_ess:12.2f} | {ppa_cost_per_kwh_ess:11.2f} | {grid_energy_cost_per_kwh_ess:14.2f} | {grid_demand_cost_per_kwh_ess:14.2f} | {ess_cost_per_kwh_ess:11.2f}")

    # Find optimal PPA coverage with ESS
    optimal_ess_idx = min(range(len(results_ess)), key=lambda i: results_ess[i]['total_cost'])
    optimal_ess_ppa = results_ess[optimal_ess_idx]['ppa_percent']
    optimal_ess_cost = results_ess[optimal_ess_idx]['total_cost']

    print(f"\nOptimal PPA coverage (With ESS): {optimal_ess_ppa}% with total cost: {optimal_ess_cost:,.0f} KRW")
    print(f"ESS Capacity: {ess_capacity_kwh:.1f} kWh ({ess_capacity*100:.0f}% of peak solar {peak_solar_mw:.1f} MW)")
else:
    results_ess = []
    optimal_ess_ppa = "N/A"
    optimal_ess_cost = "N/A"

# Summary comparison
print(f"\n=== SUMMARY ===")
print(f"ESS included: {ess_include}")
print(f"Resell enabled: {ppa_resell}, Resell rate: {ppa_resellrate}")
print(f"No ESS  - Optimal: {optimal_ppa}% PPA, Cost: {optimal_cost:,.0f} KRW")
if ess_include:
    print(f"With ESS - Optimal: {optimal_ess_ppa}% PPA, Cost: {optimal_ess_cost:,.0f} KRW")
    print(f"ESS benefit: {optimal_cost - optimal_ess_cost:,.0f} KRW savings")
else:
    print(f"With ESS - Analysis skipped (ESS disabled)")

# Save generation patterns for analysis
print(f"\n=== Saving Comprehensive Generation Patterns ===")
print(f"Analysis period: {start_date} to {end_date}")

# Create base analysis DataFrame
analysis_df = pd.DataFrame({
    'datetime': grid_df.index,
    'load': load_df['load'].values,
    'grid_rate': grid_df['rate'].values,
    'solar_generation': ppa_df['generation'].values
})

# Filter data to user-defined timeframe
mask = (analysis_df['datetime'] >= start_date) & (analysis_df['datetime'] <= end_date)
analysis_df = analysis_df[mask].copy()

print(f"Data filtered to {len(analysis_df):,} hours ({len(analysis_df)/24:.1f} days)")

# Add time components for analysis
analysis_df['hour'] = analysis_df['datetime'].dt.hour
analysis_df['month'] = analysis_df['datetime'].dt.month
analysis_df['day_of_year'] = analysis_df['datetime'].dt.dayofyear

# Scale load and solar data to MW using the capacity parameter
analysis_df['load_mw'] = analysis_df['load'] * load_capacity_mw
analysis_df['solar_generation_mw'] = analysis_df['solar_generation'] * load_capacity_mw

# Add PPA generation, purchases, costs, and excess for ALL coverage levels (0-200%)
print("Calculating hourly patterns for all PPA scenarios...")

# Build all scenario columns efficiently
scenario_columns = {}
for ppa_percent in range(0, 201, 10):
    ppa_coverage = ppa_percent / 100
    
    # Calculate scaled PPA generation for this scenario (in MW)
    ppa_generation_mw = analysis_df['solar_generation'] * load_capacity_mw * ppa_coverage
    
    # Convert to kWh for cost calculations
    load_kwh = analysis_df['load_mw'] * 1000
    ppa_generation_kwh = ppa_generation_mw * 1000
    
    # Calculate mandatory PPA purchase 
    mandatory_ppa = ppa_generation_kwh * ppa_mintake
    
    # Calculate optional PPA purchase (if mintake < 100%)
    optional_ppa_available = ppa_generation_kwh - mandatory_ppa
    remaining_load_after_mandatory = (load_kwh - mandatory_ppa).clip(lower=0)
    
    # Only buy optional PPA if it's cheaper than Grid and we need the energy
    optional_ppa = pd.Series(0.0, index=analysis_df.index, dtype=float)
    if ppa_mintake < 1:  # Only if there's optional capacity
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
    
    # Calculate excess PPA (when PPA > load)
    ppa_excess = (-remaining_load).clip(lower=0)
    
    # Calculate Grid purchases (when load > PPA)
    grid_purchase = remaining_load.clip(lower=0)
    
    # Calculate Grid costs
    grid_cost = grid_purchase * analysis_df['grid_rate']
    
    # Calculate resell revenue if enabled
    if ppa_resell:
        resell_revenue = ppa_excess * ppa_price * ppa_resellrate
        # Adjust PPA cost for resell revenue
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

# Add all scenario columns at once to avoid fragmentation
scenario_df = pd.DataFrame(scenario_columns, index=analysis_df.index)
analysis_df = pd.concat([analysis_df, scenario_df], axis=1)

print(f"Generated hourly patterns for {len(range(0, 201, 10))} PPA scenarios (0-200% in 10% steps)")

# Create long-format DataFrame for pivot table analysis
print("Creating long-format DataFrame for pivot analysis...")
long_data = []

for ppa_percent in range(0, 201, 10):
    for hour in range(len(analysis_df)):
        # Base record for each hour-scenario combination
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
        
        # Energy values (kWh)
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
        
        # Cost/Rate values (KRW)
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

# Add calculated fields for analysis
long_df['Load_Coverage_by_PPA_pct'] = (long_df['PPA_Purchase_kWh'] / long_df['Load_Demand_kWh'] * 100).fillna(0)
long_df['Excess_to_Generation_pct'] = (long_df['PPA_Excess_kWh'] / long_df['PPA_Generation_kWh'] * 100).fillna(0)
long_df['Cost_Savings_vs_Grid_Only'] = (long_df['Load_Demand_kWh'] * long_df['Grid_Rate_KRW_per_kWh']) - long_df['Total_Cost_KRW']
long_df['Cost_per_kWh_KRW'] = (long_df['Total_Cost_KRW'] / long_df['Load_Demand_kWh']).fillna(0)

# Create annual summary for each PPA scenario  
annual_summary = []
total_annual_demand_kwh = (analysis_df['load_mw'].sum() * 1000)

for ppa_percent in range(0, 201, 10):
    annual_ppa_cost = analysis_df[f'ppa_cost_{ppa_percent}pct'].sum()
    annual_grid_cost = analysis_df[f'grid_cost_{ppa_percent}pct'].sum()
    annual_total_cost = analysis_df[f'total_cost_{ppa_percent}pct'].sum()
    
    annual_summary.append({
        'PPA_Coverage': f"{ppa_percent}%",
        'Annual_PPA_Gen': analysis_df[f'ppa_gen_{ppa_percent}pct'].sum(),
        'Annual_Mandatory_PPA': analysis_df[f'mandatory_ppa_{ppa_percent}pct'].sum(),
        'Annual_Optional_PPA': analysis_df[f'optional_ppa_{ppa_percent}pct'].sum(),
        'Annual_PPA_Purchase': analysis_df[f'ppa_purchase_{ppa_percent}pct'].sum(),
        'Annual_PPA_Cost': annual_ppa_cost,
        'Annual_Grid_Purchase': analysis_df[f'grid_purchase_{ppa_percent}pct'].sum(),
        'Annual_Grid_Cost': annual_grid_cost,
        'Annual_PPA_Excess': analysis_df[f'ppa_excess_{ppa_percent}pct'].sum(),
        'Annual_Resell_Revenue': analysis_df[f'resell_revenue_{ppa_percent}pct'].sum(),
        'Annual_Total_Cost': annual_total_cost,
        'Load_Coverage_Pct': (analysis_df[f'ppa_purchase_{ppa_percent}pct'].sum() / total_annual_demand_kwh * 100),
        # Per-kWh cost columns
        'PPA_Cost_per_kWh': annual_ppa_cost / total_annual_demand_kwh,
        'Grid_Cost_per_kWh': annual_grid_cost / total_annual_demand_kwh,
        'Total_Cost_per_kWh': annual_total_cost / total_annual_demand_kwh
    })

annual_summary_df = pd.DataFrame(annual_summary)

# Cost comparison results
cost_comparison = pd.DataFrame({
    'PPA_Coverage': [f"{i}%" for i in range(0, 201, 10)],
    'Total_Cost': [result['total_cost'] for result in results_summary],
    'PPA_Cost': [result['ppa_cost'] for result in results_summary],
    'Grid_Energy_Cost': [result['grid_cost'] for result in results_summary],
    'Contract_Cost': [result['grid_demand_cost'] for result in results_summary],
    'ESS_Cost': [result['ess_cost'] for result in results_summary],
    # Per-kWh cost columns
    'Total_Cost_per_kWh': [result['total_cost'] / result['total_electricity_kwh'] for result in results_summary],
    'PPA_Cost_per_kWh': [result['ppa_cost'] / result['total_electricity_kwh'] for result in results_summary],
    'Grid_Energy_Cost_per_kWh': [(result['grid_cost'] - result['grid_demand_cost']) / result['total_electricity_kwh'] for result in results_summary],
    'Contract_Cost_per_kWh': [result['grid_demand_cost'] / result['total_electricity_kwh'] for result in results_summary],
    'ESS_Cost_per_kWh': [result['ess_cost'] / result['total_electricity_kwh'] for result in results_summary]
})

# Save to Excel with only essential sheets
with pd.ExcelWriter('ppa_analysis_results.xlsx', engine='openpyxl') as writer:
    # Primary output: Long-format data for pivot analysis
    long_df.to_excel(writer, sheet_name='PPA_Analysis_Data', index=False)
    
    # Annual summary
    annual_summary_df.to_excel(writer, sheet_name='Annual_Summary', index=False)
    
    # Cost analysis
    cost_comparison.to_excel(writer, sheet_name='Cost_Analysis', index=False)

print(f"Analysis saved to: ppa_analysis_results.xlsx")
print(f"Sheets: PPA_Analysis_Data (long-format for pivot analysis), Annual_Summary, Cost_Analysis")
print(f"Long-format data contains {len(long_df):,} rows covering all {len(range(0, 201, 10))} PPA scenarios")

# Print some key insights
print(f"\n=== Key Pattern Insights ===")
print(f"Analysis period: {start_date} to {end_date} ({len(analysis_df)/24:.1f} days)")
print(f"Load capacity: {load_capacity_mw} MW")
print(f"Total load: {(analysis_df['load_mw'].sum() * 1000):,.0f} kWh")
print(f"Average Grid rate: {analysis_df['grid_rate'].mean():.1f} KRW/kWh")
print(f"Grid rate range: {analysis_df['grid_rate'].min():.1f} - {analysis_df['grid_rate'].max():.1f} KRW/kWh")
print(f"PPA price: {ppa_price} KRW/kWh")
print(f"Average savings per kWh: {analysis_df['grid_rate'].mean() - ppa_price:.1f} KRW/kWh")

# Peak vs off-peak analysis based on actual Grid rates
hourly_avg_rates = analysis_df.groupby('hour')['grid_rate'].mean().sort_values(ascending=False)
peak_hours = hourly_avg_rates.head(12).index.tolist()  # Top 12 most expensive hours
offpeak_hours = hourly_avg_rates.tail(12).index.tolist()  # Bottom 12 cheapest hours

peak_data = analysis_df[analysis_df['hour'].isin(peak_hours)]
offpeak_data = analysis_df[analysis_df['hour'].isin(offpeak_hours)]

print(f"\nPeak hours {sorted(peak_hours)}: Avg Grid rate = {peak_data['grid_rate'].mean():.1f} KRW/kWh")
print(f"Off-peak hours {sorted(offpeak_hours)}: Avg Grid rate = {offpeak_data['grid_rate'].mean():.1f} KRW/kWh")
print(f"Peak hour savings with PPA: {peak_data['grid_rate'].mean() - ppa_price:.1f} KRW/kWh")
