import pandas as pd
import libs.KEPCOutils as _kepco

# Import pattern file
pattern_df = pd.read_excel('data/pattern.xlsx', index_col=0)

# Save load and solar to separate DataFrames
load_df = pattern_df[['load']]
solar_df = pattern_df[['solar']]

# Process KEPCO data
kepco_df, contract_fee = _kepco.process_kepco_data("data/KEPCO.xlsx", 2024, "HV_C_I")

ppa_price = 150        # KRW/kWh
ppa_mintake = 1        # %, buyer must purchase the min take * generation of each hour, whether a buyer use it or not
ppa_resell = False      # buyer is allowed to resell (True), or not (False)
ppa_resellrate = 0.9   # buyer can resell the remained amount at the purchased price * resellrate

ess_include = False     # Include ESS in analysis (True/False)
ess_capacity = 0.5     # ESS capacity as percentage of solar peak generation (0.5 = 50% of peak solar)
ess_price = 0.5        # ESS to the ppa_price (ess_price * ppa_price) is the price of energy discharged from ess
ess_hours = 4          # ESS capacity (hours to shift) - only used if ess_capacity is based on hours

# Create ppa_df from solar_df (assuming PPA generation follows solar pattern)
ppa_df = solar_df.copy()
ppa_df.columns = ['generation']

def calculate_ppa_cost(load_df, ppa_df, kepco_df, ppa_coverage, ess_capacity=0):
    """
    Calculate total power cost for given PPA coverage percentage
    
    Parameters:
    - load_df: DataFrame with hourly load demand
    - ppa_df: DataFrame with hourly solar generation pattern
    - kepco_df: DataFrame with hourly KEPCO rates
    - ppa_coverage: PPA sizing as percentage of peak load (0-2.0, where 1.0 = 100% of peak load)
    - ess_capacity: ESS capacity in kWh (0 = no ESS)
    
    Returns:
    - total_cost: Total annual cost in KRW
    - ppa_cost: PPA cost component
    - kepco_cost: KEPCO cost component
    - ess_cost: ESS cost component
    """
    
    load = load_df['load']
    kepco_rate = kepco_df['rate']
    solar_generation = ppa_df['generation']
    
    # Calculate PPA sizing based on peak load
    # ppa_coverage = 1.0 means PPA solar farm peak capacity equals peak load capacity
    # ppa_coverage = 0.1 means PPA solar farm sized at 10% of peak load capacity
    peak_load = load.max()
    peak_solar = solar_generation.max()
    
    # Scale solar generation pattern so that at 100% coverage, solar peak = load peak
    # This accounts for the fact that raw solar data peak may not equal 1.0
    if peak_solar > 0:
        # At 100% coverage: scaled_solar_peak = peak_load
        # So scaling factor = (peak_load * ppa_coverage) / peak_solar
        ppa_scale_factor = (peak_load * ppa_coverage) / peak_solar
        ppa_generation = solar_generation * ppa_scale_factor
    else:
        ppa_generation = solar_generation * 0
    
    # Initialize costs
    ppa_cost = 0
    kepco_cost = 0
    ess_cost = 0
    
    # Initialize ESS state
    ess_storage = 0
    max_ess_capacity = ess_capacity
    
    results = []
    total_ppa_purchased = 0
    total_ppa_excess = 0
    total_ppa_resold = 0
    total_load_met_by_ppa = 0
    
    for hour in range(len(load)):
        hour_load = load.iloc[hour]
        hour_kepco_rate = kepco_rate.iloc[hour]
        hour_ppa_gen = ppa_generation.iloc[hour]
        
        # Mandatory PPA purchase (minimum take) - MUST buy this amount regardless of need
        mandatory_ppa = hour_ppa_gen * ppa_mintake
        ppa_cost += mandatory_ppa * ppa_price
        total_ppa_purchased += mandatory_ppa
        
        # Track how much load is actually met by PPA
        ppa_used_for_load = min(mandatory_ppa, hour_load)
        total_load_met_by_ppa += ppa_used_for_load
        
        # After mandatory purchase, determine if we have excess or deficit
        remaining_load = hour_load - mandatory_ppa
        
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
            
            # Buy any remaining load from KEPCO
            if remaining_load > 0:
                kepco_cost += remaining_load * hour_kepco_rate
        
        results.append({
            'hour': hour,
            'load': hour_load,
            'solar_gen': solar_generation.iloc[hour],
            'ppa_gen': hour_ppa_gen,
            'mandatory_ppa': mandatory_ppa,
            'ess_storage': ess_storage,
            'remaining_load': max(0, remaining_load)
        })
    
    total_cost = ppa_cost + kepco_cost + ess_cost
    
    # Calculate actual coverage achieved
    actual_coverage = total_load_met_by_ppa / load.sum() if load.sum() > 0 else 0
    annual_ppa_gen = ppa_generation.sum()
    
    # Print some statistics for debugging
    if ppa_coverage > 0:
        ppa_peak_capacity = peak_load * ppa_coverage
        print(f"  PPA Stats: Size={ppa_coverage*100:.0f}% peak load ({ppa_peak_capacity:.1f}kW), Gen={annual_ppa_gen:,.0f} kWh, Load coverage={actual_coverage*100:.1f}%, Purchased={total_ppa_purchased:,.0f} kWh, Excess={total_ppa_excess:,.0f} kWh, Resold={total_ppa_resold:,.0f} kWh")
    
    return total_cost, ppa_cost, kepco_cost, ess_cost

# Analysis loop for different PPA coverage percentages (0-200% of peak load)
print("=== PPA Coverage Analysis (No ESS) ===")
print("PPA% | Total Cost | PPA Cost | KEPCO Cost | ESS Cost")
print("-" * 55)

results_summary = []

for ppa_percent in range(0, 201, 10):
    ppa_coverage = ppa_percent / 100
    
    # Calculate costs without ESS
    total_cost, ppa_cost, kepco_cost, ess_cost = calculate_ppa_cost(
        load_df, ppa_df, kepco_df, ppa_coverage, ess_capacity=0
    )
    
    results_summary.append({
        'ppa_percent': ppa_percent,
        'total_cost': total_cost,
        'ppa_cost': ppa_cost,
        'kepco_cost': kepco_cost,
        'ess_cost': ess_cost
    })
    
    print(f"{ppa_percent:3d}% | {total_cost:10,.0f} | {ppa_cost:8,.0f} | {kepco_cost:10,.0f} | {ess_cost:8,.0f}")

# Find optimal PPA coverage without ESS
optimal_idx = min(range(len(results_summary)), key=lambda i: results_summary[i]['total_cost'])
optimal_ppa = results_summary[optimal_idx]['ppa_percent']
optimal_cost = results_summary[optimal_idx]['total_cost']

print(f"\nOptimal PPA coverage (No ESS): {optimal_ppa}% with total cost: {optimal_cost:,.0f} KRW")

# ESS Analysis (only if enabled)
if ess_include:
    print(f"\n=== PPA Coverage Analysis (With ESS: {ess_capacity*100:.0f}% solar capacity) ===")
    print("PPA% | Total Cost | PPA Cost | KEPCO Cost | ESS Cost")
    print("-" * 55)

    results_ess = []
    # Calculate ESS capacity based on solar peak generation
    peak_solar = ppa_df['generation'].max()
    ess_capacity_kwh = peak_solar * ess_capacity

    for ppa_percent in range(0, 201, 10):
        ppa_coverage = ppa_percent / 100
        
        # Calculate costs with ESS
        total_cost, ppa_cost, kepco_cost, ess_cost = calculate_ppa_cost(
            load_df, ppa_df, kepco_df, ppa_coverage, ess_capacity=ess_capacity_kwh
        )
        
        results_ess.append({
            'ppa_percent': ppa_percent,
            'total_cost': total_cost,
            'ppa_cost': ppa_cost,
            'kepco_cost': kepco_cost,
            'ess_cost': ess_cost
        })
        
        print(f"{ppa_percent:3d}% | {total_cost:10,.0f} | {ppa_cost:8,.0f} | {kepco_cost:10,.0f} | {ess_cost:8,.0f}")

    # Find optimal PPA coverage with ESS
    optimal_ess_idx = min(range(len(results_ess)), key=lambda i: results_ess[i]['total_cost'])
    optimal_ess_ppa = results_ess[optimal_ess_idx]['ppa_percent']
    optimal_ess_cost = results_ess[optimal_ess_idx]['total_cost']

    print(f"\nOptimal PPA coverage (With ESS): {optimal_ess_ppa}% with total cost: {optimal_ess_cost:,.0f} KRW")
    print(f"ESS Capacity: {ess_capacity_kwh:.1f} kWh ({ess_capacity*100:.0f}% of peak solar {peak_solar:.1f} kW)")
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

# Create base analysis DataFrame
analysis_df = pd.DataFrame({
    'datetime': kepco_df.index,
    'load': load_df['load'].values,
    'kepco_rate': kepco_df['rate'].values,
    'solar_generation': ppa_df['generation'].values
})

# Add time components for analysis
analysis_df['hour'] = analysis_df['datetime'].dt.hour
analysis_df['month'] = analysis_df['datetime'].dt.month
analysis_df['day_of_year'] = analysis_df['datetime'].dt.dayofyear

# Calculate scaling factors for all PPA scenarios
peak_load = load_df['load'].max()
peak_solar = ppa_df['generation'].max()

# Add PPA generation, purchases, costs, and excess for ALL coverage levels (0-200%)
print("Calculating hourly patterns for all PPA scenarios...")

# Build all scenario columns efficiently
scenario_columns = {}
for ppa_percent in range(0, 201, 10):
    ppa_coverage = ppa_percent / 100
    
    # Calculate scaled PPA generation for this scenario
    if peak_solar > 0:
        ppa_scale_factor = (peak_load * ppa_coverage) / peak_solar
        ppa_generation = analysis_df['solar_generation'] * ppa_scale_factor
    else:
        ppa_generation = analysis_df['solar_generation'] * 0
    
    # Calculate mandatory PPA purchase 
    ppa_purchase = ppa_generation * ppa_mintake
    
    # Calculate PPA costs
    ppa_cost = ppa_purchase * ppa_price
    
    # Calculate remaining load after PPA
    remaining_load = analysis_df['load'] - ppa_purchase
    
    # Calculate excess PPA (when PPA > load)
    ppa_excess = (-remaining_load).clip(lower=0)
    
    # Calculate KEPCO purchases (when load > PPA)
    kepco_purchase = remaining_load.clip(lower=0)
    
    # Calculate KEPCO costs
    kepco_cost = kepco_purchase * analysis_df['kepco_rate']
    
    # Calculate resell revenue if enabled
    if ppa_resell:
        resell_revenue = ppa_excess * ppa_price * ppa_resellrate
        # Adjust PPA cost for resell revenue
        ppa_cost -= resell_revenue
    else:
        resell_revenue = pd.Series(0, index=analysis_df.index)
    
    # Calculate total hourly cost
    total_cost = ppa_cost + kepco_cost
    
    # Store all columns for this scenario
    scenario_columns[f'ppa_gen_{ppa_percent}pct'] = ppa_generation
    scenario_columns[f'ppa_purchase_{ppa_percent}pct'] = ppa_purchase  
    scenario_columns[f'ppa_cost_{ppa_percent}pct'] = ppa_cost
    scenario_columns[f'ppa_excess_{ppa_percent}pct'] = ppa_excess
    scenario_columns[f'kepco_purchase_{ppa_percent}pct'] = kepco_purchase
    scenario_columns[f'kepco_cost_{ppa_percent}pct'] = kepco_cost
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
            'Load_Demand_kWh': analysis_df['load'].iloc[hour],
            'Solar_Base_Generation_kWh': analysis_df['solar_generation'].iloc[hour],
            'PPA_Generation_kWh': analysis_df[f'ppa_gen_{ppa_percent}pct'].iloc[hour],
            'PPA_Purchase_kWh': analysis_df[f'ppa_purchase_{ppa_percent}pct'].iloc[hour],
            'PPA_Excess_kWh': analysis_df[f'ppa_excess_{ppa_percent}pct'].iloc[hour],
            'KEPCO_Purchase_kWh': analysis_df[f'kepco_purchase_{ppa_percent}pct'].iloc[hour],
        }
        
        # Cost/Rate values (KRW)
        cost_values = {
            'KEPCO_Rate_KRW_per_kWh': analysis_df['kepco_rate'].iloc[hour],
            'PPA_Cost_KRW': analysis_df[f'ppa_cost_{ppa_percent}pct'].iloc[hour],
            'KEPCO_Cost_KRW': analysis_df[f'kepco_cost_{ppa_percent}pct'].iloc[hour],
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
long_df['Cost_Savings_vs_KEPCO_Only'] = (long_df['Load_Demand_kWh'] * long_df['KEPCO_Rate_KRW_per_kWh']) - long_df['Total_Cost_KRW']
long_df['Cost_per_kWh_KRW'] = (long_df['Total_Cost_KRW'] / long_df['Load_Demand_kWh']).fillna(0)

# Create annual summary for each PPA scenario  
annual_summary = []
for ppa_percent in range(0, 201, 10):
    annual_summary.append({
        'PPA_Coverage': f"{ppa_percent}%",
        'Annual_PPA_Gen': analysis_df[f'ppa_gen_{ppa_percent}pct'].sum(),
        'Annual_PPA_Purchase': analysis_df[f'ppa_purchase_{ppa_percent}pct'].sum(),
        'Annual_PPA_Cost': analysis_df[f'ppa_cost_{ppa_percent}pct'].sum(),
        'Annual_KEPCO_Purchase': analysis_df[f'kepco_purchase_{ppa_percent}pct'].sum(),
        'Annual_KEPCO_Cost': analysis_df[f'kepco_cost_{ppa_percent}pct'].sum(),
        'Annual_PPA_Excess': analysis_df[f'ppa_excess_{ppa_percent}pct'].sum(),
        'Annual_Resell_Revenue': analysis_df[f'resell_revenue_{ppa_percent}pct'].sum(),
        'Annual_Total_Cost': analysis_df[f'total_cost_{ppa_percent}pct'].sum(),
        'Load_Coverage_Pct': (analysis_df[f'ppa_purchase_{ppa_percent}pct'].sum() / analysis_df['load'].sum() * 100)
    })

annual_summary_df = pd.DataFrame(annual_summary)

# Cost comparison results
cost_comparison = pd.DataFrame({
    'PPA_Coverage': [f"{i}%" for i in range(0, 201, 10)],
    'Total_Cost': [result['total_cost'] for result in results_summary],
    'PPA_Cost': [result['ppa_cost'] for result in results_summary],
    'KEPCO_Cost': [result['kepco_cost'] for result in results_summary],
    'ESS_Cost': [result['ess_cost'] for result in results_summary]
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
print(f"Annual load: {analysis_df['load'].sum():,.0f} kWh")
print(f"Average KEPCO rate: {analysis_df['kepco_rate'].mean():.1f} KRW/kWh")
print(f"KEPCO rate range: {analysis_df['kepco_rate'].min():.1f} - {analysis_df['kepco_rate'].max():.1f} KRW/kWh")
print(f"PPA price: {ppa_price} KRW/kWh")
print(f"Average savings per kWh: {analysis_df['kepco_rate'].mean() - ppa_price:.1f} KRW/kWh")

# Peak vs off-peak analysis based on actual KEPCO rates
hourly_avg_rates = analysis_df.groupby('hour')['kepco_rate'].mean().sort_values(ascending=False)
peak_hours = hourly_avg_rates.head(12).index.tolist()  # Top 12 most expensive hours
offpeak_hours = hourly_avg_rates.tail(12).index.tolist()  # Bottom 12 cheapest hours

peak_data = analysis_df[analysis_df['hour'].isin(peak_hours)]
offpeak_data = analysis_df[analysis_df['hour'].isin(offpeak_hours)]

print(f"\nPeak hours {sorted(peak_hours)}: Avg KEPCO rate = {peak_data['kepco_rate'].mean():.1f} KRW/kWh")
print(f"Off-peak hours {sorted(offpeak_hours)}: Avg KEPCO rate = {offpeak_data['kepco_rate'].mean():.1f} KRW/kWh")
print(f"Peak hour savings with PPA: {peak_data['kepco_rate'].mean() - ppa_price:.1f} KRW/kWh")
