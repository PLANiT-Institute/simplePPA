"""
Core PPA cost calculation engine.
"""

def calculate_ppa_cost(load_df, ppa_df, grid_df, load_capacity_mw, ppa_coverage,
                       contract_fee, ppa_price, ppa_mintake, ppa_resell,
                       ppa_resellrate, ess_price, ess_capacity=0, verbose=False):
    """
    Calculate total power cost for given load capacity and PPA coverage.

    Parameters
    ----------
    load_df : pd.DataFrame
        DataFrame with hourly load demand (normalized 0-1)
    ppa_df : pd.DataFrame
        DataFrame with hourly solar generation pattern (normalized 0-1)
    grid_df : pd.DataFrame
        DataFrame with hourly Grid rates (KRW/kWh)
    load_capacity_mw : float
        Maximum load capacity in MW (e.g., 100 = 100MW peak load)
    ppa_coverage : float
        PPA sizing as percentage of peak load (0-2.0, where 1.0 = 100% of peak load)
    contract_fee : float
        Grid contract fee (KRW/kW) applied to peak Grid demand
    ppa_price : float
        PPA price in KRW/kWh
    ppa_mintake : float
        Minimum required purchase (1.0 = 100%, 0.5 = 50%)
    ppa_resell : bool
        Whether buyer is allowed to resell excess energy
    ppa_resellrate : float
        Resale rate as fraction of PPA price
    ess_price : float
        ESS discharge cost as fraction of PPA price
    ess_capacity : float, optional
        ESS capacity in kWh (0 = no ESS)
    verbose : bool, optional
        If True, print detailed statistics

    Returns
    -------
    tuple
        (total_cost, ppa_cost, grid_total_cost, grid_demand_cost, ess_cost)
    """

    load = load_df['load']
    grid_rate = grid_df['rate']
    solar_generation = ppa_df['generation']

    # Scale load by the specified capacity
    load_mw = load * load_capacity_mw

    # Scale PPA generation based on coverage percentage of the load capacity
    ppa_generation_mw = solar_generation * load_capacity_mw * ppa_coverage

    # Convert MW to kWh for cost calculations
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

    # Tracking variables
    total_ppa_purchased = 0
    total_ppa_excess = 0
    total_ppa_resold = 0
    total_load_met_by_ppa = 0

    for hour in range(len(load_kwh)):
        hour_load = load_kwh.iloc[hour]
        hour_grid_rate = grid_rate.iloc[hour]
        hour_ppa_gen = ppa_generation_kwh.iloc[hour]

        # Mandatory PPA purchase (minimum take)
        mandatory_ppa = hour_ppa_gen * ppa_mintake

        # Optional PPA purchase (if mintake < 100%)
        optional_ppa_available = hour_ppa_gen - mandatory_ppa
        optional_ppa_purchased = 0

        if optional_ppa_available > 0:
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
            # We have excess PPA energy
            excess_ppa = abs(remaining_load)
            total_ppa_excess += excess_ppa

            # Try to store excess in ESS
            if ess_storage < max_ess_capacity:
                ess_charge = min(excess_ppa, max_ess_capacity - ess_storage)
                ess_storage += ess_charge
                excess_ppa -= ess_charge

            # Try to resell remaining excess
            if excess_ppa > 0 and ppa_resell:
                resell_revenue = excess_ppa * ppa_price * ppa_resellrate
                ppa_cost -= resell_revenue
                total_ppa_resold += excess_ppa
                excess_ppa = 0

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
                grid_demand_kw = remaining_load
                peak_grid_demand_kw = max(peak_grid_demand_kw, grid_demand_kw)

    # Calculate Grid demand charge
    grid_demand_cost = peak_grid_demand_kw * contract_fee

    # Calculate total costs
    grid_total_cost = grid_energy_cost + grid_demand_cost
    total_cost = ppa_cost + grid_total_cost + ess_cost

    # Print statistics if verbose
    if verbose and ppa_coverage > 0:
        actual_coverage = total_load_met_by_ppa / load_kwh.sum() if load_kwh.sum() > 0 else 0
        annual_ppa_gen = ppa_generation_kwh.sum()
        ppa_capacity = load_capacity_mw * ppa_coverage

        print(f"  PPA Stats: Load={load_capacity_mw:.1f}MW, PPA={ppa_capacity:.1f}MW "
              f"({ppa_coverage*100:.0f}%), Gen={annual_ppa_gen:,.0f} kWh, "
              f"Load coverage={actual_coverage*100:.1f}%, "
              f"Purchased={total_ppa_purchased:,.0f} kWh, "
              f"Excess={total_ppa_excess:,.0f} kWh, "
              f"Resold={total_ppa_resold:,.0f} kWh")
        print(f"  Grid Stats: Peak demand={peak_grid_demand_kw:,.0f} kW, "
              f"Energy cost={grid_energy_cost:,.0f} KRW, "
              f"Demand cost={grid_demand_cost:,.0f} KRW")

    return total_cost, ppa_cost, grid_total_cost, grid_demand_cost, ess_cost
