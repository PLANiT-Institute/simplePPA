"""
SimplePPA - Streamlit GUI for Power Purchase Agreement Analysis

Interactive web-based interface for PPA analysis with all configurable parameters.
Run with: streamlit run main_gui.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots

import libs.KEPCOutils as kepco
from libs.config import get_default_config, validate_config, load_app_settings
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
from libs.exporter import export_to_excel, export_to_excel_bytes


# Page configuration
st.set_page_config(
    page_title="SimplePPA Analysis",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

DEFAULT_PLOTLY_CONFIG = {"responsive": True}


def render_plotly_chart(fig, *, use_container_width=True, config=None):
    """Render a Plotly figure with unified configuration handling."""
    final_config = DEFAULT_PLOTLY_CONFIG.copy()
    if config:
        final_config.update(config)
    st.plotly_chart(fig, use_container_width=use_container_width, config=final_config)

# Initialize session state
if 'analysis_done' not in st.session_state:
    st.session_state.analysis_done = False
if 'results_summary' not in st.session_state:
    st.session_state.results_summary = None
if 'config' not in st.session_state:
    st.session_state.config = get_default_config()


def display_documentation():
    """Display comprehensive documentation and user manual."""
    st.header("📖 SimplePPA Documentation & User Manual")

    # Overview
    st.subheader("Overview")
    st.markdown("""
    SimplePPA is a comprehensive tool for analyzing Power Purchase Agreement (PPA) scenarios with solar generation,
    grid electricity, and optional Energy Storage Systems (ESS). It helps energy buyers evaluate the economic impact
    of different PPA coverage levels by simulating hourly energy flows and calculating total costs.
    """)

    # Key Terminology
    st.subheader("📚 Key Terminology")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        **Energy Terms:**
        - **PPA (Power Purchase Agreement)**: Contract to buy electricity from solar farm at fixed price
        - **Load**: Your electricity demand/consumption (kWh)
        - **PPA Coverage**: Solar farm capacity relative to your peak load (e.g., 100% = solar farm size equals your peak)
        - **ESS (Energy Storage System)**: Battery system to store excess solar energy

        **PPA Contract Terms:**
        - **Minimum Take**: Percentage of solar generation you must purchase each hour (typically 100%)
        - **Optional Purchase**: Additional solar energy you can buy if cheaper than grid (when minimum take < 100%)
        - **Resale**: Ability to sell excess solar energy back to grid
        - **Resale Rate**: Price received when reselling (typically 90% of PPA price)
        """)

    with col2:
        st.markdown("""
        **Cost Components:**
        - **PPA Cost**: Payment for solar energy purchased (KRW)
        - **Grid Energy Cost**: Payment for electricity from grid based on kWh consumed (KRW)
        - **Contract Fee**: Grid demand charge based on peak kW demand (KRW)
        - **ESS Cost**: Operating cost for discharging stored energy (KRW)
        - **Carbon Cost**: Cost of carbon emissions (KRW) = Emissions (tCO2e) × Carbon Price (KRW/tCO2e)

        **Grid Charges:**
        - **Energy Charge**: Pay per kWh consumed from grid (variable by time of day/season)
        - **Contract Fee (Demand Charge)**: Pay for peak power demand (kW) during the period
            - This is a capacity charge, not consumption
            - Based on your single highest hourly demand from grid
            - From KEPCO.xlsx "contract" sheet (KRW/kW)

        **Carbon Pricing:**
        - Optional carbon price per ton of CO2 equivalent emissions
        - Applied to total emissions from grid electricity
        - PPA solar energy assumed zero-emission
        - Helps evaluate total cost including environmental externalities
        """)

    # Column Definitions
    st.subheader("📊 Output Column Definitions")

    st.markdown("**Annual Summary Columns:**")

    col_definitions = {
        "Column Name": [
            "PPA_Coverage (%)",
            "Annual_PPA_Gen (kWh)",
            "Annual_Mandatory_PPA (kWh)",
            "Annual_Optional_PPA (kWh)",
            "Annual_PPA_Purchase (kWh)",
            "Annual_PPA_Cost (KRW)",
            "Annual_Grid_Purchase (kWh)",
            "Annual_Grid_Cost (KRW)",
            "Annual_PPA_Excess (kWh)",
            "Annual_Resell_Revenue (KRW)",
            "Annual_Total_Cost (KRW)",
            "Load_Coverage (%)",
            "PPA_Cost (KRW/kWh)",
            "Grid_Cost (KRW/kWh)",
            "Total_Cost (KRW/kWh)"
        ],
        "Description": [
            "PPA solar farm capacity as % of your peak load",
            "Total solar energy generated by PPA farm",
            "Energy that MUST be purchased per contract (minimum take)",
            "Additional energy purchased because it's cheaper than grid",
            "Total energy purchased from PPA (Mandatory + Optional)",
            "Total cost for PPA purchases (minus resale revenue if applicable)",
            "Energy purchased from grid to meet remaining demand",
            "Total grid cost (energy charges + contract fee)",
            "Excess PPA energy (stored in ESS, resold, or wasted)",
            "Revenue from reselling excess PPA energy back to grid",
            "Total annual electricity cost (PPA + Grid + ESS)",
            "Percentage of your load demand met by PPA purchases",
            "PPA cost per unit of total load demand",
            "Grid cost per unit of total load demand",
            "Total cost per unit of total load demand"
        ]
    }

    st.dataframe(pd.DataFrame(col_definitions), width="stretch", hide_index=True)

    st.markdown("**Cost Analysis Columns:**")

    cost_definitions = {
        "Column Name": [
            "PPA_Coverage (%)",
            "Total_Cost (KRW)",
            "PPA_Cost (KRW)",
            "Grid_Energy_Cost (KRW)",
            "Contract_Fee (KRW)",
            "ESS_Cost (KRW)",
            "Total_Cost (KRW/kWh)",
            "PPA_Cost (KRW/kWh)",
            "Grid_Energy_Cost (KRW/kWh)",
            "Contract_Fee (KRW/kWh)",
            "ESS_Cost (KRW/kWh)"
        ],
        "Description": [
            "PPA sizing as % of peak load",
            "Total annual cost for all electricity",
            "Cost for PPA purchases (net of resale)",
            "Grid energy charges only (excludes contract fee)",
            "Grid demand charge based on peak kW demand",
            "Cost for ESS discharge operations",
            "Total cost divided by annual load demand",
            "PPA cost divided by annual load demand",
            "Grid energy cost divided by annual load demand",
            "Contract fee divided by annual load demand",
            "ESS cost divided by annual load demand"
        ]
    }

    st.dataframe(pd.DataFrame(cost_definitions), width="stretch", hide_index=True)

    # How to Use
    st.subheader("🚀 How to Use SimplePPA")

    st.markdown("""
    **Step 1: Prepare Input Data**
    - `data/pattern.xlsx`: Hourly load and solar patterns (normalized 0-1)
    - `data/KEPCO.xlsx`: Grid rates and contract fees

    **Step 2: Configure Parameters**
    - Set load capacity, PPA price, and contract terms in sidebar
    - Choose PPA coverage range to analyze (e.g., 0-200% in 10% steps)
    - Enable ESS if needed

    **Step 3: Review Input Data (Optional)**
    - Click "🔍 Review Data" to visualize your input patterns
    - Check load/solar correlation and grid rate patterns

    **Step 4: Run Analysis**
    - Click "🚀 Run Analysis" to compute all scenarios
    - Tool simulates 8,760 hours (1 year) for each PPA coverage level

    **Step 5: Review Results**
    - View optimal PPA coverage and costs
    - Analyze cost breakdown by component
    - Compare scenarios with interactive charts
    - Export detailed data to Excel
    """)

    # Algorithm Overview
    st.subheader("⚙️ How the Algorithm Works")

    st.markdown("""
    SimplePPA simulates hourly energy procurement for an entire year (8,760 hours). For each hour:

    **1. Scale Patterns to Actual Capacity**
    - Load (kWh) = Normalized Load × Load Capacity (MW) × 1000
    - PPA Generation (kWh) = Normalized Solar × Load Capacity (MW) × PPA Coverage × 1000

    **2. PPA Purchase Decision**
    - **Mandatory Purchase**: `PPA Generation × Minimum Take %` (MUST buy)
    - **Optional Purchase**: Buy additional if PPA price < grid rate AND you need energy
    - **Total PPA Purchase**: Mandatory + Optional
    - **PPA Cost**: Total Purchase × PPA Price

    **3. Energy Balance**

    If you have **excess PPA energy** (purchased more than needed):
    - Try to store in ESS (if available and not full)
    - Try to resell to grid (if enabled) at `PPA Price × Resale Rate`
    - Otherwise, excess is wasted (but already paid for)

    If you have **energy deficit** (need more than PPA):
    - Discharge from ESS (if available) at cost = `PPA Price × ESS Discharge Price`
    - Buy remaining from grid at hourly grid rate
    - Track peak grid demand for contract fee calculation

    **4. Calculate Contract Fee**
    - After all hours, find peak grid demand (kW)
    - Contract Fee = Peak Grid Demand (kW) × Contract Fee Rate (KRW/kW)

    **5. Calculate Carbon Cost (if carbon pricing enabled)**
    - Carbon Cost = (Total Emissions in kgCO2e / 1000) × Carbon Price (KRW/tCO2e)

    **6. Total Cost**
    - Total Cost = PPA Cost + Grid Energy Cost + Contract Fee + ESS Cost
    - Total Cost with Carbon = Total Cost + Carbon Cost (if carbon pricing enabled)
    """)

    with st.expander("📐 Example Calculation"):
        st.markdown("""
        **Scenario Setup:**
        - Load Capacity: 100 MW
        - PPA Coverage: 120% (120 MW solar farm)
        - PPA Price: 170 KRW/kWh
        - Grid Rate this hour: 220 KRW/kWh
        - Minimum Take: 100%
        - Contract Fee: 8,000 KRW/kW

        **Hour 1 (Sunny afternoon):**
        - Load: 0.7 normalized → 70,000 kWh needed
        - Solar: 0.8 normalized → 96,000 kWh generated
        - Must buy: 96,000 kWh × 100% = 96,000 kWh
        - PPA Cost: 96,000 × 170 = 16,320,000 KRW
        - Excess: 96,000 - 70,000 = 26,000 kWh
        - Store 10,000 kWh in ESS, waste 16,000 kWh
        - Grid purchase: 0 kWh

        **Hour 2 (Evening, no sun):**
        - Load: 0.9 normalized → 90,000 kWh needed
        - Solar: 0.3 normalized → 36,000 kWh generated
        - Must buy: 36,000 kWh
        - PPA Cost: 36,000 × 170 = 6,120,000 KRW
        - Deficit: 90,000 - 36,000 = 54,000 kWh
        - Discharge ESS: 10,000 kWh at 85 KRW/kWh = 850,000 KRW
        - Buy from grid: 44,000 kWh × 220 = 9,680,000 KRW
        - Peak grid demand: 44,000 kW

        **After 8,760 hours:**
        - Peak grid demand was 44,000 kW
        - Contract Fee: 44,000 × 8,000 = 352,000,000 KRW
        - Total Cost: Sum of all hourly costs + Contract Fee
        """)

    # Data Sources
    st.subheader("📁 Input Data Requirements")

    st.markdown("""
    **pattern.xlsx** (required columns):
    - `load`: Normalized hourly load (0-1 scale, where 1 = peak load)
    - `solar`: Normalized hourly solar generation (0-1 scale, where 1 = peak generation)
    - `emission`: Grid emission factor (kgCO2e/kWh) - carbon intensity of grid electricity
    - Must have 8,760 rows (one year, hourly)

    **KEPCO.xlsx** (required sheets):
    - `timezone`: Peak/off-peak hours by month
    - `season`: Month to season mapping
    - `contract`: Contract fees (KRW/kW) for each tariff type
    - `HV_C_I`, `HV_C_II`, `HV_C_III`: Energy rates by season and timezone

    **Emission Data**:
    - Emission factor represents carbon intensity of grid electricity (kgCO2e/kWh)
    - PPA solar energy assumed to be zero-emission (can be configured)
    - Tool calculates total emissions and emissions per kWh for each scenario
    - Helps evaluate environmental benefits of PPA adoption
    """)

    # Tips and Best Practices
    st.subheader("💡 Tips & Best Practices")

    st.markdown("""
    **Choosing PPA Coverage:**
    - Start with 0-200% range in 10% steps for overview
    - Use smaller steps (5% or 1%) around the optimal point for fine-tuning
    - Higher coverage = more solar but also more excess to manage

    **Understanding Results:**
    - Look for the "sweet spot" where total cost is minimized
    - Contract fee savings can be significant with PPA (reduces peak grid demand)
    - ESS is most valuable when there's large excess solar AND high grid rates

    **Interpreting Cost per kWh:**
    - This is total cost divided by total load (not just grid purchases)
    - Useful for comparing different scenarios on equal basis
    - Lower is better (optimal scenario has lowest cost per kWh)

    **Common Scenarios:**
    - **100% PPA + 100% minimum take**: Most common, simple contract
    - **120% PPA + 80% minimum take**: More solar, more flexibility, more complex
    - **150% PPA + ESS**: Oversized solar with storage for evening/peak use
    """)

    # Version and Contact
    st.divider()
    st.caption("SimplePPA v0.2 - Power Purchase Agreement Analysis Tool")


def main():
    st.title("⚡ SimplePPA - Power Purchase Agreement Analysis")
    st.markdown("Analyze PPA scenarios with solar generation, grid electricity, and optional ESS")

    # Add Documentation tab at the top
    main_tab1, main_tab2 = st.tabs(["📖 Documentation", "⚙️ Analysis Tool"])

    with main_tab1:
        display_documentation()

    with main_tab2:
        run_analysis_tool()


def run_analysis_tool():
    """Main analysis tool interface."""
    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        app_settings = load_app_settings()

        # Data Files
        st.subheader("📁 Data Files")
        pattern_file = st.text_input(
            "Pattern File",
            value=st.session_state.config.get('pattern_file', 'data/pattern.xlsx'),
            help="Excel file containing normalized hourly load and solar generation patterns (0-1 scale)"
        )
        kepco_file = st.text_input(
            "KEPCO File",
            value=st.session_state.config.get('kepco_file', 'data/KEPCO.xlsx'),
            help="Excel file with Korean electricity tariff data including rates and contract fees"
        )
        kepco_year = st.number_input(
            "KEPCO Year",
            value=int(st.session_state.config.get('kepco_year', 2024)),
            step=1,
            help="Year for which the tariff applies"
        )
        kepco_tariff = st.selectbox(
            "KEPCO Tariff",
            ["HV_C_III", "HV_C_I", "HV_C_II"],
            index=["HV_C_III", "HV_C_I", "HV_C_II"].index(st.session_state.config.get('kepco_tariff', 'HV_C_III')),
            help="High Voltage tariff type:\n- HV_C_I: Option I\n- HV_C_II: Option II\n- HV_C_III: Option III (default)"
        )

        # Review Data Button
        review_button = st.button("🔍 Review Data", width="stretch")

        st.divider()

        # Analysis Period
        st.subheader("📅 Analysis Period")
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input(
                "Start Date",
                value=pd.to_datetime(st.session_state.config.get('start_date', '2024-01-01')),
                help="First date to include in the analysis period"
            )
        with col2:
            end_date = st.date_input(
                "End Date",
                value=pd.to_datetime(st.session_state.config.get('end_date', '2024-12-31')),
                help=f"Last date to include in the analysis period (max {app_settings.get('max_analysis_days', 31)} days)"
            )
        # Soft guard in UI for over-long ranges
        try:
            num_days = (pd.to_datetime(end_date) - pd.to_datetime(start_date)).days + 1
            if num_days > int(app_settings.get('max_analysis_days', 31)):
                st.warning(f"Selected range is {num_days} days. Maximum allowed is {int(app_settings.get('max_analysis_days', 31))} days.")
        except Exception:
            pass

        st.divider()

        # Load Parameters
        st.subheader("⚡ Load Parameters")
        load_capacity_mw = st.number_input(
            "Load Capacity (MW)",
            value=float(st.session_state.config.get('load_capacity_mw', 3000.0)),
            min_value=0.1,
            step=100.0,
            help="Peak load capacity in megawatts. This scales the normalized load pattern to actual power consumption. Example: 100 MW means when normalized load = 1.0, actual load = 100 MW"
        )

        st.divider()

        # PPA Parameters
        st.subheader("🌞 PPA Parameters")
        ppa_price = st.number_input(
            "PPA Price (KRW/kWh)",
            value=float(st.session_state.config.get('ppa_price', 170.0)),
            min_value=0.0,
            step=1.0,
            help="Fixed price per kWh for energy purchased from the PPA solar farm. This is the contracted rate you pay for solar electricity."
        )
        ppa_mintake = st.slider(
            "Minimum Take (%)",
            min_value=0,
            max_value=100,
            value=int(st.session_state.config.get('ppa_mintake', 1.0) * 100),
            step=1,
            help="Minimum percentage of PPA generation that MUST be purchased each hour, regardless of need.\n- 100% = Must buy all generation (typical)\n- 80% = Must buy 80%, can optionally buy up to 100% if cheaper than grid\n- Lower values provide flexibility but may cost more"
        ) / 100.0
        ppa_resell = st.checkbox(
            "Allow Reselling",
            value=bool(st.session_state.config.get('ppa_resell', False)),
            help="Enable reselling excess PPA energy back to the grid. If disabled, excess energy is wasted (but already paid for)."
        )
        ppa_resellrate = st.slider(
            "Resell Rate (%)",
            min_value=0,
            max_value=100,
            value=int(st.session_state.config.get('ppa_resellrate', 0.9) * 100),
            step=1,
            disabled=not ppa_resell,
            help="Percentage of PPA price received when reselling excess energy.\n- 90% = Resell at 90% of what you paid\n- Revenue reduces net PPA cost"
        ) / 100.0

        st.divider()

        # PPA Coverage Range
        st.subheader(
            "📊 PPA Coverage Range",
            help="PPA Coverage = (PPA Solar Farm Peak Capacity) / (Your Peak Load Capacity)\n\n"
                 "Examples:\n"
                 "• 0% = No PPA, grid only (baseline)\n"
                 "• 50% = Solar farm half your peak load\n"
                 "• 100% = Solar farm peak equals your peak load\n"
                 "• 150% = Solar farm 1.5× larger than peak load\n"
                 "• 200% = Solar farm twice your peak load\n\n"
                 "Higher coverage = More solar generation but also more excess energy to manage"
        )

        col1, col2, col3 = st.columns(3)
        with col1:
            ppa_range_start = st.number_input(
                "Start (%)",
                value=int(st.session_state.config.get('ppa_range_start', 0)),
                min_value=0,
                step=10,
                help="Starting PPA coverage percentage. 0% means no PPA (grid only), used as baseline."
            )
        with col2:
            ppa_range_end = st.number_input(
                "End (%)",
                value=int(st.session_state.config.get('ppa_range_end', 200)),
                min_value=0,
                step=10,
                help="Ending PPA coverage percentage. 200% means solar farm twice as large as peak load."
            )
        with col3:
            ppa_range_step = st.number_input(
                "Step (%)",
                value=int(st.session_state.config.get('ppa_range_step', 10)),
                min_value=1,
                step=1,
                help="Increment between scenarios. Smaller steps = more detailed analysis but longer computation time.\n- 10% = Fast (21 scenarios for 0-200%)\n- 5% = Medium (41 scenarios)\n- 1% = Detailed (201 scenarios)"
            )

        num_scenarios = len(range(int(ppa_range_start), int(ppa_range_end) + 1, int(ppa_range_step)))
        st.info(f"📈 {num_scenarios} scenarios will be analyzed")

        st.divider()

        # Carbon Pricing
        st.subheader("🌍 Carbon Pricing")
        carbon_price = st.number_input(
            "Carbon Price (KRW/tCO2e)",
            value=float(st.session_state.config.get('carbon_price', 0.0)),
            min_value=0.0,
            step=1000.0,
            help="Price per ton of CO2 equivalent emissions. Used to calculate the cost of carbon emissions. Set to 0 to exclude carbon pricing from cost analysis."
        )

        st.divider()

        # ESS Parameters
        st.subheader("🔋 ESS Parameters")
        ess_include = st.checkbox(
            "Include ESS Analysis",
            value=bool(st.session_state.config.get('ess_include', False)),
            help="Enable Energy Storage System analysis. ESS stores excess PPA energy for later use, reducing grid purchases and demand charges."
        )
        ess_capacity = st.slider(
            "ESS Capacity (% of solar peak)",
            min_value=0,
            max_value=200,
            value=int(st.session_state.config.get('ess_capacity', 0.5) * 100),
            step=10,
            disabled=not ess_include,
            help="ESS storage capacity as percentage of peak solar generation.\n- 50% = Can store up to half of peak solar output\n- 100% = Can store full peak solar output\n- Larger ESS = More flexibility but higher capital cost"
        ) / 100.0
        ess_price = st.slider(
            "ESS Discharge Price (% of PPA price)",
            min_value=0,
            max_value=100,
            value=int(st.session_state.config.get('ess_price', 0.5) * 100),
            step=5,
            disabled=not ess_include,
            help="Operating cost for using stored energy, as percentage of PPA price.\n- 50% = Discharging costs 50% of PPA price (typical)\n- Accounts for efficiency losses and O&M costs\n- Does not include capital cost (assumed external)"
        ) / 100.0

        st.divider()

        # Output Options
        st.subheader("💾 Output Options")
        output_file = st.text_input(
            "Output Filename",
            value=st.session_state.config.get('output_file', 'ppa_analysis_results.xlsx'),
            help="Name of Excel file to save detailed analysis results. Contains hourly data, annual summaries, and cost breakdowns."
        )
        export_long_format = st.checkbox(
            "Export Long Format",
            value=bool(st.session_state.config.get('export_long_format', True)),
            help="Generate detailed hourly data for all scenarios in pivot-table-ready format. Enables peak analysis but increases computation time."
        )
        verbose = st.checkbox(
            "Verbose Output",
            value=bool(st.session_state.config.get('verbose', False)),
            help="Print detailed statistics for each scenario to console (useful for debugging)"
        )

        st.divider()

        # Run Analysis Button
        run_button = st.button("🚀 Run Analysis", type="primary", width="stretch")

    # Main content area
    # Review Data functionality
    if review_button:
        review_input_data(pattern_file, kepco_file, kepco_year, kepco_tariff)

    if run_button:
        # Build configuration
        config = {
            'pattern_file': pattern_file,
            'kepco_file': kepco_file,
            'kepco_year': int(kepco_year),
            'kepco_tariff': kepco_tariff,
            'start_date': start_date.strftime('%Y-%m-%d'),
            'end_date': end_date.strftime('%Y-%m-%d'),
            'load_capacity_mw': load_capacity_mw,
            'ppa_price': ppa_price,
            'ppa_mintake': ppa_mintake,
            'ppa_resell': ppa_resell,
            'ppa_resellrate': ppa_resellrate,
            'ppa_range_start': int(ppa_range_start),
            'ppa_range_end': int(ppa_range_end),
            'ppa_range_step': int(ppa_range_step),
            'carbon_price': carbon_price,
            'ess_include': ess_include,
            'ess_capacity': ess_capacity,
            'ess_price': ess_price,
            'output_file': output_file,
            'verbose': verbose,
            'export_long_format': export_long_format,
            'max_analysis_days': int(app_settings.get('max_analysis_days', 31))
        }

        # Validate configuration
        is_valid, errors = validate_config(config)
        if not is_valid:
            st.error("Configuration errors:")
            for error in errors:
                st.error(f"  - {error}")
            return

        # Run analysis
        with st.spinner("🔄 Loading data..."):
            try:
                # Load data
                load_df, solar_df, ppa_df, emission_df = load_pattern_data(config['pattern_file'])
                grid_df, contract_fee = kepco.process_kepco_data(
                    config['kepco_file'],
                    config['kepco_year'],
                    config['kepco_tariff']
                )
                st.success("✅ Data loaded successfully!")
            except Exception as e:
                st.error(f"❌ Error loading data: {str(e)}")
                return

        with st.spinner("🔄 Running scenario analysis..."):
            try:
                # Run base analysis
                results_summary = run_scenario_analysis(
                    load_df, ppa_df, grid_df, emission_df, contract_fee, config,
                    verbose=config['verbose']
                )

                optimal_ppa, optimal_cost = find_optimal_scenario(results_summary)

                # Store in session state
                st.session_state.results_summary = results_summary
                st.session_state.optimal_ppa = optimal_ppa
                st.session_state.optimal_cost = optimal_cost
                st.session_state.config = config

                st.success(f"✅ Analysis complete! Optimal: {optimal_ppa}% PPA at {optimal_cost:,.0f} KRW")
            except Exception as e:
                st.error(f"❌ Error during analysis: {str(e)}")
                return

        # ESS Analysis
        if config['ess_include']:
            with st.spinner("🔄 Running ESS analysis..."):
                try:
                    results_ess, ess_capacity_kwh, peak_solar_mw = run_ess_analysis(
                        load_df, ppa_df, grid_df, emission_df, contract_fee, config,
                        verbose=config['verbose']
                    )

                    optimal_ess_ppa, optimal_ess_cost = find_optimal_scenario(results_ess)

                    st.session_state.results_ess = results_ess
                    st.session_state.optimal_ess_ppa = optimal_ess_ppa
                    st.session_state.optimal_ess_cost = optimal_ess_cost
                    st.session_state.ess_capacity_kwh = ess_capacity_kwh
                    st.session_state.peak_solar_mw = peak_solar_mw

                    savings = optimal_cost - optimal_ess_cost
                    st.success(f"✅ ESS Analysis complete! Optimal: {optimal_ess_ppa}% PPA, Savings: {savings:,.0f} KRW")
                except Exception as e:
                    st.error(f"❌ Error during ESS analysis: {str(e)}")

        # Generate detailed data for export
        if config['export_long_format']:
            with st.spinner("🔄 Generating detailed data..."):
                try:
                    analysis_df = create_analysis_dataframe(
                        grid_df, load_df, ppa_df, emission_df,
                        config['start_date'], config['end_date'],
                        config['load_capacity_mw']
                    )

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

                    long_df = create_long_format_dataframe(
                        analysis_df,
                        config['ppa_range_start'],
                        config['ppa_range_end'],
                        config['ppa_range_step']
                    )

                    annual_summary_df = create_annual_summary(
                        analysis_df,
                        config['ppa_range_start'],
                        config['ppa_range_end'],
                        config['ppa_range_step']
                    )

                    cost_comparison_df = create_cost_comparison(results_summary)

                    peak_analysis = analyze_peak_hours(analysis_df)

                    # Store in session state
                    st.session_state.analysis_df = analysis_df
                    st.session_state.long_df = long_df
                    st.session_state.annual_summary_df = annual_summary_df
                    st.session_state.cost_comparison_df = cost_comparison_df
                    st.session_state.peak_analysis = peak_analysis

                    st.success("✅ Detailed data generated!")
                except Exception as e:
                    st.error(f"❌ Error generating detailed data: {str(e)}")

        st.session_state.analysis_done = True

    # Display results if analysis has been done
    if st.session_state.analysis_done and st.session_state.results_summary is not None:
        display_results()


def review_input_data(pattern_file, kepco_file, kepco_year, kepco_tariff):
    """Review and visualize input data files."""
    st.header("🔍 Input Data Review")

    try:
        # Load pattern data
        with st.spinner("Loading pattern data..."):
            load_df, solar_df, ppa_df, emission_df = load_pattern_data(pattern_file)
            pattern_df = pd.read_excel(pattern_file, index_col=0)

        # Load KEPCO data
        with st.spinner("Loading KEPCO data..."):
            grid_df, contract_fee = kepco.process_kepco_data(
                kepco_file,
                int(kepco_year),
                kepco_tariff
            )

        st.success("✅ Data loaded successfully!")

        # Create tabs for different datasets
        tab1, tab2, tab3 = st.tabs([
            "📊 Load & Solar Patterns",
            "⚡ Grid Rate Data",
            "📈 Summary Statistics"
        ])

        with tab1:
            st.subheader("Load and Solar Generation Patterns")

            # Show data info
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Data Points", f"{len(pattern_df):,}")
                st.metric("Days of Data", f"{len(pattern_df)/24:.1f}")
            with col2:
                st.metric("Peak Load", f"{pattern_df['load'].max():.3f}")
                st.metric("Peak Solar", f"{pattern_df['solar'].max():.3f}")

            # Plot patterns
            st.subheader("Hourly Patterns (First 7 Days)")
            fig = make_subplots(
                rows=2, cols=1,
                subplot_titles=("Load Pattern", "Solar Generation Pattern"),
                vertical_spacing=0.12
            )

            # Show first week
            hours_to_show = min(168, len(pattern_df))  # 7 days or less
            hours = list(range(hours_to_show))

            # Load pattern
            fig.add_trace(
                go.Scatter(
                    x=hours,
                    y=pattern_df['load'].iloc[:hours_to_show],
                    name='Load',
                    line=dict(color='blue', width=2)
                ),
                row=1, col=1
            )

            # Solar pattern
            fig.add_trace(
                go.Scatter(
                    x=hours,
                    y=pattern_df['solar'].iloc[:hours_to_show],
                    name='Solar',
                    line=dict(color='orange', width=2),
                    fill='tozeroy'
                ),
                row=2, col=1
            )

            fig.update_xaxes(title_text="Hour", row=2, col=1)
            fig.update_yaxes(title_text="Normalized Load", row=1, col=1)
            fig.update_yaxes(title_text="Normalized Generation", row=2, col=1)
            fig.update_layout(height=600, showlegend=True)

            render_plotly_chart(fig)

            # Daily average pattern
            st.subheader("Average Daily Pattern")
            if len(pattern_df) >= 24:
                # Calculate average for each hour of day
                pattern_df_copy = pattern_df.copy()
                pattern_df_copy['hour'] = pattern_df_copy.index.hour if hasattr(pattern_df_copy.index, 'hour') else [i % 24 for i in range(len(pattern_df_copy))]
                daily_avg = pattern_df_copy.groupby('hour').mean()

                fig_daily = go.Figure()
                fig_daily.add_trace(go.Scatter(
                    x=daily_avg.index,
                    y=daily_avg['load'],
                    name='Avg Load',
                    line=dict(color='blue', width=3),
                    mode='lines+markers'
                ))
                fig_daily.add_trace(go.Scatter(
                    x=daily_avg.index,
                    y=daily_avg['solar'],
                    name='Avg Solar',
                    line=dict(color='orange', width=3),
                    mode='lines+markers'
                ))
                fig_daily.update_layout(
                    xaxis_title="Hour of Day",
                    yaxis_title="Normalized Value",
                    height=400,
                    hovermode='x unified'
                )
                render_plotly_chart(fig_daily)

            # Show data table
            st.subheader("Data Preview (First 100 Rows)")
            st.dataframe(pattern_df.head(100), width="stretch")

            # Download button
            csv = pattern_df.to_csv()
            st.download_button(
                "📥 Download Full Pattern Data (CSV)",
                csv,
                "pattern_data.csv",
                "text/csv"
            )

        with tab2:
            st.subheader("Grid Electricity Rate Data")

            # Show data info
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Hours", f"{len(grid_df):,}")
                st.metric("Date Range", f"{len(grid_df)/24:.1f} days")
            with col2:
                st.metric("Min Rate", f"{grid_df['rate'].min():.2f} KRW/kWh")
                st.metric("Max Rate", f"{grid_df['rate'].max():.2f} KRW/kWh")
            with col3:
                st.metric("Avg Rate", f"{grid_df['rate'].mean():.2f} KRW/kWh")
                st.metric("Contract Fee", f"{contract_fee:,.0f} KRW/kW")

            # Plot rate over time
            st.subheader("Grid Rate Over Time (First 30 Days)")
            hours_to_show = min(720, len(grid_df))  # 30 days or less

            fig_rate = go.Figure()
            fig_rate.add_trace(go.Scatter(
                x=list(range(hours_to_show)),
                y=grid_df['rate'].iloc[:hours_to_show],
                mode='lines',
                name='Grid Rate',
                line=dict(color='red', width=1),
                fill='tozeroy'
            ))
            fig_rate.update_layout(
                xaxis_title="Hour",
                yaxis_title="Rate (KRW/kWh)",
                height=400,
                hovermode='x'
            )
            render_plotly_chart(fig_rate)

            # Rate distribution
            st.subheader("Rate Distribution")
            fig_hist = go.Figure()
            fig_hist.add_trace(go.Histogram(
                x=grid_df['rate'],
                nbinsx=50,
                name='Rate Distribution',
                marker=dict(color='red', line=dict(color='darkred', width=1))
            ))
            fig_hist.update_layout(
                xaxis_title="Rate (KRW/kWh)",
                yaxis_title="Frequency",
                height=400
            )
            render_plotly_chart(fig_hist)

            # Average rate by hour of day
            st.subheader("Average Rate by Hour of Day")
            grid_df_copy = grid_df.copy()
            grid_df_copy['hour'] = grid_df_copy.index.hour if hasattr(grid_df_copy.index, 'hour') else [i % 24 for i in range(len(grid_df_copy))]
            hourly_avg = grid_df_copy.groupby('hour')['rate'].mean()

            fig_hourly = go.Figure()
            fig_hourly.add_trace(go.Bar(
                x=hourly_avg.index,
                y=hourly_avg.values,
                marker=dict(
                    color=hourly_avg.values,
                    colorscale='Reds',
                    showscale=True,
                    colorbar=dict(title="KRW/kWh")
                )
            ))
            fig_hourly.update_layout(
                xaxis_title="Hour of Day",
                yaxis_title="Average Rate (KRW/kWh)",
                height=400
            )
            render_plotly_chart(fig_hourly)

            # Show data table
            st.subheader("Data Preview (First 100 Rows)")
            st.dataframe(grid_df.head(100), width="stretch")

            # Download button
            csv = grid_df.to_csv()
            st.download_button(
                "📥 Download Grid Rate Data (CSV)",
                csv,
                "grid_rate_data.csv",
                "text/csv"
            )

        with tab3:
            st.subheader("Summary Statistics")

            # Pattern statistics
            st.write("### Load & Solar Pattern Statistics")
            stats_df = pattern_df.describe()
            st.dataframe(stats_df, width="stretch")

            # Grid rate statistics
            st.write("### Grid Rate Statistics")
            grid_stats_df = grid_df['rate'].describe().to_frame()
            st.dataframe(grid_stats_df, width="stretch")

            # Correlation if datetime index available
            st.write("### Pattern Correlation")
            correlation = pattern_df['load'].corr(pattern_df['solar'])
            st.metric("Load vs Solar Correlation", f"{correlation:.3f}")

            if correlation < 0:
                st.info("Negative correlation: Solar generation peaks when load is lower (typical for residential)")
            elif correlation > 0.5:
                st.info("Strong positive correlation: Solar generation matches load demand well (typical for commercial)")
            else:
                st.info("Weak correlation: Solar generation and load are somewhat independent")

    except FileNotFoundError as e:
        st.error(f"❌ File not found: {str(e)}")
        st.info("Please check that the file paths are correct and files exist.")
    except Exception as e:
        st.error(f"❌ Error loading data: {str(e)}")
        st.exception(e)


def display_results():
    """Display analysis results with interactive charts."""

    st.header("📊 Analysis Results")

    # Summary metrics
    # Get optimal result details for cost per kWh
    optimal_result = next(r for r in st.session_state.results_summary if r['ppa_percent'] == st.session_state.optimal_ppa)

    # Check if carbon pricing is enabled
    has_carbon_price = st.session_state.config.get('carbon_price', 0.0) > 0

    # Create summary table
    st.subheader("Optimal Scenario Summary")

    summary_data = []

    # Without ESS results
    summary_data.append({
        'Scenario': 'Optimal (No ESS)',
        'PPA Coverage (%)': f"{st.session_state.optimal_ppa}%",
        'Total Cost (M KRW)': f"{st.session_state.optimal_cost/1e6:.1f}",
        'Cost per kWh (KRW/kWh)': f"{(optimal_result['total_cost_with_carbon_per_kwh'] if has_carbon_price else optimal_result['total_cost_per_kwh']):.2f}",
        'Emissions (tCO2e)': f"{optimal_result['total_emissions']/1000:.1f}",
        'Emissions (kgCO2e/kWh)': f"{optimal_result['emissions_per_kwh']:.3f}"
    })

    # Add cost breakdown columns if carbon pricing is enabled
    if has_carbon_price:
        summary_data[0]['Electricity Cost (M KRW)'] = f"{optimal_result['total_cost']/1e6:.1f}"
        summary_data[0]['Carbon Cost (M KRW)'] = f"{optimal_result['carbon_cost']/1e6:.1f}"
        summary_data[0]['Electricity (KRW/kWh)'] = f"{optimal_result['total_cost_per_kwh']:.2f}"
        summary_data[0]['Carbon (KRW/kWh)'] = f"{optimal_result['carbon_cost_per_kwh']:.2f}"

    # ESS comparison if available
    if 'results_ess' in st.session_state:
        optimal_ess_result = next(r for r in st.session_state.results_ess if r['ppa_percent'] == st.session_state.optimal_ess_ppa)
        savings = st.session_state.optimal_cost - st.session_state.optimal_ess_cost

        ess_row = {
            'Scenario': 'Optimal (With ESS)',
            'PPA Coverage (%)': f"{st.session_state.optimal_ess_ppa}%",
            'Total Cost (M KRW)': f"{st.session_state.optimal_ess_cost/1e6:.1f}",
            'Cost per kWh (KRW/kWh)': f"{(optimal_ess_result['total_cost_with_carbon_per_kwh'] if has_carbon_price else optimal_ess_result['total_cost_per_kwh']):.2f}",
            'Emissions (tCO2e)': f"{optimal_ess_result['total_emissions']/1000:.1f}",
            'Emissions (kgCO2e/kWh)': f"{optimal_ess_result['emissions_per_kwh']:.3f}"
        }

        if has_carbon_price:
            ess_row['Electricity Cost (M KRW)'] = f"{optimal_ess_result['total_cost']/1e6:.1f}"
            ess_row['Carbon Cost (M KRW)'] = f"{optimal_ess_result['carbon_cost']/1e6:.1f}"
            ess_row['Electricity (KRW/kWh)'] = f"{optimal_ess_result['total_cost_per_kwh']:.2f}"
            ess_row['Carbon (KRW/kWh)'] = f"{optimal_ess_result['carbon_cost_per_kwh']:.2f}"

        summary_data.append(ess_row)

        # Add savings row
        savings_row = {
            'Scenario': 'ESS Savings',
            'PPA Coverage (%)': '-',
            'Total Cost (M KRW)': f"{savings/1e6:.1f}",
            'Cost per kWh (KRW/kWh)': f"{(optimal_result['total_cost_per_kwh'] - optimal_ess_result['total_cost_per_kwh']):.2f}",
            'Emissions (tCO2e)': f"{(optimal_result['total_emissions'] - optimal_ess_result['total_emissions'])/1000:.1f}",
            'Emissions (kgCO2e/kWh)': f"{(optimal_result['emissions_per_kwh'] - optimal_ess_result['emissions_per_kwh']):.3f}"
        }

        if has_carbon_price:
            savings_row['Electricity Cost (M KRW)'] = f"{(optimal_result['total_cost'] - optimal_ess_result['total_cost'])/1e6:.1f}"
            savings_row['Carbon Cost (M KRW)'] = f"{(optimal_result['carbon_cost'] - optimal_ess_result['carbon_cost'])/1e6:.1f}"
            savings_row['Electricity (KRW/kWh)'] = f"{(optimal_result['total_cost_per_kwh'] - optimal_ess_result['total_cost_per_kwh']):.2f}"
            savings_row['Carbon (KRW/kWh)'] = f"{(optimal_result['carbon_cost_per_kwh'] - optimal_ess_result['carbon_cost_per_kwh']):.2f}"

        summary_data.append(savings_row)

    # Display as DataFrame
    summary_df = pd.DataFrame(summary_data)
    st.dataframe(summary_df, width='stretch', hide_index=True)

    st.divider()

    # Tabs for different visualizations
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📈 Cost Analysis",
        "💰 Cost Breakdown",
        "🔋 ESS Comparison",
        "📊 Data Tables",
        "💾 Export"
    ])

    with tab1:
        st.subheader("Cost & Emissions per kWh vs PPA Coverage")
        plot_cost_analysis(st.session_state.results_summary)

    with tab2:
        st.subheader("Cost Components Breakdown")
        plot_cost_breakdown(st.session_state.results_summary)

    with tab3:
        if 'results_ess' in st.session_state:
            st.subheader("ESS vs No ESS Comparison")
            plot_ess_comparison(
                st.session_state.results_summary,
                st.session_state.results_ess
            )
        else:
            st.info("ESS analysis was not run. Enable ESS in configuration to see comparison.")

    with tab4:
        st.subheader("Results Summary Table")
        display_results_table(st.session_state.results_summary)

        if 'annual_summary_df' in st.session_state:
            st.subheader("Annual Summary")
            st.dataframe(st.session_state.annual_summary_df, width="stretch")

        if 'peak_analysis' in st.session_state:
            st.subheader("Peak Hour Analysis")
            display_peak_analysis(st.session_state.peak_analysis)

    with tab5:
        st.subheader("Export Results")
        export_results()


def plot_cost_analysis(results_summary):
    """Plot cost per kWh and emissions per kWh vs PPA coverage with dual y-axis."""
    df = pd.DataFrame(results_summary)

    # Check if carbon pricing is enabled
    has_carbon = df['carbon_cost'].sum() > 0

    # Create figure with secondary y-axis
    fig = make_subplots(specs=[[{"secondary_y": True}]])

    # Add cost per kWh line (primary y-axis)
    # Use total cost including carbon when carbon pricing is enabled
    cost_column = 'total_cost_with_carbon_per_kwh' if has_carbon else 'total_cost_per_kwh'
    cost_label = 'Cost (incl. Carbon)' if has_carbon else 'Cost'

    fig.add_trace(
        go.Scatter(
            x=df['ppa_percent'],
            y=df[cost_column],
            mode='lines+markers',
            name=cost_label,
            line=dict(color='blue', width=3),
            marker=dict(size=8),
            hovertemplate='Cost: %{y:.1f} KRW/kWh<extra></extra>'
        ),
        secondary_y=False
    )

    # Add emissions per kWh line (secondary y-axis) - dashed line
    fig.add_trace(
        go.Scatter(
            x=df['ppa_percent'],
            y=df['emissions_per_kwh'],
            mode='lines+markers',
            name='Emissions',
            line=dict(color='green', width=3, dash='dash'),
            marker=dict(size=8, symbol='diamond'),
            hovertemplate='Emissions: %{y:.1f} kgCO2e/kWh<extra></extra>'
        ),
        secondary_y=True
    )

    # Mark optimal cost point (already determined which cost column to use above)
    optimal_cost_idx = df[cost_column].idxmin()
    optimal_cost_value = df.loc[optimal_cost_idx, cost_column]

    fig.add_trace(
        go.Scatter(
            x=[df.loc[optimal_cost_idx, 'ppa_percent']],
            y=[optimal_cost_value],
            mode='markers',
            name='Optimal Cost',
            marker=dict(size=15, color='red', symbol='star'),
            showlegend=True
        ),
        secondary_y=False
    )

    # Mark lowest emission point
    optimal_emission_idx = df['emissions_per_kwh'].idxmin()
    fig.add_trace(
        go.Scatter(
            x=[df.loc[optimal_emission_idx, 'ppa_percent']],
            y=[df.loc[optimal_emission_idx, 'emissions_per_kwh']],
            mode='markers',
            name='Lowest Emissions',
            marker=dict(size=15, color='darkgreen', symbol='star'),
            showlegend=True
        ),
        secondary_y=True
    )

    # Set axis titles
    fig.update_xaxes(title_text="PPA Coverage (%)")
    fig.update_yaxes(
        title_text="Cost (KRW/kWh)",
        secondary_y=False,
        title_font=dict(color='blue'),
        showgrid=True,
        gridcolor='lightblue'
    )
    fig.update_yaxes(
        title_text="Emissions (kgCO2e/kWh)",
        secondary_y=True,
        title_font=dict(color='green'),
        showgrid=True,
        gridcolor='lightgreen',
        griddash='dash'
    )

    fig.update_layout(
        hovermode='x',
        height=500,
        hoverlabel=dict(
            bgcolor="white",
            font_size=14,
            font_family="Arial",
            namelength=-1
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        )
    )

    render_plotly_chart(fig)


def plot_cost_breakdown(results_summary):
    """Plot stacked area chart of cost components."""
    df = pd.DataFrame(results_summary)

    fig = go.Figure()

    # Add cost components
    fig.add_trace(go.Scatter(
        x=df['ppa_percent'],
        y=df['ppa_cost'] / 1e6,
        mode='lines',
        name='PPA Cost',
        stackgroup='one',
        line=dict(width=0.5, color='green')
    ))

    fig.add_trace(go.Scatter(
        x=df['ppa_percent'],
        y=(df['grid_cost'] - df['grid_demand_cost']) / 1e6,
        mode='lines',
        name='Grid Energy Cost',
        stackgroup='one',
        line=dict(width=0.5, color='orange')
    ))

    fig.add_trace(go.Scatter(
        x=df['ppa_percent'],
        y=df['grid_demand_cost'] / 1e6,
        mode='lines',
        name='Contract Fee',
        stackgroup='one',
        line=dict(width=0.5, color='red')
    ))

    if df['ess_cost'].sum() > 0:
        fig.add_trace(go.Scatter(
            x=df['ppa_percent'],
            y=df['ess_cost'] / 1e6,
            mode='lines',
            name='ESS Cost',
            stackgroup='one',
            line=dict(width=0.5, color='purple')
        ))

    if df['carbon_cost'].sum() > 0:
        fig.add_trace(go.Scatter(
            x=df['ppa_percent'],
            y=df['carbon_cost'] / 1e6,
            mode='lines',
            name='Carbon Cost',
            stackgroup='one',
            line=dict(width=0.5, color='brown')
        ))

    fig.update_layout(
        xaxis_title="PPA Coverage (%)",
        yaxis_title="Cost (Million KRW)",
        hovermode='x unified',
        height=500
    )

    render_plotly_chart(fig)


def plot_ess_comparison(results_no_ess, results_ess):
    """Compare costs with and without ESS."""
    df_no_ess = pd.DataFrame(results_no_ess)
    df_ess = pd.DataFrame(results_ess)

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df_no_ess['ppa_percent'],
        y=df_no_ess['total_cost'] / 1e6,
        mode='lines+markers',
        name='No ESS',
        line=dict(color='blue', width=2)
    ))

    fig.add_trace(go.Scatter(
        x=df_ess['ppa_percent'],
        y=df_ess['total_cost'] / 1e6,
        mode='lines+markers',
        name='With ESS',
        line=dict(color='green', width=2)
    ))

    # Savings area
    fig.add_trace(go.Scatter(
        x=df_no_ess['ppa_percent'],
        y=df_no_ess['total_cost'] / 1e6,
        mode='lines',
        name='Savings',
        line=dict(width=0),
        showlegend=False
    ))

    fig.add_trace(go.Scatter(
        x=df_ess['ppa_percent'],
        y=df_ess['total_cost'] / 1e6,
        mode='lines',
        name='ESS Savings',
        fill='tonexty',
        line=dict(width=0),
        fillcolor='rgba(0,255,0,0.2)'
    ))

    fig.update_layout(
        xaxis_title="PPA Coverage (%)",
        yaxis_title="Total Annual Cost (Million KRW)",
        hovermode='x unified',
        height=500
    )

    render_plotly_chart(fig)


def display_results_table(results_summary):
    """Display results in a table."""
    df = pd.DataFrame(results_summary)

    # Check if carbon pricing is enabled
    has_carbon_price = df['carbon_cost'].sum() > 0

    # Format for display
    if has_carbon_price:
        display_df = pd.DataFrame({
            'PPA Coverage (%)': df['ppa_percent'],
            'Total Cost (M KRW)': (df['total_cost'] / 1e6).round(2),
            'Carbon Cost (M KRW)': (df['carbon_cost'] / 1e6).round(2),
            'Total with Carbon (M KRW)': (df['total_cost_with_carbon'] / 1e6).round(2),
            'PPA Cost (M KRW)': (df['ppa_cost'] / 1e6).round(2),
            'Grid Energy (M KRW)': ((df['grid_cost'] - df['grid_demand_cost']) / 1e6).round(2),
            'Contract Fee (M KRW)': (df['grid_demand_cost'] / 1e6).round(2),
            'Cost (KRW/kWh)': df['total_cost_per_kwh'].round(2),
            'Carbon (KRW/kWh)': df['carbon_cost_per_kwh'].round(2),
            'Total+Carbon (KRW/kWh)': df['total_cost_with_carbon_per_kwh'].round(2),
            'Emissions (tCO2e)': (df['total_emissions'] / 1000).round(2),
            'Emissions (kgCO2e/kWh)': df['emissions_per_kwh'].round(3)
        })
    else:
        display_df = pd.DataFrame({
            'PPA Coverage (%)': df['ppa_percent'],
            'Total Cost (M KRW)': (df['total_cost'] / 1e6).round(2),
            'PPA Cost (M KRW)': (df['ppa_cost'] / 1e6).round(2),
            'Grid Energy Cost (M KRW)': ((df['grid_cost'] - df['grid_demand_cost']) / 1e6).round(2),
            'Contract Fee (M KRW)': (df['grid_demand_cost'] / 1e6).round(2),
            'Total Cost (KRW/kWh)': df['total_cost_per_kwh'].round(2),
            'Emissions (tCO2e)': (df['total_emissions'] / 1000).round(2),
            'Emissions (kgCO2e/kWh)': df['emissions_per_kwh'].round(3)
        })

    st.dataframe(display_df, width="stretch")


def display_peak_analysis(peak_analysis):
    """Display peak hour analysis."""
    col1, col2 = st.columns(2)

    with col1:
        st.metric("Peak Avg Rate", f"{peak_analysis['peak_avg_rate']:.1f} KRW/kWh")
        st.write("Peak Hours:", ", ".join(map(str, peak_analysis['peak_hours'])))

    with col2:
        st.metric("Off-Peak Avg Rate", f"{peak_analysis['offpeak_avg_rate']:.1f} KRW/kWh")
        st.write("Off-Peak Hours:", ", ".join(map(str, peak_analysis['offpeak_hours'])))

    savings = peak_analysis['peak_avg_rate'] - st.session_state.config['ppa_price']
    st.metric("Peak Hour Savings with PPA", f"{savings:.1f} KRW/kWh")


def export_results():
    """Export results to Excel."""
    if 'long_df' not in st.session_state:
        st.warning("Detailed data not available. Enable 'Export Long Format' in configuration.")
        return

    st.markdown("**Download to your computer**")
    try:
        excel_bytes = export_to_excel_bytes(
            st.session_state.long_df,
            st.session_state.annual_summary_df,
            st.session_state.cost_comparison_df
        )
        st.download_button(
            label="📥 Download Excel",
            data=excel_bytes,
            file_name=st.session_state.config.get('output_file', 'ppa_analysis_results.xlsx'),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Failed to prepare download: {str(e)}")

    # Download buttons for individual tables
    st.subheader("Download Individual Tables")

    col1, col2, col3 = st.columns(3)

    with col1:
        if 'cost_comparison_df' in st.session_state:
            cost_csv_name = st.text_input(
                "Cost Analysis filename",
                value="cost_analysis.csv",
                key="cost_csv_name"
            )
            csv = st.session_state.cost_comparison_df.to_csv(index=False)
            st.download_button(
                "📥 Download Cost Analysis CSV",
                csv,
                cost_csv_name,
                "text/csv"
            )
            

    with col2:
        if 'annual_summary_df' in st.session_state:
            annual_csv_name = st.text_input(
                "Annual Summary filename",
                value="annual_summary.csv",
                key="annual_csv_name"
            )
            csv = st.session_state.annual_summary_df.to_csv(index=False)
            st.download_button(
                "📥 Download Annual Summary CSV",
                csv,
                annual_csv_name,
                "text/csv"
            )
            

    with col3:
        if 'long_df' in st.session_state:
            long_csv_name = st.text_input(
                "Full Data filename",
                value="full_data.csv",
                key="long_csv_name"
            )
            csv = st.session_state.long_df.to_csv(index=False)
            st.download_button(
                "📥 Download Full Data CSV",
                csv,
                long_csv_name,
                "text/csv"
            )
            


if __name__ == "__main__":
    main()
