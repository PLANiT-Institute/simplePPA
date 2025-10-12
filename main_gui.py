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
from libs.config import get_default_config, validate_config
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
from libs.exporter import export_to_excel


# Page configuration
st.set_page_config(
    page_title="SimplePPA Analysis",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize session state
if 'analysis_done' not in st.session_state:
    st.session_state.analysis_done = False
if 'results_summary' not in st.session_state:
    st.session_state.results_summary = None
if 'config' not in st.session_state:
    st.session_state.config = get_default_config()


def main():
    st.title("⚡ SimplePPA - Power Purchase Agreement Analysis")
    st.markdown("Analyze PPA scenarios with solar generation, grid electricity, and optional ESS")

    # Sidebar for configuration
    with st.sidebar:
        st.header("⚙️ Configuration")

        # Data Files
        st.subheader("📁 Data Files")
        pattern_file = st.text_input("Pattern File", value="data/pattern.xlsx")
        kepco_file = st.text_input("KEPCO File", value="data/KEPCO.xlsx")
        kepco_year = st.number_input("KEPCO Year", value=2024, step=1)
        kepco_tariff = st.selectbox("KEPCO Tariff",
                                     ["HV_C_III", "HV_C_I", "HV_C_II"],
                                     index=0)

        st.divider()

        # Analysis Period
        st.subheader("📅 Analysis Period")
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("Start Date",
                                      value=pd.to_datetime("2024-01-01"))
        with col2:
            end_date = st.date_input("End Date",
                                    value=pd.to_datetime("2024-12-31"))

        st.divider()

        # Load Parameters
        st.subheader("⚡ Load Parameters")
        load_capacity_mw = st.number_input("Load Capacity (MW)",
                                          value=3000.0,
                                          min_value=0.1,
                                          step=100.0)

        st.divider()

        # PPA Parameters
        st.subheader("🌞 PPA Parameters")
        ppa_price = st.number_input("PPA Price (KRW/kWh)",
                                   value=170.0,
                                   min_value=0.0,
                                   step=1.0)
        ppa_mintake = st.slider("Minimum Take (%)",
                               min_value=0,
                               max_value=100,
                               value=100,
                               step=1) / 100.0
        ppa_resell = st.checkbox("Allow Reselling", value=False)
        ppa_resellrate = st.slider("Resell Rate (%)",
                                   min_value=0,
                                   max_value=100,
                                   value=90,
                                   step=1,
                                   disabled=not ppa_resell) / 100.0

        st.divider()

        # PPA Coverage Range
        st.subheader("📊 PPA Coverage Range")
        col1, col2, col3 = st.columns(3)
        with col1:
            ppa_range_start = st.number_input("Start (%)",
                                             value=0,
                                             min_value=0,
                                             step=10)
        with col2:
            ppa_range_end = st.number_input("End (%)",
                                           value=200,
                                           min_value=0,
                                           step=10)
        with col3:
            ppa_range_step = st.number_input("Step (%)",
                                            value=10,
                                            min_value=1,
                                            step=1)

        num_scenarios = len(range(ppa_range_start, ppa_range_end + 1, ppa_range_step))
        st.info(f"📈 {num_scenarios} scenarios will be analyzed")

        st.divider()

        # ESS Parameters
        st.subheader("🔋 ESS Parameters")
        ess_include = st.checkbox("Include ESS Analysis", value=False)
        ess_capacity = st.slider("ESS Capacity (% of solar peak)",
                                min_value=0,
                                max_value=200,
                                value=50,
                                step=10,
                                disabled=not ess_include) / 100.0
        ess_price = st.slider("ESS Discharge Price (% of PPA price)",
                             min_value=0,
                             max_value=100,
                             value=50,
                             step=5,
                             disabled=not ess_include) / 100.0

        st.divider()

        # Output Options
        st.subheader("💾 Output Options")
        output_file = st.text_input("Output Filename",
                                    value="ppa_analysis_results.xlsx")
        export_long_format = st.checkbox("Export Long Format", value=True)
        verbose = st.checkbox("Verbose Output", value=False)

        st.divider()

        # Run Analysis Button
        run_button = st.button("🚀 Run Analysis", type="primary", use_container_width=True)

    # Main content area
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
            'ess_include': ess_include,
            'ess_capacity': ess_capacity,
            'ess_price': ess_price,
            'output_file': output_file,
            'verbose': verbose,
            'export_long_format': export_long_format
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
                load_df, solar_df, ppa_df = load_pattern_data(config['pattern_file'])
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
                    load_df, ppa_df, grid_df, contract_fee, config,
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
                        load_df, ppa_df, grid_df, contract_fee, config,
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
                        grid_df, load_df, ppa_df,
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


def display_results():
    """Display analysis results with interactive charts."""

    st.header("📊 Analysis Results")

    # Summary metrics
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            "Optimal PPA Coverage",
            f"{st.session_state.optimal_ppa}%"
        )

    with col2:
        st.metric(
            "Optimal Cost",
            f"{st.session_state.optimal_cost/1e6:.1f}M KRW"
        )

    with col3:
        if 'results_ess' in st.session_state:
            st.metric(
                "Optimal with ESS",
                f"{st.session_state.optimal_ess_ppa}%"
            )
        else:
            st.metric("ESS Analysis", "Not Run")

    with col4:
        if 'results_ess' in st.session_state:
            savings = st.session_state.optimal_cost - st.session_state.optimal_ess_cost
            st.metric(
                "ESS Savings",
                f"{savings/1e6:.1f}M KRW"
            )
        else:
            st.metric("ESS Savings", "N/A")

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
        st.subheader("Total Cost vs PPA Coverage")
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
            st.dataframe(st.session_state.annual_summary_df, use_container_width=True)

        if 'peak_analysis' in st.session_state:
            st.subheader("Peak Hour Analysis")
            display_peak_analysis(st.session_state.peak_analysis)

    with tab5:
        st.subheader("Export Results")
        export_results()


def plot_cost_analysis(results_summary):
    """Plot total cost vs PPA coverage."""
    df = pd.DataFrame(results_summary)

    fig = go.Figure()

    # Add cost line
    fig.add_trace(go.Scatter(
        x=df['ppa_percent'],
        y=df['total_cost'] / 1e6,
        mode='lines+markers',
        name='Total Cost',
        line=dict(color='blue', width=3),
        marker=dict(size=8)
    ))

    # Mark optimal point
    optimal_idx = df['total_cost'].idxmin()
    fig.add_trace(go.Scatter(
        x=[df.loc[optimal_idx, 'ppa_percent']],
        y=[df.loc[optimal_idx, 'total_cost'] / 1e6],
        mode='markers',
        name='Optimal',
        marker=dict(size=15, color='red', symbol='star')
    ))

    fig.update_layout(
        xaxis_title="PPA Coverage (%)",
        yaxis_title="Total Annual Cost (Million KRW)",
        hovermode='x unified',
        height=500
    )

    st.plotly_chart(fig, use_container_width=True)


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
        name='Grid Demand Charge',
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

    fig.update_layout(
        xaxis_title="PPA Coverage (%)",
        yaxis_title="Cost (Million KRW)",
        hovermode='x unified',
        height=500
    )

    st.plotly_chart(fig, use_container_width=True)


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

    st.plotly_chart(fig, use_container_width=True)


def display_results_table(results_summary):
    """Display results in a table."""
    df = pd.DataFrame(results_summary)

    # Format for display
    display_df = pd.DataFrame({
        'PPA Coverage (%)': df['ppa_percent'],
        'Total Cost (M KRW)': (df['total_cost'] / 1e6).round(2),
        'PPA Cost (M KRW)': (df['ppa_cost'] / 1e6).round(2),
        'Grid Cost (M KRW)': (df['grid_cost'] / 1e6).round(2),
        'Demand Charge (M KRW)': (df['grid_demand_cost'] / 1e6).round(2),
        'Cost per kWh (KRW)': df['total_cost_per_kwh'].round(2)
    })

    st.dataframe(display_df, use_container_width=True)


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

    st.write("Click the button below to export results to Excel file.")

    if st.button("💾 Export to Excel"):
        try:
            with st.spinner("Exporting..."):
                export_to_excel(
                    st.session_state.config['output_file'],
                    st.session_state.long_df,
                    st.session_state.annual_summary_df,
                    st.session_state.cost_comparison_df
                )
            st.success(f"✅ Results exported to {st.session_state.config['output_file']}")
        except Exception as e:
            st.error(f"❌ Export failed: {str(e)}")

    # Download buttons for individual tables
    st.subheader("Download Individual Tables")

    col1, col2, col3 = st.columns(3)

    with col1:
        if 'cost_comparison_df' in st.session_state:
            csv = st.session_state.cost_comparison_df.to_csv(index=False)
            st.download_button(
                "📥 Cost Analysis CSV",
                csv,
                "cost_analysis.csv",
                "text/csv"
            )

    with col2:
        if 'annual_summary_df' in st.session_state:
            csv = st.session_state.annual_summary_df.to_csv(index=False)
            st.download_button(
                "📥 Annual Summary CSV",
                csv,
                "annual_summary.csv",
                "text/csv"
            )

    with col3:
        if 'long_df' in st.session_state:
            csv = st.session_state.long_df.to_csv(index=False)
            st.download_button(
                "📥 Full Data CSV",
                csv,
                "full_data.csv",
                "text/csv"
            )


if __name__ == "__main__":
    main()
