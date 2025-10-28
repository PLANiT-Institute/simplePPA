# SimplePPA - Power Purchase Agreement Analysis Tool

A comprehensive tool for analyzing Power Purchase Agreement (PPA) scenarios with solar generation, grid electricity, and optional Energy Storage Systems (ESS).

## Overview

SimplePPA helps energy buyers evaluate the economic impact of different PPA coverage levels by simulating hourly energy flows and calculating total costs. The tool compares PPA solar generation against grid electricity across multiple scenarios to find the optimal balance.

## Features

- **Multi-scenario Analysis**: Evaluate PPA coverage at any range and step size (fully customizable)
- **Flexible PPA Contracts**: Support for minimum take requirements and resale options
- **ESS Integration**: Optional battery storage for excess solar energy
- **Detailed Cost Breakdown**: Separate tracking of PPA costs, grid energy costs, and demand charges
- **Three Interfaces**:
  - **Streamlit Web GUI** (`main_gui.py`) - Interactive web interface with visualizations
  - **Modular CLI** (`main_modular.py`) - Configuration-driven command line
  - **Legacy Scripts** (`main.py`, `gui_app.py`) - Original implementations
- **Modular Architecture**: Separate libraries for calculation, analysis, export, and configuration
- **Comprehensive Export**: Excel output with hourly data, annual summaries, and pivot-ready formats
- **Input Data Review**: Visual inspection of load, solar, and grid rate patterns before analysis
- **Interactive Tooltips**: Contextual help for all configuration parameters

## Installation

### Requirements

```bash
pip install -r requirements.txt
```

This installs:
- pandas (≥2.0.0) - Data processing
- numpy (≥1.24.0) - Numerical operations
- openpyxl (≥3.1.0) - Excel I/O
- matplotlib (≥3.7.0) - Static charts
- plotly (≥5.17.0) - Interactive visualizations
- streamlit (≥1.28.0) - Web GUI framework
- altair (≥5.0.0) - Additional charting

### Data Files

Place the following files in the `data/` directory:
- `pattern.xlsx`: Hourly load and solar generation patterns (normalized 0-1)
- `KEPCO.xlsx`: Grid electricity rates and tariff information

## Usage

### 🌐 Streamlit Web GUI (Recommended)

```bash
streamlit run main_gui.py
```

**Features:**
- **Interactive Configuration**: All parameters adjustable in sidebar with tooltips
- **Data Review**: Preview and visualize input data before analysis
  - Load and solar patterns with 7-day and daily average charts
  - Grid rate distribution and hourly patterns
  - Summary statistics and correlations
- **Real-time Analysis**: Click "Run Analysis" to compute all scenarios
- **Interactive Visualizations**:
  - Cost analysis with optimal point highlighting
  - Cost breakdown (stacked area chart)
  - ESS comparison (savings visualization)
  - Peak hour analysis
- **Flexible Export**: Download results as Excel or individual CSVs
- **Fully Customizable PPA Range**: Set any start/end/step (not limited to 0-200% in 10% steps)

**Key Configuration Options:**
- PPA Coverage Range: Start %, End %, Step % (e.g., 50-150% in 5% steps)
- PPA Price: Fixed rate in KRW/kWh
- Minimum Take: 0-100% (with optional purchase flexibility)
- Resell Options: Enable/disable with configurable resell rate
- ESS Parameters: Capacity and discharge price

### 🖥️ Modular Command Line

```bash
python main_modular.py
```

**Configuration Methods:**

1. **Edit in code** (`main_modular.py` lines 35-44):
```python
config = get_default_config()
config['ppa_range_start'] = 50
config['ppa_range_end'] = 150
config['ppa_range_step'] = 5
config['load_capacity_mw'] = 5000
```

2. **Use config file**:
```python
from libs.config import load_config_from_file
config = load_config_from_file('my_config.json')
```

3. **App settings** (see `app_settings.json`):
```json
{
  "load_capacity_mw": 3000,
  "ppa_price": 170,
  "ppa_range_start": 0,
  "ppa_range_end": 200,
  "ppa_range_step": 10
}
```

**Features:**
- Configuration validation with error messages
- Detailed console output with progress indicators
- Automatic Excel export with comprehensive data
- Peak hour analysis
- Supports ESS analysis

### 📊 Legacy Interfaces

**Original CLI** (`main.py`):
```bash
python main.py
```
Edit parameters at top of file. Outputs detailed Excel file and console statistics.

**Tkinter GUI** (`gui_app.py`):
```bash
python gui_app.py
```
Desktop GUI application with matplotlib charts.

## Algorithm Details

### Core Calculation Engine

The heart of SimplePPA is the `calculate_ppa_cost()` function, which simulates hourly energy procurement and cost accumulation over the analysis period.

#### Input Parameters

- **load_df**: Hourly load demand profile (normalized 0-1)
- **ppa_df**: Hourly solar generation profile (normalized 0-1)
- **grid_df**: Hourly grid electricity rates (KRW/kWh)
- **load_capacity_mw**: Peak load capacity in MW
- **ppa_coverage**: PPA sizing as fraction of peak load (e.g., 1.0 = 100%)
- **contract_fee**: Grid demand charge (KRW/kW)
- **ess_capacity**: ESS capacity in kWh (0 = no ESS)

#### Output

Returns five cost components:
1. **total_cost**: Total annual cost
2. **ppa_cost**: Cost from PPA purchases (minus resale revenue if applicable)
3. **grid_cost**: Grid energy cost + demand charges
4. **grid_demand_cost**: Grid demand charges only
5. **ess_cost**: Cost for ESS discharge energy

### Hour-by-Hour Energy Flow

For each hour of the year (8,760 hours), the algorithm processes energy flows in this sequence:

#### Step 1: Scale Normalized Patterns to Actual Capacity

```python
load_mw = load_normalized * load_capacity_mw
ppa_generation_mw = solar_normalized * load_capacity_mw * ppa_coverage
```

Convert to kWh (since each hour represents 1 hour of operation):
```python
load_kwh = load_mw * 1000
ppa_generation_kwh = ppa_generation_mw * 1000
```

**Example**: If `load_capacity_mw = 100` and `ppa_coverage = 1.2`:
- When `load_normalized = 0.8` → `load_kwh = 80,000 kWh`
- When `solar_normalized = 0.5` → `ppa_generation_kwh = 60,000 kWh`

#### Step 2: PPA Purchase Decision

The PPA contract has two components:

**A. Mandatory Purchase (Minimum Take)**
```python
mandatory_ppa = ppa_generation_kwh * ppa_mintake
```

This amount **must** be purchased regardless of need.

**Example**: If `ppa_mintake = 0.8` (80%) and `ppa_generation_kwh = 60,000`:
- `mandatory_ppa = 48,000 kWh` (must buy)
- `optional_ppa_available = 12,000 kWh` (can buy if desired)

**B. Optional Purchase (if mintake < 100%)**

The buyer can choose to purchase additional PPA energy if it's economically beneficial:

```python
if ppa_price < hour_grid_rate and remaining_load > 0:
    optional_ppa_purchased = min(optional_ppa_available, remaining_load)
```

**Decision Logic**:
- Only buy optional PPA if it's **cheaper than grid** at this hour
- Only buy what's **needed** to meet remaining load
- Don't buy more than what's **available**

**Example**: Continuing from above with `load_kwh = 80,000`:
- After mandatory: `remaining_load = 80,000 - 48,000 = 32,000 kWh`
- If PPA is cheaper: buy up to `min(12,000, 32,000) = 12,000 kWh`
- Total PPA purchased: `48,000 + 12,000 = 60,000 kWh`

**Total PPA Purchase & Cost**:
```python
total_ppa_this_hour = mandatory_ppa + optional_ppa_purchased
ppa_cost += total_ppa_this_hour * ppa_price
```

#### Step 3: Energy Balance Check

Calculate remaining energy after PPA purchase:
```python
remaining_load = load_kwh - total_ppa_this_hour
```

Two scenarios emerge:

##### Scenario A: Excess PPA Energy (remaining_load ≤ 0)

We have more PPA energy than needed. Handle the excess:

```python
excess_ppa = abs(remaining_load)
```

**3a. Try to Store in ESS**
```python
if ess_storage < max_ess_capacity:
    ess_charge = min(excess_ppa, max_ess_capacity - ess_storage)
    ess_storage += ess_charge
    excess_ppa -= ess_charge
```

**Note**: ESS charging has no additional cost (already paid for PPA energy)

**3b. Try to Resell Remaining Excess**
```python
if excess_ppa > 0 and ppa_resell:
    resell_revenue = excess_ppa * ppa_price * ppa_resellrate
    ppa_cost -= resell_revenue  # Credit against PPA cost
```

**Example**: If `ppa_resellrate = 0.9` and excess = 10,000 kWh at 170 KRW/kWh:
- Revenue = `10,000 * 170 * 0.9 = 1,530,000 KRW`
- This reduces the net PPA cost

**3c. Waste Any Remaining Excess**

If ESS is full and reselling is disabled, excess energy is wasted (but already paid for in mandatory purchase).

##### Scenario B: Energy Deficit (remaining_load > 0)

Need more energy beyond PPA. Fulfill in order:

**3d. Discharge from ESS**
```python
if ess_storage > 0:
    ess_discharge = min(remaining_load, ess_storage)
    ess_storage -= ess_discharge
    ess_cost += ess_discharge * ppa_price * ess_price
    remaining_load -= ess_discharge
```

**ESS Discharge Cost**: Energy from ESS costs `ppa_price * ess_price`. If `ess_price = 0.5`, ESS energy costs 50% of PPA price.

**Example**: Discharging 5,000 kWh with `ppa_price = 170` and `ess_price = 0.5`:
- Cost = `5,000 * 170 * 0.5 = 425,000 KRW`

**3e. Purchase Remaining from Grid**
```python
if remaining_load > 0:
    grid_energy_cost += remaining_load * hour_grid_rate
    grid_demand_kw = remaining_load
    peak_grid_demand_kw = max(peak_grid_demand_kw, grid_demand_kw)
```

**Grid Demand Tracking**: Track peak grid demand for demand charge calculation.

#### Step 4: Calculate Grid Demand Charge

After processing all hours, apply demand charge based on peak grid usage:

```python
grid_demand_cost = peak_grid_demand_kw * contract_fee
```

**Example**: If peak grid demand was 45,000 kW and `contract_fee = 8,000 KRW/kW`:
- Demand charge = `45,000 * 8,000 = 360,000,000 KRW`

#### Step 5: Calculate Total Costs

```python
grid_total_cost = grid_energy_cost + grid_demand_cost
total_cost = ppa_cost + grid_total_cost + ess_cost
```

### Multi-Scenario Analysis

The tool runs the above calculation for 21 different PPA coverage levels (0%, 10%, 20%, ..., 200%) to create a cost curve.

**PPA Coverage Interpretation**:
- **0%**: No PPA, 100% grid electricity
- **50%**: PPA solar farm capacity is 50% of peak load capacity
- **100%**: PPA capacity equals peak load capacity
- **150%**: PPA capacity is 1.5x peak load capacity
- **200%**: PPA capacity is 2x peak load capacity

Higher coverage means more solar generation but also potentially more excess energy that must be managed.

### Optimization

The algorithm finds the optimal PPA coverage by comparing total annual costs:

```python
optimal_idx = min(range(len(results)), key=lambda i: results[i]['total_cost'])
```

This identifies the PPA coverage percentage that minimizes total electricity costs.

## Key Calculation Example

**Scenario Setup**:
- Load capacity: 100 MW
- PPA coverage: 120% (120 MW solar farm)
- PPA price: 170 KRW/kWh
- Grid rate this hour: 220 KRW/kWh
- Minimum take: 100%
- Resell allowed: No
- ESS: 10,000 kWh capacity, currently empty

**Hour Conditions**:
- Load: 0.7 normalized → 70,000 kWh
- Solar: 0.8 normalized → 96,000 kWh (120 MW * 0.8 * 1000)

**Calculation Flow**:

1. **Mandatory PPA**: 96,000 kWh * 1.0 = 96,000 kWh
   - Cost: 96,000 * 170 = 16,320,000 KRW

2. **Energy Balance**: 70,000 - 96,000 = -26,000 kWh (excess)

3. **Store in ESS**: 10,000 kWh → ESS now at 10,000 kWh

4. **Remaining Excess**: 16,000 kWh (wasted, no resale)

5. **Grid Purchase**: 0 kWh

6. **Total Hour Cost**: 16,320,000 KRW

**Next Hour Conditions**:
- Load: 0.9 normalized → 90,000 kWh
- Solar: 0.3 normalized → 36,000 kWh
- Grid rate: 250 KRW/kWh
- ESS: 10,000 kWh stored

**Calculation Flow**:

1. **Mandatory PPA**: 36,000 kWh
   - Cost: 36,000 * 170 = 6,120,000 KRW

2. **Energy Balance**: 90,000 - 36,000 = 54,000 kWh (deficit)

3. **Discharge ESS**: 10,000 kWh (empties ESS)
   - Cost: 10,000 * 170 * 0.5 = 850,000 KRW
   - Remaining deficit: 44,000 kWh

4. **Buy from Grid**: 44,000 kWh
   - Cost: 44,000 * 250 = 11,000,000 KRW
   - Track peak demand: 44,000 kW

5. **Total Hour Cost**: 6,120,000 + 850,000 + 11,000,000 = 17,970,000 KRW

## Output Files

### main.py Output

Running `main.py` generates `ppa_analysis_results.xlsx` with three sheets:

#### 1. PPA_Analysis_Data
Long-format data with one row per hour per scenario (~184,000 rows for full year):
- Datetime and time components
- PPA scenario identifier
- Energy values (load, generation, purchases, excess)
- Cost values (PPA, grid, resell revenue, total)
- Calculated metrics (coverage %, cost per kWh)

**Use Case**: Create pivot tables in Excel to analyze patterns by hour, month, season, etc.

#### 2. Annual_Summary
One row per PPA scenario (21 rows):
- Annual totals for energy and costs
- Load coverage percentages
- Per-kWh cost breakdowns

**Use Case**: Quick comparison of scenarios, find optimal coverage.

#### 3. Cost_Analysis
Cost breakdown by component for each scenario:
- Total cost
- PPA cost
- Grid energy cost
- Grid demand charge
- ESS cost
- Per-kWh metrics for each component

**Use Case**: Understand cost drivers, compare grid vs PPA economics.

### Console Output

`main.py` prints:
- Per-iteration statistics (PPA stats, Grid stats)
- Cost tables for each scenario
- Optimal PPA coverage results
- Peak hour analysis
- Data filtering and export confirmation

## Algorithm Features

### Minimum Take Contracts

Many PPAs require buyers to purchase a minimum percentage of generation:

- **100% minimum take**: Buy all generation, every hour (typical for solar PPAs)
- **80% minimum take**: Must buy 80%, can buy up to 100% if profitable
- **50% minimum take**: More flexibility, buy 50-100% based on economics

Lower minimum take provides flexibility but may result in higher PPA prices.

### Resale Mechanics

If enabled (`ppa_resell = True`):
- Excess PPA energy can be resold to grid
- Resale price is typically discounted: `resell_price = ppa_price * ppa_resellrate`
- Revenue reduces net PPA cost
- Useful when PPA generation exceeds load + ESS capacity

### ESS Economics

ESS provides value by:
1. **Peak Shaving**: Store cheap PPA energy, discharge during expensive grid hours
2. **Excess Management**: Store excess PPA energy instead of wasting it
3. **Demand Charge Reduction**: Lower peak grid demand by using stored energy

ESS cost structure:
- **Capital cost**: Not modeled (assumed externally)
- **Discharge cost**: `ess_price * ppa_price` per kWh (operating cost)

### Grid Demand Charges

Korean electricity tariffs (and many worldwide) include demand charges:
- **Energy charge**: Pay for kWh consumed
- **Demand charge**: Pay for peak kW demand (highest 15-min interval)

SimplePPA tracks peak hourly grid demand and applies contract fees. This incentivizes:
- Using PPA during high-demand hours
- Using ESS to shave peaks
- Maintaining consistent grid draw

## Performance Considerations

### main.py Performance

`main.py` is slower because it:
1. Prints statistics for every iteration (42 print operations)
2. Generates long-format DataFrame with ~184,000 rows
3. Creates comprehensive Excel file with multiple sheets
4. Performs additional peak/off-peak analysis

**Typical runtime**: 30-60 seconds

### gui_app.py Performance

`gui_app.py` is faster because it:
1. No console output during calculation
2. Only stores 21 summary results
3. Optional export (user-initiated)
4. No additional analysis

**Typical runtime**: 5-10 seconds

Both use identical calculation algorithms, so results are numerically equivalent.

## Use Cases

### 1. PPA Contract Evaluation
*Should we sign a PPA for a 150 MW solar farm?*

Run analysis with:
- `load_capacity_mw` = your peak load
- `ppa_price` = offered PPA price
- `ppa_coverage` = proposed solar capacity / peak load

Compare total costs at this coverage vs. grid-only (0% coverage).

### 2. Optimal Sizing
*What size solar farm minimizes our electricity costs?*

The tool automatically finds optimal coverage by comparing all scenarios. Check the "Optimal PPA coverage" result.

### 3. ESS Investment Analysis
*Should we add battery storage?*

Run twice:
- Once with `ess_include = False`
- Once with `ess_include = True` and desired capacity

Compare optimal costs. The difference is the maximum ESS investment that makes economic sense (before considering ESS capital costs).

### 4. Contract Terms Negotiation
*How much is resale worth? What about minimum take?*

Adjust `ppa_resell`, `ppa_resellrate`, and `ppa_mintake` to see impact on total costs. This helps evaluate different contract structures.

### 5. Seasonal Analysis
*How does PPA perform in summer vs. winter?*

Use the long-format Excel output to create pivot tables grouped by month or season.

## Limitations

1. **Hourly Resolution**: Uses hourly averages, not 15-minute intervals (actual demand charge basis)
2. **Perfect Forecast**: Assumes perfect knowledge of future generation and load
3. **No Uncertainty**: Deterministic model, no weather variability or forecast errors
4. **No Degradation**: Solar and ESS performance assumed constant over time
5. **No ESS Capital Cost**: Only operational costs modeled
6. **Single Grid Tariff**: Uses one tariff type, real contracts may have options
7. **No Transmission Losses**: Assumes perfect energy delivery
8. **No Grid Limitations**: Assumes unlimited grid capacity

## Advanced Customization

### Custom Load Patterns

Edit `data/pattern.xlsx`:
- Column 'load': Normalized hourly load (0-1, where 1 = peak)
- Column 'solar': Normalized hourly solar generation (0-1, where 1 = peak)
- Must have 8,760 rows (one year, hourly)

### Custom Grid Rates

Edit `data/KEPCO.xlsx` or modify `libs/KEPCOutils.py` to load your utility's tariff structure.

### Different Time Periods

In `main.py`, adjust:
```python
start_date = '2024-01-01'
end_date = '2024-12-31'
```

Can analyze shorter periods (e.g., summer only) or specific years.

### ESS Capacity Strategies

Choose ESS sizing:
- **Hours-based**: `ess_capacity = hours * peak_solar_mw * 1000`
- **Percentage-based**: `ess_capacity = pct * peak_solar_mw * 1000`
- **Absolute**: `ess_capacity = 50000  # 50 MWh`

## Technical Details

### Dependencies
- **pandas**: Data manipulation and analysis
- **numpy**: Numerical operations
- **openpyxl**: Excel file I/O
- **matplotlib**: Visualization (GUI only)
- **tkinter**: GUI framework

### Project Structure
```
simplePPA/
├── main_gui.py              # Streamlit web GUI (recommended)
├── main_modular.py          # Modular CLI with config support
├── main.py                  # Legacy CLI script
├── gui_app.py               # Legacy Tkinter GUI
├── app_settings.json        # App-level settings (limits, defaults)
├── requirements.txt         # Python dependencies
├── libs/                    # Modular library components
│   ├── calculator.py        # Core calculation engine
│   ├── data_processor.py    # Data loading and transformation
│   ├── analyzer.py          # Analysis and optimization logic
│   ├── exporter.py          # Excel export and printing
│   ├── config.py            # Configuration management
│   └── KEPCOutils.py        # Korean grid tariff processing
├── data/                    # Input data files
│   ├── pattern.xlsx         # Load and solar patterns
│   └── KEPCO.xlsx           # Grid rate data
└── README.md                # This file
```

### Code Architecture

**Modular Design:**
- **libs/calculator.py**: Pure calculation logic (no I/O)
  - `calculate_ppa_cost()`: Hour-by-hour simulation engine
  - All parameters passed as arguments (no globals)

- **libs/data_processor.py**: Data manipulation
  - `load_pattern_data()`: Load Excel patterns
  - `generate_scenario_columns()`: Create hourly scenario data
  - `create_long_format_dataframe()`: Pivot-ready output

- **libs/analyzer.py**: Business logic
  - `run_scenario_analysis()`: Multi-scenario execution
  - `run_ess_analysis()`: ESS-enabled analysis
  - `find_optimal_scenario()`: Optimization

- **libs/exporter.py**: Output formatting
  - `export_to_excel()`: Multi-sheet Excel files
  - Console printing utilities

- **libs/config.py**: Configuration management
  - `get_default_config()`: Default parameters
  - `validate_config()`: Parameter validation
  - `load_config_from_file()`: JSON/YAML support

**Interface Implementations:**
- `main_gui.py`: Uses all libs, adds Streamlit UI
- `main_modular.py`: Uses all libs, CLI workflow
- `main.py`, `gui_app.py`: Legacy monolithic implementations

## License

MIT License - Feel free to use and modify for your analysis needs.

## Contributing

Contributions welcome! Areas for improvement:
- Sub-hourly resolution (15-minute intervals)
- Stochastic weather modeling
- Multiple grid tariff options
- ESS degradation modeling
- Multi-year financial modeling with NPV
- Uncertainty quantification
- Additional renewable sources (wind, etc.)

## Support

For issues or questions:
1. Check the console output for error messages
2. Verify data file formats match expected structure
3. Ensure all parameters are within reasonable ranges
4. Review the algorithm details above to understand calculation logic

## Changelog

### v0.2 (Current)
- **NEW**: Streamlit web GUI with interactive visualizations
- **NEW**: Modular architecture with separate library components
- **NEW**: Fully customizable PPA coverage range (any start/end/step)
- **NEW**: Input data review with charts and statistics
- **NEW**: Interactive tooltips for all parameters
- **NEW**: Configuration file support (JSON/YAML)
- **NEW**: Clean requirements.txt for easy installation
- Improved: Better code organization and reusability
- Improved: Performance and maintainability

### v0.1
- Initial release
- Hourly PPA analysis with fixed 0-200% in 10% steps
- ESS integration
- Resale mechanics
- Original CLI (`main.py`) and Tkinter GUI (`gui_app.py`)
- Comprehensive Excel export