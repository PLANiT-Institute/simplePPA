import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import libs.KEPCOutils as _kepco

class PPAAnalysisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PPA Analysis Tool")
        self.root.geometry("1200x800")
        
        # Initialize data
        self.pattern_df = None
        self.grid_df = None
        self.contract_fee = None
        self.results_data = []
        
        # Create main notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create tabs
        self.create_input_tab()
        self.create_results_tab()
        
        # Load default data
        self.load_default_data()
        
    def create_input_tab(self):
        """Create the input parameters tab"""
        input_frame = ttk.Frame(self.notebook)
        self.notebook.add(input_frame, text="Parameters")
        
        # Main container with scrollbar
        canvas = tk.Canvas(input_frame)
        scrollbar = ttk.Scrollbar(input_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Data files section
        file_frame = ttk.LabelFrame(scrollable_frame, text="Data Files", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Label(file_frame, text="Pattern File:").grid(row=0, column=0, sticky="w")
        self.pattern_file_var = tk.StringVar(value="data/pattern.xlsx")
        ttk.Entry(file_frame, textvariable=self.pattern_file_var, width=40).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_pattern_file).grid(row=0, column=2)
        
        ttk.Label(file_frame, text="KEPCO File:").grid(row=1, column=0, sticky="w")
        self.kepco_file_var = tk.StringVar(value="data/KEPCO.xlsx")
        ttk.Entry(file_frame, textvariable=self.kepco_file_var, width=40).grid(row=1, column=1, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_kepco_file).grid(row=1, column=2)
        
        # Analysis parameters
        analysis_frame = ttk.LabelFrame(scrollable_frame, text="Analysis Parameters", padding="10")
        analysis_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Label(analysis_frame, text="Start Date (YYYY-MM-DD):").grid(row=0, column=0, sticky="w")
        self.start_date_var = tk.StringVar(value="2024-01-01")
        ttk.Entry(analysis_frame, textvariable=self.start_date_var).grid(row=0, column=1, padx=5, sticky="w")
        
        ttk.Label(analysis_frame, text="End Date (YYYY-MM-DD):").grid(row=1, column=0, sticky="w")
        self.end_date_var = tk.StringVar(value="2024-12-31")
        ttk.Entry(analysis_frame, textvariable=self.end_date_var).grid(row=1, column=1, padx=5, sticky="w")
        
        ttk.Label(analysis_frame, text="Load Capacity (MW):").grid(row=2, column=0, sticky="w")
        self.load_capacity_var = tk.DoubleVar(value=3000.0)
        ttk.Entry(analysis_frame, textvariable=self.load_capacity_var).grid(row=2, column=1, padx=5, sticky="w")
        
        # PPA parameters
        ppa_frame = ttk.LabelFrame(scrollable_frame, text="PPA Parameters", padding="10")
        ppa_frame.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Label(ppa_frame, text="PPA Price (KRW/kWh):").grid(row=0, column=0, sticky="w")
        self.ppa_price_var = tk.DoubleVar(value=170.0)
        ttk.Entry(ppa_frame, textvariable=self.ppa_price_var).grid(row=0, column=1, padx=5, sticky="w")
        
        ttk.Label(ppa_frame, text="Minimum Take (%):").grid(row=1, column=0, sticky="w")
        self.ppa_mintake_var = tk.DoubleVar(value=100.0)
        ttk.Entry(ppa_frame, textvariable=self.ppa_mintake_var).grid(row=1, column=1, padx=5, sticky="w")
        
        ttk.Label(ppa_frame, text="Resell Allowed:").grid(row=2, column=0, sticky="w")
        self.ppa_resell_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(ppa_frame, variable=self.ppa_resell_var).grid(row=2, column=1, padx=5, sticky="w")
        
        ttk.Label(ppa_frame, text="Resell Rate:").grid(row=3, column=0, sticky="w")
        self.ppa_resellrate_var = tk.DoubleVar(value=0.9)
        ttk.Entry(ppa_frame, textvariable=self.ppa_resellrate_var).grid(row=3, column=1, padx=5, sticky="w")
        
        # ESS parameters
        ess_frame = ttk.LabelFrame(scrollable_frame, text="ESS Parameters", padding="10")
        ess_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Label(ess_frame, text="Include ESS:").grid(row=0, column=0, sticky="w")
        self.ess_include_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(ess_frame, variable=self.ess_include_var).grid(row=0, column=1, padx=5, sticky="w")
        
        ttk.Label(ess_frame, text="ESS Capacity (% of solar peak):").grid(row=1, column=0, sticky="w")
        self.ess_capacity_var = tk.DoubleVar(value=50.0)
        ttk.Entry(ess_frame, textvariable=self.ess_capacity_var).grid(row=1, column=1, padx=5, sticky="w")
        
        ttk.Label(ess_frame, text="ESS Price (ratio to PPA):").grid(row=2, column=0, sticky="w")
        self.ess_price_var = tk.DoubleVar(value=0.5)
        ttk.Entry(ess_frame, textvariable=self.ess_price_var).grid(row=2, column=1, padx=5, sticky="w")
        
        # Analysis buttons
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Run Analysis", command=self.run_analysis, style="Accent.TButton").pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Export Results", command=self.export_results).pack(side=tk.LEFT, padx=10)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def create_results_tab(self):
        """Create the results visualization tab"""
        results_frame = ttk.Frame(self.notebook)
        self.notebook.add(results_frame, text="Results")
        
        # Create matplotlib figure
        self.fig = Figure(figsize=(12, 8), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.fig, results_frame)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Chart control frame
        control_frame = ttk.Frame(results_frame)
        control_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(control_frame, text="Chart Type:").pack(side=tk.LEFT)
        self.chart_type_var = tk.StringVar(value="Cost Analysis")
        chart_combo = ttk.Combobox(control_frame, textvariable=self.chart_type_var, 
                                 values=["Cost Analysis", "PPA Coverage", "Monthly Patterns", "Hourly Patterns"])
        chart_combo.pack(side=tk.LEFT, padx=5)
        chart_combo.bind("<<ComboboxSelected>>", self.update_charts)
        
        ttk.Button(control_frame, text="Update Charts", command=self.update_charts).pack(side=tk.LEFT, padx=10)
        
    def load_default_data(self):
        """Load default data files"""
        try:
            self.pattern_df = pd.read_excel(self.pattern_file_var.get(), index_col=0)
            self.grid_df, self.contract_fee = _kepco.process_kepco_data(
                self.kepco_file_var.get(), 2024, "HV_C_III"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load default data: {str(e)}")
    
    def browse_pattern_file(self):
        """Browse for pattern file"""
        filename = filedialog.askopenfilename(
            title="Select Pattern File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.pattern_file_var.set(filename)
            
    def browse_kepco_file(self):
        """Browse for KEPCO file"""
        filename = filedialog.askopenfilename(
            title="Select KEPCO File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.kepco_file_var.set(filename)
    
    def calculate_ppa_cost(self, load_df, ppa_df, grid_df, load_capacity_mw, ppa_coverage, contract_fee, ess_capacity=0):
        """Calculate PPA costs - same as original function"""
        load = load_df['load']
        grid_rate = grid_df['rate']
        solar_generation = ppa_df['generation']
        
        # Convert parameters
        ppa_price = self.ppa_price_var.get()
        ppa_mintake = self.ppa_mintake_var.get() / 100.0
        ppa_resell = self.ppa_resell_var.get()
        ppa_resellrate = self.ppa_resellrate_var.get()
        ess_price = self.ess_price_var.get()
        
        load_mw = load * load_capacity_mw
        ppa_generation_mw = solar_generation * load_capacity_mw * ppa_coverage
        load_kwh = load_mw * 1000
        ppa_generation_kwh = ppa_generation_mw * 1000
        
        ppa_cost = 0
        grid_energy_cost = 0
        ess_cost = 0
        ess_storage = 0
        max_ess_capacity = ess_capacity
        peak_grid_demand_kw = 0
        
        total_ppa_purchased = 0
        total_ppa_excess = 0
        total_ppa_resold = 0
        total_load_met_by_ppa = 0
        
        for hour in range(len(load_kwh)):
            hour_load = load_kwh.iloc[hour]
            hour_grid_rate = grid_rate.iloc[hour]
            hour_ppa_gen = ppa_generation_kwh.iloc[hour]
            
            mandatory_ppa = hour_ppa_gen * ppa_mintake
            optional_ppa_available = hour_ppa_gen - mandatory_ppa
            optional_ppa_purchased = 0
            
            if optional_ppa_available > 0:
                remaining_load_after_mandatory = max(0, hour_load - mandatory_ppa)
                if ppa_price < hour_grid_rate and remaining_load_after_mandatory > 0:
                    optional_ppa_purchased = min(optional_ppa_available, remaining_load_after_mandatory)
            
            total_ppa_this_hour = mandatory_ppa + optional_ppa_purchased
            ppa_cost += total_ppa_this_hour * ppa_price
            total_ppa_purchased += total_ppa_this_hour
            
            ppa_used_for_load = min(total_ppa_this_hour, hour_load)
            total_load_met_by_ppa += ppa_used_for_load
            
            remaining_load = hour_load - total_ppa_this_hour
            
            if remaining_load <= 0:
                excess_ppa = abs(remaining_load)
                total_ppa_excess += excess_ppa
                
                if ess_storage < max_ess_capacity:
                    ess_charge = min(excess_ppa, max_ess_capacity - ess_storage)
                    ess_storage += ess_charge
                    excess_ppa -= ess_charge
                
                if excess_ppa > 0 and ppa_resell:
                    resell_revenue = excess_ppa * ppa_price * ppa_resellrate
                    ppa_cost -= resell_revenue
                    total_ppa_resold += excess_ppa
                    excess_ppa = 0
                
                remaining_load = 0
            else:
                if ess_storage > 0:
                    ess_discharge = min(remaining_load, ess_storage)
                    ess_storage -= ess_discharge
                    ess_cost += ess_discharge * ppa_price * ess_price
                    remaining_load -= ess_discharge
                
                if remaining_load > 0:
                    grid_energy_cost += remaining_load * hour_grid_rate
                    grid_demand_kw = remaining_load
                    peak_grid_demand_kw = max(peak_grid_demand_kw, grid_demand_kw)
        
        grid_demand_cost = peak_grid_demand_kw * contract_fee
        grid_total_cost = grid_energy_cost + grid_demand_cost
        total_cost = ppa_cost + grid_total_cost + ess_cost
        
        return total_cost, ppa_cost, grid_total_cost, grid_demand_cost, ess_cost
    
    def run_analysis(self):
        """Run the PPA analysis"""
        try:
            # Load data files
            self.pattern_df = pd.read_excel(self.pattern_file_var.get(), index_col=0)
            self.grid_df, self.contract_fee = _kepco.process_kepco_data(
                self.kepco_file_var.get(), 2024, "HV_C_III"
            )
            
            load_df = self.pattern_df[['load']]
            solar_df = self.pattern_df[['solar']]
            ppa_df = solar_df.copy()
            ppa_df.columns = ['generation']
            
            load_capacity_mw = self.load_capacity_var.get()
            ess_include = self.ess_include_var.get()
            ess_capacity_pct = self.ess_capacity_var.get() / 100.0
            
            self.results_data = []
            
            # Run analysis for different PPA coverage levels
            for ppa_percent in range(0, 201, 10):
                ppa_coverage = ppa_percent / 100.0
                
                # Calculate ESS capacity
                ess_capacity = 0
                if ess_include:
                    peak_solar_mw = ppa_df['generation'].max() * load_capacity_mw
                    ess_capacity = peak_solar_mw * ess_capacity_pct * 1000
                
                total_cost, ppa_cost, grid_cost, grid_demand_cost, ess_cost = self.calculate_ppa_cost(
                    load_df, ppa_df, self.grid_df, load_capacity_mw, ppa_coverage, self.contract_fee, ess_capacity
                )
                
                total_electricity_kwh = load_capacity_mw * 1000 * load_df['load'].sum()
                
                self.results_data.append({
                    'ppa_percent': ppa_percent,
                    'total_cost': total_cost,
                    'ppa_cost': ppa_cost,
                    'grid_cost': grid_cost,
                    'grid_demand_cost': grid_demand_cost,
                    'ess_cost': ess_cost,
                    'total_electricity_kwh': total_electricity_kwh,
                    'total_cost_per_kwh': total_cost / total_electricity_kwh,
                    'ppa_cost_per_kwh': ppa_cost / total_electricity_kwh,
                    'grid_cost_per_kwh': (grid_cost - grid_demand_cost) / total_electricity_kwh,
                    'grid_demand_cost_per_kwh': grid_demand_cost / total_electricity_kwh,
                    'ess_cost_per_kwh': ess_cost / total_electricity_kwh
                })
            
            # Find optimal
            optimal_idx = min(range(len(self.results_data)), key=lambda i: self.results_data[i]['total_cost'])
            optimal_ppa = self.results_data[optimal_idx]['ppa_percent']
            optimal_cost = self.results_data[optimal_idx]['total_cost']
            
            messagebox.showinfo("Analysis Complete", 
                              f"Analysis completed!\nOptimal PPA coverage: {optimal_ppa}%\nTotal cost: {optimal_cost:,.0f} KRW")
            
            # Switch to results tab and update charts
            self.notebook.select(1)
            self.update_charts()
            
        except Exception as e:
            messagebox.showerror("Error", f"Analysis failed: {str(e)}")
    
    def update_charts(self, event=None):
        """Update the visualization charts"""
        if not self.results_data:
            return
        
        self.fig.clear()
        chart_type = self.chart_type_var.get()
        
        if chart_type == "Cost Analysis":
            self.plot_cost_analysis()
        elif chart_type == "PPA Coverage":
            self.plot_ppa_coverage()
        elif chart_type == "Monthly Patterns":
            self.plot_monthly_patterns()
        elif chart_type == "Hourly Patterns":
            self.plot_hourly_patterns()
        
        self.canvas.draw()
    
    def plot_cost_analysis(self):
        """Plot cost analysis chart"""
        ax1 = self.fig.add_subplot(2, 2, 1)
        ax2 = self.fig.add_subplot(2, 2, 2)
        ax3 = self.fig.add_subplot(2, 2, 3)
        ax4 = self.fig.add_subplot(2, 2, 4)
        
        ppa_percents = [r['ppa_percent'] for r in self.results_data]
        total_costs = [r['total_cost'] for r in self.results_data]
        ppa_costs = [r['ppa_cost'] for r in self.results_data]
        grid_costs = [r['grid_cost'] for r in self.results_data]
        
        # Total cost vs PPA coverage
        ax1.plot(ppa_percents, total_costs, 'b-', linewidth=2)
        ax1.set_title('Total Cost vs PPA Coverage')
        ax1.set_xlabel('PPA Coverage (%)')
        ax1.set_ylabel('Total Cost (KRW)')
        ax1.grid(True, alpha=0.3)
        
        # Cost breakdown
        ax2.plot(ppa_percents, ppa_costs, 'g-', label='PPA Cost', linewidth=2)
        ax2.plot(ppa_percents, grid_costs, 'r-', label='Grid Cost', linewidth=2)
        ax2.set_title('Cost Breakdown')
        ax2.set_xlabel('PPA Coverage (%)')
        ax2.set_ylabel('Cost (KRW)')
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        
        # Per-kWh costs
        total_costs_per_kwh = [r['total_cost_per_kwh'] for r in self.results_data]
        ax3.plot(ppa_percents, total_costs_per_kwh, 'purple', linewidth=2)
        ax3.set_title('Cost per kWh vs PPA Coverage')
        ax3.set_xlabel('PPA Coverage (%)')
        ax3.set_ylabel('Cost per kWh (KRW/kWh)')
        ax3.grid(True, alpha=0.3)
        
        # Savings vs Grid-only
        grid_only_cost = self.results_data[0]['total_cost']  # 0% PPA
        savings = [(grid_only_cost - r['total_cost']) / 1e6 for r in self.results_data]
        ax4.plot(ppa_percents, savings, 'orange', linewidth=2)
        ax4.set_title('Cost Savings vs Grid-Only (Million KRW)')
        ax4.set_xlabel('PPA Coverage (%)')
        ax4.set_ylabel('Savings (Million KRW)')
        ax4.grid(True, alpha=0.3)
        
        self.fig.tight_layout()
    
    def plot_ppa_coverage(self):
        """Plot PPA coverage analysis"""
        ax = self.fig.add_subplot(1, 1, 1)
        
        ppa_percents = [r['ppa_percent'] for r in self.results_data]
        total_costs = [r['total_cost'] / 1e6 for r in self.results_data]  # Convert to millions
        
        ax.plot(ppa_percents, total_costs, 'b-', linewidth=3, marker='o')
        
        # Mark optimal point
        optimal_idx = min(range(len(self.results_data)), key=lambda i: self.results_data[i]['total_cost'])
        optimal_ppa = self.results_data[optimal_idx]['ppa_percent']
        optimal_cost = self.results_data[optimal_idx]['total_cost'] / 1e6
        
        ax.plot(optimal_ppa, optimal_cost, 'ro', markersize=10, label=f'Optimal: {optimal_ppa}%')
        
        ax.set_title('Total Cost vs PPA Coverage', fontsize=14)
        ax.set_xlabel('PPA Coverage (%)', fontsize=12)
        ax.set_ylabel('Total Annual Cost (Million KRW)', fontsize=12)
        ax.grid(True, alpha=0.3)
        ax.legend()
        
        # Add text annotation for optimal point
        ax.annotate(f'Optimal: {optimal_ppa}%\n{optimal_cost:.1f}M KRW', 
                   xy=(optimal_ppa, optimal_cost), 
                   xytext=(optimal_ppa + 20, optimal_cost + 100),
                   arrowprops=dict(arrowstyle='->', color='red'),
                   fontsize=10, ha='center')
    
    def plot_monthly_patterns(self):
        """Plot monthly patterns if data is available"""
        if not hasattr(self, 'pattern_df'):
            return
        
        ax1 = self.fig.add_subplot(2, 1, 1)
        ax2 = self.fig.add_subplot(2, 1, 2)
        
        # Create hourly data for visualization
        hours = list(range(24))
        load_pattern = [self.pattern_df['load'].iloc[i % len(self.pattern_df)] for i in hours]
        solar_pattern = [self.pattern_df['solar'].iloc[i % len(self.pattern_df)] for i in hours]
        
        ax1.plot(hours, load_pattern, 'b-', linewidth=2, label='Load Pattern')
        ax1.plot(hours, solar_pattern, 'orange', linewidth=2, label='Solar Pattern')
        ax1.set_title('Daily Load and Solar Generation Patterns')
        ax1.set_xlabel('Hour of Day')
        ax1.set_ylabel('Normalized Generation')
        ax1.legend()
        ax1.grid(True, alpha=0.3)
        
        # Grid rate pattern
        if hasattr(self, 'grid_df'):
            grid_pattern = [self.grid_df['rate'].iloc[i % len(self.grid_df)] for i in hours]
            ax2.plot(hours, grid_pattern, 'r-', linewidth=2, label='Grid Rate')
            ax2.axhline(y=self.ppa_price_var.get(), color='g', linestyle='--', linewidth=2, label='PPA Price')
            ax2.set_title('Hourly Grid Rates vs PPA Price')
            ax2.set_xlabel('Hour of Day')
            ax2.set_ylabel('Price (KRW/kWh)')
            ax2.legend()
            ax2.grid(True, alpha=0.3)
        
        self.fig.tight_layout()
    
    def plot_hourly_patterns(self):
        """Plot hourly analysis patterns"""
        ax = self.fig.add_subplot(1, 1, 1)
        
        # Show cost components for different PPA scenarios
        scenarios_to_show = [0, 50, 100, 150, 200]  # 0%, 50%, 100%, 150%, 200%
        
        for scenario in scenarios_to_show:
            if scenario < len(self.results_data):
                data = self.results_data[scenario]
                total_cost = data['total_cost'] / 1e6  # Convert to millions
                ax.bar(scenario, total_cost, alpha=0.7, label=f'{scenario}% PPA')
        
        ax.set_title('Total Cost Comparison for Different PPA Scenarios')
        ax.set_xlabel('PPA Coverage (%)')
        ax.set_ylabel('Total Annual Cost (Million KRW)')
        ax.legend()
        ax.grid(True, alpha=0.3, axis='y')
    
    def export_results(self):
        """Export results to Excel"""
        if not self.results_data:
            messagebox.showwarning("Warning", "No results to export. Please run analysis first.")
            return
        
        try:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Results As"
            )
            
            if filename:
                df = pd.DataFrame(self.results_data)
                with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Cost_Analysis', index=False)
                
                messagebox.showinfo("Export Complete", f"Results exported to {filename}")
        
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export results: {str(e)}")

def main():
    root = tk.Tk()
    app = PPAAnalysisGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()