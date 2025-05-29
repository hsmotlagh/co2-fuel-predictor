import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
import joblib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from docx import Document
import os
from datetime import datetime

# File paths
model_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/XGB_ML_model/"
data_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/ML_Apply copy.xlsx"
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_4_graphs/ML_Prediction_Results_1.xlsx"
report_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_4_graphs/Prediction_Report.docx"

# Ensure output directory exists
os.makedirs(os.path.dirname(output_path), exist_ok=True)

# Feature ranges
RANGES = {
    'Pe (kW)': (300, 2200),
    'Ship Speed (knots)': (9, 15),
    'Hull Efficiency (ηH)': (1.0, 1.1),
    'Propeller Open Water Efficiency (ηO)': (0.5, 0.65)
}

# Scenario sheet names
sheet_names = [
    "Original Scenar_Coeff",
    "Paint (5%) Scen_Coeff",
    "Advance Propell_Coeff",
    "Fin (2%-4%) Sce_Coeff",
    "Bulbous Bow Sce_Coeff"
]

# Load scenario data and combined dataset
excel_data = pd.read_excel(data_path, sheet_name=sheet_names)
combined_data = pd.concat([excel_data[s] for s in sheet_names], ignore_index=True)

class PredictionApp_v2:
    def __init__(self, root):
        self.root = root
        self.root.title("XGB Emission Prediction v2")
        self.root.geometry("1200x700")

        # Style
        style = ttk.Style()
        style.configure("TLabel", font=("Helvetica", 12))
        style.configure("TButton", font=("Helvetica", 10))

        # Notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)

        # Tabs: Graphs and Data
        self.tab_graphs = ttk.Frame(self.notebook)
        self.tab_data = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_graphs, text="Graphs")
        self.notebook.add(self.tab_data, text="Scenario Data")

        # Build UI in Graphs tab
        self._build_controls()
        self._build_canvas()

        # Build UI in Data tab
        self._build_data_tab()

        # Load models
        try:
            self.xgb_model_fc = joblib.load(f"{model_path}xgb_model_fc_combined.pkl")
            self.xgb_model_co2 = joblib.load(f"{model_path}xgb_model_co2_combined.pkl")
            self.scaler = joblib.load(f"{model_path}scaler_combined.pkl")
        except Exception:
            messagebox.showerror("Error", "Model files not found.")
            self.xgb_model_fc = self.xgb_model_co2 = self.scaler = None

        self.last_prediction = None
        self.user_predictions = []

    def _build_controls(self):
        frame = ttk.Frame(self.tab_graphs, padding=10)
        frame.pack(fill='x')

        # Input fields
        self.entries = {}
        for idx, (feature, (min_val, max_val)) in enumerate(RANGES.items()):
            ttk.Label(frame, text=f"{feature} ({min_val}–{max_val}):").grid(row=0, column=idx*2, sticky='w')
            entry = ttk.Entry(frame, width=10)
            entry.grid(row=0, column=idx*2+1, padx=5)
            self.entries[feature] = entry

        # Buttons
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=1, column=0, columnspan=len(RANGES)*2, pady=10)
        ttk.Button(btn_frame, text="Predict", command=self.predict).pack(side='left', padx=5)
        self.save_btn = ttk.Button(btn_frame, text="Save Results", command=self.save_results, state='disabled')
        self.save_btn.pack(side='left', padx=5)
        self.report_btn = ttk.Button(btn_frame, text="Generate Report", command=self.generate_report, state='disabled')
        self.report_btn.pack(side='left', padx=5)

    def _build_canvas(self):
        self.fig = plt.Figure(figsize=(8, 6))
        self.ax1 = self.fig.add_subplot(211)
        self.ax2 = self.fig.add_subplot(212)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.tab_graphs)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)

    def _build_data_tab(self):
        top = ttk.Frame(self.tab_data, padding=10)
        top.pack(fill='x')
        ttk.Label(top, text="Scenario:").pack(side='left')
        self.cmb = ttk.Combobox(top, values=sheet_names, state='readonly')
        self.cmb.current(0)
        self.cmb.pack(side='left', padx=5)
        self.cmb.bind("<<ComboboxSelected>>", self._on_scenario_change)

        # Table
        cols = list(excel_data[sheet_names[0]].columns)
        self.tv = ttk.Treeview(self.tab_data, columns=cols, show='headings')
        for c in cols:
            self.tv.heading(c, text=c)
            self.tv.column(c, width=100, anchor='center')
        self.tv.pack(fill='both', expand=True)

        # Load initial data
        self._load_scenario_data(sheet_names[0])

    def _on_scenario_change(self, event):
        self._load_scenario_data(self.cmb.get())

    def _load_scenario_data(self, scenario):
        df = excel_data[scenario]
        self.tv.delete(*self.tv.get_children())
        for _, row in df.iterrows():
            self.tv.insert('', 'end', values=list(row[col] for col in df.columns))

    def predict(self):
        if not (self.xgb_model_fc and self.xgb_model_co2 and self.scaler):
            messagebox.showerror("Error", "Models not loaded.")
            return

        # Validate and collect inputs
        inputs = {}
        for feature, entry in self.entries.items():
            try:
                val = float(entry.get())
                min_val, max_val = RANGES[feature]
                if not (min_val <= val <= max_val):
                    raise ValueError
                inputs[feature] = val
            except:
                messagebox.showerror("Error", f"Invalid input for {feature}")
                return

        # K-NN estimate remaining features
        speed, pe = inputs['Ship Speed (knots)'], inputs['Pe (kW)']
        dist = np.sqrt(
            ((combined_data['Pe (kW)'] - pe)/RANGES['Pe (kW)'][1])**2 +
            ((combined_data['Ship Speed (knots)'] - speed)/RANGES['Ship Speed (knots)'][1])**2
        )
        idx = dist.nsmallest(3).index
        rows = combined_data.loc[idx]
        w = 1/(dist.loc[idx]+1e-6)
        w /= w.sum()
        n_rpm = (rows['n (rpm)']*w).sum()
        t_kn = (rows['T (kN)']*w).sum()
        pd_kw, pb_kw = pe, pe*0.95

        # Prepare DataFrame
        user_df = pd.DataFrame({
            'Pe (kW)': [pe], 'n (rpm)': [n_rpm], 'T (kN)': [t_kn],
            'Hull Efficiency (ηH)': [inputs['Hull Efficiency (ηH)']],
            'Propeller Open Water Efficiency (ηO)': [inputs['Propeller Open Water Efficiency (ηO)']],
            'Delivered Power (Pd) (kW)': [pd_kw], 'Brake Power (PB) (kW)': [pb_kw],
            'Ship Speed (knots)': [speed],
            'Physics-Based FC': [pd_kw*inputs['Hull Efficiency (ηH)']*0.8],
            'Physics-Based CO2': [pd_kw*inputs['Propeller Open Water Efficiency (ηO)']*0.5],
            'Ship Speed^2': [speed**2], 'Ship Speed^3': [speed**3],
            'Pd * Hull Efficiency': [pd_kw*inputs['Hull Efficiency (ηH)']],
            'Propeller Efficiency * n': [inputs['Propeller Open Water Efficiency (ηO)']*n_rpm]
        })

        X = self.scaler.transform(user_df)
        pred_fc = self.xgb_model_fc.predict(X)[0]
        pred_co2 = self.xgb_model_co2.predict(X)[0]

        # Store
        self.last_prediction = dict(**inputs, **{
            'Predicted FC (kg/h)': pred_fc,
            'Predicted CO2 (kg)': pred_co2
        })
        self.user_predictions.append(self.last_prediction)
        self.save_btn.config(state='normal')
        self.report_btn.config(state='normal')

        # Update graphs
        self._update_graphs(pred_fc, pred_co2)

    def _update_graphs(self, pred_fc, pred_co2):
        self.ax1.clear(); self.ax2.clear()
        # Plot each scenario
        for scen in sheet_names:
            df = excel_data[scen]
            speeds = df['Ship Speed (knots)']
            fc_vals = df['Final Predicted Fuel Consumption XGB After MC (kg/h)']
            co2_vals = df['Final Predicted CO2 Emission XGB After MC (kg)']
            self.ax1.plot(speeds, fc_vals, label=scen)
            self.ax2.plot(speeds, co2_vals, label=scen)

        # User bar
        sp = self.last_prediction['Ship Speed (knots)']
        self.ax1.bar([sp], [pred_fc], width=0.2, label='You', alpha=0.7)
        self.ax2.bar([sp], [pred_co2], width=0.2, label='You', alpha=0.7)

        self.ax1.set_title('Fuel Consumption (After MC)')
        self.ax1.set_xlabel('Speed (knots)'); self.ax1.set_ylabel('FC (kg/h)')
        self.ax1.legend(fontsize=8)
        self.ax1.grid(True)

        self.ax2.set_title('CO2 Emission (After MC)')
        self.ax2.set_xlabel('Speed (knots)'); self.ax2.set_ylabel('CO2 (kg)')
        self.ax2.legend(fontsize=8)
        self.ax2.grid(True)

        self.fig.tight_layout()
        self.canvas.draw()

    def save_results(self):
        df_user = pd.DataFrame(self.user_predictions)
        df_model = combined_data[['Ship Speed (knots)',
                                  'Final Predicted Fuel Consumption XGB After MC (kg/h)',
                                  'Final Predicted CO2 Emission XGB After MC (kg)']].copy()
        df_model['Type'] = 'Model'
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as w:
                df_user.to_excel(w, sheet_name='User', index=False)
                df_model.to_excel(w, sheet_name='Model', index=False)
            messagebox.showinfo("Saved", f"Results saved to {output_path}")
            self.save_btn.config(state='disabled')
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def generate_report(self):
        if not self.user_predictions:
            messagebox.showerror("Error", "No predictions to report.")
            return
        latest = self.user_predictions[-1]
        report = [f"Prediction Report - {datetime.now():%Y-%m-%d %H:%M:%S}"]
        report.append(f"Inputs: {{}}".format(latest))
        report.append(f"Predicted FC: {latest['Predicted FC (kg/h)']:.2f} kg/h")
        report.append(f"Predicted CO2: {latest['Predicted CO2 (kg)']:.2f} kg")
        doc = Document()
        for line in report:
            doc.add_paragraph(line)
        try:
            doc.save(report_path)
            messagebox.showinfo("Saved", f"Report saved to {report_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == '__main__':
    root = tk.Tk()
    app = PredictionApp_v2(root)
    root.mainloop()
