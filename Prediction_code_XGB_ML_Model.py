import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
import joblib
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os

# File paths
model_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/XGB_ML_model/"
data_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/ML_Apply copy.xlsx"
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_4_graphs/ML_Prediction_Results_1.xlsx"

# Ensure output directory exists
os.makedirs(os.path.dirname(output_path), exist_ok=True)

# Feature ranges (based on training data)
RANGES = {
    'Pe (kW)': (300, 2200),
    'Ship Speed (knots)': (9, 15),
    'Hull Efficiency (ηH)': (1.0, 1.1),
    'Propeller Open Water Efficiency (ηO)': (0.5, 0.65)
}

# Load combined scenario data
try:
    sheet_names = [
        "Original Scenar_Coeff", "Paint (5%) Scen_Coeff", "Advance Propell_Coeff",
        "Fin (2%-4%) Sce_Coeff", "Bulbous Bow Sce_Coeff"
    ]
    excel_data = pd.read_excel(data_path, sheet_name=sheet_names)
    combined_data = pd.concat([excel_data[sheet] for sheet in sheet_names], ignore_index=True)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

class PredictionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("XGB Emission Prediction")
        self.root.geometry("1000x600")

        # Configure style for font
        style = ttk.Style()
        style.configure("TLabel", font=("Helvetica", 8))
        style.configure("TButton", font=("Helvetica", 7))

        # Main frame
        self.main_frame = ttk.Frame(root, padding="5")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Input fields
        self.entries = {}
        row = 0
        for feature, (min_val, max_val) in RANGES.items():
            label = f"{feature} ({min_val}–{max_val}):"
            ttk.Label(self.main_frame, text=label).grid(row=row, column=0, sticky=tk.W, pady=2)
            entry = ttk.Entry(self.main_frame, width=20)
            entry.grid(row=row, column=1, sticky=tk.W, pady=2)
            self.entries[feature] = entry
            row += 1

        # Buttons
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=row, column=0, columnspan=2, pady=5)
        self.predict_button = ttk.Button(button_frame, text="Predict", command=self.predict)
        self.predict_button.grid(row=0, column=0, padx=3)
        self.save_button = ttk.Button(button_frame, text="Save Results", command=self.save_results, state='disabled')
        self.save_button.grid(row=0, column=1, padx=3)
        self.zoom_in_button = ttk.Button(button_frame, text="+ Zoom In", command=self.zoom_in)
        self.zoom_in_button.grid(row=0, column=2, padx=3)
        self.zoom_out_button = ttk.Button(button_frame, text="- Zoom Out", command=self.zoom_out)
        self.zoom_out_button.grid(row=0, column=3, padx=3)

        # Result labels
        self.fc_label = ttk.Label(self.main_frame, text="Predicted FC (After MC): N/A")
        self.fc_label.grid(row=row+1, column=0, columnspan=2, pady=2)
        self.co2_label = ttk.Label(self.main_frame, text="Predicted CO2 (After MC): N/A")
        self.co2_label.grid(row=row+2, column=0, columnspan=2, pady=2)

        # Matplotlib figure with two subplots
        self.base_figsize = (6, 3)
        self.figsize_scale = 1.0
        self.fig, (self.ax1, self.ax2) = plt.subplots(2, 1, figsize=self.base_figsize)
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.main_frame)
        self.canvas.get_tk_widget().grid(row=row+3, column=0, columnspan=2, pady=5)

        # Initialize variables
        self.xgb_model_fc = None
        self.xgb_model_co2 = None
        self.scaler = None
        self.last_prediction = None

        # Load combined model
        try:
            self.xgb_model_fc = joblib.load(f"{model_path}xgb_model_fc_combined.pkl")
            self.xgb_model_co2 = joblib.load(f"{model_path}xgb_model_co2_combined.pkl")
            self.scaler = joblib.load(f"{model_path}scaler_combined.pkl")
            print("Combined model loaded successfully.")
        except FileNotFoundError:
            messagebox.showerror("Error", "Combined model files not found.")
            self.xgb_model_fc = None
            self.xgb_model_co2 = None
            self.scaler = None

    def zoom_in(self):
        self.figsize_scale *= 1.1
        self.update_figure_size()

    def zoom_out(self):
        self.figsize_scale = max(0.5, self.figsize_scale * 0.9)
        self.update_figure_size()

    def update_figure_size(self):
        self.fig.clear()
        new_figsize = (self.base_figsize[0] * self.figsize_scale, self.base_figsize[1] * self.figsize_scale)
        self.fig, (self.ax1, self.ax2) = plt.subplots(2, 1, figsize=new_figsize)
        self.canvas.figure = self.fig
        if self.last_prediction:
            self.predict()
        self.canvas.draw()
        self.canvas.flush_events()

    def predict(self):
        if not self.xgb_model_fc or not self.xgb_model_co2 or not self.scaler:
            messagebox.showerror("Error", "Model not loaded.")
            return

        # Validate inputs
        inputs = {}
        for feature, entry in self.entries.items():
            try:
                value = float(entry.get())
                min_val, max_val = RANGES[feature]
                if not (min_val <= value <= max_val):
                    messagebox.showerror("Error", f"{feature} must be between {min_val} and {max_val}.")
                    return
                inputs[feature] = value
            except ValueError:
                messagebox.showerror("Error", f"Invalid input for {feature}.")
                return

        # Estimate remaining features
        speed = inputs['Ship Speed (knots)']
        pe = inputs['Pe (kW)']
        # Find closest matching data point based on Pe and Speed
        distances = np.sqrt(
            ((combined_data['Pe (kW)'] - pe) / RANGES['Pe (kW)'][1])**2 +
            ((combined_data['Ship Speed (knots)'] - speed) / RANGES['Ship Speed (knots)'][1])**2
        )
        ref_row = combined_data.iloc[distances.idxmin()]

        n_rpm = ref_row['n (rpm)']
        t_kn = ref_row['T (kN)']
        pd_kw = pe  # Assume Pd ≈ Pe
        pb_kw = pe * 0.95

        # Prepare input data
        user_data = pd.DataFrame({
            'Pe (kW)': [pe],
            'n (rpm)': [n_rpm],
            'T (kN)': [t_kn],
            'Hull Efficiency (ηH)': [inputs['Hull Efficiency (ηH)']],
            'Propeller Open Water Efficiency (ηO)': [inputs['Propeller Open Water Efficiency (ηO)']],
            'Delivered Power (Pd) (kW)': [pd_kw],
            'Brake Power (PB) (kW)': [pb_kw],
            'Ship Speed (knots)': [speed],
            'Physics-Based FC': [pd_kw * inputs['Hull Efficiency (ηH)'] * 0.8],
            'Physics-Based CO2': [pd_kw * inputs['Propeller Open Water Efficiency (ηO)'] * 0.5],
            'Ship Speed^2': [speed ** 2],
            'Ship Speed^3': [speed ** 3],
            'Pd * Hull Efficiency': [pd_kw * inputs['Hull Efficiency (ηH)']],
            'Propeller Efficiency * n': [inputs['Propeller Open Water Efficiency (ηO)'] * n_rpm]
        })

        # Scale and predict
        user_data_scaled = self.scaler.transform(user_data)
        pred_fc = self.xgb_model_fc.predict(user_data_scaled)[0]
        pred_co2 = self.xgb_model_co2.predict(user_data_scaled)[0]

        # Update result labels
        self.fc_label.config(text=f"Predicted FC (After MC): {pred_fc:.2f} kg/h")
        self.co2_label.config(text=f"Predicted CO2 (After MC): {pred_co2:.2f} kg")

        # Store prediction for saving
        self.last_prediction = {
            'Pe (kW)': pe,
            'Ship Speed (knots)': speed,
            'Hull Efficiency (ηH)': inputs['Hull Efficiency (ηH)'],
            'Propeller Open Water Efficiency (ηO)': inputs['Propeller Open Water Efficiency (ηO)'],
            'Predicted FC (After MC) (kg/h)': pred_fc,
            'Predicted CO2 (After MC) (kg)': pred_co2,
            'Type': 'User Prediction'
        }

        # Enable save button
        self.save_button.config(state='normal')

        # Update line graphs with prediction bar
        speeds = sorted(combined_data['Ship Speed (knots)'].unique())
        fc_values = []
        co2_values = []
        for s in speeds:
            subset = combined_data[combined_data['Ship Speed (knots)'] == s]
            fc_values.append(subset['Final Predicted Fuel Consumption XGB After MC (kg/h)'].mean())
            co2_values.append(subset['Final Predicted CO2 Emission XGB After MC (kg)'].mean())

        # Clear previous plots
        self.ax1.clear()
        self.ax2.clear()

        # FC Line Graph
        self.ax1.plot(speeds, fc_values, color=SCENARIO_COLORS['Original Scenar_Coeff'], label='Combined FC (After MC)')
        self.ax1.bar([speed], [pred_fc], color=USER_PRED_COLOR, width=0.2, label='User Prediction')
        self.ax1.set_title('Fuel Consumption (After MC)')
        self.ax1.set_xlabel('Ship Speed (knots)')
        self.ax1.set_ylabel('FC (kg/h)')
        self.ax1.legend()
        self.ax1.grid(True)

        # CO2 Line Graph
        self.ax2.plot(speeds, co2_values, color=SCENARIO_COLORS['Original Scenar_Coeff'], label='Combined CO2 (After MC)')
        self.ax2.bar([speed], [pred_co2], color=USER_PRED_COLOR, width=0.2, label='User Prediction')
        self.ax2.set_title('CO2 Emissions (After MC)')
        self.ax2.set_xlabel('Ship Speed (knots)')
        self.ax2.set_ylabel('CO2 (kg)')
        self.ax2.legend()
        self.ax2.grid(True)

        # Adjust layout and redraw
        self.fig.tight_layout()
        self.canvas.draw()
        self.canvas.flush_events()

    def save_results(self):
        if not self.last_prediction:
            messagebox.showerror("Error", "No prediction to save.")
            return

        # Save ML model results (After MC) for combined data
        model_data = combined_data[[
            'Ship Speed (knots)', 'Final Predicted Fuel Consumption XGB After MC (kg/h)',
            'Final Predicted CO2 Emission XGB After MC (kg)'
        ]].copy()
        model_data['Type'] = 'ML Model (After MC)'

        # Save user prediction
        user_data = pd.DataFrame([self.last_prediction])

        # Combine data
        result_data = pd.concat([model_data, user_data], ignore_index=True)

        try:
            if os.path.exists(output_path):
                existing_data = pd.read_excel(output_path, sheet_name='User_Predictions')
                result_data = pd.concat([existing_data, result_data], ignore_index=True)
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                result_data.to_excel(writer, sheet_name='User_Predictions', index=False)
            messagebox.showinfo("Success", f"Results saved to {output_path}")
            self.save_button.config(state='disabled')
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save results: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PredictionApp(root)
    root.mainloop()