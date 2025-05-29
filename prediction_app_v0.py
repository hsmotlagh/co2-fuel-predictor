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

# Load combined scenario data
sheet_names = [
    "Original Scenar_Coeff", "Paint (5%) Scen_Coeff", "Advance Propell_Coeff",
    "Fin (2%-4%) Sce_Coeff", "Bulbous Bow Sce_Coeff"
]
excel_data = pd.read_excel(data_path, sheet_name=sheet_names)
combined_data = pd.concat([excel_data[sheet] for sheet in sheet_names], ignore_index=True)

# Functions to assist prediction and reporting
def plot_all_scenarios(ax, y_label, user_value, speeds, speed):
    colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#9467bd', '#d62728']
    for i, sheet_name in enumerate(sheet_names):
        df = excel_data[sheet_name]
        values = [
            df[df['Ship Speed (knots)'] == s][y_label].mean()
            if not df[df['Ship Speed (knots)'] == s].empty else np.nan for s in speeds
        ]
        ax.plot(speeds, values, label=sheet_name, color=colors[i])
    ax.bar([speed], [user_value], color='#4d4d4d', width=0.2, label='User Prediction')
    ax.legend(fontsize=4, loc='upper left', frameon=True)
    ax.grid(True)
    ax.tick_params(axis='both', labelsize=8)


def compute_scenario_distances(speed, pred_fc, pred_co2):
    scenario_distances = {}
    for sheet_name in sheet_names:
        df = excel_data[sheet_name]
        if speed in df['Ship Speed (knots)'].values:
            rows = df[df['Ship Speed (knots)'] == speed]
        else:
            closest = df['Ship Speed (knots)'].iloc[(df['Ship Speed (knots)'] - speed).abs().idxmin()]
            rows = df[df['Ship Speed (knots)'] == closest]
        fc = rows['Final Predicted Fuel Consumption XGB After MC (kg/h)'].mean()
        co2 = rows['Final Predicted CO2 Emission XGB After MC (kg)'].mean()
        dist = np.sqrt((pred_fc - fc) ** 2 + (pred_co2 - co2) ** 2)
        scenario_distances[sheet_name] = {'distance': dist, 'fc': fc, 'co2': co2}
    return scenario_distances

# These functions should be integrated into the PredictionApp class
# inside the appropriate methods: predict() and generate_report().
