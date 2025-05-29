import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/MSE_Figure_16.xlsx"

# MSE results provided by user
results = [
    {'Scenario': 'Original', 'MSE_XGB': 3.20129E-08, 'MSE_GPR': 34.61564091, 'MSE_RF': 2.726694526, 'MSE_SNN': 702.9967199},
    {'Scenario': 'Paint', 'MSE_XGB': 3.16396E-08, 'MSE_GPR': 26.65448824, 'MSE_RF': 1.572330678, 'MSE_SNN': 318.0265551},
    {'Scenario': 'Advance Propeller', 'MSE_XGB': 3.82875E-08, 'MSE_GPR': 858.7625862, 'MSE_RF': 2.350944354, 'MSE_SNN': 663.6611939},
    {'Scenario': 'Fin', 'MSE_XGB': 4.14829E-08, 'MSE_GPR': 11.53390635, 'MSE_RF': 5.924156541, 'MSE_SNN': 572.3498592},
    {'Scenario': 'Bulbous Bow', 'MSE_XGB': 3.16116E-08, 'MSE_GPR': 270.8768963, 'MSE_RF': 3.632037997, 'MSE_SNN': 219.6918737},
    {'Scenario': 'Combined', 'MSE_XGB': 3.49193E-08, 'MSE_GPR': 99.97322306, 'MSE_RF': 1.100682394, 'MSE_SNN': 180.5276211}
]

# Create DataFrame for results
results_df = pd.DataFrame(results)

# Save to Excel and generate chart
try:
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        results_df.to_excel(writer, sheet_name='MSE_Results', index=False)

        workbook = writer.book
        worksheet = writer.sheets['MSE_Results']
        models = ['MSE_XGB', 'MSE_GPR', 'MSE_RF', 'MSE_SNN']
        model_names = ['XGB', 'GPR', 'RF', 'SNN']
        colors = ['#1F77B4', '#FF7F0E', '#2CA02C', '#9467BD']

        # Figure 16: MSE Performance Comparison of XGB, GPR, RF, and SNN for FC Across Scenarios
        chart_fig16 = workbook.add_chart({'type': 'bar'})
        for i, (col, name) in enumerate(zip(models, model_names)):
            chart_fig16.add_series({
                'name': name,
                'categories': '=MSE_Results!$A$2:$A$7',
                'values': f'=MSE_Results!${chr(66 + i)}$2:${chr(66 + i)}$7',
                'fill': {'color': colors[i]},
                'gap': 10,
                'data_labels': {'value': True, 'num_format': '0.00E+00' if col == 'MSE_XGB' else '0.00'}
            })
        chart_fig16.set_title({'name': 'Figure 16: MSE Performance Comparison of XGB, GPR, RF, and SNN for FC Across Scenarios'})
        chart_fig16.set_x_axis({
            'name': 'MSE (kg/h)²',
            'log_base': 10,
            'min': 1e-8,
            'max': 1000,
            'major_gridlines': {'visible': True}
        })
        chart_fig16.set_y_axis({'name': 'Scenario', 'major_gridlines': {'visible': True}})
        chart_fig16.set_size({'width': 720, 'height': 576})
        chart_fig16.set_legend({'position': 'top'})
        worksheet.insert_chart('I2', chart_fig16)

    print(f"MSE results with Figure 16 saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")