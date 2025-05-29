import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define input and output paths
input_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Combined_Scenario_Data.xlsx"
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Figures_9_10.xlsx"

# Define columns for DataFrame
columns = [
    'Scenario', 'Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
    'NN FC Before MC', 'NN FC After MC', 'SVR FC Before MC', 'SVR FC After MC'
]

# Original Scenario data (corrected FC values)
original_data = pd.DataFrame([
    [9, 115.7696693, 117.476855, 134.6166003, 212.5441426, 131.1508187],
    [10, 157.0962029, 157.8509436, 177.0100822, 196.6363757, 158.6310184],
    [11, 211.3112914, 204.0828663, 200.7464994, 216.8806122, 204.0135631],
    [12, 276.2684182, 283.9687245, 249.6234641, 276.369045, 270.2385251],
    [12.5, 319.0039009, 316.2486729, 269.8809941, 318.903584, 315.0517762],
    [13, 371.3694974, 375.8887913, 348.4123989, 373.6568423, 371.653],
    [13.5, 434.4792596, 431.9160264, 437.9602642, 434.378943, 439.5563096],
    [14, 511.7771533, 509.1980774, 520.1257762, 474.5461671, 519.5712447],
    [14.5, 609.1134837, 608.4019767, 641.5809365, 481.4836737, 608.8219948],
    [15, 814.8434842, 815.0370316, 851.9997577, 477.5037548, 729.8122753]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
            'NN FC Before MC', 'NN FC After MC', 'SVR FC Before MC', 'SVR FC After MC'])

# Read Combined Scenario results
try:
    combined_data = pd.read_excel(input_path, sheet_name='Combined_Scen_Coeff')
except Exception as e:
    print(f"Error reading Combined Scenario data: {e}")
    exit()

# Verify required columns in Combined Scenario data
required_cols = [
    'Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
    'Final Predicted Fuel Consumption NN Before MC (kg/h)', 'Final Predicted Fuel Consumption NN After MC (kg/h)',
    'Final Predicted Fuel Consumption SVR Before MC (kg/h)', 'Final Predicted Fuel Consumption SVR After MC (kg/h)'
]
missing_cols = [col for col in required_cols if col not in combined_data.columns]
if missing_cols:
    print(f"Missing columns in Combined Scenario data: {missing_cols}")
    exit()

# Combine data for charting
all_data = []
for _, row in original_data.iterrows():
    all_data.append(['Original Scenario'] + row[['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
                                                'NN FC Before MC', 'NN FC After MC', 'SVR FC Before MC', 'SVR FC After MC']].tolist())
for _, row in combined_data.iterrows():
    all_data.append(['Combined Scenario', row['Ship Speed (knots)'], row['Fuel Consumption (FC) (kg/h)'],
                     row.get('Final Predicted Fuel Consumption NN Before MC (kg/h)', np.nan),
                     row.get('Final Predicted Fuel Consumption NN After MC (kg/h)', np.nan),
                     row.get('Final Predicted Fuel Consumption SVR Before MC (kg/h)', np.nan),
                     row.get('Final Predicted Fuel Consumption SVR After MC (kg/h)', np.nan)])

# Create DataFrame
try:
    df = pd.DataFrame(all_data, columns=columns)
except Exception as e:
    print(f"Error creating DataFrame: {e}")
    exit()

# Save to Excel and generate charts
try:
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Charts', index=False)

        # Pivot data for charts
        pivot_data = df.pivot_table(
            index='Ship Speed (knots)',
            columns='Scenario',
            values=['Fuel Consumption (FC) (kg/h)', 'NN FC Before MC', 'NN FC After MC', 'SVR FC Before MC', 'SVR FC After MC'],
            aggfunc='first'
        )

        # Write pivot table to Excel
        pivot_data.to_excel(writer, sheet_name='Pivot_Data', startrow=0)

        workbook = writer.book
        worksheet = writer.sheets['Charts']
        scenarios = ['Original Scenario', 'Combined Scenario']
        colors = ['#1F77B4', '#8C564B']

        # Figure 9: FC Comparison Before and After ML (NN) for Combined Scenario
        chart_fig9 = workbook.add_chart({'type': 'column'})
        for i, scenario in enumerate(scenarios):
            try:
                # Actual FC
                col_actual = pivot_data.columns.get_loc(('Fuel Consumption (FC) (kg/h)', scenario)) + 1
                chart_fig9.add_series({
                    'name': f'{scenario} Actual',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_actual)}$2:${chr(65 + col_actual)}$11',
                    'fill': {'color': colors[i], 'transparency': 20},
                    'gap': 10
                })
                # NN FC Before MC
                col_before = pivot_data.columns.get_loc(('NN FC Before MC', scenario)) + 1
                chart_fig9.add_series({
                    'name': f'{scenario} NN Before MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_before)}$2:${chr(65 + col_before)}$11',
                    'fill': {'color': colors[i], 'transparency': 50},
                    'gap': 10
                })
                # NN FC After MC
                col_after = pivot_data.columns.get_loc(('NN FC After MC', scenario)) + 1
                chart_fig9.add_series({
                    'name': f'{scenario} NN After MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_after)}$2:${chr(65 + col_after)}$11',
                    'fill': {'color': colors[i]},
                    'gap': 10
                })
            except KeyError:
                print(f"Column not found for {scenario} in Figure 9")
                continue
        chart_fig9.set_title({'name': 'Figure 9: FC Comparison Before and After ML (NN) and After ML with MC for Combined Scenario'})
        chart_fig9.set_x_axis({'name': 'Ship Speed (knots)'})
        chart_fig9.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}})
        chart_fig9.set_size({'width': 720, 'height': 576})
        chart_fig9.set_legend({'position': 'top'})
        worksheet.insert_chart('F2', chart_fig9)

        # Figure 10: FC Comparison Before and After ML (SVR) for Combined Scenario
        chart_fig10 = workbook.add_chart({'type': 'column'})
        for i, scenario in enumerate(scenarios):
            try:
                # Actual FC
                col_actual = pivot_data.columns.get_loc(('Fuel Consumption (FC) (kg/h)', scenario)) + 1
                chart_fig10.add_series({
                    'name': f'{scenario} Actual',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_actual)}$2:${chr(65 + col_actual)}$11',
                    'fill': {'color': colors[i], 'transparency': 20},
                    'gap': 10
                })
                # SVR FC Before MC
                col_before = pivot_data.columns.get_loc(('SVR FC Before MC', scenario)) + 1
                chart_fig10.add_series({
                    'name': f'{scenario} SVR Before MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_before)}$2:${chr(65 + col_before)}$11',
                    'fill': {'color': colors[i], 'transparency': 50},
                    'gap': 10
                })
                # SVR FC After MC
                col_after = pivot_data.columns.get_loc(('SVR FC After MC', scenario)) + 1
                chart_fig10.add_series({
                    'name': f'{scenario} SVR After MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_after)}$2:${chr(65 + col_after)}$11',
                    'fill': {'color': colors[i]},
                    'gap': 10
                })
            except KeyError:
                print(f"Column not found for {scenario} in Figure 10")
                continue
        chart_fig10.set_title({'name': 'Figure 10: FC Comparison Before and After ML (SVR) and After ML with MC for Combined Scenario'})
        chart_fig10.set_x_axis({'name': 'Ship Speed (knots)'})
        chart_fig10.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}})
        chart_fig10.set_size({'width': 720, 'height': 576})
        chart_fig10.set_legend({'position': 'top'})
        worksheet.insert_chart('F22', chart_fig10)

    print(f"Figures 9 and 10 saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")