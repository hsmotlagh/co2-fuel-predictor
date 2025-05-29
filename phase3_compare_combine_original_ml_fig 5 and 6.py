import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define input and output paths
input_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Combined_Scenario_Data.xlsx"
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Figures_5_6.xlsx"

# Define columns for DataFrame
columns = [
    'Scenario', 'Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
    'XGB FC Before MC', 'XGB FC After MC', 'GPR FC Before MC', 'GPR FC After MC'
]

# Original Scenario data (corrected FC values)
original_data = pd.DataFrame([
    [9, 115.7696693, 115.7706909, 115.7701187, 115.7696693, 115.5215296],
    [10, 157.0962029, 115.7706909, 157.0962982, 158.6434913, 156.3614339],
    [11, 211.3112914, 211.3114166, 211.3114014, 211.3112914, 210.0374382],
    [12, 276.2684182, 276.2680664, 276.2684326, 276.2684182, 274.3117836],
    [12.5, 319.0039009, 319.0044556, 319.0038757, 319.0039009, 317.0237926],
    [13, 371.3694974, 371.3694458, 371.3694763, 371.3694974, 371.4828272],
    [13.5, 434.4792596, 434.479126, 434.4792175, 434.4792596, 439.2622737],
    [14, 511.7771533, 511.7771301, 511.7770691, 511.7771533, 521.3440264],
    [14.5, 609.1134837, 511.7771301, 609.1134033, 628.1440622, 623.9973716],
    [15, 814.8434842, 814.8422852, 814.8432007, 814.8434842, 815.3310835]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
            'XGB FC Before MC', 'XGB FC After MC', 'GPR FC Before MC', 'GPR FC After MC'])

# Advance Propeller Scenario data (corrected FC values)
adv_propeller_data = pd.DataFrame([
    [9, 103.7743379, 103.7743379, 103.7748413, 103.7743379, 103.5641821],
    [10, 140.8707293, 103.7754974, 140.8708038, 142.4288852, 141.8556822],
    [11, 189.5548816, 189.5549316, 189.5549469, 189.5548816, 191.9589791],
    [12, 247.890597, 247.8903198, 247.8906403, 247.890597, 252.2573825],
    [12.5, 286.2080063, 286.2084656, 286.2080383, 286.2080063, 291.4768664],
    [13, 332.9373352, 332.9371948, 332.9372864, 332.9373352, 338.0240017],
    [13.5, 389.4472384, 389.447113, 389.4472046, 389.4472384, 391.1423315],
    [14, 458.6313585, 458.6312866, 458.6312866, 458.6313585, 456.4251026],
    [14.5, 545.6961681, 458.6312866, 545.696106, 562.0957495, 529.980596],
    [15, 729.3067193, 729.305542, 729.3063965, 729.3067193, 638.4584508]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
            'XGB FC Before MC', 'XGB FC After MC', 'GPR FC Before MC', 'GPR FC After MC'])

# Read Combined Scenario results
try:
    combined_data = pd.read_excel(input_path, sheet_name='Combined_Scen_Coeff')
except Exception as e:
    print(f"Error reading Combined Scenario data: {e}")
    exit()

# Verify required columns in Combined Scenario data
required_cols = [
    'Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
    'Final Predicted Fuel Consumption XGB Before MC (kg/h)', 'Final Predicted Fuel Consumption XGB After MC (kg/h)',
    'Final Predicted Fuel Consumption GPR Before MC (kg/h)', 'Final Predicted Fuel Consumption GPR After MC (kg/h)'
]
missing_cols = [col for col in required_cols if col not in combined_data.columns]
if missing_cols:
    print(f"Missing columns in Combined Scenario data: {missing_cols}")
    exit()

# Combine data for charting
all_data = []
for _, row in original_data.iterrows():
    all_data.append(['Original Scenario'] + row[['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
                                                'XGB FC Before MC', 'XGB FC After MC', 'GPR FC Before MC', 'GPR FC After MC']].tolist())
for _, row in adv_propeller_data.iterrows():
    all_data.append(['Advance Propeller Scenario'] + row[['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
                                                         'XGB FC Before MC', 'XGB FC After MC', 'GPR FC Before MC', 'GPR FC After MC']].tolist())
for _, row in combined_data.iterrows():
    all_data.append(['Combined Scenario', row['Ship Speed (knots)'], row['Fuel Consumption (FC) (kg/h)'],
                     row.get('Final Predicted Fuel Consumption XGB Before MC (kg/h)', np.nan),
                     row.get('Final Predicted Fuel Consumption XGB After MC (kg/h)', np.nan),
                     row.get('Final Predicted Fuel Consumption GPR Before MC (kg/h)', np.nan),
                     row.get('Final Predicted Fuel Consumption GPR After MC (kg/h)', np.nan)])

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
            values=['Fuel Consumption (FC) (kg/h)', 'XGB FC Before MC', 'XGB FC After MC', 'GPR FC Before MC', 'GPR FC After MC'],
            aggfunc='first'
        )

        # Write pivot table to Excel
        pivot_data.to_excel(writer, sheet_name='Pivot_Data', startrow=0)

        workbook = writer.book
        worksheet = writer.sheets['Charts']
        scenarios_combined = ['Original Scenario', 'Combined Scenario']
        scenarios_adv_prop = ['Original Scenario', 'Advance Propeller Scenario']
        colors = ['#1F77B4', '#8C564B']

        # Figure 5: FC Comparison Before and After ML (XGB) for Combined Scenario
        chart_fig5 = workbook.add_chart({'type': 'column'})
        for i, scenario in enumerate(scenarios_combined):
            try:
                # Actual FC
                col_actual = pivot_data.columns.get_loc(('Fuel Consumption (FC) (kg/h)', scenario)) + 1
                chart_fig5.add_series({
                    'name': f'{scenario} Actual',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_actual)}$2:${chr(65 + col_actual)}$11',
                    'fill': {'color': colors[i], 'transparency': 20},
                    'gap': 10
                })
                # XGB FC Before MC
                col_before = pivot_data.columns.get_loc(('XGB FC Before MC', scenario)) + 1
                chart_fig5.add_series({
                    'name': f'{scenario} XGB Before MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_before)}$2:${chr(65 + col_before)}$11',
                    'fill': {'color': colors[i], 'transparency': 50},
                    'gap': 10
                })
                # XGB FC After MC
                col_after = pivot_data.columns.get_loc(('XGB FC After MC', scenario)) + 1
                chart_fig5.add_series({
                    'name': f'{scenario} XGB After MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_before + 2)}$2:${chr(65 + col_before + 2)}$11',
                    'fill': {'color': colors[i]},
                    'gap': 10
                })
            except KeyError:
                print(f"Column not found for {scenario} in Figure 5")
                continue
        chart_fig5.set_title({'name': 'Figure 5: FC Comparison Before and After ML (XGB) and After ML with MC for Combined Scenario'})
        chart_fig5.set_x_axis({'name': 'Ship Speed (knots)'})
        chart_fig5.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}})
        chart_fig5.set_size({'width': 720, 'height': 576})
        chart_fig5.set_legend({'position': 'top'})
        worksheet.insert_chart('F2', chart_fig5)

        # Figure 6: FC Comparison Before and After ML (GPR) for Advance Propeller Scenario
        chart_fig6 = workbook.add_chart({'type': 'column'})
        for i, scenario in enumerate(scenarios_adv_prop):
            try:
                # Actual FC
                col_actual = pivot_data.columns.get_loc(('Fuel Consumption (FC) (kg/h)', scenario)) + 1
                chart_fig6.add_series({
                    'name': f'{scenario} Actual',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_actual)}$2:${chr(65 + col_actual)}$11',
                    'fill': {'color': colors[i], 'transparency': 20},
                    'gap': 10
                })
                # GPR FC Before MC
                col_before = pivot_data.columns.get_loc(('GPR FC Before MC', scenario)) + 1
                chart_fig6.add_series({
                    'name': f'{scenario} GPR Before MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_before)}$2:${chr(65 + col_before)}$11',
                    'fill': {'color': colors[i], 'transparency': 50},
                    'gap': 10
                })
                # GPR FC After MC
                col_after = pivot_data.columns.get_loc(('GPR FC After MC', scenario)) + 1
                chart_fig6.add_series({
                    'name': f'{scenario} GPR After MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_after)}$2:${chr(65 + col_after)}$11',
                    'fill': {'color': colors[i]},
                    'gap': 10
                })
            except KeyError:
                print(f"Column not found for {scenario} in Figure 6")
                continue
        chart_fig6.set_title({'name': 'Figure 6: FC Comparison Before and After ML (GPR) and After ML with MC for Advance Propeller Scenario'})
        chart_fig6.set_x_axis({'name': 'Ship Speed (knots)'})
        chart_fig6.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}})
        chart_fig6.set_size({'width': 720, 'height': 576})
        chart_fig6.set_legend({'position': 'top'})
        worksheet.insert_chart('F22', chart_fig6)

    print(f"Figures 5 and 6 saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")