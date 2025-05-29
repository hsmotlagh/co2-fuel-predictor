import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define input and output paths
input_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Combined_Scenario_Data.xlsx"
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Figure_13.xlsx"

# Define columns for DataFrame
columns = [
    'Scenario', 'Ship Speed (knots)', 'CO2 Emission (kg)',
    'XGB CO2 Before MC', 'XGB CO2 After MC'
]

# Original Scenario data (corrected CO2 values)
original_data = pd.DataFrame([
    [9, 360.0436715, 360.0446777, 360.0441284],
    [10, 488.5691911, 360.0446777, 488.5692749],
    [11, 657.1781162, 657.1782837, 657.1781616],
    [12, 859.1947807, 859.1941528, 859.1948242],
    [12.5, 992.1021319, 992.1025391, 992.1021118],
    [13, 1154.959137, 1154.959351, 1154.959106],
    [13.5, 1351.230497, 1351.230347, 1351.230469],
    [14, 1591.626947, 1591.626953, 1591.626831],
    [14.5, 1894.342934, 1591.626953, 1894.342773],
    [15, 2534.163236, 2534.162109, 2534.163086]
], columns=['Ship Speed (knots)', 'CO2 Emission (kg)', 'XGB CO2 Before MC', 'XGB CO2 After MC'])

# Read Combined Scenario results
try:
    combined_data = pd.read_excel(input_path, sheet_name='Combined_Scen_Coeff')
except Exception as e:
    print(f"Error reading Combined Scenario data: {e}")
    exit()

# Verify required columns in Combined Scenario data
required_cols = [
    'Ship Speed (knots)', 'CO2 Emission (kg)',
    'Final Predicted CO2 Emission XGB Before MC (kg)', 'Final Predicted CO2 Emission XGB After MC (kg)'
]
missing_cols = [col for col in required_cols if col not in combined_data.columns]
if missing_cols:
    print(f"Missing columns in Combined Scenario data: {missing_cols}")
    exit()

# Combine data for charting
all_data = []
for _, row in original_data.iterrows():
    all_data.append(['Original Scenario'] + row[['Ship Speed (knots)', 'CO2 Emission (kg)',
                                                'XGB CO2 Before MC', 'XGB CO2 After MC']].tolist())
for _, row in combined_data.iterrows():
    all_data.append(['Combined Scenario', row['Ship Speed (knots)'], row['CO2 Emission (kg)'],
                     row.get('Final Predicted CO2 Emission XGB Before MC (kg)', np.nan),
                     row.get('Final Predicted CO2 Emission XGB After MC (kg)', np.nan)])

# Create DataFrame
try:
    df = pd.DataFrame(all_data, columns=columns)
except Exception as e:
    print(f"Error creating DataFrame: {e}")
    exit()

# Save to Excel and generate chart
try:
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Charts', index=False)

        # Pivot data for chart
        pivot_data = df.pivot_table(
            index='Ship Speed (knots)',
            columns='Scenario',
            values=['CO2 Emission (kg)', 'XGB CO2 Before MC', 'XGB CO2 After MC'],
            aggfunc='first'
        )

        # Write pivot table to Excel
        pivot_data.to_excel(writer, sheet_name='Pivot_Data', startrow=0)

        workbook = writer.book
        worksheet = writer.sheets['Charts']
        scenarios = ['Original Scenario', 'Combined Scenario']
        colors = ['#1F77B4', '#8C564B']

        # Figure 13: CO2 Emission Comparison Before and After ML (XGB) for Combined Scenario
        chart_fig13 = workbook.add_chart({'type': 'column'})
        for i, scenario in enumerate(scenarios):
            try:
                # Actual CO2
                col_actual = pivot_data.columns.get_loc(('CO2 Emission (kg)', scenario)) + 1
                chart_fig13.add_series({
                    'name': f'{scenario} Actual',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_actual)}$2:${chr(65 + col_actual)}$11',
                    'fill': {'color': colors[i], 'transparency': 20},
                    'gap': 10
                })
                # XGB CO2 Before MC
                col_before = pivot_data.columns.get_loc(('XGB CO2 Before MC', scenario)) + 1
                chart_fig13.add_series({
                    'name': f'{scenario} XGB Before MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_before)}$2:${chr(65 + col_before)}$11',
                    'fill': {'color': colors[i], 'transparency': 50},
                    'gap': 10
                })
                # XGB CO2 After MC
                col_after = pivot_data.columns.get_loc(('XGB CO2 After MC', scenario)) + 1
                chart_fig13.add_series({
                    'name': f'{scenario} XGB After MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_after)}$2:${chr(65 + col_after)}$11',
                    'fill': {'color': colors[i]},
                    'gap': 10
                })
            except KeyError:
                print(f"Column not found for {scenario} in Figure 13")
                continue
        chart_fig13.set_title({'name': 'Figure 13: CO₂ Emission Comparison Before and After ML (XGB) and After ML with MC for Combined Scenario'})
        chart_fig13.set_x_axis({'name': 'Ship Speed (knots)'})
        chart_fig13.set_y_axis({'name': 'CO2 Emission (kg)', 'major_gridlines': {'visible': True}})
        chart_fig13.set_size({'width': 720, 'height': 576})
        chart_fig13.set_legend({'position': 'top'})
        worksheet.insert_chart('F2', chart_fig13)

    print(f"Figure 13 saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")