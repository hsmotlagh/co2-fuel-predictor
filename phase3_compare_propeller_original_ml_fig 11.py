import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Figure_11.xlsx"

# Define columns for DataFrame
columns = [
    'Scenario', 'Ship Speed (knots)', 'CO2 Emission (kg)',
    'GPR CO2 Before MC', 'GPR CO2 After MC'
]

# Original Scenario data (corrected CO2 values)
original_data = pd.DataFrame([
    [9, 360.0436715, 360.0436715, 358.813239],
    [10, 488.5691911, 493.713828, 485.6354766],
    [11, 657.1781162, 657.1781161, 653.5278652],
    [12, 859.1947807, 859.1947808, 853.706109],
    [12.5, 992.1021319, 992.1021319, 986.6739402],
    [13, 1154.959137, 1154.959137, 1156.820535],
    [13.5, 1351.230497, 1351.230497, 1368.183743],
    [14, 1591.626947, 1591.626947, 1624.497932],
    [14.5, 1894.342934, 1955.396435, 1943.520796],
    [15, 2534.163236, 2534.163236, 2519.773601]
], columns=['Ship Speed (knots)', 'CO2 Emission (kg)', 'GPR CO2 Before MC', 'GPR CO2 After MC'])

# Advance Propeller Scenario data (corrected CO2 values)
adv_propeller_data = pd.DataFrame([
    [9, 322.738191, 322.738191, 321.9157786],
    [10, 438.107968, 443.3356238, 441.4072398],
    [11, 589.5156817, 589.5156817, 598.0201232],
    [12, 770.9397568, 770.9397568, 786.7087443],
    [12.5, 890.1068997, 890.1068997, 909.3523352],
    [13, 1035.435112, 1035.435112, 1054.286444],
    [13.5, 1211.180911, 1211.180911, 1217.186155],
    [14, 1426.343525, 1426.343525, 1416.922151],
    [14.5, 1697.115083, 1750.226352, 1638.163994],
    [15, 2268.143897, 2268.143897, 1949.719923]
], columns=['Ship Speed (knots)', 'CO2 Emission (kg)', 'GPR CO2 Before MC', 'GPR CO2 After MC'])

# Combine data for charting
all_data = []
for scenario, data in [('Original Scenario', original_data), ('Advance Propeller Scenario', adv_propeller_data)]:
    for _, row in data.iterrows():
        all_data.append([scenario] + row[['Ship Speed (knots)', 'CO2 Emission (kg)',
                                         'GPR CO2 Before MC', 'GPR CO2 After MC']].tolist())

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
            values=['CO2 Emission (kg)', 'GPR CO2 Before MC', 'GPR CO2 After MC'],
            aggfunc='first'
        )

        # Write pivot table to Excel
        pivot_data.to_excel(writer, sheet_name='Pivot_Data', startrow=0)

        workbook = writer.book
        worksheet = writer.sheets['Charts']
        scenarios = ['Original Scenario', 'Advance Propeller Scenario']
        colors = ['#1F77B4', '#D62728']

        # Figure 11: CO2 Emission Comparison Before and After ML (GPR) for Advance Propeller Scenario
        chart_fig11 = workbook.add_chart({'type': 'column'})
        for i, scenario in enumerate(scenarios):
            try:
                # Actual CO2
                col_actual = pivot_data.columns.get_loc(('CO2 Emission (kg)', scenario)) + 1
                chart_fig11.add_series({
                    'name': f'{scenario} Actual',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_actual)}$2:${chr(65 + col_actual)}$11',
                    'fill': {'color': colors[i], 'transparency': 20},
                    'gap': 10
                })
                # GPR CO2 Before MC
                col_before = pivot_data.columns.get_loc(('GPR CO2 Before MC', scenario)) + 1
                chart_fig11.add_series({
                    'name': f'{scenario} GPR Before MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_before)}$2:${chr(65 + col_before)}$11',
                    'fill': {'color': colors[i], 'transparency': 50},
                    'gap': 10
                })
                # GPR CO2 After MC
                col_after = pivot_data.columns.get_loc(('GPR CO2 After MC', scenario)) + 1
                chart_fig11.add_series({
                    'name': f'{scenario} GPR After MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_after)}$2:${chr(65 + col_after)}$11',
                    'fill': {'color': colors[i]},
                    'gap': 10
                })
            except KeyError:
                print(f"Column not found for {scenario} in Figure 11")
                continue
        chart_fig11.set_title({'name': 'Figure 11: CO₂ Emission Comparison Before and After ML (GPR) and After ML with MC for Advance Propeller Scenario'})
        chart_fig11.set_x_axis({'name': 'Ship Speed (knots)'})
        chart_fig11.set_y_axis({'name': 'CO2 Emission (kg)', 'major_gridlines': {'visible': True}})
        chart_fig11.set_size({'width': 720, 'height': 576})
        chart_fig11.set_legend({'position': 'top'})
        worksheet.insert_chart('F2', chart_fig11)

    print(f"Figure 11 saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")