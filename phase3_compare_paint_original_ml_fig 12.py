import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Figure_12.xlsx"

# Define columns for DataFrame
columns = [
    'Scenario', 'Ship Speed (knots)', 'CO2 Emission (kg)',
    'RF CO2 Before MC', 'RF CO2 After MC'
]

# Original Scenario data (corrected CO2 values)
original_data = pd.DataFrame([
    [9, 360.0436715, 479.5619646, 365.1846923],
    [10, 488.5691911, 491.4473423, 487.2839359],
    [11, 657.1781162, 719.1216592, 657.1781162],
    [12, 859.1947807, 899.5066409, 859.1947807],
    [12.5, 992.1021319, 972.180545, 991.0725549],
    [13, 1154.959137, 1113.764349, 1150.741714],
    [13.5, 1351.230497, 1325.653944, 1356.479677],
    [14, 1591.626947, 1488.195197, 1586.819018],
    [14.5, 1894.342934, 1713.64544, 1904.11218],
    [15, 2534.163236, 2199.086615, 2527.765033]
], columns=['Ship Speed (knots)', 'CO2 Emission (kg)', 'RF CO2 Before MC', 'RF CO2 After MC'])

# Paint Scenario data (corrected CO2 values)
paint_data = pd.DataFrame([
    [9, 342.3120472, 453.9660716, 345.984959],
    [10, 464.7424402, 459.6397804, 464.7424402],
    [11, 625.9974908, 676.4859729, 625.9974908],
    [12, 819.4197703, 858.2467571, 820.6981743],
    [12.5, 947.2601688, 932.9003626, 946.2690489],
    [13, 1103.828976, 1067.093867, 1087.226196],
    [13.5, 1293.648219, 1268.987043, 1287.257755],
    [14, 1526.95754, 1426.649858, 1533.464514],
    [14.5, 1821.626461, 1621.449695, 1825.357815],
    [15, 2450.197529, 2122.1974, 2450.197529]
], columns=['Ship Speed (knots)', 'CO2 Emission (kg)', 'RF CO2 Before MC', 'RF CO2 After MC'])

# Combine data for charting
all_data = []
for scenario, data in [('Original Scenario', original_data), ('Paint Scenario', paint_data)]:
    for _, row in data.iterrows():
        all_data.append([scenario] + row[['Ship Speed (knots)', 'CO2 Emission (kg)',
                                         'RF CO2 Before MC', 'RF CO2 After MC']].tolist())

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
            values=['CO2 Emission (kg)', 'RF CO2 Before MC', 'RF CO2 After MC'],
            aggfunc='first'
        )

        # Write pivot table to Excel
        pivot_data.to_excel(writer, sheet_name='Pivot_Data', startrow=0)

        workbook = writer.book
        worksheet = writer.sheets['Charts']
        scenarios = ['Original Scenario', 'Paint Scenario']
        colors = ['#1F77B4', '#FF7F0E']

        # Figure 12: CO2 Emission Comparison Before and After ML (RF) for Paint Modification Scenario
        chart_fig12 = workbook.add_chart({'type': 'column'})
        for i, scenario in enumerate(scenarios):
            try:
                # Actual CO2
                col_actual = pivot_data.columns.get_loc(('CO2 Emission (kg)', scenario)) + 1
                chart_fig12.add_series({
                    'name': f'{scenario} Actual',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_actual)}$2:${chr(65 + col_actual)}$11',
                    'fill': {'color': colors[i], 'transparency': 20},
                    'gap': 10
                })
                # RF CO2 Before MC
                col_before = pivot_data.columns.get_loc(('RF CO2 Before MC', scenario)) + 1
                chart_fig12.add_series({
                    'name': f'{scenario} RF Before MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_before)}$2:${chr(65 + col_before)}$11',
                    'fill': {'color': colors[i], 'transparency': 50},
                    'gap': 10
                })
                # RF CO2 After MC
                col_after = pivot_data.columns.get_loc(('RF CO2 After MC', scenario)) + 1
                chart_fig12.add_series({
                    'name': f'{scenario} RF After MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_after)}$2:${chr(65 + col_after)}$11',
                    'fill': {'color': colors[i]},
                    'gap': 10
                })
            except KeyError:
                print(f"Column not found for {scenario} in Figure 12")
                continue
        chart_fig12.set_title({'name': 'Figure 12: CO₂ Emission Comparison Before and After ML (RF) and After ML with MC for Paint Modification Scenario'})
        chart_fig12.set_x_axis({'name': 'Ship Speed (knots)'})
        chart_fig12.set_y_axis({'name': 'CO2 Emission (kg)', 'major_gridlines': {'visible': True}})
        chart_fig12.set_size({'width': 720, 'height': 576})
        chart_fig12.set_legend({'position': 'top'})
        worksheet.insert_chart('F2', chart_fig12)

    print(f"Figure 12 saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")