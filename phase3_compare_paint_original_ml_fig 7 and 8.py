import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Figures_7_8.xlsx"

# Define columns for DataFrame
columns = [
    'Scenario', 'Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
    'RF FC Before MC', 'RF FC After MC', 'XGB FC Before MC', 'XGB FC After MC'
]

# Original Scenario data (corrected FC values)
original_data = pd.DataFrame([
    [9, 115.7696693, 153.8689349, 117.4227306, 115.7706909, 115.7701187],
    [10, 157.0962029, 157.1669438, 156.6829376, 115.7706909, 157.0962982],
    [11, 211.3112914, 229.842333, 211.3112914, 211.3114166, 211.3114014],
    [12, 276.2684182, 289.2304312, 276.2684182, 276.2680664, 276.2684326],
    [12.5, 319.0039009, 314.198828, 318.2454924, 319.0044556, 319.0038757],
    [13, 371.3694974, 362.3647279, 366.9677781, 371.3694458, 371.3694763],
    [13.5, 434.4792596, 427.659364, 434.6211409, 434.479126, 434.4792175],
    [14, 511.7771533, 478.5193559, 512.177922, 511.7771301, 511.7770691],
    [14.5, 609.1134837, 545.0306977, 609.1134837, 511.7771301, 609.1134033],
    [15, 814.8434842, 707.1018056, 812.7861842, 814.8422852, 814.8432007]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
            'RF FC Before MC', 'RF FC After MC', 'XGB FC Before MC', 'XGB FC After MC'])

# Paint Scenario data (corrected FC values)
paint_data = pd.DataFrame([
    [9, 110.0681824, 145.9697979, 112.0365167, 110.0694275, 110.0686035],
    [10, 149.4348682, 147.7941416, 149.0412014, 110.0694275, 149.434967],
    [11, 201.2853668, 217.9306678, 200.7668618, 201.2854462, 201.2854309],
    [12, 263.4790258, 275.9635875, 263.4790258, 263.4786377, 263.479126],
    [12.5, 304.5852633, 299.9679622, 304.3589495, 304.5857544, 304.5852661],
    [13, 354.9289312, 342.6251478, 352.1221403, 354.9288025, 354.928894],
    [13.5, 415.9640576, 410.7555031, 414.6594368, 415.9638672, 415.9640808],
    [14, 490.9831319, 458.7298578, 492.1279181, 490.9830627, 490.9830322],
    [14.5, 585.731981, 509.7395255, 585.0367962, 490.9830627, 585.7319946],
    [15, 787.8448646, 682.3785854, 787.8448646, 787.843689, 787.8445435]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
            'RF FC Before MC', 'RF FC After MC', 'XGB FC Before MC', 'XGB FC After MC'])

# Combine data for charting
all_data = []
for scenario, data in [('Original Scenario', original_data), ('Paint Scenario', paint_data)]:
    for _, row in data.iterrows():
        all_data.append([scenario] + row[['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
                                         'RF FC Before MC', 'RF FC After MC', 'XGB FC Before MC', 'XGB FC After MC']].tolist())

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
            values=['Fuel Consumption (FC) (kg/h)', 'RF FC Before MC', 'RF FC After MC', 'XGB FC Before MC', 'XGB FC After MC'],
            aggfunc='first'
        )

        # Write pivot table to Excel
        pivot_data.to_excel(writer, sheet_name='Pivot_Data', startrow=0)

        workbook = writer.book
        worksheet = writer.sheets['Charts']
        scenarios = ['Original Scenario', 'Paint Scenario']
        colors = ['#1F77B4', '#FF7F0E']

        # Figure 7: FC Comparison Before and After ML (RF) for Paint Scenario
        chart_fig7 = workbook.add_chart({'type': 'column'})
        for i, scenario in enumerate(scenarios):
            try:
                # Actual FC
                col_actual = pivot_data.columns.get_loc(('Fuel Consumption (FC) (kg/h)', scenario)) + 1
                chart_fig7.add_series({
                    'name': f'{scenario} Actual',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_actual)}$2:${chr(65 + col_actual)}$11',
                    'fill': {'color': colors[i], 'transparency': 20},
                    'gap': 10
                })
                # RF FC Before MC
                col_before = pivot_data.columns.get_loc(('RF FC Before MC', scenario)) + 1
                chart_fig7.add_series({
                    'name': f'{scenario} RF Before MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_before)}$2:${chr(65 + col_before)}$11',
                    'fill': {'color': colors[i], 'transparency': 50},
                    'gap': 10
                })
                # RF FC After MC
                col_after = pivot_data.columns.get_loc(('RF FC After MC', scenario)) + 1
                chart_fig7.add_series({
                    'name': f'{scenario} RF After MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_after)}$2:${chr(65 + col_after)}$11',
                    'fill': {'color': colors[i]},
                    'gap': 10
                })
            except KeyError:
                print(f"Column not found for {scenario} in Figure 7")
                continue
        chart_fig7.set_title({'name': 'Figure 7: FC Comparison Before and After ML (RF) and After ML with MC for Paint Modification Scenario'})
        chart_fig7.set_x_axis({'name': 'Ship Speed (knots)'})
        chart_fig7.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}})
        chart_fig7.set_size({'width': 720, 'height': 576})
        chart_fig7.set_legend({'position': 'top'})
        worksheet.insert_chart('F2', chart_fig7)

        # Figure 8: FC Comparison Before and After ML (XGB) for Paint Scenario
        chart_fig8 = workbook.add_chart({'type': 'column'})
        for i, scenario in enumerate(scenarios):
            try:
                # Actual FC
                col_actual = pivot_data.columns.get_loc(('Fuel Consumption (FC) (kg/h)', scenario)) + 1
                chart_fig8.add_series({
                    'name': f'{scenario} Actual',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_actual)}$2:${chr(65 + col_actual)}$11',
                    'fill': {'color': colors[i], 'transparency': 20},
                    'gap': 10
                })
                # XGB FC Before MC
                col_before = pivot_data.columns.get_loc(('XGB FC Before MC', scenario)) + 1
                chart_fig8.add_series({
                    'name': f'{scenario} XGB Before MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_before)}$2:${chr(65 + col_before)}$11',
                    'fill': {'color': colors[i], 'transparency': 50},
                    'gap': 10
                })
                # XGB FC After MC
                col_after = pivot_data.columns.get_loc(('XGB FC After MC', scenario)) + 1
                chart_fig8.add_series({
                    'name': f'{scenario} XGB After MC',
                    'categories': '=Pivot_Data!$A$2:$A$11',
                    'values': f'=Pivot_Data!${chr(65 + col_after)}$2:${chr(65 + col_after)}$11',
                    'fill': {'color': colors[i]},
                    'gap': 10
                })
            except KeyError:
                print(f"Column not found for {scenario} in Figure 8")
                continue
        chart_fig8.set_title({'name': 'Figure 8: FC Comparison Before and After ML (XGB) and After ML with MC for Paint Modification Scenario'})
        chart_fig8.set_x_axis({'name': 'Ship Speed (knots)'})
        chart_fig8.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}})
        chart_fig8.set_size({'width': 720, 'height': 576})
        chart_fig8.set_legend({'position': 'top'})
        worksheet.insert_chart('F22', chart_fig8)

    print(f"Figures 7 and 8 saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")