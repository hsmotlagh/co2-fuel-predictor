import pandas as pd
import os

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/phase2-step1_all_speed_v2.xlsx"

# Hardcode provided FC and CO2 values for all speeds
comparison_data = [
    ["Original Scenario", 9, 115.7696693, 360.0436715],
    ["Original Scenario", 10, 157.0962029, 488.5691911],
    ["Original Scenario", 11, 211.3112914, 657.1781162],
    ["Original Scenario", 12, 276.2684182, 859.1947807],
    ["Original Scenario", 12.5, 319.0039009, 992.1021319],
    ["Original Scenario", 13, 371.3694974, 1154.959137],
    ["Original Scenario", 13.5, 434.4792596, 1351.230497],
    ["Original Scenario", 14, 511.7771533, 1591.626947],
    ["Original Scenario", 14.5, 609.1134837, 1894.342934],
    ["Original Scenario", 15, 814.8434842, 2534.163236],
    ["Paint (5%) Scenario", 9, 110.0681824, 342.3120472],
    ["Paint (5%) Scenario", 10, 149.4348682, 464.7424402],
    ["Paint (5%) Scenario", 11, 201.2853668, 625.9974908],
    ["Paint (5%) Scenario", 12, 263.4790258, 819.4197703],
    ["Paint (5%) Scenario", 12.5, 304.5852633, 947.2601688],
    ["Paint (5%) Scenario", 13, 354.9289312, 1103.828976],
    ["Paint (5%) Scenario", 13.5, 415.9640576, 1293.648219],
    ["Paint (5%) Scenario", 14, 490.9831319, 1526.95754],
    ["Paint (5%) Scenario", 14.5, 585.731981, 1821.626461],
    ["Paint (5%) Scenario", 15, 787.8448646, 2450.197529],
    ["Bulbous Bow Scenario", 9, 119.6541899, 372.1245307],
    ["Bulbous Bow Scenario", 10, 162.3385399, 504.8728592],
    ["Bulbous Bow Scenario", 11, 217.9322612, 677.7693324],
    ["Bulbous Bow Scenario", 12, 284.2566267, 884.038109],
    ["Bulbous Bow Scenario", 12.5, 323.9702129, 1007.547362],
    ["Bulbous Bow Scenario", 13, 371.0228762, 1153.881145],
    ["Bulbous Bow Scenario", 13.5, 425.0311648, 1321.846922],
    ["Bulbous Bow Scenario", 14, 501.2150135, 1558.778692],
    ["Bulbous Bow Scenario", 14.5, 600.5953901, 1867.851663],
    ["Bulbous Bow Scenario", 15, 863.1846317, 2684.504205],
    ["Combined Paint and Bulbous Bow Scenario", 9, 108.589, 337.71179],
    ["Combined Paint and Bulbous Bow Scenario", 10, 147.583, 459.08333],
    ["Combined Paint and Bulbous Bow Scenario", 11, 199.028, 619.17708],
    ["Combined Paint and Bulbous Bow Scenario", 12, 260.120, 808.97320],
    ["Combined Paint and Bulbous Bow Scenario", 12.5, 301.539, 937.78629],
    ["Combined Paint and Bulbous Bow Scenario", 13, 351.128, 1092.00808],
    ["Combined Paint and Bulbous Bow Scenario", 13.5, 411.981, 1281.26091],
    ["Combined Paint and Bulbous Bow Scenario", 14, 486.464, 1512.90224],
    ["Combined Paint and Bulbous Bow Scenario", 14.5, 579.315, 1801.66965],
    ["Combined Paint and Bulbous Bow Scenario", 15, 771.037, 2397.92507]
]

# Save results to Excel
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Save comparison table
    comparison_df = pd.DataFrame(comparison_data, columns=["Scenario", "Ship Speed (knots)", "Fuel Consumption (kg/h)", "CO2 Emission (kg)"])
    comparison_df.to_excel(writer, sheet_name='Comparison_All_Speeds', index=False)

    # Pivot data for clustered column chart
    pivot_fc = comparison_df.pivot(index='Ship Speed (knots)', columns='Scenario', values='Fuel Consumption (kg/h)')
    pivot_co2 = comparison_df.pivot(index='Ship Speed (knots)', columns='Scenario', values='CO2 Emission (kg)')
    pivot_fc.to_excel(writer, sheet_name='Pivot_Data', startrow=0)
    pivot_co2.to_excel(writer, sheet_name='Pivot_Data', startrow=12)

    # Create interactive chart for FC
    workbook = writer.book
    worksheet = writer.sheets['Comparison_All_Speeds']
    chart_fc = workbook.add_chart({'type': 'column'})
    scenarios = ["Original Scenario", "Paint (5%) Scenario", "Bulbous Bow Scenario", "Combined Paint and Bulbous Bow Scenario"]
    colors = ['#1F77B4', '#2CA02C', '#D62728', '#9467BD']  # Blue, Green, Red, Purple
    for i, scenario in enumerate(scenarios):
        col = pivot_fc.columns.get_loc(scenario) + 1  # Column index in Pivot_Data (skip index column)
        chart_fc.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$2:$A$11',
            'values': f'=Pivot_Data!${chr(66 + i)}$2:${chr(66 + i)}$11',
            'fill': {'color': colors[i]},
            'gap': 20
        })
    chart_fc.set_title({'name': 'Fuel Consumption Comparison Across Speeds'})
    chart_fc.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_fc.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}})
    chart_fc.set_size({'width': 720, 'height': 576})
    chart_fc.set_legend({'position': 'top'})
    worksheet.insert_chart('F2', chart_fc)

    # Create interactive chart for CO2
    chart_co2 = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios):
        col = pivot_co2.columns.get_loc(scenario) + 1
        chart_co2.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$14:$A$23',
            'values': f'=Pivot_Data!${chr(66 + i)}$14:${chr(66 + i)}$23',
            'fill': {'color': colors[i]},
            'gap': 20
        })
    chart_co2.set_title({'name': 'CO2 Emission Comparison Across Speeds'})
    chart_co2.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_co2.set_y_axis({'name': 'CO2 Emission (kg)', 'major_gridlines': {'visible': True}})
    chart_co2.set_size({'width': 720, 'height': 576})
    chart_co2.set_legend({'position': 'top'})
    worksheet.insert_chart('F20', chart_co2)

print(f"Results saved to {output_path}")