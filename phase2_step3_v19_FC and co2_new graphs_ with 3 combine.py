import pandas as pd
import os

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/phase2-step1_fc_co2_graphs_v4.xlsx"

# Hardcode provided FC and CO2 values
comparison_data = [
    # Original Scenario
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
    # Paint (5%) Scenario
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
    # Bulbous Bow Scenario
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
    # Advance Propeller Scenario
    ["Advance Propeller Scenario", 9, 103.7743379, 322.738191],
    ["Advance Propeller Scenario", 10, 140.8707293, 438.107968],
    ["Advance Propeller Scenario", 11, 189.5548816, 589.5156817],
    ["Advance Propeller Scenario", 12, 247.890597, 770.9397568],
    ["Advance Propeller Scenario", 12.5, 286.2080063, 890.1068997],
    ["Advance Propeller Scenario", 13, 332.9373352, 1035.435112],
    ["Advance Propeller Scenario", 13.5, 389.4472384, 1211.180911],
    ["Advance Propeller Scenario", 14, 458.6313585, 1426.343525],
    ["Advance Propeller Scenario", 14.5, 545.6961681, 1697.115083],
    ["Advance Propeller Scenario", 15, 729.3067193, 2268.143897],
    # Fin Installation Scenario
    ["Fin Installation Scenario", 9, 115.4654418, 359.097524],
    ["Fin Installation Scenario", 10, 156.6860586, 487.2936421],
    ["Fin Installation Scenario", 11, 210.7942821, 655.5702172],
    ["Fin Installation Scenario", 12, 275.6058519, 857.1341993],
    ["Fin Installation Scenario", 12.5, 318.2599935, 989.7885799],
    ["Fin Installation Scenario", 13, 370.516624, 1152.306701],
    ["Fin Installation Scenario", 13.5, 433.5948752, 1348.480062],
    ["Fin Installation Scenario", 14, 510.8733655, 1588.816167],
    ["Fin Installation Scenario", 14.5, 608.260007, 1891.688622],
    ["Fin Installation Scenario", 15, 814.3657819, 2532.677582],
    # Combined Paint and Bulbous Bow Scenario
    ["Combined Paint and Bulbous Bow Scenario", 9, 108.59, 337.71179],
    ["Combined Paint and Bulbous Bow Scenario", 10, 147.58, 459.08333],
    ["Combined Paint and Bulbous Bow Scenario", 11, 199.03, 619.17708],
    ["Combined Paint and Bulbous Bow Scenario", 12, 260.12, 808.9732],
    ["Combined Paint and Bulbous Bow Scenario", 12.5, 301.54, 937.78629],
    ["Combined Paint and Bulbous Bow Scenario", 13, 351.13, 1092.00808],
    ["Combined Paint and Bulbous Bow Scenario", 13.5, 411.98, 1281.26091],
    ["Combined Paint and Bulbous Bow Scenario", 14, 486.46, 1512.90224],
    ["Combined Paint and Bulbous Bow Scenario", 14.5, 579.32, 1801.66965],
    ["Combined Paint and Bulbous Bow Scenario", 15, 771.04, 2397.92507]
]

# Save results to Excel
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Save comparison table
    comparison_df = pd.DataFrame(comparison_data, columns=["Scenario", "Ship Speed (knots)", "Fuel Consumption (kg/h)",
                                                           "CO2 Emission (kg)"])
    comparison_df.to_excel(writer, sheet_name='Comparison_All_Speeds', index=False)

    # Pivot data for graphs
    pivot_fc = comparison_df.pivot(index='Ship Speed (knots)', columns='Scenario', values='Fuel Consumption (kg/h)')
    pivot_co2 = comparison_df.pivot(index='Ship Speed (knots)', columns='Scenario', values='CO2 Emission (kg)')
    pivot_fc.to_excel(writer, sheet_name='Pivot_Data', startrow=0)
    pivot_co2.to_excel(writer, sheet_name='Pivot_Data', startrow=12)

    # Debug: Print pivot table to verify data
    print("Pivot FC Table:\n", pivot_fc)
    print("Pivot CO2 Table:\n", pivot_co2)

    # Create graphs
    workbook = writer.book
    worksheet = writer.sheets['Comparison_All_Speeds']

    # Scenarios excluding Combined
    scenarios_excl = ["Original Scenario", "Paint (5%) Scenario", "Bulbous Bow Scenario", "Advance Propeller Scenario",
                      "Fin Installation Scenario"]
    colors_excl = ['#1F77B4', '#2CA02C', '#D62728', '#FF7F0E', '#9467BD']  # Blue, Green, Red, Orange, Purple

    # FC graph (excluding Combined)
    chart_fc_excl = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios_excl):
        col = pivot_fc.columns.get_loc(scenario) + 1
        print(f"FC Excl Series: {scenario}, Column: {chr(66 + i)}, Values: {pivot_fc[scenario].values}")
        chart_fc_excl.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$2:$A$11',
            'values': f'=Pivot_Data!${chr(66 + i)}$2:${chr(66 + i)}$11',
            'fill': {'color': colors_excl[i]},
            'gap': 20
        })
    chart_fc_excl.set_title({'name': 'Comparative FC Trends (Excl. Combined)'})
    chart_fc_excl.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_fc_excl.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}, 'min': 0})
    chart_fc_excl.set_size({'width': 720, 'height': 576})
    chart_fc_excl.set_legend({'position': 'top'})
    worksheet.insert_chart('F2', chart_fc_excl)

    # CO2 graph (excluding Combined)
    chart_co2_excl = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios_excl):
        col = pivot_co2.columns.get_loc(scenario) + 1
        print(f"CO2 Excl Series: {scenario}, Column: {chr(66 + i)}, Values: {pivot_co2[scenario].values}")
        chart_co2_excl.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$14:$A$23',
            'values': f'=Pivot_Data!${chr(66 + i)}$14:${chr(66 + i)}$23',
            'fill': {'color': colors_excl[i]},
            'gap': 20
        })
    chart_co2_excl.set_title({'name': 'CO2 Emission Trends (Excl. Combined)'})
    chart_co2_excl.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_co2_excl.set_y_axis({'name': 'CO2 Emission (kg)', 'major_gridlines': {'visible': True}, 'min': 0})
    chart_co2_excl.set_size({'width': 720, 'height': 576})
    chart_co2_excl.set_legend({'position': 'top'})
    worksheet.insert_chart('F20', chart_co2_excl)

    # Scenarios including Combined
    scenarios_incl = ["Original Scenario", "Paint (5%) Scenario", "Bulbous Bow Scenario",
                      "Combined Paint and Bulbous Bow Scenario"]
    colors_incl = ['#1F77B4', '#2CA02C', '#D62728', '#9467BD']  # Blue, Green, Red, Purple

    # FC graph (including Combined)
    chart_fc_incl = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios_incl):
        col = pivot_fc.columns.get_loc(scenario) + 1
        print(f"FC Incl Series: {scenario}, Column: {chr(66 + i)}, Values: {pivot_fc[scenario].values}")
        chart_fc_incl.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$2:$A$11',
            'values': f'=Pivot_Data!${chr(66 + i)}$2:${chr(66 + i)}$11',
            'fill': {'color': colors_incl[i]},
            'gap': 20
        })
    chart_fc_incl.set_title({'name': 'Comparative FC Trends for Original and Modified Scenarios'})
    chart_fc_incl.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_fc_incl.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}, 'min': 0})
    chart_fc_incl.set_size({'width': 720, 'height': 576})
    chart_fc_incl.set_legend({'position': 'top'})
    worksheet.insert_chart('P2', chart_fc_incl)

    # CO2 graph (including Combined)
    chart_co2_incl = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios_incl):
        col = pivot_co2.columns.get_loc(scenario) + 1
        print(f"CO2 Incl Series: {scenario}, Column: {chr(66 + i)}, Values: {pivot_co2[scenario].values}")
        chart_co2_incl.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$14:$A$23',
            'values': f'=Pivot_Data!${chr(66 + i)}$14:${chr(66 + i)}$23',
            'fill': {'color': colors_incl[i]},
            'gap': 20
        })
    chart_co2_incl.set_title({'name': 'CO2 Emission Trends Across Scenarios'})
    chart_co2_incl.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_co2_incl.set_y_axis({'name': 'CO2 Emission (kg)', 'major_gridlines': {'visible': True}, 'min': 0})
    chart_co2_incl.set_size({'width': 720, 'height': 576})
    chart_co2_incl.set_legend({'position': 'top'})
    worksheet.insert_chart('P20', chart_co2_incl)

print(f"Graphs saved to {output_path}")