import pandas as pd
import os

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/phase2-step1_combined_params_v2.xlsx"

# Hardcode provided Rt, Pe, Treq values for all scenarios
comparison_data = [
    # Original Scenario
    ["Original Scenario", 9, 67.83484964, 314.0482199, 89.7286371],
    ["Original Scenario", 10, 82.91383244, 426.5087541, 109.5295012],
    ["Original Scenario", 11, 101.2616638, 572.9789987, 133.5905855],
    ["Original Scenario", 12, 121.6124557, 750.6893668, 160.0163891],
    ["Original Scenario", 12.5, 134.6913255, 866.0652227, 176.7602696],
    ["Original Scenario", 13, 149.8513221, 1002.085761, 196.140474],
    ["Original Scenario", 13.5, 167.5756882, 1163.712609, 218.767217],
    ["Original Scenario", 14, 188.8585755, 1360.083917, 245.5898251],
    ["Original Scenario", 14.5, 214.5647195, 1600.39533, 278.2940591],
    ["Original Scenario", 15, 269.1337357, 2076.635905, 346.3754643],
    # Paint (5%) Scenario
    ["Paint (5%) Scenario", 9, 65.00455308, 300.945079, 85.98485858],
    ["Paint (5%) Scenario", 10, 79.46765766, 408.781631, 104.9770907],
    ["Paint (5%) Scenario", 11, 97.14332675, 549.6758001, 128.1574232],
    ["Paint (5%) Scenario", 12, 116.7662969, 720.7749977, 153.6398644],
    ["Paint (5%) Scenario", 12.5, 129.4605611, 832.4314076, 169.8957494],
    ["Paint (5%) Scenario", 13, 144.2222434, 964.442986, 188.7725699],
    ["Paint (5%) Scenario", 13.5, 161.5346485, 1121.761213, 210.8807421],
    ["Paint (5%) Scenario", 14, 182.3919873, 1313.514136, 237.1807377],
    ["Paint (5%) Scenario", 14.5, 207.6590523, 1548.887339, 269.3372922],
    ["Paint (5%) Scenario", 15, 261.7755133, 2019.85986, 336.9054225],
    # Bulbous Bow Scenario
    ["Bulbous Bow Scenario", 9, 69.29085975, 320.7889643, 92.25005292],
    ["Bulbous Bow Scenario", 10, 84.699344, 435.6934255, 112.6111415],
    ["Bulbous Bow Scenario", 11, 103.2905863, 584.4594532, 137.1429527],
    ["Bulbous Bow Scenario", 12, 123.8226029, 764.3321633, 163.9600144],
    ["Bulbous Bow Scenario", 12.5, 135.6252412, 872.0703012, 179.1046976],
    ["Bulbous Bow Scenario", 13, 148.8083127, 995.1109488, 195.9860825],
    ["Bulbous Bow Scenario", 13.5, 163.4970007, 1135.388572, 214.7546376],
    ["Bulbous Bow Scenario", 14, 184.4681193, 1328.465608, 241.3303844],
    ["Bulbous Bow Scenario", 14.5, 210.7977341, 1572.298139, 275.0420581],
    ["Bulbous Bow Scenario", 15, 280.5197137, 2164.490111, 363.1135134],
    # Advance Propeller Scenario
    ["Advance Propeller Scenario", 9, 67.83484964, 314.0482199, 89.7286371],
    ["Advance Propeller Scenario", 10, 82.91383244, 426.5087541, 109.5295012],
    ["Advance Propeller Scenario", 11, 101.2616638, 572.9789987, 133.5905855],
    ["Advance Propeller Scenario", 12, 121.6124557, 750.6893668, 160.0163891],
    ["Advance Propeller Scenario", 12.5, 134.6913255, 866.0652227, 176.7602696],
    ["Advance Propeller Scenario", 13, 149.8513221, 1002.085761, 196.140474],
    ["Advance Propeller Scenario", 13.5, 167.5756882, 1163.712609, 218.767217],
    ["Advance Propeller Scenario", 14, 188.8585755, 1360.083917, 245.5898251],
    ["Advance Propeller Scenario", 14.5, 214.5647195, 1600.39533, 278.2940591],
    ["Advance Propeller Scenario", 15, 269.1337357, 2076.635905, 346.3754643],
    # Fin Installation Scenario
    ["Fin Installation Scenario", 9, 67.83484964, 314.0482199, 90.31160087],
    ["Fin Installation Scenario", 10, 82.91383244, 426.5087541, 110.237233],
    ["Fin Installation Scenario", 11, 101.2616638, 572.9789987, 134.449073],
    ["Fin Installation Scenario", 12, 121.6124557, 750.6893668, 161.0334424],
    ["Fin Installation Scenario", 12.5, 134.6913255, 866.0652227, 177.8713822],
    ["Fin Installation Scenario", 13, 149.8513221, 1002.085761, 197.3597647],
    ["Fin Installation Scenario", 13.5, 167.5756882, 1163.712609, 220.1120268],
    ["Fin Installation Scenario", 14, 188.8585755, 1360.083917, 247.074198],
    ["Fin Installation Scenario", 14.5, 214.5647195, 1600.39533, 279.9570986],
    ["Fin Installation Scenario", 15, 269.1337357, 2076.635905, 348.3751466],
    # Combined Paint and Bulbous Bow Scenario
    ["Combined Paint and Bulbous Bow Scenario", 9, 64.54, 298.91, 85.45],
    ["Combined Paint and Bulbous Bow Scenario", 10, 79.02, 406.47, 104.51],
    ["Combined Paint and Bulbous Bow Scenario", 11, 96.68, 547.37, 127.60],
    ["Combined Paint and Bulbous Bow Scenario", 12, 116.21, 717.23, 153.07],
    ["Combined Paint and Bulbous Bow Scenario", 12.5, 129.01, 829.45, 169.36],
    ["Combined Paint and Bulbous Bow Scenario", 13, 143.87, 961.24, 188.28],
    ["Combined Paint and Bulbous Bow Scenario", 13.5, 161.23, 1119.50, 210.34],
    ["Combined Paint and Bulbous Bow Scenario", 14, 181.84, 1309.58, 236.62],
    ["Combined Paint and Bulbous Bow Scenario", 14.5, 206.82, 1542.61, 268.74],
    ["Combined Paint and Bulbous Bow Scenario", 15, 260.44, 2009.38, 335.85]
]

# Save results to Excel
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Save comparison table
    comparison_df = pd.DataFrame(comparison_data,
                                 columns=["Scenario", "Ship Speed (knots)", "Total Resistance (Rt) (kN)",
                                          "Effective Power (Pe) (kW)", "Thrust Required (Treq) (kN)"])
    comparison_df.to_excel(writer, sheet_name='Comparison_All_Speeds', index=False)

    # Pivot data for graphs
    pivot_rt = comparison_df.pivot(index='Ship Speed (knots)', columns='Scenario', values='Total Resistance (Rt) (kN)')
    pivot_pe = comparison_df.pivot(index='Ship Speed (knots)', columns='Scenario', values='Effective Power (Pe) (kW)')
    pivot_treq = comparison_df.pivot(index='Ship Speed (knots)', columns='Scenario',
                                     values='Thrust Required (Treq) (kN)')
    pivot_rt.to_excel(writer, sheet_name='Pivot_Data', startrow=0)
    pivot_pe.to_excel(writer, sheet_name='Pivot_Data', startrow=12)
    pivot_treq.to_excel(writer, sheet_name='Pivot_Data', startrow=24)

    # Create graphs excluding Combined Scenario
    workbook = writer.book
    worksheet = writer.sheets['Comparison_All_Speeds']

    # Rt (excluding Combined)
    chart_rt = workbook.add_chart({'type': 'column'})
    scenarios_excl = ["Original Scenario", "Paint (5%) Scenario", "Bulbous Bow Scenario", "Advance Propeller Scenario",
                      "Fin Installation Scenario"]
    colors_excl = ['#1F77B4', '#2CA02C', '#D62728', '#FF7F0E', '#9467BD']  # Blue, Green, Red, Orange, Purple
    for i, scenario in enumerate(scenarios_excl):
        col = pivot_rt.columns.get_loc(scenario) + 1
        chart_rt.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$2:$A$11',
            'values': f'=Pivot_Data!${chr(66 + i)}$2:${chr(66 + i)}$11',
            'fill': {'color': colors_excl[i]},
            'gap': 20
        })
    chart_rt.set_title({'name': 'Total Resistance (Rt) vs Ship Speed (Excl. Combined)'})
    chart_rt.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_rt.set_y_axis({'name': 'Total Resistance (Rt) (kN)', 'major_gridlines': {'visible': True}})
    chart_rt.set_size({'width': 720, 'height': 576})
    chart_rt.set_legend({'position': 'top'})
    worksheet.insert_chart('F2', chart_rt)

    # Pe (excluding Combined)
    chart_pe = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios_excl):
        col = pivot_pe.columns.get_loc(scenario) + 1
        chart_pe.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$14:$A$23',
            'values': f'=Pivot_Data!${chr(66 + i)}$14:${chr(66 + i)}$23',
            'fill': {'color': colors_excl[i]},
            'gap': 20
        })
    chart_pe.set_title({'name': 'Effective Power (Pe) vs Ship Speed (Excl. Combined)'})
    chart_pe.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_pe.set_y_axis({'name': 'Effective Power (Pe) (kW)', 'major_gridlines': {'visible': True}})
    chart_pe.set_size({'width': 720, 'height': 576})
    chart_pe.set_legend({'position': 'top'})
    worksheet.insert_chart('F20', chart_pe)

    # Treq (excluding Combined)
    chart_treq = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios_excl):
        col = pivot_treq.columns.get_loc(scenario) + 1
        chart_treq.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$26:$A$35',
            'values': f'=Pivot_Data!${chr(66 + i)}$26:${chr(66 + i)}$35',
            'fill': {'color': colors_excl[i]},
            'gap': 20
        })
    chart_treq.set_title({'name': 'Thrust Required (Treq) vs Ship Speed (Excl. Combined)'})
    chart_treq.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_treq.set_y_axis({'name': 'Thrust Required (Treq) (kN)', 'major_gridlines': {'visible': True}})
    chart_treq.set_size({'width': 720, 'height': 576})
    chart_treq.set_legend({'position': 'top'})
    worksheet.insert_chart('F38', chart_treq)

    # Create graphs including Combined Scenario
    scenarios_incl = ["Original Scenario", "Paint (5%) Scenario", "Bulbous Bow Scenario",
                      "Combined Paint and Bulbous Bow Scenario"]
    colors_incl = ['#1F77B4', '#2CA02C', '#D62728', '#9467BD']  # Blue, Green, Red, Purple

    # Rt (including Combined)
    chart_rt_comb = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios_incl):
        col = pivot_rt.columns.get_loc(scenario) + 1
        chart_rt_comb.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$2:$A$11',
            'values': f'=Pivot_Data!${chr(66 + i)}$2:${chr(66 + i)}$11',
            'fill': {'color': colors_incl[i]},
            'gap': 20
        })
    chart_rt_comb.set_title({'name': 'Total Resistance (Rt) vs Ship Speed'})
    chart_rt_comb.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_rt_comb.set_y_axis({'name': 'Total Resistance (Rt) (kN)', 'major_gridlines': {'visible': True}})
    chart_rt_comb.set_size({'width': 720, 'height': 576})
    chart_rt_comb.set_legend({'position': 'top'})
    worksheet.insert_chart('P2', chart_rt_comb)

    # Pe (including Combined)
    chart_pe_comb = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios_incl):
        col = pivot_pe.columns.get_loc(scenario) + 1
        chart_pe_comb.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$14:$A$23',
            'values': f'=Pivot_Data!${chr(66 + i)}$14:${chr(66 + i)}$23',
            'fill': {'color': colors_incl[i]},
            'gap': 20
        })
    chart_pe_comb.set_title({'name': 'Effective Power (Pe) vs Ship Speed'})
    chart_pe_comb.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_pe_comb.set_y_axis({'name': 'Effective Power (Pe) (kW)', 'major_gridlines': {'visible': True}})
    chart_pe_comb.set_size({'width': 720, 'height': 576})
    chart_pe_comb.set_legend({'position': 'top'})
    worksheet.insert_chart('P20', chart_pe_comb)

    # Treq (including Combined)
    chart_treq_comb = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios_incl):
        col = pivot_treq.columns.get_loc(scenario) + 1
        chart_treq_comb.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$A$26:$A$35',
            'values': f'=Pivot_Data!${chr(66 + i)}$26:${chr(66 + i)}$35',
            'fill': {'color': colors_incl[i]},
            'gap': 20
        })
    chart_treq_comb.set_title({'name': 'Thrust Required (Treq) vs Ship Speed'})
    chart_treq_comb.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_treq_comb.set_y_axis({'name': 'Thrust Required (Treq) (kN)', 'major_gridlines': {'visible': True}})
    chart_treq_comb.set_size({'width': 720, 'height': 576})
    chart_treq_comb.set_legend({'position': 'top'})
    worksheet.insert_chart('P38', chart_treq_comb)

print(f"Results saved to {output_path}")