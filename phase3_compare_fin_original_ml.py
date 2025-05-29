import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/ML_Fin_Original_Comparison.xlsx"

# Define required columns
columns = [
    'Scenario', 'Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
    'SVR FC Before MC', 'SVR CO2 Before MC', 'SVR FC After MC', 'SVR CO2 After MC',
    'GPR FC Before MC', 'GPR CO2 Before MC', 'GPR FC After MC', 'GPR CO2 After MC',
    'RF FC Before MC', 'RF CO2 Before MC', 'RF FC After MC', 'RF CO2 After MC',
    'XGB FC Before MC', 'XGB CO2 Before MC', 'XGB FC After MC', 'XGB CO2 After MC',
    'NN FC Before MC', 'NN CO2 Before MC', 'NN FC After MC', 'NN CO2 After MC'
]

# Data for Original Scenario (verified to match corrected values provided on May 21, 2025)
original_data = pd.DataFrame([
    [9, 115.7696693, 360.0436715, 212.5441426, 916.5662006, 131.1508187, 447.1187997, 115.7696693, 360.0436715, 115.5215296, 358.813239, 153.8689349, 479.5619646, 117.4227306, 365.1846923, 115.7706909, 360.0446777, 115.7701187, 360.0441284, 117.476855, 358.8237916, 134.6166003, 448.428842],
    [10, 157.0962029, 488.5691911, 196.6363757, 870.0368642, 158.6310184, 526.2801538, 158.6434913, 493.713828, 156.3614339, 485.6354766, 157.1669438, 491.4473423, 156.6829376, 487.2839359, 115.7706909, 360.0446777, 157.0962982, 488.5692749, 157.8509436, 435.609839, 177.0100822, 537.1271601],
    [11, 211.3112914, 657.1781162, 216.8806122, 859.0989233, 204.0135631, 660.4106049, 211.3112914, 657.1781161, 210.0374382, 653.5278652, 229.842333, 719.1216592, 211.3112914, 657.1781162, 211.3114166, 657.1782837, 211.3114014, 657.1781616, 204.0828663, 611.9956983, 200.7464994, 630.8712249],
    [12, 276.2684182, 859.1947807, 276.369045, 929.6611844, 270.2385251, 855.6961727, 276.2684182, 859.1947808, 274.3117836, 853.706109, 289.2304312, 899.5066409, 276.2684182, 859.1947807, 276.2680664, 859.1941528, 276.2684326, 859.1948242, 283.9687245, 918.0395544, 249.6234641, 775.7901294],
    [12.5, 319.0039009, 992.1021319, 318.903584, 1011.335582, 315.0517762, 988.4495129, 319.0039009, 992.1021319, 317.0237926, 986.6739402, 314.198828, 972.180545, 318.2454924, 991.0725549, 319.0044556, 992.1025391, 319.0038757, 992.1021118, 316.2486729, 1006.558513, 269.8809941, 835.7401571],
    [13, 371.3694974, 1154.959137, 373.6568423, 1135.725684, 371.653, 1156.598379, 371.3694974, 1154.959137, 371.4828272, 1156.820535, 362.3647279, 1113.764349, 366.9677781, 1150.741714, 371.3694458, 1154.959351, 371.3694763, 1154.959106, 375.8887913, 1125.634976, 348.4123989, 1076.905761],
    [13.5, 434.4792596, 1351.230497, 434.378943, 1255.635094, 439.5563096, 1350.774585, 434.4792596, 1351.230497, 439.2622737, 1368.183743, 427.659364, 1325.653944, 434.6211409, 1356.479677, 434.479126, 1351.230347, 434.4792175, 1351.230469, 431.9160264, 1324.754651, 437.9602642, 1355.812357],
    [14, 511.7771533, 1591.626947, 474.5461671, 1298.03187, 519.5712447, 1566.882446, 511.7771533, 1591.626947, 521.3440264, 1624.497932, 478.5193559, 1488.195197, 512.177922, 1586.819018, 511.7771301, 1591.626953, 511.7770691, 1591.626831, 509.1980774, 1597.243742, 520.1257762, 1617.521944],
    [14.5, 609.1134837, 1894.342934, 481.4836737, 1259.176407, 608.8219948, 1786.114488, 628.1440622, 1955.396435, 623.9973716, 1943.520796, 545.0306977, 1713.64544, 609.1134837, 1904.11218, 511.7771301, 1591.626953, 609.1134033, 1894.342773, 608.4019767, 1950.373941, 641.5809365, 1992.570497],
    [15, 814.8434842, 2534.163236, 477.5037548, 1212.024875, 729.8122753, 2007.914881, 814.8434842, 2534.163236, 815.3310835, 2519.773601, 707.1018056, 2199.086615, 812.7861842, 2527.765033, 814.8422852, 2534.162109, 814.8432007, 2534.163086, 815.0370316, 2539.708027, 851.9997577, 2607.000819]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
            'SVR FC Before MC', 'SVR CO2 Before MC', 'SVR FC After MC', 'SVR CO2 After MC',
            'GPR FC Before MC', 'GPR CO2 Before MC', 'GPR FC After MC', 'GPR CO2 After MC',
            'RF FC Before MC', 'RF CO2 Before MC', 'RF FC After MC', 'RF CO2 After MC',
            'XGB FC Before MC', 'XGB CO2 Before MC', 'XGB FC After MC', 'XGB CO2 After MC',
            'NN FC Before MC', 'NN CO2 Before MC', 'NN FC After MC', 'NN CO2 After MC'])

# Data for Fin Scenario (verified to match corrected values provided on May 21, 2025)
fin_data = pd.DataFrame([
    [9, 115.4654418, 359.097524, 212.0958908, 913.770889, 132.2985703, 432.2445051, 115.4654418, 359.097524, 115.4334622, 359.7679461, 153.8108697, 477.0252609, 116.2898542, 361.6614464, 115.4665375, 359.0986328, 115.4659271, 359.0979919, 115.0758962, 357.2560946, 92.21518494, 318.8370427],
    [10, 156.6860586, 487.2936421, 196.2611189, 867.2527734, 159.6467296, 516.3940987, 158.1669174, 492.2309249, 156.6452291, 488.1874082, 154.7641581, 479.9899878, 156.6860586, 487.2936421, 115.4665375, 359.0986328, 156.6861877, 487.2937317, 140.7286946, 432.8192804, 151.8848255, 518.9360556],
    [11, 210.7942821, 655.5702172, 216.4905454, 856.3071276, 203.8605224, 657.270079, 210.7942821, 655.5702172, 210.8205532, 656.4474727, 228.9630601, 717.381292, 210.7942821, 653.8874515, 210.7943573, 655.5702515, 210.7943268, 655.5702515, 199.8009168, 614.1815599, 213.8552332, 675.1926038],
    [12, 275.6058519, 857.1341993, 275.7063009, 927.0795856, 270.2442762, 859.667355, 275.6058519, 857.1341993, 276.1299273, 860.1094409, 288.5443227, 897.3728436, 277.7385589, 861.7512785, 275.6055298, 857.133606, 275.6059265, 857.1343384, 293.6497456, 915.8674439, 267.1862687, 803.11335],
    [12.5, 318.2599935, 989.7885799, 318.1597652, 1009.083118, 314.9616482, 994.5758689, 318.2599935, 989.7885799, 319.2214946, 994.5699582, 310.2706482, 961.0182486, 316.9803693, 983.1558609, 318.260376, 989.7889404, 318.2599792, 989.7885742, 319.6082213, 1000.317825, 281.3093046, 830.594657],
    [13, 370.516624, 1152.306701, 372.8123109, 1133.012162, 370.4364224, 1161.800517, 370.516624, 1152.306701, 372.4533639, 1160.137876, 355.0749839, 1103.594104, 367.5976586, 1154.326015, 370.516571, 1152.306519, 370.516571, 1152.306641, 361.3500475, 1122.179805, 341.3365451, 1035.605585],
    [13.5, 433.5948752, 1348.480062, 433.4946492, 1252.54077, 437.6316248, 1354.98607, 433.5948752, 1348.480062, 437.1422487, 1359.871399, 425.3773728, 1340.384008, 434.6516648, 1354.611666, 433.5947571, 1348.480103, 433.5947571, 1348.480103, 428.8722687, 1324.758744, 441.6574526, 1368.795645],
    [14, 510.8733655, 1588.816167, 473.8840309, 1295.051641, 518.1438336, 1565.763989, 510.8733655, 1588.816167, 516.1530362, 1604.662059, 477.626285, 1485.417746, 506.2366561, 1592.470255, 510.8733521, 1588.816284, 510.8733215, 1588.81604, 513.2912434, 1595.878574, 528.4768684, 1664.443742],
    [14.5, 608.260007, 1891.688622, 480.8145243, 1256.27127, 608.4796762, 1776.293066, 626.9943073, 1951.814179, 615.8368798, 1914.260509, 545.6078108, 1694.878558, 612.3821225, 1885.631173, 510.8733521, 1588.816284, 608.2599487, 1891.688477, 627.7010608, 1946.36451, 638.2565635, 1970.512321],
    [15, 814.3657819, 2532.677582, 476.842285, 1209.179179, 728.9895, 1972.527355, 814.3657819, 2532.677582, 810.8277316, 2510.328947, 706.4854355, 2197.169704, 812.3047242, 2532.677582, 814.364624, 2532.67627, 814.3654175, 2532.677246, 814.8688945, 2537.091747, 854.2167696, 2676.540663]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
            'SVR FC Before MC', 'SVR CO2 Before MC', 'SVR FC After MC', 'SVR CO2 After MC',
            'GPR FC Before MC', 'GPR CO2 Before MC', 'GPR FC After MC', 'GPR CO2 After MC',
            'RF FC Before MC', 'RF CO2 Before MC', 'RF FC After MC', 'RF CO2 After MC',
            'XGB FC Before MC', 'XGB CO2 Before MC', 'XGB FC After MC', 'XGB CO2 After MC',
            'NN FC Before MC', 'NN CO2 Before MC', 'NN FC After MC', 'NN CO2 After MC'])

# Combine scenario data
all_data = []
for scenario, data in [('Original Scenario', original_data), ('Fin Scenario', fin_data)]:
    for _, row in data.iterrows():
        row_data = [scenario] + row[columns[1:]].tolist()
        if len(row_data) != len(columns):
            print(f"Row length mismatch in {scenario}: {len(row_data)} elements, expected {len(columns)}. Row: {row_data}")
            exit()
        all_data.append(row_data)

# Create DataFrame
try:
    df = pd.DataFrame(all_data, columns=columns)
except Exception as e:
    print(f"Error creating DataFrame: {e}")
    exit()

# Save to Excel and generate charts
try:
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Fin_Original', index=False)

        # Pivot data for charts
        pivot_fc_before = df.pivot_table(index='Ship Speed (knots)', columns='Scenario',
                                        values=['SVR FC Before MC', 'GPR FC Before MC', 'RF FC Before MC',
                                                'XGB FC Before MC', 'NN FC Before MC'])
        pivot_fc_after = df.pivot_table(index='Ship Speed (knots)', columns='Scenario',
                                       values=['SVR FC After MC', 'GPR FC After MC', 'RF FC After MC',
                                               'XGB FC After MC', 'NN FC After MC'])
        pivot_co2_before = df.pivot_table(index='Ship Speed (knots)', columns='Scenario',
                                         values=['SVR CO2 Before MC', 'GPR CO2 Before MC', 'RF CO2 Before MC',
                                                 'XGB CO2 Before MC', 'NN CO2 Before MC'])
        pivot_co2_after = df.pivot_table(index='Ship Speed (knots)', columns='Scenario',
                                        values=['SVR CO2 After MC', 'GPR CO2 After MC', 'RF CO2 After MC',
                                                'XGB CO2 After MC', 'NN CO2 After MC'])

        pivot_fc_before.to_excel(writer, sheet_name='Pivot_Data', startrow=0)
        pivot_fc_after.to_excel(writer, sheet_name='Pivot_Data', startrow=12)
        pivot_co2_before.to_excel(writer, sheet_name='Pivot_Data', startrow=24)
        pivot_co2_after.to_excel(writer, sheet_name='Pivot_Data', startrow=36)

        # Create charts
        workbook = writer.book
        worksheet = writer.sheets['Fin_Original']
        scenarios = ['Original Scenario', 'Fin Scenario']
        colors = ['#1F77B4', '#2CA02C']
        models = ['SVR', 'GPR', 'RF', 'XGB', 'NN']

        for model in models:
            # FC Before MC
            chart_fc_before = workbook.add_chart({'type': 'column'})
            for i, scenario in enumerate(scenarios):
                try:
                    col = pivot_fc_before.columns.get_loc((f'{model} FC Before MC', scenario)) + 2
                    chart_fc_before.add_series({
                        'name': scenario,
                        'categories': '=Pivot_Data!$A$2:$A$11',
                        'values': f'=Pivot_Data!${chr(65 + col)}$2:${chr(65 + col)}$11',
                        'fill': {'color': colors[i]},
                        'gap': 10
                    })
                except KeyError:
                    print(f"Column not found for {model} FC Before MC, {scenario}")
                    continue
            chart_fc_before.set_title({'name': f'{model} FC Predictions Before MC'})
            chart_fc_before.set_x_axis({'name': 'Ship Speed (knots)'})
            chart_fc_before.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}})
            chart_fc_before.set_size({'width': 720, 'height': 576})
            chart_fc_before.set_legend({'position': 'top'})
            worksheet.insert_chart(f'F{(models.index(model) * 18) + 2}', chart_fc_before)

            # FC After MC
            chart_fc_after = workbook.add_chart({'type': 'column'})
            for i, scenario in enumerate(scenarios):
                try:
                    col = pivot_fc_after.columns.get_loc((f'{model} FC After MC', scenario)) + 2
                    chart_fc_after.add_series({
                        'name': scenario,
                        'categories': '=Pivot_Data!$A$14:$A$23',
                        'values': f'=Pivot_Data!${chr(65 + col)}$14:${chr(65 + col)}$23',
                        'fill': {'color': colors[i]},
                        'gap': 10
                    })
                except KeyError:
                    print(f"Column not found for {model} FC After MC, {scenario}")
                    continue
            chart_fc_after.set_title({'name': f'{model} FC Predictions After MC'})
            chart_fc_after.set_x_axis({'name': 'Ship Speed (knots)'})
            chart_fc_after.set_y_axis({'name': 'Fuel Consumption (kg/h)', 'major_gridlines': {'visible': True}})
            chart_fc_after.set_size({'width': 720, 'height': 576})
            chart_fc_after.set_legend({'position': 'top'})
            worksheet.insert_chart(f'P{(models.index(model) * 18) + 2}', chart_fc_after)

            # CO2 Before MC
            chart_co2_before = workbook.add_chart({'type': 'column'})
            for i, scenario in enumerate(scenarios):
                try:
                    col = pivot_co2_before.columns.get_loc((f'{model} CO2 Before MC', scenario)) + 2
                    chart_co2_before.add_series({
                        'name': scenario,
                        'categories': '=Pivot_Data!$A$26:$A$35',
                        'values': f'=Pivot_Data!${chr(65 + col)}$26:${chr(65 + col)}$35',
                        'fill': {'color': colors[i]},
                        'gap': 10
                    })
                except KeyError:
                    print(f"Column not found for {model} CO2 Before MC, {scenario}")
                    continue
            chart_co2_before.set_title({'name': f'{model} CO2 Predictions Before MC'})
            chart_co2_before.set_x_axis({'name': 'Ship Speed (knots)'})
            chart_co2_before.set_y_axis({'name': 'CO2 Emission (kg)', 'major_gridlines': {'visible': True}})
            chart_co2_before.set_size({'width': 720, 'height': 576})
            chart_co2_before.set_legend({'position': 'top'})
            worksheet.insert_chart(f'F{(models.index(model) * 18) + 20}', chart_co2_before)

            # CO2 After MC
            chart_co2_after = workbook.add_chart({'type': 'column'})
            for i, scenario in enumerate(scenarios):
                try:
                    col = pivot_co2_after.columns.get_loc((f'{model} CO2 After MC', scenario)) + 2
                    chart_co2_after.add_series({
                        'name': scenario,
                        'categories': '=Pivot_Data!$A$38:$A$47',
                        'values': f'=Pivot_Data!${chr(65 + col)}$38:${chr(65 + col)}$47',
                        'fill': {'color': colors[i]},
                        'gap': 10
                    })
                except KeyError:
                    print(f"Column not found for {model} CO2 After MC, {scenario}")
                    continue
            chart_co2_after.set_title({'name': f'{model} CO2 Predictions After MC'})
            chart_co2_after.set_x_axis({'name': 'Ship Speed (knots)'})
            chart_co2_after.set_y_axis({'name': 'CO2 Emission (kg)', 'major_gridlines': {'visible': True}})
            chart_co2_after.set_size({'width': 720, 'height': 576})
            chart_co2_after.set_legend({'position': 'top'})
            worksheet.insert_chart(f'P{(models.index(model) * 18) + 20}', chart_co2_after)

    print(f"Comparison results saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")