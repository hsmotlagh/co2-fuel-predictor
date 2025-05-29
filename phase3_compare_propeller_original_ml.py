import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/ML_Propeller_Original_Comparison.xlsx"

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

# Data for Advance Propeller Scenario (verified to match corrected values provided on May 21, 2025)
adv_propeller_data = pd.DataFrame([
    [9, 103.7743379, 322.738191, 180.6946793, 805.8991647, 119.1119081, 380.0966852, 103.7743379, 322.738191, 103.5641821, 321.9157786, 137.5077847, 427.6492104, 104.8872297, 325.0455865, 103.7754974, 322.7393188, 103.7748413, 322.7387085, 103.3978574, 322.0538208, 106.4254376, 377.2094133],
    [10, 140.8707293, 438.107968, 168.2066289, 759.4009522, 144.3497792, 453.8118531, 142.4288852, 443.3356238, 141.8556822, 441.4072398, 138.3655901, 443.6558599, 140.4997653, 438.107968, 103.7754974, 322.7393188, 140.8708038, 438.1080627, 129.8300881, 397.590206, 141.2367734, 459.8997054],
    [11, 189.5548816, 589.5156817, 191.1641988, 748.5982487, 184.6487631, 578.4961849, 189.5548816, 589.5156817, 191.9589791, 598.0201232, 205.5157148, 646.0192317, 189.0680401, 589.5156817, 189.5549316, 589.515625, 189.5549469, 589.5157471, 180.2472957, 552.0813686, 175.16391, 584.7836301],
    [12, 247.890597, 770.9397568, 247.99085, 819.7103562, 244.1215673, 763.9002434, 247.890597, 770.9397568, 252.2573825, 786.7087443, 259.5024821, 807.0527194, 248.6569452, 770.9397568, 247.8903198, 770.93927, 247.8906403, 770.9396973, 263.0853598, 820.2599773, 222.064907, 705.8975354],
    [12.5, 286.2080063, 890.1068997, 286.1075057, 901.5164683, 283.931854, 887.8096963, 286.2080063, 890.1068997, 291.4768664, 909.3523352, 279.0127234, 870.4313061, 285.6098965, 891.8217925, 286.2084656, 890.1072998, 286.2080383, 890.1069336, 287.0265893, 899.457106, 244.3001162, 736.7144522],
    [13, 332.9373352, 1035.435112, 333.8740144, 1024.025543, 333.1273451, 1039.677876, 332.9373352, 1035.435112, 338.0240017, 1054.286444, 323.3617529, 993.9776255, 330.2313812, 1023.808855, 332.9371948, 1035.434937, 332.9372864, 1035.435059, 326.349306, 1015.526764, 310.9270433, 942.737159],
    [13.5, 389.4472384, 1211.180911, 389.347495, 1143.887905, 393.1309773, 1217.755984, 389.4472384, 1211.180911, 391.1423315, 1217.186155, 385.8533167, 1192.185647, 389.7007227, 1210.21179, 389.447113, 1211.180786, 389.4472046, 1211.180908, 384.3666688, 1189.890251, 403.1325328, 1220.971881],
    [14, 458.6313585, 1426.343525, 429.4692586, 1186.499088, 464.5089952, 1415.87227, 458.6313585, 1426.343525, 456.4251026, 1416.922151, 428.8637319, 1333.766206, 461.2433028, 1424.191899, 458.6312866, 1426.343506, 458.6312866, 1426.343384, 460.9869834, 1430.125371, 482.192159, 1464.036762],
    [14.5, 545.6961681, 1697.115083, 440.675173, 1147.827106, 544.9027781, 1616.591779, 562.0957495, 1750.226352, 529.980596, 1638.163994, 482.8819581, 1522.902149, 547.5322736, 1697.115083, 458.6312866, 1426.343506, 545.696106, 1697.11499, 559.478636, 1735.170489, 569.6322537, 1759.253633],
    [15, 729.3067193, 2268.143897, 441.2868586, 1101.041104, 652.5308001, 1810.595579, 729.3067193, 2268.143897, 638.4584508, 1949.719923, 633.0654382, 1968.833513, 727.4706137, 2268.143897, 729.305542, 2268.142578, 729.3063965, 2268.143555, 729.7857526, 2272.424563, 776.2036733, 2380.76935]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
            'SVR FC Before MC', 'SVR CO2 Before MC', 'SVR FC After MC', 'SVR CO2 After MC',
            'GPR FC Before MC', 'GPR CO2 Before MC', 'GPR FC After MC', 'GPR CO2 After MC',
            'RF FC Before MC', 'RF CO2 Before MC', 'RF FC After MC', 'RF CO2 After MC',
            'XGB FC Before MC', 'XGB CO2 Before MC', 'XGB FC After MC', 'XGB CO2 After MC',
            'NN FC Before MC', 'NN CO2 Before MC', 'NN FC After MC', 'NN CO2 After MC'])

# Combine scenario data
all_data = []
for scenario, data in [('Original Scenario', original_data), ('Advance Propeller Scenario', adv_propeller_data)]:
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
        df.to_excel(writer, sheet_name='Propeller_Original', index=False)

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
        worksheet = writer.sheets['Propeller_Original']
        scenarios = ['Original Scenario', 'Advance Propeller Scenario']
        colors = ['#1F77B4', '#D62728']
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