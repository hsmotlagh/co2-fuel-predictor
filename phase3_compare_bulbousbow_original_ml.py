import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/ML_BulbousBow_Original_Comparison.xlsx"

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

# Data for Bulbous Bow Scenario (verified to match corrected values provided on May 21, 2025)
bulbousbow_data = pd.DataFrame([
    [9, 119.6541899, 372.1245307, 216.4660808, 918.9686993, 132.077147, 443.1769183, 119.6541899, 372.1245307, 120.1282541, 375.5055149, 159.0816657, 499.1220685, 120.5078769, 374.7794972, 119.6553192, 372.1254883, 119.6546326, 372.125061, 120.3363807, 374.0680898, 105.2021918, 399.7885449],
    [10, 162.3385399, 504.8728592, 203.1716661, 874.1118535, 162.3196739, 528.523451, 162.9791912, 507.1562235, 163.0493357, 509.4299438, 164.97835, 514.4043086, 162.3385399, 504.8728592, 119.6553192, 372.1254883, 162.3386383, 504.8729858, 187.0887422, 579.5929166, 166.0555277, 577.9062393],
    [11, 217.9322612, 677.7693324, 225.4047631, 868.1678728, 210.0158736, 673.5510916, 217.9322612, 677.7693324, 218.3486253, 681.7203112, 237.7285392, 741.0154147, 217.9322612, 677.7693324, 217.9316406, 677.7685547, 217.9323578, 677.7694092, 218.422111, 679.0384899, 206.7959789, 708.5779733],
    [12, 284.2566267, 884.038109, 284.3572593, 946.8520474, 277.4218541, 874.8795296, 284.2566267, 884.038109, 284.667013, 887.3271867, 297.3238248, 924.6770952, 283.593383, 885.2732015, 284.2573547, 884.0386963, 284.2566528, 884.0380859, 288.5116364, 897.1619963, 278.1867048, 857.9204903],
    [12.5, 323.9702129, 1007.547362, 323.8698972, 1025.585585, 319.6531282, 1000.764194, 323.9702129, 1007.547362, 324.6184827, 1011.524823, 316.2215187, 992.9711596, 323.6464678, 1006.31227, 323.9700623, 1007.547058, 323.9702148, 1007.547363, 317.7140773, 987.6864396, 304.7972329, 909.5844777],
    [13, 371.0228762, 1153.881145, 372.5251372, 1135.842929, 371.2404083, 1156.766912, 371.0228762, 1153.881145, 372.5009879, 1161.723442, 359.8512847, 1126.813364, 368.4083852, 1142.607082, 371.0230103, 1153.881348, 371.0228577, 1153.881226, 367.8551809, 1141.334013, 349.7062282, 1074.050383],
    [13.5, 425.0311648, 1321.846922, 424.930848, 1241.929406, 431.3176299, 1333.615466, 425.0311648, 1321.846922, 428.9075528, 1340.881052, 412.6247838, 1292.824366, 420.7800579, 1315.817951, 425.0310669, 1321.846924, 425.031189, 1321.846924, 432.4448148, 1348.230195, 435.4264877, 1352.764498],
    [14, 501.2150135, 1558.778692, 462.0915107, 1287.143545, 509.1874479, 1544.013479, 501.2150135, 1558.778692, 509.8519426, 1597.165493, 466.1656241, 1459.272315, 500.453175, 1556.409374, 501.2149658, 1558.778809, 501.2149353, 1558.778564, 499.1217642, 1549.979887, 489.9020399, 1570.165135],
    [14.5, 600.5953901, 1867.851663, 463.1069245, 1245.099951, 602.6333002, 1769.859126, 615.8438274, 1916.745181, 609.5364334, 1904.301113, 527.8524522, 1608.558974, 602.2274788, 1881.093984, 501.2149658, 1558.778809, 600.5953979, 1867.851563, 577.0773365, 1788.746498, 608.6858274, 1861.606628],
    [15, 863.1846317, 2684.504205, 472.5114681, 1207.230073, 757.5131097, 2016.107919, 863.1846317, 2684.504205, 812.8306153, 2487.647371, 736.1162728, 2289.321608, 860.5587393, 2684.504205, 863.1834717, 2684.50293, 863.1843262, 2684.503906, 862.7095994, 2683.686106, 889.4501571, 2807.182587]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
            'SVR FC Before MC', 'SVR CO2 Before MC', 'SVR FC After MC', 'SVR CO2 After MC',
            'GPR FC Before MC', 'GPR CO2 Before MC', 'GPR FC After MC', 'GPR CO2 After MC',
            'RF FC Before MC', 'RF CO2 Before MC', 'RF FC After MC', 'RF CO2 After MC',
            'XGB FC Before MC', 'XGB CO2 Before MC', 'XGB FC After MC', 'XGB CO2 After MC',
            'NN FC Before MC', 'NN CO2 Before MC', 'NN FC After MC', 'NN CO2 After MC'])

# Combine scenario data
all_data = []
for scenario, data in [('Original Scenario', original_data), ('Bulbous Bow Scenario', bulbousbow_data)]:
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
        df.to_excel(writer, sheet_name='BulbousBow_Original', index=False)

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
        worksheet = writer.sheets['BulbousBow_Original']
        scenarios = ['Original Scenario', 'Bulbous Bow Scenario']
        colors = ['#1F77B4', '#9467BD']
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