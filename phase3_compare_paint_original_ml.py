import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define input and output paths
input_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Combined_Scenario_Data.xlsx"
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/ML_Paint_Original_Comparison.xlsx"

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

# Data for Paint Scenario (verified to match corrected values provided on May 21, 2025)
paint_data = pd.DataFrame([
    [9, 110.0681824, 342.3120472, 198.9500958, 868.3139517, 122.5078835, 406.610355, 110.0681824, 342.3120472, 110.9468804, 345.827258, 145.9697979, 453.9660716, 112.0365167, 345.984959, 110.0694275, 342.3130493, 110.0686035, 342.3125305, 109.7261409, 341.3546769, 104.3293647, 369.3603125],
    [10, 149.4348682, 464.7424402, 184.4212582, 821.9489526, 148.252523, 484.2087486, 151.1317976, 470.3726079, 150.4267954, 469.3596164, 147.7941416, 459.6397804, 149.0412014, 464.7424402, 110.0694275, 342.3130493, 149.434967, 464.7426147, 136.466916, 417.9267183, 159.4908743, 472.5635381],
    [11, 201.2853668, 625.9974908, 205.5268493, 811.4365762, 191.8302572, 618.5010343, 201.2853668, 625.9974908, 202.1326087, 630.0039312, 217.9306678, 676.4859729, 200.7668618, 625.9974908, 201.2854462, 625.9974976, 201.2854309, 625.9976196, 190.9735841, 585.9480329, 197.0763132, 579.7292394],
    [12, 263.4790258, 819.4197703, 263.5794429, 882.4097629, 254.4450196, 811.3285214, 263.4790258, 819.4197703, 264.4943841, 823.901072, 275.9635875, 858.2467571, 263.4790258, 820.6981743, 263.4786377, 819.4193726, 263.479126, 819.4199219, 280.0992688, 873.3001888, 237.5931262, 673.9866465],
    [12.5, 304.5852633, 947.2601688, 304.4850525, 964.1036723, 297.0075927, 941.239089, 304.5852633, 947.2601688, 305.9407542, 952.9826716, 299.9679622, 932.9003626, 304.3589495, 946.2690489, 304.5857544, 947.260498, 304.5852661, 947.2601929, 305.8465382, 957.7176548, 267.2496331, 753.8638453],
    [13, 354.9289312, 1103.828976, 356.7247847, 1086.98547, 352.766029, 1106.209366, 354.9289312, 1103.828976, 356.5729771, 1110.169015, 342.6251478, 1067.093867, 352.1221403, 1087.226196, 354.9288025, 1103.829102, 354.928894, 1103.828979, 347.0132561, 1078.978708, 328.4503802, 944.3038956],
    [13.5, 415.9640576, 1293.648219, 415.8638479, 1207.031549, 423.01583, 1302.449508, 415.9640576, 1293.648219, 417.0407844, 1295.161495, 410.7555031, 1268.987043, 414.6594368, 1287.257755, 415.9638672, 1293.648193, 415.9640808, 1293.648193, 411.2245197, 1271.538847, 417.1987376, 1196.729475],
    [14, 490.9831319, 1526.95754, 456.1786509, 1249.671125, 503.15546, 1515.700765, 490.9831319, 1526.95754, 491.9418897, 1527.291712, 458.7298578, 1426.649858, 492.1279181, 1533.464514, 490.9830627, 1526.957642, 490.9830322, 1526.95752, 493.1734249, 1531.564039, 493.8534728, 1480.261427],
    [14.5, 585.731981, 1821.626461, 464.3280106, 1210.609411, 594.70329, 1734.062779, 604.0632426, 1880.573319, 581.251839, 1799.758685, 509.7395255, 1621.449695, 585.0367962, 1825.357815, 490.9830627, 1526.957642, 585.7319946, 1821.626343, 603.0627193, 1871.194434, 587.7390545, 1836.58612],
    [15, 787.8448646, 2450.197529, 462.2979811, 1164.175625, 713.9108748, 1945.897448, 787.8448646, 2450.197529, 772.4711985, 2402.145048, 682.3785854, 2122.1974, 787.8448646, 2450.197529, 787.843689, 2450.196289, 787.8445435, 2450.197266, 788.3603698, 2454.717365, 803.6422616, 2465.61085]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
            'SVR FC Before MC', 'SVR CO2 Before MC', 'SVR FC After MC', 'SVR CO2 After MC',
            'GPR FC Before MC', 'GPR CO2 Before MC', 'GPR FC After MC', 'GPR CO2 After MC',
            'RF FC Before MC', 'RF CO2 Before MC', 'RF FC After MC', 'RF CO2 After MC',
            'XGB FC Before MC', 'XGB CO2 Before MC', 'XGB FC After MC', 'XGB CO2 After MC',
            'NN FC Before MC', 'NN CO2 Before MC', 'NN FC After MC', 'NN CO2 After MC'])

# Combine scenario data
all_data = []
for scenario, data in [('Original Scenario', original_data), ('Paint Scenario', paint_data)]:
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
        df.to_excel(writer, sheet_name='Paint_Original', index=False)

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
        pivot_co2_after.to_excel(writer, sheet_name='Pipeline', startrow=36)

        # Create charts
        workbook = writer.book
        worksheet = writer.sheets['Paint_Original']
        scenarios = ['Original Scenario', 'Paint Scenario']
        colors = ['#1F77B4', '#FF7F0E']
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
                        'categories': '=Pipeline!$A$38:$A$47',
                        'values': f'=Pipeline!${chr(65 + col)}$38:${chr(65 + col)}$47',
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