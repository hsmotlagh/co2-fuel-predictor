import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

# Define input and output paths
input_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Combined_Scenario_Data.xlsx"
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/SVR_MSE_RMSE_Figures_14_15.xlsx"

# Define scenarios
scenarios = ['Original', 'Paint', 'Advance Propeller', 'Fin', 'Bulbous Bow', 'Combined']

# Data for Original Scenario
original_data = pd.DataFrame([
    [9, 115.7696693, 360.0436715, 212.5441426, 131.1508187, 916.5662006, 447.1187997],
    [10, 157.0962029, 488.5691911, 196.6363757, 158.6310184, 870.0368642, 526.2801538],
    [11, 211.3112914, 657.1781162, 216.8806122, 204.0135631, 859.0989233, 660.4106049],
    [12, 276.2684182, 859.1947807, 276.369045, 270.2385251, 929.6611844, 855.6961727],
    [12.5, 319.0039009, 992.1021319, 318.903584, 315.0517762, 1011.335582, 988.4495129],
    [13, 371.3694974, 1154.959137, 373.6568423, 371.653, 1135.725684, 1156.598379],
    [13.5, 434.4792596, 1351.230497, 434.378943, 439.5563096, 1255.635094, 1350.774585],
    [14, 511.7771533, 1591.626947, 474.5461671, 519.5712447, 1298.03187, 1566.882446],
    [14.5, 609.1134837, 1894.342934, 481.4836737, 608.8219948, 1259.176407, 1786.114488],
    [15, 814.8434842, 2534.163236, 477.5037548, 729.8122753, 1212.024875, 2007.914881]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
            'SVR FC Before MC', 'SVR FC After MC', 'SVR CO2 Before MC', 'SVR CO2 After MC'])

# Data for Paint Scenario
paint_data = pd.DataFrame([
    [9, 110.0681824, 342.3120472, 198.9500958, 122.5078835, 868.3139517, 406.610355],
    [10, 149.4348682, 464.7424402, 184.4212582, 148.252523, 821.9489526, 484.2087486],
    [11, 201.2853668, 625.9974908, 205.5268493, 191.8302572, 811.4365762, 618.5010343],
    [12, 263.4790258, 819.4197703, 263.5794429, 254.4450196, 882.4097629, 811.3285214],
    [12.5, 304.5852633, 947.2601688, 304.4850525, 297.0075927, 964.1036723, 941.239089],
    [13, 354.9289312, 1103.828976, 356.7247847, 352.766029, 1086.98547, 1106.209366],
    [13.5, 415.9640576, 1293.648219, 415.8638479, 423.01583, 1207.031549, 1302.449508],
    [14, 490.9831319, 1526.95754, 456.1786509, 503.15546, 1249.671125, 1515.700765],
    [14.5, 585.731981, 1821.626461, 464.3280106, 594.70329, 1210.609411, 1734.062779],
    [15, 787.8448646, 2450.197529, 462.2979811, 713.9108748, 1164.175625, 1945.897448]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
            'SVR FC Before MC', 'SVR FC After MC', 'SVR CO2 Before MC', 'SVR CO2 After MC'])

# Data for Advance Propeller Scenario
adv_propeller_data = pd.DataFrame([
    [9, 103.7743379, 322.738191, 180.6946793, 119.1119081, 805.8991647, 380.0966852],
    [10, 140.8707293, 438.107968, 168.2066289, 144.3497792, 759.4009522, 453.8118531],
    [11, 189.5548816, 589.5156817, 191.1641988, 184.6487631, 748.5982487, 578.4961849],
    [12, 247.890597, 770.9397568, 247.99085, 244.1215673, 819.7103562, 763.9002434],
    [12.5, 286.2080063, 890.1068997, 286.1075057, 283.931854, 901.5164683, 887.8096963],
    [13, 332.9373352, 1035.435112, 333.8740144, 333.1273451, 1024.025543, 1039.677876],
    [13.5, 389.4472384, 1211.180911, 389.347495, 393.1309773, 1143.887905, 1217.755984],
    [14, 458.6313585, 1426.343525, 429.4692586, 464.5089952, 1186.499088, 1415.87227],
    [14.5, 545.6961681, 1697.115083, 440.675173, 544.9027781, 1147.827106, 1616.591779],
    [15, 729.3067193, 2268.143897, 441.2868586, 652.5308001, 1101.041104, 1810.595579]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
            'SVR FC Before MC', 'SVR FC After MC', 'SVR CO2 Before MC', 'SVR CO2 After MC'])

# Data for Fin Scenario
fin_data = pd.DataFrame([
    [9, 115.4654418, 359.097524, 212.0958908, 132.2985703, 913.770889, 432.2445051],
    [10, 156.6860586, 487.2936421, 196.2611189, 159.6467296, 867.2527734, 516.3940987],
    [11, 210.7942821, 655.5702172, 216.4905454, 203.8605224, 856.3071276, 657.270079],
    [12, 275.6058519, 857.1341993, 275.7063009, 270.2442762, 927.0795856, 859.667355],
    [12.5, 318.2599935, 989.7885799, 318.1597652, 314.9616482, 1009.083118, 994.5758689],
    [13, 370.516624, 1152.306701, 372.8123109, 370.4364224, 1133.012162, 1161.800517],
    [13.5, 433.5948752, 1348.480062, 433.4946492, 437.6316248, 1252.54077, 1354.98607],
    [14, 510.8733655, 1588.816167, 473.8840309, 518.1438336, 1295.051641, 1565.763989],
    [14.5, 608.260007, 1891.688622, 480.8145243, 608.4796762, 1256.27127, 1776.293066],
    [15, 814.3657819, 2532.677582, 476.842285, 728.9895, 1209.179179, 1972.527355]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
            'SVR FC Before MC', 'SVR FC After MC', 'SVR CO2 Before MC', 'SVR CO2 After MC'])

# Data for Bulbous Bow Scenario
bulbousbow_data = pd.DataFrame([
    [9, 119.6541899, 372.1245307, 216.4660808, 132.077147, 918.9686993, 443.1769183],
    [10, 162.3385399, 504.8728592, 203.1716661, 162.3196739, 874.1118535, 528.523451],
    [11, 217.9322612, 677.7693324, 225.4047631, 210.0158736, 868.1678728, 673.5510916],
    [12, 284.2566267, 884.038109, 284.3572593, 277.4218541, 946.8520474, 874.8795296],
    [12.5, 323.9702129, 1007.547362, 323.8698972, 319.6531282, 1025.585585, 1000.764194],
    [13, 371.0228762, 1153.881145, 372.5251372, 371.2404083, 1135.842929, 1156.766912],
    [13.5, 425.0311648, 1321.846922, 424.930848, 431.3176299, 1241.929406, 1333.615466],
    [14, 501.2150135, 1558.778692, 462.0915107, 509.1874479, 1287.143545, 1544.013479],
    [14.5, 600.5953901, 1867.851663, 463.1069245, 602.6333002, 1245.099951, 1769.859126],
    [15, 863.1846317, 2684.504205, 472.5114681, 757.5131097, 1207.230073, 2016.107919]
], columns=['Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
            'SVR FC Before MC', 'SVR FC After MC', 'SVR CO2 Before MC', 'SVR CO2 After MC'])

# Read Combined Scenario results
try:
    combined_data = pd.read_excel(input_path, sheet_name='Combined_Scen_Coeff')
except Exception as e:
    print(f"Error reading Combined Scenario data: {e}")
    exit()

# Verify required columns in Combined Scenario data
required_cols = [
    'Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)',
    'Final Predicted Fuel Consumption SVR Before MC (kg/h)', 'Final Predicted Fuel Consumption SVR After MC (kg/h)',
    'Final Predicted CO2 Emission SVR Before MC (kg)', 'Final Predicted CO2 Emission SVR After MC (kg)'
]
missing_cols = [col for col in required_cols if col not in combined_data.columns]
if missing_cols:
    print(f"Missing columns in Combined Scenario data: {missing_cols}")
    exit()


# Function to calculate MSE
def calculate_mse(actual, predicted):
    return np.mean((actual - predicted) ** 2)


# Function to calculate RMSE
def calculate_rmse(mse):
    return np.sqrt(mse)


# Calculate MSE and RMSE for all scenarios
results = []
for scenario, data in [
    ('Original', original_data),
    ('Paint', paint_data),
    ('Advance Propeller', adv_propeller_data),
    ('Fin', fin_data),
    ('Bulbous Bow', bulbousbow_data),
    ('Combined', combined_data)
]:
    if scenario == 'Combined':
        mse_fc_before = calculate_mse(data['Fuel Consumption (FC) (kg/h)'],
                                      data['Final Predicted Fuel Consumption SVR Before MC (kg/h)'])
        mse_fc_after = calculate_mse(data['Fuel Consumption (FC) (kg/h)'],
                                     data['Final Predicted Fuel Consumption SVR After MC (kg/h)'])
        mse_co2_before = calculate_mse(data['CO2 Emission (kg)'],
                                       data['Final Predicted CO2 Emission SVR Before MC (kg)'])
        mse_co2_after = calculate_mse(data['CO2 Emission (kg)'], data['Final Predicted CO2 Emission SVR After MC (kg)'])
    else:
        mse_fc_before = calculate_mse(data['Fuel Consumption (FC) (kg/h)'], data['SVR FC Before MC'])
        mse_fc_after = calculate_mse(data['Fuel Consumption (FC) (kg/h)'], data['SVR FC After MC'])
        mse_co2_before = calculate_mse(data['CO2 Emission (kg)'], data['SVR CO2 Before MC'])
        mse_co2_after = calculate_mse(data['CO2 Emission (kg)'], data['SVR CO2 After MC'])

    rmse_fc_before = calculate_rmse(mse_fc_before)
    rmse_fc_after = calculate_rmse(mse_fc_after)
    rmse_co2_before = calculate_rmse(mse_co2_before)
    rmse_co2_after = calculate_rmse(mse_co2_after)

    results.append({
        'Scenario': scenario,
        'MSE_FC_Before_MC': mse_fc_before,
        'MSE_FC_After_MC': mse_fc_after,
        'MSE_CO2_Before_MC': mse_co2_before,
        'MSE_CO2_After_MC': mse_co2_after,
        'RMSE_FC_Before_MC': rmse_fc_before,
        'RMSE_FC_After_MC': rmse_fc_after,
        'RMSE_CO2_Before_MC': rmse_co2_before,
        'RMSE_CO2_After_MC': rmse_co2_after
    })

# Create DataFrame for results
results_df = pd.DataFrame(results)

# Save to Excel and generate charts
try:
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        results_df.to_excel(writer, sheet_name='MSE_RMSE', index=False)

        workbook = writer.book
        worksheet = writer.sheets['MSE_RMSE']

        # Figure 14: MSE Performance of SVR Across Scenarios FC
        chart_fig14 = workbook.add_chart({'type': 'bar'})
        for i, col in enumerate(['MSE_FC_Before_MC', 'MSE_FC_After_MC']):
            chart_fig14.add_series({
                'name': 'Before MC' if i == 0 else 'After MC',
                'categories': '=MSE_RMSE!$A$2:$A$7',
                'values': f'=MSE_RMSE!${chr(66 + i)}$2:${chr(66 + i)}$7',
                'fill': {'color': '#1F77B4' if i == 0 else '#FF7F0E'},
                'gap': 20
            })
        chart_fig14.set_title({'name': 'Figure 14: MSE Performance of SVR Across Scenarios FC'})
        chart_fig14.set_x_axis({'name': 'MSE (kg/h)²'})
        chart_fig14.set_y_axis({'name': 'Scenario', 'major_gridlines': {'visible': True}})
        chart_fig14.set_size({'width': 720, 'height': 576})
        chart_fig14.set_legend({'position': 'top'})
        worksheet.insert_chart('K2', chart_fig14)

        # Figure 15: RMSE Performance of SVR Across Scenarios for CO2 Emissions
        chart_fig15 = workbook.add_chart({'type': 'bar'})
        for i, col in enumerate(['RMSE_CO2_Before_MC', 'RMSE_CO2_After_MC']):
            chart_fig15.add_series({
                'name': 'Before MC' if i == 0 else 'After MC',
                'categories': '=MSE_RMSE!$A$2:$A$7',
                'values': f'=MSE_RMSE!${chr(70 + i)}$2:${chr(70 + i)}$7',
                'fill': {'color': '#1F77B4' if i == 0 else '#FF7F0E'},
                'gap': 20
            })
        chart_fig15.set_title({'name': 'Figure 15: RMSE Performance of SVR Across Scenarios for CO₂ Emissions'})
        chart_fig15.set_x_axis({'name': 'RMSE (kg)'})
        chart_fig15.set_y_axis({'name': 'Scenario', 'major_gridlines': {'visible': True}})
        chart_fig15.set_size({'width': 720, 'height': 576})
        chart_fig15.set_legend({'position': 'top'})
        worksheet.insert_chart('K22', chart_fig15)

    print(f"MSE and RMSE results with Figures 14 and 15 saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")