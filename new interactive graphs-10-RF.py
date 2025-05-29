import pandas as pd
import numpy as np
import os

# Complete data for RF for all scenarios
data_rf = {
    "Scenario": [
        "Original Scenar_Coeff",
        "Paint (5%) Scen_Coeff",
        "Advance Propell_Coeff",
        "Fin (2%-4%) Sce_Coeff",
        "Bulbous Bow Sce_Coeff"
    ],
    "Ship_Speed": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
    "Original_FC": [
        [115.77, 157.10, 211.31, 276.27, 319.00, 371.37, 434.48, 511.78, 609.11, 814.84],
        [110.07, 149.43, 201.29, 263.48, 304.59, 354.93, 415.96, 490.98, 585.73, 787.84],
        [103.77, 140.87, 189.55, 247.89, 286.21, 332.94, 389.45, 458.63, 545.70, 729.31],
        [115.47, 156.69, 210.79, 275.61, 318.26, 370.52, 433.59, 510.87, 608.26, 814.37],
        [119.65, 162.34, 217.93, 284.26, 323.97, 371.02, 425.03, 501.22, 600.60, 863.18],
    ],
    "FC_Before_MC": [
        [153.87, 157.17, 229.84, 289.23, 314.20, 362.36, 427.66, 478.52, 545.03, 707.10],
        [145.97, 147.79, 217.93, 275.96, 299.97, 342.63, 410.76, 458.73, 509.74, 682.38],
        [137.51, 138.37, 205.52, 259.50, 279.01, 323.36, 385.85, 428.86, 482.88, 633.07],
        [153.81, 154.76, 228.96, 288.54, 310.27, 355.07, 425.38, 477.63, 545.61, 706.49],
        [159.08, 164.98, 237.73, 297.32, 316.22, 359.85, 412.62, 466.17, 527.85, 736.12],
    ],
    "FC_After_MC": [
        [117.42, 156.68, 211.31, 276.27, 318.25, 366.97, 434.62, 512.18, 609.11, 812.79],
        [112.04, 149.04, 200.77, 263.48, 304.36, 352.12, 414.66, 492.13, 585.04, 787.84],
        [104.89, 140.50, 189.07, 248.66, 285.61, 330.23, 389.70, 461.24, 547.53, 727.47],
        [116.29, 156.69, 210.79, 277.74, 316.98, 367.60, 434.65, 506.24, 612.38, 812.30],
        [120.51, 162.34, 217.93, 283.59, 323.65, 368.41, 420.78, 500.45, 602.23, 860.56],
    ],
    "Original_CO2": [
        [360.04, 488.57, 657.18, 859.19, 992.10, 1154.96, 1351.23, 1591.63, 1894.34, 2534.16],
        [342.31, 464.74, 625.99, 819.42, 947.26, 1103.83, 1293.65, 1526.96, 1821.63, 2450.20],
        [322.74, 438.11, 589.52, 770.94, 890.11, 1035.44, 1211.18, 1426.34, 1697.12, 2268.14],
        [359.10, 487.29, 655.57, 857.13, 989.79, 1152.31, 1348.48, 1588.82, 1891.69, 2532.68],
        [372.12, 504.87, 677.77, 884.04, 1007.55, 1153.88, 1321.85, 1558.78, 1867.85, 2684.50],
    ],
    "CO2_Before_MC": [
        [479.56, 491.45, 719.12, 899.51, 972.18, 1113.76, 1325.65, 1488.20, 1713.65, 2199.09],
        [453.97, 459.64, 676.49, 858.25, 932.90, 1067.09, 1268.99, 1426.65, 1621.45, 2122.20],
        [427.65, 443.66, 646.02, 807.05, 870.43, 993.98, 1192.19, 1333.77, 1522.90, 1968.83],
        [477.03, 479.99, 717.38, 897.37, 961.02, 1103.59, 1340.38, 1485.42, 1694.88, 2197.17],
        [499.12, 514.40, 741.02, 924.68, 992.97, 1126.81, 1292.82, 1459.27, 1608.56, 2289.32],
    ],
    "CO2_After_MC": [
        [365.18, 487.28, 657.18, 859.19, 991.07, 1150.74, 1356.48, 1586.82, 1904.11, 2527.77],
        [345.98, 464.74, 626.00, 820.70, 946.27, 1087.23, 1287.26, 1533.46, 1825.36, 2450.20],
        [325.05, 438.11, 589.52, 770.94, 891.82, 1023.81, 1210.21, 1424.19, 1697.12, 2268.14],
        [361.66, 487.29, 653.89, 861.75, 983.16, 1154.33, 1354.61, 1592.47, 1885.63, 2532.68],
        [374.78, 504.87, 677.77, 885.27, 1006.31, 1142.61, 1315.82, 1556.41, 1881.09, 2684.50],
    ],
}

# Convert data to DataFrame
scenarios = data_rf["Scenario"]

mse_results = []
rmse_results = []

# Calculating MSE and RMSE for FC and CO2 across all scenarios
for idx, scenario in enumerate(scenarios):
    mse_fc_before_mc = np.mean([(o - p) ** 2 for o, p in zip(data_rf["Original_FC"][idx], data_rf["FC_Before_MC"][idx])])
    rmse_fc_before_mc = np.sqrt(mse_fc_before_mc)

    mse_fc_after_mc = np.mean([(o - p) ** 2 for o, p in zip(data_rf["Original_FC"][idx], data_rf["FC_After_MC"][idx])])
    rmse_fc_after_mc = np.sqrt(mse_fc_after_mc)

    mse_co2_before_mc = np.mean([(o - p) ** 2 for o, p in zip(data_rf["Original_CO2"][idx], data_rf["CO2_Before_MC"][idx])])
    rmse_co2_before_mc = np.sqrt(mse_co2_before_mc)

    mse_co2_after_mc = np.mean([(o - p) ** 2 for o, p in zip(data_rf["Original_CO2"][idx], data_rf["CO2_After_MC"][idx])])
    rmse_co2_after_mc = np.sqrt(mse_co2_after_mc)

    mse_results.append([scenario, mse_fc_before_mc, mse_fc_after_mc, mse_co2_before_mc, mse_co2_after_mc])
    rmse_results.append([scenario, rmse_fc_before_mc, rmse_fc_after_mc, rmse_co2_before_mc, rmse_co2_after_mc])

# Create DataFrames for MSE and RMSE
columns = ["Scenario", "MSE_FC_Before_MC", "MSE_FC_After_MC", "MSE_CO2_Before_MC", "MSE_CO2_After_MC"]
mse_df = pd.DataFrame(mse_results, columns=columns)

columns = ["Scenario", "RMSE_FC_Before_MC", "RMSE_FC_After_MC", "RMSE_CO2_Before_MC", "RMSE_CO2_After_MC"]
rmse_df = pd.DataFrame(rmse_results, columns=columns)

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML-RESULT/RF_MSE_RMSE_Results_Updated.xlsx"
if not os.path.exists(os.path.dirname(output_path)):
    os.makedirs(os.path.dirname(output_path))

# Save to Excel with updated chart configurations using XlsxWriter
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    mse_df.to_excel(writer, sheet_name='MSE_Results', index=False)
    rmse_df.to_excel(writer, sheet_name='RMSE_Results', index=False)
    workbook = writer.book
    mse_sheet = writer.sheets['MSE_Results']
    rmse_sheet = writer.sheets['RMSE_Results']

    def add_bar_chart(sheet, data_df, title, cell_location, sheet_name):
        # Get the number of rows
        num_rows = len(data_df)
        # Define data ranges
        categories = f"='{sheet_name}'!$A$2:$A${num_rows + 1}"
        values1 = f"='{sheet_name}'!$B$2:$B${num_rows + 1}"
        values2 = f"='{sheet_name}'!$C$2:$C${num_rows + 1}"

        chart = workbook.add_chart({'type': 'column'})
        chart.add_series({
            'name':       data_df.columns[1],
            'categories': categories,
            'values':     values1,
        })
        chart.add_series({
            'name':       data_df.columns[2],
            'categories': categories,
            'values':     values2,
        })
        chart.set_title({'name': title})
        chart.set_x_axis({'name': 'Scenarios', 'label_position': 'low'})
        chart.set_y_axis({'name': 'MSE' if 'MSE' in title else 'RMSE'})
        chart.set_legend({'position': 'bottom'})

        # Set the size of the chart
        chart.set_size({'width': 720, 'height': 576})

        # Insert the chart into the worksheet
        sheet.insert_chart(cell_location, chart)

    # Add charts to the MSE sheet
    add_bar_chart(mse_sheet, mse_df[["Scenario", "MSE_FC_Before_MC", "MSE_FC_After_MC"]],
                  "MSE for Fuel Consumption (FC)", 'E2', 'MSE_Results')
    add_bar_chart(mse_sheet, mse_df[["Scenario", "MSE_CO2_Before_MC", "MSE_CO2_After_MC"]],
                  "MSE for CO2 Emission", 'E20', 'MSE_Results')

    # Add charts to the RMSE sheet
    add_bar_chart(rmse_sheet, rmse_df[["Scenario", "RMSE_FC_Before_MC", "RMSE_FC_After_MC"]],
                  "RMSE for Fuel Consumption (FC)", 'E2', 'RMSE_Results')
    add_bar_chart(rmse_sheet, rmse_df[["Scenario", "RMSE_CO2_Before_MC", "RMSE_CO2_After_MC"]],
                  "RMSE for CO2 Emission", 'E20', 'RMSE_Results')
