import pandas as pd
import numpy as np
import os

# Complete data for NN for all scenarios
data_nn = {
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
        [117.48, 157.85, 204.08, 283.97, 316.25, 375.89, 431.92, 509.20, 608.40, 815.04],
        [109.73, 136.47, 190.97, 280.10, 305.85, 347.01, 411.22, 493.17, 603.06, 788.36],
        [103.40, 129.83, 180.25, 263.09, 287.03, 326.35, 384.37, 460.99, 559.48, 729.79],
        [115.08, 140.73, 199.80, 293.65, 319.61, 361.35, 428.87, 513.29, 627.70, 814.87],
        [120.34, 187.09, 218.42, 288.51, 317.71, 367.86, 432.44, 499.12, 577.08, 862.71],
    ],
    "FC_After_MC": [
        [128.84, 179.68, 213.68, 265.85, 282.16, 353.85, 438.64, 519.05, 611.01, 838.81],
        [97.73, 158.41, 209.97, 267.20, 300.96, 354.45, 406.52, 472.89, 592.71, 776.50],
        [93.38, 132.24, 172.23, 235.49, 265.77, 324.34, 371.57, 443.85, 532.31, 730.76],
        [118.51, 164.21, 210.48, 271.76, 297.64, 363.14, 436.37, 512.93, 614.66, 820.08],
        [102.42, 167.46, 213.59, 255.23, 276.38, 340.86, 445.29, 532.33, 608.73, 883.14],
    ],
    "Original_CO2": [
        [360.04, 488.57, 657.18, 859.19, 992.10, 1154.96, 1351.23, 1591.63, 1894.34, 2534.16],
        [342.31, 464.74, 625.99, 819.42, 947.26, 1103.83, 1293.65, 1526.96, 1821.63, 2450.20],
        [322.74, 438.11, 589.52, 770.94, 890.11, 1035.44, 1211.18, 1426.34, 1697.12, 2268.14],
        [359.10, 487.29, 655.57, 857.13, 989.79, 1152.31, 1348.48, 1588.82, 1891.69, 2532.68],
        [372.12, 504.87, 677.77, 884.04, 1007.55, 1153.88, 1321.85, 1558.78, 1867.85, 2684.50],
    ],
    "CO2_Before_MC": [
        [358.82, 435.61, 611.99, 918.04, 1006.56, 1125.63, 1324.75, 1597.24, 1950.37, 2539.71],
        [341.35, 417.93, 585.95, 873.30, 957.72, 1078.98, 1271.54, 1531.56, 1871.19, 2454.72],
        [322.05, 397.59, 552.08, 820.26, 899.46, 1015.53, 1189.89, 1430.13, 1735.17, 2272.42],
        [357.26, 432.82, 614.18, 915.87, 1000.32, 1122.18, 1324.76, 1595.88, 1946.36, 2537.09],
        [374.07, 579.59, 679.04, 897.16, 987.69, 1141.33, 1348.23, 1549.98, 1788.75, 2683.69],
    ],
    "CO2_After_MC": [
        [385.77, 544.22, 699.84, 852.30, 902.10, 1097.65, 1316.51, 1545.43, 1829.04, 2507.27],
        [337.35, 513.09, 669.51, 804.99, 835.58, 1062.75, 1316.96, 1545.85, 1824.72, 2441.33],
        [284.67, 449.68, 602.62, 752.68, 786.17, 959.40, 1201.08, 1410.60, 1657.71, 2276.62],
        [353.16, 522.78, 658.91, 832.05, 879.13, 1080.81, 1331.95, 1553.76, 1914.72, 2598.91],
        [438.46, 578.74, 667.46, 743.69, 817.67, 1056.39, 1357.72, 1601.87, 1860.39, 2665.96],
    ],
}

# Convert data to DataFrame
scenarios = data_nn["Scenario"]

mse_results = []
rmse_results = []

# Calculating MSE and RMSE for FC and CO2 across all scenarios
for idx, scenario in enumerate(scenarios):
    mse_fc_before_mc = np.mean([(o - p) ** 2 for o, p in zip(data_nn["Original_FC"][idx], data_nn["FC_Before_MC"][idx])])
    rmse_fc_before_mc = np.sqrt(mse_fc_before_mc)

    mse_fc_after_mc = np.mean([(o - p) ** 2 for o, p in zip(data_nn["Original_FC"][idx], data_nn["FC_After_MC"][idx])])
    rmse_fc_after_mc = np.sqrt(mse_fc_after_mc)

    mse_co2_before_mc = np.mean([(o - p) ** 2 for o, p in zip(data_nn["Original_CO2"][idx], data_nn["CO2_Before_MC"][idx])])
    rmse_co2_before_mc = np.sqrt(mse_co2_before_mc)

    mse_co2_after_mc = np.mean([(o - p) ** 2 for o, p in zip(data_nn["Original_CO2"][idx], data_nn["CO2_After_MC"][idx])])
    rmse_co2_after_mc = np.sqrt(mse_co2_after_mc)

    mse_results.append([scenario, mse_fc_before_mc, mse_fc_after_mc, mse_co2_before_mc, mse_co2_after_mc])
    rmse_results.append([scenario, rmse_fc_before_mc, rmse_fc_after_mc, rmse_co2_before_mc, rmse_co2_after_mc])

# Create DataFrames for MSE and RMSE
columns = ["Scenario", "MSE_FC_Before_MC", "MSE_FC_After_MC", "MSE_CO2_Before_MC", "MSE_CO2_After_MC"]
mse_df = pd.DataFrame(mse_results, columns=columns)

columns = ["Scenario", "RMSE_FC_Before_MC", "RMSE_FC_After_MC", "RMSE_CO2_Before_MC", "RMSE_CO2_After_MC"]
rmse_df = pd.DataFrame(rmse_results, columns=columns)

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML-RESULT/NN_MSE_RMSE_Results_Updated.xlsx"
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
