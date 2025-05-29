import pandas as pd
import numpy as np
import os

# Complete data for GPR for all scenarios
data_gpr = {
    "Scenario": [
        "Original\nScenar_Coeff",
        "Paint (5%)\nSce_Coeff",
        "Advance\nPropell_Coeff",
        "Fin (2%-4%)\nSce_Coeff",
        "Bulbous Bow\nSce_Coeff"
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
        [115.77, 158.64, 211.31, 276.27, 319.00, 371.37, 434.48, 511.78, 628.14, 814.84],
        [110.07, 151.13, 201.29, 263.48, 304.59, 354.93, 415.96, 490.98, 604.06, 787.84],
        [103.77, 142.43, 189.55, 247.89, 286.21, 332.94, 389.45, 458.63, 562.10, 729.31],
        [115.47, 158.17, 210.79, 275.61, 318.26, 370.52, 433.59, 510.87, 626.99, 814.37],
        [119.65, 162.98, 217.93, 284.26, 323.97, 371.02, 425.03, 501.22, 615.84, 863.18],
    ],
    "FC_After_MC": [
        [115.52, 156.36, 210.04, 274.31, 317.02, 371.48, 439.26, 521.34, 624.00, 815.33],
        [110.95, 150.43, 202.13, 264.49, 305.94, 356.57, 417.04, 491.94, 581.25, 772.47],
        [103.56, 141.86, 191.96, 252.26, 291.48, 338.02, 391.14, 456.43, 529.98, 638.46],
        [115.43, 156.65, 210.82, 276.13, 319.22, 372.45, 437.14, 516.15, 615.84, 810.83],
        [120.13, 163.05, 218.35, 284.67, 324.62, 372.50, 428.91, 509.85, 609.54, 812.83],
    ],
    "Original_CO2": [
        [360.04, 488.57, 657.18, 859.19, 992.10, 1154.96, 1351.23, 1591.63, 1894.34, 2534.16],
        [342.31, 464.74, 625.99, 819.42, 947.26, 1103.83, 1293.65, 1526.96, 1821.63, 2450.20],
        [322.74, 438.11, 589.52, 770.94, 890.11, 1035.44, 1211.18, 1426.34, 1697.12, 2268.14],
        [359.10, 487.29, 655.57, 857.13, 989.79, 1152.31, 1348.48, 1588.82, 1891.69, 2532.68],
        [372.12, 504.87, 677.77, 884.04, 1007.55, 1153.88, 1321.85, 1558.78, 1867.85, 2684.50],
    ],
    "CO2_Before_MC": [
        [360.04, 493.71, 657.18, 859.19, 992.10, 1154.96, 1351.23, 1591.63, 1955.40, 2534.16],
        [342.31, 470.37, 625.99, 819.42, 947.26, 1103.83, 1293.65, 1526.96, 1880.57, 2450.20],
        [322.74, 443.34, 589.52, 770.94, 890.11, 1035.44, 1211.18, 1426.34, 1750.23, 2268.14],
        [359.10, 492.23, 655.57, 857.13, 989.79, 1152.31, 1348.48, 1588.82, 1951.81, 2532.68],
        [372.12, 507.16, 677.77, 884.04, 1007.55, 1153.88, 1321.85, 1558.78, 1916.75, 2684.50],
    ],
    "CO2_After_MC": [
        [358.81, 485.64, 653.53, 853.71, 986.67, 1156.82, 1368.18, 1624.50, 1943.52, 2519.77],
        [345.83, 469.36, 630.00, 823.90, 952.98, 1110.17, 1295.16, 1527.29, 1799.76, 2402.15],
        [321.92, 441.41, 598.02, 786.71, 909.35, 1054.29, 1217.19, 1416.92, 1638.16, 1949.72],
        [359.77, 488.19, 656.45, 860.11, 994.57, 1160.14, 1359.87, 1604.66, 1914.26, 2510.33],
        [375.51, 509.43, 681.72, 887.33, 1011.52, 1161.72, 1340.88, 1597.17, 1904.30, 2487.65],
    ],
}

# Convert data to DataFrame
scenarios = data_gpr["Scenario"]

mse_results = []
rmse_results = []

# Calculating MSE and RMSE for FC and CO2 across all scenarios
for idx, scenario in enumerate(scenarios):
    mse_fc_before_mc = np.mean([(o - p) ** 2 for o, p in zip(data_gpr["Original_FC"][idx], data_gpr["FC_Before_MC"][idx])])
    rmse_fc_before_mc = np.sqrt(mse_fc_before_mc)

    mse_fc_after_mc = np.mean([(o - p) ** 2 for o, p in zip(data_gpr["Original_FC"][idx], data_gpr["FC_After_MC"][idx])])
    rmse_fc_after_mc = np.sqrt(mse_fc_after_mc)

    mse_co2_before_mc = np.mean([(o - p) ** 2 for o, p in zip(data_gpr["Original_CO2"][idx], data_gpr["CO2_Before_MC"][idx])])
    rmse_co2_before_mc = np.sqrt(mse_co2_before_mc)

    mse_co2_after_mc = np.mean([(o - p) ** 2 for o, p in zip(data_gpr["Original_CO2"][idx], data_gpr["CO2_After_MC"][idx])])
    rmse_co2_after_mc = np.sqrt(mse_co2_after_mc)

    mse_results.append([scenario, mse_fc_before_mc, mse_fc_after_mc, mse_co2_before_mc, mse_co2_after_mc])
    rmse_results.append([scenario, rmse_fc_before_mc, rmse_fc_after_mc, rmse_co2_before_mc, rmse_co2_after_mc])

# Create DataFrames for MSE and RMSE
columns = ["Scenario", "MSE_FC_Before_MC", "MSE_FC_After_MC", "MSE_CO2_Before_MC", "MSE_CO2_After_MC"]
mse_df = pd.DataFrame(mse_results, columns=columns)

columns = ["Scenario", "RMSE_FC_Before_MC", "RMSE_FC_After_MC", "RMSE_CO2_Before_MC", "RMSE_CO2_After_MC"]
rmse_df = pd.DataFrame(rmse_results, columns=columns)

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML-RESULT/GPR_MSE_RMSE_Results_Updated.xlsx"
if not os.path.exists(output_path):
    os.makedirs(output_path)

# Save to Excel with updated chart configurations using XlsxWriter
with pd.ExcelWriter(os.path.join(output_path, 'GPR_MSE_RMSE_Results_Updated.xlsx'), engine='xlsxwriter') as writer:
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
