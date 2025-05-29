import pandas as pd
import numpy as np
import os

# Complete data for XGB for all scenarios
data_xgb = {
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
        [115.77, 115.77, 211.31, 276.27, 319.00, 371.37, 434.48, 511.78, 511.78, 814.84],
        [110.07, 110.07, 201.29, 263.48, 304.59, 354.93, 415.96, 490.98, 490.98, 787.84],
        [103.77, 103.77, 189.55, 247.89, 286.21, 332.94, 389.45, 458.63, 458.63, 729.31],
        [115.47, 115.47, 210.79, 275.61, 318.26, 370.52, 433.59, 510.87, 510.87, 814.37],
        [119.65, 119.65, 217.93, 284.26, 323.97, 371.02, 425.03, 501.22, 501.22, 863.18],
    ],
    "FC_After_MC": [
        [115.77, 157.10, 211.31, 276.27, 319.00, 371.37, 434.48, 511.78, 609.11, 814.84],
        [110.07, 149.43, 201.29, 263.48, 304.59, 354.93, 415.96, 490.98, 585.73, 787.84],
        [103.77, 140.87, 189.55, 247.89, 286.21, 332.94, 389.45, 458.63, 545.70, 729.31],
        [115.47, 156.69, 210.79, 275.61, 318.26, 370.52, 433.59, 510.87, 608.26, 814.37],
        [119.65, 162.34, 217.93, 284.26, 323.97, 371.02, 425.03, 501.22, 600.60, 863.18],
    ],
    "Original_CO2": [
        [360.04, 488.57, 657.18, 859.19, 992.10, 1154.96, 1351.23, 1591.63, 1894.34, 2534.16],
        [342.31, 464.74, 625.99, 819.42, 947.26, 1103.83, 1293.65, 1526.96, 1821.63, 2450.20],
        [322.74, 438.11, 589.52, 770.94, 890.11, 1035.44, 1211.18, 1426.34, 1697.12, 2268.14],
        [359.10, 487.29, 655.57, 857.13, 989.79, 1152.31, 1348.48, 1588.82, 1891.69, 2532.68],
        [372.12, 504.87, 677.77, 884.04, 1007.55, 1153.88, 1321.85, 1558.78, 1867.85, 2684.50],
    ],
    "CO2_Before_MC": [
        [360.04, 360.04, 657.18, 859.19, 992.10, 1154.96, 1351.23, 1591.63, 1591.63, 2534.16],
        [342.31, 342.31, 625.99, 819.42, 947.26, 1103.83, 1293.65, 1526.96, 1526.96, 2450.20],
        [322.74, 322.74, 589.52, 770.94, 890.11, 1035.44, 1211.18, 1426.34, 1426.34, 2268.14],
        [359.10, 359.10, 655.57, 857.13, 989.79, 1152.31, 1348.48, 1588.82, 1588.82, 2532.68],
        [372.12, 372.12, 677.77, 884.04, 1007.55, 1153.88, 1321.85, 1558.78, 1558.78, 2684.50],
    ],
    "CO2_After_MC": [
        [360.04, 488.57, 657.18, 859.19, 992.10, 1154.96, 1351.23, 1591.63, 1894.34, 2534.16],
        [342.31, 464.74, 625.99, 819.42, 947.26, 1103.83, 1293.65, 1526.96, 1821.63, 2450.20],
        [322.74, 438.11, 589.52, 770.94, 890.11, 1035.44, 1211.18, 1426.34, 1697.12, 2268.14],
        [359.10, 487.29, 655.57, 857.13, 989.79, 1152.31, 1348.48, 1588.82, 1891.69, 2532.68],
        [372.12, 504.87, 677.77, 884.04, 1007.55, 1153.88, 1321.85, 1558.78, 1867.85, 2684.50],
    ],
}

# Convert data to DataFrame
scenarios = data_xgb["Scenario"]

mse_results = []
rmse_results = []

# Calculating MSE and RMSE for FC and CO2 across all scenarios
for idx, scenario in enumerate(scenarios):
    mse_fc_before_mc = np.mean([(o - p) ** 2 for o, p in zip(data_xgb["Original_FC"][idx], data_xgb["FC_Before_MC"][idx])])
    rmse_fc_before_mc = np.sqrt(mse_fc_before_mc)

    mse_fc_after_mc = np.mean([(o - p) ** 2 for o, p in zip(data_xgb["Original_FC"][idx], data_xgb["FC_After_MC"][idx])])
    rmse_fc_after_mc = np.sqrt(mse_fc_after_mc)

    mse_co2_before_mc = np.mean([(o - p) ** 2 for o, p in zip(data_xgb["Original_CO2"][idx], data_xgb["CO2_Before_MC"][idx])])
    rmse_co2_before_mc = np.sqrt(mse_co2_before_mc)

    mse_co2_after_mc = np.mean([(o - p) ** 2 for o, p in zip(data_xgb["Original_CO2"][idx], data_xgb["CO2_After_MC"][idx])])
    rmse_co2_after_mc = np.sqrt(mse_co2_after_mc)

    mse_results.append([scenario, mse_fc_before_mc, mse_fc_after_mc, mse_co2_before_mc, mse_co2_after_mc])
    rmse_results.append([scenario, rmse_fc_before_mc, rmse_fc_after_mc, rmse_co2_before_mc, rmse_co2_after_mc])

# Create DataFrames for MSE and RMSE
columns = ["Scenario", "MSE_FC_Before_MC", "MSE_FC_After_MC", "MSE_CO2_Before_MC", "MSE_CO2_After_MC"]
mse_df = pd.DataFrame(mse_results, columns=columns)

columns = ["Scenario", "RMSE_FC_Before_MC", "RMSE_FC_After_MC", "RMSE_CO2_Before_MC", "RMSE_CO2_After_MC"]
rmse_df = pd.DataFrame(rmse_results, columns=columns)

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML-RESULT/XGB_MSE_RMSE_Results_Updated.xlsx"
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
