import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# Data for SVR for all scenarios
data_svr = {
    "Scenario": ["Original Scenar_Coeff", "Paint (5%) Sce_Coeff", "Advance Propell_Coeff", "Fin (2%-4%) Sce_Coeff", "Bulbous Bow Sce_Coeff"],
    "Ship_Speed": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
    "Original_FC": [
        [115.77, 157.10, 211.31, 276.27, 319.00, 371.37, 434.48, 511.78, 609.11, 814.84],
        [110.07, 149.43, 201.29, 263.48, 304.59, 354.93, 415.96, 490.98, 585.73, 787.84],
        [103.77, 140.87, 189.55, 247.89, 286.21, 332.94, 389.45, 458.63, 545.70, 729.31],
        [115.47, 156.69, 210.79, 275.61, 318.26, 370.52, 433.59, 510.87, 608.26, 814.37],
        [119.65, 162.34, 217.93, 284.26, 323.97, 371.02, 425.03, 501.22, 600.60, 863.18],
    ],
    "FC_Before_MC": [
        [212.54, 196.64, 216.88, 276.37, 318.90, 373.66, 434.38, 474.55, 481.48, 477.50],
        [198.95, 184.42, 205.53, 263.58, 304.49, 356.72, 415.86, 456.18, 464.33, 462.30],
        [180.69, 168.21, 191.16, 247.99, 286.11, 333.87, 389.35, 429.47, 440.68, 441.29],
        [212.10, 196.26, 216.49, 275.71, 318.16, 372.81, 433.49, 473.88, 480.81, 476.84],
        [216.47, 203.17, 225.40, 284.36, 323.87, 372.53, 424.93, 462.09, 463.11, 472.51],
    ],
    "FC_After_MC": [
        [131.15, 158.63, 204.01, 270.24, 315.05, 371.65, 439.56, 519.57, 608.82, 729.81],
        [122.51, 148.25, 191.83, 254.45, 297.01, 352.77, 423.02, 503.16, 594.70, 713.91],
        [119.11, 144.35, 184.65, 244.12, 283.93, 333.13, 393.13, 464.51, 544.90, 652.53],
        [132.30, 159.65, 203.86, 270.24, 314.96, 370.44, 437.63, 518.14, 608.48, 728.99],
        [132.08, 162.32, 210.02, 277.42, 319.65, 371.24, 431.32, 509.19, 602.63, 757.51],
    ],
    "Original_CO2": [
        [360.04, 488.57, 657.18, 859.19, 992.10, 1154.96, 1351.23, 1591.63, 1894.34, 2534.16],
        [342.31, 464.74, 625.99, 819.42, 947.26, 1103.83, 1293.65, 1526.96, 1821.63, 2450.20],
        [322.74, 438.11, 589.52, 770.94, 890.11, 1035.44, 1211.18, 1426.34, 1697.12, 2268.14],
        [359.10, 487.29, 655.57, 857.13, 989.79, 1152.31, 1348.48, 1588.82, 1891.69, 2532.68],
        [372.12, 504.87, 677.77, 884.04, 1007.55, 1153.88, 1321.85, 1558.78, 1867.85, 2684.50],
    ],
    "CO2_Before_MC": [
        [916.57, 870.04, 859.10, 929.66, 1011.34, 1135.73, 1255.64, 1298.03, 1259.18, 1212.02],
        [868.31, 821.95, 811.44, 882.41, 964.10, 1086.99, 1207.03, 1249.67, 1210.61, 1164.18],
        [805.90, 759.40, 748.60, 819.71, 901.52, 1024.03, 1143.89, 1186.50, 1147.83, 1101.04],
        [913.77, 867.25, 856.31, 927.08, 1009.08, 1133.01, 1252.54, 1295.05, 1256.27, 1209.18],
        [918.97, 874.11, 868.17, 946.85, 1025.59, 1135.84, 1241.93, 1287.14, 1245.10, 1207.23],
    ],
    "CO2_After_MC": [
        [447.12, 526.28, 660.41, 855.70, 988.45, 1156.60, 1350.77, 1566.88, 1786.11, 2007.91],
        [406.61, 484.21, 618.50, 811.33, 941.24, 1106.21, 1302.45, 1515.70, 1734.06, 1945.90],
        [380.10, 453.81, 578.50, 763.90, 887.81, 1039.68, 1217.76, 1415.87, 1616.59, 1810.60],
        [432.24, 516.39, 657.27, 859.67, 994.58, 1161.80, 1354.99, 1565.76, 1776.29, 1972.53],
        [443.18, 528.52, 673.55, 874.88, 1000.76, 1156.77, 1333.62, 1544.01, 1769.86, 2016.11],
    ],
}

# Convert data to DataFrame
scenarios = data_svr["Scenario"]
metrics = ["Original_FC", "FC_Before_MC", "FC_After_MC", "Original_CO2", "CO2_Before_MC", "CO2_After_MC"]

mse_results = []
rmse_results = []

# Calculating MSE and RMSE for FC and CO2 across all scenarios
for idx, scenario in enumerate(scenarios):
    mse_fc_before_mc = np.mean([(o - p) ** 2 for o, p in zip(data_svr["Original_FC"][idx], data_svr["FC_Before_MC"][idx])])
    rmse_fc_before_mc = np.sqrt(mse_fc_before_mc)

    mse_fc_after_mc = np.mean([(o - p) ** 2 for o, p in zip(data_svr["Original_FC"][idx], data_svr["FC_After_MC"][idx])])
    rmse_fc_after_mc = np.sqrt(mse_fc_after_mc)

    mse_co2_before_mc = np.mean([(o - p) ** 2 for o, p in zip(data_svr["Original_CO2"][idx], data_svr["CO2_Before_MC"][idx])])
    rmse_co2_before_mc = np.sqrt(mse_co2_before_mc)

    mse_co2_after_mc = np.mean([(o - p) ** 2 for o, p in zip(data_svr["Original_CO2"][idx], data_svr["CO2_After_MC"][idx])])
    rmse_co2_after_mc = np.sqrt(mse_co2_after_mc)

    mse_results.append([scenario, mse_fc_before_mc, mse_fc_after_mc, mse_co2_before_mc, mse_co2_after_mc])
    rmse_results.append([scenario, rmse_fc_before_mc, rmse_fc_after_mc, rmse_co2_before_mc, rmse_co2_after_mc])

# Create DataFrames for MSE and RMSE
columns = ["Scenario", "MSE_FC_Before_MC", "MSE_FC_After_MC", "MSE_CO2_Before_MC", "MSE_CO2_After_MC"]
mse_df = pd.DataFrame(mse_results, columns=columns)

columns = ["Scenario", "RMSE_FC_Before_MC", "RMSE_FC_After_MC", "RMSE_CO2_Before_MC", "RMSE_CO2_After_MC"]
rmse_df = pd.DataFrame(rmse_results, columns=columns)

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML-RESULT"
if not os.path.exists(output_path):
    os.makedirs(output_path)

# Save to Excel
with pd.ExcelWriter(os.path.join(output_path, 'SVR_MSE_RMSE_Results.xlsx')) as writer:
    mse_df.to_excel(writer, sheet_name='MSE_Results', index=False)
    rmse_df.to_excel(writer, sheet_name='RMSE_Results', index=False)

# Plotting the MSE and RMSE for visualization
plt.figure(figsize=(12, 8))

# Plot MSE for FC Before and After MC
plt.subplot(2, 2, 1)
for scenario, mse_fc_before, mse_fc_after in zip(scenarios, mse_df['MSE_FC_Before_MC'], mse_df['MSE_FC_After_MC']):
    plt.bar(scenario, mse_fc_before, alpha=0.5, label='Before MC')
    plt.bar(scenario, mse_fc_after, alpha=0.5, label='After MC')
plt.title('MSE for Fuel Consumption (FC)')
plt.ylabel('MSE')
plt.xticks(rotation=45)
plt.legend()

# Plot MSE for CO2 Before and After MC
plt.subplot(2, 2, 2)
for scenario, mse_co2_before, mse_co2_after in zip(scenarios, mse_df['MSE_CO2_Before_MC'], mse_df['MSE_CO2_After_MC']):
    plt.bar(scenario, mse_co2_before, alpha=0.5, label='Before MC')
    plt.bar(scenario, mse_co2_after, alpha=0.5, label='After MC')
plt.title('MSE for CO2 Emission')
plt.ylabel('MSE')
plt.xticks(rotation=45)
plt.legend()

# Plot RMSE for FC Before and After MC
plt.subplot(2, 2, 3)
for scenario, rmse_fc_before, rmse_fc_after in zip(scenarios, rmse_df['RMSE_FC_Before_MC'], rmse_df['RMSE_FC_After_MC']):
    plt.bar(scenario, rmse_fc_before, alpha=0.5, label='Before MC')
    plt.bar(scenario, rmse_fc_after, alpha=0.5, label='After MC')
plt.title('RMSE for Fuel Consumption (FC)')
plt.ylabel('RMSE')
plt.xticks(rotation=45)
plt.legend()

# Plot RMSE for CO2 Before and After MC
plt.subplot(2, 2, 4)
for scenario, rmse_co2_before, rmse_co2_after in zip(scenarios, rmse_df['RMSE_CO2_Before_MC'], rmse_df['RMSE_CO2_After_MC']):
    plt.bar(scenario, rmse_co2_before, alpha=0.5, label='Before MC')
    plt.bar(scenario, rmse_co2_after, alpha=0.5, label='After MC')
plt.title('RMSE for CO2 Emission')
plt.ylabel('RMSE')
plt.xticks(rotation=45)
plt.legend()

plt.tight_layout()
plt.savefig(os.path.join(output_path, 'SVR_MSE_RMSE_Plots.png'))
plt.show()
