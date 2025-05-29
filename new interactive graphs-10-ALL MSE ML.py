import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pandas import ExcelWriter
import xlsxwriter

# Data Setup for MSE and RMSE results
# SVR, GPR, RF, XGB, NN models in different scenarios

# Scenario Names
scenarios = ["Original Scenario", "Paint (5%) Scenario", "Advance Propeller Scenario", "Fin (2%-4%) Scenario", "Bulbous Bow Scenario"]

# SVR MSE and RMSE values before and after Monte Carlo (MC)
svr_mse_before_mc = [14243.79089, 13107.04578, 10150.26754, 14247.91694, 18415.70182]
svr_mse_after_mc = [766.08926, 613.36023, 623.48086, 773.8736, 1155.59246]
svr_rmse_before_mc = [119.3473539, 114.4860069, 100.7485362, 119.3646386, 135.704465]
svr_rmse_after_mc = [27.67831751, 24.76611051, 24.96959872, 27.81858372, 33.99400624]

# GPR MSE and RMSE values before and after MC
gpr_mse_before_mc = [36.45125, 33.88789, 27.13936, 35.30033, 23.26672]
gpr_mse_after_mc = [34.61913, 26.64075, 858.80455, 11.53899, 270.82693]
gpr_rmse_before_mc = [6.03748706, 5.821330604, 5.209545086, 5.941408082, 4.823558852]
gpr_rmse_after_mc = [5.883802342, 5.161467814, 29.30536726, 3.396908889, 16.45682017]

# RF MSE and RMSE values before and after MC
rf_mse_before_mc = [1893.40515, 1986.01682, 1578.57899, 1900.89755, 2512.84462]
rf_mse_after_mc = [2.72035, 1.5741, 2.35278, 5.91939, 3.62797]
rf_rmse_before_mc = [43.51327556, 44.56474862, 39.73133511, 43.59928382, 50.12828164]
rf_rmse_after_mc = [1.649348356, 1.25463142, 1.53387744, 2.432979655, 1.904723077]

# XGB MSE and RMSE values before and after MC
xgb_mse_before_mc = [1118.12978, 1052.67721, 895.75949, 1118.39005, 1169.88205]
xgb_mse_after_mc = [0, 0, 0, 0, 0]
xgb_rmse_before_mc = [33.43844763, 32.44498744, 29.92924139, 33.44233918, 34.20353856]
xgb_rmse_after_mc = [0, 0, 0, 0, 0]

# NN MSE and RMSE values before and after MC
nn_mse_before_mc = [15.67965, 94.2981, 70.51438, 119.33052, 129.32438]
nn_mse_after_mc = [310.734, 92.91818, 184.74624, 64.58955, 620.18957]
nn_rmse_before_mc = [3.95975378, 9.71072088, 8.397284085, 10.92385097, 11.37208776]
nn_rmse_after_mc = [17.62764874, 9.639407658, 13.5921389, 8.036762408, 24.90360556]

# Create an ExcelWriter object to save the charts to Excel
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML-RESULT/ML_Model_Evaluation_Charts.xlsx"
with ExcelWriter(output_path, engine='xlsxwriter') as writer:
    workbook = writer.book
    # Plot Figure 9: MSE and RMSE Performance of SVR Across Scenarios
    plt.figure(figsize=(14, 6))
    plt.subplot(1, 2, 1)
    plt.barh(scenarios, svr_mse_before_mc, color='lightblue', label='MSE Before MC')
    plt.barh(scenarios, svr_mse_after_mc, color='navy', alpha=0.7, label='MSE After MC')
    plt.xlabel('MSE')
    plt.title('SVR MSE Performance Across Scenarios')
    plt.legend()
    plt.subplot(1, 2, 2)
    plt.barh(scenarios, svr_rmse_before_mc, color='orange', label='RMSE Before MC')
    plt.barh(scenarios, svr_rmse_after_mc, color='darkred', alpha=0.7, label='RMSE After MC')
    plt.xlabel('RMSE')
    plt.title('SVR RMSE Performance Across Scenarios')
    plt.legend()

    plt.suptitle('Figure 9: MSE and RMSE Performance of SVR Across Scenarios')
    plt.tight_layout()
    plt.savefig("Figure_9.png")
    worksheet = workbook.add_worksheet('Figure 9')
    worksheet.insert_image('B2', 'Figure_9.png')
    plt.close()

    # Plot Figure 10: Comparative Analysis of XGBoost vs. GPR MSE Performance
    x = np.arange(len(scenarios))
    width = 0.35

    plt.figure(figsize=(12, 6))
    plt.bar(x - width/2, xgb_mse_before_mc, width, label='XGBoost MSE Before MC', color='lightgreen')
    plt.bar(x + width/2, gpr_mse_before_mc, width, label='GPR MSE Before MC', color='skyblue')
    plt.xlabel('Scenarios')
    plt.ylabel('MSE')
    plt.title('Figure 10: Comparative Analysis of XGBoost vs. GPR MSE Performance')
    plt.xticks(ticks=x, labels=scenarios, rotation=45)
    plt.legend()
    plt.tight_layout()
    plt.savefig("Figure_10.png")
    worksheet = workbook.add_worksheet('Figure 10')
    worksheet.insert_image('B2', 'Figure_10.png')
    plt.close()

    # Plot Figure 11: Random Forest vs. Neural Network MSE Performance Comparison
    plt.figure(figsize=(12, 6))
    plt.bar(x - width/2, rf_mse_before_mc, width, label='RF MSE Before MC', color='lightcoral')
    plt.bar(x + width/2, nn_mse_before_mc, width, label='NN MSE Before MC', color='lightblue')
    plt.xlabel('Scenarios')
    plt.ylabel('MSE')
    plt.title('Figure 11: Random Forest vs. Neural Network MSE Performance Comparison')
    plt.xticks(ticks=x, labels=scenarios, rotation=45)
    plt.legend()
    plt.tight_layout()
    plt.savefig("Figure_11.png")
    worksheet = workbook.add_worksheet('Figure 11')
    worksheet.insert_image('B2', 'Figure_11.png')
    plt.close()
