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
n_mse_after_mc = [310.734, 92.91818, 184.74624, 64.58955, 620.18957]
n_rmse_before_mc = [3.95975378, 9.71072088, 8.397284085, 10.92385097, 11.37208776]
n_rmse_after_mc = [17.62764874, 9.639407658, 13.5921389, 8.036762408, 24.90360556]

# Create an ExcelWriter object to save the charts to Excel
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML-RESULT/ML_Model_Evaluation_Charts_2.xlsx"
with ExcelWriter(output_path, engine='xlsxwriter') as writer:
    workbook = writer.book
    
    # Add SVR MSE and RMSE data to worksheet for Figure 9
    worksheet_svr = workbook.add_worksheet('Figure 9')
    worksheet_svr.write_row('A1', ['Scenario', 'MSE Before MC', 'MSE After MC', 'RMSE Before MC', 'RMSE After MC'])
    for idx, scenario in enumerate(scenarios):
        worksheet_svr.write(idx + 1, 0, scenario)
        worksheet_svr.write(idx + 1, 1, svr_mse_before_mc[idx])
        worksheet_svr.write(idx + 1, 2, svr_mse_after_mc[idx])
        worksheet_svr.write(idx + 1, 3, svr_rmse_before_mc[idx])
        worksheet_svr.write(idx + 1, 4, svr_rmse_after_mc[idx])
    chart_svr = workbook.add_chart({'type': 'column'})
    chart_svr.add_series({
        'name': 'MSE Before MC',
        'categories': '=Figure 9!$A$2:$A$6',
        'values': '=Figure 9!$B$2:$B$6'
    })
    chart_svr.add_series({
        'name': 'MSE After MC',
        'categories': '=Figure 9!$A$2:$A$6',
        'values': '=Figure 9!$C$2:$C$6'
    })
    chart_svr.add_series({
        'name': 'RMSE Before MC',
        'categories': '=Figure 9!$A$2:$A$6',
        'values': '=Figure 9!$D$2:$D$6'
    })
    chart_svr.add_series({
        'name': 'RMSE After MC',
        'categories': '=Figure 9!$A$2:$A$6',
        'values': '=Figure 9!$E$2:$E$6'
    })
    chart_svr.set_title({'name': 'Figure 9: MSE and RMSE Performance of SVR Across Scenarios'})
    chart_svr.set_x_axis({'name': 'Scenarios'})
    chart_svr.set_y_axis({'name': 'Error Values'})
    worksheet_svr.insert_chart('G2', chart_svr)

    # Add XGBoost and GPR data for Figure 10
    worksheet_xgb_gpr = workbook.add_worksheet('Figure 10')
    worksheet_xgb_gpr.write_row('A1', ['Scenario', 'XGBoost MSE Before MC', 'GPR MSE Before MC'])
    for idx, scenario in enumerate(scenarios):
        worksheet_xgb_gpr.write(idx + 1, 0, scenario)
        worksheet_xgb_gpr.write(idx + 1, 1, xgb_mse_before_mc[idx])
        worksheet_xgb_gpr.write(idx + 1, 2, gpr_mse_before_mc[idx])
    chart_xgb_gpr = workbook.add_chart({'type': 'column'})
    chart_xgb_gpr.add_series({
        'name': 'XGBoost MSE Before MC',
        'categories': '=Figure 10!$A$2:$A$6',
        'values': '=Figure 10!$B$2:$B$6'
    })
    chart_xgb_gpr.add_series({
        'name': 'GPR MSE Before MC',
        'categories': '=Figure 10!$A$2:$A$6',
        'values': '=Figure 10!$C$2:$C$6'
    })
    chart_xgb_gpr.set_title({'name': 'Figure 10: Comparative Analysis of XGBoost vs. GPR MSE Performance'})
    chart_xgb_gpr.set_x_axis({'name': 'Scenarios'})
    chart_xgb_gpr.set_y_axis({'name': 'MSE'})
    worksheet_xgb_gpr.insert_chart('E2', chart_xgb_gpr)

    # Add Random Forest and NN data for Figure 11
    worksheet_rf_nn = workbook.add_worksheet('Figure 11')
    worksheet_rf_nn.write_row('A1', ['Scenario', 'RF MSE Before MC', 'NN MSE Before MC'])
    for idx, scenario in enumerate(scenarios):
        worksheet_rf_nn.write(idx + 1, 0, scenario)
        worksheet_rf_nn.write(idx + 1, 1, rf_mse_before_mc[idx])
        worksheet_rf_nn.write(idx + 1, 2, nn_mse_before_mc[idx])
    chart_rf_nn = workbook.add_chart({'type': 'column'})
    chart_rf_nn.add_series({
        'name': 'RF MSE Before MC',
        'categories': '=Figure 11!$A$2:$A$6',
        'values': '=Figure 11!$B$2:$B$6'
    })
    chart_rf_nn.add_series({
        'name': 'NN MSE Before MC',
        'categories': '=Figure 11!$A$2:$A$6',
        'values': '=Figure 11!$C$2:$C$6'
    })
    chart_rf_nn.set_title({'name': 'Figure 11: Random Forest vs. Neural Network MSE Performance Comparison'})
    chart_rf_nn.set_x_axis({'name': 'Scenarios'})
    chart_rf_nn.set_y_axis({'name': 'MSE'})
    worksheet_rf_nn.insert_chart('E2', chart_rf_nn)
