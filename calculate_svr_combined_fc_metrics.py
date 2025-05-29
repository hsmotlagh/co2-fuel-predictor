import pandas as pd
import numpy as np
import xlsxwriter

# Define input and output paths
input_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Combined_Scenario_Data.xlsx"
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Table_D2_SVR_Combined.xlsx"

# Read Combined Scenario data
try:
    data = pd.read_excel(input_path, sheet_name='Combined_Scen_Coeff')
except Exception as e:
    print(f"Error reading Combined Scenario data: {e}")
    exit()

# Verify required columns
required_cols = [
    'Ship Speed (knots)', 'Fuel Consumption (FC) (kg/h)',
    'Final Predicted Fuel Consumption SVR Before MC (kg/h)',
    'Final Predicted Fuel Consumption SVR After MC (kg/h)'
]
missing_cols = [col for col in required_cols if col not in data.columns]
if missing_cols:
    print(f"Missing columns in Combined Scenario data: {missing_cols}")
    exit()

# Calculate MSE and RMSE
def calculate_mse(actual, predicted):
    return np.mean((actual - predicted) ** 2)

def calculate_rmse(mse):
    return np.sqrt(mse)

mse_fc_before = calculate_mse(data['Fuel Consumption (FC) (kg/h)'], data['Final Predicted Fuel Consumption SVR Before MC (kg/h)'])
rmse_fc_before = calculate_rmse(mse_fc_before)
mse_fc_after = calculate_mse(data['Fuel Consumption (FC) (kg/h)'], data['Final Predicted Fuel Consumption SVR After MC (kg/h)'])
rmse_fc_after = calculate_rmse(mse_fc_after)

# Prepare results
results = [{
    'Scenario': 'Combined',
    'MSE_FC_Before': mse_fc_before,
    'RMSE_FC_Before': rmse_fc_before,
    'MSE_FC_After': mse_fc_after,
    'RMSE_FC_After': rmse_fc_after
}]

# Create Excel file
try:
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet('SVR_Combined')

    # Define formats
    merge_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True
    })
    cell_format = workbook.add_format({'border': 1, 'align': 'center'})
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'text_wrap': True})

    # Write title and caption
    worksheet.merge_range('A1:E1', 'Table D2: SVR Model Performance Metrics for FC in Combined Scenario', merge_format)
    worksheet.merge_range('A2:E2', 'Appendix D - Model Performance Reference', merge_format)

    # Define headers
    headers = [
        'Scenario', 'MSE FC Before\n(kg/h)²', 'RMSE FC Before\n(kg/h)',
        'MSE FC After\n(kg/h)²', 'RMSE FC After\n(kg/h)'
    ]

    # Write headers
    for col, header in enumerate(headers):
        worksheet.write(2, col, header, header_format)

    # Write data
    for row, record in enumerate(results, 3):
        worksheet.write(row, 0, record['Scenario'], cell_format)
        worksheet.write(row, 1, record['MSE_FC_Before'], cell_format)
        worksheet.write(row, 2, record['RMSE_FC_Before'], cell_format)
        worksheet.write(row, 3, record['MSE_FC_After'], cell_format)
        worksheet.write(row, 4, record['RMSE_FC_After'], cell_format)

    # Write footnotes
    footnote_row = len(results) + 4
    worksheet.write(footnote_row, 0, 'Notes:', header_format)
    worksheet.write(footnote_row + 1, 0, '1. Metrics calculated for FC only; CO₂ metrics are pending.', cell_format)
    worksheet.write(footnote_row + 2, 0, '2. MSE = (1/n) * Σ(Actual - Predicted)², RMSE = √MSE.', cell_format)

    # Adjust column widths
    for col, header in enumerate(headers):
        worksheet.set_column(col, col, max(len(header.split('\n')[0]), 12))

    # Adjust row height for headers
    worksheet.set_row(2, 60)

    workbook.close()
    print(f"SVR Combined Scenario FC metrics saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")