import xlsxwriter

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Table_D2_SVR.xlsx"

# Data for SVR
data = [
    {'Scenario': 'Original', 'MSE_FC_Before': 14243.79089, 'RMSE_FC_Before': 119.3473539, 'MSE_CO2_Before': 274854.3493, 'RMSE_CO2_Before': 524.2655332, 'MSE_FC_After': 766.08926, 'RMSE_FC_After': 27.67831751, 'MSE_CO2_After': 29830.91651, 'RMSE_CO2_After': 172.7162891},
    {'Scenario': 'Paint', 'MSE_FC_Before': 13107.04578, 'RMSE_FC_Before': 114.4860069, 'MSE_CO2_Before': 255478.7247, 'RMSE_CO2_Before': 505.4490328, 'MSE_FC_After': 613.36023, 'RMSE_FC_After': 24.76611051, 'MSE_CO2_After': 26686.82464, 'RMSE_CO2_After': 163.3610255},
    {'Scenario': 'Advance Propeller', 'MSE_FC_Before': 10150.26754, 'RMSE_FC_Before': 100.7485362, 'MSE_CO2_Before': 209050.9269, 'RMSE_CO2_Before': 457.2208732, 'MSE_FC_After': 623.48086, 'RMSE_FC_After': 24.96959872, 'MSE_CO2_After': 21971.1779, 'RMSE_CO2_After': 148.2267786},
    {'Scenario': 'Fin', 'MSE_FC_Before': 14247.91694, 'RMSE_FC_Before': 119.3646386, 'MSE_CO2_Before': 274887.8678, 'RMSE_CO2_Before': 524.2974993, 'MSE_FC_After': 773.8736, 'RMSE_FC_After': 27.81858372, 'MSE_CO2_After': 33397.79416, 'RMSE_CO2_After': 182.7506338},
    {'Scenario': 'Bulbous Bow', 'MSE_FC_Before': 18415.70182, 'RMSE_FC_Before': 135.704465, 'MSE_CO2_Before': 312655.0951, 'RMSE_CO2_Before': 559.1556984, 'MSE_FC_After': 1155.59246, 'RMSE_FC_After': 33.99400624, 'MSE_CO2_After': 46246.89343, 'RMSE_CO2_After': 215.0509089},
    {'Scenario': 'Combined', 'MSE_FC_Before': None, 'RMSE_FC_Before': None, 'MSE_CO2_Before': None, 'RMSE_CO2_Before': None, 'MSE_FC_After': None, 'RMSE_FC_After': None, 'MSE_CO2_After': None, 'RMSE_CO2_After': None}
]

try:
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet('SVR')

    # Define formats
    merge_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True
    })
    cell_format = workbook.add_format({'border': 1, 'align': 'center'})
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'text_wrap': True})

    # Write title and caption
    worksheet.merge_range('A1:I1', 'Table D2: SVR Model Performance Metrics for FC and CO₂ Emissions', merge_format)
    worksheet.merge_range('A2:I2', 'Appendix D - Model Performance Reference', merge_format)

    # Define headers
    headers = [
        'Scenario', 'MSE FC Before\n(kg/h)²', 'RMSE FC Before\n(kg/h)', 'MSE CO₂ Before\n(kg)²', 'RMSE CO₂ Before\n(kg)',
        'MSE FC After\n(kg/h)²', 'RMSE FC After\n(kg/h)', 'MSE CO₂ After\n(kg)²', 'RMSE CO₂ After\n(kg)'
    ]

    # Write headers
    for col, header in enumerate(headers):
        worksheet.write(2, col, header, header_format)

    # Write data
    for row, record in enumerate(data, 3):
        worksheet.write(row, 0, record['Scenario'], cell_format)
        for col, key in enumerate(['MSE_FC_Before', 'RMSE_FC_Before', 'MSE_CO2_Before', 'RMSE_CO2_Before', 'MSE_FC_After', 'RMSE_FC_After', 'MSE_CO2_After', 'RMSE_CO2_After'], 1):
            value = record[key]
            if value is None:
                worksheet.write(row, col, '--', cell_format)
            else:
                worksheet.write(row, col, value, cell_format)

    # Write footnotes
    footnote_row = len(data) + 4
    worksheet.write(footnote_row, 0, 'Notes:', header_format)
    worksheet.write(footnote_row + 1, 0, '1. Combined Scenario metrics are pending calculation from Combined_Scenario_Data.xlsx.', cell_format)

    # Adjust column widths
    for col, header in enumerate(headers):
        worksheet.set_column(col, col, max(len(header.split('\n')[0]), 12))

    # Adjust row height for headers
    worksheet.set_row(2, 60)

    workbook.close()
    print(f"SVR Table D2 saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")