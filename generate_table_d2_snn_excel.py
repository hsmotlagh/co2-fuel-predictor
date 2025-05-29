import xlsxwriter
import numpy as np

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Table_D2_SNN.xlsx"

# Data for SNN
data = [
    {'Scenario': 'Original', 'MSE': 702.9967199},
    {'Scenario': 'Paint', 'MSE': 318.0265551},
    {'Scenario': 'Advance Propeller', 'MSE': 663.6611939},
    {'Scenario': 'Fin', 'MSE': 572.3498592},
    {'Scenario': 'Bulbous Bow', 'MSE': 219.6918737},
    {'Scenario': 'Combined', 'MSE': 180.5276211}
]

try:
    workbook = xlsxwriter.Workbook(output_path)
    worksheet = workbook.add_worksheet('SNN')

    # Define formats
    merge_format = workbook.add_format({
        'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True
    })
    cell_format = workbook.add_format({'border': 1, 'align': 'center'})
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'border': 1, 'text_wrap': True})

    # Write title and caption
    worksheet.merge_range('A1:C1', 'Table D2: SNN Model Performance Metrics for FC After MC', merge_format)
    worksheet.merge_range('A2:C2', 'Appendix D - Model Performance Reference', merge_format)

    # Define headers
    headers = ['Scenario', 'MSE FC After\n(kg/h)²', 'RMSE FC After\n(kg/h)']

    # Write headers
    for col, header in enumerate(headers):
        worksheet.write(2, col, header, header_format)

    # Write data
    for row, record in enumerate(data, 3):
        worksheet.write(row, 0, record['Scenario'], cell_format)
        mse = record['MSE']
        rmse = np.sqrt(mse)
        worksheet.write(row, 1, mse, cell_format)
        worksheet.write(row, 2, rmse, cell_format)

    # Write footnotes
    footnote_row = len(data) + 4
    worksheet.write(footnote_row, 0, 'Notes:', header_format)
    worksheet.write(footnote_row + 1, 0, '1. SNN metrics are provided only for FC after MC simulations; CO₂ metrics are unavailable.', cell_format)
    worksheet.write(footnote_row + 2, 0, '2. RMSE is computed as √MSE.', cell_format)

    # Adjust column widths
    for col, header in enumerate(headers):
        worksheet.set_column(col, col, max(len(header.split('\n')[0]), 12))

    # Adjust row height for headers
    worksheet.set_row(2, 60)

    workbook.close()
    print(f"SNN Table D2 saved to {output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")