import pandas as pd
import xlsxwriter
import os

# File paths for input and output
input_file = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML-RESULT/phase1-step1.xlsx"
output_file = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML_Interactive_Chart.xlsx"

# Function to create interactive Excel charts for five sheets of new data
def create_interactive_chart(input_file, output_file):
    # Load the data from the first five sheets of the Excel file
    sheet_names = ["Sheet 1 (Original Scenario)", "Sheet 2 (Paint (5%) Scenario)", "Sheet 3 (Advance Propeller Scenario)", "Sheet 4 (Fin (2%-4%) Scenario)", "Sheet 5 (Bulbous Bow Scenario)"]
    dataframes = [pd.read_excel(input_file, sheet_name=i) for i in range(5)]
    
    # Create a new workbook and add a worksheet using xlsxwriter
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet('Data')

    # Write the data from all five sheets into the worksheet separately
    row_index = 0
    start_end_rows = []  # List to store start and end rows for each sheet
    
    for df in dataframes:
        # Write the header row if it's the first DataFrame
        worksheet.write_row(row_index, 0, df.columns.tolist())  # Write header row
        row_start = row_index + 1
        for _, row in df.iterrows():
            worksheet.write_row(row_index + 1, 0, row.tolist())
            row_index += 1
        row_end = row_index
        row_index += 2  # Leave a gap between datasets
        
        # Store the start and end rows for each sheet
        start_end_rows.append((row_start, row_end))

    # Create chart objects for Total Resistance (Rt), Effective Power (Pe), and Thrust Required (Treq)
    rt_chart = workbook.add_chart({'type': 'line'})
    pe_chart = workbook.add_chart({'type': 'line'})
    treq_chart = workbook.add_chart({'type': 'line'})

    # Get necessary columns for plotting from the first DataFrame (assuming all have the same structure)
    try:
        speed_column = dataframes[0].columns.get_loc('Ship Speed (knots)')
        rt_column = dataframes[0].columns.get_loc('Rt (kN)')
        pe_column = dataframes[0].columns.get_loc('Pe(kW)')
        treq_column = dataframes[0].columns.get_loc('Treq (kN)')
    except KeyError as e:
        print(f"Column not found: {e}")
        return

    # Add series for Total Resistance (Rt) chart from all sheets
    for i, (row_start, row_end) in enumerate(start_end_rows):
        rt_chart.add_series({
            'name': f'Rt ({sheet_names[i]})',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, rt_column, row_end, rt_column],
            'line': {'dash_type': 'solid'}
        })

    # Add series for Effective Power (Pe) chart from all sheets
    for i, (row_start, row_end) in enumerate(start_end_rows):
        pe_chart.add_series({
            'name': f'Pe ({sheet_names[i]})',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, pe_column, row_end, pe_column],
            'line': {'dash_type': 'solid'}
        })

    # Add series for Thrust Required (Treq) chart from all sheets
    for i, (row_start, row_end) in enumerate(start_end_rows):
        treq_chart.add_series({
            'name': f'Treq ({sheet_names[i]})',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, treq_column, row_end, treq_column],
            'line': {'dash_type': 'solid'}
        })

    # Configure chart title, axes, and legend for Total Resistance (Rt) chart
    rt_chart.set_title({'name': 'Total Resistance (Rt) vs Ship Speed for Different Scenarios'})
    rt_chart.set_x_axis({'name': 'Ship Speed (knots)'})
    rt_chart.set_y_axis({'name': 'Total Resistance (Rt) (kN)'})
    rt_chart.set_legend({'position': 'bottom'})

    # Configure chart title, axes, and legend for Effective Power (Pe) chart
    pe_chart.set_title({'name': 'Effective Power (Pe) vs Ship Speed for Different Scenarios'})
    pe_chart.set_x_axis({'name': 'Ship Speed (knots)'})
    pe_chart.set_y_axis({'name': 'Effective Power (Pe) (kW)'})
    pe_chart.set_legend({'position': 'bottom'})

    # Configure chart title, axes, and legend for Thrust Required (Treq) chart
    treq_chart.set_title({'name': 'Thrust Required (Treq) vs Ship Speed for Different Scenarios'})
    treq_chart.set_x_axis({'name': 'Ship Speed (knots)'})
    treq_chart.set_y_axis({'name': 'Thrust Required (Treq) (kN)'})
    treq_chart.set_legend({'position': 'bottom'})

    # Insert the charts into the worksheet
    worksheet.insert_chart('J2', rt_chart)
    worksheet.insert_chart('J30', pe_chart)
    worksheet.insert_chart('J58', treq_chart)

    # Close the workbook
    workbook.close()

# Generate Excel file with interactive charts for the first five sheets
create_interactive_chart(input_file, output_file)
