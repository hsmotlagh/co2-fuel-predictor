import pandas as pd
import xlsxwriter
import os

# File paths for input and output
input_file = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/SNN/SNN.xlsx"
output_file = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/SNN_Interactive_Chart.xlsx"

# Function to create interactive Excel charts for five sheets of SNN data
def create_interactive_chart(input_file, output_file):
    # Load the data from the first five sheets of the Excel file
    sheet_names = ["Sheet 1 (Original scenario)", "Sheet 2 (Paint (5%) Scenario)", "Sheet 3 (Advanced Propeller Scenario)", "Sheet 4 (Fin (2%-4%) Scenario)", "Sheet 5 (Bulbous Bow Scenario)"]
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

    # Create a new chart object for Fuel Consumption (FC)
    fc_chart = workbook.add_chart({'type': 'line'})

    # Create a new chart object for CO2 Emission
    co2_chart = workbook.add_chart({'type': 'line'})

    # Get necessary columns for plotting from the first DataFrame (assuming all have the same structure)
    try:
        speed_column = dataframes[0].columns.get_loc('Ship Speed (knots)')
        original_fc_column = dataframes[0].columns.get_loc('Fuel Consumption (FC) (kg/h)')
        original_co2_column = dataframes[0].columns.get_loc('CO2 Emission (kg)')
        fc_before_ml_column = dataframes[0].columns.get_loc('Final Predicted Fuel Consumption NN Before MC (kg/h)')
        co2_before_ml_column = dataframes[0].columns.get_loc('Final Predicted CO2 Emission NN Before MC (kg)')
        fc_after_mc_column = dataframes[0].columns.get_loc('Final Predicted Fuel Consumption NN After MC (kg/h)')
        co2_after_mc_column = dataframes[0].columns.get_loc('Final Predicted CO2 Emission NN After MC (kg)')
    except KeyError as e:
        print(f"Column not found: {e}")
        return

    # Add series for Fuel Consumption (FC) chart from all sheets
    for i, (row_start, row_end) in enumerate(start_end_rows):
        fc_chart.add_series({
            'name': f'Original FC ({sheet_names[i]})',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, original_fc_column, row_end, original_fc_column],
            'line': {'dash_type': 'solid'}
        })
        fc_chart.add_series({
            'name': f'Predicted FC Before ML ({sheet_names[i]})',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, fc_before_ml_column, row_end, fc_before_ml_column],
            'line': {'dash_type': 'dash'}
        })
        fc_chart.add_series({
            'name': f'Predicted FC After MC ({sheet_names[i]})',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, fc_after_mc_column, row_end, fc_after_mc_column],
            'line': {'dash_type': 'dot'}
        })

    # Add series for CO2 Emission chart from all sheets
    for i, (row_start, row_end) in enumerate(start_end_rows):
        co2_chart.add_series({
            'name': f'Original CO2 ({sheet_names[i]})',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, original_co2_column, row_end, original_co2_column],
            'line': {'dash_type': 'solid'}
        })
        co2_chart.add_series({
            'name': f'Predicted CO2 Before ML ({sheet_names[i]})',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, co2_before_ml_column, row_end, co2_before_ml_column],
            'line': {'dash_type': 'dash'}
        })
        co2_chart.add_series({
            'name': f'Predicted CO2 After MC ({sheet_names[i]})',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, co2_after_mc_column, row_end, co2_after_mc_column],
            'line': {'dash_type': 'dot'}
        })

    # Configure chart title, axes, and legend for Fuel Consumption chart
    fc_chart.set_title({'name': 'Fuel Consumption Comparison Before and After ML(SNN) and after ML with Monte Carlo'})
    fc_chart.set_x_axis({'name': 'Ship Speed (knots)'})
    fc_chart.set_y_axis({'name': 'Fuel Consumption (kg/h)'})
    fc_chart.set_legend({'position': 'bottom'})

    # Configure chart title, axes, and legend for CO2 Emission chart
    co2_chart.set_title({'name': 'CO2 Emission Comparison Before and After ML(SNN) and after ML with Monte Carlo'})
    co2_chart.set_x_axis({'name': 'Ship Speed (knots)'})
    co2_chart.set_y_axis({'name': 'CO2 Emission (kg)'})
    co2_chart.set_legend({'position': 'bottom'})

    # Insert the charts into the worksheet
    worksheet.insert_chart('J2', fc_chart)
    worksheet.insert_chart('J30', co2_chart)

    # Close the workbook
    workbook.close()

# Generate Excel file with interactive charts for the first five sheets
create_interactive_chart(input_file, output_file)
