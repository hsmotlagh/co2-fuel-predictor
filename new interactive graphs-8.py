import pandas as pd
import xlsxwriter
import os

# File paths for input and output
input_file = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML-RESULT/phase2.xlsx"
output_file = "/Users/hamid/Documents/CO2 PAPER FINAL/REPORT/graphs/BF-ML_Interactive_Chart-4.xlsx"

# Function to create interactive Excel charts for five sheets of new data
def create_interactive_chart(input_file, output_file):
    # Load the data from the first five sheets of the Excel file
    sheet_names = [
        "Original Scenar_Coeff", 
        "Paint (5%) Scen_Coeff", 
        "Advance Propell_Coeff", 
        "Fin (2%-4%) Sce_Coeff", 
        "Bulbous Bow Sce_Coeff"
    ]
    dataframes = [pd.read_excel(input_file, sheet_name=name) for name in sheet_names]
    
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

    # Create chart objects for the graphs
    fuel_chart = workbook.add_chart({'type': 'line'})
    co2_chart = workbook.add_chart({'type': 'line'})
    performance_chart = workbook.add_chart({'type': 'line'})
    total_efficiency_chart = workbook.add_chart({'type': 'line'})

    # Get necessary columns for plotting from the first DataFrame (assuming all have the same structure)
    try:
        speed_column = dataframes[0].columns.get_loc('Ship Speed (knots)')
        fuel_column = dataframes[0].columns.get_loc('Fuel Consumption (FC) (kg/h)')
        co2_column = dataframes[0].columns.get_loc('CO2 Emission (kg)')
        hull_efficiency_column = dataframes[0].columns.get_loc('Hull Efficiency (ηH)')
        relative_rotative_efficiency_column = dataframes[0].columns.get_loc('Relative rotative efficiency')
    except KeyError as e:
        print(f"Column not found: {e}")
        return

    # Add series for Fuel Consumption chart from all sheets
    symbols = ['circle', 'square', 'diamond', 'triangle', 'x']
    for i, (row_start, row_end) in enumerate(start_end_rows):
        fuel_chart.add_series({
            'name': sheet_names[i],
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, fuel_column, row_end, fuel_column],
            'line': {'dash_type': 'solid'},
            'marker': {'type': symbols[i], 'size': 6}
        })

    # Add series for CO2 Emission chart from all sheets
    for i, (row_start, row_end) in enumerate(start_end_rows):
        co2_chart.add_series({
            'name': sheet_names[i],
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, co2_column, row_end, co2_column],
            'line': {'dash_type': 'solid'},
            'marker': {'type': symbols[i], 'size': 6}
        })

    # Add series for Performance Analysis chart (combining Fuel and CO2 for each scenario)
    for i, (row_start, row_end) in enumerate(start_end_rows):
        performance_chart.add_series({
            'name': f'{sheet_names[i]} Fuel Consumption',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, fuel_column, row_end, fuel_column],
            'line': {'dash_type': 'dash'},
            'marker': {'type': symbols[i], 'size': 6}
        })
        performance_chart.add_series({
            'name': f'{sheet_names[i]} CO2 Emission',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, co2_column, row_end, co2_column],
            'line': {'dash_type': 'solid'},
            'marker': {'type': symbols[i], 'size': 6}
        })

    # Add series for Total Efficiency chart by summing efficiencies
    for i, (row_start, row_end) in enumerate(start_end_rows):
        total_efficiency_series = (dataframes[i]['Hull Efficiency (ηH)'] + dataframes[i]['Relative rotative efficiency'])
        worksheet.write_column(row_start, len(dataframes[0].columns), total_efficiency_series)
        total_efficiency_chart.add_series({
            'name': f'{sheet_names[i]} Total Efficiency',
            'categories': [worksheet.name, row_start, speed_column, row_end, speed_column],
            'values': [worksheet.name, row_start, len(dataframes[0].columns), row_end, len(dataframes[0].columns)],
            'line': {'dash_type': 'solid'},
            'marker': {'type': symbols[i], 'size': 6}
        })

    # Configure chart title, axes, and legend for Fuel Consumption chart
    fuel_chart.set_title({'name': 'Comparative Fuel Consumption Trends for Original and Modified Scenarios'})
    fuel_chart.set_x_axis({'name': 'Ship Speed (knots)'})
    fuel_chart.set_y_axis({'name': 'Fuel Consumption (kg/h)'})
    fuel_chart.set_legend({'position': 'bottom'})

    # Configure chart title, axes, and legend for CO2 Emission chart
    co2_chart.set_title({'name': 'CO2 Emission Trends Across Scenarios'})
    co2_chart.set_x_axis({'name': 'Ship Speed (knots)'})
    co2_chart.set_y_axis({'name': 'CO2 Emission (kg)'})
    co2_chart.set_legend({'position': 'bottom'})

    # Configure chart title, axes, and legend for Performance Analysis chart
    performance_chart.set_title({'name': 'Scenario-Specific Performance Analysis for Technical Modifications'})
    performance_chart.set_x_axis({'name': 'Ship Speed (knots)'})
    performance_chart.set_y_axis({'name': 'Performance Metrics'})
    performance_chart.set_legend({'position': 'bottom'})

    # Configure chart title, axes, and legend for Total Efficiency chart
    total_efficiency_chart.set_title({'name': 'Total Efficiency Across Different Scenarios'})
    total_efficiency_chart.set_x_axis({'name': 'Ship Speed (knots)'})
    total_efficiency_chart.set_y_axis({'name': 'Total Efficiency'})
    total_efficiency_chart.set_legend({'position': 'bottom'})

    # Insert the charts into the worksheet
    worksheet.insert_chart('J2', fuel_chart)
    worksheet.insert_chart('J30', co2_chart)
    worksheet.insert_chart('J58', performance_chart)
    worksheet.insert_chart('J86', total_efficiency_chart)

    # Close the workbook
    workbook.close()

# Generate Excel file with interactive charts for the first five sheets
create_interactive_chart(input_file, output_file)
