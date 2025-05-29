import pandas as pd
import os

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/phase2-step1_combined.xlsx"

# Hardcode provided FC and CO2 values at 12 knots
comparison_data = [
    ["Original Scenario", 12, 276.2684182, 859.1947807],
    ["Paint (5%) Scenario", 12, 263.4790258, 819.4197703],
    ["Bulbous Bow Scenario", 12, 284.2566267, 884.038109],
    ["Combined Paint and Bulbous Bow Scenario", 12, 260.12, 260.12 * 3.11]
]

# Save results to Excel
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    # Save comparison table
    comparison_df = pd.DataFrame(comparison_data, columns=["Scenario", "Ship Speed (knots)", "Fuel Consumption (kg/h)", "CO2 Emission (kg)"])
    comparison_df.to_excel(writer, sheet_name='Comparison_12knots', index=False)

    # Create interactive chart for FC
    workbook = writer.book
    worksheet = writer.sheets['Comparison_12knots']
    chart_fc = workbook.add_chart({'type': 'column'})
    chart_fc.add_series({
        'name': 'Fuel Consumption (kg/h)',
        'categories': ['Comparison_12knots', 1, 0, len(comparison_data), 0],
        'values': ['Comparison_12knots', 1, 2, len(comparison_data), 2],
    })
    chart_fc.set_title({'name': 'Fuel Consumption Comparison at 12 Knots'})
    chart_fc.set_x_axis({'name': 'Scenario'})
    chart_fc.set_y_axis({'name': 'Fuel Consumption (kg/h)'})
    chart_fc.set_size({'width': 720, 'height': 576})
    worksheet.insert_chart('F2', chart_fc)

    # Create interactive chart for CO2
    chart_co2 = workbook.add_chart({'type': 'column'})
    chart_co2.add_series({
        'name': 'CO2 Emission (kg)',
        'categories': ['Comparison_12knots', 1, 0, len(comparison_data), 0],
        'values': ['Comparison_12knots', 1, 3, len(comparison_data), 3],
    })
    chart_co2.set_title({'name': 'CO2 Emission Comparison at 12 Knots'})
    chart_co2.set_x_axis({'name': 'Scenario'})
    chart_co2.set_y_axis({'name': 'CO2 Emission (kg)'})
    chart_co2.set_size({'width': 720, 'height': 576})
    worksheet.insert_chart('F20', chart_co2)

print(f"Results saved to {output_path}")