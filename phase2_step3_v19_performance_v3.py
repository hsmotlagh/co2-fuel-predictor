import pandas as pd
import os

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/phase2-step1_performance_analysis_v3.xlsx"

# Hardcode provided FC, Treq, and assumed ηO values for all speeds
performance_data = [
    # Original Scenario
    ["Original Scenario", 9, 115.7696693, 0.464, 89.7286371],
    ["Original Scenario", 10, 157.0962029, 0.464, 109.5295012],
    ["Original Scenario", 11, 211.3112914, 0.464, 133.5905855],
    ["Original Scenario", 12, 276.2684182, 0.464, 160.0163891],
    ["Original Scenario", 12.5, 319.0039009, 0.464, 176.7602696],
    ["Original Scenario", 13, 371.3694974, 0.464, 196.140474],
    ["Original Scenario", 13.5, 434.4792596, 0.464, 218.767217],
    ["Original Scenario", 14, 511.7771533, 0.464, 245.5898251],
    ["Original Scenario", 14.5, 609.1134837, 0.464, 278.2940591],
    ["Original Scenario", 15, 814.8434842, 0.464, 346.3754643],
    # Paint (5%) Scenario
    ["Paint (5%) Scenario", 9, 110.0681824, 0.464, 85.98485858],
    ["Paint (5%) Scenario", 10, 149.4348682, 0.464, 104.9770907],
    ["Paint (5%) Scenario", 11, 201.2853668, 0.464, 128.1574232],
    ["Paint (5%) Scenario", 12, 263.4790258, 0.464, 153.6398644],
    ["Paint (5%) Scenario", 12.5, 304.5852633, 0.464, 169.8957494],
    ["Paint (5%) Scenario", 13, 354.9289312, 0.464, 188.7725699],
    ["Paint (5%) Scenario", 13.5, 415.9640576, 0.464, 210.8807421],
    ["Paint (5%) Scenario", 14, 490.9831319, 0.464, 237.1807377],
    ["Paint (5%) Scenario", 14.5, 585.731981, 0.464, 269.3372922],
    ["Paint (5%) Scenario", 15, 787.8448646, 0.464, 336.9054225],
    # Bulbous Bow Scenario
    ["Bulbous Bow Scenario", 9, 119.6541899, 0.464, 92.25005292],
    ["Bulbous Bow Scenario", 10, 162.3385399, 0.464, 112.6111415],
    ["Bulbous Bow Scenario", 11, 217.9322612, 0.464, 137.1429527],
    ["Bulbous Bow Scenario", 12, 284.2566267, 0.464, 163.9600144],
    ["Bulbous Bow Scenario", 12.5, 323.9702129, 0.464, 179.1046976],
    ["Bulbous Bow Scenario", 13, 371.0228762, 0.464, 195.9860825],
    ["Bulbous Bow Scenario", 13.5, 425.0311648, 0.464, 214.7546376],
    ["Bulbous Bow Scenario", 14, 501.2150135, 0.464, 241.3303844],
    ["Bulbous Bow Scenario", 14.5, 600.5953901, 0.464, 275.0420581],
    ["Bulbous Bow Scenario", 15, 863.1846317, 0.464, 363.1135134],
    # Advance Propeller Scenario
    ["Advance Propeller Scenario", 9, 103.7743379, 0.579, 89.7286371],
    ["Advance Propeller Scenario", 10, 140.8707293, 0.579, 109.5295012],
    ["Advance Propeller Scenario", 11, 189.5548816, 0.579, 133.5905855],
    ["Advance Propeller Scenario", 12, 247.890597, 0.579, 160.0163891],
    ["Advance Propeller Scenario", 12.5, 286.2080063, 0.579, 176.7602696],
    ["Advance Propeller Scenario", 13, 332.9373352, 0.579, 196.140474],
    ["Advance Propeller Scenario", 13.5, 389.4472384, 0.579, 218.767217],
    ["Advance Propeller Scenario", 14, 458.6313585, 0.579, 245.5898251],
    ["Advance Propeller Scenario", 14.5, 545.6961681, 0.579, 278.2940591],
    ["Advance Propeller Scenario", 15, 729.3067193, 0.579, 346.3754643],
    # Fin Installation Scenario
    ["Fin Installation Scenario", 9, 115.4654418, 0.464, 90.31160087],
    ["Fin Installation Scenario", 10, 156.6860586, 0.464, 110.237233],
    ["Fin Installation Scenario", 11, 210.7942821, 0.464, 134.449073],
    ["Fin Installation Scenario", 12, 275.6058519, 0.464, 161.0334424],
    ["Fin Installation Scenario", 12.5, 318.2599935, 0.464, 177.8713822],
    ["Fin Installation Scenario", 13, 370.516624, 0.464, 197.3597647],
    ["Fin Installation Scenario", 13.5, 433.5948752, 0.464, 220.1120268],
    ["Fin Installation Scenario", 14, 510.8733655, 0.464, 247.074198],
    ["Fin Installation Scenario", 14.5, 608.260007, 0.464, 279.9570986],
    ["Fin Installation Scenario", 15, 814.3657819, 0.464, 348.3751466],
    # Combined Paint and Bulbous Bow Scenario
    ["Combined Paint and Bulbous Bow Scenario", 9, 108.59, 0.464, 85.45],
    ["Combined Paint and Bulbous Bow Scenario", 10, 147.58, 0.464, 104.51],
    ["Combined Paint and Bulbous Bow Scenario", 11, 199.03, 0.464, 127.60],
    ["Combined Paint and Bulbous Bow Scenario", 12, 260.12, 0.464, 153.07],
    ["Combined Paint and Bulbous Bow Scenario", 12.5, 301.54, 0.464, 169.36],
    ["Combined Paint and Bulbous Bow Scenario", 13, 351.13, 0.464, 188.28],
    ["Combined Paint and Bulbous Bow Scenario", 13.5, 411.98, 0.464, 210.34],
    ["Combined Paint and Bulbous Bow Scenario", 14, 486.46, 0.464, 236.62],
    ["Combined Paint and Bulbous Bow Scenario", 14.5, 579.32, 0.464, 268.74],
    ["Combined Paint and Bulbous Bow Scenario", 15, 771.04, 0.464, 335.85]
]

# Create DataFrame
performance_df = pd.DataFrame(performance_data,
                              columns=["Scenario", "Ship Speed (knots)", "FC (kg/h)", "ηO", "Treq (kN)"])

# Calculate efficiency gains relative to Original Scenario for each speed
efficiency_data = []
for speed in performance_df["Ship Speed (knots)"].unique():
    df_speed = performance_df[performance_df["Ship Speed (knots)"] == speed]
    original_fc = df_speed[df_speed["Scenario"] == "Original Scenario"]["FC (kg/h)"].iloc[0]
    original_treq = df_speed[df_speed["Scenario"] == "Original Scenario"]["Treq (kN)"].iloc[0]

    for _, row in df_speed.iterrows():
        fc_reduction = ((original_fc - row["FC (kg/h)"]) / original_fc) * 100
        treq_reduction = ((original_treq - row["Treq (kN)"]) / original_treq) * 100
        efficiency_data.append([
            row["Scenario"], row["Ship Speed (knots)"], fc_reduction, treq_reduction
        ])

# Create efficiency DataFrame
efficiency_df = pd.DataFrame(efficiency_data,
                             columns=["Scenario", "Ship Speed (knots)", "FC Reduction (%)", "Treq Reduction (%)"])

# Save to Excel and generate charts
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    efficiency_df.to_excel(writer, sheet_name='Performance_Analysis', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Performance_Analysis']

    # Pivot data for charts
    pivot_fc = efficiency_df.pivot_table(index='Scenario', columns='Ship Speed (knots)', values='FC Reduction (%)')
    pivot_treq = efficiency_df.pivot_table(index='Scenario', columns='Ship Speed (knots)', values='Treq Reduction (%)')
    pivot_fc.to_excel(writer, sheet_name='Pivot_Data', startrow=0)
    pivot_treq.to_excel(writer, sheet_name='Pivot_Data', startrow=10)

    # Scenarios to plot (excluding Original)
    scenarios = ["Paint (5%) Scenario", "Bulbous Bow Scenario", "Advance Propeller Scenario",
                 "Fin Installation Scenario", "Combined Paint and Bulbous Bow Scenario"]
    colors = ['#2CA02C', '#D62728', '#FF7F0E', '#9467BD', '#8C564B']  # Green, Red, Orange, Purple, Brown
    speeds = [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15]

    # Chart for FC Reduction
    chart_fc = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios):
        chart_fc.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$B$1:$K$1',
            'values': f'=Pivot_Data!$B${i + 2}:$K${i + 2}',
            'fill': {'color': colors[i]},
            'gap': 10
        })
    chart_fc.set_title({'name': 'FC Reduction Across Speeds'})
    chart_fc.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_fc.set_y_axis({'name': 'FC Reduction (%)', 'major_gridlines': {'visible': True}, 'min': -10, 'max': 15})
    chart_fc.set_size({'width': 720, 'height': 576})
    chart_fc.set_legend({'position': 'top'})
    worksheet.insert_chart('I2', chart_fc)

    # Chart for Treq Reduction
    chart_treq = workbook.add_chart({'type': 'column'})
    for i, scenario in enumerate(scenarios):
        chart_treq.add_series({
            'name': scenario,
            'categories': '=Pivot_Data!$B$11:$K$11',
            'values': f'=Pivot_Data!$B${i + 12}:$K${i + 12}',
            'fill': {'color': colors[i]},
            'gap': 10
        })
    chart_treq.set_title({'name': 'Treq Reduction Across Speeds'})
    chart_treq.set_x_axis({'name': 'Ship Speed (knots)'})
    chart_treq.set_y_axis({'name': 'Treq Reduction (%)', 'major_gridlines': {'visible': True}, 'min': -10, 'max': 15})
    chart_treq.set_size({'width': 720, 'height': 576})
    chart_treq.set_legend({'position': 'top'})
    worksheet.insert_chart('I20', chart_treq)

print(f"Performance analysis saved to {output_path}")