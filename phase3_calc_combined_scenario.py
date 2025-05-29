import pandas as pd
import numpy as np

# Define output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Combined_Scenario_Data.xlsx"

# Table X data for Combined Scenario
combined_data = [
    [9, 64.54, 298.91, 85.45, 108.59, 337.71],
    [10, 79.02, 406.47, 104.51, 147.58, 459.08],
    [11, 96.68, 547.37, 127.60, 199.03, 619.18],
    [12, 116.21, 717.23, 153.07, 260.12, 808.97],
    [12.5, 129.01, 829.45, 169.36, 301.54, 937.79],
    [13, 143.87, 961.24, 188.28, 351.13, 1092.01],
    [13.5, 161.23, 1119.50, 210.34, 411.98, 1281.26],
    [14, 181.84, 1309.58, 236.62, 486.46, 1512.90],
    [14.5, 206.82, 1542.61, 268.74, 579.32, 1801.67],
    [15, 260.44, 2009.38, 335.85, 771.04, 2397.93]
]

# Create DataFrame
data = pd.DataFrame(combined_data, columns=[
    'Ship Speed (knots)', 'Rt (kN)', 'Pe (kW)', 'Treq (kN)',
    'Fuel Consumption (FC) (kg/h)', 'CO2 Emission (kg)'
])

# Interpolate parameters from Paint and Bulbous Bow (assumed average)
# Paint and Bulbous Bow data for reference (from your provided results)
paint_data = {
    'Ship Speed (knots)': [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
    'n (rpm)': [89.50, 99.49, 109.79, 119.72, 125.23, 131.72, 138.61, 146.06, 154.24, 168.16],
    'T (kN)': [86.59, 106.85, 130.19, 154.10, 168.93, 188.75, 211.25, 237.84, 270.02, 336.88],
    'ηH': [1.0588, 1.0587, 1.0572, 1.0570, 1.0569, 1.0552, 1.0522, 1.0505, 1.0476, 1.0429],
    'ηO': [0.5817, 0.5821, 0.5819, 0.5830, 0.5826, 0.5801, 0.5774, 0.5737, 0.5687, 0.5538],
    'Pd (kW)': [482.80, 655.48, 882.91, 1155.71, 1336.02, 1556.85, 1824.57, 2153.63, 2569.23, 3455.77],
    'PB (kW)': [500.31, 679.25, 914.93, 1197.63, 1384.48, 1613.31, 1890.75, 2231.74, 2662.42, 3581.11]
}

bulbous_data = {
    'Ship Speed (knots)': [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
    'n (rpm)': [91.47, 101.59, 111.94, 121.92, 127.11, 133.21, 139.37, 146.83, 155.23, 172.23],
    'T (kN)': [93.14, 114.62, 138.95, 163.86, 177.64, 196.05, 215.16, 242.00, 275.70, 362.79],
    'ηH': [1.0520, 1.0519, 1.0504, 1.0503, 1.0503, 1.0487, 1.0458, 1.0442, 1.0413, 1.0370],
    'ηO': [0.5741, 0.5748, 0.5751, 0.5767, 0.5774, 0.5761, 0.5754, 0.5718, 0.5663, 0.5448],
    'Pd (kW)': [524.85, 712.08, 955.93, 1246.85, 1421.05, 1627.44, 1864.34, 2198.51, 2634.43, 3786.24],
    'PB (kW)': [543.88, 737.90, 990.60, 1292.08, 1472.59, 1686.47, 1931.96, 2278.25, 2729.98, 3923.57]
}

# Interpolate n, T, ηH, ηO, Pd, PB (average of Paint and Bulbous Bow)
for param in ['n (rpm)', 'T (kN)', 'ηH', 'ηO', 'Pd (kW)', 'PB (kW)']:
    paint_vals = np.array(paint_data[param])
    bulbous_vals = np.array(bulbous_data[param])
    combined_vals = (paint_vals + bulbous_vals) / 2
    data[param] = combined_vals

# Save to Excel
data.to_excel(output_path, sheet_name='Combined_Scen_Coeff', index=False)
print(f"Combined Scenario data saved to {output_path}")