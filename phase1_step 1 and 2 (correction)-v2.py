# Define scenarios and calculations
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import os
from numpy.linalg import lstsq

# Define constants
viscosity = 0.00000118831  # kinematic viscosity in m^2/s
density = 1025  # density of water in kg/m^3
wetted_surface_area = 2350  # wetted surface area in m^2 for all scenarios except bulbous bow
model_factor_k = 0.27  # updated model factor value
relative_rotative_efficiency = 1.012  # given for all scenarios

# Define speed range (knots)
speeds_knots = np.array([9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15])

# Convert speed to m/s using Equation 1
speeds_ms = speeds_knots * 0.5144

# Tanker ship model test results
data = {
    "Ship Speed (knots)": speeds_knots,
    "Wave Resistance Coefficient (without bulbous bow) * 1000": [0.435, 0.439, 0.49, 0.538, 0.604, 0.692, 0.805, 0.953, 1.141, 1.701],
    "Wave Resistance Coefficient (with bulbous bow) * 1000": [0.444, 0.448, 0.495, 0.538, 0.574, 0.623, 0.684, 0.829, 1.027, 1.786],
    "Thrust Deduction": [0.244, 0.243, 0.242, 0.24, 0.238, 0.236, 0.234, 0.231, 0.229, 0.223],
    "Wake Fraction": [0.286, 0.285, 0.283, 0.281, 0.279, 0.276, 0.272, 0.268, 0.264, 0.255]
}
model_test_results = pd.DataFrame(data)

# Define scenarios
scenarios = [
    "Original Scenario",
    "Paint (5%) Scenario",
    "Advance Propeller Scenario",
    "Fin (2%-4%) Scenario",
    "Bulbous Bow Scenario"
]

# Prepare DataFrames to store results for each scenario
scenario_results = {scenario: [] for scenario in scenarios}

# Define the output path for Excel file
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/phase1-step1.xlsx"

# Loop through each speed and calculate parameters for each scenario
for i, V in enumerate(speeds_ms):
    wave_resistance_coeff_without_bulb = model_test_results.loc[i, "Wave Resistance Coefficient (without bulbous bow) * 1000"] / 1000
    wave_resistance_coeff_with_bulb = model_test_results.loc[i, "Wave Resistance Coefficient (with bulbous bow) * 1000"] / 1000
    thrust_deduction = model_test_results.loc[i, "Thrust Deduction"]
    wake_fraction = model_test_results.loc[i, "Wake Fraction"]
    waterline_length = 100  # Default waterline length for all scenarios except bulbous bow
    wetted_surface_area_scenario = wetted_surface_area  # Default wetted surface area for all scenarios

    for scenario in scenarios:
        # Set wave resistance coefficient, thrust deduction, wake fraction, waterline length, and wetted surface area based on the scenario
        if scenario == "Original Scenario":
            wave_resistance_coeff = wave_resistance_coeff_without_bulb
            CF_multiplier = 1.0
        elif scenario == "Paint (5%) Scenario":
            wave_resistance_coeff = wave_resistance_coeff_without_bulb
            CF_multiplier = 0.95  # Apply 5% reduction for Paint (5%)
        elif scenario == "Advance Propeller Scenario":
            wave_resistance_coeff = wave_resistance_coeff_without_bulb
            CF_multiplier = 1.0
        elif scenario == "Fin (2%-4%) Scenario":
            wave_resistance_coeff = wave_resistance_coeff_without_bulb
            thrust_deduction *= 1.02
            wake_fraction *= 1.04  # Multiply wake fraction by 1.04 for Fin (2%-4%) Scenario
            CF_multiplier = 1.0
        elif scenario == "Bulbous Bow Scenario":
            wave_resistance_coeff = wave_resistance_coeff_with_bulb
            waterline_length = 103
            wetted_surface_area_scenario = 2400  # Wetted surface area for bulbous bow scenario
            wake_fraction = [0.286, 0.285, 0.283, 0.281, 0.279, 0.276, 0.272, 0.268, 0.264, 0.255][i]  # Specific wake fraction for Bulbous Bow Scenario
            CF_multiplier = 1.0

        # Calculate Reynolds number using corrected Equation 2
        Re = (V * waterline_length) / viscosity

        # Calculate frictional resistance coefficient using Equation 3
        CF = (0.075 / (np.log10(Re) - 2) ** 2) * CF_multiplier

        # Calculate viscous resistance coefficient using Equation 4
        CV = (1 + model_factor_k) * CF

        # Calculate total resistance coefficient (CT) using updated formula
        CT = CV + wave_resistance_coeff

        # Calculate total resistance value using Equation 5
        RT = (density / 2) * wetted_surface_area_scenario * V ** 2 * CT  # updated formula for RT in kN

        # Calculate effective power (Pe) using updated formula (Equation 6)
        Pe = RT * V

        # Calculate advance speed (VA) using corrected Equation 7
        VA = V * (1 - wake_fraction)

        # Calculate thrust required (Treq) using corrected Equation 8
        Treq = RT / (1 - thrust_deduction)

        # Append results to scenario list
        scenario_results[scenario].append([
            speeds_knots[i], wave_resistance_coeff * 1000, thrust_deduction, wake_fraction, relative_rotative_efficiency,
            V, Re, CF, CV, CT, RT / 1000, Pe / 1000, VA, Treq / 1000
        ])

# Update propeller data based on new values
original_propeller_data = pd.DataFrame({
    "J": [0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5, 0.55, 0.6, 0.65, 0.7, 0.75, 0.8, 0.85],
    "Kt": [0.2993, 0.2882, 0.2762, 0.2632, 0.2493, 0.2346, 0.219, 0.2026, 0.1855, 0.1678, 0.1492, 0.1302, 0.1105, 0.0903, 0.0694, 0.0482, 0.0264],
    "10*Kq": [0.3696, 0.359, 0.3479, 0.336, 0.3234, 0.31, 0.2956, 0.2803, 0.2638, 0.2462, 0.2274, 0.2072, 0.1857, 0.1626, 0.138, 0.1118, 0.0838]
})

advance_propeller_data = pd.DataFrame({
    "J": [0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5, 0.55, 0.6, 0.65, 0.7, 0.75, 0.8, 0.85],
    "Kt": [0.3151, 0.3034, 0.2907, 0.277, 0.2624, 0.2469, 0.2305, 0.2133, 0.1953, 0.1766, 0.1571, 0.137, 0.1163, 0.095, 0.0731, 0.0507, 0.0278],
    "10*Kq": [0.35196, 0.34191, 0.33129, 0.32002, 0.30803, 0.29523, 0.28155, 0.26692, 0.25126, 0.2345, 0.21656, 0.19735, 0.17682, 0.15488, 0.13145, 0.10646, 0.07983]
})

# Update original Kt and Kq data, advance Kt and Kq data based on new propeller data
original_kt_data = pd.DataFrame({
    "J": original_propeller_data["J"],
    "J^2": original_propeller_data["J"] ** 2,
    "Kt": original_propeller_data["Kt"]
})

original_kq_data = pd.DataFrame({
    "J": original_propeller_data["J"],
    "J^2": original_propeller_data["J"] ** 2,
    "10*Kq": original_propeller_data["10*Kq"]
})

advance_kt_data = pd.DataFrame({
    "J": advance_propeller_data["J"],
    "J^2": advance_propeller_data["J"] ** 2,
    "Kt": advance_propeller_data["Kt"]
})

advance_kq_data = pd.DataFrame({
    "J": advance_propeller_data["J"],
    "J^2": advance_propeller_data["J"] ** 2,
    "10*Kq": advance_propeller_data["10*Kq"]
})

# Calculate coefficients for Kt and Kq using least squares fit
original_kt_coeff = lstsq(np.vstack([original_kt_data["J^2"], original_kt_data["J"], np.ones(len(original_kt_data))]).T, original_kt_data["Kt"], rcond=None)[0]
original_kq_coeff = lstsq(np.vstack([original_kq_data["J^2"], original_kq_data["J"], np.ones(len(original_kq_data))]).T, original_kq_data["10*Kq"], rcond=None)[0]
advance_kt_coeff = lstsq(np.vstack([advance_kt_data["J^2"], advance_kt_data["J"], np.ones(len(advance_kt_data))]).T, advance_kt_data["Kt"], rcond=None)[0]
advance_kq_coeff = lstsq(np.vstack([advance_kq_data["J^2"], advance_kq_data["J"], np.ones(len(advance_kq_data))]).T, advance_kq_data["10*Kq"], rcond=None)[0]

# Create coefficients DataFrame
coefficients_data = pd.DataFrame({
    "Coefficient": ["cJ^2", "cJ", "c"],
    "Original Kt Coefficients": original_kt_coeff,
    "Original Kq Coefficients": original_kq_coeff,
    "Advance Kt Coefficients": advance_kt_coeff,
    "Advance Kq Coefficients": advance_kq_coeff
})

# Save updated data to Excel
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    for scenario in scenarios:
        columns = [
            "Ship Speed (knots)", "Wave Resistance Coefficient * 1000", "Thrust deduction", "Wake Fraction",
            "Relative rotative efficiency", "Speed m/s", "Re", "CF", "Cv", "Ct", "Rt (kN)", "Pe(kW)",
            "Va (m/s)", "Treq (kN)"
        ]
        scenario_df = pd.DataFrame(scenario_results[scenario], columns=columns)
        scenario_df.to_excel(writer, sheet_name=f'{scenario}', index=False)

        # Set column width to 20
        worksheet = writer.sheets[f'{scenario}']
        for idx, col in enumerate(scenario_df.columns):
            worksheet.set_column(idx, idx, 20)

    # Save propeller data to Excel
    original_propeller_data.to_excel(writer, sheet_name='Original Propeller Data', index=False)
    advance_propeller_data.to_excel(writer, sheet_name='Advance Propeller Data', index=False)
    original_kt_data.to_excel(writer, sheet_name='Original Kt Data', index=False)
    original_kq_data.to_excel(writer, sheet_name='Original Kq Data', index=False)
    advance_kt_data.to_excel(writer, sheet_name='Advance Kt Data', index=False)
    advance_kq_data.to_excel(writer, sheet_name='Advance Kq Data', index=False)
    coefficients_data.to_excel(writer, sheet_name='Coefficients', index=False)

# Plot graphs for step 1 results
for scenario in scenarios:
    scenario_df = pd.DataFrame(scenario_results[scenario], columns=columns)

    plt.figure()
    plt.plot(scenario_df["Ship Speed (knots)"], scenario_df["Pe(kW)"], marker='o')
    plt.xlabel("Speed (knots)")
    plt.ylabel("Effective Power (Pe) (kW)")
    plt.title(f"Effective Power vs Ship Speed for {scenario}")
    plt.grid(True)
    plt.savefig(f"/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/{scenario.lower().replace(' ', '_')}_effective_power_vs_speed.png")

    plt.figure()
    plt.plot(scenario_df["Ship Speed (knots)"], scenario_df["Rt (kN)"], marker='o')
    plt.xlabel("Speed (knots)")
    plt.ylabel("Total Resistance (RT) (kN)")
    plt.title(f"Total Resistance vs Ship Speed for {scenario}")
    plt.grid(True)
    plt.savefig(f"/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/{scenario.lower().replace(' ', '_')}_total_resistance_vs_speed.png")

    plt.figure()
    plt.plot(scenario_df["Ship Speed (knots)"], scenario_df["Treq (kN)"], marker='o')
    plt.xlabel("Speed (knots)")
    plt.ylabel("Thrust Required (Treq) (kN)")
    plt.title(f"Thrust Required vs Ship Speed for {scenario}")
    plt.grid(True)
    plt.savefig(f"/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/{scenario.lower().replace(' ', '_')}_thrust_required_vs_speed.png")

plt.show()
