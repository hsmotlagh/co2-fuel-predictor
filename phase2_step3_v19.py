import numpy as np
import pandas as pd
import os
from sklearn.linear_model import LinearRegression

# Load input data from Excel file
input_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/phase1-step1.xlsx"
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/phase2-step1_combined.xlsx"

# Load Kt coefficients from 'Coefficients' sheet
coefficients_data = pd.read_excel(input_path, sheet_name='Coefficients')
original_kt_coefficients = coefficients_data['Original Kt Coefficients'].values
advance_kt_coefficients = coefficients_data['Advance Kt Coefficients'].values

# Load Kq coefficients from 'Coefficients' sheet
original_kq_coefficients = coefficients_data['Original Kq Coefficients'].values
advance_kq_coefficients = coefficients_data['Advance Kq Coefficients'].values

# Define the dictionary containing Kt coefficients for different scenarios
kt_coefficients_dict = {
    "Original Scenario": original_kt_coefficients,
    "Paint (5%) Scenario": original_kt_coefficients,
    "Advance Propeller Scenario": advance_kt_coefficients,
    "Fin (2%-4%) Scenario": original_kt_coefficients,
    "Bulbous Bow Scenario": original_kt_coefficients,
    "Combined Paint and Bulbous Bow Scenario": original_kt_coefficients
}

# Define the dictionary containing Kq coefficients for different scenarios
kq_coefficients_dict = {
    "Original Scenario": original_kq_coefficients,
    "Paint (5%) Scenario": original_kq_coefficients,
    "Advance Propeller Scenario": advance_kq_coefficients,
    "Fin (2%-4%) Scenario": original_kq_coefficients,
    "Bulbous Bow Scenario": original_kq_coefficients,
    "Combined Paint and Bulbous Bow Scenario": original_kq_coefficients
}

# Define constants
density = 1025  # density of water in kg/m^3
propeller_diameter = 4  # propeller diameter in meters
transmission_efficiency = 0.965
sfoc = 220  # Specific Fuel Oil Consumption in g/kWh

# Define n (rpm) values for ship speeds 9 to 12.5 and 13 to 15
n_rpm_values_dict = {
    "low_speeds": np.arange(50, 180, 10),  # For ship speeds 9 to 12.5
    "high_speeds": np.arange(100, 230, 10)  # For ship speeds 13 to 15
}

# Load Va values for each scenario from the Excel file
scenarios = ["Original Scenario", "Paint (5%) Scenario", "Advance Propeller Scenario",
             "Fin (2%-4%) Scenario", "Bulbous Bow Scenario", "Combined Paint and Bulbous Bow Scenario"]

# Prepare DataFrame to store results for each scenario and each ship speed
phase2_step1_results = {}
thrust_coefficients_results = {}

for scenario in scenarios:
    # Load Va values from the corresponding sheet in the Excel file
    try:
        data = pd.read_excel(input_path, sheet_name=scenario)
    except ValueError:
        print(f"Worksheet named '{scenario}' not found in the input file.")
        continue

    va_values = data['Va (m/s)']
    treq_values = data['Treq (kN)']
    thrust_deduction_values = data['Thrust deduction']
    wake_fraction_values = data['Wake Fraction']
    relative_rotative_efficiency_values = data['Relative rotative efficiency']
    pe_values = data['Pe(kW)']

    # Get Kt coefficients for the scenario
    kt_coefficient_2, kt_coefficient_1, kt_coefficient_0 = kt_coefficients_dict[scenario]

    # Get Kq coefficients for the scenario
    kq_coefficient_2, kq_coefficient_1, kq_coefficient_0 = kq_coefficients_dict[scenario]

    # Prepare results for the current scenario
    scenario_results = []
    coefficients_results = []

    # Loop through each ship speed and calculate parameters
    for i, va in enumerate(va_values):
        ship_speed = data['Ship Speed (knots)'][i]
        if ship_speed not in [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15]:
            continue  # Only process specific ship speeds

        # Determine appropriate n (rpm) values based on ship speed
        if ship_speed in [9, 10, 11, 12, 12.5]:
            n_rpm_values = n_rpm_values_dict["low_speeds"]
        else:
            n_rpm_values = n_rpm_values_dict["high_speeds"]

        # Convert rpm to rps
        n_rps_values = n_rpm_values / 60

        speed_results = []
        thrust_values = []

        for n_rps in n_rps_values:
            # Calculate J (advance coefficient) using corrected formula
            J = va / (n_rps * propeller_diameter)

            # Calculate Kt using the coefficients from the input sheet
            Kt = (kt_coefficient_2 * J ** 2) + (kt_coefficient_1 * J) + kt_coefficient_0

            # Calculate propeller thrust (T) in kN
            T = (Kt * density * n_rps ** 2 * propeller_diameter ** 4) / 1000

            # Append the results for this speed
            speed_results.append([n_rpm_values[np.where(n_rps_values == n_rps)[0][0]], n_rps, J, Kt, T])
            thrust_values.append(T)

        # Calculate T^2, T^3, T^4 for each thrust value
        thrust_df = pd.DataFrame({
            'T (thrust)': thrust_values,
            'T^2': np.power(thrust_values, 2),
            'T^3': np.power(thrust_values, 3),
            'T^4': np.power(thrust_values, 4),
            'n (rpm)': n_rpm_values
        })

        # Fit polynomial to calculate coefficients for T, T^2, T^3, T^4 using Linear Regression
        X = thrust_df[['T (thrust)', 'T^2', 'T^3', 'T^4']]
        y = thrust_df['n (rpm)']

        reg = LinearRegression().fit(X, y)
        coefficients = [reg.intercept_] + reg.coef_.tolist()

        # Record the thrust coefficients for the scenario and speed
        coefficients_results.append(coefficients)

        # Append speed-specific results to scenario results
        scenario_results.append(speed_results)

    # Store results for the scenario
    phase2_step1_results[scenario] = scenario_results
    thrust_coefficients_results[scenario] = coefficients_results

# Save Phase 2, Step 1 results to Excel
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    for scenario in scenarios:
        if scenario not in phase2_step1_results:
            continue
        for i, speed_result in enumerate(phase2_step1_results[scenario]):
            ship_speed = [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15][i]
            sheet_name = f'{scenario[:15]}_S{ship_speed}kn'
            scenario_df = pd.DataFrame(speed_result, columns=["n (rpm)", "n (rps)", "J", "Kt", "T (kN)"])
            scenario_df['Thrust deduction'] = thrust_deduction_values.iloc[i]
            scenario_df['Wake Fraction'] = wake_fraction_values.iloc[i]
            # Calculate hull efficiency (ηH)
            scenario_df['Hull Efficiency (ηH)'] = (1 - scenario_df['Thrust deduction']) / (
                        1 - scenario_df['Wake Fraction'])
            # Calculate Kq using the coefficients from the input sheet
            scenario_df['Kq'] = ((kq_coefficient_2 * scenario_df['J'] ** 2) + (
                        kq_coefficient_1 * scenario_df['J']) + kq_coefficient_0) / 10
            # Calculate propeller open water efficiency (ηO)
            scenario_df['Propeller Open Water Efficiency (ηO)'] = (scenario_df['J'] / (2 * np.pi)) * (
                        scenario_df['Kt'] / scenario_df['Kq'])
            scenario_df['Pe (kW)'] = pe_values.iloc[i]
            # Calculate delivered power (Pd)
            scenario_df['Delivered Power (Pd) (kW)'] = scenario_df['Pe (kW)'] / (
                        scenario_df['Hull Efficiency (ηH)'] * scenario_df['Propeller Open Water Efficiency (ηO)'] *
                        relative_rotative_efficiency_values.iloc[i])
            # Calculate brake power (PB)
            scenario_df['Brake Power (PB) (kW)'] = scenario_df['Delivered Power (Pd) (kW)'] / transmission_efficiency
            # Calculate fuel consumption (FC)
            scenario_df['Fuel Consumption (FC) (kg/h)'] = (sfoc * scenario_df['Brake Power (PB) (kW)']) / 1000
            # Calculate CO2 emission (CO2)
            scenario_df['CO2 Emission (kg)'] = scenario_df['Fuel Consumption (FC) (kg/h)'] * 3.11
            scenario_df.to_excel(writer, sheet_name=sheet_name, index=False)

    # Save thrust coefficients to a separate sheet
    for scenario in thrust_coefficients_results:
        sheet_name = f'{scenario[:15]}_Coeff'
        treq_values = pd.read_excel(input_path, sheet_name=scenario)['Treq (kN)']
        va_values = pd.read_excel(input_path, sheet_name=scenario)['Va (m/s)']
        thrust_deduction_values = pd.read_excel(input_path, sheet_name=scenario)['Thrust deduction']
        wake_fraction_values = pd.read_excel(input_path, sheet_name=scenario)['Wake Fraction']
        relative_rotative_efficiency_values = pd.read_excel(input_path, sheet_name=scenario)[
            'Relative rotative efficiency']
        pe_values = pd.read_excel(input_path, sheet_name=scenario)['Pe(kW)']
        coeff_df = pd.DataFrame(thrust_coefficients_results[scenario],
                                columns=["c (constant)", "c T", "c T^2", "c T^3", "c T^4"])
        coeff_df['Ship Speed (knots)'] = [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15][:len(coeff_df)]
        coeff_df['Treq (thrust required)'] = treq_values[:len(coeff_df)].values
        coeff_df['Va (m/s)'] = va_values[:len(coeff_df)].values
        coeff_df['Thrust deduction'] = thrust_deduction_values[:len(coeff_df)].values
        coeff_df['Wake Fraction'] = wake_fraction_values[:len(coeff_df)].values
        coeff_df['Relative rotative efficiency'] = relative_rotative_efficiency_values[:len(coeff_df)].values
        coeff_df['Pe (kW)'] = pe_values[:len(coeff_df)].values

        # Calculate n (rpm) for each ship speed and scenario
        coeff_df['n (rpm)'] = (
                (coeff_df['c T^4'] * coeff_df['Treq (thrust required)'] ** 4) +
                (coeff_df['c T^3'] * coeff_df['Treq (thrust required)'] ** 3) +
                (coeff_df['c T^2'] * coeff_df['Treq (thrust required)'] ** 2) +
                (coeff_df['c T'] * coeff_df['Treq (thrust required)']) +
                coeff_df['c (constant)']
        )

        # Calculate n (rps) for each ship speed and scenario
        coeff_df['n (rps)'] = coeff_df['n (rpm)'] / 60

        # Calculate J for each ship speed and scenario
        coeff_df['J'] = coeff_df['Va (m/s)'] / (coeff_df['n (rps)'] * propeller_diameter)

        # Calculate Kt using the coefficients from the input sheet
        kt_coefficient_2, kt_coefficient_1, kt_coefficient_0 = kt_coefficients_dict[scenario]
        coeff_df['Kt'] = (kt_coefficient_2 * coeff_df['J'] ** 2) + (kt_coefficient_1 * coeff_df['J']) + kt_coefficient_0

        # Calculate T (kN) for each ship speed and scenario
        coeff_df['T (kN)'] = (density * coeff_df['Kt'] * (coeff_df['n (rps)'] ** 2) * (propeller_diameter ** 4)) / 1000

        # Calculate hull efficiency (ηH)
        coeff_df['Hull Efficiency (ηH)'] = (1 - coeff_df['Thrust deduction']) / (1 - coeff_df['Wake Fraction'])

        # Calculate Kq using the coefficients from the input sheet
        kq_coefficient_2, kq_coefficient_1, kq_coefficient_0 = kq_coefficients_dict[scenario]
        coeff_df['Kq'] = ((kq_coefficient_2 * coeff_df['J'] ** 2) + (
                    kq_coefficient_1 * coeff_df['J']) + kq_coefficient_0) / 10

        # Calculate propeller open water efficiency (ηO)
        coeff_df['Propeller Open Water Efficiency (ηO)'] = (coeff_df['J'] / (2 * np.pi)) * (
                    coeff_df['Kt'] / coeff_df['Kq'])

        # Calculate delivered power (Pd)
        coeff_df['Delivered Power (Pd) (kW)'] = coeff_df['Pe (kW)'] / (
                    coeff_df['Hull Efficiency (ηH)'] * coeff_df['Propeller Open Water Efficiency (ηO)'] * coeff_df[
                'Relative rotative efficiency'])

        # Calculate brake power (PB)
        coeff_df['Brake Power (PB) (kW)'] = coeff_df['Delivered Power (Pd) (kW)'] / transmission_efficiency
        # Calculate fuel consumption (FC)
        coeff_df['Fuel Consumption (FC) (kg/h)'] = (sfoc * coeff_df['Brake Power (PB) (kW)']) / 1000
        # Calculate CO2 emission (CO2)
        coeff_df['CO2 Emission (kg)'] = coeff_df['Fuel Consumption (FC) (kg/h)'] * 3.11
        coeff_df.to_excel(writer, sheet_name=sheet_name, index=False)

print(f"Results saved to {output_path}")