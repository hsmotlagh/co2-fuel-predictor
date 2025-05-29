import numpy as np
import pandas as pd
import os
from sklearn.linear_model import LinearRegression

# Define input and output paths
input_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/phase1-step1.xlsx"
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/phase2-step1_base.xlsx"

# Load Kt and Kq coefficients
coefficients_data = pd.read_excel(input_path, sheet_name='Coefficients')
original_kt_coefficients = coefficients_data['Original Kt Coefficients'].values
original_kq_coefficients = coefficients_data['Original Kq Coefficients'].values

# Define coefficient dictionaries
kt_coefficients_dict = {
    "Original Scenario": original_kt_coefficients,
    "Paint (5%) Scenario": original_kt_coefficients,
    "Bulbous Bow Scenario": original_kt_coefficients
}

kq_coefficients_dict = {
    "Original Scenario": original_kq_coefficients,
    "Paint (5%) Scenario": original_kq_coefficients,
    "Bulbous Bow Scenario": original_kq_coefficients
}

# Define constants
density = 1025  # kg/m^3
propeller_diameter = 4  # meters
transmission_efficiency = 0.965
sfoc = 220  # g/kWh
model_factor_k = 0.27
viscosity = 0.00000118831  # m^2/s
waterline_length_default = 100  # m
waterline_length_bulbous = 103  # m
wetted_surface_area_default = 2350  # m²
wetted_surface_area_bulbous = 2400  # m²
required_columns = [
    'Ship Speed (knots)', 'Speed m/s', 'CF', 'Thrust deduction',
    'Wake Fraction', 'Relative rotative efficiency', 'Va (m/s)'
]

# Hardcode Bulbous Bow wave resistance coefficients (divided by 1000)
bulbous_wave_resistance = np.array([0.444, 0.448, 0.495, 0.538, 0.574, 0.623, 0.684, 0.829, 1.027, 1.786]) / 1000

# Define n (rpm) values
n_rpm_values_dict = {
    "low_speeds": np.arange(50, 180, 10),
    "high_speeds": np.arange(100, 230, 10)
}

# Define scenarios
scenarios = ["Original Scenario", "Paint (5%) Scenario", "Bulbous Bow Scenario"]

# Initialize results storage
phase2_step1_results = {}
thrust_coefficients_results = {}

for scenario in scenarios:
    # Load Original Scenario data
    try:
        data = pd.read_excel(input_path, sheet_name='Original Scenario')
    except ValueError:
        print(f"Error: Worksheet 'Original Scenario' not found in {input_path}.")
        continue

    # Verify columns
    missing_cols = [col for col in required_columns if col not in data.columns]
    if missing_cols:
        print(f"Error: Missing columns {missing_cols} in 'Original Scenario' sheet.")
        continue

    va_values = data['Va (m/s)'].values
    thrust_deduction_values = data['Thrust deduction'].values
    wake_fraction_values = data['Wake Fraction'].values
    relative_rotative_efficiency_values = data['Relative rotative efficiency'].values
    ship_speeds = data['Ship Speed (knots)'].values
    speed_ms_values = data['Speed m/s'].values

    # Apply scenario-specific parameters
    cf_multiplier = 1.0
    wetted_surface_area = wetted_surface_area_default
    waterline_length = waterline_length_default
    wave_resistance_values = np.zeros_like(speed_ms_values)
    if scenario == "Paint (5%) Scenario":
        cf_multiplier = 0.95
    elif scenario == "Bulbous Bow Scenario":
        wetted_surface_area = wetted_surface_area_bulbous
        waterline_length = waterline_length_bulbous
        wave_resistance_values = bulbous_wave_resistance

    # Calculate resistance and power
    re_values = (speed_ms_values * waterline_length) / viscosity
    cf_values = 0.075 / (np.log10(re_values) - 2) ** 2 * cf_multiplier
    cv_values = (1 + model_factor_k) * cf_values
    ct_values = cv_values + wave_resistance_values
    rt_values = ct_values * (density / 2) * wetted_surface_area * (speed_ms_values ** 2) / 1000  # kN
    pe_values = rt_values * speed_ms_values  # kW
    treq_values = rt_values / (1 - thrust_deduction_values)  # kN

    kt_coefficient_2, kt_coefficient_1, kt_coefficient_0 = kt_coefficients_dict[scenario]
    kq_coefficient_2, kq_coefficient_1, kq_coefficient_0 = kq_coefficients_dict[scenario]

    scenario_results = []
    coefficients_results = []

    for i, va in enumerate(va_values):
        ship_speed = ship_speeds[i]
        if ship_speed not in [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15]:
            continue

        n_rpm_values = n_rpm_values_dict["low_speeds"] if ship_speed <= 12.5 else n_rpm_values_dict["high_speeds"]
        n_rps_values = n_rpm_values / 60

        speed_results = []
        thrust_values = []

        for n_rps in n_rps_values:
            J = va / (n_rps * propeller_diameter)
            Kt = (kt_coefficient_2 * J ** 2) + (kt_coefficient_1 * J) + kt_coefficient_0
            T = (Kt * density * n_rps ** 2 * propeller_diameter ** 4) / 1000  # kN
            speed_results.append([n_rpm_values[np.where(n_rps_values == n_rps)[0][0]], n_rps, J, Kt, T])
            thrust_values.append(T)

        thrust_df = pd.DataFrame({
            'T (thrust)': thrust_values,
            'T^2': np.power(thrust_values, 2),
            'T^3': np.power(thrust_values, 3),
            'T^4': np.power(thrust_values, 4),
            'n (rpm)': n_rpm_values
        })

        X = thrust_df[['T (thrust)', 'T^2', 'T^3', 'T^4']]
        y = thrust_df['n (rpm)']
        reg = LinearRegression().fit(X, y)
        coefficients = [reg.intercept_] + reg.coef_.tolist()

        coefficients_results.append(coefficients)
        scenario_results.append(speed_results)

    phase2_step1_results[scenario] = scenario_results
    thrust_coefficients_results[scenario] = coefficients_results

# Save results to Excel
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    comparison_data = []
    for scenario in scenarios:
        if scenario not in phase2_step1_results:
            print(f"Warning: No results for '{scenario}'.")
            continue
        for i, speed_result in enumerate(phase2_step1_results[scenario]):
            ship_speed = [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15][i]
            sheet_name = f'{scenario[:15]}_S{ship_speed}kn'
            scenario_df = pd.DataFrame(speed_result, columns=["n (rpm)", "n (rps)", "J", "Kt", "T (kN)"])
            scenario_df['Thrust deduction'] = thrust_deduction_values[i]
            scenario_df['Wake Fraction'] = wake_fraction_values[i]
            scenario_df['Hull Efficiency (ηH)'] = (1 - scenario_df['Thrust deduction']) / (
                        1 - scenario_df['Wake Fraction'])
            scenario_df['Kq'] = ((kq_coefficient_2 * scenario_df['J'] ** 2) + (
                        kq_coefficient_1 * scenario_df['J']) + kq_coefficient_0) / 10
            scenario_df['Propeller Open Water Efficiency (ηO)'] = (scenario_df['J'] / (2 * np.pi)) * (
                        scenario_df['Kt'] / scenario_df['Kq'])
            scenario_df['Pe (kW)'] = pe_values[i]
            scenario_df['Delivered Power (Pd) (kW)'] = scenario_df['Pe (kW)'] / (
                        scenario_df['Hull Efficiency (ηH)'] * scenario_df['Propeller Open Water Efficiency (ηO)'] *
                        relative_rotative_efficiency_values[i])
            scenario_df['Brake Power (PB) (kW)'] = scenario_df['Delivered Power (Pd) (kW)'] / transmission_efficiency
            scenario_df['Fuel Consumption (FC) (kg/h)'] = (sfoc * scenario_df['Brake Power (PB) (kW)']) / 1000
            scenario_df['CO2 Emission (kg)'] = scenario_df['Fuel Consumption (FC) (kg/h)'] * 3.11
            scenario_df.to_excel(writer, sheet_name=sheet_name, index=False)

            if ship_speed == 12:
                mean_fc = scenario_df['Fuel Consumption (FC) (kg/h)'].mean()
                mean_co2 = scenario_df['CO2 Emission (kg)'].mean()
                comparison_data.append([scenario, ship_speed, mean_fc, mean_co2])

    # Save comparison table
    comparison_df = pd.DataFrame(comparison_data, columns=["Scenario", "Ship Speed (knots)", "Fuel Consumption (kg/h)",
                                                           "CO2 Emission (kg)"])
    comparison_df.to_excel(writer, sheet_name='Comparison_12knots', index=False)

    # Create interactive chart
    workbook = writer.book
    worksheet = writer.sheets['Comparison_12knots']
    chart = workbook.add_chart({'type': 'column'})
    chart.add_series({
        'name': 'Fuel Consumption (kg/h)',
        'categories': ['Comparison_12knots', 1, 0, len(comparison_data), 0],
        'values': ['Comparison_12knots', 1, 2, len(comparison_data), 2],
    })
    chart.set_title({'name': 'Fuel Consumption Comparison at 12 Knots'})
    chart.set_x_axis({'name': 'Scenario'})
    chart.set_y_axis({'name': 'Fuel Consumption (kg/h)'})
    chart.set_size({'width': 720, 'height': 576})
    worksheet.insert_chart('F2', chart)

print(f"Results saved to {output_path}")