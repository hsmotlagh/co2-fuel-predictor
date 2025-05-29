import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import os
from numpy.linalg import lstsq

# Define constants
waterline_length = 100  # meters
viscosity = 0.00000118831  # kinematic viscosity in m^2/s
density = 1025  # density of water in kg/m^3
wetted_surface_area = 2350  # wetted surface area in m^2
wake_fraction = 0.1  # placeholder value (to be updated)
thrust_deduction = 0.1  # placeholder value (to be updated)
model_factor_k = 0.27  # placeholder value (to be updated)

# Define speed range (knots)
speeds_knots = np.arange(9, 15.5, 0.5)

# Convert speed to m/s using Equation 1
speeds_ms = speeds_knots * 0.5144

# Prepare lists to store results
results = []

for V in speeds_ms:
    # Calculate Reynolds number using Equation 2
    Re = (V * waterline_length) / viscosity

    # Calculate frictional resistance coefficient using Equation 3
    CF = 0.075 / (np.log10(Re) - 2) ** 2

    # Calculate viscous resistance coefficient using Equation 4
    CV = (1 + model_factor_k) * CF

    # Assume Wave Resistance Coefficient (CW) - placeholder value, can be updated later
    CW = 0.002

    # Calculate total resistance coefficient (CT)
    CT = CV + CW

    # Calculate total resistance value using Equation 5
    RT = CT * (density / 2) * wetted_surface_area * V ** 2

    # Calculate effective power (Pe) using Equation 6
    Pe = RT * V

    # Calculate advance speed (VA) using Equation 7
    VA = V * (1 - wake_fraction)

    # Calculate thrust required (Treq) using Equation 8
    Treq = RT / (1 - thrust_deduction)

    # Append results to list
    results.append([V / 0.5144, Re, CF, CV, CT, RT, Pe, VA, Treq])

# Create DataFrame to store results
columns = ["Speed (knots)", "Reynolds Number", "Frictional Resistance Coefficient (CF)",
           "Viscous Resistance Coefficient (CV)", "Total Resistance Coefficient (CT)",
           "Total Resistance (RT)", "Effective Power (Pe)", "Advance Speed (VA)", "Thrust Required (Treq)"]
df = pd.DataFrame(results, columns=columns)

# Save results to Excel file
output_path = "/Users/hamid/Documents/CO2 PAPER/python code-co2 papper/phase1-step1.xlsx"

# Define data for propellers
J = np.array([0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5, 0.55, 0.6, 0.65, 0.7, 0.75, 0.8, 0.85])
original_kt = np.array([0.2993, 0.2882, 0.2762, 0.2632, 0.2493, 0.2346, 0.219, 0.2026, 0.1855, 0.1678, 0.1492, 0.1302, 0.1105, 0.0903, 0.0694, 0.0482, 0.0264])
original_kq = np.array([0.3696, 0.359, 0.3479, 0.336, 0.3234, 0.31, 0.2956, 0.2803, 0.2638, 0.2462, 0.2274, 0.2072, 0.1857, 0.1626, 0.138, 0.1118, 0.0838])
advance_kt = np.array([0.3151, 0.3034, 0.2907, 0.277, 0.2624, 0.2469, 0.2305, 0.2133, 0.1953, 0.1766, 0.1571, 0.137, 0.1163, 0.095, 0.0731, 0.0507, 0.0278])
advance_kq = np.array([0.35196, 0.34191, 0.33129, 0.32002, 0.30803, 0.29523, 0.28155, 0.26692, 0.25126, 0.2345, 0.21656, 0.19735, 0.17682, 0.15488, 0.13145, 0.10646, 0.07983])

# Create separate tables for original and advance propeller
original_propeller_data = pd.DataFrame({
    "J": J,
    "Kt": original_kt,
    "10*Kq": original_kq
})
advance_propeller_data = pd.DataFrame({
    "J": J,
    "Kt": advance_kt,
    "10*Kq": advance_kq
})

# Create additional tables for calculations
original_kt_data = pd.DataFrame({
    "J": J,
    "J^2": J**2,
    "Kt": original_kt
})
original_kq_data = pd.DataFrame({
    "J": J,
    "J^2": J**2,
    "10*Kq": original_kq
})
advance_kt_data = pd.DataFrame({
    "J": J,
    "J^2": J**2,
    "Kt": advance_kt
})
advance_kq_data = pd.DataFrame({
    "J": J,
    "J^2": J**2,
    "10*Kq": advance_kq
})

# Fit polynomial models for original propeller using least squares
A = np.vstack([J**2, J, np.ones(len(J))]).T
kt_orig_coeff, _, _, _ = lstsq(A, original_kt, rcond=None)
kq_orig_coeff, _, _, _ = lstsq(A, original_kq, rcond=None)

# Fit polynomial models for advance propeller using least squares
kt_adv_coeff, _, _, _ = lstsq(A, advance_kt, rcond=None)
kq_adv_coeff, _, _, _ = lstsq(A, advance_kq, rcond=None)

# Create DataFrames for coefficients
coefficients_data = pd.DataFrame({
    "Coefficient": ["cJ^2", "cJ", "c"],
    "Original Kt Coefficients": kt_orig_coeff,
    "Original Kq Coefficients": kq_orig_coeff,
    "Advance Kt Coefficients": kt_adv_coeff,
    "Advance Kq Coefficients": kq_adv_coeff
})

# Save all data to the Excel file
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Step 1 Results', index=False)
    original_propeller_data.to_excel(writer, sheet_name='Original Propeller Data', index=False)
    advance_propeller_data.to_excel(writer, sheet_name='Advance Propeller Data', index=False)
    original_kt_data.to_excel(writer, sheet_name='Original Kt Data', index=False)
    original_kq_data.to_excel(writer, sheet_name='Original Kq Data', index=False)
    advance_kt_data.to_excel(writer, sheet_name='Advance Kt Data', index=False)
    advance_kq_data.to_excel(writer, sheet_name='Advance Kq Data', index=False)
    coefficients_data.to_excel(writer, sheet_name='Coefficients', index=False)

# Plot graphs for step 1 results
plt.figure()
plt.plot(df["Speed (knots)"], df["Effective Power (Pe)"], marker='o')
plt.xlabel("Speed (knots)")
plt.ylabel("Effective Power (Pe) (W)")
plt.title("Effective Power vs Ship Speed")
plt.grid(True)
plt.savefig("/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/effective_power_vs_speed.png")

plt.figure()
plt.plot(df["Speed (knots)"], df["Total Resistance (RT)"], marker='o')
plt.xlabel("Speed (knots)")
plt.ylabel("Total Resistance (RT) (N)")
plt.title("Total Resistance vs Ship Speed")
plt.grid(True)
plt.savefig("/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/total_resistance_vs_speed.png")

plt.figure()
plt.plot(df["Speed (knots)"], df["Thrust Required (Treq)"], marker='o')
plt.xlabel("Speed (knots)")
plt.ylabel("Thrust Required (Treq) (N)")
plt.title("Thrust Required vs Ship Speed")
plt.grid(True)
plt.savefig("/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/thrust_required_vs_speed.png")

# Plot graphs for original propeller
plt.figure()
plt.plot(J, original_kt, marker='o', label="Original Kt")
plt.plot(J, kt_orig_coeff[0] * J**2 + kt_orig_coeff[1] * J + kt_orig_coeff[2], linestyle='--', label="Original Kt (fit)")
plt.xlabel("Advance Coefficient (J)")
plt.ylabel("Thrust Coefficient (Kt)")
plt.title("J vs Kt for Original Propeller")
plt.legend()
plt.grid(True)
plt.savefig("/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/original_propeller_J_vs_Kt.png")

plt.figure()
plt.plot(J, original_kq, marker='o', label="Original Kq")
plt.plot(J, kq_orig_coeff[0] * J**2 + kq_orig_coeff[1] * J + kq_orig_coeff[2], linestyle='--', label="Original Kq (fit)")
plt.xlabel("Advance Coefficient (J)")
plt.ylabel("Torque Coefficient (10*Kq)")
plt.title("J vs Kq for Original Propeller")
plt.legend()
plt.grid(True)
plt.savefig("/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/original_propeller_J_vs_Kq.png")

# Plot graphs for advance propeller
plt.figure()
plt.plot(J, advance_kt, marker='o', label="Advance Kt")
plt.plot(J, kt_adv_coeff[0] * J**2 + kt_adv_coeff[1] * J + kt_adv_coeff[2], linestyle='--', label="Advance Kt (fit)")
plt.xlabel("Advance Coefficient (J)")
plt.ylabel("Thrust Coefficient (Kt)")
plt.title("J vs Kt for Advance Propeller")
plt.legend()
plt.grid(True)
plt.savefig("/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/advance_propeller_J_vs_Kt.png")

plt.figure()
plt.plot(J, advance_kq, marker='o', label="Advance Kq")
plt.plot(J, kq_adv_coeff[0] * J**2 + kq_adv_coeff[1] * J + kq_adv_coeff[2], linestyle='--', label="Advance Kq (fit)")
plt.xlabel("Advance Coefficient (J)")
plt.ylabel("Torque Coefficient (10*Kq)")
plt.title("J vs Kq for Advance Propeller")
plt.legend()
plt.grid(True)
plt.savefig("/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_1_graphs/advance_propeller_J_vs_Kq.png")

plt.show()
