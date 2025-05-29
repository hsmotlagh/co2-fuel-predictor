import pandas as pd
import matplotlib.pyplot as plt
import os

# File output path
output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/"

# Data for different scenarios
scenarios = {
    "Original": {
        "Ship Speed (knots)": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "Fuel Consumption (FC) (kg/h)": [115.7696693, 157.0962029, 211.3112914, 276.2684182, 319.0039009, 371.3694974, 434.4792596, 511.7771533, 609.1134837, 814.8434842],
        "CO2 Emission (kg)": [360.0436715, 488.5691911, 657.1781162, 859.1947807, 992.1021319, 1154.959137, 1351.230497, 1591.626947, 1894.342934, 2534.163236]
    },
    "Paint (5%)": {
        "Ship Speed (knots)": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "Fuel Consumption (FC) (kg/h)": [110.0681824, 149.4348682, 201.2853668, 263.4790258, 304.5852633, 354.9289312, 415.9640576, 490.9831319, 585.731981, 787.8448646],
        "CO2 Emission (kg)": [342.3120472, 464.7424402, 625.9974908, 819.4197703, 947.2601688, 1103.828976, 1293.648219, 1526.95754, 1821.626461, 2450.197529]
    },
    "Advanced Propeller": {
        "Ship Speed (knots)": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "Fuel Consumption (FC) (kg/h)": [103.7743379, 140.8707293, 189.5548816, 247.890597, 286.2080063, 332.9373352, 389.4472384, 458.6313585, 545.6961681, 729.3067193],
        "CO2 Emission (kg)": [322.738191, 438.107968, 589.5156817, 770.9397568, 890.1068997, 1035.435112, 1211.180911, 1426.343525, 1697.115083, 2268.143897]
    },
    "Fin (2%-4%)": {
        "Ship Speed (knots)": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "Fuel Consumption (FC) (kg/h)": [115.4654418, 156.6860586, 210.7942821, 275.6058519, 318.2599935, 370.516624, 433.5948752, 510.8733655, 608.260007, 814.3657819],
        "CO2 Emission (kg)": [359.097524, 487.2936421, 655.5702172, 857.1341993, 989.7885799, 1152.306701, 1348.480062, 1588.816167, 1891.688622, 2532.677582]
    },
    "Bulbous Bow": {
        "Ship Speed (knots)": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "Fuel Consumption (FC) (kg/h)": [119.6541899, 162.3385399, 217.9322612, 284.2566267, 323.9702129, 371.0228762, 425.0311648, 501.2150135, 600.5953901, 863.1846317],
        "CO2 Emission (kg)": [372.1245307, 504.8728592, 677.7693324, 884.038109, 1007.547362, 1153.881145, 1321.846922, 1558.778692, 1867.851663, 2684.504205]
    }
}

# Figure 1: Comparative Fuel Consumption Trends for Original and Modified Scenarios
plt.figure(figsize=(10, 6))
for scenario, data in scenarios.items():
    plt.plot(data["Ship Speed (knots)"], data["Fuel Consumption (FC) (kg/h)"], label=f"{scenario} Scenario")
plt.xlabel("Ship Speed (knots)")
plt.ylabel("Fuel Consumption (kg/h)")
plt.title("Figure 1: Comparative Fuel Consumption Trends for Original and Modified Scenarios")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig(os.path.join(output_path, "Figure_1_Comparative_Fuel_Consumption.png"))
plt.show()

# Figure 2: CO₂ Emission Trends Across Scenarios
plt.figure(figsize=(10, 6))
for scenario, data in scenarios.items():
    plt.plot(data["Ship Speed (knots)"], data["CO2 Emission (kg)"], label=f"{scenario} Scenario")
plt.xlabel("Ship Speed (knots)")
plt.ylabel("CO₂ Emission (kg)")
plt.title("Figure 2: CO₂ Emission Trends Across Scenarios")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig(os.path.join(output_path, "Figure_2_CO2_Emission_Trends.png"))
plt.show()

# Figure 3: Scenario-Specific Performance Analysis for Technical Modifications
plt.figure(figsize=(10, 6))
for scenario, data in scenarios.items():
    plt.plot(data["Ship Speed (knots)"], data["Fuel Consumption (FC) (kg/h)"], label=f"{scenario} Fuel Consumption", linestyle='--')
    plt.plot(data["Ship Speed (knots)"], data["CO2 Emission (kg)"], label=f"{scenario} CO2 Emission")
plt.xlabel("Ship Speed (knots)")
plt.ylabel("Performance Metrics")
plt.title("Figure 3: Scenario-Specific Performance Analysis for Technical Modifications")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.savefig(os.path.join(output_path, "Figure_3_Scenario_Specific_Performance.png"))
plt.show()
