import matplotlib.pyplot as plt
import numpy as np

# Ship Speed Data
ship_speed = np.array([9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15])

# Data for Different Scenarios (Original, Paint, Advanced Propeller, Fin, Bulbous Bow)
scenarios = {
    "Original": {
        "Treq": [89.7286371, 109.5295012, 133.5905855, 160.0163891, 176.7602696, 196.140474, 218.767217, 245.5898251, 278.2940591, 346.3754643],
        "Pe": [314.0482199, 426.5087541, 572.9789987, 750.6893668, 866.0652227, 1002.085761, 1163.712609, 1360.083917, 1600.39533, 2076.635905],
        "Rt": [67.83484964, 82.91383244, 101.2616638, 121.6124557, 134.6913255, 149.8513221, 167.5756882, 188.8585755, 214.5647195, 269.1337357]
    },
    "Paint (5%)": {
        "Treq": [85.98485858, 104.9770907, 128.1574232, 153.6398644, 169.8957494, 188.7725699, 210.8807421, 237.1807377, 269.3372922, 336.9054225],
        "Pe": [300.945079, 408.781631, 549.6758001, 720.7749977, 832.4314076, 964.442986, 1121.761213, 1313.514136, 1548.887339, 2019.85986],
        "Rt": [65.00455308, 79.46765766, 97.14332675, 116.7662969, 129.4605611, 144.2222434, 161.5346485, 182.3919873, 207.6590523, 261.7755133]
    },
    "Advanced Propeller": {
        "Treq": [89.7286371, 109.5295012, 133.5905855, 160.0163891, 176.7602696, 196.140474, 218.767217, 245.5898251, 278.2940591, 346.3754643],
        "Pe": [314.0482199, 426.5087541, 572.9789987, 750.6893668, 866.0652227, 1002.085761, 1163.712609, 1360.083917, 1600.39533, 2076.635905],
        "Rt": [67.83484964, 82.91383244, 101.2616638, 121.6124557, 134.6913255, 149.8513221, 167.5756882, 188.8585755, 214.5647195, 269.1337357]
    },
    "Fin (2%-4%)": {
        "Treq": [90.31160087, 110.237233, 134.449073, 161.0334424, 177.8713822, 197.3597647, 220.1120268, 247.074198, 279.9570986, 348.3751466],
        "Pe": [314.0482199, 426.5087541, 572.9789987, 750.6893668, 866.0652227, 1002.085761, 1163.712609, 1360.083917, 1600.39533, 2076.635905],
        "Rt": [67.83484964, 82.91383244, 101.2616638, 121.6124557, 134.6913255, 149.8513221, 167.5756882, 188.8585755, 214.5647195, 269.1337357]
    },
    "Bulbous Bow": {
        "Treq": [92.25005292, 112.6111415, 137.1429527, 163.9600144, 179.1046976, 195.9860825, 214.7546376, 241.3303844, 275.0420581, 363.1135134],
        "Pe": [320.7889643, 435.6934255, 584.4594532, 764.3321633, 872.0703012, 995.1109488, 1135.388572, 1328.465608, 1572.298139, 2164.490111],
        "Rt": [69.29085975, 84.699344, 103.2905863, 123.8226029, 135.6252412, 148.8083127, 163.4970007, 184.4681193, 210.7977341, 280.5197137]
    }
}

# Figure 1: Total Resistance (R_T) vs. Ship Speed for Different Scenarios
plt.figure(figsize=(10, 6))
for scenario, data in scenarios.items():
    plt.plot(ship_speed, data['Rt'], label=scenario)
plt.xlabel('Ship Speed (knots)')
plt.ylabel('Total Resistance (R_T) (kN)')
plt.title('Figure 1: Total Resistance (R_T) vs. Ship Speed for Different Scenarios')
plt.legend()
plt.grid(True)
plt.savefig("/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/Figure_1_Total_Resistance_vs_Ship_Speed.png")
plt.show()

# Figure 2: Effective Power (P_e) vs. Ship Speed for Different Scenarios
plt.figure(figsize=(10, 6))
for scenario, data in scenarios.items():
    plt.plot(ship_speed, data['Pe'], label=scenario)
plt.xlabel('Ship Speed (knots)')
plt.ylabel('Effective Power (P_e) (kW)')
plt.title('Figure 2: Effective Power (P_e) vs. Ship Speed for Different Scenarios')
plt.legend()
plt.grid(True)
plt.savefig("/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/Figure_2_Effective_Power_vs_Ship_Speed.png")
plt.show()

# Figure 3: Thrust Required (T_req) vs. Ship Speed for Different Scenarios
plt.figure(figsize=(10, 6))
for scenario, data in scenarios.items():
    plt.plot(ship_speed, data['Treq'], label=scenario)
plt.xlabel('Ship Speed (knots)')
plt.ylabel('Thrust Required (T_req) (kN)')
plt.title('Figure 3: Thrust Required (T_req) vs. Ship Speed for Different Scenarios')
plt.legend()
plt.grid(True)
plt.savefig("/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_2_graphs/Figure_3_Thrust_Required_vs_Ship_Speed.png")
plt.show()
