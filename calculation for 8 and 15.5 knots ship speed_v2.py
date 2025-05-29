import numpy as np
from sklearn.linear_model import LinearRegression


# Function to calculate physics-based FC (for reference, power-law model)
def predict_physics_fc(scenario_name, speeds, fc_values):
    X = np.array(speeds)
    y = np.array(fc_values)
    X_log = np.log(X).reshape(-1, 1)
    y_log = np.log(y)
    model = LinearRegression()
    model.fit(X_log, y_log)
    b = model.coef_[0]
    a = np.exp(model.intercept_)
    X_pred = np.array([8.0, 15.5])
    y_pred = a * np.power(X_pred, b)
    print(f"\n{scenario_name} Physics-Based:")
    print(f"FC at 8 knots: {y_pred[0]:.2f} kg/h")
    print(f"FC at 15.5 knots: {y_pred[1]:.2f} kg/h")
    return y_pred[0], y_pred[1]


# Function to extrapolate ML FC using physics-based ratio
def predict_ml_fc(scenario_name, speeds, ml_data, model_name, physics_fc_8, physics_fc_9, physics_fc_15,
                  physics_fc_15_5):
    fc_before = ml_data[f"{model_name}_Before_MC"]
    fc_after = ml_data[f"{model_name}_After_MC"]

    # Find 9-knot and 15-knot ML values
    idx_9 = speeds.index(9)
    idx_15 = speeds.index(15)
    fc_before_9 = fc_before[idx_9]
    fc_after_9 = fc_after[idx_9]
    fc_before_15 = fc_before[idx_15]
    fc_after_15 = fc_after[idx_15]

    # Extrapolate using physics-based FC ratio
    fc_before_8 = fc_before_9 * (physics_fc_8 / physics_fc_9)
    fc_after_8 = fc_after_9 * (physics_fc_8 / physics_fc_9)
    fc_before_15_5 = fc_before_15 * (physics_fc_15_5 / physics_fc_15)
    fc_after_15_5 = fc_after_15 * (physics_fc_15_5 / physics_fc_15)

    print(f"\n{scenario_name} - {model_name}:")
    print(f"Before MC at 8 knots: {fc_before_8:.2f} kg/h, After MC: {fc_after_8:.2f} kg/h")
    print(f"Before MC at 15.5 knots: {fc_before_15_5:.2f} kg/h, After MC: {fc_after_15_5:.2f} kg/h")

    return {
        "Before_8": fc_before_8, "After_8": fc_after_8,
        "Before_15_5": fc_before_15_5, "After_15_5": fc_after_15_5
    }


# Data for each scenario
scenarios = {
    "Advance Propeller": {
        "speeds": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "Physics_Based": [103.7743379, 140.8707293, 189.5548816, 247.890597, 286.2080063,
                          332.9373352, 389.4472384, 458.6313585, 545.6961681, 729.3067193],
        "SVR_Before_MC": [180.6946793, 168.2066289, 191.1641988, 247.99085, 286.1075057,
                          333.8740144, 389.347495, 429.4692586, 440.675173, 441.2868586],
        "SVR_After_MC": [119.1119081, 144.3497792, 184.6487631, 244.1215673, 283.931854,
                         333.1273451, 393.1309773, 464.5089952, 544.9027781, 652.5308001],
        "RF_Before_MC": [137.5077847, 138.3655901, 205.5157148, 259.5024821, 279.0127234,
                         323.3617529, 385.8533167, 428.8637319, 482.8819581, 633.0654382],
        "RF_After_MC": [104.8872297, 140.4997653, 189.0680401, 248.6569452, 285.6098965,
                        330.2313812, 389.7007227, 461.2433028, 547.5322736, 727.4706137]
    },
    "Combined": {
        "speeds": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "Physics_Based": [108.59, 147.58, 199.03, 260.12, 301.54, 351.13, 411.98, 486.46, 579.32, 771.04],
        "XGB_Before_MC": [108.5910568, 108.5910568, 199.0301666, 260.1192932, 301.5406799,
                          351.1300964, 411.9798584, 486.4600525, 486.4600525, 771.0388794],
        "XGB_After_MC": [108.5904694, 147.5800934, 199.0300293, 260.1200256, 301.5400391,
                         351.1299744, 411.9800415, 486.4599304, 579.3199463, 771.0396729],
        "GPR_Before_MC": [108.59, 152.8555864, 199.03, 260.12, 301.54, 351.13, 411.98,
                          486.46, 590.5928489, 771.04],
        "GPR_After_MC": [107.8253492, 146.9748097, 198.7271214, 259.8045197, 301.6543918,
                         352.5227859, 417.5965021, 499.5042938, 603.6922852, 756.8650185]
    },
    "Paint": {
        "speeds": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "Physics_Based": [110.0681824, 149.4348682, 201.2853668, 263.4790258, 304.5852633,
                          354.9289312, 415.9640576, 490.9831319, 585.731981, 787.8448646],
        "NN_Before_MC": [109.7261409, 136.466916, 190.9735841, 280.0992688, 305.8465382,
                         347.0132561, 411.2245197, 493.1734249, 603.0627193, 788.3603698],
        "NN_After_MC": [104.3293647, 159.4908743, 197.0763132, 237.5931262, 267.2496331,
                        328.4503802, 417.1987376, 493.8534728, 587.7390545, 803.6422616]
    }
}

# Your provided physics-based FC values
physics_fc_values = {
    "Advance Propeller": {"FC_8_knots": 61.41, "FC_15_5_knots": 677.97, "FC_9_knots": 103.7743379,
                          "FC_15_knots": 729.3067193},
    "Combined": {"FC_8_knots": 63.94, "FC_15_5_knots": 719.18, "FC_9_knots": 108.59, "FC_15_knots": 771.04},
    "Paint": {"FC_8_knots": 64.67, "FC_15_5_knots": 728.40, "FC_9_knots": 110.0681824, "FC_15_knots": 787.8448646}
}

# Calculate FC for physics-based (for reference) and ML models
results = {}
for scenario_name, data in scenarios.items():
    results[scenario_name] = {}
    # Physics-based FC (for reference, using code)
    if "Physics_Based" in data:
        fc_8, fc_15_5 = predict_physics_fc(scenario_name, data["speeds"], data["Physics_Based"])
        results[scenario_name]["Physics_Based"] = {"FC_8_knots": fc_8, "FC_15_5_knots": fc_15_5}
    # ML models (using your provided physics-based FC for scaling)
    for model in ["SVR", "RF", "XGB", "GPR", "NN"]:
        if f"{model}_Before_MC" in data:
            ml_results = predict_ml_fc(scenario_name, data["speeds"], data, model,
                                       physics_fc_values[scenario_name]["FC_8_knots"],
                                       physics_fc_values[scenario_name]["FC_9_knots"],
                                       physics_fc_values[scenario_name]["FC_15_knots"],
                                       physics_fc_values[scenario_name]["FC_15_5_knots"])
            results[scenario_name][model] = ml_results