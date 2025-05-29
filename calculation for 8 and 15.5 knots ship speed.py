import numpy as np
from sklearn.linear_model import LinearRegression


# Function to fit power-law model (FC = a * V^b) and predict FC at 8 and 15.5 knots
def predict_fc(scenario_name, speeds, fc_values):
    # Convert lists to numpy arrays
    X = np.array(speeds)
    y = np.array(fc_values)

    # Apply logarithmic transformation for linear regression
    X_log = np.log(X).reshape(-1, 1)
    y_log = np.log(y)

    # Fit linear regression on log-transformed data
    model = LinearRegression()
    model.fit(X_log, y_log)

    # Extract coefficients: b is slope, a is exp(intercept)
    b = model.coef_[0]
    a = np.exp(model.intercept_)

    # Predict FC at 8 and 15.5 knots
    X_pred = np.array([8.0, 15.5])
    y_pred = a * np.power(X_pred, b)

    print(f"\n{scenario_name} Scenario:")
    print(f"FC at 8 knots: {y_pred[0]:.2f} kg/h")
    print(f"FC at 15.5 knots: {y_pred[1]:.2f} kg/h")

    return y_pred[0], y_pred[1]


# Data for each scenario (from your provided FC values)
scenarios = {
    "Original": {
        "speeds": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "fc": [115.7696693, 157.0962029, 211.3112914, 276.2684182, 319.0039009,
               371.3694974, 434.4792596, 511.7771533, 609.1134837, 814.8434842]
    },
    "Paint": {
        "speeds": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "fc": [110.0681824, 149.4348682, 201.2853668, 263.4790258, 304.5852633,
               354.9289312, 415.9640576, 490.9831319, 585.731981, 787.8448646]
    },
    "Advance Propeller": {
        "speeds": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "fc": [103.7743379, 140.8707293, 189.5548816, 247.890597, 286.2080063,
               332.9373352, 389.4472384, 458.6313585, 545.6961681, 729.3067193]
    },
    "Fin": {
        "speeds": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "fc": [115.4654418, 156.6860586, 210.7942821, 275.6058519, 318.2599935,
               370.516624, 433.5948752, 510.8733655, 608.260007, 814.3657819]
    },
    "Bulbous Bow": {
        "speeds": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "fc": [119.6541899, 162.3385399, 217.9322612, 284.2566267, 323.9702129,
               371.0228762, 425.0311648, 501.2150135, 600.5953901, 863.1846317]
    },
    "Combined": {
        "speeds": [9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15],
        "fc": [108.59, 147.58, 199.03, 260.12, 301.54, 351.13, 411.98, 486.46, 579.32, 771.04]
    }
}

# Calculate FC for each scenario
results = {}
for scenario_name, data in scenarios.items():
    fc_8, fc_15_5 = predict_fc(scenario_name, data["speeds"], data["fc"])
    results[scenario_name] = {"FC_8_knots": fc_8, "FC_15_5_knots": fc_15_5}