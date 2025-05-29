import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.svm import SVR
from sklearn.gaussian_process import GaussianProcessRegressor
from sklearn.gaussian_process.kernels import RBF, ConstantKernel as C
from sklearn.ensemble import RandomForestRegressor
from xgboost import XGBRegressor
from sklearn.neural_network import MLPRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_squared_error
import random

# Define input and output file path
input_output_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/Combined_Scenario_Data.xlsx"

# Read Combined Scenario data
try:
    data = pd.read_excel(input_output_path, sheet_name='Combined_Scen_Coeff')
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

# Required feature columns (matching first code's output)
required_columns = [
    'Pe (kW)', 'n (rpm)', 'T (kN)', 'ηH', 'ηO', 'Pd (kW)', 'PB (kW)', 'Ship Speed (knots)'
]

# Check for missing columns
missing_columns = [col for col in required_columns if col not in data.columns]
if missing_columns:
    print(f"Missing columns: {missing_columns}")
    print(f"Available columns: {data.columns.tolist()}")
    exit()

# Extract features and targets
features = data[required_columns].copy()
target_fc = data['Fuel Consumption (FC) (kg/h)'].copy()
target_co2 = data['CO2 Emission (kg)'].copy()

# Physics-Based Prediction
physics_based_fc = data['Pd (kW)'] * data['ηH'] * 0.8
physics_based_co2 = data['Pd (kW)'] * data['ηO'] * 0.5
features['Physics-Based FC'] = physics_based_fc
features['Physics-Based CO2'] = physics_based_co2

# Feature Engineering
features['Ship Speed^2'] = features['Ship Speed (knots)'] ** 2
features['Ship Speed^3'] = features['Ship Speed (knots)'] ** 3
features['Pd * Hull Efficiency'] = data['Pd (kW)'] * data['ηH']
features['Propeller Efficiency * n'] = data['ηO'] * data['n (rpm)']

# Scale features
scaler = StandardScaler()
features_scaled = scaler.fit_transform(features)

# Initialize ML models
models = {
    'SVR': SVR(kernel='rbf'),
    'GPR': GaussianProcessRegressor(kernel=C(1.0, (1e-3, 1e3)) * RBF(1.0, (1e-3, 1e3)), n_restarts_optimizer=10,
                                    random_state=42),
    'RF': RandomForestRegressor(n_estimators=100, random_state=42),
    'XGB': XGBRegressor(random_state=42),
    'NN': MLPRegressor(hidden_layer_sizes=(100, 50), max_iter=1000, random_state=42)
}

# Results storage
mse_results = []
predictions = {model: {'FC Before MC': [], 'CO2 Before MC': [], 'FC After MC': [], 'CO2 After MC': []} for model in
               models}

# Train and evaluate models
for model_name, model in models.items():
    print(f"Training {model_name} for Combined Scenario")

    # Split data
    X_train_fc, X_test_fc, y_train_fc, y_test_fc = train_test_split(features_scaled, target_fc, test_size=0.2,
                                                                    random_state=42)
    X_train_co2, X_test_co2, y_train_co2, y_test_co2 = train_test_split(features_scaled, target_co2, test_size=0.2,
                                                                        random_state=42)

    # Train for FC (Before MC)
    model.fit(X_train_fc, y_train_fc)
    predicted_fc = model.predict(features_scaled)
    mse_fc_before_mc = mean_squared_error(target_fc, predicted_fc)
    predictions[model_name]['FC Before MC'] = predicted_fc

    # Train for CO2 (Before MC)
    model.fit(X_train_co2, y_train_co2)
    predicted_co2 = model.predict(features_scaled)
    mse_co2_before_mc = mean_squared_error(target_co2, predicted_co2)
    predictions[model_name]['CO2 Before MC'] = predicted_co2

    # Monte Carlo Augmentation
    augmented_features = features.copy()
    augmented_target_fc = target_fc.copy()
    augmented_target_co2 = target_co2.copy()
    for _ in range(5):
        noise = features.apply(lambda x: x * (1 + random.uniform(-0.05, 0.05)))
        augmented_features = pd.concat([augmented_features, noise], ignore_index=True)
        augmented_target_fc = pd.concat([augmented_target_fc, target_fc], ignore_index=True)
        augmented_target_co2 = pd.concat([augmented_target_co2, target_co2], ignore_index=True)

    # Scale augmented features
    augmented_features_scaled = scaler.fit_transform(augmented_features)

    # Split augmented data
    X_train_fc_mc, X_test_fc_mc, y_train_fc_mc, y_test_fc_mc = train_test_split(augmented_features_scaled,
                                                                                augmented_target_fc, test_size=0.2,
                                                                                random_state=42)
    X_train_co2_mc, X_test_co2_mc, y_train_co2_mc, y_test_co2_mc = train_test_split(augmented_features_scaled,
                                                                                    augmented_target_co2, test_size=0.2,
                                                                                    random_state=42)

    # Train for FC (After MC)
    model.fit(X_train_fc_mc, y_train_fc_mc)
    predicted_fc_mc = model.predict(features_scaled)
    mse_fc_after_mc = mean_squared_error(target_fc, predicted_fc_mc)
    predictions[model_name]['FC After MC'] = predicted_fc_mc

    # Train for CO2 (After MC)
    model.fit(X_train_co2_mc, y_train_co2_mc)
    predicted_co2_mc = model.predict(features_scaled)
    mse_co2_after_mc = mean_squared_error(target_co2, predicted_co2_mc)
    predictions[model_name]['CO2 After MC'] = predicted_co2_mc

    # Store MSE results
    mse_results.append({
        'Model': model_name,
        'MSE Fuel Consumption Before MC': mse_fc_before_mc,
        'MSE Fuel Consumption After MC': mse_fc_after_mc,
        'MSE CO2 Emission Before MC': mse_co2_before_mc,
        'MSE CO2 Emission After MC': mse_co2_after_mc
    })

# Add predictions to data
for model_name in models:
    data[f'Final Predicted Fuel Consumption {model_name} Before MC (kg/h)'] = predictions[model_name]['FC Before MC']
    data[f'Final Predicted CO2 Emission {model_name} Before MC (kg)'] = predictions[model_name]['CO2 Before MC']
    data[f'Final Predicted Fuel Consumption {model_name} After MC (kg/h)'] = predictions[model_name]['FC After MC']
    data[f'Final Predicted CO2 Emission {model_name} After MC (kg)'] = predictions[model_name]['CO2 After MC']

# Save to Excel
try:
    with pd.ExcelWriter(input_output_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        data.to_excel(writer, sheet_name='Combined_Scen_Coeff', index=False)
        pd.DataFrame(mse_results).to_excel(writer, sheet_name='ML_MSE_Combined', index=False)
    print(f"Results saved to {input_output_path}")
except Exception as e:
    print(f"Error saving Excel file: {e}")