import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split, GridSearchCV
from xgboost import XGBRegressor
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import mean_squared_error
import random
import joblib
import os

# File paths
input_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/ML_Apply copy.xlsx"
output_vis_path = "/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/XGB_ML_model/"

# Ensure output directory exists
os.makedirs(output_vis_path, exist_ok=True)

# Scenarios
sheet_names = [
    "Original Scenar_Coeff",
    "Paint (5%) Scen_Coeff",
    "Advance Propell_Coeff",
    "Fin (2%-4%) Sce_Coeff",
    "Bulbous Bow Sce_Coeff"
]

# Feature columns
required_columns = [
    'Pe (kW)', 'n (rpm)', 'T (kN)', 'Hull Efficiency (ηH)',
    'Propeller Open Water Efficiency (ηO)', 'Delivered Power (Pd) (kW)',
    'Brake Power (PB) (kW)', 'Ship Speed (knots)'
]

# Load and combine data
try:
    excel_data = pd.read_excel(input_path, sheet_name=sheet_names)
    # Combine all scenario data
    combined_data = pd.concat([excel_data[sheet] for sheet in sheet_names], ignore_index=True)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

# Check required columns
missing_columns = [col for col in required_columns if col not in combined_data.columns]
if missing_columns:
    print(f"Missing columns: {missing_columns}")
    exit()

# Extract features and targets
features = combined_data[required_columns].copy()
if 'Fuel Consumption (FC) (kg/h)' not in combined_data.columns or 'CO2 Emission (kg)' not in combined_data.columns:
    print("Target columns missing.")
    exit()
target_fc = combined_data['Fuel Consumption (FC) (kg/h)'].copy()
target_co2 = combined_data['CO2 Emission (kg)'].copy()

# Debugging: Check initial shapes
print(f"Initial features shape: {features.shape}")
print(f"Initial target_fc shape: {target_fc.shape}")
print(f"Initial target_co2 shape: {target_co2.shape}")

# Feature engineering
features['Physics-Based FC'] = features['Delivered Power (Pd) (kW)'] * features['Hull Efficiency (ηH)'] * 0.8
features['Physics-Based CO2'] = features['Delivered Power (Pd) (kW)'] * features['Propeller Open Water Efficiency (ηO)'] * 0.5
features['Ship Speed^2'] = features['Ship Speed (knots)'] ** 2
features['Ship Speed^3'] = features['Ship Speed (knots)'] ** 3
features['Pd * Hull Efficiency'] = features['Delivered Power (Pd) (kW)'] * features['Hull Efficiency (ηH)']
features['Propeller Efficiency * n'] = features['Propeller Open Water Efficiency (ηO)'] * features['n (rpm)']
features['Brake Power (PB) (kW)'] = features['Brake Power (PB) (kW)'].where(
    features['Brake Power (PB) (kW)'] > 0, features['Pe (kW)'] * 0.95
)

# Scale features
scaler = StandardScaler()
features_scaled = scaler.fit_transform(features)

# Split data
X_train_fc, X_test_fc, y_train_fc, y_test_fc = train_test_split(
    features_scaled, target_fc, test_size=0.2, random_state=42
)
X_train_co2, X_test_co2, y_train_co2, y_test_co2 = train_test_split(
    features_scaled, target_co2, test_size=0.2, random_state=42
)

# Hyperparameter tuning for FC model
param_grid = {
    'n_estimators': [50, 100, 200],
    'max_depth': [3, 5, 7],
    'learning_rate': [0.01, 0.1, 0.3],
    'min_child_weight': [1, 3, 5],
    'gamma': [0, 0.1, 0.3],
    'subsample': [0.7, 0.9, 1.0]
}
xgb_model_fc = XGBRegressor(random_state=42)
grid_search_fc = GridSearchCV(xgb_model_fc, param_grid, cv=5, scoring='neg_mean_squared_error', n_jobs=-1)
grid_search_fc.fit(X_train_fc, y_train_fc)
xgb_model_fc = grid_search_fc.best_estimator_

# Hyperparameter tuning for CO2 model
xgb_model_co2 = XGBRegressor(random_state=42)
grid_search_co2 = GridSearchCV(xgb_model_co2, param_grid, cv=5, scoring='neg_mean_squared_error', n_jobs=-1)
grid_search_co2.fit(X_train_co2, y_train_co2)
xgb_model_co2 = grid_search_co2.best_estimator_

# Monte Carlo augmentation (5x)
augmented_features = features.copy()
augmented_target_fc = target_fc.copy()
augmented_target_co2 = target_co2.copy()

for _ in range(5):
    noise = features.apply(lambda x: x * (1 + random.uniform(-0.05, 0.05)))
    augmented_features = pd.concat([augmented_features, noise], ignore_index=True)
    augmented_target_fc = pd.concat([augmented_target_fc, target_fc], ignore_index=True)
    augmented_target_co2 = pd.concat([augmented_target_co2, target_co2], ignore_index=True)

# Ensure targets are 1D arrays
augmented_target_fc = augmented_target_fc.values.ravel() if hasattr(augmented_target_fc, 'values') else augmented_target_fc.ravel()
augmented_target_co2 = augmented_target_co2.values.ravel() if hasattr(augmented_target_co2, 'values') else augmented_target_co2.ravel()

# Debugging: Check augmented shapes
print(f"Augmented features shape: {augmented_features.shape}")
print(f"Augmented target_fc shape: {augmented_target_fc.shape}")
print(f"Augmented target_co2 shape: {augmented_target_co2.shape}")

# Scale augmented features
augmented_features_scaled = scaler.transform(augmented_features)

# Train models (After MC)
xgb_model_fc.fit(augmented_features_scaled, augmented_target_fc)
xgb_model_co2.fit(augmented_features_scaled, augmented_target_co2)

# Save models and scaler
try:
    joblib.dump(xgb_model_fc, f"{output_vis_path}xgb_model_fc_combined.pkl")
    joblib.dump(xgb_model_co2, f"{output_vis_path}xgb_model_co2_combined.pkl")
    joblib.dump(scaler, f"{output_vis_path}scaler_combined.pkl")
    print("Combined models and scaler saved")
except Exception as e:
    print(f"Error saving models: {e}")