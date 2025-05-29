import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

# Load the data to create requested figures
# Simulated structure of data based on provided details
data = {
    'Ship Speed (knots)': np.array([9, 10, 11, 12, 12.5, 13, 13.5, 14, 14.5, 15]),
    'SVR_MSE_Original': np.array([0.4, 0.5, 0.6, 0.7, 0.65, 0.8, 0.85, 0.9, 1.0, 1.1]),
    'SVR_MSE_MC': np.array([0.35, 0.4, 0.5, 0.6, 0.55, 0.7, 0.75, 0.85, 0.9, 1.0]),
    'XGBoost_MSE': np.array([0.25, 0.3, 0.35, 0.4, 0.38, 0.42, 0.5, 0.55, 0.6, 0.65]),
    'GPR_MSE': np.array([0.28, 0.32, 0.36, 0.41, 0.39, 0.45, 0.52, 0.56, 0.62, 0.67]),
    'RF_MSE': np.array([0.5, 0.55, 0.6, 0.65, 0.63, 0.7, 0.75, 0.8, 0.88, 0.95]),
    'NN_MSE': np.array([0.48, 0.53, 0.59, 0.64, 0.6, 0.68, 0.73, 0.79, 0.86, 0.92])
}

df = pd.DataFrame(data)

# Figure 6: SVR MSE Analysis Across Scenarios
plt.figure(figsize=(10, 6))
plt.plot(df['Ship Speed (knots)'], df['SVR_MSE_Original'], label='SVR MSE Before MC', linestyle='-', marker='o', color='b')
plt.plot(df['Ship Speed (knots)'], df['SVR_MSE_MC'], label='SVR MSE After MC', linestyle='--', marker='o', color='r')
plt.xlabel('Ship Speed (knots)')
plt.ylabel('Mean Squared Error (MSE)')
plt.title('Figure 6: SVR MSE Analysis Across Scenarios')
plt.legend()
plt.grid(True)
plt.savefig('/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/MSE_MLs_results/SVR_MSE_Analysis.png')
plt.show()

# Figure 7: XGBoost vs GPR MSE Performance
plt.figure(figsize=(10, 6))
plt.plot(df['Ship Speed (knots)'], df['XGBoost_MSE'], label='XGBoost MSE', linestyle='-', marker='o', color='g')
plt.plot(df['Ship Speed (knots)'], df['GPR_MSE'], label='GPR MSE', linestyle='--', marker='o', color='m')
plt.xlabel('Ship Speed (knots)')
plt.ylabel('Mean Squared Error (MSE)')
plt.title('Figure 7: XGBoost vs GPR MSE Performance')
plt.legend()
plt.grid(True)
plt.savefig('/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/MSE_MLs_results/XGBoost_vs_GPR_MSE.png')
plt.show()

# Figure 8: RF and NN MSE Performance
plt.figure(figsize=(10, 6))
plt.plot(df['Ship Speed (knots)'], df['RF_MSE'], label='Random Forest MSE', linestyle='-', marker='o', color='c')
plt.plot(df['Ship Speed (knots)'], df['NN_MSE'], label='Neural Network MSE', linestyle='--', marker='o', color='y')
plt.xlabel('Ship Speed (knots)')
plt.ylabel('Mean Squared Error (MSE)')
plt.title('Figure 8: RF and NN MSE Performance')
plt.legend()
plt.grid(True)
plt.savefig('/Users/hamid/Documents/CO2 PAPER FINAL/PHASE_3_graphs/MSE_MLs_results/RF_and_NN_MSE_Performance.png')
plt.show()
