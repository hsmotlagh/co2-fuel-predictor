from sklearn.metrics import r2_score
import numpy as np

# Data from your Excel
J = np.array([0.05, 0.1, 0.15, 0.2, 0.25, 0.3, 0.35, 0.4, 0.45, 0.5, 0.55, 0.6, 0.65, 0.7, 0.75, 0.8, 0.85])
original_kt = np.array([0.2993, 0.2882, 0.2762, 0.2632, 0.2493, 0.2346, 0.219, 0.2026, 0.1855, 0.1678, 0.1492, 0.1302, 0.1105, 0.0903, 0.0694, 0.0482, 0.0264])
original_kq = np.array([0.3696, 0.359, 0.3479, 0.336, 0.3234, 0.31, 0.2956, 0.2803, 0.2638, 0.2462, 0.2274, 0.2072, 0.1857, 0.1626, 0.138, 0.1118, 0.0838])

# Coefficients from Excel
kt_coeff = [-0.141738906, -0.216126161, 0.311516176]
kq_coeff = [-0.231996904, -0.142815531, 0.375066176]

# Predicted values
predicted_kt = kt_coeff[0] * J**2 + kt_coeff[1] * J + kt_coeff[2]
predicted_kq = kq_coeff[0] * J**2 + kq_coeff[1] * J + kq_coeff[2]

# Calculate R²
r2_kt = r2_score(original_kt, predicted_kt)
r2_kq = r2_score(original_kq, predicted_kq)
print(f"R² for Kt: {r2_kt:.3f}, R² for Kq: {r2_kq:.3f}")