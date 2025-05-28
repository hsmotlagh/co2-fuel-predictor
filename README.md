# 🚢 XGB CO₂ & Fuel-Consumption Predictor

A secure, local web-based application for predicting ship fuel consumption and CO₂ emissions using XGBoost machine learning models. Built in Streamlit with integrated scenario analysis, Excel output, and visualizations.

---

## 🔐 Password Protection
This app is protected with a login system using `.streamlit/secrets.toml`. If you are a collaborator or evaluator and need access credentials (username and password), please contact the creator directly.

---

## 📁 Folder Structure
```
co2-fuel-predictor/
├── app.py
├── requirements.txt
├── AppIcon.png
├── ML_Apply copy.xlsx
├── XGB_ML_model/
│   ├── xgb_model_fc_combined.pkl
│   ├── xgb_model_co2_combined.pkl
│   └── scaler_combined.pkl
├── .streamlit/
│   └── secrets.toml
└── README.md
```

---

## 💻 Run Locally (Recommended)

### 1. Clone the Repository
```bash
git clone https://github.com/hsmotlagh/co2-fuel-predictor.git
cd co2-fuel-predictor
```

### 2. Create a Virtual Environment
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Install Dependencies
```bash
pip install -r requirements.txt
```

### 4. Create Secrets File
Create the file `.streamlit/secrets.toml` and enter the credentials provided to you:
```toml
[auth]
username = "your-username"
password = "your-password"
# contact me for username and password
```

### 5. Run the App
```bash
streamlit run app.py
```

### 6. Open in Browser
Once launched, a browser tab will open. Enter your username and password to access the interface.

---

## 🧪 Features
- XGBoost model for FC and CO₂ prediction
- Interactive scenario comparison (Original, Paint, Fin, etc.)
- Excel and Word report export
- Secure login system via `secrets.toml`
- Visual charts using matplotlib

---

## 👤 Creator
**Hamid**, PhD Candidate – Marine Engineering & AI for Sustainability  
© 2025 All Rights Reserved

---

## 📜 License
This repository is for academic and non-commercial use. For permission to reuse or adapt, contact the author.

