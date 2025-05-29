from setuptools import setup

APP = ["prediction_app.py"]
# any packages you use
OPTIONS = {
    "argv_emulation": True,
    "packages": [
        "tkinter",
        "pandas",
        "numpy",
        "scipy",
        "sklearn",
        "xgboost",
        "matplotlib",
        "docx",
        "openpyxl"
    ],
    # force copy of the macOS Tk/Tcl frameworks
    "frameworks": [
        "/Library/Frameworks/Tk.framework",
        "/Library/Frameworks/Tcl.framework"
    ],
}

setup(
    app=APP,
    options={"py2app": OPTIONS},
    setup_requires=["py2app"],
)

