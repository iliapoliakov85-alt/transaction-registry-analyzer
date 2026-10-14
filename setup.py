from setuptools import setup

APP = ["analyzer.py"]
DATA_FILES = [("", ["update_config.json"])]
OPTIONS = {
    "argv_emulation": False,
    "iconfile": "assets/AppIcon.icns",
    "packages": ["duckdb", "pandas", "openpyxl", "PySide6"],
    "plist": {
        "CFBundleDisplayName": "Анализатор реестра",
        "CFBundleName": "Анализатор реестра",
        "CFBundleIdentifier": "ru.asop.transaction-analyzer",
        "CFBundleShortVersionString": "8.4.2",
        "CFBundleVersion": "8.4.2",
        "LSMinimumSystemVersion": "12.0",
        "NSHighResolutionCapable": True,
    },
}

setup(app=APP, data_files=DATA_FILES, options={"py2app": OPTIONS}, setup_requires=["py2app"])
