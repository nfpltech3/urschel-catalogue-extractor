# Urschel PDF Catalogue Extractor

A desktop utility tool designed for Nagarkot Forwarders Pvt. Ltd. to automate the rapid identification and extraction of specific product models from large OEM PDF catalogues.

## Tech Stack
- Python 3.10+
- PyMuPDF (fitz) for PDF Processing
- Pandas for Excel interfacing
- Custom Tkinter GUI matching Nagarkot aesthetic

---

## Installation

### Clone
```bash
git clone https://github.com/username/urschel-catalogue-extractor.git
cd urschel-catalogue-extractor
```

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. **Create virtual environment**
```bash
python -m venv venv
```

2. **Activate (REQUIRED)**

Windows:
```cmd
venv\Scripts\activate
```

Mac/Linux:
```bash
source venv/bin/activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Run application**
```bash
python urschel_tool.py
```

---

### Build Executable (For Desktop Apps)

1. Make sure PyInstaller and modules are installed inside venv:
```bash
pip install -r requirements.txt
```

2. Build using the included Spec file (which includes the `logo.png` dependency):
```bash
pyinstaller Urschel_Catalogue_Tool.spec
```

3. Locate Executable:
The `Urschel_Catalogue_Tool.exe` will be generated in the `dist/` folder.

---

## Environment Variables

Copy:
```bash
cp .env.example .env
```
No specific API keys are required for this app as it runs entirely locally.

---

## Notes
- ALWAYS use virtual environment for Python development.
- Do not commit `venv` or `__pycache__`.
- Run and test before pushing any changes.
