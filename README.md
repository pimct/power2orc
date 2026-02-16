# Aspen Python Integrator (Power → ORC Simulation)

This project automates Aspen Plus simulations using Python via COM automation.
It supports:

- Running a **power plant model**
- Running an **Organic Rankine Cycle (ORC) model**
- Passing flue gas output from power plant → ORC automatically
- YAML-based Aspen path mapping for easy modification

Typical use:
Waste heat recovery studies, ORC optimization, process integration research.

## Project layout
aspen_python_integrator/
├─ main.py
├─ requirements.txt
├─ environment.yml
├─ README.md
├─ .gitignore
├─ LICENSE
├─ CITATION.cff
└─ aspen_models/
    └─ ORC/
        ├─ ORC_paths.yaml
        └─ ORC.apw
    └─ power/
        ├─ power_paths.yaml
        └─ power.apw

## ⚙️ Requirements
### Software
- Windows OS
- Aspen Plus installed locally
- Aspen COM automation enabled

### Python
Recommended:
- Python 3.10
- Conda environment preferred

---

## 🚀 Installation

### Option 1 — Conda (Recommended)
conda env create -f environment.yml
conda activate aspen-python

### Option 2 — pip
pip install -r requirements.txt

## ▶️ Running the Simulation
Edit options at the top of main.py:
MODE = "power_to_orc"

Available modes:

Mode	Description
power_only	Run power plant model only
orc_only	Run ORC model only
power_to_orc	Run power plant → send flue gas to ORC

## 📄 Aspen Path Configuration

All Aspen variables are defined in YAML:
Power model:
aspen_models/power/power_paths.yaml
ORC model:
aspen_models/ORC/ORC_paths.yaml

Example:
fgastemp: '\Data\Streams\FLUEGAS\Input\TEMP\MIXED'


## 🔄 Power → ORC Integration Logic

The script:
1. Runs power plant simulation
2. Extracts:
   - Flue gas temperature
   - Mass flow
   - Composition
3. Passes these to ORC simulation automatically

This enables:
    - Waste heat recovery studies
    - Combined cycle analysis
    - Sensitivity studies

# 👤 Author
Prathana Nimmanterdwong