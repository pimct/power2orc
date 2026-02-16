import os
import yaml
import win32com.client

# ============================================================
# OPTION: choose what you want to run
# ============================================================
MODE = "power_to_orc"     # "power_only", "orc_only", "power_to_orc"
VISIBLE = 0              # 1 to show Aspen GUI

MODEL_EXT = ".apw"       # ".bkp" recommended; or ".apw"

# ============================================================
# USER INPUTS
# ============================================================
# --- Power plant (gas turbine) inputs ---
POWER_CASE = {
    "fuelfeed": 50.0,   # fuel mass flow (kg/s)
    "airfeed":  1000.0   # air mass flow  (kg/s)
}

# --- ORC inputs (if orc_only, you must provide all ORC required inputs) ---
ORC_CASE = {
    "fgastemp":   250,     # °C
    "fgaspres":   1.013,   # bar
    "fgasmsflow": 200.0,   # kg/s

    # flue gas composition (mole fraction) - example
    "fgasco2":   0.12,
    "fgasn2":    0.75,
    "fgasco":    0.00,
    "fgaswater": 0.08,
    "fgaso2":    0.05
}

# ============================================================
# PATHS (repo layout)
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

POWER_DIR = os.path.normpath(os.path.join(BASE_DIR, "aspen_models", "power"))
ORC_DIR   = os.path.normpath(os.path.join(BASE_DIR, "aspen_models", "ORC"))

POWER_YAML  = os.path.join(POWER_DIR, "power_paths.yaml")
ORC_YAML    = os.path.join(ORC_DIR, "ORC_paths.yaml")

POWER_MODEL = os.path.join(POWER_DIR, f"power{MODEL_EXT}")
ORC_MODEL   = os.path.join(ORC_DIR, f"ORC{MODEL_EXT}")


# ============================================================
# CORE HELPERS
# ============================================================
def load_yaml(path: str) -> dict:
    if not os.path.isfile(path):
        raise FileNotFoundError(f"YAML not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    if not isinstance(data, dict):
        raise ValueError(f"Invalid YAML: expected dict at {path}")
    if "input_paths" not in data or not isinstance(data["input_paths"], dict):
        raise KeyError(f"YAML must contain input_paths dict: {path}")
    if "output_paths" in data and not isinstance(data["output_paths"], dict):
        raise KeyError(f"output_paths must be dict: {path}")
    return data


def print_mapping(title: str, mapping: dict):
    print(f"\n--- {title} ---")
    if not mapping:
        print("(empty)")
        return
    for k, v in mapping.items():
        print(f"{k:12s} -> {v}")


def open_aspen(model_path: str, visible: int = 0):
    if not os.path.isfile(model_path):
        raise FileNotFoundError(f"Model not found: {model_path}")
    aspen = win32com.client.DispatchEx("Apwn.Document")
    aspen.InitFromArchive2(os.path.normpath(model_path))
    aspen.Visible = int(bool(visible))
    return aspen


def close_aspen(aspen):
    try:
        aspen.Close(False)  # don't save
    except Exception:
        pass


def set_value(aspen, path: str, value):
    node = aspen.Tree.FindNode(path)
    if node is None:
        raise RuntimeError(f"Aspen path not found: {path}")
    node.Value = value


def get_value(aspen, path: str):
    node = aspen.Tree.FindNode(path)
    if node is None:
        return None
    return node.Value


def run_case_fresh(model_path: str, io_map: dict, inputs: dict, visible: int = 0) -> dict:
    """
    Reinitialize each run by opening a fresh Aspen instance.
    io_map must have: input_paths (dict), output_paths (dict optional)
    """
    input_paths = io_map["input_paths"]
    output_paths = io_map.get("output_paths", {})

    aspen = open_aspen(model_path, visible=visible)
    try:
        # set inputs
        for name, val in inputs.items():
            if name not in input_paths:
                raise KeyError(f"Input '{name}' not found in YAML input_paths")
            set_value(aspen, input_paths[name], val)

        # run
        aspen.Engine.Run2()

        # collect outputs
        out = {"inputs": dict(inputs)}
        for name, path in output_paths.items():
            out[name] = get_value(aspen, path)
        return out
    finally:
        close_aspen(aspen)


def extract_for_orc(power_out: dict) -> dict:
    """
    Map power model outputs -> ORC model required inputs.
    Assumes variable names align (fgastemp, fgasmsflow, fgasco2, etc.)
    """
    required = [
        "fgastemp", "fgasmsflow", "fgasco2", "fgasn2", "fgasco", "fgaswater"
    ]
    orc_in = {}
    for k in required:
        orc_in[k] = power_out.get(k)

    # ORC also needs fgaspres and fgaso2 (power yaml doesn't provide these)
    # Choose defaults / keep user-provided values
    return orc_in


# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    print("BASE_DIR    =", BASE_DIR)
    print("POWER_YAML  =", POWER_YAML,  "| exists:", os.path.exists(POWER_YAML))
    print("ORC_YAML    =", ORC_YAML,    "| exists:", os.path.exists(ORC_YAML))
    print("POWER_MODEL =", POWER_MODEL, "| exists:", os.path.exists(POWER_MODEL))
    print("ORC_MODEL   =", ORC_MODEL,   "| exists:", os.path.exists(ORC_MODEL))

    power_map = load_yaml(POWER_YAML)
    orc_map = load_yaml(ORC_YAML)

    # Print for cross-check
    print_mapping("POWER input_paths", power_map["input_paths"])
    print_mapping("POWER output_paths", power_map.get("output_paths", {}))
    print_mapping("ORC input_paths", orc_map["input_paths"])
    print_mapping("ORC output_paths", orc_map.get("output_paths", {}))

    if MODE == "power_only":
        print("\nMODE = power_only")
        power_res = run_case_fresh(POWER_MODEL, power_map, POWER_CASE, visible=VISIBLE)
        print("\nRESULT (POWER)")
        for k, v in power_res.items():
            if k != "inputs":
                print(f"{k}: {v}")

    elif MODE == "orc_only":
        print("\nMODE = orc_only")
        orc_res = run_case_fresh(ORC_MODEL, orc_map, ORC_CASE, visible=VISIBLE)
        print("\nRESULT (ORC)")
        for k, v in orc_res.items():
            if k != "inputs":
                print(f"{k}: {v}")

    elif MODE == "power_to_orc":
        print("\nMODE = power_to_orc")

        # 1) Run power model
        power_res = run_case_fresh(POWER_MODEL, power_map, POWER_CASE, visible=VISIBLE)
        power_out = {k: v for k, v in power_res.items() if k != "inputs"}

        # 2) Build ORC inputs from power outputs + fill missing items
        orc_from_power = extract_for_orc(power_out)

        # fill missing ORC-required inputs that are not provided by power model
        # - pressure and O2 mole fraction are not in your power_paths.yaml
        orc_inputs = dict(ORC_CASE)  # start from user defaults
        orc_inputs.update({k: v for k, v in orc_from_power.items() if v is not None})

        # 3) Run ORC model
        orc_res = run_case_fresh(ORC_MODEL, orc_map, orc_inputs, visible=VISIBLE)

        print("\nRESULT (POWER -> ORC)")
        print("Power outputs (key):")
        for k in ["work", "fgastemp", "fgasmsflow", "fgasco2", "fgasn2", "fgasco", "fgaswater", "steammsflow"]:
            if k in power_out:
                print(f"  {k}: {power_out[k]}")

        print("\nORC outputs:")
        for k in ["work", "qevap"]:
            print(f"  {k}: {orc_res.get(k)}")

    else:
        raise ValueError("MODE must be: power_only | orc_only | power_to_orc")
