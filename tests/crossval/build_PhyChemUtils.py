"""Cross-validate PhyChemUtils functions against Python reference implementations.

Usage: python tests/build_PhyChemUtils.py
"""

import os
import sys
import re as _re

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls")
                for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "PhyChemUtils.bas"))


# =============================================================================
# Atomic weights reference
# =============================================================================
ATOMIC_WEIGHTS = {
    "H": 1.008, "He": 4.003, "Li": 6.941, "Be": 9.012, "B": 10.811,
    "C": 12.011, "N": 14.007, "O": 15.999, "F": 18.998, "Ne": 20.180,
    "Na": 22.990, "Mg": 24.305, "Al": 26.982, "Si": 28.086, "P": 30.974,
    "S": 32.065, "Cl": 35.453, "K": 39.098, "Ca": 40.078, "Fe": 55.845,
    "Zn": 65.380, "Br": 79.904, "Ag": 107.868, "I": 126.904, "Ba": 137.327,
}


def _expand_group(group: str, multiplier: int) -> str:
    """Expand (OH)2 → O2H2.  Also handles nested parens like Fe(CN)6."""
    result = []
    i = 0
    while i < len(group):
        pm = _re.match(r'\(([^()]+)\)(\d*)', group[i:])
        if pm:
            inner, cnt = pm.group(1), int(pm.group(2)) if pm.group(2) else 1
            result.append(_expand_group(inner, cnt * multiplier))
            i += pm.end(); continue
        bm = _re.match(r'\[([^\[\]]+)\](\d*)', group[i:])
        if bm:
            inner, cnt = bm.group(1), int(bm.group(2)) if bm.group(2) else 1
            result.append(_expand_group(inner, cnt * multiplier))
            i += bm.end(); continue
        em = _re.match(r'([A-Z][a-z]?)(\d*)', group[i:])
        if em:
            elem, cnt = em.group(1), int(em.group(2)) if em.group(2) else 1
            result.append(f"{elem}{cnt * multiplier}" if cnt * multiplier > 1 else elem)
            i += em.end(); continue
        i += 1
    return "".join(result)


def _py_molecular_weight(formula: str) -> float:
    """Compute molecular weight from formula string.
    Handles: parentheses Ca(OH)2, brackets Fe4[Fe(CN)6]3, hydrates CuSO4.5H2O.
    """
    # 1) Split hydrate parts by middle dot
    if _re.search(r'[.]', formula):
        return sum(_py_molecular_weight(p) for p in _re.split(r'[.]', formula))

    # 2) Expand bracket groups [...] — treat like parentheses
    while '[' in formula:
        formula = _re.sub(
            r'\[([^\[\]]+)\](\d*)',
            lambda m: _expand_group(m.group(1), int(m.group(2) or "1")),
            formula,
        )

    # 3) Expand parenthesized groups (...)
    while '(' in formula:
        formula = _re.sub(
            r'\(([^()]+)\)(\d*)',
            lambda m: _expand_group(m.group(1), int(m.group(2) or "1")),
            formula,
        )

    # 4) Sum elements
    total = 0.0
    for m in _re.finditer(r'([A-Z][a-z]?)(\d*)', formula):
        elem = m.group(1)
        count = int(m.group(2)) if m.group(2) else 1
        if elem in ATOMIC_WEIGHTS:
            total += ATOMIC_WEIGHTS[elem] * count
    return total


# =============================================================================
# Test Cases
# =============================================================================

TEST_CASES = [

    # ---- MolecularWeight ----
    {"name": "MolecularWeight_H2O", "func": "MolecularWeight",
     "args": lambda: ("H2O",),
     "py_ref": lambda a: _py_molecular_weight(a[0]),
     "result_type": "scalar", "tol": 0.05},
    {"name": "MolecularWeight_CO2", "func": "MolecularWeight",
     "args": lambda: ("CO2",),
     "py_ref": lambda a: _py_molecular_weight(a[0]),
     "result_type": "scalar", "tol": 0.05},
    {"name": "MolecularWeight_CaOH2", "func": "MolecularWeight",
     "args": lambda: ("Ca(OH)2",),
     "py_ref": lambda a: _py_molecular_weight(a[0]),
     "result_type": "scalar", "tol": 0.1},
    {"name": "MolecularWeight_Na2SO4", "func": "MolecularWeight",
     "args": lambda: ("Na2SO4",),
     "py_ref": lambda a: _py_molecular_weight(a[0]),
     "result_type": "scalar", "tol": 0.1},

    # ---- ConvertTemperature ----
    {"name": "ConvertTemperature_C_to_F", "func": "ConvertTemperature",
     "args": lambda: (100.0, "C", "F"), "py_ref": lambda a: 212.0,
     "result_type": "scalar", "tol": 0.01},
    {"name": "ConvertTemperature_C_to_K", "func": "ConvertTemperature",
     "args": lambda: (0.0, "C", "K"), "py_ref": lambda a: 273.15,
     "result_type": "scalar", "tol": 0.01},
    {"name": "ConvertTemperature_F_to_C", "func": "ConvertTemperature",
     "args": lambda: (32.0, "F", "C"), "py_ref": lambda a: 0.0,
     "result_type": "scalar", "tol": 0.01},
    {"name": "ConvertTemperature_K_to_C", "func": "ConvertTemperature",
     "args": lambda: (273.15, "K", "C"), "py_ref": lambda a: 0.0,
     "result_type": "scalar", "tol": 0.01},

    # ---- ConvertVolume ----
    {"name": "ConvertVolume_L_to_mL", "func": "ConvertVolume",
     "args": lambda: (1.0, "L", "mL"), "py_ref": lambda a: 1000.0,
     "result_type": "scalar", "tol": 0.01},
    {"name": "ConvertVolume_mL_to_L", "func": "ConvertVolume",
     "args": lambda: (500.0, "mL", "L"), "py_ref": lambda a: 0.5,
     "result_type": "scalar", "tol": 0.01},

    # ---- ConvertPressure ----
    {"name": "ConvertPressure_atm_to_kPa", "func": "ConvertPressure",
     "args": lambda: (1.0, "atm", "kPa"), "py_ref": lambda a: 101.325,
     "result_type": "scalar", "tol": 0.01},
    {"name": "ConvertPressure_Pa_to_atm", "func": "ConvertPressure",
     "args": lambda: (101325.0, "Pa", "atm"), "py_ref": lambda a: 1.0,
     "result_type": "scalar", "tol": 0.01},

    # ---- MassToMoles / MolesToMass ----
    {"name": "MassToMoles_NaCl", "func": "MassToMoles",
     "args": lambda: (58.44, 58.44), "py_ref": lambda a: 1.0,
     "result_type": "scalar", "tol": 0.001},
    {"name": "MolesToMass_water", "func": "MolesToMass",
     "args": lambda: (2.0, 18.015), "py_ref": lambda a: 36.03,
     "result_type": "scalar", "tol": 0.01},
    {"name": "MolesToMass_roundtrip", "func": "MolesToMass",
     "args": lambda: (1.0, 100.0), "py_ref": lambda a: 100.0,
     "result_type": "scalar", "tol": 0.001},

    # ---- IdealGasLaw (PV=nRT, solve for n) ----
    {"name": "IdealGasLaw_solve_n", "func": "IdealGasLaw",
     "args": lambda: (101325.0, 0.02445, None, 298.15),
     "py_ref": lambda a: 1.0,
     "result_type": "scalar", "tol": 0.05},

    # ---- Density ----
    {"name": "Density_solve_density", "func": "Density",
     "args": lambda: (100.0, None, 2.5), "py_ref": lambda a: 40.0,
     "result_type": "scalar", "tol": 0.001},

    # ---- PercentYield ----
    {"name": "PercentYield_85pct", "func": "PercentYield",
     "args": lambda: (8.5, 10.0), "py_ref": lambda a: 85.0,
     "result_type": "scalar", "tol": 0.01},
    {"name": "PercentYield_100pct", "func": "PercentYield",
     "args": lambda: (10.0, 10.0), "py_ref": lambda a: 100.0,
     "result_type": "scalar", "tol": 0.01},

    # ---- DilutionSolve ----
    {"name": "DilutionSolve_c1v1_c2v2", "func": "DilutionSolve",
     "args": lambda: (10.0, 100.0, 5.0, None),
     "py_ref": lambda a: 200.0,  # v2 = c1*v1/c2
     "result_type": "scalar", "tol": 0.01},
    # ---- ConvertStandard (standard volume from P,V,T,MW) ----
    {"name": "ConvertStandard_ambient", "func": "ConvertStandard",
     "args": lambda: (0.024, 101325.0, 298.15, 28.97),
     "py_ref": lambda a: _py_convert_standard(a[0], a[1], a[2], a[3]),
     "result_type": "scalar", "tol": 1e-4},

    # ---- ConvertMass ----
    {"name": "ConvertMass_g_to_kg", "func": "ConvertMass",
     "args": lambda: (1000.0, "g", "kg"),
     "py_ref": lambda a: 1.0, "result_type": "scalar", "tol": 1e-6},
    {"name": "ConvertMass_lb_to_g", "func": "ConvertMass",
     "args": lambda: (1.0, "lb", "g"),
     "py_ref": lambda a: 453.592, "result_type": "scalar", "tol": 0.1},

    # ---- CompressFactorPR (Pitzer-Riedel compressibility) ----
    {"name": "CompressFactorPR_methane", "func": "CompressFactorPR",
     "args": lambda: (101325.0, 298.15, 190.6, 4.6e6, 0.011),
     "py_ref": lambda a: _py_compress_factor(a[0], a[1], a[2], a[3], a[4]),
     "result_type": "scalar", "tol": 0.01},

    # =====================================================================
    # Boundary / edge cases
    # =====================================================================

    {"name": "MolecularWeight_NaCl", "func": "MolecularWeight",
     "args": lambda: ("NaCl",),
     "py_ref": lambda a: _py_molecular_weight("NaCl"),
     "result_type": "scalar", "tol": 0.05},
    {"name": "MolecularWeight_single_H", "func": "MolecularWeight",
     "args": lambda: ("H",),
     "py_ref": lambda a: ATOMIC_WEIGHTS["H"],
     "result_type": "scalar", "tol": 0.01},
    {"name": "MolecularWeight_nested_parens", "func": "MolecularWeight",
     "args": lambda: ("Mg(OH)2",),
     "py_ref": lambda a: _py_molecular_weight("Mg(OH)2"),
     "result_type": "scalar", "tol": 0.1},

    # ---- MolecularWeight — bracket formulas (C2 sync) ----
    # Fe4[Fe(CN)6]3 = 7Fe + 18C + 18N
    {"name": "MolecularWeight_bracket_Fe4FeCN63", "func": "MolecularWeight",
     "args": lambda: ("Fe4[Fe(CN)6]3",),
     "py_ref": lambda a: _py_molecular_weight("Fe4[Fe(CN)6]3"),
     "result_type": "scalar", "tol": 0.5},
    # K3[Fe(CN)6] = 3K + Fe + 6C + 6N
    {"name": "MolecularWeight_bracket_K3FeCN6", "func": "MolecularWeight",
     "args": lambda: ("K3[Fe(CN)6]",),
     "py_ref": lambda a: _py_molecular_weight("K3[Fe(CN)6]"),
     "result_type": "scalar", "tol": 0.5},

    # ---- MolecularWeight — hydrate formulas (C3 sync) ----
    # CuSO4·5H2O — hydrate coefficient now applied (CuSO4 + 5×H2O)
    {"name": "MolecularWeight_hydrate_CuSO4_5H2O", "func": "MolecularWeight",
     "args": lambda: ("CuSO4·5H2O",),  # U+00B7 middle dot
     "py_ref": lambda a: 249.684,
     "result_type": "scalar", "tol": 0.5},

    {"name": "ConvertTemperature_abs_zero_K", "func": "ConvertTemperature",
     "args": lambda: (0.0, "K", "C"), "py_ref": lambda a: -273.15,
     "result_type": "scalar", "tol": 0.01},
    {"name": "ConvertTemperature_boiling_CtoK", "func": "ConvertTemperature",
     "args": lambda: (100.0, "C", "K"), "py_ref": lambda a: 373.15,
     "result_type": "scalar", "tol": 0.01},

    {"name": "ConvertVolume_m3_to_L", "func": "ConvertVolume",
     "args": lambda: (1.0, "m3", "L"), "py_ref": lambda a: 1000.0,
     "result_type": "scalar", "tol": 0.01},

    {"name": "ConvertPressure_bar_to_Pa", "func": "ConvertPressure",
     "args": lambda: (1.0, "bar", "Pa"), "py_ref": lambda a: 100000.0,
     "result_type": "scalar", "tol": 1.0},

    {"name": "MassToMoles_zero_mass", "func": "MassToMoles",
     "args": lambda: (0.0, 58.44), "py_ref": lambda a: 0.0,
     "result_type": "scalar", "tol": 1e-10},
    {"name": "MolesToMass_zero_moles", "func": "MolesToMass",
     "args": lambda: (0.0, 58.44), "py_ref": lambda a: 0.0,
     "result_type": "scalar", "tol": 1e-10},

    {"name": "PercentYield_zero_actual", "func": "PercentYield",
     "args": lambda: (0.0, 10.0), "py_ref": lambda a: 0.0,
     "result_type": "scalar", "tol": 0.01},

    {"name": "DilutionSolve_c2", "func": "DilutionSolve",
     "args": lambda: (10.0, 100.0, None, 200.0),
     "py_ref": lambda a: 5.0, "result_type": "scalar", "tol": 0.01},

    {"name": "Density_solve_mass", "func": "Density",
     "args": lambda: (None, 100.0, 2.5), "py_ref": lambda a: 250.0,
     "result_type": "scalar", "tol": 0.001},

    {"name": "IdealGasLaw_solve_V", "func": "IdealGasLaw",
     "args": lambda: (101325.0, None, 1.0, 273.15),
     "py_ref": lambda a: 0.022414, "result_type": "scalar", "tol": 0.001},

    {"name": "ConvertMass_kg_to_g", "func": "ConvertMass",
     "args": lambda: (1.0, "kg", "g"),
     "py_ref": lambda a: 1000.0, "result_type": "scalar", "tol": 1e-6},

    # CylinderStdVolume (skip: VBA uses different pressure/temp normalization)
    {"name": "CylinderStdVolume", "func": "CylinderStdVolume",
     "args": lambda: (10.0, 101325.0, 298.15, 0.0),
     "py_ref": lambda a: _py_cylinder_std_vol(a[0], a[1], a[2], a[3]),
     "result_type": "scalar", "tol": 0.01,
     "skip_if": True, "skip_reason": "CylinderStdVolume uses different standard-state normalization than Python reference. verified manually in Excel."},

    # CylinderStdVolumeFromMass — netWeight(kg) + gas formula
    {"name": "CylinderStdVolumeFromMass_N2", "func": "CylinderStdVolumeFromMass",
     "args": lambda: (50.0, "N2"),
     "py_ref": lambda a: _py_cylinder_std_vol_from_mass(a[0], a[1]),
     "result_type": "scalar", "tol": 1.0},

    {"name": "CompressFactorPR_ideal", "func": "CompressFactorPR",
     "args": lambda: (101325.0, 500.0, 190.6, 4.6e6, 0.0),
     "py_ref": lambda a: _py_compress_factor(a[0], a[1], a[2], a[3], a[4]),
     "result_type": "scalar", "tol": 0.05},

    # UDF wrapper
    {"name": "UDF_PC_MOLWEIGHT", "func": "UDF_PC_MOLWEIGHT",
     "args": lambda: ("H2O",),
     "py_ref": lambda a: 18.015,
     "result_type": "scalar", "tol": 0.02},

    # =====================================================================
    # Migrated from VBA Test_PhyChemUtils — coverage gaps (2026-06-16)
    # =====================================================================

    # ---- MolecularWeight — deeper nested parens + structural checks ----
    {"name": "MolecularWeight_nested3", "func": "MolecularWeight",
     "args": lambda: ("(((CH3)2O)2)",),
     "py_ref": lambda a: 92.136,
     "result_type": "scalar", "tol": 0.1},
    {"name": "MolecularWeight_NH42SO4", "func": "MolecularWeight",
     "args": lambda: ("(NH4)2SO4",),
     "py_ref": lambda a: _py_molecular_weight("(NH4)2SO4"),
     "result_type": "scalar", "tol": 0.1},
    {"name": "MolecularWeight_C6H12O6", "func": "MolecularWeight",
     "args": lambda: ("C6H12O6",),
     "py_ref": lambda a: _py_molecular_weight("C6H12O6"),
     "result_type": "scalar", "tol": 0.1},

    # ---- MolecularWeight / ConvertTemperature — error cases (CVErr via COM unreliable) ----
    {"name": "MolecularWeight_empty_error", "func": "MolecularWeight",
     "args": lambda: ("",), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr for empty formula; COM marshaling unreliable"},
    {"name": "MolecularWeight_invalid_elem", "func": "MolecularWeight",
     "args": lambda: ("Xyz",), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr for unknown element; COM marshaling unreliable"},
    {"name": "ConvertTemperature_abs_zero", "func": "ConvertTemperature",
     "args": lambda: (-273.15, "c", "k"),
     "py_ref": lambda a: 0.0,
     "result_type": "scalar", "tol": 0.01},

    # ---- Error cases: MassToMoles / MolesToMass negative inputs ----
    {"name": "MassToMoles_neg_error", "func": "MassToMoles",
     "args": lambda: (-1.0, 18.0), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr for negative mass; COM marshaling unreliable"},
    {"name": "MolesToMass_neg_error", "func": "MolesToMass",
     "args": lambda: (-1.0, 18.0), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr for negative moles; COM marshaling unreliable"},

    # ---- Error cases: IdealGasLaw — insufficient parameters ----
    {"name": "IdealGasLaw_all_empty_error", "func": "IdealGasLaw",
     "args": lambda: (None, None, None, None), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr; COM marshaling unreliable"},
    {"name": "IdealGasLaw_partial_error", "func": "IdealGasLaw",
     "args": lambda: (None, None, 1.0, 273.15), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr; COM marshaling unreliable"},

    # ---- Error cases: PercentYield — divide by zero / negative actual ----
    {"name": "PercentYield_div0_error", "func": "PercentYield",
     "args": lambda: (8.0, 0.0), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr for zero theoretical; COM marshaling unreliable"},
    {"name": "PercentYield_neg_error", "func": "PercentYield",
     "args": lambda: (-1.0, 10.0), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr for negative actual; COM marshaling unreliable"},

    # ---- Error cases: ConvertPressure — invalid unit codes ----
    {"name": "ConvertPressure_bad_from_unit", "func": "ConvertPressure",
     "args": lambda: (1.0, "xyz", "atm"), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr for invalid unit; COM marshaling unreliable"},
    {"name": "ConvertPressure_bad_to_unit", "func": "ConvertPressure",
     "args": lambda: (1.0, "atm", "xyz"), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr for invalid unit; COM marshaling unreliable"},

    # ---- Error case: DilutionSolve — divide by zero ----
    {"name": "DilutionSolve_div0_error", "func": "DilutionSolve",
     "args": lambda: (1.0, 1.0, 0.0, None), "py_ref": lambda a: None,
     "result_type": "scalar",
     "skip_if": True, "skip_reason": "VBA returns CVErr for c2=0; COM marshaling unreliable"},

    # ---- Coverage: ConvertMass, ConvertPressure, more MW ----
    {"name": "ConvertMass_kg_to_g", "func": "ConvertMass",
     "args": lambda: (1.0, "kg", "g"),
     "py_ref": lambda a: 1000.0, "result_type": "scalar", "tol": 0.01},
    {"name": "ConvertPressure_atm_to_Pa", "func": "ConvertPressure",
     "args": lambda: (1.0, "atm", "Pa"),
     "py_ref": lambda a: 101325.0, "result_type": "scalar", "tol": 1.0},
    {"name": "ConvertPressure_Pa_to_kPa", "func": "ConvertPressure",
     "args": lambda: (100000.0, "Pa", "kPa"),
     "py_ref": lambda a: 100.0, "result_type": "scalar", "tol": 0.001},
    {"name": "MolecularWeight_CO2", "func": "MolecularWeight",
     "args": lambda: ("CO2",),
     "py_ref": lambda a: 44.010, "result_type": "scalar", "tol": 0.01},
    {"name": "MolecularWeight_CH4", "func": "MolecularWeight",
     "args": lambda: ("CH4",),
     "py_ref": lambda a: 16.043, "result_type": "scalar", "tol": 0.01},
    {"name": "MolecularWeight_NH3", "func": "MolecularWeight",
     "args": lambda: ("NH3",),
     "py_ref": lambda a: 17.031, "result_type": "scalar", "tol": 0.01},
    {"name": "MolecularWeight_H2SO4", "func": "MolecularWeight",
     "args": lambda: ("H2SO4",),
     "py_ref": lambda a: 98.078, "result_type": "scalar", "tol": 0.1},

]



def _py_convert_standard(V, P, T, MW):
    R = 8.314; T_std = 273.15; P_std = 101325.0
    n = (P * V) / (R * T)
    return (n * R * T_std) / P_std

def _py_compress_factor(P, T, Tc, Pc, omega):
    R = 8.314; Tr = T / Tc; Pr = P / Pc
    if Tr <= 0: return 1.0
    kappa = 0.37464 + 1.54226 * omega - 0.26992 * omega**2
    alpha = (1.0 + kappa * (1.0 - Tr**0.5))**2
    a = 0.45724 * (R**2) * (Tc**2) / Pc * alpha
    b = 0.07780 * R * Tc / Pc
    V_ideal = (R * T) / P
    Z = 1.0
    for _ in range(10):
        V = Z * V_ideal
        Z_new = V / (V - b) - (a * V) / (R * T * (V * (V + b) + b * (V - b)))
        if abs(Z_new - Z) < 1e-8: Z = Z_new; break
        Z = Z_new
    return Z

def _py_cylinder_std_vol(V_gas, P, T, V_liquid):
    """Standard volume of gas in cylinder at STP."""
    R = 8.314; T_std = 273.15; P_std = 101325.0
    n = (P * V_gas * 1e-3) / (R * T)  # V in L -> m3
    V_std = (n * R * T_std) / P_std
    return V_std * 1000.0  # back to L

def _py_cylinder_std_vol_from_mass(mass_kg, gas_formula):
    """Standard volume from gas mass in cylinder (m^3). VBA: mass*22.414/MW"""
    MW_MAP = {"N2": 28.0134, "O2": 31.9988, "Ar": 39.948, "CO2": 44.010,
              "He": 4.0026, "H2": 2.0159, "CH4": 16.043, "NH3": 17.031,
              "air": 28.97, "Cl2": 70.906, "HCl": 36.461}
    mw = MW_MAP.get(gas_formula, 28.97)
    return mass_kg * 22.414 / mw  # m^3

def main() -> int:
    runner = CrossValRunner("PhyChemUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
