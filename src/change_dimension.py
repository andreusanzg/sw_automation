import win32com.client
from pathlib import Path

# ----- CONSTANTES -----
SW_DOC_PART = 1
SW_OPEN_SILENT = 1

BASE_DIR = Path(__file__).resolve().parent

PART_PATH = BASE_DIR.parent / "data" / "caja.sldprt"

sw = win32com.client.Dispatch("SldWorks.Application")
sw.Visible = True

# --- Crear especificación de apertura ---
spec = sw.GetOpenDocSpec(PART_PATH)

spec.DocumentType = SW_DOC_PART
spec.Silent = True
spec.ReadOnly = False

# --- Abrir documento  ---
model = sw.OpenDoc7(spec)

if model is None:
    raise RuntimeError("No se pudo abrir la pieza")

print("✅ Pieza abierta:", model.GetTitle)

# --- CAMBIAR DIMENSIÓN ---
dim = model.Parameter("D1@Croquis1")


if dim is None:
    raise RuntimeError("Dimensión no encontrada")

print("Valor actual (m):", dim.SystemValue)

dim.SystemValue = 0.05  # 50 mm

# ✅ REBUILD CORRECTO
model.ForceRebuild3(False)

print("✅ Dimensión modificada y modelo reconstruido")

