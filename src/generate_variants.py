import csv
import time
from pathlib import Path
import win32com.client
import pythoncom

# ---------------- CONFIGURACIÓN ----------------


BASE_DIR = Path(__file__).resolve().parent
PART_PATH = BASE_DIR.parent / "data" / "caja.sldprt"
CSV_PATH = BASE_DIR.parent / "data" / "variants_example.csv"
OUTPUT_DIR = BASE_DIR.parent / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

SW_DOC_PART = 1
SW_SAVE_SILENT = 1

# ---------------- SOLIDWORKS ----------------
sw = win32com.client.Dispatch("SldWorks.Application")
sw.Visible = False

# ---------------- CSV ----------------
with open(CSV_PATH, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f)

    for row in reader:
        name = row["name"]
        print(f"\n🔧 Generando variante: {name}")

        # ---------- 1️⃣ Abrir SIEMPRE la pieza base ----------
        spec = sw.GetOpenDocSpec(str(PART_PATH))
        spec.DocumentType = SW_DOC_PART
        spec.Silent = True
        spec.ReadOnly = False

        model = sw.OpenDoc7(spec)
        if model is None:
            raise RuntimeError("No se pudo abrir la pieza base")

        eq_mgr = model.GetEquationMgr

        # ---------- 2️⃣ Modificar Global Variables ----------
        for i in range(eq_mgr.GetCount):
            eq = eq_mgr.Equation(i)

            for var, value in row.items():
                if var == "name":
                    continue

                if eq.startswith(f'"{var}"'):
                    eq_mgr.Equation(
                        i,
                        f'"{var}" = {float(value)}'
                    )

        # ---------- 3️⃣ Resolver y reconstruir ----------
        eq_mgr.EvaluateAll
        model.ForceRebuild3(True)
        model.ForceRebuild3(True)

        time.sleep(0.05)

        # ---------- 4️⃣ GUARDAR VARIANTE (AQUÍ VA Save3) ----------
        output_path = OUTPUT_DIR / f"{name}.SLDPRT"

        # Fijar el nombre UNA vez
        model.SaveAs(str(output_path))

        # Guardar de forma robusta
        errors = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, 0
        )
        warnings = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, 0
        )

        ok = model.Save3(
            SW_SAVE_SILENT,
            errors,
            warnings
        )

        if not ok:
            raise RuntimeError(
                f"Save3 falló: errors={errors.value}, warnings={warnings.value}"
            )

        print(f"✅ Guardado: {output_path.name}")

        # ---------- 5️⃣ Cerrar documento (CRÍTICO) ----------
        sw.CloseDoc(model.GetTitle)
sw.ExitApp()
print("\n🎉 Todas las variantes generadas correctamente")