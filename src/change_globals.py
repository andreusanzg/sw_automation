import win32com.client

sw = win32com.client.Dispatch("SldWorks.Application")
model = sw.ActiveDoc

if model is None:
    raise RuntimeError("No hay documento activo")

eq_mgr = model.GetEquationMgr

updated = False

for i in range(eq_mgr.GetCount):
    eq = eq_mgr.Equation(i)

    if '"Width"' in eq:
        eq_mgr.Equation(i, '"Width" = 100')   # 250 mm
        updated = True

    if '"Height"' in eq:
        eq_mgr.Equation(i, '"Height" = 50')  # 120 mm
        updated = True

    if '"Depth"' in eq:
        eq_mgr.Equation(i, '"Depth" = 80')  # 80 mm
        updated = True

if not updated:
    raise RuntimeError("No se encontraron las variables esperadas")

# ✅ Rebuild correcto
model.ForceRebuild3(False)

print("✅ Global Variables actualizadas")