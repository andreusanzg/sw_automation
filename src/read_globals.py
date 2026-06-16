import win32com.client

sw = win32com.client.Dispatch("SldWorks.Application")
model = sw.ActiveDoc

if model is None:
    raise RuntimeError("No hay documento activo")

eq_mgr = model.GetEquationMgr

count = eq_mgr.GetCount

print("🌍 Global Variables encontradas:\n")

for i in range(count):
    equation = eq_mgr.Equation(i)
    print(equation)
