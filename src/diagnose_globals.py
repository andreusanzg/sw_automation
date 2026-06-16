import win32com.client

sw = win32com.client.Dispatch("SldWorks.Application")
model = sw.ActiveDoc

eq_mgr = model.GetEquationMgr

print("=== ECUACIONES ===")
for i in range(eq_mgr.GetCount):
    print(eq_mgr.Equation(i))

print("\n=== DIMENSIONES ===")
feat = model.FirstFeature()
while feat:
    disp = feat.GetFirstDisplayDimension()
    while disp:
        dim = disp.GetDimension()
        if dim:
            print(dim.FullName, "=", dim.SystemValue)
        disp = feat.GetNextDisplayDimension(disp)
    feat = feat.GetNextFeature()