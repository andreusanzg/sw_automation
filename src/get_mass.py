import win32com.client
import json

sw = win32com.client.Dispatch("SldWorks.Application")
model = sw.ActiveDoc

if model is None:
    print(json.dumps({"error": "No active SolidWorks document"}))
    exit(1)

mass = model.Extension.CreateMassProperty.Mass

print(json.dumps({"mass": mass}))
