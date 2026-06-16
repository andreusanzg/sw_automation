import win32com.client

try:
    sw = win32com.client.Dispatch("SldWorks.Application")
    sw.Visible = True
    print("✅ SolidWorks COM OK")
    print("Version:", sw.RevisionNumber)
except Exception as e:
    print("❌ SolidWorks COM FAILED")
    raise