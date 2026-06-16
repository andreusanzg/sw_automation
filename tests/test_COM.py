import win32com.client

shell = win32com.client.Dispatch("WScript.Shell")
print("COM OK:", shell)
