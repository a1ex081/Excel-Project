Option Explicit

Sub RunPythonScript()

'Declare CMD and Python Variables
Dim objShell As Object
Dim PythonExe, PythonScript As String

'Create a New Shell Object
Set objShell = VBA.CreateObject("Wscript.Shell")

'Provide the file path to the Python Exe
PythonExe = """C:\Python38\python.exe"""

'Provide the file path to the python Script
PythonScript = """C:\Users\alexb\Documents\GitHub\Excel-Project\Python Script.py"""

'Run the Python Script
objShell.Run PythonExe & PythonScript

End Sub 