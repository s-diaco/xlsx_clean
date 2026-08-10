' XlsxClean.vbs - double-click launcher (no console window)
' Resolves repo root from this script's folder (desktop\ -> repo root).
Option Explicit

Dim fso, sh, scriptDir, root, pythonExe, cmd
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
root = fso.GetParentFolderName(scriptDir)

' Prefer project venv, then py launcher, then python on PATH.
pythonExe = root & "\.venv\Scripts\pythonw.exe"
If Not fso.FileExists(pythonExe) Then
  pythonExe = root & "\.venv\Scripts\python.exe"
End If
If Not fso.FileExists(pythonExe) Then
  pythonExe = "py"
End If

sh.CurrentDirectory = root
' Window style 0 = hidden.
cmd = """" & pythonExe & """ -m xlsx_clean.gui_app"
If pythonExe = "py" Then
  cmd = "py -3 -m xlsx_clean.gui_app"
End If

sh.Run cmd, 0, False
