Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
appPath = scriptDir & "\src\borrowing_base_workbench\app.py"
pythonwPath = shell.ExpandEnvironmentStrings("%LocalAppData%") & "\Microsoft\WindowsApps\pythonw.exe"

shell.Run """" & pythonwPath & """ """ & appPath & """", 0, False
