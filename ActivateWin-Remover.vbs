Dim objShell
Set objShell = WScript.CreateObject("WScript.Shell")

strScriptDir = Replace(WScript.ScriptFullName, WScript.ScriptName, "")
strPythonScript = strScriptDir & "remover.py"

objShell.Run "python """ & strPythonScript & """", 1, True

Set objShell = Nothing
