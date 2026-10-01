Set fso = CreateObject("Scripting.FileSystemObject")
pasta = fso.GetParentFolderName(WScript.ScriptFullName)
cmd = "cmd /c """ & pasta & "\iniciar_servidor.bat"""
CreateObject("WScript.Shell").Run cmd, 0, False
