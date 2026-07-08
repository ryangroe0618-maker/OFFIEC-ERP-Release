Option Explicit

Dim shell
Dim fso
Dim root
Dim command

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetParentFolderName(WScript.ScriptFullName)

command = "cmd.exe /c """ & root & "\START_WINDOWS.cmd"""
shell.Run command, 0, False
