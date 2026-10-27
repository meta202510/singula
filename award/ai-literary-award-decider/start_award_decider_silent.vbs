Option Explicit

Dim shell, fso, folder, launcher, checkResult, appPath
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
folder = fso.GetParentFolderName(WScript.ScriptFullName)
appPath = fso.BuildPath(folder, "launch_award_decider.pyw")

If CommandExists("pyw") Then
    launcher = "pyw -3"
ElseIf CommandExists("pythonw") Then
    launcher = "pythonw"
Else
    MsgBox "Python 3 was not found. Install Python 3 with Tcl/Tk enabled, then try again.", 16, "Startup error"
    WScript.Quit 1
End If

checkResult = shell.Run(launcher & " -c ""import tkinter""", 0, True)
If checkResult <> 0 Then
    MsgBox "Tkinter is not available. Repair or reinstall Python with Tcl/Tk enabled.", 16, "Startup error"
    WScript.Quit 1
End If

shell.Run launcher & " """ & appPath & """", 0, False

Function CommandExists(commandName)
    CommandExists = (shell.Run("cmd.exe /c where " & commandName & " >nul 2>nul", 0, True) = 0)
End Function
