Attribute VB_Name = "Module1"
Public Sub SaveDesktop()
Set WshShell = WScript.CreateObject("WScript.Shell")
strDesktop = WshShell.SpecialFolders("Desktop")
Set oShellLink = WshShell.CreateShortcut(strDesktop & "\Starboy Notepad.lnk")

oShellLink.TargetPath = Form1.Text1.Text & "\StarboyKB75.exe" 'WScript.ScriptFullName
oShellLink.WindowStyle = 1
oShellLink.Hotkey = "CTRL+SHIFT+N"
oShellLink.IconLocation = Form1.Text1.Text & "\StarboyKB75.exe, 0"
oShellLink.Description = "Shortcut To Starboy Notepad."
oShellLink.WorkingDirectory = strDesktop
oShellLink.Save
End Sub
