Option Explicit

Dim objShell, objFSO, objFile, objInput, objOutput, strPassword, strSite, strLine

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.OpenTextFile("passwords.txt", 8, True)

Do
    strSite = InputBox("Enter the name of the website or application", "Password Manager")
    If strSite = "" Then Exit Do
    strPassword = InputBox("Enter the password for " & strSite, "Password Manager")
    If strPassword = "" Then Exit Do
    objFile.WriteLine(strSite & "," & strPassword)
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile("passwords.txt", 1, False)

Do Until objFile.AtEndOfStream
    strLine = objFile.ReadLine
    If InStr(strLine, ",") > 0 Then
        strSite = Split(strLine, ",")(0)
        strPassword = Split(strLine, ",")(1)
        If strSite <> "" And strPassword <> "" Then
            objOutput = objOutput & strSite & vbCrLf
        End If
    End If
Loop

objFile.Close

objShell.Popup objOutput, 0, "Password Manager"
