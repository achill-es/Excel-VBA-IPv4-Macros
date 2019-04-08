Public Function Nslookup(ByVal hostname As String) As String

    Const TemporaryFolder = 2
    Const ForReading = 1
    
    Dim objFSO As Object
    Dim tmpFileName As String
    Dim objShell As Object
    Dim objTempFile As Object
    Dim cmdStr As String
    Dim intFound As Integer


    Set objFSO = CreateObject("Scripting.FileSystemObject")
    tmpFileName = objFSO.BuildPath(objFSO.GetSpecialFolder(TemporaryFolder), objFSO.GetTempName)
    Set objShell = CreateObject("Wscript.Shell")

    'Run the nslookup command
    objShell.Run "cmd /c nslookup " & hostname & " > " & tmpFileName, 0, True

    cmdStr = ""
    Set objTempFile = objFSO.OpenTextFile(tmpFileName, ForReading)

    Do While objTempFile.AtEndOfStream <> True
        cmdStr = cmdStr & Trim(objTempFile.Readline) & vbCrLf
    Loop

    objTempFile.Close
    objFSO.DeleteFile tmpFileName, True

    'Process the result
    intFound = InStr(1, cmdStr, "Name:", vbTextCompare)
    If intFound = 0 Then
        Nslookup = ""
    Else
        loc1 = InStr(intFound, cmdStr, "Address:", vbTextCompare) + InStr(intFound, cmdStr, "Addresses:", vbTextCompare)
        loc2 = InStr(loc1, cmdStr, vbCrLf, vbTextCompare)
        Nslookup = Trim(Mid(cmdStr, loc1 + 8, loc2 - loc1 - 8))
    End If

End Function
