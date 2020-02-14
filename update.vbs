sourceURL = "https://github.com/v2ray/v2ray-core/releases/latest/download/v2ray-windows-64.zip"
path = "v2ray-windows-64.zip"

answer = MsgBox("Press OK to proceed to downloading this file: " & vbcrlf & vbcrlf & sourceURL & vbcrlf & vbcrlf & "You will be notified if the download is successful.", 1, "Confirmation")
If answer = 2 Then
    Wscript.Echo "Cancelled."
    WScript.Quit
End If

Dim xHttp: Set xHttp = createobject("MSXML2.ServerXMLHTTP")
Dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.setOption 2, 13056
xHttp.Open "GET", sourceURL, False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile path, 2 '//overwrite
end with

extractTo = "v2ray"
zipFile = path

Set fso = CreateObject("Scripting.FileSystemObject")
If NOT fso.FolderExists(extractTo) Then
    fso.CreateFolder(extractTo)
ElseIf fso.FileExists(extractTo & "\" & "config.json") Then
    If fso.FileExists(extractTo & "\" & "config.jsonbak") Then
        fso.DeleteFile extractTo & "\" & "config.jsonbak"
    End If
    fso.MoveFile extractTo & "\" & "config.json", extractTo & "\" & "config.jsonbak"
End If


parentDir = fso.GetParentFolderName(WScript.ScriptFullName) & "\"
dim objShell: Set objShell = CreateObject("Shell.Application")
dim filesInZip: Set filesInZip = objShell.NameSpace(parentDir & zipFile).items
objShell.NameSpace(parentDir & extractTo).CopyHere(filesInZip)

If fso.FileExists(extractTo & "\" & "config.jsonbak") Then
    fso.DeleteFile extractTo & "\" & "config.json"
    fso.MoveFile extractTo & "\" & "config.jsonbak", extractTo & "\" & "config.json"
End If

Wscript.Echo "Success!"