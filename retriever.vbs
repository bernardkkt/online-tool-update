' File downloader with Microsoft VBScript
'
' Usage:
'   wscript.exe retriever.vbs [URL] [local path with filename] [options]
'
' Options:
'   /p - Prompt for confirmation before running download task

Set objArgs = Wscript.Arguments
sourceURL = objArgs(0)
path = objArgs(1)
silent = "None"

If objArgs.Count > 2 then
    silent = objArgs(2)
    If silent = "/p" Then
        Call Ask()
    End If
End If

Dim xHttp: Set xHttp = createobject("MSXML2.ServerXMLHTTP")
Dim bStrm: Set bStrm = createobject("Adodb.Stream")
' Workaround for certificates issue: see https://stackoverflow.com/a/9238141
xHttp.setOption 2, 13056
xHttp.Open "GET", sourceURL, False
xHttp.Send

with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile path, 2 '//overwrite
end with

If silent = "/p" Then
        Wscript.Echo "Success!"
End If

Sub Ask()
    answer = MsgBox("Press OK to proceed to downloading this file: " & vbcrlf & vbcrlf & sourceURL & vbcrlf & vbcrlf & "You will be notified if the update is successful.", 1, "Confirmation")
    If answer = 2 Then
        Wscript.Echo "Cancelled."
        WScript.Quit 1
    End If
End Sub
