Attribute VB_Name = "AIExtract_Mod1"

Option Explicit
' ============================================================
' MODULE: modAPI (V3 - Cross-platform Windows + Mac)
' PURPOSE: Send chat text + prompt to Claude API and return response
' Detects OS at runtime and uses the right HTTP mechanism.
' ============================================================

#If Mac Then
    Private Const IS_MAC As Boolean = True
#Else
    Private Const IS_MAC As Boolean = False
#End If

Public Function CallClaudeAPI(ByVal promptText As String) As String
    Dim apiKey As String, apiUrl As String, apiVersion As String, model As String
    Dim maxTokens As Long, temperature As Double, timeoutSec As Long, maxRetries As Long
    Dim payload As String
   
    ' Read settings
    apiKey = GetSetting("Anthropic API Key")
    apiUrl = GetSetting("API Endpoint")
    apiVersion = GetSetting("API Version")
    model = GetSetting("Model")
    maxTokens = CLng(GetSetting("Max Tokens (Response)"))
    temperature = CDbl(GetSetting("Temperature"))
    timeoutSec = CLng(GetSetting("Request Timeout (sec)"))
    maxRetries = CLng(GetSetting("Max Retries On Failure"))
   
    If apiKey = "" Or InStr(apiKey, "<<<") > 0 Then
        CallClaudeAPI = "ERROR: API key not set in Settings sheet"
        Exit Function
    End If
   
    ' Build JSON payload
    payload = "{""model"":""" & model & """,""max_tokens"":" & maxTokens & _
              ",""temperature"":" & temperature & _
              ",""messages"":[{""role"":""user"",""content"":" & JsonEncode(promptText) & "}]}"
   
    ' Route to correct implementation based on OS
    If IS_MAC Then
        CallClaudeAPI = CallClaudeAPI_Mac(payload, apiKey, apiUrl, apiVersion, timeoutSec, maxRetries)
    Else
        CallClaudeAPI = CallClaudeAPI_Windows(payload, apiKey, apiUrl, apiVersion, timeoutSec, maxRetries)
    End If
End Function

' ============================================================
' WINDOWS implementation - uses MSXML2.ServerXMLHTTP
' ============================================================
Private Function CallClaudeAPI_Windows(ByVal payload As String, ByVal apiKey As String, _
                                       ByVal apiUrl As String, ByVal apiVersion As String, _
                                       ByVal timeoutSec As Long, ByVal maxRetries As Long) As String
    Dim http As Object
    Dim retry As Long, success As Boolean, Response As String
   
    For retry = 1 To maxRetries
        On Error Resume Next
        Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        http.setTimeouts timeoutSec * 1000, timeoutSec * 1000, timeoutSec * 1000, timeoutSec * 1000
        http.Open "POST", apiUrl, False
        http.setRequestHeader "Content-Type", "application/json"
        http.setRequestHeader "x-api-key", apiKey
        http.setRequestHeader "anthropic-version", apiVersion
        http.Send payload
       
        If Err.Number = 0 And http.status = 200 Then
            Response = http.ResponseText
            success = True
            On Error GoTo 0
            Exit For
        End If
        On Error GoTo 0
       
        If retry < maxRetries Then
            Application.Wait Now + TimeSerial(0, 0, 2)
        End If
    Next retry
   
    If Not success Then
        If http Is Nothing Then
            CallClaudeAPI_Windows = "ERROR: Could not connect to API"
        Else
            CallClaudeAPI_Windows = "ERROR: API returned status " & http.status & " - " & http.ResponseText
        End If
        Exit Function
    End If
   
    CallClaudeAPI_Windows = Response
End Function

' ============================================================
' MAC implementation - uses shell + curl + temp files
' ============================================================
Private Function CallClaudeAPI_Mac(ByVal payload As String, ByVal apiKey As String, _
                                   ByVal apiUrl As String, ByVal apiVersion As String, _
                                   ByVal timeoutSec As Long, ByVal maxRetries As Long) As String
#If Mac Then
    Dim tmpDir As String, tmpPayloadFile As String, tmpResponseFile As String
    Dim tmpConfigFile As String, tmpStderrFile As String
    Dim Response As String
    Dim retry As Long, success As Boolean
   
    tmpDir = Environ("TMPDIR")
    If Right(tmpDir, 1) <> "/" Then tmpDir = tmpDir & "/"
   
    Dim ts As String
    ts = Format(Now, "yyyymmddhhnnss") & "_" & Int(Rnd * 100000)
    tmpPayloadFile = tmpDir & "claude_payload_" & ts & ".json"
    tmpResponseFile = tmpDir & "claude_response_" & ts & ".json"
    tmpConfigFile = tmpDir & "claude_config_" & ts & ".txt"
    tmpStderrFile = tmpDir & "claude_stderr_" & ts & ".txt"
   
    ' Write payload as UTF-8
    On Error Resume Next
    WriteUTF8File tmpPayloadFile, payload
    If Err.Number <> 0 Then
        CallClaudeAPI_Mac = "ERROR: Cannot write payload file - " & Err.Description
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
   
    ' Build and write config file as UTF-8
    Dim cfg As String
    cfg = "url = """ & apiUrl & """" & vbLf & _
          "request = ""POST""" & vbLf & _
          "header = ""Content-Type: application/json""" & vbLf & _
          "header = ""x-api-key: " & apiKey & """" & vbLf & _
          "header = ""anthropic-version: " & apiVersion & """" & vbLf & _
          "data-binary = ""@" & tmpPayloadFile & """" & vbLf & _
          "max-time = " & timeoutSec & vbLf & _
          "output = """ & tmpResponseFile & """" & vbLf & _
          "silent" & vbLf
   
    On Error Resume Next
    WriteUTF8File tmpConfigFile, cfg
    If Err.Number <> 0 Then
        CallClaudeAPI_Mac = "ERROR: Cannot write config file - " & Err.Description
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
   
    ' Run curl
    Dim runCmd As String
    runCmd = "do shell script ""/usr/bin/curl --config '" & tmpConfigFile & "' 2> '" & tmpStderrFile & "'"""
   
    For retry = 1 To maxRetries
        On Error Resume Next
        MacScript runCmd
        On Error GoTo 0
       
        Response = ReadMacTextFile(tmpResponseFile)
       
        If Len(Response) > 0 And InStr(Response, """content""") > 0 Then
            success = True
            Exit For
        End If
       
        If retry < maxRetries Then
            Application.Wait Now + TimeSerial(0, 0, 2)
        End If
    Next retry
   
    ' Cleanup temp files
    On Error Resume Next
    Kill tmpPayloadFile
    Kill tmpResponseFile
    Kill tmpConfigFile
    Kill tmpStderrFile
    On Error GoTo 0
   
    If Not success Then
        If Len(Response) > 0 Then
            CallClaudeAPI_Mac = "ERROR: " & Left(Response, 500)
        Else
            CallClaudeAPI_Mac = "ERROR: No response from API. Check internet connection and API key."
        End If
        Exit Function
    End If
   
    CallClaudeAPI_Mac = Response
#Else
    CallClaudeAPI_Mac = "ERROR: Mac path called on non-Mac system"
#End If
End Function

Private Function ReadMacTextFile(ByVal path As String) As String
    Dim fnum As Integer, Result As String, lineText As String
    On Error Resume Next
    fnum = FreeFile
    Open path For Input As #fnum
    If Err.Number <> 0 Then
        ReadMacTextFile = ""
        On Error GoTo 0
        Exit Function
    End If
    Do While Not EOF(fnum)
        Line Input #fnum, lineText
        Result = Result & lineText & vbLf
    Loop
    Close #fnum
    On Error GoTo 0
    ReadMacTextFile = Result
End Function
Private Function ReadTextFileDiag(ByVal path As String) As String
    Dim fnum As Integer, Result As String, lineText As String
    On Error Resume Next
    fnum = FreeFile
    Open path For Input As #fnum
    If Err.Number <> 0 Then
        ReadTextFileDiag = ""
        Exit Function
    End If
    Do While Not EOF(fnum)
        Line Input #fnum, lineText
        Result = Result & lineText & vbLf
    Loop
    Close #fnum
    On Error GoTo 0
    ReadTextFileDiag = Result
End Function

Private Function EscapeForAppleScriptDiag(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    EscapeForAppleScriptDiag = t
End Function
' ============================================================
' Helper: Read entire text file into a string
' ============================================================
Private Function ReadTextFile(ByVal path As String) As String
    Dim fnum As Integer, Result As String, lineText As String
    On Error Resume Next
    fnum = FreeFile
    Open path For Input As #fnum
    If Err.Number <> 0 Then
        ReadTextFile = ""
        Exit Function
    End If
    Do While Not EOF(fnum)
        Line Input #fnum, lineText
        Result = Result & lineText & vbLf
    Loop
    Close #fnum
    On Error GoTo 0
    ReadTextFile = Result
End Function

' ============================================================
' Helper: Escape string for embedding in AppleScript "do shell script"
' ============================================================
Private Function EscapeForAppleScript(ByVal s As String) As String
    Dim t As String
    t = s
    ' AppleScript inside VBA double-quotes needs:
    ' - Backslash ? \\
    ' - Double quote ? \"
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    EscapeForAppleScript = t
End Function

' ============================================================
' JSON encoder - used by both platforms
' ============================================================
Public Function JsonEncode(ByVal s As String) As String
    Dim i As Long, Ch As String, Result As String
    Result = """"
    For i = 1 To Len(s)
        Ch = Mid(s, i, 1)
        Select Case Ch
            Case """": Result = Result & "\"""
            Case "\": Result = Result & "\\"
            Case vbCr: Result = Result & "\r"
            Case vbLf: Result = Result & "\n"
            Case vbTab: Result = Result & "\t"
            Case Chr(8): Result = Result & "\b"
            Case Chr(12): Result = Result & "\f"
            Case Else
                If Asc(Ch) < 32 Then
                    Result = Result & "\u" & Right("0000" & hex(Asc(Ch)), 4)
                Else
                    Result = Result & Ch
                End If
        End Select
    Next i
    JsonEncode = Result & """"
End Function

Private Sub WriteUTF8File(ByVal path As String, ByVal content As String)
    ' Convert string to UTF-8 byte array and write binary
    ' This works on both Mac and Windows
    Dim utf8Bytes() As Byte
    utf8Bytes = StrToUTF8Bytes(content)
   
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Binary Access Write As #fnum
    Put #fnum, 1, utf8Bytes
    Close #fnum
End Sub

Private Function StrToUTF8Bytes(ByVal s As String) As Byte()
    ' Convert VBA string (UTF-16) to UTF-8 byte array
    Dim i As Long, codePoint As Long, Ch As Long
    Dim bytes() As Byte
    Dim outIdx As Long
   
    ReDim bytes(0 To Len(s) * 4)  ' worst case: 4 bytes per char
    outIdx = 0
   
    i = 1
    Do While i <= Len(s)
        Ch = AscW(Mid(s, i, 1))
        If Ch < 0 Then Ch = Ch + 65536  ' AscW returns signed; convert to unsigned
       
        ' Check for UTF-16 surrogate pair
        If Ch >= &HD800 And Ch <= &HDBFF And i < Len(s) Then
            ' High surrogate - combine with low surrogate
            Dim low As Long
            low = AscW(Mid(s, i + 1, 1))
            If low < 0 Then low = low + 65536
            codePoint = &H10000 + (Ch - &HD800) * &H400 + (low - &HDC00)
            i = i + 2
        Else
            codePoint = Ch
            i = i + 1
        End If
       
        ' Encode code point as UTF-8
        If codePoint < &H80 Then
            ' 1 byte: 0xxxxxxx
            bytes(outIdx) = CByte(codePoint)
            outIdx = outIdx + 1
        ElseIf codePoint < &H800 Then
            ' 2 bytes: 110xxxxx 10xxxxxx
            bytes(outIdx) = CByte(&HC0 Or (codePoint \ &H40))
            bytes(outIdx + 1) = CByte(&H80 Or (codePoint And &H3F))
            outIdx = outIdx + 2
        ElseIf codePoint < &H10000 Then
            ' 3 bytes: 1110xxxx 10xxxxxx 10xxxxxx
            bytes(outIdx) = CByte(&HE0 Or (codePoint \ &H1000))
            bytes(outIdx + 1) = CByte(&H80 Or ((codePoint \ &H40) And &H3F))
            bytes(outIdx + 2) = CByte(&H80 Or (codePoint And &H3F))
            outIdx = outIdx + 3
        Else
            ' 4 bytes: 11110xxx 10xxxxxx 10xxxxxx 10xxxxxx
            bytes(outIdx) = CByte(&HF0 Or (codePoint \ &H40000))
            bytes(outIdx + 1) = CByte(&H80 Or ((codePoint \ &H1000) And &H3F))
            bytes(outIdx + 2) = CByte(&H80 Or ((codePoint \ &H40) And &H3F))
            bytes(outIdx + 3) = CByte(&H80 Or (codePoint And &H3F))
            outIdx = outIdx + 4
        End If
    Loop
   
    ' Trim to actual size
    If outIdx = 0 Then
        ReDim bytes(0 To 0)
        StrToUTF8Bytes = bytes
        Exit Function
    End If
   
    ReDim Preserve bytes(0 To outIdx - 1)
    StrToUTF8Bytes = bytes
End Function

' ============================================================
' Settings reader - used everywhere
' ============================================================
Public Function GetSetting(ByVal settingName As String) As String
    Dim ws As Worksheet, LastRow As Long, i As Long
    Set ws = ThisWorkbook.Sheets("Settings")
    LastRow = ws.Cells(ws.rows.Count, "B").End(xlUp).Row
    For i = 4 To LastRow
        If LCase(Trim(ws.Cells(i, 2).Value)) = LCase(Trim(settingName)) Then
            GetSetting = CStr(ws.Cells(i, 3).Value)
            Exit Function
        End If
    Next i
    GetSetting = ""
End Function



