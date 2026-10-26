Attribute VB_Name = "AIExtract_Mod4"

Option Explicit
' ============================================================
' MODULE: modExtraction
' PURPOSE: Main orchestrator _ read chat, call API, write preview
' Connect this to the green "EXTRACT CHAT" button on AI Extract sheet.
' ============================================================

Public Sub AIPreview_Automation()
    Dim ws As Worksheet, chatText As String, driverName As String, mode As String
    Dim prompt As String, Response As String, t0 As Single
   
   
    Set ws = ThisWorkbook.Sheets("AI Extract")
   
    ws.Range("J5:AD1000").Value = ""
    ws.Range("J5:AD1000").Font.Italic = False
    ws.Range("J5:AD1000").Interior.Color = xlNone
   
    ' --- Validate inputs ---
    chatText = ReadChatFromPasteArea(ws)
    driverName = CStr(ws.Range("E4").Value)
    mode = CStr(ws.Range("E5").Value)
   
    If GetSetting("Warn if paste box is empty") = "YES" And Trim(chatText) = "" Then
        MsgBox "Please paste a chat into columns A and B before clicking Extract.", vbExclamation, "Empty Chat"
        Exit Sub
    End If
   
    If GetSetting("Warn if driver name not selected") = "YES" And Trim(driverName) = "" Then
        If MsgBox("No driver name selected. Use default '" & GetSetting("Default staff name (if dropdown empty)") & "'?", _
                  vbYesNo + vbQuestion, "Driver Name") = vbNo Then Exit Sub
        driverName = GetSetting("Default staff name (if dropdown empty)")
    End If
   
    If mode = "" Then mode = GetSetting("Default Start/Finish mode")
   
    ' --- Update status ---
    SetStatus ws, "Building prompt...", "FFE699"
    DoEvents
   
    prompt = BuildPrompt(chatText, driverName, mode)
   
    SetStatus ws, "Calling Claude API... (this can take 10-30 seconds)", "FFE699"
    DoEvents
   
    t0 = Timer
    Response = CallClaudeAPI(prompt)
'    ThisWorkbook.Sheets("AI Extract").Range("AF1").Value = response
   
    If Left(Response, 5) = "ERROR" Then
        SetStatus ws, "FAILED", "FFC7CE"
        MsgBox Response, vbCritical, "Extraction Failed"
'        LogExtraction driverName, mode, 0, 0, "FAILED", response
        Exit Sub
    End If
   
    SetStatus ws, "Parsing response...", "FFE699"
    DoEvents
   

    Dim rows As Variant
    rows = ParseClaudeResponse(Response)

'' === DIAGNOSTIC V2 ===
'    Dim diagWs As Worksheet
'    Set diagWs = ThisWorkbook.Sheets("AI Extract")
'    diagWs.Range("AF1").Value = response
'
'    Dim rows As Variant
'    rows = ParseClaudeResponse(response)
'
'    diagWs.Range("AF7").Value = "ROWS_COUNT=" & SafeRowCountDiag(rows)
'
'    ' Test direct write to cells - bypass WritePreview entirely
'    On Error Resume Next
'    diagWs.Range("AF10").Value = "TEST_ROW1_TIME=" & rows(1, 1)
'    diagWs.Range("AF11").Value = "TEST_ROW1_TRAILER=" & rows(1, 2)
'    diagWs.Range("AF12").Value = "TEST_ROW1_ADDR=" & rows(1, 3)
'    diagWs.Range("AF13").Value = "TEST_ROW1_CITY=" & rows(1, 4)
'    diagWs.Range("AF14").Value = "TEST_ROW1_LVAR=" & rows(1, 5)
'    diagWs.Range("AF15").Value = "TEST_ROW1_NOTES=" & Left(CStr(rows(1, 6)), 50)
'    diagWs.Range("AF16").Value = "TEST_ROW9_TIME=" & rows(9, 1)
'    diagWs.Range("AF17").Value = "TEST_ROW9_TRAILER=" & rows(9, 2)
'    diagWs.Range("AF18").Value = "TEST_ROW9_LVAR=" & rows(9, 5)
'    diagWs.Range("AF19").Value = "ERR_AFTER_TEST=" & Err.Number & " - " & Err.Description
'    On Error GoTo 0
'    ' === END DIAGNOSTIC V2 ===
   
    Dim rowCount As Long
    rowCount = SafeRowCount(rows)
   
    If rowCount = 0 Then
        SetStatus ws, "No events extracted", "FFC7CE"
        MsgBox "AI returned 0 event rows. Check that the chat has actual driver events.", vbInformation, "No Events"
'        LogExtraction driverName, mode, 0, EstimateTokens(prompt, response), "EMPTY", "No events parsed"
        Exit Sub
    End If
   
' --- Write preview to output area (cols J..AD) - inline write ---
    On Error GoTo 0
   
    ' Clear previous preview first
    ws.Range("J5:AD200").ClearContents
   
    Dim writeRowCount As Long
    writeRowCount = 0
    On Error Resume Next
    writeRowCount = UBound(rows, 1)
    On Error GoTo 0
   
    Dim todayStr As String
    todayStr = Format(Date, "mm/dd/yyyy")
    Dim upperCity As String, upperTrailer As String, markInferred As String
    upperCity = GetSetting("Auto-uppercase cities")
    upperTrailer = GetSetting("Auto-uppercase trailer numbers")
    markInferred = GetSetting("Mark inferred times visually")
   
    Dim wIdx As Long
    For wIdx = 1 To writeRowCount
        Dim targetRow As Long
        targetRow = 4 + wIdx
       
        Dim vTime As String, vTrailer As String, vAddr As String, vCity As String
        Dim vLvAr As String, vNotes As String, vRnp As String, vDelTime As String, vInf As String
        vTime = CStr(rows(wIdx, 1))
        vTrailer = CStr(rows(wIdx, 2))
        vAddr = CStr(rows(wIdx, 3))
        vCity = CStr(rows(wIdx, 4))
        vLvAr = CStr(rows(wIdx, 5))
        vNotes = CStr(rows(wIdx, 6))
        vRnp = CStr(rows(wIdx, 7))
        vDelTime = CStr(rows(wIdx, 8))
        vInf = LCase(CStr(rows(wIdx, 9)))
       
        If upperCity = "YES" And Len(vCity) > 0 Then vCity = UCase(vCity)
        If upperTrailer = "YES" And Len(vTrailer) > 0 Then vTrailer = UCase(vTrailer)
       
        ws.Cells(targetRow, 10).Value = todayStr        ' J = date
        ws.Cells(targetRow, 11).Value = vTime           ' K = time
        ws.Cells(targetRow, 12).Value = driverName      ' L = staff
        ws.Cells(targetRow, 13).Value = vTrailer        ' M = trailer
        ws.Cells(targetRow, 14).Value = vAddr           ' N = address
        ws.Cells(targetRow, 15).Value = vCity           ' O = city
        ws.Cells(targetRow, 17).Value = vLvAr           ' Q = leave/arrive
        ws.Cells(targetRow, 18).Value = vNotes          ' R = notes
        ws.Cells(targetRow, 21).Value = vRnp            ' U = rtn/new/picked
        ws.Cells(targetRow, 26).Value = vDelTime        ' Z = del time
       
        ' Highlight inferred time
        If markInferred = "YES" And vInf = "true" Then
            ws.Cells(targetRow, 11).Interior.Color = RGB(255, 230, 153)
            ws.Cells(targetRow, 11).Font.Italic = True
            ws.Cells(targetRow, 11).Font.Color = RGB(156, 87, 0)
        End If
    Next wIdx
   
'    Dim rowCount As Long
    rowCount = writeRowCount
   
    ' --- Update status indicators ---
    SetStatus ws, "Extracted " & rowCount & " rows", "E2EFDA"
'    ws.Range("E16").Value = Format(Now, "mm/dd/yyyy hh:mm:ss")
'    ws.Range("E17").Value = rowCount
'    ws.Range("E18").Value = EstimateTokens(prompt, response)
   
'    LogExtraction driverName, mode, rowCount, EstimateTokens(prompt, response), "SUCCESS", _
                  "Took " & Format(Timer - t0, "0.0") & "s"


    Dim Last_Rw As Variant
    Last_Rw = ws.Range("J1000").End(xlUp).Row
       
    If Last_Rw > 4 Then
        Range("P5").Value = "Start"
        Range("P" & Last_Rw).Value = "Finish"
        If Range("Q" & Last_Rw).Value = "Finish" Then Range("Q" & Last_Rw).Value = "leave"
   
        Range("S5:S" & Last_Rw).Formula = "=IFERROR(IF(N5=""yard"","""",VLOOKUP(N5,'driver main sheet'!$E:$J,6,0)),"""")"
        Range("T5:T" & Last_Rw).Formula = "=IFERROR(IF(N5=""yard"","""",VLOOKUP(N5,'driver main sheet'!$E:$K,7,0)),"""")"
   
        Range("S5:T" & Last_Rw).Select
        Selection.Copy
        Selection.PasteSpecial xlPasteValues
       
        Application.CutCopyMode = False
       
        Range("J4").Select
   
    End If
   
    MsgBox "Data Extracted..", vbInformation, "Status"
   

End Sub

'Public Sub btnConfirm_Click()
'    Dim ws As Worksheet, mainWs As Worksheet
'    Set ws = ThisWorkbook.Sheets("AI Extract")
'
'    Dim sheetName As String
'    sheetName = GetSetting("Write to which sheet on Confirm")
'
'    On Error Resume Next
'    Set mainWs = ThisWorkbook.Sheets(sheetName)
'    On Error GoTo 0
'
'    If mainWs Is Nothing Then
'        MsgBox "Could not find sheet: " & sheetName, vbCritical, "Confirm Error"
'        Exit Sub
'    End If
'
'    ' Find last filled row in AI Extract output area (column J)
'    Dim lastPreviewRow As Long, firstPreviewRow As Long
'    firstPreviewRow = 5
'    lastPreviewRow = ws.Cells(ws.rows.Count, "J").End(xlUp).Row
'    If lastPreviewRow < firstPreviewRow Then
'        MsgBox "No preview rows to confirm.", vbExclamation, "Nothing To Confirm"
'        Exit Sub
'    End If
'
'    ' Find last filled row in main sheet (column A)
'    Dim mainLastRow As Long
'    mainLastRow = mainWs.Cells(mainWs.rows.Count, "A").End(xlUp).Row
'
'    Dim rowsToCopy As Long
'    rowsToCopy = lastPreviewRow - firstPreviewRow + 1
'
'    If MsgBox("Copy " & rowsToCopy & " rows to sheet '" & sheetName & "' starting at row " & (mainLastRow + 1) & "?", _
'              vbYesNo + vbQuestion, "Confirm Write") = vbNo Then Exit Sub
'
'    ' Copy values from cols J..AD to A..U of main sheet
'    Dim previewRange As Range, mainTarget As Range
'    Set previewRange = ws.Range("J" & firstPreviewRow & ":AD" & lastPreviewRow)
'    Set mainTarget = mainWs.Range("A" & (mainLastRow + 1))
'
'    previewRange.Copy
'    mainTarget.PasteSpecial xlPasteValues
'    Application.CutCopyMode = False
'
'    SetStatus ws, "? " & rowsToCopy & " rows written to " & sheetName, "E2EFDA"
'    MsgBox "Done! " & rowsToCopy & " rows added to '" & sheetName & "'.", vbInformation, "Success"
'End Sub
'
'Public Sub btnClearOutput_Click()
'    Dim ws As Worksheet
'    Set ws = ThisWorkbook.Sheets("AI Extract")
'    ws.Range("J5:AD200").ClearContents
'
'    ' Reset formatting back to defaults
'    Dim r As Long, col As String
'    For r = 5 To 104
'        For Each col In Array("J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD")
'            With ws.Range(col & r)
'                If col = "S" Or col = "T" Or col = "W" Or col = "X" Then
'                    .Interior.Color = RGB(242, 242, 242)
'                ElseIf col = "V" Or col = "Y" Or col = "AA" Or col = "AB" Or col = "AC" Or col = "AD" Then
'                    .Interior.Color = RGB(255, 255, 255)
'                Else
'                    .Interior.Color = RGB(255, 242, 204)
'                End If
'                .Font.Italic = False
'                .Font.Color = RGB(0, 0, 0)
'            End With
'        Next col
'    Next r
'
'    SetStatus ws, "Output cleared", "DDEBF7"
'    ws.Range("E17").Value = 0
'End Sub

Private Function ReadChatFromPasteArea(ByRef ws As Worksheet) As String
    Dim lastA As Long, lastB As Long, LastRow As Long, i As Long, Result As String
    lastA = ws.Cells(ws.rows.Count, "A").End(xlUp).Row
    lastB = ws.Cells(ws.rows.Count, "B").End(xlUp).Row
    LastRow = Application.Max(lastA, lastB)
   
    For i = 6 To LastRow  ' chat starts at row 6
        Dim a As String, b As String
        a = CStr(ws.Cells(i, 1).Value)
        b = CStr(ws.Cells(i, 2).Value)
        If Trim(a) <> "" Then Result = Result & a & vbLf
        If Trim(b) <> "" Then Result = Result & b & vbLf
    Next i
   
    ReadChatFromPasteArea = Result
End Function

Private Sub SetStatus(ByRef ws As Worksheet, ByVal Msg As String, ByVal colorHex As String)
    ws.Range("A1").Value = Msg
    ws.Range("A1").Interior.Color = HexToColor(colorHex)
End Sub

Private Function HexToColor(ByVal hex As String) As Long
    HexToColor = RGB(CLng("&H" & Mid(hex, 1, 2)), CLng("&H" & Mid(hex, 3, 2)), CLng("&H" & Mid(hex, 5, 2)))
End Function

Private Function SafeRowCount(ByVal arr As Variant) As Long
    On Error Resume Next
    SafeRowCount = UBound(arr, 1)
    If Err.Number <> 0 Then SafeRowCount = 0
    On Error GoTo 0
End Function

Private Function EstimateTokens(ByVal prompt As String, ByVal Response As String) As Long
    ' Rough estimate: ~4 chars per token
    EstimateTokens = (Len(prompt) + Len(Response)) \ 4
End Function

'Private Sub LogExtraction(ByVal driver As String, ByVal mode As String, _
'                          ByVal rowCount As Long, ByVal tokens As Long, _
'                          ByVal status As String, ByVal notes As String)
'    If GetSetting("Log every extraction") <> "YES" Then Exit Sub
'
'    Dim logWs As Worksheet, nextRow As Long, lastNum As Long
'    Set logWs = ThisWorkbook.Sheets("Extraction Log")
'    nextRow = logWs.Cells(logWs.rows.Count, "B").End(xlUp).Row + 1
'    If nextRow < 4 Then nextRow = 4
'
'    ' If row 4 still has the example placeholder, overwrite it
'    If logWs.Range("I4").Value Like "*Sample log*" Then nextRow = 4
'
'    lastNum = logWs.Cells(nextRow + 1, "A").Value
'    If Not IsNumeric(lastNum) Then lastNum = 0
'
'    logWs.Cells(nextRow, "A").Value = lastNum + 1
'    logWs.Cells(nextRow, "B").Value = Format(Date, "mm/dd/yyyy")
'    logWs.Cells(nextRow, "C").Value = Format(Now, "hh:mm:ss")
'    logWs.Cells(nextRow, "D").Value = driver
'    logWs.Cells(nextRow, "E").Value = mode
'    logWs.Cells(nextRow, "F").Value = rowCount
'    logWs.Cells(nextRow, "G").Value = tokens
'    logWs.Cells(nextRow, "H").Value = status
'    logWs.Cells(nextRow, "I").Value = notes
'End Sub

'Public Function ExtractContentFromResponseDiag(ByVal s As String) As String
'    Dim startIdx As Long
'    Dim marker As String
'    marker = """text"":"""
'    startIdx = InStr(1, s, marker)
'    If startIdx = 0 Then
'        ExtractContentFromResponseDiag = "ERR: no text marker found"
'        Exit Function
'    End If
'    startIdx = startIdx + Len(marker)
'
'    Dim i As Long, result As String, ch As String, prevCh As String
'    For i = startIdx To Len(s)
'        ch = Mid(s, i, 1)
'        If ch = """" And prevCh <> "\" Then
'            Exit For
'        End If
'        result = result & ch
'        prevCh = ch
'    Next i
'
'    result = Replace(result, "\n", vbLf)
'    result = Replace(result, "\r", vbCr)
'    result = Replace(result, "\t", vbTab)
'    result = Replace(result, "\""", """")
'    result = Replace(result, "\\", "\")
'
'    ExtractContentFromResponseDiag = result
'End Function
'
'Public Function CleanJsonStringDiag(ByVal s As String) As String
'    Dim t As String
'    t = Trim(s)
'    If Left(t, 7) = "```json" Then t = Mid(t, 8)
'    If Left(t, 3) = "```" Then t = Mid(t, 4)
'    If Right(t, 3) = "```" Then t = Left(t, Len(t) - 3)
'    CleanJsonStringDiag = Trim(t)
'End Function
'
'Public Function SafeRowCountDiag(ByVal arr As Variant) As Long
'    On Error Resume Next
'    SafeRowCountDiag = UBound(arr, 1)
'    If Err.Number <> 0 Then SafeRowCountDiag = 0
'    On Error GoTo 0
'End Function



