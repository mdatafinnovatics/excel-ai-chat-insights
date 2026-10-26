Attribute VB_Name = "AIExtract_Mod3"

Option Explicit
' ============================================================
' MODULE: modJSON (V3 - rewritten with state machine parser)
' Robust JSON array parser for Claude API responses
' ============================================================

Public Function ParseClaudeResponse(ByVal ResponseText As String) As Variant
    Dim contentText As String
    contentText = ExtractContentFromResponse(ResponseText)
    
    If Len(contentText) = 0 Then
        ParseClaudeResponse = Array()
        Exit Function
    End If
    
    Dim cleaned As String
    cleaned = CleanJsonString(contentText)
    
    ParseClaudeResponse = ParseJsonArray(cleaned)
End Function

Private Function ExtractContentFromResponse(ByVal s As String) As String
    ' Find "text":"..." in API response and extract the value
    Dim marker As String
    marker = """text"":"""
    Dim startIdx As Long
    startIdx = InStr(1, s, marker)
    If startIdx = 0 Then
        ExtractContentFromResponse = ""
        Exit Function
    End If
    startIdx = startIdx + Len(marker)
    
    ' Walk forward, respecting backslash escapes
    Dim i As Long, Result As String, Ch As String
    Dim n As Long
    n = Len(s)
    i = startIdx
    Do While i <= n
        Ch = Mid(s, i, 1)
        If Ch = "\" And i < n Then
            ' Escape sequence
            Dim nextCh As String
            nextCh = Mid(s, i + 1, 1)
            Select Case nextCh
                Case """": Result = Result & """"
                Case "\": Result = Result & "\"
                Case "n": Result = Result & vbLf
                Case "r": Result = Result & vbCr
                Case "t": Result = Result & vbTab
                Case "/": Result = Result & "/"
                Case "b": Result = Result & Chr(8)
                Case "f": Result = Result & Chr(12)
                Case "u"
                    ' Unicode escape \uXXXX
                    If i + 5 <= n Then
                        Dim hexStr As String
                        hexStr = Mid(s, i + 2, 4)
                        Result = Result & ChrW(CLng("&H" & hexStr))
                        i = i + 4
                    End If
                Case Else: Result = Result & nextCh
            End Select
            i = i + 2
        ElseIf Ch = """" Then
            ' End of string value
            Exit Do
        Else
            Result = Result & Ch
            i = i + 1
        End If
    Loop
    
    ExtractContentFromResponse = Result
End Function

Private Function CleanJsonString(ByVal s As String) As String
    Dim t As String
    t = TrimAll(s)
    
    ' Strip markdown code fences
    If Left(t, 7) = "```json" Then
        t = Mid(t, 8)
        t = TrimAll(t)
    ElseIf Left(t, 3) = "```" Then
        t = Mid(t, 4)
        t = TrimAll(t)
    End If
    
    If Right(t, 3) = "```" Then
        t = Left(t, Len(t) - 3)
        t = TrimAll(t)
    End If
    
    CleanJsonString = t
End Function

Private Function TrimAll(ByVal s As String) As String
    ' Trim spaces, tabs, newlines from both ends
    Dim t As String
    t = s
    Do While Len(t) > 0
        Dim fc As String
        fc = Left(t, 1)
        If fc = " " Or fc = vbTab Or fc = vbLf Or fc = vbCr Then
            t = Mid(t, 2)
        Else
            Exit Do
        End If
    Loop
    Do While Len(t) > 0
        Dim lc As String
        lc = Right(t, 1)
        If lc = " " Or lc = vbTab Or lc = vbLf Or lc = vbCr Then
            t = Left(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    TrimAll = t
End Function

Private Function ParseJsonArray(ByVal s As String) As Variant
    ' State-machine parser: walks through array, finds each {...} object, parses fields
    Dim objects As Collection
    Set objects = New Collection
    
    Dim n As Long
    n = Len(s)
    
    Dim i As Long
    i = 1
    
    ' Find opening [
    Do While i <= n
        If Mid(s, i, 1) = "[" Then
            i = i + 1
            Exit Do
        End If
        i = i + 1
    Loop
    
    ' Parse objects
    Do While i <= n
        Dim Ch As String
        Ch = Mid(s, i, 1)
        
        If Ch = "{" Then
            ' Parse this object
            Dim objEnd As Long
            objEnd = FindMatchingBrace(s, i)
            If objEnd > 0 Then
                Dim objStr As String
                objStr = Mid(s, i + 1, objEnd - i - 1)
                
                Dim fields(1 To 9) As String
                ParseObjectFields objStr, fields
                
                objects.Add fields
                i = objEnd + 1
            Else
                Exit Do
            End If
        ElseIf Ch = "]" Then
            Exit Do
        Else
            i = i + 1
        End If
    Loop
    
    ' Convert to 2D array
    If objects.Count = 0 Then
        ParseJsonArray = Array()
        Exit Function
    End If
    
    Dim arr() As Variant
    ReDim arr(1 To objects.Count, 1 To 9)
    Dim rNum As Long, k As Long
    For rNum = 1 To objects.Count
        Dim f As Variant
        f = objects(rNum)
        For k = 1 To 9
            arr(rNum, k) = f(k)
        Next k
    Next rNum
    
    ParseJsonArray = arr
End Function

Private Function FindMatchingBrace(ByVal s As String, ByVal StartPos As Long) As Long
    ' Given position of {, find position of matching }
    Dim depth As Long, i As Long, inString As Boolean, prevWasBackslash As Boolean
    depth = 0
    i = StartPos
    Dim n As Long
    n = Len(s)
    
    Do While i <= n
        Dim Ch As String
        Ch = Mid(s, i, 1)
        
        If inString Then
            If Ch = "\" And Not prevWasBackslash Then
                prevWasBackslash = True
            ElseIf Ch = """" And Not prevWasBackslash Then
                inString = False
                prevWasBackslash = False
            Else
                prevWasBackslash = False
            End If
        Else
            If Ch = """" Then
                inString = True
                prevWasBackslash = False
            ElseIf Ch = "{" Then
                depth = depth + 1
            ElseIf Ch = "}" Then
                depth = depth - 1
                If depth = 0 Then
                    FindMatchingBrace = i
                    Exit Function
                End If
            End If
        End If
        i = i + 1
    Loop
    
    FindMatchingBrace = 0
End Function

Private Sub ParseObjectFields(ByVal objStr As String, ByRef fields() As String)
    ' Fields order: 1=time, 2=trailer, 3=address, 4=city, 5=leave_or_arrive, 6=notes, 7=rtn_new_picked, 8=del_time, 9=time_inferred
    fields(1) = ExtractField(objStr, "time")
    fields(2) = ExtractField(objStr, "trailer")
    fields(3) = ExtractField(objStr, "address")
    fields(4) = ExtractField(objStr, "city")
    fields(5) = ExtractField(objStr, "leave_or_arrive")
    fields(6) = ExtractField(objStr, "notes")
    fields(7) = ExtractField(objStr, "rtn_new_picked")
    fields(8) = ExtractField(objStr, "del_time")
    fields(9) = ExtractField(objStr, "time_inferred")
End Sub

Private Function ExtractField(ByVal objStr As String, ByVal fieldName As String) As String
    ' Find "fieldName": ... and return the value (string or unquoted)
    Dim searchKey As String
    searchKey = """" & fieldName & """"
    
    Dim keyPos As Long
    keyPos = InStr(1, objStr, searchKey, vbBinaryCompare)
    If keyPos = 0 Then
        ExtractField = ""
        Exit Function
    End If
    
    ' Advance past the key and the colon
    Dim i As Long
    i = keyPos + Len(searchKey)
    Dim n As Long
    n = Len(objStr)
    
    ' Skip whitespace and colon
    Do While i <= n
        Dim Ch As String
        Ch = Mid(objStr, i, 1)
        If Ch = ":" Or Ch = " " Or Ch = vbTab Or Ch = vbLf Or Ch = vbCr Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    
    If i > n Then
        ExtractField = ""
        Exit Function
    End If
    
    Dim firstCh As String
    firstCh = Mid(objStr, i, 1)
    
    If firstCh = """" Then
        ' Quoted string value - read until unescaped closing quote
        Dim Result As String, prevBackslash As Boolean
        i = i + 1
        Do While i <= n
            Dim c As String
            c = Mid(objStr, i, 1)
            If c = "\" And Not prevBackslash Then
                prevBackslash = True
                i = i + 1
            ElseIf c = """" And Not prevBackslash Then
                ExtractField = Result
                Exit Function
            Else
                If prevBackslash Then
                    Select Case c
                        Case "n": Result = Result & vbLf
                        Case "r": Result = Result & vbCr
                        Case "t": Result = Result & vbTab
                        Case """": Result = Result & """"
                        Case "\": Result = Result & "\"
                        Case "/": Result = Result & "/"
                        Case Else: Result = Result & c
                    End Select
                    prevBackslash = False
                Else
                    Result = Result & c
                End If
                i = i + 1
            End If
        Loop
        ExtractField = Result
    Else
        ' Unquoted value (true, false, null, number) - read until comma, }, or whitespace
        Dim unq As String
        Do While i <= n
            Dim uc As String
            uc = Mid(objStr, i, 1)
            If uc = "," Or uc = "}" Or uc = vbLf Or uc = vbCr Or uc = vbTab Then
                Exit Do
            End If
            unq = unq & uc
            i = i + 1
        Loop
        ExtractField = Trim(unq)
    End If
End Function



