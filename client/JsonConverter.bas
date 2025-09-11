Option Explicit
Option Compare Binary
' =============================================================================
' VBA-JSON (modernized 64-bit only) ? based on Tim Hall's v2.3.1
' Drop-in module for Office 64-bit (VBA7+). No 32-bit code paths.
' =============================================================================

' ===== 64-bit platform declares =====
#If Mac Then
    ' ---- macOS (64-bit) ----
    Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _
        (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
    Private Declare PtrSafe Function utc_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" _
        (ByVal utc_File As LongPtr) As LongPtr
    Private Declare PtrSafe Function utc_fread Lib "/usr/lib/libc.dylib" Alias "fread" _
        (ByVal utc_Buffer As String, ByVal utc_Size As LongPtr, ByVal utc_Number As LongPtr, ByVal utc_File As LongPtr) As LongPtr
    Private Declare PtrSafe Function utc_feof Lib "/usr/lib/libc.dylib" Alias "feof" _
        (ByVal utc_File As LongPtr) As LongPtr

    Private Type utc_ShellResult
        utc_Output As String
        utc_ExitCode As LongPtr
    End Type
#Else
    ' ---- Windows (64-bit) ----
    Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
        (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
    Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
        (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
    Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
        (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

    Private Type utc_SYSTEMTIME
        utc_wYear As Integer
        utc_wMonth As Integer
        utc_wDayOfWeek As Integer
        utc_wDay As Integer
        utc_wHour As Integer
        utc_wMinute As Integer
        utc_wSecond As Integer
        utc_wMilliseconds As Integer
    End Type

    Private Type utc_TIME_ZONE_INFORMATION
        utc_Bias As Long
        utc_StandardName(0 To 31) As Integer
        utc_StandardDate As utc_SYSTEMTIME
        utc_StandardBias As Long
        utc_DaylightName(0 To 31) As Integer
        utc_DaylightDate As utc_SYSTEMTIME
        utc_DaylightBias As Long
    End Type
#End If
' ===== end declares =====

' ===== Options =====
Private Type json_Options
    UseDoubleForLargeNumbers As Boolean
    AllowUnquotedKeys As Boolean
    EscapeSolidus As Boolean
End Type
Public JsonOptions As json_Options

' ===== Public API =====

Public Function ParseJson(ByVal JsonString As String) As Object
    Dim i As Long: i = 1
    ' strip common whitespace (keep NBSP etc.)
    JsonString = Replace(Replace(Replace(JsonString, vbCr, ""), vbLf, ""), vbTab, "")
    json_SkipSpaces JsonString, i

    Select Case Mid$(JsonString, i, 1)
        Case "{": Set ParseJson = json_ParseObject(JsonString, i)
        Case "[": Set ParseJson = json_ParseArray(JsonString, i)
        Case Else
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(JsonString, i, "Expecting '{' or '['")
    End Select
End Function

Public Function ConvertToJson(ByVal JsonValue As Variant, _
                              Optional ByVal Whitespace As Variant, _
                              Optional ByVal lvl As Long = 0) As String
    Dim buf As String, pos As Long, blen As Long
    Dim i As Long, l1 As Long, u1 As Long, l2 As Long, u2 As Long
    Dim first As Boolean, first2 As Boolean
    Dim k As Variant, v As Variant, s As String
    Dim pretty As Boolean, ind As String, ind2 As String

    pretty = Not IsMissing(Whitespace)
    Select Case VarType(JsonValue)
        Case vbNull: ConvertToJson = "null": Exit Function
        Case vbDate: ConvertToJson = """" & ConvertToIso(CDate(JsonValue)) & """": Exit Function
        Case vbString
            If Not JsonOptions.UseDoubleForLargeNumbers And json_StringIsLargeNumber(JsonValue) Then
                ConvertToJson = JsonValue
            Else
                ConvertToJson = """" & json_Encode(JsonValue) & """"
            End If
            Exit Function
        Case vbBoolean: ConvertToJson = IIf(JsonValue, "true", "false"): Exit Function
        Case vbArray To vbArray + vbByte
            If pretty Then
                If VarType(Whitespace) = vbString Then
                    ind = String$(lvl + 1, Whitespace): ind2 = String$(lvl + 2, Whitespace)
                Else
                    ind = Space$((lvl + 1) * Whitespace): ind2 = Space$((lvl + 2) * Whitespace)
                End If
            End If
            json_BufferAppend buf, "[", pos, blen
            On Error Resume Next
            l1 = LBound(JsonValue, 1): u1 = UBound(JsonValue, 1)
            l2 = LBound(JsonValue, 2): u2 = UBound(JsonValue, 2)
            On Error GoTo 0
            first = True: first2 = True
            If l1 >= 0 And u1 >= 0 Then
                For i = l1 To u1
                    If Not first Then json_BufferAppend buf, ",", pos, blen Else first = False
                    If l2 >= 0 And u2 >= 0 Then
                        If pretty Then json_BufferAppend buf, vbNewLine, pos, blen
                        json_BufferAppend buf, ind & "[", pos, blen
                        Dim j As Long
                        For j = l2 To u2
                            If Not first2 Then json_BufferAppend buf, ",", pos, blen Else first2 = False
                            s = ConvertToJson(JsonValue(i, j), Whitespace, lvl + 2)
                            If s = "" And json_IsUndefined(JsonValue(i, j)) Then s = "null"
                            If pretty Then s = vbNewLine & ind2 & s
                            json_BufferAppend buf, s, pos, blen
                        Next j
                        If pretty Then json_BufferAppend buf, vbNewLine, pos, blen
                        json_BufferAppend buf, ind & "]", pos, blen
                        first2 = True
                    Else
                        s = ConvertToJson(JsonValue(i), Whitespace, lvl + 1)
                        If s = "" And json_IsUndefined(JsonValue(i)) Then s = "null"
                        If pretty Then s = vbNewLine & ind & s
                        json_BufferAppend buf, s, pos, blen
                    End If
                Next i
            End If
            If pretty Then
                json_BufferAppend buf, vbNewLine, pos, blen
                If VarType(Whitespace) = vbString Then ind = String$(lvl, Whitespace) Else ind = Space$(lvl * Whitespace)
            End If
            json_BufferAppend buf, ind & "]", pos, blen
            ConvertToJson = json_BufferToString(buf, pos)
            Exit Function

        Case vbObject
            If pretty Then
                If VarType(Whitespace) = vbString Then ind = String$(lvl + 1, Whitespace) Else ind = Space$((lvl + 1) * Whitespace)
            End If

            If TypeName(JsonValue) = "Dictionary" Or TypeName(JsonValue) = "Scripting.Dictionary" Then
                json_BufferAppend buf, "{", pos, blen
                first = True
                For Each k In JsonValue.Keys
                    s = ConvertToJson(JsonValue(k), Whitespace, lvl + 1)
                    If s = "" Then
                        If json_IsUndefined(JsonValue(k)) Then GoTo NextKey
                    End If
                    If Not first Then json_BufferAppend buf, ",", pos, blen Else first = False
                    If pretty Then
                        s = vbNewLine & ind & """" & k & """: " & s
                    Else
                        s = """" & k & """:" & s
                    End If
                    json_BufferAppend buf, s, pos, blen
NextKey:
                Next k
                If pretty Then
                    json_BufferAppend buf, vbNewLine, pos, blen
                    If VarType(Whitespace) = vbString Then ind = String$(lvl, Whitespace) Else ind = Space$(lvl * Whitespace)
                End If
                json_BufferAppend buf, ind & "}", pos, blen
                ConvertToJson = json_BufferToString(buf, pos)
                Exit Function

            ElseIf TypeName(JsonValue) = "Collection" Then
                json_BufferAppend buf, "[", pos, blen
                first = True
                For Each v In JsonValue
                    If Not first Then json_BufferAppend buf, ",", pos, blen Else first = False
                    s = ConvertToJson(v, Whitespace, lvl + 1)
                    If s = "" And json_IsUndefined(v) Then s = "null"
                    If pretty Then s = vbNewLine & ind & s
                    json_BufferAppend buf, s, pos, blen
                Next v
                If pretty Then
                    json_BufferAppend buf, vbNewLine, pos, blen
                    If VarType(Whitespace) = vbString Then ind = String$(lvl, Whitespace) Else ind = Space$(lvl * Whitespace)
                End If
                json_BufferAppend buf, ind & "]", pos, blen
                ConvertToJson = json_BufferToString(buf, pos)
                Exit Function
            End If

        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            ConvertToJson = Replace(CStr(JsonValue), ",", "."): Exit Function

        Case Else
            On Error Resume Next
            ConvertToJson = JsonValue
            On Error GoTo 0
            Exit Function
    End Select
End Function

' ===== Private: Parsing =====

Private Function json_ParseObject(ByVal s As String, ByRef i As Long) As Object
    Dim key As String, nextc As String
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Set json_ParseObject = dict

    json_SkipSpaces s, i
    If Mid$(s, i, 1) <> "{" Then Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(s, i, "Expecting '{'") Else i = i + 1
    Do
        json_SkipSpaces s, i
        Select Case Mid$(s, i, 1)
            Case "}": i = i + 1: Exit Function
            Case ",": i = i + 1: json_SkipSpaces s, i
        End Select
        key = json_ParseKey(s, i)
        nextc = json_Peek(s, i)
        If nextc = "[" Or nextc = "{" Then
            Set dict.Item(key) = json_ParseValue(s, i)
        Else
            dict.Item(key) = json_ParseValue(s, i)
        End If
    Loop
End Function

Private Function json_ParseArray(ByVal s As String, ByRef i As Long) As Collection
    Dim col As New Collection
    Set json_ParseArray = col

    json_SkipSpaces s, i
    If Mid$(s, i, 1) <> "[" Then Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(s, i, "Expecting '['") Else i = i + 1
    Do
        json_SkipSpaces s, i
        Select Case Mid$(s, i, 1)
            Case "]": i = i + 1: Exit Function
            Case ",": i = i + 1: json_SkipSpaces s, i
        End Select
        col.add json_ParseValue(s, i)
    Loop
End Function

Private Function json_ParseValue(ByVal s As String, ByRef i As Long) As Variant
    json_SkipSpaces s, i
    Select Case Mid$(s, i, 1)
        Case "{": Set json_ParseValue = json_ParseObject(s, i)
        Case "[": Set json_ParseValue = json_ParseArray(s, i)
        Case """", "'": json_ParseValue = json_ParseString(s, i)
        Case Else
            If Mid$(s, i, 4) = "true" Then json_ParseValue = True: i = i + 4: Exit Function
            If Mid$(s, i, 5) = "false" Then json_ParseValue = False: i = i + 5: Exit Function
            If Mid$(s, i, 4) = "null" Then json_ParseValue = Null: i = i + 4: Exit Function
            If InStr("+-0123456789", Mid$(s, i, 1)) > 0 Then
                json_ParseValue = json_ParseNumber(s, i): Exit Function
            End If
            Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(s, i, "Expecting 'STRING', 'NUMBER', null, true, false, '{', or '['")
    End Select
End Function

Private Function json_ParseString(ByVal s As String, ByRef i As Long) As String
    Dim q As String, ch As String, code As String
    Dim buf As String, pos As Long, blen As Long

    json_SkipSpaces s, i
    q = Mid$(s, i, 1): i = i + 1
    Do While i > 0 And i <= Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "\"
                i = i + 1: ch = Mid$(s, i, 1)
                Select Case ch
                    Case """", "\", "/", "'": json_BufferAppend buf, ch, pos, blen: i = i + 1
                    Case "b": json_BufferAppend buf, vbBack, pos, blen: i = i + 1
                    Case "f": json_BufferAppend buf, vbFormFeed, pos, blen: i = i + 1
                    Case "n": json_BufferAppend buf, vbCrLf, pos, blen: i = i + 1
                    Case "r": json_BufferAppend buf, vbCr, pos, blen: i = i + 1
                    Case "t": json_BufferAppend buf, vbTab, pos, blen: i = i + 1
                    Case "u"
                        i = i + 1: code = Mid$(s, i, 4)
                        json_BufferAppend buf, ChrW$(Val("&h" & code)), pos, blen
                        i = i + 4
                End Select
            Case q
                json_ParseString = json_BufferToString(buf, pos): i = i + 1: Exit Function
            Case Else
                json_BufferAppend buf, ch, pos, blen: i = i + 1
        End Select
    Loop
End Function

Private Function json_ParseNumber(ByVal s As String, ByRef i As Long) As Variant
    Dim ch As String, valS As String, isLarge As Boolean
    json_SkipSpaces s, i
    Do While i > 0 And i <= Len(s)
        ch = Mid$(s, i, 1)
        If InStr("+-0123456789.eE", ch) > 0 Then
            valS = valS & ch: i = i + 1
        Else
            isLarge = IIf(InStr(valS, "."), Len(valS) >= 17, Len(valS) >= 16)
            If Not JsonOptions.UseDoubleForLargeNumbers And isLarge Then
                json_ParseNumber = valS
            Else
                json_ParseNumber = Val(valS)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function json_ParseKey(ByVal s As String, ByRef i As Long) As String
    If Mid$(s, i, 1) = """" Or Mid$(s, i, 1) = "'" Then
        json_ParseKey = json_ParseString(s, i)
    ElseIf JsonOptions.AllowUnquotedKeys Then
        Dim ch As String
        Do While i > 0 And i <= Len(s)
            ch = Mid$(s, i, 1)
            If (ch <> " ") And (ch <> ":") Then
                json_ParseKey = json_ParseKey & ch: i = i + 1
            Else
                Exit Do
            End If
        Loop
    Else
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(s, i, "Expecting '""' or '''")
    End If

    json_SkipSpaces s, i
    If Mid$(s, i, 1) <> ":" Then
        Err.Raise 10001, "JSONConverter", json_ParseErrorMessage(s, i, "Expecting ':'")
    Else
        i = i + 1
    End If
End Function

Private Function json_IsUndefined(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbEmpty: json_IsUndefined = True
        Case vbObject
            Select Case TypeName(v)
                Case "Empty", "Nothing": json_IsUndefined = True
            End Select
    End Select
End Function

Private Function json_Encode(ByVal t As Variant) As String
    Dim i As Long, ch As String, code As Long
    Dim buf As String, pos As Long, blen As Long
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        code = AscW(ch): If code < 0 Then code = code + 65536
        Select Case code
            Case 34: ch = "\"""        ' "
            Case 92: ch = "\\"         ' \
            Case 47: If JsonOptions.EscapeSolidus Then ch = "\/"
            Case 8:  ch = "\b"
            Case 12: ch = "\f"
            Case 10: ch = "\n"
            Case 13: ch = "\r"
            Case 9:  ch = "\t"
            Case 0 To 31, 127 To 65535: ch = "\u" & Right$("0000" & Hex$(code), 4)
        End Select
        json_BufferAppend buf, ch, pos, blen
    Next
    json_Encode = json_BufferToString(buf, pos)
End Function

Private Function json_Peek(ByVal s As String, ByVal i As Long, Optional ByVal n As Long = 1) As String
    json_SkipSpaces s, i
    json_Peek = Mid$(s, i, n)
End Function

Private Sub json_SkipSpaces(ByVal s As String, ByRef i As Long)
    Do While i > 0 And i <= Len(s)
        Dim c As String: c = Mid$(s, i, 1)
        Dim a As Long: a = AscW(c)
        If c = " " Or c = vbTab Or c = vbCr Or c = vbLf Or a = 160 Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Function json_StringIsLargeNumber(ByVal s As Variant) As Boolean
    Dim n As Long, i As Long, cc As Integer
    n = Len(s)
    If n >= 16 And n <= 100 Then
        json_StringIsLargeNumber = True
        For i = 1 To n
            cc = Asc(Mid$(s, i, 1))
            Select Case cc
                Case 46, 48 To 57, 69, 101
                Case Else: json_StringIsLargeNumber = False: Exit Function
            End Select
        Next
    End If
End Function

Private Function json_ParseErrorMessage(ByVal s As String, ByVal i As Long, ByVal msg As String) As String
    Dim i0 As Long, i1 As Long
    i0 = i - 20: If i0 < 1 Then i0 = 1
    i1 = i + 20: If i1 > Len(s) Then i1 = Len(s)
    json_ParseErrorMessage = "Error parsing JSON:" & vbNewLine & _
                             Mid$(s, i0, i1 - i0 + 1) & vbNewLine & _
                             Space$(i - i0) & "^" & vbNewLine & msg
End Function

Private Sub json_BufferAppend(ByRef buf As String, ByRef app As Variant, ByRef pos As Long, ByRef blen As Long)
    Dim alen As Long, need As Long, add As Long
    alen = Len(app): need = pos + alen
    If need > blen Then
        add = IIf(alen > blen, alen, blen)
        If add < 1024 Then add = 1024
        buf = buf & Space$(add)
        blen = blen + add
    End If
    Mid$(buf, pos + 1, alen) = CStr(app)
    pos = pos + alen
End Sub

Private Function json_BufferToString(ByRef buf As String, ByVal pos As Long) As String
    If pos > 0 Then json_BufferToString = Left$(buf, pos)
End Function

' ===== UTC/ISO 8601 (64-bit only) =====

Public Function ParseUtc(ByVal utc_UtcDate As Date) As Date
#If Mac Then
    ParseUtc = utc_ConvertDate(utc_UtcDate)
#Else
    Dim tz As utc_TIME_ZONE_INFORMATION, loc As utc_SYSTEMTIME
    utc_GetTimeZoneInformation tz
    utc_SystemTimeToTzSpecificLocalTime tz, utc_DateToSystemTime(utc_UtcDate), loc
    ParseUtc = utc_SystemTimeToDate(loc)
#End If
End Function

Public Function ConvertToUtc(ByVal utc_LocalDate As Date) As Date
#If Mac Then
    ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else
    Dim tz As utc_TIME_ZONE_INFORMATION, ut As utc_SYSTEMTIME
    utc_GetTimeZoneInformation tz
    utc_TzSpecificLocalTimeToSystemTime tz, utc_DateToSystemTime(utc_LocalDate), ut
    ConvertToUtc = utc_SystemTimeToDate(ut)
#End If
End Function

Public Function ParseIso(ByVal utc_IsoString As String) As Date
    Dim p() As String, d() As String, t() As String
    Dim ofsIdx As Long, hasOfs As Boolean, negOfs As Boolean
    Dim o() As String, offs As Date

    p = Split(utc_IsoString, "T")
    d = Split(p(0), "-")
    ParseIso = DateSerial(CInt(d(0)), CInt(d(1)), CInt(d(2)))

    If UBound(p) > 0 Then
        If InStr(p(1), "Z") > 0 Then
            t = Split(Replace(p(1), "Z", ""), ":")
        Else
            ofsIdx = InStr(1, p(1), "+")
            If ofsIdx = 0 Then negOfs = True: ofsIdx = InStr(1, p(1), "-")
            If ofsIdx > 0 Then
                hasOfs = True
                t = Split(Left$(p(1), ofsIdx - 1), ":")
                o = Split(Right$(p(1), Len(p(1)) - ofsIdx), ":")
                Select Case UBound(o)
                    Case 0: offs = TimeSerial(CInt(o(0)), 0, 0)
                    Case 1: offs = TimeSerial(CInt(o(0)), CInt(o(1)), 0)
                    Case 2: offs = TimeSerial(CInt(o(0)), CInt(o(1)), Int(Val(o(2))))
                End Select
                If negOfs Then offs = -offs
            Else
                t = Split(p(1), ":")
            End If
        End If

        Select Case UBound(t)
            Case 0: ParseIso = ParseIso + TimeSerial(CInt(t(0)), 0, 0)
            Case 1: ParseIso = ParseIso + TimeSerial(CInt(t(0)), CInt(t(1)), 0)
            Case 2: ParseIso = ParseIso + TimeSerial(CInt(t(0)), CInt(t(1)), Int(Val(t(2))))
        End Select

        ParseIso = ParseUtc(ParseIso)
        If hasOfs Then ParseIso = ParseIso - offs
    End If
End Function

Public Function ConvertToIso(ByVal utc_LocalDate As Date) As String
    ConvertToIso = Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-dd\THH:nn:ss.000\Z")
End Function

#If Mac Then
Private Function utc_ConvertDate(ByVal v As Date, Optional ByVal utc_ConvertToUtc As Boolean = False) As Date
    Dim cmd As String, res As utc_ShellResult
    Dim p() As String, d() As String, t() As String
    If utc_ConvertToUtc Then
        cmd = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' '" & Format$(v, "yyyy-mm-dd HH:mm:ss") & "' +'%s'` +'%Y-%m-%d %H:%M:%S'"
    Else
        cmd = "date -jf '%Y-%m-%d %H:%M:%S %z' '" & Format$(v, "yyyy-mm-dd HH:mm:ss") & " +0000' +'%Y-%m-%d %H:%M:%S'"
    End If
    res = utc_ExecuteInShell(cmd)
    If res.utc_Output = "" Then Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
    p = Split(res.utc_Output, " ")
    d = Split(p(0), "-"): t = Split(p(1), ":")
    utc_ConvertDate = DateSerial(d(0), d(1), d(2)) + TimeSerial(t(0), t(1), t(2))
End Function

Private Function utc_ExecuteInShell(ByVal cmd As String) As utc_ShellResult
    Dim f As LongPtr, r As LongPtr, chunk As String
    On Error GoTo eh
    f = utc_popen(cmd, "r")
    If f = 0 Then Exit Function
    Do While utc_feof(f) = 0
        chunk = Space$(256)
        r = CLng(utc_fread(chunk, 1, Len(chunk) - 1, f))
        If r > 0 Then
            chunk = Left$(chunk, CLng(r))
            utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & chunk
        End If
    Loop
eh:
    utc_ExecuteInShell.utc_ExitCode = CLng(utc_pclose(f))
End Function
#Else
Private Function utc_DateToSystemTime(ByVal v As Date) As utc_SYSTEMTIME
    With utc_DateToSystemTime
        .utc_wYear = Year(v): .utc_wMonth = Month(v): .utc_wDay = Day(v)
        .utc_wHour = Hour(v): .utc_wMinute = Minute(v): .utc_wSecond = Second(v)
        .utc_wMilliseconds = 0
    End With
End Function

Private Function utc_SystemTimeToDate(v As utc_SYSTEMTIME) As Date
    utc_SystemTimeToDate = DateSerial(v.utc_wYear, v.utc_wMonth, v.utc_wDay) + _
                           TimeSerial(v.utc_wHour, v.utc_wMinute, v.utc_wSecond)
End Function
#End If
