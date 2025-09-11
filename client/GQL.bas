Option Explicit

Public GQL_BASE As String   ' e.g. "http://localhost:8000/graphql"
Public GQL_API  As String   ' e.g. "devkey"

' ───────────────────────── Core Call ─────────────────────────

Public Function GqlCall(ByVal query As String, ByVal variablesJson As String) As Object
    If Len(Trim$(GQL_BASE)) = 0 Then Err.Raise 511, , "GQL Error: GQL_BASE is empty."
    If Len(variablesJson) = 0 Then variablesJson = "{}"

    Dim body As String
    body = "{""query"":" & JsonQuote(query) & ",""variables"":" & variablesJson & "}"

    Dim http As Object
    Dim usedServerHttp As Boolean
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    usedServerHttp = Not http Is Nothing
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo 0

    If http Is Nothing Then Err.Raise 512, , "GQL Error: Cannot create HTTP client."

    If usedServerHttp Then
        On Error Resume Next
        ' resolve/connect/send/receive (ms)
        http.setTimeouts 5000, 5000, 15000, 30000
        On Error GoTo 0
    End If

    http.Open "POST", GQL_BASE, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    If Len(GQL_API) > 0 Then http.setRequestHeader "X-API-Key", GQL_API
    http.send body

    Dim statusCode As Long
    On Error Resume Next
    statusCode = http.Status
    On Error GoTo 0

    Dim respText As String: respText = ""
    On Error Resume Next
    respText = http.responseText
    On Error GoTo 0

    ' 200이 아니어도 GraphQL errors 본문이 담겨 있을 수 있으므로 먼저 파싱 시도
    Dim resp As Object
    On Error Resume Next
    Set resp = JsonConverter.ParseJson(respText)
    On Error GoTo 0

    If statusCode <> 200 Then
        If Not resp Is Nothing Then
            Dim msg As String: msg = ExtractGraphQLErrorMessage(resp)
            If Len(msg) > 0 Then
                Err.Raise 513, , "GQL HTTP " & statusCode & ": " & msg
            End If
        End If
        Err.Raise 513, , "GQL HTTP " & statusCode & ": " & respText
    End If

    If resp Is Nothing Then
        Err.Raise 514, , "GQL Error: Empty/invalid JSON response."
    End If

    ' GraphQL spec: {"data":..., "errors":[...]} 형태
    If resp.Exists("errors") Then
        Dim errs As Object: Set errs = resp("errors")
        If Not errs Is Nothing Then
            Dim firstMsg As String: firstMsg = ExtractGraphQLErrorMessage(resp)
            If Len(firstMsg) = 0 Then firstMsg = "Unknown GraphQL errors."
            Err.Raise 514, , "GQL errors: " & firstMsg
        End If
    End If

    If Not resp.Exists("data") Then
        Err.Raise 515, , "GQL Error: Missing 'data' in response."
    End If

    Set GqlCall = resp("data")
End Function

Private Function ExtractGraphQLErrorMessage(ByVal resp As Object) As String
    On Error GoTo NOPE
    If resp.Exists("errors") Then
        Dim errs As Object: Set errs = resp("errors")
        ' VBA-JSON 배열은 Collection인 경우가 흔함
        Dim e As Variant
        If TypeName(errs) = "Collection" Then
            If errs.Count >= 1 Then
                Set e = errs.Item(1)
                If Not e Is Nothing Then
                    If TypeName(e) = "Dictionary" Then
                        If e.Exists("message") Then
                            ExtractGraphQLErrorMessage = CStr(e("message"))
                            Exit Function
                        End If
                    End If
                End If
            End If
        ElseIf TypeName(errs) = "Dictionary" Then
            If errs.Exists("message") Then
                ExtractGraphQLErrorMessage = CStr(errs("message"))
                Exit Function
            End If
        End If
    End If
NOPE:
    ExtractGraphQLErrorMessage = ""
End Function

' ───────────────────────── JSON helpers ─────────────────────────

Public Function JsonQuote(ByVal s As String) As String
    ' 최소 안전 이스케이프: \, ", 제어문자
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbBack, "\b")
    t = Replace(t, vbFormFeed, "\f")
    t = Replace(t, vbCrLf, "\n")
    t = Replace(t, vbCr, "\n")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbTab, "\t")
    JsonQuote = """" & t & """"
End Function

Public Function JsonValue(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        JsonValue = "null"
    ElseIf VarType(v) = vbBoolean Then
        JsonValue = IIf(v, "true", "false")
    ElseIf IsDate(v) Then
        JsonValue = """" & Format$(v, "yyyy-mm-dd\Thh:nn:ss") & """"
    ElseIf IsNumeric(v) Then
        JsonValue = JNum(v)
    Else
        JsonValue = JsonQuote(CStr(v))
    End If
End Function

Private Function JNum(ByVal v As Variant) As String
    ' 소수점 구분자 강제 '.' (로케일 무시)
    Dim s As String: s = CStr(v)
    JNum = Replace(s, Application.International(xlDecimalSeparator), ".")
End Function

' ───────────────────────── High-level ops ─────────────────────────

Public Function Gql_LoadRows(ByVal tableName As String, Optional ByVal sinceVer As Long = 0) As Object
    If sinceVer < 0 Then sinceVer = 0
    Dim q As String
    q = "query Rows($table:String!,$sv:Int){ rows(table:$table,since_version:$sv) }"
    Dim vars As String
    vars = "{""table"":" & JsonQuote(tableName) & ",""sv"":" & CStr(sinceVer) & "}"
    Dim data As Object: Set data = GqlCall(q, vars)
    If data Is Nothing Then Set Gql_LoadRows = Nothing: Exit Function
    If data.Exists("rows") Then Set Gql_LoadRows = data("rows") Else Set Gql_LoadRows = Nothing
End Function

Public Function Gql_UpsertRows(ByVal tableName As String, ByVal actor As String, ByVal rowsJson As String) As Object
    Dim q As String
    q = "mutation Upsert($table:String!,$actor:String,$rows:[RowIn!]!){ upsertRows(table:$table,actor:$actor,rows:$rows){ results{ id status db_version message } snapshot } }"
    Dim vars As String
    vars = "{""table"":" & JsonQuote(tableName) & ",""actor"":" & JsonQuote(actor) & ",""rows"":" & rowsJson & "}"
    Dim data As Object: Set data = GqlCall(q, vars)
    If data Is Nothing Then Set Gql_UpsertRows = Nothing: Exit Function
    If data.Exists("upsertRows") Then Set Gql_UpsertRows = data("upsertRows") Else Set Gql_UpsertRows = Nothing
End Function

Public Sub Gql_CreateTable(ByVal tableName As String, ByVal headers As Variant, ByVal firstRow As Variant)
    If Not IsArray(headers) Then Exit Sub
    On Error GoTo SAFE_EXIT

    Dim n As Long
    n = UBound(headers, 2)
    If n <= 0 Then GoTo SAFE_EXIT

    Dim cols() As String: ReDim cols(1 To n)
    Dim i As Long
    For i = 1 To n
        Dim cname As String: cname = CStr(headers(1, i))
        Dim ctype As String: ctype = "TEXT"

        Dim samp As Variant
        If IsArray(firstRow) Then
            On Error Resume Next
            samp = firstRow(1, i)
            On Error GoTo 0
            If Not IsEmpty(samp) Then
                If IsNumeric(samp) Then
                    If samp = CLng(samp) Then ctype = "INTEGER" Else ctype = "REAL"
                ElseIf IsDate(samp) Then
                    ctype = "TEXT" ' ?쒕쾭媛 ISO8601 泥섎━
                End If
            End If
        End If

        cols(i) = "{""name"":" & JsonQuote(cname) & ",""type"":" & JsonQuote(ctype) & "}"
    Next i

    Dim q As String: q = "mutation($input: CreateTableInput!){ createTable(input:$input) }"
    Dim vars As String
    vars = "{""input"":{""table"":" & JsonQuote(tableName) & ",""with_meta"":true,""columns"":[" & Join(cols, ",") & "]}}"

    On Error Resume Next: Call GqlCall(q, vars): On Error GoTo 0
SAFE_EXIT:
End Sub

Public Sub Gql_AddColumns(ByVal tableName As String, ByVal headers As Variant, ByVal firstRow As Variant)
    If Not IsArray(headers) Then Exit Sub
    On Error GoTo SAFE_EXIT

    Dim n As Long
    n = UBound(headers, 2)
    If n <= 0 Then GoTo SAFE_EXIT

    Dim cols() As String: ReDim cols(1 To n)
    Dim i As Long
    For i = 1 To n
        Dim cname As String: cname = CStr(headers(1, i))
        Dim ctype As String: ctype = "TEXT"

        Dim samp As Variant
        If IsArray(firstRow) Then
            On Error Resume Next
            samp = firstRow(1, i)
            On Error GoTo 0
            If Not IsEmpty(samp) Then
                If IsNumeric(samp) Then
                    If samp = CLng(samp) Then ctype = "INTEGER" Else ctype = "REAL"
                ElseIf IsDate(samp) Then
                    ctype = "TEXT"
                End If
            End If
        End If

        cols(i) = "{""name"":" & JsonQuote(cname) & ",""type"":" & JsonQuote(ctype) & "}"
    Next i

    Dim q As String: q = "mutation($input: AddColumnsInput!){ addColumns(input:$input) }"
    Dim vars As String
    vars = "{""input"":{""table"":" & JsonQuote(tableName) & ",""columns"":[" & Join(cols, ",") & "]}}"

    On Error Resume Next: Call GqlCall(q, vars): On Error GoTo 0
SAFE_EXIT:
End Sub
