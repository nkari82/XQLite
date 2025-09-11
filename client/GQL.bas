Option Explicit

Public GQL_BASE As String   ' 예: "http://localhost:8000/graphql"
Public GQL_API  As String   ' 예: "devkey"

Public Function GqlCall(ByVal query As String, ByVal variablesJson As String) As Object
    Dim http As Object: Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", GQL_BASE, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "X-API-Key", GQL_API
    Dim body As String
    body = "{""query"":" & JsonQuote(query) & ",""variables"":" & IIf(Len(variablesJson) = 0, "{}", variablesJson) & "}"
    http.send body
    If http.Status <> 200 Then Err.Raise 513, , "GQL Error: " & http.Status & " " & http.responseText
    Dim resp As Object: Set resp = JsonConverter.ParseJson(http.responseText)
    If Not resp("errors") Is Nothing Then
        Dim e As Object: Set e = resp("errors")(1)
        Err.Raise 514, , "GQL errors: " & e("message")
    End If
    Set GqlCall = resp("data")
End Function

Public Function JsonQuote(ByVal s As String) As String
    Dim t As String
    t = Replace(s, """", "\""")
    t = Replace(t, vbCrLf, "\n"): t = Replace(t, vbCr, "\n"): t = Replace(t, vbLf, "\n")
    JsonQuote = """" & t & """"
End Function

Public Function JsonValue(v As Variant) As String
    If IsEmpty(v) Or IsNull(v) Then
        JsonValue = "null"
    ElseIf VarType(v) = vbBoolean Then
        JsonValue = IIf(v, "true", "false")
    ElseIf IsNumeric(v) Then
        JsonValue = CStr(v)
    Else
        JsonValue = JsonQuote(CStr(v))
    End If
End Function

Public Function Gql_LoadRows(ByVal tableName As String, Optional ByVal sinceVer As Long = 0) As Object
    Dim q As String
    q = "query Rows($table:String!,$sv:Int){ rows(table:$table,since_version:$sv) }"
    Dim vars As String: vars = "{""table"":" & JsonQuote(tableName) & ",""sv"":" & CStr(sinceVer) & "}"
    Dim data As Object: Set data = GqlCall(q, vars)
    Set Gql_LoadRows = data("rows")
End Function

Public Function Gql_UpsertRows(ByVal tableName As String, ByVal actor As String, ByVal rowsJson As String) As Object
    Dim q As String
    q = "mutation Upsert($table:String!,$actor:String,$rows:[RowIn!]!){ upsertRows(table:$table,actor:$actor,rows:$rows){ results{ id status db_version message } snapshot } }"
    Dim vars As String
    vars = "{""table"":" & JsonQuote(tableName) & ",""actor"":" & JsonQuote(actor) & ",""rows"":" & rowsJson & "}"
    Dim data As Object: Set data = GqlCall(q, vars)
    Set Gql_UpsertRows = data("upsertRows")
End Function

Public Sub Gql_CreateTable(ByVal tableName As String, ByVal headers As Variant, ByVal firstRow As Variant)
    Dim i As Long, n As Long: n = UBound(headers, 2)
    Dim cols() As String: ReDim cols(1 To n)
    For i = 1 To n
        Dim cname As String: cname = CStr(headers(1, i))
        Dim ctype As String: ctype = "TEXT"
        On Error Resume Next
        Dim samp As Variant: samp = IIf(IsEmpty(firstRow), Empty, firstRow(1, i))
        On Error GoTo 0
        If Not IsEmpty(samp) And IsNumeric(samp) Then
            If samp = CLng(samp) Then ctype = "INTEGER" Else ctype = "REAL"
        End If
        cols(i) = "{""name"":" & JsonQuote(cname) & ",""type"":" & JsonQuote(ctype) & "}"
    Next i
    Dim q As String: q = "mutation($input: CreateTableInput!){ createTable(input:$input) }"
    Dim vars As String
    vars = "{""input"":{""table"":" & JsonQuote(tableName) & ",""with_meta"":true,""columns"":[" & Join(cols, ",") & "]}}"
    On Error Resume Next: Call GqlCall(q, vars): On Error GoTo 0
End Sub

Public Sub Gql_AddColumns(ByVal tableName As String, ByVal headers As Variant, ByVal firstRow As Variant)
    Dim i As Long, n As Long: n = UBound(headers, 2)
    Dim cols() As String: ReDim cols(1 To n)
    For i = 1 To n
        Dim cname As String: cname = CStr(headers(1, i))
        Dim ctype As String: ctype = "TEXT"
        On Error Resume Next
        Dim samp As Variant: samp = IIf(IsEmpty(firstRow), Empty, firstRow(1, i))
        On Error GoTo 0
        If Not IsEmpty(samp) And IsNumeric(samp) Then
            If samp = CLng(samp) Then ctype = "INTEGER" Else ctype = "REAL"
        End If
        cols(i) = "{""name"":" & JsonQuote(cname) & ",""type"":" & JsonQuote(ctype) & "}"
    Next i
    Dim q As String: q = "mutation($input: AddColumnsInput!){ addColumns(input:$input) }"
    Dim vars As String
    vars = "{""input"":{""table"":" & JsonQuote(tableName) & ",""columns"":[" & Join(cols, ",") & "]}}"
    On Error Resume Next: Call GqlCall(q, vars): On Error GoTo 0
End Sub
