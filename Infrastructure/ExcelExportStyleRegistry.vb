Imports System.Data

Namespace KKY_Tool_Revit.Infrastructure

    Public Module ExcelExportStyleRegistry

        Public Delegate Function RowStatusResolver(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus

        Private ReadOnly _resolvers As New Dictionary(Of String, RowStatusResolver)(StringComparer.OrdinalIgnoreCase)

        Sub New()
            ' 기본 규칙 등록 (필요하면 기능별로 Register로 덮어쓰기 가능)
            Register("connector", AddressOf ResolveConnector)
            Register("connector diagnostics", AddressOf ResolveConnector)
            Register("familylink", AddressOf ResolveIssueLike)
            Register("familylinkaudit", AddressOf ResolveIssueLike)
            Register("family link audit", AddressOf ResolveIssueLike)

            Register("pms vs segment size검토", AddressOf ResolveResultLike)
            Register("pipe segment class검토", AddressOf ResolveResultLike)
            Register("routing class검토", AddressOf ResolveResultLike)
        End Sub

        Public Sub Register(key As String, resolver As RowStatusResolver)
            If String.IsNullOrWhiteSpace(key) OrElse resolver Is Nothing Then Return
            _resolvers(key.Trim()) = resolver
        End Sub

        Public Function Resolve(sheetNameOrKey As String, row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            If row Is Nothing OrElse table Is Nothing Then Return ExcelStyleHelper.RowStatus.None

            Dim key = NormalizeKey(sheetNameOrKey)
            Dim fn As RowStatusResolver = Nothing

            If Not String.IsNullOrWhiteSpace(key) AndAlso _resolvers.TryGetValue(key, fn) Then
                Return SafeResolve(fn, row, table)
            End If

            ' 시트명이 key로 등록되어있을 수도 있음
            If Not String.IsNullOrWhiteSpace(sheetNameOrKey) AndAlso _resolvers.TryGetValue(sheetNameOrKey.Trim(), fn) Then
                Return SafeResolve(fn, row, table)
            End If

            ' 최후: 범용 규칙
            Return ResolveGeneric(row, table)
        End Function

        Private Function SafeResolve(fn As RowStatusResolver, row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Try
                Return fn(row, table)
            Catch
                Return ExcelStyleHelper.RowStatus.None
            End Try
        End Function

        Private Function NormalizeKey(sheetNameOrKey As String) As String
            If String.IsNullOrWhiteSpace(sheetNameOrKey) Then Return ""
            Dim s = sheetNameOrKey.Trim().ToLowerInvariant()

            If s.Contains("connector") Then Return "connector"
            If s.Contains("familylink") OrElse s.Contains("family link") Then Return "familylinkaudit"
            If s.Contains("pms") OrElse s.Contains("segment") Then Return s ' PMS 시트는 시트명 그대로도 등록해둠

            Return s
        End Function

        ' ---- 기능별 기본 Resolver들 ----

        Private Function ResolveConnector(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            ' Multi 쪽에서 "Status" = "오류 없음" 형태가 기본값임:contentReference[oaicite:5]{index=5}
            Dim statusText = GetColText(row, table, "Status")
            If IsOkLike(statusText) Then Return ExcelStyleHelper.RowStatus.None

            If LooksError(statusText) Then Return ExcelStyleHelper.RowStatus.Error
            Return ExcelStyleHelper.RowStatus.Warning
        End Function

        Private Function ResolveIssueLike(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Dim issue = GetColText(row, table, "Issue")
            If IsOkLike(issue) Then Return ExcelStyleHelper.RowStatus.None
            If LooksError(issue) Then Return ExcelStyleHelper.RowStatus.Error
            Return ExcelStyleHelper.RowStatus.Warning
        End Function

        Private Function ResolveResultLike(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Dim result = GetColText(row, table, "Result")
            If IsOkLike(result) Then Return ExcelStyleHelper.RowStatus.None
            If LooksError(result) Then Return ExcelStyleHelper.RowStatus.Error
            Return ExcelStyleHelper.RowStatus.Warning
        End Function

        Private Function ResolveGeneric(row As DataRow, table As DataTable) As ExcelStyleHelper.RowStatus
            Dim candidates = New String() {
                "Status", "Result", "Issue",
                "Class검토결과", "Class검토", "검토결과",
                "Error", "ErrorMessage", "Notes"
            }

            Dim txt As String = ""
            For Each c In candidates
                txt = GetColText(row, table, c)
                If Not String.IsNullOrWhiteSpace(txt) Then Exit For
            Next

            If String.IsNullOrWhiteSpace(txt) Then Return ExcelStyleHelper.RowStatus.None
            If IsOkLike(txt) Then Return ExcelStyleHelper.RowStatus.None
            If LooksError(txt) Then Return ExcelStyleHelper.RowStatus.Error
            Return ExcelStyleHelper.RowStatus.Warning
        End Function

        ' ---- 헬퍼 ----

        Private Function GetColText(row As DataRow, table As DataTable, colName As String) As String
            If table.Columns.Contains(colName) Then
                Dim v = row(colName)
                If v Is Nothing OrElse v Is DBNull.Value Then Return ""
                Return v.ToString().Trim()
            End If
            Return ""
        End Function

        Private Function IsOkLike(s As String) As Boolean
            If String.IsNullOrWhiteSpace(s) Then Return True
            Dim t = s.Trim().ToLowerInvariant()
            If t = "ok" OrElse t = "pass" OrElse t = "success" Then Return True
            If t.Contains("오류 없음") OrElse t.Contains("정상") OrElse t.Contains("이상 없음") Then Return True
            Return False
        End Function

        Private Function LooksError(s As String) As Boolean
            If String.IsNullOrWhiteSpace(s) Then Return False
            Dim t = s.Trim().ToLowerInvariant()
            If t.Contains("error") OrElse t.Contains("fail") Then Return True
            If t.Contains("실패") OrElse t.Contains("오류") Then Return True
            Return False
        End Function

        ' ConnectorExport 등에서 호출 중일 수 있는 호환용 (저장 시 이미 스타일을 먹이므로 no-op)
        Public Sub ApplyStylesForKey(styleKey As String, xlsxPath As String)
            ' intentionally no-op
        End Sub

    End Module

End Namespace
