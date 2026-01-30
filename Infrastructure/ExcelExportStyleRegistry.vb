Imports System
Imports System.Collections.Generic
Imports KKY_Tool_Revit.Infrastructure
Imports NPOI.SS.UserModel

Namespace Global.KKY_Tool_Revit.Infrastructure

    Public Module ExcelExportStyleRegistry

        Private _inited As Boolean = False
        Private ReadOnly _map As New Dictionary(Of String, ExcelStyleHelper.RowStatusResolver)(StringComparer.OrdinalIgnoreCase)

        Public Sub Register(key As String, resolver As ExcelStyleHelper.RowStatusResolver)
            If String.IsNullOrWhiteSpace(key) Then Exit Sub
            If resolver Is Nothing Then Exit Sub
            EnsureInitialized()
            _map(key.Trim()) = resolver
        End Sub

        Public Function TryGet(key As String) As ExcelStyleHelper.RowStatusResolver
            EnsureInitialized()
            If String.IsNullOrWhiteSpace(key) Then Return Nothing
            Dim r As ExcelStyleHelper.RowStatusResolver = Nothing
            If _map.TryGetValue(key.Trim(), r) Then Return r
            Return Nothing
        End Function

        ' Multi-export 저장 완료 직후 호출: key에 맞는 resolver로 스타일 적용
        Public Sub ApplyStylesForKey(exportKey As String, filePath As String)
            EnsureInitialized()

            If String.IsNullOrWhiteSpace(exportKey) Then Exit Sub
            If String.IsNullOrWhiteSpace(filePath) Then Exit Sub

            Dim resolver As ExcelStyleHelper.RowStatusResolver = Nothing
            If Not _map.TryGetValue(exportKey.Trim(), resolver) Then Exit Sub
            If resolver Is Nothing Then Exit Sub

            ExcelStyleHelper.ApplyRowStyles(filePath, resolver)
        End Sub

        Private Sub EnsureInitialized()
            If _inited Then Exit Sub
            _inited = True

            ' 지금 단계: connector만 우선 등록
            Register("connector", AddressOf ConnectorResolver)
        End Sub

        ' ===== connector 전용 resolver =====
        Private Function ConnectorResolver(row As IRow,
                                           rowIndex As Integer,
                                           headerIndex As Dictionary(Of String, Integer)) As ExcelStyleHelper.RowStatus
            If row Is Nothing Then Return ExcelStyleHelper.RowStatus.None
            If headerIndex Is Nothing Then Return ExcelStyleHelper.RowStatus.None

            Dim statusCol As Integer = ExcelStyleHelper.GetHeaderCol(headerIndex, "Status", "상태")
            If statusCol < 0 Then Return ExcelStyleHelper.RowStatus.None

            Dim txt As String = ExcelStyleHelper.GetCellText(row, statusCol)
            Return ExcelStyleHelper.ResolveConnectorStatus(txt)
        End Function

    End Module

End Namespace
