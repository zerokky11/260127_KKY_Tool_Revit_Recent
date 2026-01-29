Imports System
Imports System.Data
Imports NPOI.SS.UserModel

Namespace Exports

    Public Module ConnectorExport

        ' 저장 대화상자 버전(기존 흐름 유지)
        Public Function SaveWithDialog(resultTable As DataTable) As String
            If resultTable Is Nothing Then Return String.Empty

            Dim outPath As String = Infrastructure.ExcelCore.PickAndSaveXlsx(
                "Connector Diagnostics",
                resultTable,
                "ConnectorDiagnostics.xlsx"
            )

            If String.IsNullOrWhiteSpace(outPath) Then
                Return outPath
            End If

            ' 저장 성공 후: Status 기준 배경색 적용(엑셀 추출 로직에만 영향)
            TryApplyExcelRowStyles(outPath, resultTable)

            Return outPath
        End Function

        ' 경로 고정 저장 버전(브리지에서 경로를 이미 받은 경우)
        Public Sub Save(outPath As String, resultTable As DataTable)
            If String.IsNullOrWhiteSpace(outPath) Then Exit Sub
            If resultTable Is Nothing Then Exit Sub

            Infrastructure.ExcelCore.SaveXlsx(outPath, "Connector Diagnostics", resultTable)

            ' 저장 성공 후: Status 기준 배경색 적용(엑셀 추출 로직에만 영향)
            TryApplyExcelRowStyles(outPath, resultTable)
        End Sub

        Private Sub TryApplyExcelRowStyles(outPath As String, resultTable As DataTable)
            If String.IsNullOrWhiteSpace(outPath) Then Exit Sub
            If resultTable Is Nothing Then Exit Sub
            If resultTable.Columns Is Nothing OrElse resultTable.Columns.Count = 0 Then Exit Sub
            If resultTable.Rows Is Nothing OrElse resultTable.Rows.Count = 0 Then Exit Sub

            Dim resolver As Infrastructure.ExcelStyleHelper.StatusResolver = BuildStatusResolver(resultTable)
            If resolver Is Nothing Then Exit Sub

            Dim lastColIndex As Integer = resultTable.Columns.Count - 1
            If lastColIndex < 0 Then Exit Sub

            ' 엑셀 시트는 0행이 헤더, 데이터는 1행부터 시작
            Dim startRowIndex As Integer = 1
            Dim endRowIndex As Integer = resultTable.Rows.Count ' (헤더 제외 데이터 row 수 == 마지막 rowIndex)

            Try
                Infrastructure.ExcelStyleHelper.ApplyRowStyleByStatus(
                    outPath,
                    startRowIndex,
                    endRowIndex,
                    lastColIndex,
                    resolver
                )
            Catch
                ' 스타일 실패해도 "저장 자체"는 성공으로 유지 (요청: 다른 논리 구조 영향 최소화)
            End Try
        End Sub

        Private Function BuildStatusResolver(resultTable As DataTable) As Infrastructure.ExcelStyleHelper.StatusResolver
            If resultTable Is Nothing Then Return Nothing

            Dim statusIndex As Integer = FindColumnIndex(resultTable, "Status")
            If statusIndex < 0 Then Return Nothing

            Return Function(row As IRow, rowIndex As Integer) As Infrastructure.ExcelStyleHelper.RowStatus
                       Dim statusText As String = Infrastructure.ExcelStyleHelper.GetCellText(row, statusIndex)
                       Return ResolveConnectorStatus(statusText)
                   End Function
        End Function

        ' PR에서 정의한 매핑(OK/WARN/ERROR 계열) 그대로 계승하되, 문자열은 더 안전하게(부분 포함) 처리
        ' - 빈 값은 OK 취급
        Private Function ResolveConnectorStatus(statusText As String) As Infrastructure.ExcelStyleHelper.RowStatus
            If String.IsNullOrWhiteSpace(statusText) Then Return Infrastructure.ExcelStyleHelper.RowStatus.Ok

            Dim s As String = statusText.Trim()

            If s.Equals("OK", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("Match", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("PASS", StringComparison.OrdinalIgnoreCase) Then
                Return Infrastructure.ExcelStyleHelper.RowStatus.Ok
            End If

            If s.Equals("WARN", StringComparison.OrdinalIgnoreCase) OrElse
               s.IndexOf("Check", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("Tolerance", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("Proximity", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return Infrastructure.ExcelStyleHelper.RowStatus.Warn
            End If

            If s.Equals("FAIL", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("ERROR", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("Mismatch", StringComparison.OrdinalIgnoreCase) OrElse
               s.IndexOf("Missing", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("Shared Parameter", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return Infrastructure.ExcelStyleHelper.RowStatus.Error
            End If

            Return Infrastructure.ExcelStyleHelper.RowStatus.Ok
        End Function

        Private Function FindColumnIndex(dt As DataTable, columnName As String) As Integer
            If dt Is Nothing OrElse String.IsNullOrWhiteSpace(columnName) Then Return -1

            For i As Integer = 0 To dt.Columns.Count - 1
                Dim name As String = dt.Columns(i).ColumnName
                If name IsNot Nothing AndAlso name.Equals(columnName, StringComparison.OrdinalIgnoreCase) Then
                    Return i
                End If
            Next

            Return -1
        End Function

    End Module

End Namespace
