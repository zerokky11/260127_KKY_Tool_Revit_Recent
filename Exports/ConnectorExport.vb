Imports System.Data
Imports KKY_Tool_Revit.Infrastructure

Namespace Exports

    Public Module ConnectorExport

        Public Function SaveWithDialog(resultTable As DataTable) As String
            If resultTable Is Nothing Then Return String.Empty
            Dim resolver = BuildStatusResolver(resultTable)
            Return ExcelCore.PickAndSaveXlsx("Connector Diagnostics", resultTable, "ConnectorDiagnostics.xlsx", False, Nothing, resolver)
        End Function

        ' 경로 고정 저장 버전(브리지에서 경로를 이미 받은 경우)
        Public Sub Save(outPath As String, resultTable As DataTable)
            If resultTable Is Nothing Then Exit Sub
            Dim resolver = BuildStatusResolver(resultTable)
            ExcelCore.SaveXlsx(outPath, "Connector Diagnostics", resultTable, False, Nothing, resolver)
        End Sub

        Private Function BuildStatusResolver(resultTable As DataTable) As ExcelStyleHelper.StatusResolver
            If resultTable Is Nothing Then Return Nothing
            Dim statusIndex As Integer = FindColumnIndex(resultTable, "Status")
            If statusIndex < 0 Then Return Nothing

            Return Function(row As NPOI.SS.UserModel.IRow, rowIndex As Integer) As ExcelStyleHelper.RowStatus
                       Dim statusText As String = ExcelStyleHelper.GetCellText(row, statusIndex)
                       Return ResolveConnectorStatus(statusText)
                   End Function
        End Function

        Private Function ResolveConnectorStatus(statusText As String) As ExcelStyleHelper.RowStatus
            If String.IsNullOrWhiteSpace(statusText) Then Return ExcelStyleHelper.RowStatus.Ok
            Dim s = statusText.Trim()

            If s.Equals("OK", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("Match", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("PASS", StringComparison.OrdinalIgnoreCase) Then
                Return ExcelStyleHelper.RowStatus.Ok
            End If

            If s.Equals("WARN", StringComparison.OrdinalIgnoreCase) OrElse
               s.IndexOf("Check", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("Tolerance", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.Equals(" ʿ(Proximity)", StringComparison.OrdinalIgnoreCase) Then
                Return ExcelStyleHelper.RowStatus.Warn
            End If

            If s.Equals("FAIL", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("ERROR", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("Mismatch", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("  ü ", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("Shared Parameter  ʿ", StringComparison.OrdinalIgnoreCase) OrElse
               s.IndexOf("Missing", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return ExcelStyleHelper.RowStatus.Error
            End If

            Return ExcelStyleHelper.RowStatus.Ok
        End Function

        Private Function FindColumnIndex(dt As DataTable, columnName As String) As Integer
            If dt Is Nothing OrElse String.IsNullOrWhiteSpace(columnName) Then Return -1
            For i = 0 To dt.Columns.Count - 1
                Dim name = dt.Columns(i).ColumnName
                If name IsNot Nothing AndAlso name.Equals(columnName, StringComparison.OrdinalIgnoreCase) Then
                    Return i
                End If
            Next
            Return -1
        End Function

    End Module
End Namespace
