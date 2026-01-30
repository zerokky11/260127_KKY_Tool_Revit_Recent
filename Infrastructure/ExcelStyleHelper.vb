Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Threading
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace Global.KKY_Tool_Revit.Infrastructure

    Public Module ExcelStyleHelper

        Public Enum RowStatus
            None = 0
            Ok = 1
            Warn = 2
            [Error] = 3
        End Enum

        ' headerIndex: 헤더명(정규화) -> col index
        Public Delegate Function RowStatusResolver(row As IRow,
                                                   rowIndex As Integer,
                                                   headerIndex As Dictionary(Of String, Integer)) As RowStatus

        Public Sub ApplyRowStyles(filePath As String, resolver As RowStatusResolver)
            If String.IsNullOrWhiteSpace(filePath) Then Exit Sub
            If resolver Is Nothing Then Exit Sub
            If Not File.Exists(filePath) Then Exit Sub

            ' 저장 직후 잠금 대비(짧게 재시도)
            For attempt As Integer = 1 To 3
                Try
                    ApplyRowStylesOnce(filePath, resolver)
                    Return
                Catch
                    Thread.Sleep(80)
                End Try
            Next
        End Sub

        Private Sub ApplyRowStylesOnce(filePath As String, resolver As RowStatusResolver)
            Dim wb As XSSFWorkbook = Nothing

            Try
                Using fsRead As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    wb = New XSSFWorkbook(fsRead)
                End Using
                If wb Is Nothing OrElse wb.NumberOfSheets <= 0 Then Exit Sub

                Dim sh As ISheet = wb.GetSheetAt(0)
                If sh Is Nothing Then Exit Sub

                Dim header As IRow = sh.GetRow(0)
                If header Is Nothing Then Exit Sub

                Dim headerIndex As Dictionary(Of String, Integer) = BuildHeaderIndex(header)
                Dim lastCol As Integer = header.LastCellNum - 1
                If lastCol < 0 Then Exit Sub

                Dim lastRow As Integer = sh.LastRowNum
                If lastRow < 1 Then Exit Sub

                Dim styleCache As New Dictionary(Of String, ICellStyle)(StringComparer.Ordinal)

                For ri As Integer = 1 To lastRow
                    Dim row As IRow = sh.GetRow(ri)
                    If row Is Nothing Then Continue For

                    Dim st As RowStatus = resolver(row, ri, headerIndex)
                    If st = RowStatus.None Then Continue For

                    For ci As Integer = 0 To lastCol
                        Dim cell As ICell = row.GetCell(ci)
                        If cell Is Nothing Then
                            ' 빈칸도 색칠되도록 셀 생성
                            cell = row.CreateCell(ci, CellType.Blank)
                        End If

                        Dim baseStyle As ICellStyle = cell.CellStyle
                        Dim styled As ICellStyle = GetOrCreateStyledStyle(wb, baseStyle, st, styleCache)
                        If styled IsNot Nothing Then cell.CellStyle = styled
                    Next
                Next

                Using fsWrite As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fsWrite)
                End Using

            Finally
                If wb IsNot Nothing Then
                    Try : wb.Close() : Catch : End Try
                End If
            End Try
        End Sub

        Private Function BuildHeaderIndex(header As IRow) As Dictionary(Of String, Integer)
            Dim dict As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            If header Is Nothing Then Return dict

            Dim lastCol As Integer = header.LastCellNum - 1
            If lastCol < 0 Then Return dict

            For ci As Integer = 0 To lastCol
                Dim txt As String = GetCellText(header, ci)
                If String.IsNullOrWhiteSpace(txt) Then Continue For

                Dim key As String = NormalizeHeader(txt)
                If Not dict.ContainsKey(key) Then dict(key) = ci
            Next

            Return dict
        End Function

        Private Function NormalizeHeader(s As String) As String
            If s Is Nothing Then Return String.Empty
            Return s.Trim().ToLowerInvariant()
        End Function

        Public Function GetHeaderCol(headerIndex As Dictionary(Of String, Integer), ParamArray candidates As String()) As Integer
            If headerIndex Is Nothing OrElse candidates Is Nothing OrElse candidates.Length = 0 Then Return -1

            For Each c As String In candidates
                If String.IsNullOrWhiteSpace(c) Then Continue For
                Dim key As String = NormalizeHeader(c)
                Dim idx As Integer = -1
                If headerIndex.TryGetValue(key, idx) Then Return idx
            Next

            Return -1
        End Function

        Private Function GetOrCreateStyledStyle(wb As IWorkbook,
                                               baseStyle As ICellStyle,
                                               st As RowStatus,
                                               cache As Dictionary(Of String, ICellStyle)) As ICellStyle
            If wb Is Nothing Then Return Nothing
            If st = RowStatus.None Then Return baseStyle

            Dim baseIdx As Short = -1
            If baseStyle IsNot Nothing Then baseIdx = baseStyle.Index

            Dim key As String = CInt(st).ToString() & "|" & baseIdx.ToString()
            Dim existing As ICellStyle = Nothing
            If cache IsNot Nothing AndAlso cache.TryGetValue(key, existing) Then
                Return existing
            End If

            Dim ns As ICellStyle = wb.CreateCellStyle()
            If baseStyle IsNot Nothing Then ns.CloneStyleFrom(baseStyle)

            ns.FillPattern = FillPattern.SolidForeground
            Select Case st
                Case RowStatus.Ok
                    ns.FillForegroundColor = CType(IndexedColors.LightGreen.Index, Short)
                Case RowStatus.Warn
                    ns.FillForegroundColor = CType(IndexedColors.LightYellow.Index, Short)
                Case RowStatus.Error
                    ns.FillForegroundColor = CType(IndexedColors.Rose.Index, Short)
                Case Else
                    ns.FillPattern = FillPattern.NoFill
            End Select

            If cache IsNot Nothing Then cache(key) = ns
            Return ns
        End Function

        Public Function GetCellText(row As IRow, colIndex As Integer) As String
            If row Is Nothing OrElse colIndex < 0 Then Return String.Empty
            Return GetCellText(row.GetCell(colIndex))
        End Function

        Public Function GetCellText(cell As ICell) As String
            If cell Is Nothing Then Return String.Empty
            Try
                Select Case cell.CellType
                    Case CellType.String
                        Return If(cell.StringCellValue, String.Empty)
                    Case CellType.Numeric
                        Return cell.NumericCellValue.ToString()
                    Case CellType.Boolean
                        Return cell.BooleanCellValue.ToString()
                    Case CellType.Formula
                        Return cell.ToString()
                    Case CellType.Blank
                        Return String.Empty
                    Case Else
                        Return cell.ToString()
                End Select
            Catch
                Return cell.ToString()
            End Try
        End Function

        ' ===== connector 상태 매핑 =====
        Public Function ResolveConnectorStatus(statusText As String) As RowStatus
            If String.IsNullOrWhiteSpace(statusText) Then Return RowStatus.Ok
            Dim s As String = statusText.Trim()

            If s.Equals("OK", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("Match", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("PASS", StringComparison.OrdinalIgnoreCase) Then
                Return RowStatus.Ok
            End If

            If s.Equals("WARN", StringComparison.OrdinalIgnoreCase) OrElse
               s.IndexOf("Check", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("Tolerance", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("Proximity", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return RowStatus.Warn
            End If

            If s.Equals("FAIL", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("ERROR", StringComparison.OrdinalIgnoreCase) OrElse
               s.Equals("Mismatch", StringComparison.OrdinalIgnoreCase) OrElse
               s.IndexOf("Missing", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               s.IndexOf("Shared Parameter", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return RowStatus.Error
            End If

            Return RowStatus.Ok
        End Function

    End Module

End Namespace
