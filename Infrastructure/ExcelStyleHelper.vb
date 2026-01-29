Imports System
Imports System.Collections.Generic
Imports System.IO
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace Infrastructure

    Public Module ExcelStyleHelper

        Public Enum RowStatus
            None = 0
            Ok = 1
            Warn = 2
            [Error] = 3
        End Enum

        ' rowIndex: 시트의 Row 인덱스(0-based)
        Public Delegate Function StatusResolver(row As IRow, rowIndex As Integer) As RowStatus

        ' 저장된 xlsx를 다시 열어, resolver 규칙대로 행 스타일을 적용한 뒤 다시 저장한다.
        ' - startRowIndex/endRowIndex: 시트 Row 인덱스(0-based, inclusive)
        ' - lastColIndex: 0-based, inclusive
        Public Sub ApplyRowStyleByStatus(filePath As String,
                                        startRowIndex As Integer,
                                        endRowIndex As Integer,
                                        lastColIndex As Integer,
                                        resolver As StatusResolver)

            If String.IsNullOrWhiteSpace(filePath) Then Exit Sub
            If resolver Is Nothing Then Exit Sub
            If Not File.Exists(filePath) Then Exit Sub
            If lastColIndex < 0 Then Exit Sub

            Dim wb As XSSFWorkbook = Nothing

            Try
                ' 1) Read
                Using fsRead As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    wb = New XSSFWorkbook(fsRead)
                End Using

                If wb Is Nothing OrElse wb.NumberOfSheets <= 0 Then Exit Sub

                Dim sh As ISheet = wb.GetSheetAt(0)
                If sh Is Nothing Then Exit Sub

                ApplyRowStyleByStatusInternal(sh, startRowIndex, endRowIndex, lastColIndex, resolver)

                ' 2) Write (overwrite)
                Using fsWrite As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fsWrite)
                End Using

            Finally
                If wb IsNot Nothing Then
                    Try
                        wb.Close()
                    Catch
                        ' ignore
                    End Try
                End If
            End Try
        End Sub

        Private Sub ApplyRowStyleByStatusInternal(sh As ISheet,
                                                 startRowIndex As Integer,
                                                 endRowIndex As Integer,
                                                 lastColIndex As Integer,
                                                 resolver As StatusResolver)

            If sh Is Nothing Then Exit Sub
            If resolver Is Nothing Then Exit Sub
            If lastColIndex < 0 Then Exit Sub

            Dim wb As IWorkbook = sh.Workbook
            If wb Is Nothing Then Exit Sub

            Dim realStart As Integer = Math.Max(0, startRowIndex)
            Dim realEnd As Integer = endRowIndex
            If realEnd <= 0 OrElse realEnd > sh.LastRowNum Then realEnd = sh.LastRowNum
            If realEnd < realStart Then Exit Sub

            ' (status + baseStyleIndex) 단위로만 스타일 생성/재사용해서 파일 비대화 방지
            Dim styleCache As New Dictionary(Of String, ICellStyle)(StringComparer.Ordinal)

            For ri As Integer = realStart To realEnd
                Dim row As IRow = sh.GetRow(ri)
                If row Is Nothing Then Continue For

                Dim st As RowStatus = resolver(row, ri)
                If st = RowStatus.None Then Continue For

                For ci As Integer = 0 To lastColIndex
                    Dim cell As ICell = row.GetCell(ci)
                    If cell Is Nothing Then Continue For ' 셀 생성은 하지 않음(시트 구조 변경 최소화)

                    Dim baseStyle As ICellStyle = cell.CellStyle
                    Dim styled As ICellStyle = GetOrCreateStyledStyle(wb, baseStyle, st, styleCache)
                    If styled IsNot Nothing Then
                        cell.CellStyle = styled
                    End If
                Next
            Next
        End Sub

        Private Function GetOrCreateStyledStyle(wb As IWorkbook,
                                               baseStyle As ICellStyle,
                                               st As RowStatus,
                                               cache As Dictionary(Of String, ICellStyle)) As ICellStyle
            If wb Is Nothing Then Return Nothing
            If st = RowStatus.None Then Return baseStyle

            Dim baseIdx As Short = -1
            If baseStyle IsNot Nothing Then baseIdx = baseStyle.Index

            Dim key As String = (CInt(st).ToString() & "|" & baseIdx.ToString())
            Dim existing As ICellStyle = Nothing
            If cache IsNot Nothing AndAlso cache.TryGetValue(key, existing) Then
                Return existing
            End If

            Dim ns As ICellStyle = wb.CreateCellStyle()
            If baseStyle IsNot Nothing Then
                ns.CloneStyleFrom(baseStyle)
            End If

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

            If cache IsNot Nothing Then
                cache(key) = ns
            End If

            Return ns
        End Function

        Public Function GetCellText(row As IRow, colIndex As Integer) As String
            If row Is Nothing OrElse colIndex < 0 Then Return String.Empty

            Dim cell As ICell = row.GetCell(colIndex)
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

    End Module

End Namespace
