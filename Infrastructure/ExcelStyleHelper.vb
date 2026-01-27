Imports System
Imports System.Collections.Generic
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace Infrastructure

    Public Module ExcelStyleHelper

        Public Enum RowStatus
            None = 0
            Ok = 1
            Warn = 2
            [Error] = 3
            Candidate = 4
            Deleted = 5
        End Enum

        Public Delegate Function StatusResolver(row As IRow, rowIndex As Integer) As RowStatus

        Public Sub ApplyHeaderStyle(sheet As ISheet, Optional headerRowIndex As Integer = 0, Optional lastColIndex As Integer = -1)
            If sheet Is Nothing Then Return
            Dim headerRow = sheet.GetRow(headerRowIndex)
            If headerRow Is Nothing Then Return

            Dim resolvedLastCol As Integer = lastColIndex
            If resolvedLastCol < 0 Then
                resolvedLastCol = headerRow.LastCellNum - 1
            End If
            If resolvedLastCol < 0 Then Return

            Dim wb = sheet.Workbook
            Dim headerStyle = CreateHeaderStyle(wb)

            For ci As Integer = 0 To resolvedLastCol
                Dim cell = headerRow.GetCell(ci)
                If cell Is Nothing Then
                    cell = headerRow.CreateCell(ci)
                End If
                cell.CellStyle = headerStyle
            Next

            sheet.CreateFreezePane(0, headerRowIndex + 1)
        End Sub

        Public Sub ApplyRowStyleByStatus(sheet As ISheet,
                                         Optional firstDataRowIndex As Integer = 1,
                                         Optional lastRowIndex As Integer = -1,
                                         Optional lastColIndex As Integer = -1,
                                         Optional resolver As StatusResolver = Nothing)
            If sheet Is Nothing OrElse resolver Is Nothing Then Return

            Dim resolvedLastRow As Integer = lastRowIndex
            If resolvedLastRow < 0 Then resolvedLastRow = sheet.LastRowNum
            If resolvedLastRow < firstDataRowIndex Then Return

            Dim headerRow = sheet.GetRow(Math.Max(0, firstDataRowIndex - 1))
            Dim resolvedLastCol As Integer = lastColIndex
            If resolvedLastCol < 0 Then
                If headerRow IsNot Nothing AndAlso headerRow.LastCellNum > 0 Then
                    resolvedLastCol = headerRow.LastCellNum - 1
                Else
                    resolvedLastCol = GuessLastColumn(sheet, firstDataRowIndex, resolvedLastRow)
                End If
            End If
            If resolvedLastCol < 0 Then Return

            Dim wb = sheet.Workbook
            Dim baseStyle As ICellStyle = FindBaseStyle(sheet, firstDataRowIndex, resolvedLastRow, resolvedLastCol, wb)
            Dim styles = CreateStatusStyles(wb, baseStyle)

            For ri As Integer = firstDataRowIndex To resolvedLastRow
                Dim row = sheet.GetRow(ri)
                If row Is Nothing Then Continue For
                Dim status As RowStatus = resolver(row, ri)
                If status = RowStatus.None Then Continue For

                Dim st As ICellStyle = Nothing
                If Not styles.TryGetValue(status, st) Then Continue For

                For ci As Integer = 0 To resolvedLastCol
                    Dim cell = row.GetCell(ci)
                    If cell Is Nothing Then Continue For
                    cell.CellStyle = st
                Next
            Next
        End Sub

        Private Function CreateHeaderStyle(wb As IWorkbook) As ICellStyle
            Dim st As ICellStyle = wb.CreateCellStyle()
            st.FillPattern = FillPattern.SolidForeground
            st.Alignment = HorizontalAlignment.Center
            st.VerticalAlignment = VerticalAlignment.Center
            st.WrapText = True
            SetThinBorders(st)

            Dim font = wb.CreateFont()
            font.IsBold = True
            font.Color = IndexedColors.White.Index
            st.SetFont(font)

            Dim xssfStyle = TryCast(st, XSSFCellStyle)
            If xssfStyle IsNot Nothing Then
                xssfStyle.SetFillForegroundColor(New XSSFColor(New Byte() {&H2F, &H2F, &H2F}))
            Else
                st.FillForegroundColor = IndexedColors.Grey80Percent.Index
            End If

            Return st
        End Function

        Private Function CreateStatusStyles(wb As IWorkbook, baseStyle As ICellStyle) As Dictionary(Of RowStatus, ICellStyle)
            Dim styles As New Dictionary(Of RowStatus, ICellStyle)()

            styles(RowStatus.Ok) = CreateStatusStyle(wb, baseStyle, IndexedColors.LightGreen, IndexedColors.DarkGreen, False)
            styles(RowStatus.Warn) = CreateStatusStyle(wb, baseStyle, IndexedColors.LightYellow, IndexedColors.DarkYellow, False)
            styles(RowStatus.Error) = CreateStatusStyle(wb, baseStyle, IndexedColors.Rose, IndexedColors.Red, False)
            styles(RowStatus.Candidate) = CreateStatusStyle(wb, baseStyle, IndexedColors.LightOrange, IndexedColors.DarkOrange, False)
            styles(RowStatus.Deleted) = CreateStatusStyle(wb, baseStyle, IndexedColors.Grey25Percent, IndexedColors.Grey50Percent, True)

            Return styles
        End Function

        Private Function CreateStatusStyle(wb As IWorkbook,
                                           baseStyle As ICellStyle,
                                           fillColor As IndexedColors,
                                           fontColor As IndexedColors,
                                           strikeout As Boolean) As ICellStyle
            Dim st As ICellStyle = wb.CreateCellStyle()
            If baseStyle IsNot Nothing Then
                st.CloneStyleFrom(baseStyle)
            Else
                SetThinBorders(st)
            End If
            st.FillPattern = FillPattern.SolidForeground
            st.FillForegroundColor = fillColor.Index

            Dim baseFont As IFont = Nothing
            If baseStyle IsNot Nothing Then
                baseFont = wb.GetFontAt(baseStyle.FontIndex)
            End If

            Dim f = wb.CreateFont()
            If baseFont IsNot Nothing Then
                f.FontName = baseFont.FontName
                f.FontHeightInPoints = baseFont.FontHeightInPoints
                f.IsBold = baseFont.IsBold
                f.IsItalic = baseFont.IsItalic
                f.Underline = baseFont.Underline
                f.IsStrikeout = baseFont.IsStrikeout
            End If
            f.Color = fontColor.Index
            If strikeout Then f.IsStrikeout = True
            st.SetFont(f)

            Return st
        End Function

        Private Function FindBaseStyle(sheet As ISheet,
                                       firstRow As Integer,
                                       lastRow As Integer,
                                       lastCol As Integer,
                                       wb As IWorkbook) As ICellStyle
            Dim fallback As ICellStyle = Nothing
            For r As Integer = firstRow To lastRow
                Dim row = sheet.GetRow(r)
                If row Is Nothing Then Continue For
                For c As Integer = 0 To lastCol
                    Dim cell = row.GetCell(c)
                    If cell Is Nothing Then Continue For
                    Dim st = cell.CellStyle
                    If st Is Nothing Then Continue For
                    If fallback Is Nothing Then fallback = st
                    Dim font = wb.GetFontAt(st.FontIndex)
                    If font IsNot Nothing AndAlso Not font.IsBold Then
                        Return st
                    End If
                Next
            Next

            If fallback IsNot Nothing Then Return fallback

            Dim st = wb.CreateCellStyle()
            SetThinBorders(st)
            Return st
        End Function

        Private Function GuessLastColumn(sheet As ISheet, firstRow As Integer, lastRow As Integer) As Integer
            Dim lastCol As Integer = -1
            For r As Integer = firstRow To lastRow
                Dim row = sheet.GetRow(r)
                If row Is Nothing Then Continue For
                Dim rowLast = row.LastCellNum - 1
                If rowLast > lastCol Then lastCol = rowLast
            Next
            Return lastCol
        End Function

        Private Sub SetThinBorders(st As ICellStyle)
            st.BorderBottom = BorderStyle.Thin
            st.BorderTop = BorderStyle.Thin
            st.BorderLeft = BorderStyle.Thin
            st.BorderRight = BorderStyle.Thin
        End Sub

        Public Function GetCellText(row As IRow, colIndex As Integer) As String
            If row Is Nothing OrElse colIndex < 0 Then Return String.Empty
            Dim cell = row.GetCell(colIndex)
            If cell Is Nothing Then Return String.Empty
            Dim v = cell.ToString()
            Return If(v, String.Empty)
        End Function

    End Module

End Namespace
