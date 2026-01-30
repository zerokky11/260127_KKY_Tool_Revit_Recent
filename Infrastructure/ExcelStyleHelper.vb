Imports System.Runtime.CompilerServices
Imports NPOI.SS.UserModel

Namespace KKY_Tool_Revit.Infrastructure

    Public Module ExcelStyleHelper

        Public Enum RowStatus
            None = 0
            Info = 1
            Warning = 2
            [Error] = 3
        End Enum

        Private Class StyleSet
            Public HeaderStyle As ICellStyle
            Public InfoStyle As ICellStyle
            Public WarningStyle As ICellStyle
            Public ErrorStyle As ICellStyle
        End Class

        Private ReadOnly _wbStyles As New ConditionalWeakTable(Of IWorkbook, StyleSet)()

        Private Function GetStyleSet(wb As IWorkbook) As StyleSet
            Return _wbStyles.GetValue(wb, Function(_) CreateStyleSet(wb))
        End Function

        Private Function CreateStyleSet(wb As IWorkbook) As StyleSet
            Dim setx As New StyleSet()

            ' Header
            Dim header = wb.CreateCellStyle()
            Dim hf = wb.CreateFont()
            hf.IsBold = True
            header.SetFont(hf)
            header.Alignment = HorizontalAlignment.Center
            header.VerticalAlignment = VerticalAlignment.Center
            header.FillForegroundColor = IndexedColors.Grey25Percent.Index
            header.FillPattern = FillPattern.SolidForeground
            header.BorderBottom = BorderStyle.Thin
            header.BorderTop = BorderStyle.Thin
            header.BorderLeft = BorderStyle.Thin
            header.BorderRight = BorderStyle.Thin
            setx.HeaderStyle = header

            ' Info
            setx.InfoStyle = CreateFillStyle(wb, IndexedColors.LightCornflowerBlue.Index)

            ' Warning
            setx.WarningStyle = CreateFillStyle(wb, IndexedColors.LightYellow.Index)

            ' Error
            setx.ErrorStyle = CreateFillStyle(wb, IndexedColors.Rose.Index)

            Return setx
        End Function

        Private Function CreateFillStyle(wb As IWorkbook, fillColor As Short) As ICellStyle
            Dim s = wb.CreateCellStyle()
            s.WrapText = True
            s.VerticalAlignment = VerticalAlignment.Center
            s.FillForegroundColor = fillColor
            s.FillPattern = FillPattern.SolidForeground
            s.BorderBottom = BorderStyle.Thin
            s.BorderTop = BorderStyle.Thin
            s.BorderLeft = BorderStyle.Thin
            s.BorderRight = BorderStyle.Thin
            Return s
        End Function

        Public Function GetHeaderStyle(wb As IWorkbook) As ICellStyle
            Return GetStyleSet(wb).HeaderStyle
        End Function

        Public Function GetRowStyle(wb As IWorkbook, status As RowStatus) As ICellStyle
            Dim setx = GetStyleSet(wb)
            Select Case status
                Case RowStatus.Info
                    Return setx.InfoStyle
                Case RowStatus.Warning
                    Return setx.WarningStyle
                Case RowStatus.Error
                    Return setx.ErrorStyle
                Case Else
                    Return Nothing
            End Select
        End Function

        Public Sub ApplyStyleToRow(row As IRow, colCount As Integer, style As ICellStyle)
            If row Is Nothing OrElse style Is Nothing Then Return
            Dim lastCol As Integer = Math.Max(0, colCount - 1)

            For c As Integer = 0 To lastCol
                Dim cell = row.GetCell(c)
                If cell Is Nothing Then
                    cell = row.CreateCell(c)
                    cell.SetCellValue("")
                End If
                cell.CellStyle = style
            Next
        End Sub

    End Module

End Namespace
