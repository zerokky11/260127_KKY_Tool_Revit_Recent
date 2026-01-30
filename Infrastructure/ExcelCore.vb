Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Namespace KKY_Tool_Revit.Infrastructure

    Public Module ExcelCore

        Public Function PickAndSaveXlsx(title As String,
                                        table As DataTable,
                                        defaultFileName As String,
                                        Optional autoFit As Boolean = False,
                                        Optional progressKey As String = Nothing) As String

            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))

            Dim path = PickSavePath("Excel Workbook (*.xlsx)|*.xlsx", defaultFileName, title)
            If String.IsNullOrWhiteSpace(path) Then Return ""

            SaveXlsx(path, If(String.IsNullOrWhiteSpace(table.TableName), title, table.TableName), table, autoFit, sheetKey:=title, progressKey:=progressKey)
            Return path
        End Function

        Public Function PickAndSaveXlsxMulti(sheets As IList(Of KeyValuePair(Of String, DataTable)),
                                             defaultFileName As String,
                                             Optional autoFit As Boolean = False,
                                             Optional progressKey As String = Nothing) As String

            If sheets Is Nothing OrElse sheets.Count = 0 Then Throw New ArgumentException("Sheets is empty.", NameOf(sheets))

            Dim path = PickSavePath("Excel Workbook (*.xlsx)|*.xlsx", defaultFileName, "엑셀 저장")
            If String.IsNullOrWhiteSpace(path) Then Return ""

            SaveXlsxMulti(path, sheets, autoFit, progressKey)
            Return path
        End Function

        Public Sub SaveXlsx(filePath As String,
                            sheetName As String,
                            table As DataTable,
                            Optional autoFit As Boolean = False,
                            Optional sheetKey As String = Nothing,
                            Optional progressKey As String = Nothing)

            If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentNullException(NameOf(filePath))
            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))

            EnsureDir(filePath)

            Using wb As IWorkbook = New XSSFWorkbook()
                Dim safeSheet = NormalizeSheetName(If(sheetName, "Sheet1"))
                Dim sheet = wb.CreateSheet(safeSheet)

                WriteTableToSheet(wb, sheet, safeSheet, table, sheetKey, autoFit, progressKey)

                Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fs)
                End Using
            End Using
        End Sub

        Public Sub SaveXlsxMulti(filePath As String,
                                 sheets As IList(Of KeyValuePair(Of String, DataTable)),
                                 Optional autoFit As Boolean = False,
                                 Optional progressKey As String = Nothing)

            If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentNullException(NameOf(filePath))
            If sheets Is Nothing OrElse sheets.Count = 0 Then Throw New ArgumentException("Sheets is empty.", NameOf(sheets))

            EnsureDir(filePath)

            Using wb As IWorkbook = New XSSFWorkbook()
                Dim usedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For i As Integer = 0 To sheets.Count - 1
                    Dim name = If(sheets(i).Key, $"Sheet{i + 1}")
                    Dim table = sheets(i).Value
                    If table Is Nothing Then Continue For

                    Dim safe = MakeUniqueSheetName(NormalizeSheetName(name), usedNames)
                    usedNames.Add(safe)

                    Dim sheet = wb.CreateSheet(safe)
                    WriteTableToSheet(wb, sheet, safe, table, sheetKey:=name, autoFit:=autoFit, progressKey:=progressKey)
                Next

                Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                    wb.Write(fs)
                End Using
            End Using
        End Sub

        Public Sub SaveCsv(filePath As String, table As DataTable)
            If String.IsNullOrWhiteSpace(filePath) Then Throw New ArgumentNullException(NameOf(filePath))
            If table Is Nothing Then Throw New ArgumentNullException(NameOf(table))
            EnsureDir(filePath)

            Using sw As New StreamWriter(filePath, False, New UTF8Encoding(encoderShouldEmitUTF8Identifier:=True))
                ' header
                For c As Integer = 0 To table.Columns.Count - 1
                    If c > 0 Then sw.Write(",")
                    sw.Write(EscapeCsv(table.Columns(c).ColumnName))
                Next
                sw.WriteLine()

                ' rows
                For r As Integer = 0 To table.Rows.Count - 1
                    Dim dr = table.Rows(r)
                    For c As Integer = 0 To table.Columns.Count - 1
                        If c > 0 Then sw.Write(",")
                        Dim v = dr(c)
                        Dim s = If(v Is Nothing OrElse v Is DBNull.Value, "", v.ToString())
                        sw.Write(EscapeCsv(s))
                    Next
                    sw.WriteLine()
                Next
            End Using
        End Sub

        ' ---------------- internal ----------------

        Private Sub WriteTableToSheet(wb As IWorkbook,
                                      sheet As ISheet,
                                      sheetName As String,
                                      table As DataTable,
                                      sheetKey As String,
                                      autoFit As Boolean,
                                      progressKey As String)

            Dim colCount As Integer = table.Columns.Count
            If colCount = 0 Then Return

            ' header
            Dim headerRow = sheet.CreateRow(0)
            Dim headerStyle = ExcelStyleHelper.GetHeaderStyle(wb)

            For c As Integer = 0 To colCount - 1
                Dim cell = headerRow.CreateCell(c)
                cell.SetCellValue(table.Columns(c).ColumnName)
                cell.CellStyle = headerStyle
            Next

            sheet.CreateFreezePane(0, 1)

            Dim total As Integer = table.Rows.Count
            For r As Integer = 0 To total - 1
                Dim dr = table.Rows(r)
                Dim row = sheet.CreateRow(r + 1)

                For c As Integer = 0 To colCount - 1
                    WriteCell(row, c, dr(c))
                Next

                ' ---- 핵심: 저장하면서 행 상태를 판정해서 배경색 적용 ----
                Dim status = ExcelExportStyleRegistry.Resolve(If(sheetKey, sheetName), dr, table)
                If status <> ExcelStyleHelper.RowStatus.None Then
                    Dim style = ExcelStyleHelper.GetRowStyle(wb, status)
                    ExcelStyleHelper.ApplyStyleToRow(row, colCount, style)
                End If

                If (r Mod 200) = 0 Then
                    TryReportProgress(progressKey, r, total, sheetName)
                End If
            Next

            If autoFit Then
                Try
                    sheet.TrackAllColumnsForAutoSizing()
                Catch
                    ' ignore
                End Try
                For c As Integer = 0 To colCount - 1
                    Try
                        sheet.AutoSizeColumn(c)
                    Catch
                    End Try
                Next
            End If
        End Sub

        Private Sub WriteCell(row As IRow, colIndex As Integer, value As Object)
            Dim cell = row.CreateCell(colIndex)

            If value Is Nothing OrElse value Is DBNull.Value Then
                cell.SetCellValue("")
                Return
            End If

            If TypeOf value Is Boolean Then
                cell.SetCellValue(CBool(value))
                Return
            End If

            If TypeOf value Is DateTime Then
                cell.SetCellValue(DirectCast(value, DateTime))
                Return
            End If

            If TypeOf value Is Byte OrElse TypeOf value Is Short OrElse TypeOf value Is Integer OrElse
               TypeOf value Is Long OrElse TypeOf value Is Single OrElse TypeOf value Is Double OrElse
               TypeOf value Is Decimal Then

                Dim d As Double
                If Double.TryParse(value.ToString(), d) Then
                    cell.SetCellValue(d)
                Else
                    cell.SetCellValue(value.ToString())
                End If
                Return
            End If

            cell.SetCellValue(value.ToString())
        End Sub

        Private Function PickSavePath(filter As String, defaultFileName As String, title As String) As String
            Using dlg As New SaveFileDialog()
                dlg.Filter = filter
                dlg.Title = If(String.IsNullOrWhiteSpace(title), "저장", title)
                dlg.FileName = If(String.IsNullOrWhiteSpace(defaultFileName), "export.xlsx", defaultFileName)
                dlg.RestoreDirectory = True
                If dlg.ShowDialog() <> DialogResult.OK Then Return ""
                Return dlg.FileName
            End Using
        End Function

        Private Sub EnsureDir(filePath As String)
            Dim dir = Path.GetDirectoryName(filePath)
            If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                Directory.CreateDirectory(dir)
            End If
        End Sub

        Private Function NormalizeSheetName(name As String) As String
            Dim s = If(name, "Sheet1").Trim()
            If s.Length = 0 Then s = "Sheet1"

            ' Excel 금지 문자: : \ / ? * [ ]
            Dim bad = New Char() {":"c, "\"c, "/"c, "?"c, "*"c, "["c, "]"c}
            For Each ch In bad
                s = s.Replace(ch, "_"c)
            Next

            If s.Length > 31 Then s = s.Substring(0, 31)
            Return s
        End Function

        Private Function MakeUniqueSheetName(baseName As String, used As HashSet(Of String)) As String
            Dim s = baseName
            Dim i As Integer = 1
            While used.Contains(s)
                Dim suffix = $"({i})"
                Dim cut = Math.Min(31 - suffix.Length, baseName.Length)
                s = baseName.Substring(0, cut) & suffix
                i += 1
            End While
            Return s
        End Function

        Private Function EscapeCsv(s As String) As String
            If s Is Nothing Then Return ""
            Dim needs = s.Contains(","c) OrElse s.Contains(""""c) OrElse s.Contains(vbCr) OrElse s.Contains(vbLf)
            Dim t = s.Replace("""", """""")
            If needs Then Return $"""{t}"""
            Return t
        End Function

        ' progressKey는 UiBridge에서 "hub:multi-progress" 같은 채널로 쓰는 구조가 있어서:contentReference[oaicite:6]{index=6}
        ' 여기서는 있으면 최대한 조용히 반영(리플렉션)하고, 없어도 기능은 정상 동작하게 처리
        Private Sub TryReportProgress(progressKey As String, current As Integer, total As Integer, sheetName As String)
            If String.IsNullOrWhiteSpace(progressKey) Then Return
            Try
                Dim t = Type.GetType("KKY_Tool_Revit.UI.Hub.ExcelProgressReporter, " & GetType(ExcelCore).Assembly.FullName, throwOnError:=False)
                If t Is Nothing Then Return

                Dim mi = t.GetMethod("Report", Reflection.BindingFlags.Public Or Reflection.BindingFlags.Static)
                If mi Is Nothing Then Return

                Dim percent As Double = 0
                If total > 0 Then percent = (CDbl(current) / CDbl(total)) * 100.0R
                mi.Invoke(Nothing, New Object() {progressKey, percent, $"Exporting {sheetName}...", $"{current}/{total}"})
            Catch
            End Try
        End Sub

    End Module

End Namespace
