Imports System
Imports System.Data

Namespace Exports

    Public Module ConnectorExport

        Public Function SaveWithDialog(resultTable As DataTable) As String
            If resultTable Is Nothing Then Return String.Empty

            Dim outPath As String = Global.KKY_Tool_Revit.Infrastructure.ExcelCore.PickAndSaveXlsx(
                "Connector Diagnostics",
                resultTable,
                "ConnectorDiagnostics.xlsx"
            )

            If String.IsNullOrWhiteSpace(outPath) Then Return outPath

            Try
                Global.KKY_Tool_Revit.Infrastructure.ExcelExportStyleRegistry.ApplyStylesForKey("connector", outPath)
            Catch
            End Try

            Return outPath
        End Function

        Public Sub Save(outPath As String, resultTable As DataTable)
            If String.IsNullOrWhiteSpace(outPath) Then Exit Sub
            If resultTable Is Nothing Then Exit Sub

            Global.KKY_Tool_Revit.Infrastructure.ExcelCore.SaveXlsx(outPath, "Connector Diagnostics", resultTable)

            Try
                Global.KKY_Tool_Revit.Infrastructure.ExcelExportStyleRegistry.ApplyStylesForKey("connector", outPath)
            Catch
            End Try
        End Sub

    End Module

End Namespace
