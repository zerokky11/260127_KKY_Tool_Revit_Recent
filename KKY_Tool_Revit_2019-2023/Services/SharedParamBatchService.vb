Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Web.Script.Serialization
Imports System.Windows.Forms
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports KKY_Tool_Revit.Infrastructure
Imports RevitApp = Autodesk.Revit.ApplicationServices.Application
Imports WinForms = System.Windows.Forms

Namespace Services

    Public NotInheritable Class SharedParamBatchService

        Private Sub New()
        End Sub

        Public Class RunSummary
            Public Property OkCount As Integer
            Public Property FailCount As Integer
            Public Property SkipCount As Integer
        End Class

        Public Class LogEntry
            Public Property Level As String = String.Empty
            Public Property [File] As String = String.Empty
            Public Property Message As String = String.Empty
        End Class

        Public Class RunResult
            Public Property Ok As Boolean
            Public Property Message As String = String.Empty
            Public Property Summary As RunSummary
            Public Property Logs As List(Of LogEntry)
            Public Property LogTextPath As String = String.Empty
        End Class

        Public Class ParamGroupOption
            Public Property Key As String = String.Empty
            Public Property Id As String = String.Empty
            Public Property Label As String = String.Empty
        End Class

        Friend Class BatchSettings
            Public Property RvtPaths As List(Of String)
            Public Property Parameters As List(Of ParamToBind)
            Public Property CloseAllWorksetsOnOpen As Boolean
            Public Property SyncComment As String
        End Class

        Friend Class CategoryRef
            Public Property IdInt As Integer
            Public Property Path As String
            Public Property Name As String

            Public Function Clone() As CategoryRef
                Return New CategoryRef() With {.IdInt = Me.IdInt, .Path = Me.Path, .Name = Me.Name}
            End Function
        End Class

        Friend Class ParamBindingSettings
            Public Property IsInstanceBinding As Boolean
            Public Property ParamGroup As BuiltInParameterGroup
            Public Property AllowVaryBetweenGroups As Boolean
            Public Property Categories As List(Of CategoryRef)

            Public Function CloneDeep() As ParamBindingSettings
                Dim c As New ParamBindingSettings()
                c.IsInstanceBinding = Me.IsInstanceBinding
                c.ParamGroup = Me.ParamGroup
                c.AllowVaryBetweenGroups = Me.AllowVaryBetweenGroups

                Dim src As List(Of CategoryRef) = If(Me.Categories, New List(Of CategoryRef)())
                Dim dst As New List(Of CategoryRef)()
                For Each x As CategoryRef In src
                    If x IsNot Nothing Then dst.Add(x.Clone())
                Next
                c.Categories = dst

                Return c
            End Function
        End Class

        Friend Class ParamToBind
            Public Property GroupName As String
            Public Property ParamName As String
            Public Property GuidValue As Guid
            Public Property ParamTypeLabel As String
            Public Property Description As String
            Public Property Settings As ParamBindingSettings

            Public ReadOnly Property GuidString As String
                Get
                    Return GuidValue.ToString()
                End Get
            End Property
        End Class

        Friend Class CategoryTreeItem
            Public Property IdInt As Integer
            Public Property Name As String
            Public Property Path As String
            Public Property CatType As CategoryType
            Public Property IsBindable As Boolean
            Public Property Children As List(Of CategoryTreeItem)

            Public Sub New()
                Children = New List(Of CategoryTreeItem)()
            End Sub
        End Class

        Private Class CategoryMaps
            Public Sub New()
                ById = New Dictionary(Of Integer, Category)()
                ByPath = New Dictionary(Of String, Category)(StringComparer.OrdinalIgnoreCase)
            End Sub
            Public Property ById As Dictionary(Of Integer, Category)
            Public Property ByPath As Dictionary(Of String, Category)
        End Class

        Public Shared Function Init(uiapp As UIApplication) As Object
            If uiapp Is Nothing OrElse uiapp.ActiveUIDocument Is Nothing OrElse uiapp.ActiveUIDocument.Document Is Nothing Then
                Return New With {
                    .ok = False,
                    .message = "활성 프로젝트 문서가 필요합니다.",
                    .spFilePath = String.Empty,
                    .groups = New List(Of String)(),
                    .defsByGroup = New Dictionary(Of String, Object)(),
                    .categoryTree = New List(Of Object)(),
                    .paramGroups = New List(Of Object)()
                }
            End If

            Dim baseDoc As Document = uiapp.ActiveUIDocument.Document
            If baseDoc Is Nothing OrElse baseDoc.IsFamilyDocument Then
                Return New With {
                    .ok = False,
                    .message = "활성 프로젝트 문서가 필요합니다. (패밀리 문서 불가)",
                    .spFilePath = String.Empty,
                    .groups = New List(Of String)(),
                    .defsByGroup = New Dictionary(Of String, Object)(),
                    .categoryTree = New List(Of Object)(),
                    .paramGroups = New List(Of Object)()
                }
            End If

            Dim app As RevitApp = uiapp.Application
            Dim spFilePath As String = app.SharedParametersFilename
            If String.IsNullOrWhiteSpace(spFilePath) OrElse Not File.Exists(spFilePath) Then
                Return New With {
                    .ok = False,
                    .message = "현재 Revit에 설정된 Shared Parameter TXT를 찾을 수 없습니다.",
                    .spFilePath = If(spFilePath, String.Empty),
                    .groups = New List(Of String)(),
                    .defsByGroup = New Dictionary(Of String, Object)(),
                    .categoryTree = New List(Of Object)(),
                    .paramGroups = New List(Of Object)()
                }
            End If

            Dim defFile As DefinitionFile = Nothing
            Try
                defFile = app.OpenSharedParameterFile()
            Catch ex As Exception
                Return New With {
                    .ok = False,
                    .message = "Shared Parameter 파일을 열 수 없습니다: " & ex.Message,
                    .spFilePath = spFilePath,
                    .groups = New List(Of String)(),
                    .defsByGroup = New Dictionary(Of String, Object)(),
                    .categoryTree = New List(Of Object)(),
                    .paramGroups = New List(Of Object)()
                }
            End Try

            If defFile Is Nothing Then
                Return New With {
                    .ok = False,
                    .message = "Shared Parameter 파일을 열 수 없습니다.",
                    .spFilePath = spFilePath,
                    .groups = New List(Of String)(),
                    .defsByGroup = New Dictionary(Of String, Object)(),
                    .categoryTree = New List(Of Object)(),
                    .paramGroups = New List(Of Object)()
                }
            End If

            Dim groupNames As New List(Of String)()
            Dim defsByGroup As New Dictionary(Of String, List(Of Object))(StringComparer.OrdinalIgnoreCase)
            For Each g As DefinitionGroup In defFile.Groups
                If g Is Nothing Then Continue For

                groupNames.Add(g.Name)
                Dim defs As New List(Of Object)()
                For Each d As Definition In g.Definitions
                    Dim ext As ExternalDefinition = TryCast(d, ExternalDefinition)
                    If ext Is Nothing Then Continue For

                    defs.Add(New With {
                        .name = ext.Name,
                        .guid = ext.GUID.ToString("D"),
                        .paramTypeLabel = GetExternalDefinitionTypeLabel(ext),
                        .desc = If(ext.Description, String.Empty)
                    })
                Next

                defsByGroup(g.Name) = defs
            Next

            Dim categoryTreeRoots As List(Of CategoryTreeItem) = BuildCategoryTree(baseDoc)
            If categoryTreeRoots.Count = 0 Then
                Return New With {
                    .ok = False,
                    .message = "바인딩 가능한 카테고리를 찾지 못했습니다.",
                    .spFilePath = spFilePath,
                    .groups = groupNames,
                    .defsByGroup = defsByGroup.ToDictionary(Function(kv) kv.Key, Function(kv) DirectCast(kv.Value, Object)),
                    .categoryTree = New List(Of Object)(),
                    .paramGroups = BuildParamGroupOptions()
                }
            End If

            Dim categoryDto As List(Of Object) = categoryTreeRoots.Select(Function(x) ToCategoryDto(x)).ToList()
            Dim defsDto As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            For Each kv In defsByGroup
                defsDto(kv.Key) = kv.Value
            Next

            Return New With {
                .ok = True,
                .message = String.Empty,
                .spFilePath = spFilePath,
                .groups = groupNames,
                .defsByGroup = defsDto,
                .categoryTree = categoryDto,
                .paramGroups = BuildParamGroupOptions()
            }
        End Function

        Public Shared Function BrowseRvts() As Object
            Using dlg As New OpenFileDialog()
                dlg.Filter = "Revit Project (*.rvt)|*.rvt"
                dlg.Multiselect = True
                dlg.RestoreDirectory = True
                Dim result = dlg.ShowDialog()
                If result <> DialogResult.OK Then
                    Return New With {.ok = False, .message = "파일 선택이 취소되었습니다."}
                End If

                Dim paths As New List(Of String)()
                If dlg.FileNames IsNot Nothing Then
                    For Each p In dlg.FileNames
                        If Not String.IsNullOrWhiteSpace(p) Then
                            paths.Add(p)
                        End If
                    Next
                End If

                Return New With {.ok = True, .rvtPaths = paths}
            End Using
        End Function

        Public Shared Function BrowseRvtFolder() As Object
            Dim selectedPath As String = ""
            Using dlg As New FolderBrowserDialog()
                dlg.Description = "RVT 폴더 선택"
                dlg.ShowNewFolderButton = False
                Dim result = dlg.ShowDialog()
                If result <> DialogResult.OK Then
                    Return New With {.ok = False, .message = "폴더 선택이 취소되었습니다."}
                End If
                selectedPath = dlg.SelectedPath
            End Using

            If String.IsNullOrWhiteSpace(selectedPath) OrElse Not Directory.Exists(selectedPath) Then
                Return New With {.ok = False, .message = "선택된 폴더를 찾을 수 없습니다."}
            End If

            Dim paths As List(Of String) = CollectRvtFiles(selectedPath)
            Return New With {.ok = True, .rvtPaths = paths, .fromFolder = True}
        End Function

        Public Shared Function Run(uiapp As UIApplication, payloadJson As String, progress As IProgress(Of Object)) As Object
            If uiapp Is Nothing Then
                Return New RunResult With {.Ok = False, .Message = "Revit UIApplication이 없습니다."}
            End If

            Dim serializer As New JavaScriptSerializer()
            Dim payload As Dictionary(Of String, Object) = Nothing
            Try
                payload = serializer.Deserialize(Of Dictionary(Of String, Object))(payloadJson)
            Catch ex As Exception
                Return New RunResult With {.Ok = False, .Message = "설정 JSON 파싱 실패: " & ex.Message}
            End Try

            Dim spFilePath As String = TryGetString(payload, "spFilePath")
            If String.IsNullOrWhiteSpace(spFilePath) Then
                spFilePath = uiapp.Application.SharedParametersFilename
            End If

            Dim settings As BatchSettings = ParseBatchSettings(payload)
            Dim validationError As String = ValidateBatchSettings(settings, spFilePath)
            If validationError <> "" Then
                Return New RunResult With {.Ok = False, .Message = validationError}
            End If

            Dim logs As New List(Of LogEntry)()
            Dim logLines As New List(Of String)()
            Dim okCount As Integer = 0
            Dim failCount As Integer = 0
            Dim skipCount As Integer = 0

            Dim app As RevitApp = uiapp.Application
            Dim originalSpFile As String = app.SharedParametersFilename

            Try
                ReportProgress(progress, 0, settings.RvtPaths.Count, "Shared Parameter 정의 로드 중...")

                app.SharedParametersFilename = spFilePath
                Dim defFile As DefinitionFile = app.OpenSharedParameterFile()
                If defFile Is Nothing Then
                    Return New RunResult With {.Ok = False, .Message = "Shared Parameter 파일을 열 수 없습니다: " & spFilePath}
                End If

                Dim extDefByGuid As New Dictionary(Of Guid, ExternalDefinition)()
                For Each p As ParamToBind In settings.Parameters
                    Dim extDef As ExternalDefinition = FindExternalDefinitionByGuid(defFile, p.GuidValue)
                    If extDef Is Nothing Then
                        Return New RunResult With {
                            .Ok = False,
                            .Message = "Shared Parameter 정의를 찾지 못했습니다: " & p.ParamName & " (" & p.GuidValue.ToString() & ")"
                        }
                    End If
                    If Not extDefByGuid.ContainsKey(p.GuidValue) Then
                        extDefByGuid.Add(p.GuidValue, extDef)
                    End If
                Next

                Dim total As Integer = settings.RvtPaths.Count
                For i As Integer = 0 To total - 1
                    Dim rvtPath As String = settings.RvtPaths(i)
                    ReportProgress(progress, i + 1, total, "처리 중: " & rvtPath)

                    If String.IsNullOrWhiteSpace(rvtPath) OrElse Not File.Exists(rvtPath) Then
                        skipCount += 1
                        AddLog(logs, logLines, "SKIP", rvtPath, "파일 없음")
                        Continue For
                    End If

                    If IsAlreadyOpen(app, rvtPath) Then
                        skipCount += 1
                        AddLog(logs, logLines, "SKIP", rvtPath, "이미 열려있음")
                        Continue For
                    End If

                    Dim doc As Document = Nothing
                    Try
                        doc = OpenProjectDocument(app, rvtPath, settings.CloseAllWorksetsOnOpen)

                        If doc Is Nothing Then
                            failCount += 1
                            AddLog(logs, logLines, "FAIL", rvtPath, "열기 실패(Unknown)")
                            Continue For
                        End If

                        If doc.IsFamilyDocument Then
                            skipCount += 1
                            AddLog(logs, logLines, "SKIP", rvtPath, "패밀리 문서")
                            SafeClose(doc, False)
                            Continue For
                        End If

                        Dim perDocNotes As String = ""
                        Dim applied As Boolean = ApplyAllSharedParameterBindings(doc, app, extDefByGuid, settings.Parameters, perDocNotes)

                        If Not applied Then
                            failCount += 1
                            AddLog(logs, logLines, "FAIL", rvtPath, perDocNotes)
                            SafeClose(doc, False)
                            Continue For
                        End If

                        If Not String.IsNullOrWhiteSpace(perDocNotes) Then
                            Dim lines As String() = perDocNotes.Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                            For Each ln As String In lines
                                AddLog(logs, logLines, "WARN", rvtPath, ln)
                            Next
                        End If

                        If doc.IsWorkshared Then
                            Dim syncLog As String = ""
                            Dim synced As Boolean = SyncWithCentral(doc, settings.SyncComment, syncLog)
                            If Not synced Then
                                failCount += 1
                                AddLog(logs, logLines, "FAIL", rvtPath, "Sync 실패 :: " & syncLog)
                                SafeClose(doc, False)
                                Continue For
                            End If

                            okCount += 1
                            AddLog(logs, logLines, "OK", rvtPath, "적용(" & settings.Parameters.Count & "개) + Sync 완료 (Comment: " & settings.SyncComment & ")")
                            SafeClose(doc, False)
                        Else
                            doc.Save()
                            okCount += 1
                            AddLog(logs, logLines, "OK", rvtPath, "적용(" & settings.Parameters.Count & "개) + Save 완료")
                            SafeClose(doc, False)
                        End If

                    Catch ex As Exception
                        failCount += 1
                        AddLog(logs, logLines, "FAIL", rvtPath, "예외: " & ex.Message)
                        If doc IsNot Nothing Then
                            SafeClose(doc, False)
                        End If
                    End Try
                Next

                Dim logPath As String = Path.Combine(Path.GetTempPath(),
                                                    "KKY_SharedParamBatch_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt")
                Try
                    File.WriteAllLines(logPath, logLines.ToArray(), Encoding.UTF8)
                Catch
                End Try

                Dim summary As New RunSummary With {
                    .OkCount = okCount,
                    .FailCount = failCount,
                    .SkipCount = skipCount
                }

                ReportProgress(progress, total, total, "완료")

                Return New RunResult With {
                    .Ok = True,
                    .Message = "완료",
                    .Summary = summary,
                    .Logs = logs,
                    .LogTextPath = logPath
                }

            Catch exTop As Exception
                Return New RunResult With {.Ok = False, .Message = "치명적 예외: " & exTop.Message}

            Finally
                Try
                    app.SharedParametersFilename = originalSpFile
                Catch
                End Try
            End Try
        End Function

        Public Shared Function ExportExcel(resultDtoOrLogsJson As String) As Object
            If String.IsNullOrWhiteSpace(resultDtoOrLogsJson) Then
                Return New With {.ok = False, .message = "로그 데이터가 없습니다."}
            End If

            Dim serializer As New JavaScriptSerializer()
            Dim payload As Dictionary(Of String, Object) = Nothing
            Try
                payload = serializer.Deserialize(Of Dictionary(Of String, Object))(resultDtoOrLogsJson)
            Catch ex As Exception
                Return New With {.ok = False, .message = "로그 JSON 파싱 실패: " & ex.Message}
            End Try

            Dim logs As List(Of LogEntry) = ParseLogEntries(payload)
            If logs Is Nothing OrElse logs.Count = 0 Then
                Return New With {.ok = False, .message = "내보낼 로그가 없습니다."}
            End If

            Dim rvtPaths As List(Of String) = ParseStringList(payload, "rvtPaths")
            Dim parameters As List(Of ParamToBind) = ParseParamList(payload)
            If rvtPaths.Count = 0 OrElse parameters.Count = 0 Then
                Return New With {.ok = False, .message = "내보낼 RVT/파라미터 정보가 없습니다."}
            End If

            Dim doAutoFit As Boolean = False
            Try
                Dim mode As String = TryGetString(payload, "excelMode")
                If String.Equals(mode, "normal", StringComparison.OrdinalIgnoreCase) AndAlso logs.Count <= 30000 Then
                    doAutoFit = True
                End If
            Catch
                doAutoFit = False
            End Try

            Dim dt As New DataTable("SharedParamBatch")
            Dim headers As String() = {
                "파일명",
                "파라미터명",
                "GUID",
                "바인딩",
                "파라미터 그룹",
                "성공여부",
                "메시지"
            }
            For Each h As String In headers
                dt.Columns.Add(h)
            Next

            If Not ValidateExportSchema(dt, headers) Then
                Return New With {.ok = False, .message = "엑셀 스키마가 올바르지 않습니다."}
            End If

            Dim statusMap As Dictionary(Of String, LogEntry) = BuildStatusMap(logs)

            For Each rvtPath As String In rvtPaths
                Dim rvtFileName As String = If(String.IsNullOrWhiteSpace(rvtPath), "", Path.GetFileName(rvtPath))
                Dim statusEntry As LogEntry = Nothing
                Dim status As String = "SKIP"
                Dim message As String = ""

                If Not String.IsNullOrWhiteSpace(rvtPath) AndAlso statusMap.TryGetValue(rvtPath, statusEntry) Then
                    status = NormalizeStatus(statusEntry.Level)
                    If String.Equals(status, "FAIL", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(status, "SKIP", StringComparison.OrdinalIgnoreCase) Then
                        message = If(statusEntry.Message, String.Empty)
                    End If
                End If

                For Each p As ParamToBind In parameters
                    Dim row = dt.NewRow()
                    row("파일명") = rvtFileName
                    row("파라미터명") = If(p.ParamName, String.Empty)
                    row("GUID") = p.GuidString
                    row("바인딩") = If(p.Settings IsNot Nothing AndAlso p.Settings.IsInstanceBinding, "Instance", "Type")
                    row("파라미터 그룹") = GetParamGroupLabel(If(p.Settings IsNot Nothing, p.Settings.ParamGroup, BuiltInParameterGroup.INVALID))
                    row("성공여부") = status
                    row("메시지") = message
                    dt.Rows.Add(row)
                Next
            Next

            Dim fileName As String = $"SharedParamBatch_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            Dim saved As String = ExcelCore.PickAndSaveXlsx("Logs", dt, fileName, doAutoFit, "sharedparambatch:progress")
            If String.IsNullOrWhiteSpace(saved) Then
                Return New With {.ok = False, .message = "엑셀 내보내기가 취소되었습니다."}
            End If

            Return New With {.ok = True, .filePath = saved}
        End Function

        Private Shared Function CollectRvtFiles(rootPath As String) As List(Of String)
            Dim results As New List(Of String)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim stack As New Stack(Of String)()
            If Not String.IsNullOrWhiteSpace(rootPath) Then stack.Push(rootPath)

            While stack.Count > 0
                Dim current As String = stack.Pop()
                Dim subDirs As IEnumerable(Of String) = Enumerable.Empty(Of String)()
                Dim files As IEnumerable(Of String) = Enumerable.Empty(Of String)()

                Try
                    subDirs = Directory.EnumerateDirectories(current)
                Catch
                End Try

                For Each dir As String In subDirs
                    stack.Push(dir)
                Next

                Try
                    files = Directory.EnumerateFiles(current, "*.rvt")
                Catch
                End Try

                For Each f As String In files
                    If IsBackupRvt(f) Then Continue For
                    If Not seen.Contains(f) Then
                        seen.Add(f)
                        results.Add(f)
                    End If
                Next
            End While

            Return results
        End Function

        Private Shared Function IsBackupRvt(filePath As String) As Boolean
            If String.IsNullOrWhiteSpace(filePath) Then Return False
            Dim nameOnly As String = Path.GetFileNameWithoutExtension(filePath)
            If String.IsNullOrWhiteSpace(nameOnly) Then Return False
            Dim idx As Integer = nameOnly.LastIndexOf("."c)
            If idx < 0 OrElse idx >= nameOnly.Length - 1 Then Return False
            Dim suffix As String = nameOnly.Substring(idx + 1)
            If suffix.Length <> 4 Then Return False
            Return suffix.All(Function(ch) Char.IsDigit(ch))
        End Function

        Private Shared Function BuildStatusMap(logs As List(Of LogEntry)) As Dictionary(Of String, LogEntry)
            Dim map As New Dictionary(Of String, LogEntry)(StringComparer.OrdinalIgnoreCase)
            If logs Is Nothing Then Return map

            For Each l As LogEntry In logs
                If l Is Nothing Then Continue For
                Dim path As String = If(l.File, String.Empty)
                If String.IsNullOrWhiteSpace(path) Then Continue For
                Dim status As String = NormalizeStatus(l.Level)
                If status = "" Then Continue For

                Dim existing As LogEntry = Nothing
                If map.TryGetValue(path, existing) Then
                    Dim currentRank As Integer = GetStatusRank(status)
                    Dim existingRank As Integer = GetStatusRank(existing.Level)
                    If currentRank > existingRank Then
                        map(path) = New LogEntry With {.Level = status, .File = path, .Message = l.Message}
                    End If
                Else
                    map.Add(path, New LogEntry With {.Level = status, .File = path, .Message = l.Message})
                End If
            Next

            Return map
        End Function

        Private Shared Function NormalizeStatus(level As String) As String
            If String.IsNullOrWhiteSpace(level) Then Return ""
            Dim upper As String = level.Trim().ToUpperInvariant()
            Select Case upper
                Case "OK", "FAIL", "SKIP"
                    Return upper
                Case Else
                    Return ""
            End Select
        End Function

        Private Shared Function GetStatusRank(level As String) As Integer
            Dim normalized As String = NormalizeStatus(level)
            Select Case normalized
                Case "FAIL"
                    Return 3
                Case "SKIP"
                    Return 2
                Case "OK"
                    Return 1
                Case Else
                    Return 0
            End Select
        End Function

        Private Shared Function GetParamGroupLabel(groupValue As BuiltInParameterGroup) As String
            If groupValue = BuiltInParameterGroup.INVALID Then Return ""
            Try
                Return LabelUtils.GetLabelFor(groupValue)
            Catch
                Return groupValue.ToString()
            End Try
        End Function

        Private Shared Function ValidateExportSchema(dt As DataTable, headers As String()) As Boolean
            If dt Is Nothing OrElse headers Is Nothing Then Return False
            If dt.Columns.Count <> headers.Length Then Return False
            For i As Integer = 0 To headers.Length - 1
                If Not String.Equals(dt.Columns(i).ColumnName, headers(i), StringComparison.Ordinal) Then
                    Return False
                End If
            Next
            Return True
        End Function

        Private Shared Sub ReportProgress(progress As IProgress(Of Object), stepIndex As Integer, total As Integer, text As String)
            If progress Is Nothing Then Return
            progress.Report(New With {
                .step = stepIndex,
                .total = total,
                .text = text
            })
        End Sub

        Private Shared Function GetExternalDefinitionTypeLabel(ed As Autodesk.Revit.DB.ExternalDefinition) As String
            If ed Is Nothing Then Return ""

            ' (A) Old API: ParameterType
            Try
                Dim pi = ed.GetType().GetProperty("ParameterType")
                If pi IsNot Nothing Then
                    Dim ptObj As Object = pi.GetValue(ed, Nothing)
                    If ptObj IsNot Nothing Then
                        Dim mi = GetType(Autodesk.Revit.DB.LabelUtils).GetMethod("GetLabelFor", New Type() {ptObj.GetType()})
                        If mi IsNot Nothing Then
                            Return Convert.ToString(mi.Invoke(Nothing, New Object() {ptObj}))
                        End If
                        Return ptObj.ToString()
                    End If
                End If
            Catch
                ' ignore
            End Try

            ' (B) New API: GetDataType() -> ForgeTypeId (do not type-reference)
            Try
                Dim miGetDt = ed.GetType().GetMethod("GetDataType", Type.EmptyTypes)
                If miGetDt IsNot Nothing Then
                    Dim dtObj As Object = miGetDt.Invoke(ed, Nothing)
                    If dtObj IsNot Nothing Then
                        Dim miLabel = GetType(Autodesk.Revit.DB.LabelUtils).GetMethod("GetLabelForSpec", New Type() {dtObj.GetType()})
                        If miLabel Is Nothing Then
                            miLabel = GetType(Autodesk.Revit.DB.LabelUtils).GetMethod("GetLabelFor", New Type() {dtObj.GetType()})
                        End If
                        If miLabel IsNot Nothing Then
                            Return Convert.ToString(miLabel.Invoke(Nothing, New Object() {dtObj}))
                        End If

                        Dim piTypeId = dtObj.GetType().GetProperty("TypeId")
                        If piTypeId IsNot Nothing Then
                            Return Convert.ToString(piTypeId.GetValue(dtObj, Nothing))
                        End If

                        Return dtObj.ToString()
                    End If
                End If
            Catch
                ' ignore
            End Try

            ' final fallback (never empty if possible)
            Try
                Return ed.ToString()
            Catch
                Return ""
            End Try
        End Function

        Private Shared Function BuildParamGroupOptions() As List(Of ParamGroupOption)
            Dim opts As New List(Of ParamGroupOption)()
            Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each v As BuiltInParameterGroup In [Enum].GetValues(GetType(BuiltInParameterGroup))
                If v = BuiltInParameterGroup.INVALID Then Continue For
                Dim key As String = v.ToString()
                If seen.Contains(key) Then Continue For
                seen.Add(key)

                Dim label As String = key
                Try
                    label = LabelUtils.GetLabelFor(v)
                Catch
                    label = key
                End Try

                opts.Add(New ParamGroupOption With {.Key = key, .Id = key, .Label = label})
            Next

            Dim hasOther As Boolean = opts.Any(Function(o) String.Equals(o.Id, "PG_OTHER", StringComparison.OrdinalIgnoreCase) OrElse String.Equals(o.Key, "PG_OTHER", StringComparison.OrdinalIgnoreCase))
            If Not hasOther Then
                Try
                    Dim otherKey As String = "PG_OTHER"
                    Dim otherLabel As String = "Other"
                    Try
                        Dim otherEnum As BuiltInParameterGroup = CType([Enum].Parse(GetType(BuiltInParameterGroup), otherKey), BuiltInParameterGroup)
                        otherLabel = LabelUtils.GetLabelFor(otherEnum)
                    Catch
                        otherLabel = "Other"
                    End Try
                    opts.Add(New ParamGroupOption With {.Key = otherKey, .Id = otherKey, .Label = otherLabel})
                Catch
                End Try
            End If

            Return opts.OrderBy(Function(o) o.Label).ToList()
        End Function

        Private Shared Sub AddLog(logs As List(Of LogEntry), lines As List(Of String), level As String, filePath As String, message As String)
            logs.Add(New LogEntry With {
                .Level = level,
                .File = filePath,
                .Message = message
            })
            lines.Add("[" & level & "] " & filePath & " :: " & message)
        End Sub

        Private Shared Function ParseBatchSettings(payload As Dictionary(Of String, Object)) As BatchSettings
            Dim settings As New BatchSettings()
            settings.RvtPaths = ParseStringList(payload, "rvtPaths")
            settings.CloseAllWorksetsOnOpen = ParseBoolean(payload, "closeAllWorksetsOnOpen", True)
            settings.SyncComment = TryGetString(payload, "syncComment")
            settings.Parameters = ParseParamList(payload)
            Return settings
        End Function

        Private Shared Function ParseParamList(payload As Dictionary(Of String, Object)) As List(Of ParamToBind)
            Dim list As New List(Of ParamToBind)()
            If payload Is Nothing OrElse Not payload.ContainsKey("parameters") Then Return list

            Dim raw = payload("parameters")
            Dim arr = TryCast(raw, System.Collections.IEnumerable)
            If arr Is Nothing OrElse TypeOf raw Is String Then Return list

            For Each o In arr
                Dim d = ParsePayloadDict(o)
                Dim p As New ParamToBind()
                p.GroupName = TryGetString(d, "groupName")
                p.ParamName = TryGetString(d, "paramName")
                p.ParamTypeLabel = TryGetString(d, "paramTypeLabel")
                p.Description = TryGetString(d, "description")

                Dim guidRaw As String = TryGetString(d, "guid")
                Dim gv As Guid
                If Guid.TryParse(guidRaw, gv) Then
                    p.GuidValue = gv
                End If

                Dim settingsObj As Object = Nothing
                If d.TryGetValue("settings", settingsObj) Then
                    p.Settings = ParseBindingSettings(ParsePayloadDict(settingsObj))
                Else
                    p.Settings = New ParamBindingSettings()
                End If

                list.Add(p)
            Next

            Return list
        End Function

        Private Shared Function ParseBindingSettings(dict As Dictionary(Of String, Object)) As ParamBindingSettings
            Dim s As New ParamBindingSettings()
            If dict Is Nothing Then Return s

            s.IsInstanceBinding = ParseBoolean(dict, "isInstanceBinding", True)
            s.AllowVaryBetweenGroups = ParseBoolean(dict, "allowVaryBetweenGroups", False)
            s.ParamGroup = ParseParamGroup(dict, "paramGroup")
            s.Categories = ParseCategoryRefs(dict)
            Return s
        End Function

        Private Shared Function ParseCategoryRefs(dict As Dictionary(Of String, Object)) As List(Of CategoryRef)
            Dim list As New List(Of CategoryRef)()
            If dict Is Nothing OrElse Not dict.ContainsKey("categories") Then Return list

            Dim raw = dict("categories")
            Dim arr = TryCast(raw, System.Collections.IEnumerable)
            If arr Is Nothing OrElse TypeOf raw Is String Then Return list

            For Each o In arr
                Dim d = ParsePayloadDict(o)
                Dim c As New CategoryRef()
                c.Name = TryGetString(d, "name")
                c.Path = TryGetString(d, "path")
                Dim idRaw As String = TryGetString(d, "idInt")
                Dim idValue As Integer
                If Integer.TryParse(idRaw, idValue) Then c.IdInt = idValue
                list.Add(c)
            Next

            Return list
        End Function

        Private Shared Function ParseLogEntries(payload As Dictionary(Of String, Object)) As List(Of LogEntry)
            Dim list As New List(Of LogEntry)()
            If payload Is Nothing OrElse Not payload.ContainsKey("logs") Then Return list

            Dim raw = payload("logs")
            Dim arr = TryCast(raw, System.Collections.IEnumerable)
            If arr Is Nothing OrElse TypeOf raw Is String Then Return list

            For Each o In arr
                Dim d = ParsePayloadDict(o)
                list.Add(New LogEntry With {
                    .Level = TryGetString(d, "level"),
                    .File = TryGetString(d, "file"),
                    .Message = TryGetString(d, "msg")
                })
            Next

            Return list
        End Function

        Private Shared Function ParsePayloadDict(payload As Object) As Dictionary(Of String, Object)
            Dim dict = TryCast(payload, Dictionary(Of String, Object))
            If dict IsNot Nothing Then Return dict

            Dim result As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            If payload Is Nothing Then Return result

            Dim t = payload.GetType()
            For Each p In t.GetProperties()
                Try
                    result(p.Name) = p.GetValue(payload, Nothing)
                Catch
                End Try
            Next
            Return result
        End Function

        Private Shared Function ParseBoolean(payload As Dictionary(Of String, Object), key As String, defaultValue As Boolean) As Boolean
            If payload Is Nothing OrElse Not payload.ContainsKey(key) Then Return defaultValue
            Try
                Return Convert.ToBoolean(payload(key))
            Catch
                Return defaultValue
            End Try
        End Function

        Private Shared Function ParseStringList(payload As Dictionary(Of String, Object), key As String) As List(Of String)
            Dim list As New List(Of String)()
            If payload Is Nothing OrElse Not payload.ContainsKey(key) Then Return list

            Dim raw = payload(key)
            Dim enumerable = TryCast(raw, System.Collections.IEnumerable)
            If enumerable IsNot Nothing AndAlso Not (TypeOf raw Is String) Then
                For Each o As Object In enumerable
                    Dim s = If(o, String.Empty).ToString()
                    If Not String.IsNullOrWhiteSpace(s) AndAlso Not list.Contains(s) Then
                        list.Add(s)
                    End If
                Next
            Else
                Dim s = If(raw, String.Empty).ToString()
                If Not String.IsNullOrWhiteSpace(s) Then
                    list.Add(s)
                End If
            End If
            Return list
        End Function

        Private Shared Function TryGetString(payload As Dictionary(Of String, Object), key As String) As String
            If payload Is Nothing OrElse Not payload.ContainsKey(key) Then Return String.Empty
            Return If(payload(key), String.Empty).ToString()
        End Function

        Private Shared Function ParseParamGroup(payload As Dictionary(Of String, Object), key As String) As BuiltInParameterGroup
            Dim raw As Object = Nothing
            If payload IsNot Nothing Then
                payload.TryGetValue(key, raw)
            End If

            If raw Is Nothing Then Return BuiltInParameterGroup.INVALID

            Try
                Dim s As String = raw.ToString()
                Dim iv As Integer
                If Integer.TryParse(s, iv) Then
                    Return CType(iv, BuiltInParameterGroup)
                End If
                Dim parsed As BuiltInParameterGroup
                If [Enum].TryParse(s, True, parsed) Then
                    Return parsed
                End If
            Catch
            End Try

            Try
                Dim iv As Integer = Convert.ToInt32(raw)
                Return CType(iv, BuiltInParameterGroup)
            Catch
            End Try

            Return BuiltInParameterGroup.INVALID
        End Function

        Private Shared Function ValidateBatchSettings(s As BatchSettings, spFilePath As String) As String
            If s Is Nothing Then Return "설정이 비어있습니다."
            If String.IsNullOrWhiteSpace(spFilePath) OrElse Not File.Exists(spFilePath) Then
                Return "Shared Parameter TXT가 유효하지 않습니다."
            End If

            If s.RvtPaths Is Nothing OrElse s.RvtPaths.Count = 0 Then
                Return "대상 RVT 파일을 1개 이상 선택하세요."
            End If

            If s.Parameters Is Nothing OrElse s.Parameters.Count = 0 Then
                Return "추가할 Shared Parameter를 1개 이상 등록하세요."
            End If

            Dim noCat As New List(Of String)()
            For Each p As ParamToBind In s.Parameters
                If p Is Nothing OrElse p.Settings Is Nothing OrElse p.Settings.Categories Is Nothing OrElse p.Settings.Categories.Count = 0 Then
                    noCat.Add(If(p IsNot Nothing, p.ParamName, "(unknown)"))
                End If
            Next
            If noCat.Count > 0 Then
                Return "카테고리가 지정되지 않은 파라미터가 있습니다." & vbCrLf &
                       String.Join(", ", noCat.ToArray())
            End If

            Return ""
        End Function

        Private Shared Function FindExternalDefinitionByGuid(defFile As DefinitionFile, targetGuid As Guid) As ExternalDefinition
            If defFile Is Nothing Then Return Nothing

            For Each g As DefinitionGroup In defFile.Groups
                If g Is Nothing Then Continue For
                For Each d As Definition In g.Definitions
                    Dim ed As ExternalDefinition = TryCast(d, ExternalDefinition)
                    If ed Is Nothing Then Continue For
                    If ed.GUID = targetGuid Then Return ed
                Next
            Next

            Return Nothing
        End Function

        Private Shared Function ApplyAllSharedParameterBindings(doc As Document,
                                                               app As RevitApp,
                                                               extDefByGuid As Dictionary(Of Guid, ExternalDefinition),
                                                               paramList As List(Of ParamToBind),
                                                               ByRef notes As String) As Boolean
            notes = ""

            If doc Is Nothing OrElse app Is Nothing OrElse extDefByGuid Is Nothing OrElse paramList Is Nothing OrElse paramList.Count = 0 Then
                notes = "입력값 오류(doc/app/defs/params)."
                Return False
            End If

            Dim maps As CategoryMaps = BuildAvailableCategoryMaps(doc)

            Dim warnLines As New List(Of String)()

            Using tg As New TransactionGroup(doc, "Batch Add Shared Parameters")
                tg.Start()

                For Each p As ParamToBind In paramList
                    If p Is Nothing OrElse p.Settings Is Nothing Then
                        notes = "파라미터 설정이 비어있습니다."
                        tg.RollBack()
                        Return False
                    End If

                    Dim extDef As ExternalDefinition = Nothing
                    If Not extDefByGuid.TryGetValue(p.GuidValue, extDef) OrElse extDef Is Nothing Then
                        notes = "정의 누락: " & p.ParamName
                        tg.RollBack()
                        Return False
                    End If

                    Dim perParamWarn As String = ""
                    Dim perParamErr As String = ""

                    Dim ok As Boolean = ApplyOneSharedParameterBinding(doc, app, maps, extDef, p, perParamWarn, perParamErr)

                    If Not ok Then
                        notes = "Param [" & p.ParamName & "] :: " & perParamErr
                        tg.RollBack()
                        Return False
                    End If

                    If Not String.IsNullOrWhiteSpace(perParamWarn) Then
                        Dim lines As String() = perParamWarn.Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                        For Each ln As String In lines
                            warnLines.Add("Param [" & p.ParamName & "] :: " & ln)
                        Next
                    End If
                Next

                tg.Assimilate()
            End Using

            If warnLines.Count > 0 Then
                notes = String.Join(vbCrLf, warnLines.ToArray())
            Else
                notes = ""
            End If

            Return True
        End Function

        Private Shared Function ApplyOneSharedParameterBinding(doc As Document,
                                                              app As RevitApp,
                                                              maps As CategoryMaps,
                                                              extDef As ExternalDefinition,
                                                              p As ParamToBind,
                                                              ByRef warnLog As String,
                                                              ByRef errLog As String) As Boolean
            warnLog = ""
            errLog = ""

            If HasParameterNameConflict(doc, extDef) Then
                errLog = "동일한 파라미터 이름이 이미 문서에 존재합니다(다른 GUID 가능). 이름 충돌."
                Return False
            End If

            Dim req As List(Of CategoryRef) = If(p.Settings.Categories, New List(Of CategoryRef)())
            Dim warn As New List(Of String)()

            Dim catSet As CategorySet = app.Create.NewCategorySet()
            Dim inserted As New Dictionary(Of Integer, CategoryRef)()

            For Each cref As CategoryRef In req
                If cref Is Nothing Then Continue For

                Dim resolveBy As String = ""
                Dim cat As Category = ResolveCategoryInDoc(maps, cref, resolveBy)

                If cat Is Nothing Then
                    warn.Add("CAT_NOT_FOUND: " & DescribeCategoryRef(cref))
                    Continue For
                End If

                Dim bindable As Boolean = False
                Try
                    bindable = cat.AllowsBoundParameters
                Catch
                    bindable = False
                End Try
                If Not bindable Then
                    warn.Add("CAT_NOT_BINDABLE(AllowsBoundParameters=False): " & DescribeCategoryRef(cref) &
                             " (ResolvedBy=" & resolveBy & ", ResolvedId=" & cat.Id.IntegerValue & ")")
                    Continue For
                End If

                Dim rid As Integer = cat.Id.IntegerValue
                If Not inserted.ContainsKey(rid) Then
                    catSet.Insert(cat)
                    inserted.Add(rid, cref)
                End If
            Next

            If inserted.Count = 0 Then
                errLog = "선택된 카테고리가 이 문서에서 유효하지 않음(미존재/바인딩 불가/해석 실패)."
                Return False
            End If

            Dim binding As Autodesk.Revit.DB.Binding = Nothing
            If p.Settings.IsInstanceBinding Then
                binding = app.Create.NewInstanceBinding(catSet)
            Else
                binding = app.Create.NewTypeBinding(catSet)
            End If

            If binding Is Nothing Then
                errLog = "Binding 생성 실패."
                Return False
            End If

            Try
                Using t As New Autodesk.Revit.DB.Transaction(doc, "Bind Shared Parameter: " & p.ParamName)
                    t.Start()

                    Dim map As BindingMap = doc.ParameterBindings

                    Dim insertedOk As Boolean = map.Insert(extDef, binding, p.Settings.ParamGroup)
                    If Not insertedOk Then
                        insertedOk = map.ReInsert(extDef, binding, p.Settings.ParamGroup)
                    End If

                    If Not insertedOk Then
                        t.RollBack()
                        errLog = "ParameterBindings Insert/ReInsert 실패(이미 존재/제약)."
                        Return False
                    End If

                    If p.Settings.IsInstanceBinding AndAlso p.Settings.AllowVaryBetweenGroups Then
                        Try
                            Dim spe2 As SharedParameterElement = SharedParameterElement.Lookup(doc, extDef.GUID)
                            If spe2 IsNot Nothing Then
                                Dim idef As InternalDefinition = TryCast(spe2.GetDefinition(), InternalDefinition)
                                If idef IsNot Nothing Then
                                    idef.SetAllowVaryBetweenGroups(doc, True)
                                End If
                            End If
                        Catch exVar As Exception
                            warn.Add("VARY_SET_FAIL: " & exVar.Message)
                        End Try
                    End If

                    t.Commit()
                End Using

                Dim boundIds As HashSet(Of Integer) = GetBoundCategoryIds(doc, extDef.GUID)
                If boundIds.Count > 0 Then
                    For Each kv As KeyValuePair(Of Integer, CategoryRef) In inserted
                        If Not boundIds.Contains(kv.Key) Then
                            warn.Add("CAT_DROPPED_BY_REVIT: " & DescribeCategoryRef(kv.Value) & " (ResolvedId=" & kv.Key & ")")
                        End If
                    Next
                End If

                If warn.Count > 0 Then
                    warnLog = String.Join(vbCrLf, warn.ToArray())
                Else
                    warnLog = ""
                End If

                Return True

            Catch ex As Exception
                errLog = "트랜잭션 예외: " & ex.Message
                Return False
            End Try
        End Function

        Private Shared Function DescribeCategoryRef(cref As CategoryRef) As String
            If cref Is Nothing Then Return "(null)"
            Dim p As String = If(cref.Path, "").Trim()
            If p <> "" Then
                Return p & " (SavedId=" & cref.IdInt & ")"
            End If
            Dim n As String = If(cref.Name, "").Trim()
            If n <> "" Then
                Return n & " (SavedId=" & cref.IdInt & ")"
            End If
            Return "(SavedId=" & cref.IdInt & ")"
        End Function

        Private Shared Function ResolveCategoryInDoc(maps As CategoryMaps, cref As CategoryRef, ByRef resolvedBy As String) As Category
            resolvedBy = ""

            If maps Is Nothing OrElse cref Is Nothing Then
                resolvedBy = "none"
                Return Nothing
            End If

            Dim cat As Category = Nothing

            Dim path As String = If(cref.Path, "").Trim()
            If path <> "" Then
                If maps.ByPath.TryGetValue(path, cat) AndAlso cat IsNot Nothing Then
                    Dim supportsId As Integer = CInt(BuiltInCategory.OST_StairsSupports)
                    If cat.Id.IntegerValue = supportsId Then
                        Dim rep As Category = Nothing
                        If maps.ById.TryGetValue(supportsId, rep) AndAlso rep IsNot Nothing Then
                            resolvedBy = "path->supportsfix"
                            Return rep
                        End If
                    End If

                    resolvedBy = "path"
                    Return cat
                End If
            End If

            If maps.ById.TryGetValue(cref.IdInt, cat) AndAlso cat IsNot Nothing Then
                resolvedBy = "id"
                Return cat
            End If

            resolvedBy = "notfound"
            Return Nothing
        End Function

        Private Shared Function GetBoundCategoryIds(doc As Document, extDefGuid As Guid) As HashSet(Of Integer)
            Dim hs As New HashSet(Of Integer)()

            Try
                Dim map As BindingMap = doc.ParameterBindings
                Dim it As DefinitionBindingMapIterator = map.ForwardIterator()
                it.Reset()

                While it.MoveNext()
                    Dim kDef As Definition = TryCast(it.Key, Definition)
                    Dim kExt As ExternalDefinition = TryCast(kDef, ExternalDefinition)

                    If kExt IsNot Nothing AndAlso kExt.GUID = extDefGuid Then
                        Dim eb As ElementBinding = TryCast(it.Current, ElementBinding)
                        If eb IsNot Nothing Then
                            Dim cs As CategorySet = eb.Categories
                            If cs IsNot Nothing Then
                                For Each c As Category In cs
                                    If c IsNot Nothing Then
                                        hs.Add(c.Id.IntegerValue)
                                    End If
                                Next
                            End If
                        End If

                        Exit While
                    End If
                End While

            Catch
            End Try

            Return hs
        End Function

        Private Shared Function HasParameterNameConflict(doc As Document, extDef As ExternalDefinition) As Boolean
            Dim spe As SharedParameterElement = SharedParameterElement.Lookup(doc, extDef.GUID)
            If spe IsNot Nothing Then
                Return False
            End If

            Try
                Dim collector As New FilteredElementCollector(doc)
                collector.OfClass(GetType(ParameterElement))
                For Each pe As ParameterElement In collector
                    If pe Is Nothing Then Continue For

                    Dim def As Definition = Nothing
                    Try
                        def = pe.GetDefinition()
                    Catch
                        def = Nothing
                    End Try
                    If def Is Nothing Then Continue For

                    If String.Equals(def.Name, extDef.Name, StringComparison.OrdinalIgnoreCase) Then
                        Dim speExisting As SharedParameterElement = TryCast(pe, SharedParameterElement)
                        If speExisting IsNot Nothing Then
                            If speExisting.GuidValue <> extDef.GUID Then
                                Return True
                            End If
                        Else
                            Return True
                        End If
                    End If
                Next
            Catch
                Return False
            End Try

            Return False
        End Function

        Private Shared Function OpenProjectDocument(app As RevitApp,
                                                   userVisiblePath As String,
                                                   closeAllWorksets As Boolean) As Document
            Dim mp As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(userVisiblePath)

            Dim openOpts As New OpenOptions()
            openOpts.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach

            If closeAllWorksets Then
                Dim wsc As New WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
                openOpts.SetOpenWorksetsConfiguration(wsc)
            End If

            Return app.OpenDocumentFile(mp, openOpts)
        End Function

        Private Shared Function SyncWithCentral(doc As Document,
                                               comment As String,
                                               ByRef err As String) As Boolean
            err = ""
            If doc Is Nothing OrElse Not doc.IsWorkshared Then
                err = "Workshared 문서가 아님."
                Return False
            End If

            Try
                Dim twc As New TransactWithCentralOptions()

                Dim swc As New SynchronizeWithCentralOptions()
                swc.Comment = If(comment, "")

                Try
                    Dim rel As New RelinquishOptions(True)
                    swc.SetRelinquishOptions(rel)
                Catch
                End Try

                doc.SynchronizeWithCentral(twc, swc)
                Return True

            Catch ex As Exception
                err = ex.Message
                Return False
            End Try
        End Function

        Private Shared Sub SafeClose(doc As Document, saveModified As Boolean)
            If doc Is Nothing Then Return
            Try
                doc.Close(saveModified)
            Catch
            End Try
        End Sub

        Private Shared Function IsAlreadyOpen(app As RevitApp, userVisiblePath As String) As Boolean
            Try
                For Each d As Document In app.Documents
                    If d Is Nothing Then Continue For
                    If String.Equals(d.PathName, userVisiblePath, StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Next
            Catch
                Return False
            End Try
            Return False
        End Function

        Private Shared Function BuildCategoryTree(doc As Document) As List(Of CategoryTreeItem)
            Dim roots As New List(Of CategoryTreeItem)()
            Dim visitedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            For Each c As Category In doc.Settings.Categories
                Dim node As CategoryTreeItem = BuildCategoryTreeItemRecursive(c, Nothing, visitedPaths)
                If node IsNot Nothing Then roots.Add(node)
            Next

            Return roots.OrderBy(Function(x) x.Name).ToList()
        End Function

        Private Shared Function BuildCategoryTreeItemRecursive(cat As Category,
                                                              parentPath As String,
                                                              visitedPaths As HashSet(Of String)) As CategoryTreeItem
            If cat Is Nothing Then Return Nothing

            Dim path As String = If(String.IsNullOrEmpty(parentPath), cat.Name, parentPath & "\" & cat.Name)

            If visitedPaths IsNot Nothing Then
                If visitedPaths.Contains(path) Then Return Nothing
                visitedPaths.Add(path)
            End If

            Dim bindable As Boolean = False
            Try
                bindable = cat.AllowsBoundParameters
            Catch
                bindable = False
            End Try

            Try
                If cat.Id.IntegerValue = CInt(BuiltInCategory.OST_StairsSupports) Then
                    bindable = True
                End If
            Catch
            End Try

            Dim it As New CategoryTreeItem() With {
                .IdInt = cat.Id.IntegerValue,
                .Name = cat.Name,
                .Path = path,
                .CatType = cat.CategoryType,
                .IsBindable = bindable
            }

            Try
                Dim subs As CategoryNameMap = cat.SubCategories
                If subs IsNot Nothing Then
                    For Each sc As Category In subs
                        Dim child As CategoryTreeItem = BuildCategoryTreeItemRecursive(sc, path, visitedPaths)
                        If child IsNot Nothing Then it.Children.Add(child)
                    Next
                End If
            Catch
            End Try

            If it.IsBindable Then Return it
            If it.Children IsNot Nothing AndAlso it.Children.Count > 0 Then Return it
            Return Nothing
        End Function

        Private Shared Function BuildAvailableCategoryMaps(doc As Document) As CategoryMaps
            Dim maps As New CategoryMaps()
            Dim visitedPaths As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Try
                For Each c As Category In doc.Settings.Categories
                    CollectCategoryRecursive(c, Nothing, maps, visitedPaths)
                Next
            Catch
            End Try

            Try
                InjectStairsSupportsBindingFix(doc, maps)
            Catch
            End Try

            Return maps
        End Function

        Private Shared Sub CollectCategoryRecursive(cat As Category,
                                                    parentPath As String,
                                                    maps As CategoryMaps,
                                                    visitedPaths As HashSet(Of String))
            If cat Is Nothing OrElse maps Is Nothing Then Return

            Dim path As String = If(String.IsNullOrEmpty(parentPath), cat.Name, parentPath & "\" & cat.Name)

            If visitedPaths IsNot Nothing Then
                If visitedPaths.Contains(path) Then Return
                visitedPaths.Add(path)
            End If

            Dim idInt As Integer = cat.Id.IntegerValue

            If Not maps.ById.ContainsKey(idInt) Then
                maps.ById.Add(idInt, cat)
            End If
            If Not maps.ByPath.ContainsKey(path) Then
                maps.ByPath.Add(path, cat)
            End If

            Try
                Dim subs As CategoryNameMap = cat.SubCategories
                If subs IsNot Nothing Then
                    For Each sc As Category In subs
                        CollectCategoryRecursive(sc, path, maps, visitedPaths)
                    Next
                End If
            Catch
            End Try
        End Sub

        Private Shared Function TryGetCategoryFromElementType(doc As Document, bic As BuiltInCategory) As Category
            If doc Is Nothing Then Return Nothing

            Try
                Dim e As Element = New FilteredElementCollector(doc) _
                    .OfCategory(bic) _
                    .WhereElementIsElementType() _
                    .FirstElement()

                If e Is Nothing Then Return Nothing
                Return e.Category

            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Sub InjectStairsSupportsBindingFix(doc As Document, maps As CategoryMaps)
            If doc Is Nothing OrElse maps Is Nothing Then Return

            Dim replacement As Category = TryGetCategoryFromElementType(doc, BuiltInCategory.OST_StairsStringerCarriage)
            If replacement Is Nothing Then Return

            Dim supportsId As Integer = CInt(BuiltInCategory.OST_StairsSupports)

            Try
                maps.ById(supportsId) = replacement
            Catch
            End Try

            Try
                Dim keys As List(Of String) = maps.ByPath _
                    .Where(Function(kv) kv.Value IsNot Nothing AndAlso kv.Value.Id.IntegerValue = supportsId) _
                    .Select(Function(kv) kv.Key) _
                    .ToList()

                For Each k As String In keys
                    maps.ByPath(k) = replacement
                Next
            Catch
            End Try
        End Sub

        Private Shared Function ToCategoryDto(item As CategoryTreeItem) As Object
            If item Is Nothing Then Return Nothing
            Return New With {
                .idInt = item.IdInt,
                .name = item.Name,
                .path = item.Path,
                .catType = item.CatType.ToString(),
                .isBindable = item.IsBindable,
                .children = item.Children.Select(Function(c) ToCategoryDto(c)).ToList()
            }
        End Function

    End Class

End Namespace
