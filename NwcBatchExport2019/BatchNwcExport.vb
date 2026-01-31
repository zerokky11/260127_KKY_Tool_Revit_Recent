Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Automation
Imports System.Drawing

Imports Autodesk.Revit.Attributes
Imports RvtDB = Autodesk.Revit.DB
Imports RvtUI = Autodesk.Revit.UI
Imports WF = System.Windows.Forms

<Transaction(TransactionMode.Manual)>
Public Class BatchNwcExportCommand
    Implements RvtUI.IExternalCommand

    Public Function Execute(commandData As RvtUI.ExternalCommandData,
                            ByRef message As String,
                            elements As RvtDB.ElementSet) As RvtUI.Result Implements RvtUI.IExternalCommand.Execute
        Try
            UiManager.Show(commandData.Application)
            Return RvtUI.Result.Succeeded
        Catch ex As Exception
            message = ex.ToString()
            Return RvtUI.Result.Failed
        End Try
    End Function
End Class

Friend NotInheritable Class UiManager
    Private Sub New()
    End Sub

    Private Shared _form As BatchForm
    Private Shared _handler As ExportExternalEventHandler
    Private Shared _exEvent As RvtUI.ExternalEvent

    Public Shared Sub Show(uiapp As RvtUI.UIApplication)
        If _form IsNot Nothing AndAlso Not _form.IsDisposed Then
            _form.BringToFront()
            _form.Focus()
            Return
        End If

        Dim logSink As Action(Of String) =
          Sub(msg As String)
              If _form Is Nothing OrElse _form.IsDisposed Then Return
              _form.SafeLog(msg)
          End Sub

        _handler = New ExportExternalEventHandler(uiapp, logSink)
        _exEvent = RvtUI.ExternalEvent.Create(_handler)

        _form = New BatchForm(uiapp, _exEvent, _handler)
        _form.Show()
        _form.SafeLog("Ready.")
    End Sub
End Class

Friend Class ViewExportItem
    Public Property ViewId As RvtDB.ElementId
    Public Property ViewName As String
    Public Property FacetingFactor As Double
    Public Property IncludeLinks As Boolean
End Class

Friend Class NwcOptions
    Public Property ConvertConstructionParts As Boolean = False
    Public Property ConvertElementIds As Boolean = True
    Public Property ConvertElementProperties As Boolean = True
    Public Property ConvertLights As Boolean = False
    Public Property ConvertLinkedCadFormats As Boolean = False
    Public Property ConvertRoomAsAttribute As Boolean = False
    Public Property ConvertUrls As Boolean = False
    Public Property ExportRoomGeometry As Boolean = False

    Public Property ConvertElementParameters As String = "All"
    Public Property Coordinates As String = "Shared"
    Public Property DivideFileIntoLevels As Boolean = True

    ' 중요: Current View 강제(원하는 동작)
    Public Property ExportScope As String = "Current View"

    Public Property TryFindMissingMaterials As Boolean = True
End Class

Friend Class ExportJob
    Public Property Items As List(Of ViewExportItem)
    Public Property OutputFolder As String

    Public Property AllowOverwrite As Boolean
    Public Property ContinueOnError As Boolean
    Public Property VerifyFactorAfterSet As Boolean

    Public Property ExporterClientId As String
    Public Property ExporterClassName As String

    Public Property GlobalOptions As NwcOptions
End Class

Friend Class ExportExternalEventHandler
    Implements RvtUI.IExternalEventHandler

    Private ReadOnly _uiapp As RvtUI.UIApplication
    Private ReadOnly _controller As ExportController
    Private ReadOnly _log As Action(Of String)

    Private ReadOnly _lockObj As New Object()
    Private _pendingJob As ExportJob = Nothing
    Private _pendingCancel As Boolean = False

    Public Sub New(uiapp As RvtUI.UIApplication, logSink As Action(Of String))
        _uiapp = uiapp
        _log = logSink
        _controller = New ExportController(_uiapp, _log)
    End Sub

    Public Sub EnqueueStart(job As ExportJob)
        SyncLock _lockObj
            _pendingJob = job
            _pendingCancel = False
        End SyncLock
    End Sub

    Public Sub EnqueueCancel()
        SyncLock _lockObj
            _pendingCancel = True
        End SyncLock
    End Sub

    Public Sub Execute(app As RvtUI.UIApplication) Implements RvtUI.IExternalEventHandler.Execute
        Dim job As ExportJob = Nothing
        Dim cancel As Boolean = False

        SyncLock _lockObj
            job = _pendingJob
            cancel = _pendingCancel
            _pendingJob = Nothing
            _pendingCancel = False
        End SyncLock

        If cancel Then
            _controller.Cancel()
            Return
        End If

        If job IsNot Nothing Then
            _controller.Start(job)
        End If
    End Sub

    Public Function GetName() As String Implements RvtUI.IExternalEventHandler.GetName
        Return "NWC Batch Export 2019 ExternalEvent"
    End Function
End Class

Friend Class BatchForm
    Inherits WF.Form

    Private ReadOnly _uiapp As RvtUI.UIApplication
    Private ReadOnly _exEvent As RvtUI.ExternalEvent
    Private ReadOnly _handler As ExportExternalEventHandler

    Private ReadOnly _grid As WF.DataGridView
    Private ReadOnly _namesBox As WF.TextBox

    Private ReadOnly _folderBox As WF.TextBox
    Private ReadOnly _browseBtn As WF.Button

    Private ReadOnly _defaultFactorBox As WF.NumericUpDown
    Private ReadOnly _applyFactorBtn As WF.Button

    Private ReadOnly _globalLinksChk As WF.CheckBox
    Private ReadOnly _applyLinksBtn As WF.Button

    Private ReadOnly _overwriteChk As WF.CheckBox
    Private ReadOnly _continueChk As WF.CheckBox
    Private ReadOnly _verifyChk As WF.CheckBox

    Private ReadOnly _clientIdBox As WF.TextBox
    Private ReadOnly _classBox As WF.TextBox

    ' Options Editor UI
    Private ReadOnly _optConvertConstructionParts As WF.CheckBox
    Private ReadOnly _optConvertElementIds As WF.CheckBox
    Private ReadOnly _optConvertElementProps As WF.CheckBox
    Private ReadOnly _optConvertLights As WF.CheckBox

    Private ReadOnly _optConvertLinkedCad As WF.CheckBox
    Private ReadOnly _optDivideLevels As WF.CheckBox
    Private ReadOnly _optConvertRoomAsAttribute As WF.CheckBox
    Private ReadOnly _optConvertUrls As WF.CheckBox

    Private ReadOnly _optConvertElementParams As WF.ComboBox
    Private ReadOnly _optCoordinates As WF.ComboBox
    Private ReadOnly _optExportScope As WF.ComboBox

    Private ReadOnly _optExportRoomGeometry As WF.CheckBox
    Private ReadOnly _optMissingMaterials As WF.CheckBox

    Private ReadOnly _startBtn As WF.Button
    Private ReadOnly _cancelBtn As WF.Button
    Private ReadOnly _logBox As WF.TextBox

    Private Const COL_EXPORT As String = "Export"
    Private Const COL_NAME As String = "ViewName"
    Private Const COL_FACTOR As String = "Factor"
    Private Const COL_LINKS As String = "Links"

    Public Sub New(uiapp As RvtUI.UIApplication, exEvent As RvtUI.ExternalEvent, handler As ExportExternalEventHandler)
        _uiapp = uiapp
        _exEvent = exEvent
        _handler = handler

        Text = "NWC Batch Export (Revit 2019) - UIA (optionseditor aware)"
        Width = 1180
        Height = 900
        StartPosition = WF.FormStartPosition.CenterScreen
        AutoScaleMode = WF.AutoScaleMode.Dpi
        MinimumSize = New Size(980, 720)

        Dim main As New WF.TableLayoutPanel() With {
          .RowCount = 6,
          .ColumnCount = 2,
          .AutoSize = True,
          .AutoSizeMode = WF.AutoSizeMode.GrowAndShrink,
          .Dock = WF.DockStyle.Top
        }
        main.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Percent, 60))
        main.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Percent, 40))

        main.RowStyles.Add(New WF.RowStyle(WF.SizeType.Absolute, 320))
        main.RowStyles.Add(New WF.RowStyle(WF.SizeType.Absolute, 120))
        main.RowStyles.Add(New WF.RowStyle(WF.SizeType.Absolute, 190))
        main.RowStyles.Add(New WF.RowStyle(WF.SizeType.Absolute, 120))
        main.RowStyles.Add(New WF.RowStyle(WF.SizeType.Absolute, 220))
        main.RowStyles.Add(New WF.RowStyle(WF.SizeType.Absolute, 260))

        ' ===== Views Grid =====
        _grid = New WF.DataGridView() With {
          .Dock = WF.DockStyle.Fill,
          .AllowUserToAddRows = False,
          .AllowUserToDeleteRows = False,
          .RowHeadersVisible = False,
          .SelectionMode = WF.DataGridViewSelectionMode.FullRowSelect,
          .MultiSelect = True,
          .AutoSizeColumnsMode = WF.DataGridViewAutoSizeColumnsMode.Fill
        }
        _grid.Columns.Add(New WF.DataGridViewCheckBoxColumn() With {.Name = COL_EXPORT, .HeaderText = "Export", .FillWeight = 12})
        _grid.Columns.Add(New WF.DataGridViewTextBoxColumn() With {.Name = COL_NAME, .HeaderText = "3D View Name", .ReadOnly = True, .FillWeight = 58})
        _grid.Columns.Add(New WF.DataGridViewTextBoxColumn() With {.Name = COL_FACTOR, .HeaderText = "Factor", .FillWeight = 15})
        _grid.Columns.Add(New WF.DataGridViewCheckBoxColumn() With {.Name = COL_LINKS, .HeaderText = "Links", .FillWeight = 15})
        AddHandler _grid.CellValidating, AddressOf OnGridCellValidating

        Dim viewsGroup As New WF.GroupBox() With {.Dock = WF.DockStyle.Fill, .Text = "3D Views (각 행 Factor/Links)"}
        viewsGroup.Controls.Add(_grid)
        main.Controls.Add(viewsGroup, 0, 0)
        main.SetColumnSpan(viewsGroup, 2)

        ' ===== Names Box =====
        _namesBox = New WF.TextBox() With {.Dock = WF.DockStyle.Fill, .Multiline = True, .ScrollBars = WF.ScrollBars.Vertical}
        Dim namesGroup As New WF.GroupBox() With {
          .Dock = WF.DockStyle.Fill,
          .Text = "추가 입력(줄바꿈). 형식: 뷰이름<TAB>factor<TAB>true/false  또는  뷰이름|factor|true (뷰이름 변경 금지)"
        }
        namesGroup.Controls.Add(_namesBox)
        main.Controls.Add(namesGroup, 0, 1)
        main.SetColumnSpan(namesGroup, 2)

        ' ===== Basic Options =====
        Dim optPanel As New WF.TableLayoutPanel() With {.Dock = WF.DockStyle.Fill, .RowCount = 6, .ColumnCount = 3}
        optPanel.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Absolute, 220))
        optPanel.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Percent, 100))
        optPanel.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Absolute, 190))
        For i As Integer = 0 To 5
            optPanel.RowStyles.Add(New WF.RowStyle(WF.SizeType.Absolute, 26))
        Next

        optPanel.Controls.Add(New WF.Label() With {.Text = "Output Folder", .Dock = WF.DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 0)
        _folderBox = New WF.TextBox() With {.Dock = WF.DockStyle.Fill}
        _browseBtn = New WF.Button() With {.Text = "Browse", .Dock = WF.DockStyle.Fill}
        AddHandler _browseBtn.Click, AddressOf OnBrowse
        optPanel.Controls.Add(_folderBox, 1, 0)
        optPanel.Controls.Add(_browseBtn, 2, 0)

        optPanel.Controls.Add(New WF.Label() With {.Text = "Default/Apply Factor", .Dock = WF.DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 1)
        Dim factorBar As New WF.FlowLayoutPanel() With {.Dock = WF.DockStyle.Fill, .FlowDirection = WF.FlowDirection.LeftToRight}
        _defaultFactorBox = New WF.NumericUpDown() With {.Minimum = 0, .Maximum = 1000, .DecimalPlaces = 2, .Increment = 0.5D, .Value = 1D, .Width = 90}
        _applyFactorBtn = New WF.Button() With {.Text = "Apply to selected rows", .Width = 170}
        AddHandler _applyFactorBtn.Click, AddressOf OnApplyFactorToSelected
        factorBar.Controls.Add(_defaultFactorBox)
        factorBar.Controls.Add(_applyFactorBtn)
        optPanel.Controls.Add(factorBar, 1, 1)

        optPanel.Controls.Add(New WF.Label() With {.Text = "Default Links", .Dock = WF.DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 2)
        Dim linksBar As New WF.FlowLayoutPanel() With {.Dock = WF.DockStyle.Fill, .FlowDirection = WF.FlowDirection.LeftToRight}
        _globalLinksChk = New WF.CheckBox() With {.Text = "Links=ON", .Checked = True, .AutoSize = True}
        _applyLinksBtn = New WF.Button() With {.Text = "Apply to selected rows", .Width = 170}
        AddHandler _applyLinksBtn.Click, AddressOf OnApplyLinksToSelected
        linksBar.Controls.Add(_globalLinksChk)
        linksBar.Controls.Add(_applyLinksBtn)
        optPanel.Controls.Add(linksBar, 1, 2)

        optPanel.Controls.Add(New WF.Label() With {.Text = "Flags", .Dock = WF.DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 3)
        Dim flagsBar As New WF.FlowLayoutPanel() With {.Dock = WF.DockStyle.Fill, .FlowDirection = WF.FlowDirection.LeftToRight}
        _overwriteChk = New WF.CheckBox() With {.Text = "Overwrite", .Checked = True, .AutoSize = True}
        _continueChk = New WF.CheckBox() With {.Text = "Continue on error", .Checked = False, .AutoSize = True}
        _verifyChk = New WF.CheckBox() With {.Text = "Verify factor after set", .Checked = True, .AutoSize = True}
        flagsBar.Controls.Add(_overwriteChk)
        flagsBar.Controls.Add(_continueChk)
        flagsBar.Controls.Add(_verifyChk)
        optPanel.Controls.Add(flagsBar, 1, 3)

        Dim btnPanel As New WF.FlowLayoutPanel() With {.Dock = WF.DockStyle.Fill, .FlowDirection = WF.FlowDirection.LeftToRight}
        _startBtn = New WF.Button() With {.Text = "Start Batch Export", .Width = 200, .Height = 26}
        _cancelBtn = New WF.Button() With {.Text = "Cancel", .Width = 100, .Height = 26}
        AddHandler _startBtn.Click, AddressOf OnStart
        AddHandler _cancelBtn.Click, AddressOf OnCancel
        btnPanel.Controls.Add(_startBtn)
        btnPanel.Controls.Add(_cancelBtn)
        optPanel.Controls.Add(btnPanel, 1, 5)

        Dim optGroup As New WF.GroupBox() With {.Dock = WF.DockStyle.Fill, .Text = "Basic Options"}
        optGroup.Controls.Add(optPanel)
        main.Controls.Add(optGroup, 0, 2)
        main.SetColumnSpan(optGroup, 2)

        ' ===== Exporter Command =====
        Dim cmdPanel As New WF.TableLayoutPanel() With {.Dock = WF.DockStyle.Fill, .RowCount = 2, .ColumnCount = 2}
        cmdPanel.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Absolute, 220))
        cmdPanel.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Percent, 100))
        cmdPanel.RowStyles.Add(New WF.RowStyle(WF.SizeType.Absolute, 26))
        cmdPanel.RowStyles.Add(New WF.RowStyle(WF.SizeType.Absolute, 26))

        cmdPanel.Controls.Add(New WF.Label() With {.Text = "Exporter ClientId (GUID)", .Dock = WF.DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 0)
        _clientIdBox = New WF.TextBox() With {.Dock = WF.DockStyle.Fill, .Text = "2c0e1f44-77d9-4745-aa8a-4657534613c8"}
        cmdPanel.Controls.Add(_clientIdBox, 1, 0)

        cmdPanel.Controls.Add(New WF.Label() With {.Text = "Exporter ClassName", .Dock = WF.DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 1)
        _classBox = New WF.TextBox() With {.Dock = WF.DockStyle.Fill, .Text = "NavisWorks16.LcRevitExportCommand"}
        cmdPanel.Controls.Add(_classBox, 1, 1)

        Dim cmdGroup As New WF.GroupBox() With {.Dock = WF.DockStyle.Fill, .Text = "Navisworks Exporter Command (PostCommand)"}
        cmdGroup.Controls.Add(cmdPanel)
        main.Controls.Add(cmdGroup, 0, 3)
        main.SetColumnSpan(cmdGroup, 2)

        ' ===== Options Editor (UI에서 지정) =====
        Dim navPanel As New WF.TableLayoutPanel() With {.Dock = WF.DockStyle.Fill, .RowCount = 4, .ColumnCount = 4}
        navPanel.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Percent, 25))
        navPanel.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Percent, 25))
        navPanel.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Percent, 25))
        navPanel.ColumnStyles.Add(New WF.ColumnStyle(WF.SizeType.Percent, 25))
        For i As Integer = 0 To 3
            navPanel.RowStyles.Add(New WF.RowStyle(WF.SizeType.Absolute, 26))
        Next

        _optConvertConstructionParts = New WF.CheckBox() With {.Text = "Convert construction parts", .Checked = False, .AutoSize = True}
        _optConvertElementIds = New WF.CheckBox() With {.Text = "Convert element IDs", .Checked = True, .AutoSize = True}
        _optConvertElementProps = New WF.CheckBox() With {.Text = "Convert element properties", .Checked = True, .AutoSize = True}
        _optConvertLights = New WF.CheckBox() With {.Text = "Convert lights", .Checked = False, .AutoSize = True}

        _optConvertLinkedCad = New WF.CheckBox() With {.Text = "Convert linked CAD formats", .Checked = False, .AutoSize = True}
        _optDivideLevels = New WF.CheckBox() With {.Text = "Divide File into Levels", .Checked = True, .AutoSize = True}
        _optConvertRoomAsAttribute = New WF.CheckBox() With {.Text = "Convert room as attribute", .Checked = False, .AutoSize = True}
        _optConvertUrls = New WF.CheckBox() With {.Text = "Convert URLs", .Checked = False, .AutoSize = True}

        _optConvertElementParams = New WF.ComboBox() With {.DropDownStyle = WF.ComboBoxStyle.DropDown}
        _optConvertElementParams.Items.AddRange(New Object() {"All", "None"})
        _optConvertElementParams.Text = "All"

        _optCoordinates = New WF.ComboBox() With {.DropDownStyle = WF.ComboBoxStyle.DropDown}
        _optCoordinates.Items.AddRange(New Object() {"Shared", "Internal"})
        _optCoordinates.Text = "Shared"

        _optExportScope = New WF.ComboBox() With {.DropDownStyle = WF.ComboBoxStyle.DropDown}
        _optExportScope.Items.AddRange(New Object() {"Current View", "View", "현재 뷰", "Selection", "Entire model"})
        _optExportScope.Text = "Current View"

        _optExportRoomGeometry = New WF.CheckBox() With {.Text = "Export room geometry", .Checked = False, .AutoSize = True}
        _optMissingMaterials = New WF.CheckBox() With {.Text = "Try and find missing materials", .Checked = True, .AutoSize = True}

        navPanel.Controls.Add(_optConvertConstructionParts, 0, 0)
        navPanel.Controls.Add(_optConvertElementIds, 1, 0)
        navPanel.Controls.Add(_optConvertElementProps, 2, 0)
        navPanel.Controls.Add(_optConvertLights, 3, 0)

        navPanel.Controls.Add(_optConvertLinkedCad, 0, 1)
        navPanel.Controls.Add(_optDivideLevels, 1, 1)
        navPanel.Controls.Add(_optConvertRoomAsAttribute, 2, 1)
        navPanel.Controls.Add(_optConvertUrls, 3, 1)

        navPanel.Controls.Add(New WF.Label() With {.Text = "Convert element parameters:", .Dock = WF.DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 2)
        navPanel.Controls.Add(_optConvertElementParams, 1, 2)
        navPanel.Controls.Add(New WF.Label() With {.Text = "Coordinates:", .Dock = WF.DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 2, 2)
        navPanel.Controls.Add(_optCoordinates, 3, 2)

        navPanel.Controls.Add(New WF.Label() With {.Text = "Export:", .Dock = WF.DockStyle.Fill, .TextAlign = ContentAlignment.MiddleLeft}, 0, 3)
        navPanel.Controls.Add(_optExportScope, 1, 3)
        navPanel.Controls.Add(_optExportRoomGeometry, 2, 3)
        navPanel.Controls.Add(_optMissingMaterials, 3, 3)

        Dim navGroup As New WF.GroupBox() With {.Dock = WF.DockStyle.Fill, .Text = "Navisworks Options Editor - Revit (applied every export)"}
        navGroup.Controls.Add(navPanel)
        main.Controls.Add(navGroup, 0, 4)
        main.SetColumnSpan(navGroup, 2)

        ' ===== Log =====
        _logBox = New WF.TextBox() With {
          .Dock = WF.DockStyle.Fill, .Multiline = True, .ReadOnly = True,
          .ScrollBars = WF.ScrollBars.Both, .WordWrap = False
        }
        Dim logGroup As New WF.GroupBox() With {.Dock = WF.DockStyle.Fill, .Text = "Log"}
        logGroup.Controls.Add(_logBox)
        main.Controls.Add(logGroup, 0, 5)
        main.SetColumnSpan(logGroup, 2)

        Dim scrollPanel As New WF.Panel() With {.Dock = WF.DockStyle.Fill, .AutoScroll = True}
        scrollPanel.Controls.Add(main)
        Controls.Add(scrollPanel)

        Dim docPath As String = ""
        Try : docPath = _uiapp.ActiveUIDocument.Document.PathName : Catch : End Try
        If Not String.IsNullOrWhiteSpace(docPath) AndAlso File.Exists(docPath) Then
            _folderBox.Text = Path.Combine(Path.GetDirectoryName(docPath), "NWC")
        Else
            _folderBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
        End If

        LoadViewsIntoGrid()
    End Sub

    Private Sub LoadViewsIntoGrid()
        _grid.Rows.Clear()

        Dim doc = _uiapp.ActiveUIDocument.Document
        Dim views = New RvtDB.FilteredElementCollector(doc).
          OfClass(GetType(RvtDB.View3D)).
          Cast(Of RvtDB.View3D)().
          Where(Function(v) Not v.IsTemplate).
          OrderBy(Function(v) v.Name, StringComparer.OrdinalIgnoreCase).
          ToList()

        For Each v In views
            Dim idx = _grid.Rows.Add(False, v.Name, _defaultFactorBox.Value.ToString(CultureInfo.InvariantCulture), _globalLinksChk.Checked)
            _grid.Rows(idx).Tag = v.Id
        Next
    End Sub

    Private Sub OnGridCellValidating(sender As Object, e As WF.DataGridViewCellValidatingEventArgs)
        If e.ColumnIndex < 0 OrElse e.RowIndex < 0 Then Return
        If _grid.Columns(e.ColumnIndex).Name <> COL_FACTOR Then Return

        Dim text = Convert.ToString(e.FormattedValue, CultureInfo.InvariantCulture)
        If String.IsNullOrWhiteSpace(text) Then Return

        Dim tmp As Double
        If Not DoubleTryParseFlexible(text, tmp) Then
            e.Cancel = True
            SafeLog("ERROR: Factor는 숫자여야 합니다. 입력=" & text)
        End If
    End Sub

    Private Sub OnApplyFactorToSelected(sender As Object, e As EventArgs)
        Dim v = _defaultFactorBox.Value.ToString(CultureInfo.InvariantCulture)
        For Each r As WF.DataGridViewRow In _grid.SelectedRows
            r.Cells(COL_FACTOR).Value = v
        Next
    End Sub

    Private Sub OnApplyLinksToSelected(sender As Object, e As EventArgs)
        For Each r As WF.DataGridViewRow In _grid.SelectedRows
            r.Cells(COL_LINKS).Value = _globalLinksChk.Checked
        Next
    End Sub

    Private Sub OnBrowse(sender As Object, e As EventArgs)
        Using dlg As New WF.FolderBrowserDialog()
            dlg.SelectedPath = _folderBox.Text
            If dlg.ShowDialog() = WF.DialogResult.OK Then
                _folderBox.Text = dlg.SelectedPath
            End If
        End Using
    End Sub

    Private Sub OnStart(sender As Object, e As EventArgs)
        Dim outDir = _folderBox.Text.Trim()
        If String.IsNullOrWhiteSpace(outDir) Then
            SafeLog("ERROR: Output Folder가 비었습니다.")
            Return
        End If
        Directory.CreateDirectory(outDir)

        Dim doc = _uiapp.ActiveUIDocument.Document
        Dim viewMap = New RvtDB.FilteredElementCollector(doc).
          OfClass(GetType(RvtDB.View3D)).
          Cast(Of RvtDB.View3D)().
          Where(Function(v) Not v.IsTemplate).
          ToDictionary(Function(v) v.Name, Function(v) v.Id, StringComparer.Ordinal)

        Dim itemsById As New Dictionary(Of Integer, ViewExportItem)()

        For Each row As WF.DataGridViewRow In _grid.Rows
            Dim doExport As Boolean = False
            Try : doExport = Convert.ToBoolean(row.Cells(COL_EXPORT).Value) : Catch : End Try
            If Not doExport Then Continue For

            Dim idObj = TryCast(row.Tag, RvtDB.ElementId)
            If idObj Is Nothing Then Continue For

            Dim name = Convert.ToString(row.Cells(COL_NAME).Value, CultureInfo.InvariantCulture)

            Dim factorText = Convert.ToString(row.Cells(COL_FACTOR).Value, CultureInfo.InvariantCulture)
            Dim factorVal As Double = Convert.ToDouble(_defaultFactorBox.Value, CultureInfo.InvariantCulture)
            If Not String.IsNullOrWhiteSpace(factorText) Then
                If Not DoubleTryParseFlexible(factorText, factorVal) Then
                    SafeLog("ERROR: Factor 파싱 실패: " & name & " / " & factorText)
                    Return
                End If
            End If

            Dim linksVal As Boolean = _globalLinksChk.Checked
            Try : linksVal = Convert.ToBoolean(row.Cells(COL_LINKS).Value) : Catch : End Try

            itemsById(idObj.IntegerValue) = New ViewExportItem() With {
              .ViewId = idObj, .ViewName = name, .FacetingFactor = factorVal, .IncludeLinks = linksVal
            }
        Next

        Dim lines = _namesBox.Text.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries).
          Select(Function(x) x.Trim()).
          Where(Function(x) x.Length > 0).ToList()

        For Each ln In lines
            Dim parsed = ParseNameLine(ln)
            If parsed Is Nothing Then
                SafeLog("WARN: 라인 파싱 실패(스킵): " & ln)
                Continue For
            End If

            Dim nm = parsed.Item1
            Dim fac = parsed.Item2
            Dim linksOpt = parsed.Item3

            If Not viewMap.ContainsKey(nm) Then
                SafeLog("WARN: 입력한 뷰 이름을 찾지 못함: " & nm)
                Continue For
            End If

            Dim id = viewMap(nm)
            Dim linksVal = If(linksOpt.HasValue, linksOpt.Value, _globalLinksChk.Checked)
            Dim factorVal = If(fac.HasValue, fac.Value, Convert.ToDouble(_defaultFactorBox.Value, CultureInfo.InvariantCulture))

            itemsById(id.IntegerValue) = New ViewExportItem() With {
              .ViewId = id, .ViewName = nm, .FacetingFactor = factorVal, .IncludeLinks = linksVal
            }
        Next

        Dim items = itemsById.Values.OrderBy(Function(it) it.ViewName, StringComparer.Ordinal).ToList()
        If items.Count = 0 Then
            SafeLog("ERROR: Export할 뷰가 없습니다.")
            Return
        End If

        For Each it In items
            If it.ViewName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 Then
                SafeLog("ERROR: 뷰 이름에 파일명 불가 문자가 있습니다: " & it.ViewName)
                Return
            End If
        Next

        Dim cid = _clientIdBox.Text.Trim()
        Dim cls = _classBox.Text.Trim()
        If String.IsNullOrWhiteSpace(cid) OrElse String.IsNullOrWhiteSpace(cls) Then
            SafeLog("ERROR: Exporter ClientId/ClassName이 비었습니다.")
            Return
        End If

        Dim opt As New NwcOptions() With {
          .ConvertConstructionParts = _optConvertConstructionParts.Checked,
          .ConvertElementIds = _optConvertElementIds.Checked,
          .ConvertElementProperties = _optConvertElementProps.Checked,
          .ConvertLights = _optConvertLights.Checked,
          .ConvertLinkedCadFormats = _optConvertLinkedCad.Checked,
          .ConvertRoomAsAttribute = _optConvertRoomAsAttribute.Checked,
          .ConvertUrls = _optConvertUrls.Checked,
          .ExportRoomGeometry = _optExportRoomGeometry.Checked,
          .ConvertElementParameters = _optConvertElementParams.Text.Trim(),
          .Coordinates = _optCoordinates.Text.Trim(),
          .DivideFileIntoLevels = _optDivideLevels.Checked,
          .ExportScope = _optExportScope.Text.Trim(),
          .TryFindMissingMaterials = _optMissingMaterials.Checked
        }

        Dim job As New ExportJob() With {
          .Items = items,
          .OutputFolder = outDir,
          .AllowOverwrite = _overwriteChk.Checked,
          .ContinueOnError = _continueChk.Checked,
          .VerifyFactorAfterSet = _verifyChk.Checked,
          .ExporterClientId = cid,
          .ExporterClassName = cls,
          .GlobalOptions = opt
        }

        _handler.EnqueueStart(job)
        _exEvent.Raise()
        SafeLog("Start requested. Items=" & job.Items.Count.ToString(CultureInfo.InvariantCulture))
    End Sub

    Private Sub OnCancel(sender As Object, e As EventArgs)
        _handler.EnqueueCancel()
        _exEvent.Raise()
        SafeLog("Cancel requested.")
    End Sub

    Public Sub SafeLog(msg As String)
        If InvokeRequired Then
            BeginInvoke(New Action(Of String)(AddressOf SafeLog), msg)
            Return
        End If
        _logBox.AppendText(DateTime.Now.ToString("HH:mm:ss") & "  " & msg & Environment.NewLine)
        _logBox.SelectionStart = _logBox.TextLength
        _logBox.ScrollToCaret()
    End Sub

    Private Shared Function DoubleTryParseFlexible(text As String, ByRef value As Double) As Boolean
        Dim t = text.Trim()
        If Double.TryParse(t, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, value) Then Return True
        If Double.TryParse(t, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, value) Then Return True
        Dim swapped = t.Replace(","c, "."c)
        If Double.TryParse(swapped, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, value) Then Return True
        swapped = t.Replace("."c, ","c)
        If Double.TryParse(swapped, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, value) Then Return True
        Return False
    End Function

    Private Shared Function ParseNameLine(line As String) As Tuple(Of String, Nullable(Of Double), Nullable(Of Boolean))
        Dim seps As Char() = {ControlChars.Tab, "|"c, ","c, ";"c}
        Dim parts = line.Split(seps, StringSplitOptions.RemoveEmptyEntries).Select(Function(x) x.Trim()).ToList()
        If parts.Count = 0 Then Return Nothing
        Dim name = parts(0)
        If String.IsNullOrWhiteSpace(name) Then Return Nothing

        Dim fac As Nullable(Of Double) = Nothing
        If parts.Count >= 2 Then
            Dim tmp As Double
            If Double.TryParse(parts(1), NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, tmp) OrElse
               Double.TryParse(parts(1), NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, tmp) Then
                fac = tmp
            End If
        End If

        Dim links As Nullable(Of Boolean) = Nothing
        If parts.Count >= 3 Then
            Dim v = parts(2).Trim().ToLowerInvariant()
            If v = "true" OrElse v = "t" OrElse v = "1" OrElse v = "yes" OrElse v = "y" OrElse v = "예" Then links = True
            If v = "false" OrElse v = "f" OrElse v = "0" OrElse v = "no" OrElse v = "n" OrElse v = "아니오" Then links = False
        End If

        Return Tuple.Create(name, fac, links)
    End Function
End Class

Friend Class ExportController
    Private Enum RunState
        Idle = 0
        NextItem = 10
        WaitViewActivated = 11
        WaitAutomation = 12
        WaitFileStable = 13
        Done = 90
        ErrorState = 91
        Cancelled = 92
    End Enum

    Private ReadOnly _uiapp As RvtUI.UIApplication
    Private ReadOnly _log As Action(Of String)

    Private _state As RunState = RunState.Idle
    Private _job As ExportJob = Nothing
    Private _queue As Queue(Of ViewExportItem) = Nothing
    Private _cancel As Boolean = False

    Private ReadOnly _revitPid As Integer = Process.GetCurrentProcess().Id

    Private _current As ViewExportItem = Nothing
    Private _currentTargetPath As String = Nothing
    Private _automationTask As Task(Of Boolean) = Nothing

    Private _stableSize As Long = -1
    Private _stableHits As Integer = 0

    Private _navCmdId As RvtUI.RevitCommandId = Nothing
    Private _postedForThisItem As Boolean = False

    Public Sub New(uiapp As RvtUI.UIApplication, logSink As Action(Of String))
        _uiapp = uiapp
        _log = logSink
    End Sub

    Public Sub Start(job As ExportJob)
        If _state <> RunState.Idle AndAlso _state <> RunState.Done AndAlso _state <> RunState.ErrorState AndAlso _state <> RunState.Cancelled Then
            _log("이미 실행 중입니다.")
            Return
        End If

        _cancel = False
        _job = job
        _queue = New Queue(Of ViewExportItem)(job.Items)

        _navCmdId = ResolveNavisworksCommandId(_uiapp, _job.ExporterClientId, _job.ExporterClassName, _log)
        If _navCmdId Is Nothing Then
            Finish(RunState.ErrorState, "ERROR: Navisworks Exporter CommandId를 찾지 못했습니다. ClientId/ClassName 확인 필요")
            Return
        End If

        AddHandler _uiapp.Idling, AddressOf OnIdling
        _state = RunState.NextItem
        _log("Batch started. PID=" & _revitPid.ToString(CultureInfo.InvariantCulture))
    End Sub

    Public Sub Cancel()
        _cancel = True
    End Sub

    Private Sub Finish(finalState As RunState, msg As String)
        Try : RemoveHandler _uiapp.Idling, AddressOf OnIdling : Catch : End Try
        _state = finalState
        _automationTask = Nothing
        _log(msg)
    End Sub

    Private Sub OnIdling(sender As Object, e As RvtUI.Events.IdlingEventArgs)
        If _cancel Then
            Finish(RunState.Cancelled, "Cancelled.")
            Return
        End If

        Try
            e.SetRaiseWithoutDelay()

            Dim uidoc = _uiapp.ActiveUIDocument
            If uidoc Is Nothing OrElse uidoc.Document Is Nothing Then
                Finish(RunState.ErrorState, "ERROR: Active document가 없습니다.")
                Return
            End If

            Select Case _state

                Case RunState.NextItem
                    If _queue.Count = 0 Then
                        Finish(RunState.Done, "Done.")
                        Return
                    End If

                    _current = _queue.Dequeue()
                    _postedForThisItem = False

                    Dim v = TryCast(uidoc.Document.GetElement(_current.ViewId), RvtDB.View3D)
                    If v Is Nothing OrElse v.IsTemplate Then
                        _log("WARN: 유효하지 않은 3D 뷰 스킵: " & _current.ViewName)
                        _state = RunState.NextItem
                        Return
                    End If

                    _currentTargetPath = Path.Combine(_job.OutputFolder, _current.ViewName & ".nwc")
                    _stableSize = -1
                    _stableHits = 0

                    uidoc.RequestViewChange(v)
                    _state = RunState.WaitViewActivated
                    _log("Switching view: " & _current.ViewName &
                         " (Factor=" & _current.FacetingFactor.ToString(CultureInfo.InvariantCulture) &
                         ", Links=" & _current.IncludeLinks.ToString() & ")")

                Case RunState.WaitViewActivated
                    Dim av = uidoc.ActiveView
                    If av Is Nothing Then Return
                    If av.Id.IntegerValue = _current.ViewId.IntegerValue Then
                        _state = RunState.WaitAutomation
                    End If

                Case RunState.WaitAutomation
                    If _automationTask Is Nothing Then

                        If Not _postedForThisItem Then
                            If Not _uiapp.CanPostCommand(_navCmdId) Then
                                Dim msg = "ERROR: CanPostCommand=False (Navisworks exporter PostCommand 불가)"
                                If _job.ContinueOnError Then
                                    _log(msg & " (continue)")
                                    _state = RunState.NextItem
                                    Return
                                Else
                                    Finish(RunState.ErrorState, msg)
                                    Return
                                End If
                            End If

                            _uiapp.PostCommand(_navCmdId)
                            _postedForThisItem = True
                            _log("Navisworks exporter PostCommand posted.")
                        End If

                        Dim item = _current
                        Dim target = _currentTargetPath
                        Dim overwrite = _job.AllowOverwrite
                        Dim verify = _job.VerifyFactorAfterSet
                        Dim opt = _job.GlobalOptions
                        Dim pid = _revitPid
                        Dim logSink = _log

                        _automationTask = UiAutomationSta.Run(Function()
                                                                  Return UiNwcAutomator.RunNavisworksExportDialogs(
                                                                    pid, target, overwrite, item, opt, verify, logSink)
                                                              End Function)
                    End If

                    If _automationTask.IsCompleted Then
                        Dim ok As Boolean = (_automationTask.Status = TaskStatus.RanToCompletion AndAlso _automationTask.Result)
                        _automationTask = Nothing

                        If ok Then
                            _state = RunState.WaitFileStable
                        Else
                            Dim msg = "ERROR: export automation failed for " & _currentTargetPath
                            If _job.ContinueOnError Then
                                _log(msg & " (continue)")
                                _state = RunState.NextItem
                            Else
                                Finish(RunState.ErrorState, msg)
                            End If
                        End If
                    End If

                Case RunState.WaitFileStable
                    If IsFileStable(_currentTargetPath) Then
                        _log("OK: " & _currentTargetPath)
                        _state = RunState.NextItem
                    End If

            End Select

        Catch ex As Exception
            Finish(RunState.ErrorState, "ERROR: " & ex.Message)
        End Try
    End Sub

    Private Function IsFileStable(path As String) As Boolean
        Try
            If Not File.Exists(path) Then Return False
            Dim fi As New FileInfo(path)
            Dim size = fi.Length

            If size = _stableSize Then
                _stableHits += 1
            Else
                _stableSize = size
                _stableHits = 0
            End If

            Return _stableHits >= 2
        Catch
            Return False
        End Try
    End Function

    Private Shared Function ResolveNavisworksCommandId(uiapp As RvtUI.UIApplication,
                                                      clientId As String,
                                                      className As String,
                                                      log As Action(Of String)) As RvtUI.RevitCommandId
        Dim cid = (If(clientId, "")).Trim()
        Dim cls = (If(className, "")).Trim()
        If cid.Length = 0 Then Return Nothing

        Dim guidNoBraces As String = cid.Trim("{"c, "}"c)
        Dim guidBraces As String = "{" & guidNoBraces & "}"

        Dim candidates As New List(Of String)()
        candidates.Add(guidNoBraces)
        candidates.Add(guidBraces)

        If cls.Length > 0 Then
            candidates.Add(guidNoBraces & ":" & cls)
            candidates.Add(guidBraces & ":" & cls)
            candidates.Add("Execute external command:" & guidNoBraces & ":" & cls)
            candidates.Add("Execute external command:" & guidBraces & ":" & cls)
        End If

        For Each s In candidates
            Try
                Dim id = RvtUI.RevitCommandId.LookupCommandId(s)
                If id IsNot Nothing AndAlso uiapp.CanPostCommand(id) Then
                    log("Resolved exporter CommandId via: " & s)
                    Return id
                End If
            Catch
            End Try
        Next

        Return Nothing
    End Function
End Class

Friend NotInheritable Class UiAutomationSta
    Private Sub New()
    End Sub

    Public Shared Function Run(Of T)(work As Func(Of T)) As Task(Of T)
        Dim tcs As New TaskCompletionSource(Of T)()
        Dim th As New Thread(Sub()
                                 Try
                                     tcs.SetResult(work())
                                 Catch ex As Exception
                                     tcs.SetException(ex)
                                 End Try
                             End Sub)
        th.IsBackground = True
        th.SetApartmentState(ApartmentState.STA)
        th.Start()
        Return tcs.Task
    End Function
End Class

Friend NotInheritable Class UiNwcAutomator
    Private Sub New()
    End Sub

    ' =========================
    ' Win32
    ' =========================
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowPos(hWnd As IntPtr,
                                       hWndInsertAfter As IntPtr,
                                       X As Integer,
                                       Y As Integer,
                                       cx As Integer,
                                       cy As Integer,
                                       uFlags As UInteger) As Boolean
    End Function
    Private Const SWP_NOSIZE As UInteger = &H1UI
    Private Const SWP_NOZORDER As UInteger = &H4UI
    Private Const SWP_NOACTIVATE As UInteger = &H10UI

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function PostMessage(hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As Boolean
    End Function
    Private Const BM_CLICK As UInteger = &HF5UI

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function GetWindowText(hWnd As IntPtr, lpString As System.Text.StringBuilder, cch As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function EnumWindows(lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function

    Private Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean

    Private Shared Function GetWin32Title(hwnd As IntPtr) As String
        Try
            If hwnd = IntPtr.Zero Then Return ""
            Dim sb As New System.Text.StringBuilder(512)
            Dim n = GetWindowText(hwnd, sb, sb.Capacity)
            If n <= 0 Then Return ""
            Return sb.ToString()
        Catch
            Return ""
        End Try
    End Function
    Private Const AddinTitleToken As String = "NWC Batch Export"

    Private Shared ReadOnly ExportDialogTitleCandidates As String() = {
  "Export scene as", "Export Scene As", "Export scene as...",
  "장면 내보내기", "장면 내보내기..."
}

    Private Shared Function IsAddinWindow(win As AutomationElement) As Boolean
        If win Is Nothing Then Return False

        Dim hwnd As IntPtr = IntPtr.Zero
        Try : hwnd = New IntPtr(win.Current.NativeWindowHandle) : Catch : hwnd = IntPtr.Zero : End Try

        Dim t As String = GetWin32Title(hwnd)
        If String.IsNullOrWhiteSpace(t) Then t = SafeName(win)
        If String.IsNullOrWhiteSpace(t) Then Return False

        Return t.IndexOf(AddinTitleToken, StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    Private Shared Sub MoveOffscreen(win As AutomationElement)
        Try
            Dim hwnd As IntPtr = New IntPtr(win.Current.NativeWindowHandle)
            If hwnd = IntPtr.Zero Then Return
            SetWindowPos(hwnd, IntPtr.Zero, -32000, -32000, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE)
        Catch
        End Try
    End Sub

    ' =========================
    ' Names / Candidates
    ' =========================
    Private Const OptionsEditorProcessName As String = "navisworks.optionseditor"

    Private Shared ReadOnly OptionsEditorTitleCandidates As String() = {
    "Navisworks Options Editor - Revit",
    "Navisworks Options Editor",
    "Options Editor - Revit",
    "Options Editor"
  }

    Private Shared ReadOnly ExportSettingsButtonCandidates As String() = {
    "Navisworks settings", "Navisworks Settings", "Navisworks settings...", "Navisworks Settings...",
    "Navisworks 설정", "Navisworks 설정..."
  }

    Private Shared ReadOnly SaveButtonCandidates As String() = {"Save", "저장", "저장(S)", "OK", "확인", "Export", "내보내기"}
    Private Shared ReadOnly CancelButtonCandidates As String() = {"Cancel", "취소"}
    Private Shared ReadOnly OkButtonCandidates As String() = {"OK", "확인"}
    Private Shared ReadOnly YesButtonCandidates As String() = {"Yes", "예", "Overwrite", "덮어쓰기", "Replace", "대체", "확인"}

    Private Shared ReadOnly TreeFileReadersCandidates As String() = {"File Readers", "파일 리더"}
    Private Shared ReadOnly TreeRevitCandidates As String() = {"Revit"}

    Private Shared ReadOnly LblFacetingCandidates As String() = {"Faceting Factor", "Faceting"}
    Private Shared ReadOnly LblConvertElementParams As String() = {"Convert element parameters"}
    Private Shared ReadOnly LblCoordinates As String() = {"Coordinates"}
    Private Shared ReadOnly LblExportScope As String() = {"Export"}

    Private Shared ReadOnly ChkConvertConstructionParts As String() = {"Convert construction parts"}
    Private Shared ReadOnly ChkConvertElementIds As String() = {"Convert element IDs", "Convert element ids"}
    Private Shared ReadOnly ChkConvertElementProps As String() = {"Convert element properties"}
    Private Shared ReadOnly ChkConvertLights As String() = {"Convert lights"}
    Private Shared ReadOnly ChkConvertLinkedCad As String() = {"Convert linked CAD formats"}
    Private Shared ReadOnly ChkConvertLinkedFiles As String() = {"Convert linked files"}
    Private Shared ReadOnly ChkConvertRoomAsAttribute As String() = {"Convert room as attribute"}
    Private Shared ReadOnly ChkConvertUrls As String() = {"Convert URLs", "Convert Urls"}
    Private Shared ReadOnly ChkDivideLevels As String() = {"Divide File into Levels"}
    Private Shared ReadOnly ChkExportRoomGeometry As String() = {"Export room geometry"}
    Private Shared ReadOnly ChkMissingMaterials As String() = {"Try and find missing materials"}

    ' =========================
    ' MAIN
    ' =========================
    Public Shared Function RunNavisworksExportDialogs(revitPid As Integer,
                                                   targetFullPath As String,
                                                   overwrite As Boolean,
                                                   item As ViewExportItem,
                                                   opt As NwcOptions,
                                                   verifyFactorAfterSet As Boolean,
                                                   log As Action(Of String)) As Boolean
        Dim L As Action(Of String) =
      Sub(s As String)
          If log IsNot Nothing Then log("UIA: " & s)
      End Sub

        L("Wait export save dialog...")
        Dim exportDlg = WaitForExportDialog(revitPid, 180000, L) ' 3분
        If exportDlg Is Nothing Then
            L("ERROR export save dialog not found (timeout).")
            Return False
        End If

        Dim exHwnd As IntPtr = IntPtr.Zero
        Try : exHwnd = New IntPtr(exportDlg.Current.NativeWindowHandle) : Catch : exHwnd = IntPtr.Zero : End Try
        L("Export dialog found. UIAName=""" & SafeName(exportDlg) & """ WinTitle=""" & GetWin32Title(exHwnd) &
      """ pid=" & exportDlg.Current.ProcessId.ToString() & " hwnd=" & exHwnd.ToString())
        MoveOffscreen(exportDlg)

        Dim settingsBtn = FindFirstClickableByContains(exportDlg, ExportSettingsButtonCandidates)
        If settingsBtn Is Nothing Then
            L("ERROR Navisworks settings... button not found.")
            DumpDialogButtons(exportDlg, L)
            CloseByCancel(exportDlg)
            Return False
        End If

        ' ===== 핵심: Settings 클릭을 "논블로킹"으로 =====
        L("Click Navisworks settings... (NoBlock)")
        ClickNoBlock(settingsBtn, L)

        ' Options Editor는 환경에 따라 늦게 뜰 수 있음(내부 로딩)
        L("Wait Options Editor... (process-aware, timeout=300s)")
        Dim optDlg = WaitForOptionsEditorDialog(revitPid, 300000, L) ' 5분
        If optDlg Is Nothing Then
            L("ERROR Options Editor not found (timeout).")
            CloseByCancel(exportDlg)
            Return False
        End If

        Dim opHwnd As IntPtr = IntPtr.Zero
        Try : opHwnd = New IntPtr(optDlg.Current.NativeWindowHandle) : Catch : opHwnd = IntPtr.Zero : End Try
        L("Options Editor found. UIAName=""" & SafeName(optDlg) & """ WinTitle=""" & GetWin32Title(opHwnd) &
      """ pid=" & optDlg.Current.ProcessId.ToString() & " hwnd=" & opHwnd.ToString())
        MoveOffscreen(optDlg)

        ' --- Options Editor 설정 적용 ---
        EnsureRevitReaderPage(optDlg)

        TrySetCheckboxByNameContains(optDlg, ChkConvertConstructionParts, opt.ConvertConstructionParts)
        TrySetCheckboxByNameContains(optDlg, ChkConvertElementIds, opt.ConvertElementIds)
        TrySetCheckboxByNameContains(optDlg, ChkConvertElementProps, opt.ConvertElementProperties)
        TrySetCheckboxByNameContains(optDlg, ChkConvertLights, opt.ConvertLights)
        TrySetCheckboxByNameContains(optDlg, ChkConvertLinkedCad, opt.ConvertLinkedCadFormats)

        ' 뷰별 링크 포함
        TrySetCheckboxByNameContains(optDlg, ChkConvertLinkedFiles, item.IncludeLinks)

        TrySetCheckboxByNameContains(optDlg, ChkConvertRoomAsAttribute, opt.ConvertRoomAsAttribute)
        TrySetCheckboxByNameContains(optDlg, ChkConvertUrls, opt.ConvertUrls)
        TrySetCheckboxByNameContains(optDlg, ChkDivideLevels, opt.DivideFileIntoLevels)
        TrySetCheckboxByNameContains(optDlg, ChkExportRoomGeometry, opt.ExportRoomGeometry)
        TrySetCheckboxByNameContains(optDlg, ChkMissingMaterials, opt.TryFindMissingMaterials)

        If Not String.IsNullOrWhiteSpace(opt.ConvertElementParameters) Then
            TrySelectComboToRightOfLabel(optDlg, LblConvertElementParams,
                                   New List(Of String) From {opt.ConvertElementParameters, "All", "None"}, L)
        End If

        If Not String.IsNullOrWhiteSpace(opt.Coordinates) Then
            TrySelectComboToRightOfLabel(optDlg, LblCoordinates,
                                   New List(Of String) From {opt.Coordinates, "Shared", "Internal"}, L)
        End If

        ' Export scope
        If Not TrySetExportScope(optDlg, opt.ExportScope, L) Then
            L("WARN Export scope set/verify failed (may already be Current View).")
        End If

        ' Faceting factor
        If Not TrySetAndVerifyFactor(optDlg, item.FacetingFactor, verifyFactorAfterSet, L) Then
            L("ERROR Faceting Factor set/verify failed.")
            CloseByOk(optDlg)
            CloseByCancel(exportDlg)
            Return False
        End If

        ' OK close
        CloseByOk(optDlg)
        WaitUntilClosed(optDlg, 30000)

        ' --- Back to export dialog ---
        exportDlg = WaitForExportDialog(revitPid, 180000, L)
        If exportDlg Is Nothing Then
            L("ERROR export dialog not found after Options Editor.")
            Return False
        End If
        MoveOffscreen(exportDlg)

        ' Set file path
        If Not SetTargetPath(exportDlg, targetFullPath, L) Then
            L("ERROR cannot set target path.")
            DumpDialogEdits(exportDlg, L)
            CloseByCancel(exportDlg)
            Return False
        End If

        Dim saveBtn = FindFirstClickableByContains(exportDlg, SaveButtonCandidates)
        If saveBtn Is Nothing Then
            L("ERROR Save/OK button not found.")
            DumpDialogButtons(exportDlg, L)
            CloseByCancel(exportDlg)
            Return False
        End If

        L("Click Save/OK")
        ClickNoBlock(saveBtn, L)

        If overwrite Then
            HandleOverwriteDialogIfAny()
        End If

        L("Wait export dialog closed...")
        WaitUntilClosed(exportDlg, 10 * 60 * 1000)

        L("Done.")
        Return True
    End Function

    ' =========================
    ' Click NoBlock
    ' =========================
    Private Shared Sub ClickNoBlock(el As AutomationElement, L As Action(Of String))
        If el Is Nothing Then Return

        ' 1) hwnd 있으면 BM_CLICK(PostMessage)로 끝 (모달이어도 블로킹 없음)
        Try
            Dim hwnd As IntPtr = New IntPtr(el.Current.NativeWindowHandle)
            If hwnd <> IntPtr.Zero Then
                PostMessage(hwnd, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
                Return
            End If
        Catch
        End Try

        ' 2) hwnd 없으면 Invoke를 별도 스레드에서 (블로킹되더라도 우리 흐름은 계속)
        Dim th As New Thread(Sub()
                                 Try
                                     Invoke(el)
                                 Catch
                                 End Try
                             End Sub)
        th.IsBackground = True
        th.SetApartmentState(ApartmentState.STA)
        th.Start()
    End Sub

    ' =========================
    ' Export dialog detection
    ' - 설정버튼 + (파일이름 1148 Edit/Combo) + (Save/Cancel)
    ' =========================
    Private Shared Function WaitForExportDialog(revitPid As Integer, timeoutMs As Integer, L As Action(Of String)) As AutomationElement
        Dim until = DateTime.UtcNow.AddMilliseconds(timeoutMs)
        While DateTime.UtcNow < until
            Dim w = FindTopLevelWindowByPid(revitPid, Function(win) IsExportDialogWindow(win))
            If w IsNot Nothing Then Return w

            Dim w2 = FindTopLevelWindowAnyProcess(Function(win) IsExportDialogWindow(win))
            If w2 IsNot Nothing Then Return w2

            Thread.Sleep(120)
        End While
        Return Nothing
    End Function

    Private Shared Function IsExportDialogWindow(win As AutomationElement) As Boolean
        If win Is Nothing Then Return False
        If IsAddinWindow(win) Then Return False

        Dim hwnd As IntPtr = IntPtr.Zero
        Try : hwnd = New IntPtr(win.Current.NativeWindowHandle) : Catch : hwnd = IntPtr.Zero : End Try

        Dim title As String = GetWin32Title(hwnd)
        If String.IsNullOrWhiteSpace(title) Then Return False

        Dim okTitle As Boolean =
    ExportDialogTitleCandidates.Any(Function(k) title.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0)
        If Not okTitle Then Return False

        Dim hasSettings = (FindFirstClickableByContains(win, ExportSettingsButtonCandidates) IsNot Nothing)
        If Not hasSettings Then Return False

        ' 파일이름 입력칸(1148) 존재 확인
        Dim fileBox = FindFileNameValueTarget(win, AddressOf NoLog)
        If fileBox Is Nothing Then Return False

        Dim hasSave = (FindFirstClickableByContains(win, SaveButtonCandidates) IsNot Nothing)
        If Not hasSave Then Return False

        Return True
    End Function


    ' =========================
    ' Options Editor detection
    ' - navisworks.optionseditor 프로세스 우선 (MainWindowHandle/EnumWindows)
    ' - UIA fallback
    ' =========================
    Private Shared Function WaitForOptionsEditorDialog(revitPid As Integer, timeoutMs As Integer, L As Action(Of String)) As AutomationElement
        Dim until = DateTime.UtcNow.AddMilliseconds(timeoutMs)

        While DateTime.UtcNow < until
            ' 1) 프로세스 기반 (가장 안정)
            Dim found = FindOptionsEditorViaProcess(L)
            If found IsNot Nothing Then Return found

            ' 2) Revit PID 내부로 뜨는 환경 대비
            Dim w2 = FindTopLevelWindowByPid(revitPid, AddressOf IsOptionsEditorLike)
            If w2 IsNot Nothing Then Return w2

            ' 3) 전체 fallback
            Dim w3 = FindTopLevelWindowAnyProcess(AddressOf IsOptionsEditorLike)
            If w3 IsNot Nothing Then Return w3

            Thread.Sleep(150)
        End While

        Return Nothing
    End Function

    Private Shared Function FindOptionsEditorViaProcess(L As Action(Of String)) As AutomationElement
        Try
            Dim ps = Process.GetProcessesByName(OptionsEditorProcessName)
            For Each p In ps
                If p Is Nothing OrElse p.HasExited Then Continue For

                ' MainWindowHandle 우선
                If p.MainWindowHandle <> IntPtr.Zero Then
                    Dim ae = AutomationElement.FromHandle(p.MainWindowHandle)
                    If ae IsNot Nothing AndAlso IsOptionsEditorLike(ae) Then Return ae
                End If

                ' MainWindowHandle이 0인 경우: EnumWindows로 pid 매칭
                For Each hwnd In EnumerateTopLevelWindowsByPid(p.Id)
                    If hwnd = IntPtr.Zero Then Continue For
                    Dim ae = AutomationElement.FromHandle(hwnd)
                    If ae IsNot Nothing AndAlso IsOptionsEditorLike(ae) Then Return ae
                Next
            Next
        Catch
        End Try

        Return Nothing
    End Function

    Private Shared Function EnumerateTopLevelWindowsByPid(pid As Integer) As List(Of IntPtr)
        Dim results As New List(Of IntPtr)()
        Try
            EnumWindows(Function(hWnd, lParam)
                            Dim p As Integer = 0
                            GetWindowThreadProcessId(hWnd, p)
                            If p = pid Then results.Add(hWnd)
                            Return True
                        End Function, IntPtr.Zero)
        Catch
        End Try
        Return results
    End Function

    Private Shared Function IsOptionsEditorLike(win As AutomationElement) As Boolean
        If win Is Nothing Then Return False
        If IsAddinWindow(win) Then Return False

        Dim hwnd As IntPtr = IntPtr.Zero
        Try : hwnd = New IntPtr(win.Current.NativeWindowHandle) : Catch : hwnd = IntPtr.Zero : End Try

        Dim title As String = GetWin32Title(hwnd)
        If String.IsNullOrWhiteSpace(title) Then title = SafeName(win)
        If String.IsNullOrWhiteSpace(title) Then Return False

        Dim titleOk As Boolean =
    OptionsEditorTitleCandidates.Any(Function(t) title.IndexOf(t, StringComparison.OrdinalIgnoreCase) >= 0) OrElse
    (title.IndexOf("Navisworks Options Editor", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso
     title.IndexOf("Revit", StringComparison.OrdinalIgnoreCase) >= 0)

        If Not titleOk Then Return False

        ' 반드시 좌측 Tree에 File Readers / Revit 노드가 있어야 함 (애드인 폼은 이런 트리가 없음)
        Dim hasFileReaders As Boolean = False
        For Each nm In TreeFileReadersCandidates
            Dim ti = win.FindFirst(TreeScope.Descendants,
      New AndCondition(
        New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem),
        New PropertyCondition(AutomationElement.NameProperty, nm)
      ))
            If ti IsNot Nothing Then hasFileReaders = True : Exit For
        Next
        If Not hasFileReaders Then Return False

        Dim hasRevitNode As Boolean = False
        For Each nm In TreeRevitCandidates
            Dim ti = win.FindFirst(TreeScope.Descendants,
      New AndCondition(
        New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem),
        New PropertyCondition(AutomationElement.NameProperty, nm)
      ))
            If ti IsNot Nothing Then hasRevitNode = True : Exit For
        Next
        If Not hasRevitNode Then Return False

        Return True
    End Function


    Private Shared Function FindTopLevelWindowByPid(pid As Integer, predicate As Func(Of AutomationElement, Boolean)) As AutomationElement
        Try
            Dim root = AutomationElement.RootElement
            Dim cond = New AndCondition(
        New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
        New PropertyCondition(AutomationElement.ProcessIdProperty, pid)
      )
            Dim wins = root.FindAll(TreeScope.Children, cond)
            For i As Integer = 0 To wins.Count - 1
                Dim win = TryCast(wins(i), AutomationElement)
                If win Is Nothing Then Continue For
                If predicate(win) Then Return win
            Next
        Catch
        End Try
        Return Nothing
    End Function

    Private Shared Function FindTopLevelWindowAnyProcess(predicate As Func(Of AutomationElement, Boolean)) As AutomationElement
        Try
            Dim root = AutomationElement.RootElement
            Dim cond = New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window)
            Dim wins = root.FindAll(TreeScope.Children, cond)
            For i As Integer = 0 To wins.Count - 1
                Dim win = TryCast(wins(i), AutomationElement)
                If win Is Nothing Then Continue For
                If predicate(win) Then Return win
            Next
        Catch
        End Try
        Return Nothing
    End Function

    ' =========================
    ' Find clickable (Button + SplitButton)
    ' =========================
    Private Shared Function FindFirstClickableByContains(window As AutomationElement, candidates As IEnumerable(Of String)) As AutomationElement
        If window Is Nothing Then Return Nothing

        Dim types As ControlType() = {ControlType.Button, ControlType.SplitButton}
        For Each ct In types
            Dim els = window.FindAll(TreeScope.Descendants, New PropertyCondition(AutomationElement.ControlTypeProperty, ct))
            For i As Integer = 0 To els.Count - 1
                Dim e = TryCast(els(i), AutomationElement)
                If e Is Nothing Then Continue For
                Dim nm = SafeName(e)
                If String.IsNullOrWhiteSpace(nm) Then Continue For
                If candidates.Any(Function(c) nm.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0) Then
                    Return e
                End If
            Next
        Next

        Return Nothing
    End Function

    ' =========================
    ' File name/path setter (ComboBox 1148 대응)
    ' =========================
    Private Shared Function SetTargetPath(exportDlg As AutomationElement, targetFullPath As String, L As Action(Of String)) As Boolean
        Dim valueTarget = FindFileNameValueTarget(exportDlg, L)
        If valueTarget Is Nothing Then Return False

        If SupportsWritableValuePattern(valueTarget) Then
            If SetValue(valueTarget, targetFullPath) Then Return True
        End If

        ' ComboBox 안의 Edit 재시도
        Try
            Dim innerEdit = valueTarget.FindFirst(TreeScope.Descendants, New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit))
            If innerEdit IsNot Nothing AndAlso SupportsWritableValuePattern(innerEdit) Then
                Return SetValue(innerEdit, targetFullPath)
            End If
        Catch
        End Try

        Return False
    End Function

    Private Shared Function FindFileNameValueTarget(window As AutomationElement, L As Action(Of String)) As AutomationElement
        If window Is Nothing Then Return Nothing

        ' ComboBox(1148) 우선
        Dim cb1148 = window.FindFirst(TreeScope.Descendants,
                                  New AndCondition(
                                    New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox),
                                    New PropertyCondition(AutomationElement.AutomationIdProperty, "1148")
                                  ))
        If cb1148 IsNot Nothing Then
            Dim ed = cb1148.FindFirst(TreeScope.Descendants, New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit))
            If ed IsNot Nothing Then Return ed
            Return cb1148
        End If

        ' Edit(1148)
        Dim ed1148 = window.FindFirst(TreeScope.Descendants,
                                  New AndCondition(
                                    New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                                    New PropertyCondition(AutomationElement.AutomationIdProperty, "1148")
                                  ))
        If ed1148 IsNot Nothing Then Return ed1148

        Return Nothing
    End Function

    ' =========================
    ' Options Editor navigation / setters
    ' =========================
    Private Shared Sub EnsureRevitReaderPage(settingsDlg As AutomationElement)
        Try
            Dim fileReaders As AutomationElement = Nothing
            For Each nm In TreeFileReadersCandidates
                fileReaders = settingsDlg.FindFirst(TreeScope.Descendants,
                    New AndCondition(
                      New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem),
                      New PropertyCondition(AutomationElement.NameProperty, nm)
                    ))
                If fileReaders IsNot Nothing Then Exit For
            Next

            If fileReaders IsNot Nothing Then
                Dim exp = TryCast(fileReaders.GetCurrentPattern(ExpandCollapsePattern.Pattern), ExpandCollapsePattern)
                If exp IsNot Nothing Then exp.Expand()
            End If

            Dim revitNode As AutomationElement = Nothing
            For Each nm In TreeRevitCandidates
                revitNode = settingsDlg.FindFirst(TreeScope.Descendants,
                    New AndCondition(
                      New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TreeItem),
                      New PropertyCondition(AutomationElement.NameProperty, nm)
                    ))
                If revitNode IsNot Nothing Then Exit For
            Next

            If revitNode IsNot Nothing Then
                Dim sel = TryCast(revitNode.GetCurrentPattern(SelectionItemPattern.Pattern), SelectionItemPattern)
                If sel IsNot Nothing Then sel.Select() Else Invoke(revitNode)
            End If
        Catch
        End Try
    End Sub

    Private Shared Sub TrySetCheckboxByNameContains(window As AutomationElement, nameCandidates As IEnumerable(Of String), desired As Boolean)
        Try
            Dim cbs = window.FindAll(TreeScope.Descendants, New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.CheckBox))
            For i As Integer = 0 To cbs.Count - 1
                Dim cb = TryCast(cbs(i), AutomationElement)
                If cb Is Nothing Then Continue For
                Dim nm = SafeName(cb)
                If String.IsNullOrWhiteSpace(nm) Then Continue For

                If nameCandidates.Any(Function(k) nm.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0) Then
                    Dim t = TryCast(cb.GetCurrentPattern(TogglePattern.Pattern), TogglePattern)
                    If t Is Nothing Then Exit Sub
                    Dim nowOn = (t.Current.ToggleState = ToggleState.On)
                    If nowOn <> desired Then t.Toggle()
                    Exit Sub
                End If
            Next
        Catch
        End Try
    End Sub

    Private Shared Function TrySetExportScope(window As AutomationElement, desired As String, L As Action(Of String)) As Boolean
        Dim combo = FindNearestToRight(window, ControlType.ComboBox, LblExportScope)
        If combo Is Nothing Then Return False

        Dim candidates = GetExportScopeCandidates(desired)

        Dim ok = SelectComboItemRobust(combo, candidates)
        If Not ok Then Return False

        Return True
    End Function

    Private Shared Function GetExportScopeCandidates(desired As String) As List(Of String)
        Dim d = (If(desired, "")).Trim().ToLowerInvariant()
        If d.Length = 0 OrElse d.Contains("current") OrElse d.Contains("view") OrElse d.Contains("현재") OrElse d.Contains("뷰") Then
            Return New List(Of String) From {"Current view", "Current View", "View", "현재 뷰", "현재뷰"}
        End If
        If d.Contains("select") OrElse d.Contains("선택") Then
            Return New List(Of String) From {"Selection", "선택"}
        End If
        If d.Contains("entire") OrElse d.Contains("model") OrElse d.Contains("전체") OrElse d.Contains("all") Then
            Return New List(Of String) From {"Entire model", "Entire Model", "전체 모델", "전체모델", "All"}
        End If
        Return New List(Of String) From {desired, "Current View", "View", "현재 뷰"}
    End Function

    Private Shared Sub TrySelectComboToRightOfLabel(window As AutomationElement,
                                                 labelCandidates As IEnumerable(Of String),
                                                 desiredCandidates As IEnumerable(Of String),
                                                 L As Action(Of String))
        Dim combo = FindNearestToRight(window, ControlType.ComboBox, labelCandidates)
        If combo Is Nothing Then Exit Sub
        SelectComboItemRobust(combo, desiredCandidates)
    End Sub

    Private Shared Function SelectComboItemRobust(combo As AutomationElement, desiredCandidates As IEnumerable(Of String)) As Boolean
        Try
            Dim exp = TryCast(combo.GetCurrentPattern(ExpandCollapsePattern.Pattern), ExpandCollapsePattern)
            If exp IsNot Nothing Then exp.Expand()

            Dim items = combo.FindAll(TreeScope.Descendants, New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem))
            For Each d In desiredCandidates
                If String.IsNullOrWhiteSpace(d) Then Continue For
                For i As Integer = 0 To items.Count - 1
                    Dim it = TryCast(items(i), AutomationElement)
                    If it Is Nothing Then Continue For
                    Dim nm = SafeName(it)
                    If nm.IndexOf(d, StringComparison.OrdinalIgnoreCase) >= 0 Then
                        SelectItem(it)
                        If exp IsNot Nothing Then exp.Collapse()
                        Return True
                    End If
                Next
            Next

            If exp IsNot Nothing Then exp.Collapse()
        Catch
        End Try
        Return False
    End Function

    Private Shared Sub SelectItem(item As AutomationElement)
        Try
            Dim si = TryCast(item.GetCurrentPattern(SelectionItemPattern.Pattern), SelectionItemPattern)
            If si IsNot Nothing Then si.Select() Else Invoke(item)
        Catch
            Invoke(item)
        End Try
    End Sub

    Private Shared Function TrySetAndVerifyFactor(window As AutomationElement,
                                               desired As Double,
                                               verify As Boolean,
                                               L As Action(Of String)) As Boolean
        Dim edit = FindNearestToRight(window, ControlType.Edit, LblFacetingCandidates)
        If edit Is Nothing OrElse Not SupportsWritableValuePattern(edit) Then
            Return False
        End If

        Dim s = desired.ToString(CultureInfo.InvariantCulture)
        If Not SetValue(edit, s) Then Return False
        If Not verify Then Return True

        Dim readBack As Double
        If TryReadDouble(edit, readBack) Then
            Return Math.Abs(readBack - desired) <= 0.01
        End If
        Return False
    End Function

    Private Shared Function FindNearestToRight(window As AutomationElement,
                                             targetType As ControlType,
                                             labelCandidates As IEnumerable(Of String)) As AutomationElement
        Dim label = FindAnyLabelContains(window, labelCandidates)
        If label Is Nothing Then Return Nothing

        Dim lr = label.Current.BoundingRectangle
        Dim targets = window.FindAll(TreeScope.Descendants, New PropertyCondition(AutomationElement.ControlTypeProperty, targetType))

        Dim best As AutomationElement = Nothing
        Dim bestScore As Double = Double.MaxValue

        For i As Integer = 0 To targets.Count - 1
            Dim t = TryCast(targets(i), AutomationElement)
            If t Is Nothing Then Continue For
            Dim tr = t.Current.BoundingRectangle
            If tr.Width <= 0 OrElse tr.Height <= 0 Then Continue For

            Dim dx = tr.Left - lr.Right
            If dx < -5 Then Continue For

            Dim lmidY = (lr.Top + lr.Bottom) / 2.0
            Dim tmidY = (tr.Top + tr.Bottom) / 2.0
            Dim dy = Math.Abs(tmidY - lmidY)

            Dim score = dx * dx + dy * dy
            If score < bestScore Then
                bestScore = score
                best = t
            End If
        Next

        Return best
    End Function

    Private Shared Function FindAnyLabelContains(window As AutomationElement, candidates As IEnumerable(Of String)) As AutomationElement
        Dim types As ControlType() = {ControlType.Text, ControlType.Custom}
        For Each ct In types
            Dim els = window.FindAll(TreeScope.Descendants, New PropertyCondition(AutomationElement.ControlTypeProperty, ct))
            For i As Integer = 0 To els.Count - 1
                Dim el = TryCast(els(i), AutomationElement)
                If el Is Nothing Then Continue For
                Dim nm = SafeName(el)
                If String.IsNullOrWhiteSpace(nm) Then Continue For
                If candidates.Any(Function(c) nm.IndexOf(c, StringComparison.OrdinalIgnoreCase) >= 0) Then
                    Return el
                End If
            Next
        Next
        Return Nothing
    End Function

    Private Shared Function SupportsWritableValuePattern(el As AutomationElement) As Boolean
        Try
            Dim p As Object = Nothing
            If Not el.TryGetCurrentPattern(ValuePattern.Pattern, p) Then Return False
            Dim vp = TryCast(p, ValuePattern)
            If vp Is Nothing Then Return False
            Return Not vp.Current.IsReadOnly
        Catch
            Return False
        End Try
    End Function

    Private Shared Function TryReadDouble(el As AutomationElement, ByRef value As Double) As Boolean
        Try
            Dim vp = TryCast(el.GetCurrentPattern(ValuePattern.Pattern), ValuePattern)
            If vp Is Nothing Then Return False
            Dim s = vp.Current.Value
            If Double.TryParse(s, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.InvariantCulture, value) Then Return True
            If Double.TryParse(s, NumberStyles.Float Or NumberStyles.AllowThousands, CultureInfo.CurrentCulture, value) Then Return True
        Catch
        End Try
        Return False
    End Function

    Private Shared Function SetValue(el As AutomationElement, value As String) As Boolean
        Try
            Dim vp = TryCast(el.GetCurrentPattern(ValuePattern.Pattern), ValuePattern)
            If vp Is Nothing Then Return False
            If vp.Current.IsReadOnly Then Return False
            vp.SetValue(value)
            Return True
        Catch
            Return False
        End Try
    End Function

    Private Shared Sub Invoke(el As AutomationElement)
        If el Is Nothing Then Return
        Try
            Dim ip = TryCast(el.GetCurrentPattern(InvokePattern.Pattern), InvokePattern)
            If ip IsNot Nothing Then
                ip.Invoke()
                Return
            End If
        Catch
        End Try
    End Sub

    Private Shared Function SafeName(el As AutomationElement) As String
        Try
            Return If(el.Current.Name, "")
        Catch
            Return ""
        End Try
    End Function

    Private Shared Sub CloseByOk(dlg As AutomationElement)
        Dim okBtn = FindFirstClickableByContains(dlg, OkButtonCandidates)
        If okBtn IsNot Nothing Then ClickNoBlock(okBtn, Sub(_unused As String) Return)
    End Sub

    Private Shared Sub CloseByCancel(dlg As AutomationElement)
        Dim cancelBtn = FindFirstClickableByContains(dlg, CancelButtonCandidates)
        If cancelBtn IsNot Nothing Then ClickNoBlock(cancelBtn, Sub(_unused As String) Return)
    End Sub

    Private Shared Sub WaitUntilClosed(window As AutomationElement, timeoutMs As Integer)
        Dim until = DateTime.UtcNow.AddMilliseconds(timeoutMs)
        While DateTime.UtcNow < until
            Try
                Dim r = window.Current.BoundingRectangle
            Catch ex As ElementNotAvailableException
                Return
            Catch
                Return
            End Try
            Thread.Sleep(200)
        End While
    End Sub

    Private Shared Sub HandleOverwriteDialogIfAny()
        Dim until = DateTime.UtcNow.AddMilliseconds(8000)
        While DateTime.UtcNow < until
            Try
                Dim w = FindTopLevelWindowAnyProcess(Function(win) FindFirstClickableByContains(win, YesButtonCandidates) IsNot Nothing)
                If w IsNot Nothing Then
                    Dim yesBtn = FindFirstClickableByContains(w, YesButtonCandidates)
                    If yesBtn IsNot Nothing Then
                        ClickNoBlock(yesBtn, Sub(_unused As String) Return)
                        Return
                    End If
                End If
            Catch
            End Try
            Thread.Sleep(150)
        End While
    End Sub

    Private Shared Sub DumpDialogButtons(dlg As AutomationElement, L As Action(Of String))
        Try
            Dim btns = dlg.FindAll(TreeScope.Descendants, New OrCondition(
        New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button),
        New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.SplitButton)
      ))
            For i As Integer = 0 To btns.Count - 1
                Dim b = TryCast(btns(i), AutomationElement)
                If b Is Nothing Then Continue For
                Dim nm = SafeName(b)
                If nm.Length > 0 Then L("DEBUG Button=" & nm & " type=" & b.Current.ControlType.ProgrammaticName)
            Next
        Catch
        End Try
    End Sub

    Private Shared Sub DumpDialogEdits(dlg As AutomationElement, L As Action(Of String))
        Try
            Dim els = dlg.FindAll(TreeScope.Descendants, New OrCondition(
        New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
        New PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ComboBox)
      ))
            For i As Integer = 0 To els.Count - 1
                Dim e = TryCast(els(i), AutomationElement)
                If e Is Nothing Then Continue For
                Dim nm = SafeName(e)
                Dim aid As String = ""
                Try : aid = If(e.Current.AutomationId, "") : Catch : aid = "" : End Try
                L("DEBUG Field type=" & e.Current.ControlType.ProgrammaticName & " name=""" & nm & """ aid=" & aid)
            Next
        Catch
        End Try
    End Sub

End Class



