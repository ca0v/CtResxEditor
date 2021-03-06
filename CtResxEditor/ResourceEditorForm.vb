Option Explicit On
Option Strict On
Option Infer On

Imports System.Globalization

Public Class ResourceEditorForm

#Region "Feedback"

    Private Class FeedbackListener
        Inherits TextWriterTraceListener

        Private ReadOnly m_MainForm As ResourceEditorForm

        Friend Sub New(ByVal piMainForm As ResourceEditorForm)
            Me.m_MainForm = piMainForm
        End Sub

        Public Overrides Sub WriteLine(ByVal message As String)
            Me.m_MainForm.Feedback(message, TraceLevel.Info)
        End Sub
    End Class

    Private Delegate Sub FeedbackDelagate(ByVal piValue As String, ByVal piMessageType As System.Diagnostics.TraceLevel)

    Private Sub Feedback(ByVal piValue As String, ByVal piMessageType As System.Diagnostics.TraceLevel)
        If String.IsNullOrEmpty(piValue) Then Return
        If Me.InvokeRequired Then
            Me.Invoke(New FeedbackDelagate(AddressOf Me.Feedback), piValue, piMessageType)
        Else
            Me.ToolStripStatusLabel1.Text = piValue
            Me.m_StatusHistory.Add(New Status(piValue, piMessageType))
            Select Case piMessageType
                Case TraceLevel.Error
                    Me.ToolStripStatusLabel1.Image = System.Drawing.SystemIcons.Error.ToBitmap
                Case TraceLevel.Info
                    Me.ToolStripStatusLabel1.Image = System.Drawing.SystemIcons.Information.ToBitmap
                Case TraceLevel.Warning
                    Me.ToolStripStatusLabel1.Image = System.Drawing.SystemIcons.Warning.ToBitmap
                Case Else
                    Me.ToolStripStatusLabel1.Image = System.Drawing.SystemIcons.WinLogo.ToBitmap
            End Select
        End If
    End Sub

    Private Sub Feedback(ByVal piException As Exception)
        Dim lvException As Exception = piException
        Dim lvTabs As String = String.Empty
        Do While Not IsNothing(lvException)
            Me.Feedback(lvTabs & lvException.Message, TraceLevel.Error)
            lvException = lvException.InnerException
            lvTabs &= "  "
        Loop
    End Sub

    Private Structure Status
        Public ReadOnly DateTime As DateTime
        Public ReadOnly Message As String
        Public ReadOnly Level As Diagnostics.TraceLevel
        Public Sub New(ByVal piMessage As String, ByVal piLevel As Diagnostics.TraceLevel)
            Me.DateTime = Now
            Me.Message = piMessage
            Me.Level = piLevel
        End Sub
    End Structure

    Private m_StatusHistory As New Generic.List(Of Status)

    Private Sub ShowStatusHistory()
        Dim lvFilterText As New TextBox
        lvFilterText.Dock = DockStyle.Top
        Dim lvDataGridView As New DataGridView
        Dim lvTable As New DataTable
        lvTable.Columns.Add("Time", GetType(DateTime))
        lvTable.Columns.Add("Status", GetType(String))
        lvTable.Columns.Add("Level", GetType(String))
        For Each lvItem As Status In Me.m_StatusHistory
            lvTable.Rows.Add(lvItem.DateTime, lvItem.Message, lvItem.Level.ToString)
        Next

        lvDataGridView.DataSource = lvTable
        Dim lvForm As New Form
        lvForm.Controls.Add(lvDataGridView)
        lvForm.Controls.Add(lvFilterText)
        lvDataGridView.Dock = DockStyle.Fill
        lvForm.Size = New Drawing.Size(400, 380)
        lvForm.Text = "Status History"
        FilterUtility.RegisterFilter(lvFilterText, lvTable)
        lvForm.Show()
        lvDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
    End Sub

    Private Sub ToolStripDropDownButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripDropDownButton1.Click
        ShowStatusHistory()
    End Sub

#End Region

#Region "Menu Management"

    Private Sub RefreshFileMenuItems()
        Me.RefreshFileHistory()
    End Sub

    Private Sub RefreshToolMenuItems()
        GenerateResxFilesToolStripMenuItem.Enabled = Not IsNothing(Me.CurrentDataSet)
        AddLanguageToolStripMenuItem.Enabled = Not IsNothing(Me.CurrentDataSet)
    End Sub

    Private Sub RefreshViewMenuItems()
        RowsWithBlanksToolStripMenuItem.Enabled = Not IsNothing(Me.CurrentDataSet)
        Dim lvFastFilter = FilterUtility.FindFilter(Me.FastFilterTextBox)
        If IsNothing(lvFastFilter) Then
            ' do nothing
        Else
            Me.RowsWithBlanksToolStripMenuItem.Checked = lvFastFilter.HasNullRow
        End If
    End Sub

    Private Sub FileToolStripMenuItem_DropDownOpening(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FileToolStripMenuItem.DropDownOpening
        Me.RefreshFileMenuItems()
    End Sub

    Private Sub ToolsToolStripMenuItem_DropDownOpening(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolsToolStripMenuItem.DropDownOpening
        Me.RefreshToolMenuItems()
    End Sub

    Private Sub FileClickHandler(ByVal sender As Object, ByVal e As EventArgs)
        Dim lvMenuItem As ToolStripMenuItem = DirectCast(sender, ToolStripMenuItem)
        Dim lvFile As String = lvMenuItem.Text
        If Not IO.File.Exists(lvFile) Then
            Me.Feedback(String.Format("{0} does not exist", lvFile), TraceLevel.Warning)
            Me.RefreshFileHistory()
        Else
            Me.ForceFileDialogUtility(lvFile).FileOpen(lvFile)
        End If
    End Sub

    Private Sub RefreshFileHistory()
        Dim lvMenuItem As Windows.Forms.ToolStripMenuItem = Me.FileHistoryToolStripMenuItem
        lvMenuItem.DropDownItems.Clear()
        Dim lvFileHistory As Specialized.StringCollection = My.Settings.FileHistory
        If Not IsNothing(lvFileHistory) Then
            Dim lvIndex As Int32 = 0
            Do While (lvIndex < lvFileHistory.Count)
                Dim lvFile As String = lvFileHistory(lvIndex)
                If IO.File.Exists(lvFile) Then
                    Dim lvChildItem As New Windows.Forms.ToolStripMenuItem(lvFile, Nothing, AddressOf FileClickHandler)
                    lvChildItem.Text = lvFile
                    lvMenuItem.DropDownItems.Add(lvChildItem)
                    lvIndex += 1
                Else
                    lvFileHistory.RemoveAt(lvIndex)
                End If
            Loop
        End If
    End Sub

    Private Sub InsertFileToFrontOfHistoryList(ByVal piFileName As String)
        If IsNothing(My.Settings.FileHistory) Then
            My.Settings.FileHistory = New Specialized.StringCollection
        End If
        Dim lvFileHistory As Specialized.StringCollection = My.Settings.FileHistory
        Do While lvFileHistory.Contains(piFileName)
            lvFileHistory.Remove(piFileName)
        Loop
        lvFileHistory.Insert(0, piFileName)
    End Sub

#End Region

#Region "File Open/Save"

    Private Const cXmlFileExt As String = ".resxml"
    Private m_FileSaveHelper As FileDialogUtility

    Private Function ForceFileDialogUtility(Optional ByVal piFileName As String = "") As FileDialogUtility
        If IsNothing(Me.m_FileSaveHelper) Then
            Me.m_FileSaveHelper = New FileDialogUtility(piFileName)
            AddHandler Me.m_FileSaveHelper.FileSaving, AddressOf Me.FileSavingHandler
            AddHandler Me.m_FileSaveHelper.FileOpening, AddressOf Me.FileOpeningHandler
            Me.m_FileSaveHelper.AddFileExtension("All Languages", cXmlFileExt)
            Me.m_FileSaveHelper.AddFileExtension(".NET ResX File", ".resx")
            Me.m_FileSaveHelper.AddFileExtension(".NET Resources File", ".resources")
            Me.m_FileSaveHelper.AddFileExtension("WiX Localization File", ".wxl")
        End If
        Return Me.m_FileSaveHelper
    End Function

    Private Sub FileSavingHandler(ByVal sender As Object, ByVal e As FileDialogUtility.FileEventArgs)
        Me.SaveFile(e.FileName)
    End Sub

    Private Sub FileOpeningHandler(ByVal sender As Object, ByVal e As FileDialogUtility.FileEventArgs)
        Dim lvFileName As String = String.Empty
        Dim lvCountry As String = String.Empty
        Dim lvCulture As String = String.Empty
        Dim lvExt As String = String.Empty
        ResUtility.ExtractCultureInfo(IO.Path.GetFileName(e.FileName), lvFileName, lvCountry, lvCulture, lvExt)
        If cXmlFileExt.Equals(IO.Path.GetExtension(e.FileName), StringComparison.InvariantCultureIgnoreCase) Then
            Me.OpenFile(e.FileName)
        Else
            Me.ImportFile(e.FileName)
            Me.Text = String.Format("{1}:{0} - Resource Editor", lvFileName, lvExt)
            e.FileName = String.Format("{0}{1}", lvFileName, cXmlFileExt)
        End If
        Me.m_ExportToJsonText = lvFileName
        Me.m_ExportToResxText = lvFileName
        Me.m_ExportToResxText = lvFileName
        Me.m_ExportToWxlText = lvFileName
    End Sub

    Private Function ForceDataSet() As NewDataSet
        If IsNothing(Me.CurrentDataSet) Then
            NewResources()
        End If
        Return Me.CurrentDataSet
    End Function

    Private Function CurrentDataSet() As NewDataSet
        If IsNothing(Me.DataGridView1.DataSource) Then
            Return Nothing
        End If
        Return DirectCast(DirectCast(Me.DataGridView1.DataSource, DataTable).DataSet, NewDataSet)
    End Function

    Private Sub SaveFile(ByVal piFileName As String)
        Debug.Assert(cXmlFileExt.Equals(IO.Path.GetExtension(piFileName)), "Invalid file extension")
        Dim lvDataSet As NewDataSet = CurrentDataSet()
        lvDataSet.ResxData.WriteXml(piFileName, XmlWriteMode.WriteSchema)
        lvDataSet.AcceptChanges()
        Me.InsertFileToFrontOfHistoryList(piFileName)
        Me.Text = String.Format("{0} - Resource Editor", IO.Path.GetFileName(piFileName))
    End Sub

    Private Sub AddFile(ByVal piFileName As String)
        Dim lvDataSet As NewDataSet = Me.ForceDataSet
        Select Case IO.Path.GetExtension(piFileName).ToUpper
            Case ".WXL"
                WixUtility.MergeWxlDataSet(piFileName, lvDataSet)
            Case ".RESX"
                ResUtility.MergeResXDataSet(piFileName, lvDataSet, ResUtility.OverwriteOptions.PromptForOverwrite)
            Case ".RESOURCES"
                ResUtility.MergeResourceDataSet(piFileName, lvDataSet)
            Case Else
                Debug.Fail(String.Format("{0} is an unexpected extension", IO.Path.GetExtension(piFileName)))
        End Select
    End Sub

    Private Sub AddDirectory(ByVal piDirectoryName As String)
        Dim lvDataSet As NewDataSet = Me.ForceDataSet
        For Each lvFileName As String In IO.Directory.GetFiles(piDirectoryName, "*.properties", IO.SearchOption.AllDirectories)
            Dim lvResourceName As String = IO.Path.GetFileNameWithoutExtension(IO.Path.GetDirectoryName(lvFileName))
            lvResourceName = lvResourceName.Substring(0, lvResourceName.IndexOf("-"))
            Dim lvCulture As New System.Globalization.CultureInfo(Int32.Parse(lvResourceName))
            ResUtility.MergeNameValuePairDataSet(lvFileName, lvDataSet, lvCulture.Name)
        Next
    End Sub

    Private m_DataSet As NewDataSet

    Private Sub OpenFile(ByVal piFileName As String)
        If Me.DataChanged Then
            Select Case PromptToSaveChanges()
                Case Windows.Forms.DialogResult.Yes
                    If Not Me.ForceFileDialogUtility.FileSave() Then
                        Return
                    End If
                Case Windows.Forms.DialogResult.No
                    ' do nothing
                Case Windows.Forms.DialogResult.Cancel
                    Return
            End Select
        End If
        Dim lvDataSet As New NewDataSet
        lvDataSet.ResxData.Rows.Clear()
        lvDataSet.ReadXml(piFileName, XmlReadMode.InferSchema)
        lvDataSet.AcceptChanges()
        Me.m_DataSet = lvDataSet
        Me.m_ExportToResxText = IO.Path.GetFileNameWithoutExtension(piFileName)
        Me.DataGridView1.DataSource = lvDataSet.ResxData
        Me.InsertFileToFrontOfHistoryList(piFileName)
        FilterUtility.RegisterFilter(Me.FastFilterTextBox, lvDataSet.ResxData)
        Me.Text = String.Format("{0} - Resource Editor", IO.Path.GetFileName(piFileName))
    End Sub

    Private Sub ImportFile(ByVal piFileName As String)
        Dim lvDataSet As New NewDataSet
        Dim lvPath As String = IO.Path.GetDirectoryName(piFileName)
        Dim lvFileName As String = String.Empty
        Dim lvCountry As String = String.Empty
        Dim lvCulture As String = String.Empty
        Dim lvExt As String = String.Empty
        ResUtility.ExtractCultureInfo(IO.Path.GetFileName(piFileName), lvFileName, lvCountry, lvCulture, lvExt)
        ResUtility.MergeDataSet(lvPath, lvFileName, lvExt, lvDataSet)
        Me.DataGridView1.DataSource = lvDataSet.ResxData
        FilterUtility.RegisterFilter(Me.FastFilterTextBox, lvDataSet.ResxData)
    End Sub

#End Region

#Region "Auto Update"
    Private m_AutoUpdate As New AutoUpdate("CtResxEditor.MSI")
    Private m_UpdateAvailableHandlerReturned As Boolean = False
    Private m_InstallOnExit As Boolean = False

    Private Sub UpdateAvailableHandler(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)
        m_UpdateAvailableHandlerReturned = True
        m_InstallOnExit = CBool(e.Result)
    End Sub

    Private Sub ApplicationExitHandler(ByVal sender As Object, ByVal e As EventArgs)
        Do While Not m_UpdateAvailableHandlerReturned
            MsgBox("Waiting for the installer download to complete")
            Threading.Thread.Sleep(1000)
        Loop
        If (m_InstallOnExit) Then
            m_AutoUpdate.StartInstaller()
        End If
    End Sub
#End Region

    'TODO: add drag-and-drop of resx to merge with current data (add a datestamp column too)
#Region "Drag Drop"

    Private Enum DropMode
        None = 0
        Folder = 1
        File = 2
    End Enum
    Private m_DropMode As DropMode = DropMode.None

    Protected Overrides Sub OnDragEnter(ByVal drgevent As System.Windows.Forms.DragEventArgs)
        MyBase.OnDragEnter(drgevent)
        If drgevent.Data.GetDataPresent("FileDrop", False) Then
            For Each lvFileName As String In DirectCast(drgevent.Data.GetData("FileDrop", False), String())
                If System.IO.File.Exists(lvFileName) Then
                    If Me.m_DropMode = DropMode.Folder Then
                        Debug.Fail("Cannot mix files and folders, clear the grid and try again")
                        Exit Sub
                    End If
                ElseIf System.IO.Directory.Exists(lvFileName) Then
                    If Me.m_DropMode = DropMode.File Then
                        Debug.Fail("Cannot mix files and folders, clear the grid and try again")
                        Exit Sub
                    End If
                End If
            Next
        Else
            Exit Sub
        End If
        drgevent.Effect = Windows.Forms.DragDropEffects.Copy
    End Sub

    Protected Overrides Sub OnDragDrop(ByVal drgevent As System.Windows.Forms.DragEventArgs)
        MyBase.OnDragDrop(drgevent)
        For Each lvFileName As String In DirectCast(drgevent.Data.GetData("FileDrop", False), String())
            Select Case Me.m_DropMode
                Case DropMode.None
                    If IO.File.Exists(lvFileName) Then
                        Me.AddFile(lvFileName)
                        Me.m_DropMode = DropMode.File
                    ElseIf IO.Directory.Exists(lvFileName) Then
                        Me.AddDirectory(lvFileName)
                        Me.m_DropMode = DropMode.Folder
                    End If
                Case DropMode.File
                    If IO.File.Exists(lvFileName) Then
                        Me.AddFile(lvFileName)
                        Me.m_DropMode = DropMode.File
                    End If
                Case DropMode.Folder
                    If IO.Directory.Exists(lvFileName) Then
                        Me.AddDirectory(lvFileName)
                        Me.m_DropMode = DropMode.Folder
                    End If
            End Select
        Next
    End Sub

#End Region

    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        MyBase.OnLoad(e)
        Me.AllowDrop = True
        Try
            m_AutoUpdate.StartIsUpdateAvailable(AddressOf UpdateAvailableHandler)
            AddHandler Application.ApplicationExit, AddressOf ApplicationExitHandler
        Catch ex As Exception
            My.Application.Log.WriteException(ex)
        End Try
        If (0 < My.Application.CommandLineArgs.Count) Then
            Dim lvFileName = My.Application.CommandLineArgs(0)
            If IO.File.Exists(lvFileName) Then
                Me.ForceFileDialogUtility.FileOpen(lvFileName)
            Else
                Debug.Fail(String.Format("File {0} does not exist", lvFileName))
            End If
        End If
        AutoHotKey.AssignHotKeys(Me)
    End Sub

    Private Function DataChanged() As Boolean
        If IsNothing(Me.m_DataSet) Then Return False
        Return Me.m_DataSet.HasChanges()
    End Function

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        MyBase.OnClosing(e)
        My.Settings.Save()
        If Me.DataChanged Then
            Select Case PromptToSaveChanges
                Case Windows.Forms.DialogResult.Yes
                    If Not Me.ForceFileDialogUtility.FileSave() Then
                        e.Cancel = True
                    End If
                Case Windows.Forms.DialogResult.No
                    ' do nothing
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub

    Private Function PromptToSaveChanges() As DialogResult
        Return System.Windows.Forms.MessageBox.Show("Save Changes?", "Changes have been made", MessageBoxButtons.YesNoCancel)
    End Function

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        ForceFileDialogUtility.FileOpen()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        ForceFileDialogUtility.FileSave()
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsToolStripMenuItem.Click
        ForceFileDialogUtility.FileSaveAs()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        NewResources()
    End Sub

    Private Sub NewResources()
        Dim lvDataSet As New NewDataSet
        Me.DataGridView1.DataSource = lvDataSet.ResxData
        FilterUtility.RegisterFilter(Me.FastFilterTextBox, lvDataSet.ResxData)
    End Sub

    Private Sub AddLanguageToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddLanguageToolStripMenuItem.Click
        AddLanguage(Me.ForceDataSet.ResxData)
    End Sub

    Private Shared Sub AddLanguage(ByVal piTable As NewDataSet.ResxDataDataTable)
        Dim lvLanguagePicker As New LanguagePicker
        For Each lvColumn As DataColumn In piTable.Columns
            Dim lvColumnName As String = lvColumn.ColumnName
            Try
                Dim lvCulture As New System.Globalization.CultureInfo(lvColumnName)
                lvLanguagePicker.SelectCulture(lvCulture)
            Catch
                ' ignore failures
            End Try
        Next
        Dim lvCultures As Generic.List(Of CultureInfo) = lvLanguagePicker.PickCultures()
        If Not IsNothing(lvCultures) Then
            For Each lvCulture As CultureInfo In lvCultures
                ResUtility.ForceCultureColumn(lvCulture.Name, piTable)
            Next
        End If
    End Sub

#Region "Quick Input - Find Node"

    Private Delegate Sub QuickInputDelegate(ByVal piValue As String)
    Private m_QuickInputDelegate As QuickInputDelegate
    Private m_Accepted As Boolean

    Private Sub QuickInputToolStripTextBox_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles QuickInputToolStripTextBox.KeyDown
        If e.KeyCode = Windows.Forms.Keys.Enter Then
            Me.m_Accepted = True
            e.SuppressKeyPress = True
            Me.LeaveQuickInputBox()
        End If
    End Sub

    Private Sub LeaveQuickInputBox()
        Me.QuickInputToolStripTextBox.Visible = False
        Dim lvDelegate As QuickInputDelegate = Me.m_QuickInputDelegate
        If Not IsNothing(lvDelegate) Then
            Me.m_QuickInputDelegate = Nothing
            If Me.m_Accepted Then
                lvDelegate(Me.QuickInputToolStripTextBox.Text)
            End If
        End If
    End Sub

    Private Sub QuickInputToolStripTextBox_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles QuickInputToolStripTextBox.LostFocus
        Me.m_Accepted = False
        Me.LeaveQuickInputBox()
    End Sub

    Private m_ExportToResxText As String = String.Empty

    Private m_SearchResults() As TreeNode = New TreeNode() {}

    Private Sub SetExportToResxText(ByVal piValue As String)
        Me.m_ExportToResxText = piValue
        If Not String.IsNullOrEmpty(Me.m_ExportToResxText) Then
            ExportToResxFiles(Me.m_ExportToResxText, False)
            ExportToResxFiles(String.Format("{0}.Missing", Me.m_ExportToResxText), True)
        End If
    End Sub

    Private Sub SetExportToResourceText(ByVal piValue As String)
        Me.m_ExportToResxText = piValue
        If Not String.IsNullOrEmpty(Me.m_ExportToResxText) Then
            ExportToResourceFiles(Me.m_ExportToResxText, False)
            ExportToResourceFiles(String.Format("{0}.Missing", Me.m_ExportToResxText), True)
        End If
    End Sub

    Private Sub SetExportToPropertiesText(ByVal piValue As String)
        Me.m_ExportToResxText = piValue
        If Not String.IsNullOrEmpty(Me.m_ExportToResxText) Then
            ExportToPropertiesFiles(Me.m_ExportToResxText, False)
            ExportToPropertiesFiles(String.Format("{0}.Missing", Me.m_ExportToResxText), True)
        End If
    End Sub

    Private Sub SetExportToWxlText(ByVal piValue As String)
        Me.m_ExportToWxlText = piValue
        If Not String.IsNullOrEmpty(Me.m_ExportToWxlText) Then
            ExportToWxlFiles(Me.m_ExportToWxlText, False)
            ExportToWxlFiles(String.Format("{0}.Missing", Me.m_ExportToWxlText), True)
        End If
    End Sub

    Private Sub SetExportToJsonText(ByVal piValue As String)
        Me.m_ExportToJsonText = piValue
        If Not String.IsNullOrEmpty(Me.m_ExportToJsonText) Then
            ExportToJsonFiles(Me.m_ExportToJsonText, False)
            ExportToJsonFiles(String.Format("{0}.Missing", Me.m_ExportToJsonText), True)
        End If
    End Sub

    Private Function ForceWorkingDirectory() As String
        Dim lvRootPath As String = My.Settings.WorkingDirectory
        If String.IsNullOrEmpty(lvRootPath) OrElse Not IO.Directory.Exists(lvRootPath) Then
            lvRootPath = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData
            My.Settings.WorkingDirectory = lvRootPath
            My.Settings.Save()
        End If
        Return lvRootPath
    End Function

    Private Sub ExportToResxFiles(ByVal piRootName As String, ByVal piNullOnly As Boolean)
        Dim lvDataset As NewDataSet = Me.CurrentDataSet
        Debug.Assert(Not IsNothing(lvDataset), "No dataset defined")
        Dim lvRootPath As String = ForceWorkingDirectory()
        Dim lvTable As NewDataSet.ResxDataDataTable = lvDataset.ResxData
        Dim lvRows = lvTable.Select(String.Empty, "Id")
        Dim lvEnglishColumn As DataColumn = lvTable.Columns("en")
        For Each lvColumn As DataColumn In lvTable.Columns
            If lvColumn Is lvTable.idColumn Then
                Continue For
            End If
            If lvColumn Is lvTable.VersionColumn Then
                Continue For
            End If
            Dim lvResxFile As String
            If lvColumn Is lvTable.DefaultColumn Then
                lvResxFile = String.Format("{0}.resx", piRootName)
            Else
                lvResxFile = String.Format("{0}.{1}.resx", piRootName, lvColumn.ColumnName)
            End If
            lvResxFile = IO.Path.Combine(lvRootPath, lvResxFile)
            Using lvWriter As New System.Resources.ResXResourceWriter(lvResxFile)
                If piNullOnly Then
                    For Each lvRow As NewDataSet.ResxDataRow In lvRows
                        If DBNull.Value Is lvRow(lvColumn) Then
                            If lvRow.Is_DefaultNull Then
                                If IsNothing(lvEnglishColumn) Then
                                    lvWriter.AddResource(lvRow.id, String.Empty)
                                ElseIf lvRow(lvEnglishColumn) Is DBNull.Value Then
                                    lvWriter.AddResource(lvRow.id, String.Empty)
                                Else
                                    lvWriter.AddResource(lvRow.id, CStr(lvRow(lvEnglishColumn)))
                                End If
                            Else
                                lvWriter.AddResource(lvRow.id, lvRow._Default)
                            End If
                        End If
                    Next
                Else
                    For Each lvRow As NewDataSet.ResxDataRow In lvRows
                        If lvRow(lvColumn) Is DBNull.Value Then
                            ' do nothing
                        Else
                            lvWriter.AddResource(lvRow.id, CStr(lvRow(lvColumn)))
                        End If

                    Next
                End If
                lvWriter.Close()
            End Using
        Next
        Process.Start(lvRootPath)
    End Sub

    Private Sub ExportToResourceFiles(ByVal piRootName As String, ByVal piNullOnly As Boolean)
        Dim lvDataset As NewDataSet = Me.CurrentDataSet
        Debug.Assert(Not IsNothing(lvDataset), "No dataset defined")
        Dim lvRootPath As String = ForceWorkingDirectory()
        Dim lvTable As NewDataSet.ResxDataDataTable = lvDataset.ResxData
        Dim lvRows = lvTable.Select(String.Empty, "Id")
        Dim lvEnglishColumn As DataColumn = lvTable.Columns("en")
        For Each lvColumn As DataColumn In lvTable.Columns
            If lvColumn Is lvTable.idColumn Then
                Continue For
            End If
            If lvColumn Is lvTable.VersionColumn Then
                Continue For
            End If
            Dim lvResxFile As String
            If lvColumn Is lvTable.DefaultColumn Then
                lvResxFile = String.Format("{0}.resources", piRootName)
            Else
                lvResxFile = String.Format("{0}.{1}.resources", piRootName, lvColumn.ColumnName)
            End If
            lvResxFile = IO.Path.Combine(lvRootPath, lvResxFile)
            Using lvWriter As New System.Resources.ResourceWriter(lvResxFile)
                If piNullOnly Then
                    For Each lvRow As NewDataSet.ResxDataRow In lvRows
                        If DBNull.Value Is lvRow(lvColumn) Then
                            If lvRow.Is_DefaultNull Then
                                If IsNothing(lvEnglishColumn) Then
                                    lvWriter.AddResource(lvRow.id, String.Empty)
                                ElseIf lvRow(lvEnglishColumn) Is DBNull.Value Then
                                    lvWriter.AddResource(lvRow.id, String.Empty)
                                Else
                                    lvWriter.AddResource(lvRow.id, CStr(lvRow(lvEnglishColumn)))
                                End If
                            Else
                                lvWriter.AddResource(lvRow.id, lvRow._Default)
                            End If
                        End If
                    Next
                Else
                    For Each lvRow As NewDataSet.ResxDataRow In lvRows
                        If lvRow(lvColumn) Is DBNull.Value Then
                            ' do nothing
                        Else
                            lvWriter.AddResource(lvRow.id, CStr(lvRow(lvColumn)))
                        End If

                    Next
                End If
                lvWriter.Close()
            End Using
        Next
        Process.Start(lvRootPath)
    End Sub

    Private Sub ExportToPropertiesFiles(ByVal piRootName As String, ByVal piNullOnly As Boolean)
        Dim lvDataset As NewDataSet = Me.CurrentDataSet
        Debug.Assert(Not IsNothing(lvDataset), "No dataset defined")
        Dim lvRootPath As String = ForceWorkingDirectory()
        Dim lvTable As NewDataSet.ResxDataDataTable = lvDataset.ResxData
        Dim lvRows = lvTable.Select(String.Empty, "Id")
        Dim lvEnglishColumn As DataColumn = lvTable.Columns("en")
        For Each lvColumn As DataColumn In lvTable.Columns
            If lvColumn Is lvTable.idColumn Then
                Continue For
            End If
            If lvColumn Is lvTable.VersionColumn Then
                Continue For
            End If
            Dim lvResxFile As String
            If lvColumn Is lvTable.DefaultColumn Then
                lvResxFile = IO.Path.Combine(lvRootPath, String.Format("{0}.properties", piRootName))
            Else
                Dim lvCultureInfo As New CultureInfo(lvColumn.ColumnName)
                Dim lvDirectory As String = IO.Path.Combine(lvRootPath, String.Format("{0} - {1}", lvCultureInfo.LCID, lvCultureInfo.EnglishName))
                IO.Directory.CreateDirectory(lvDirectory)
                lvResxFile = IO.Path.Combine(lvDirectory, String.Format("{0}.properties", piRootName))
            End If
            Using lvWriter = IO.File.CreateText(lvResxFile)
                If piNullOnly Then
                    For Each lvRow As NewDataSet.ResxDataRow In lvRows
                        If DBNull.Value Is lvRow(lvColumn) Then
                            If lvRow.Is_DefaultNull Then
                                If IsNothing(lvEnglishColumn) Then
                                    lvWriter.Write(lvRow.id, String.Empty)
                                ElseIf lvRow(lvEnglishColumn) Is DBNull.Value Then
                                    lvWriter.WriteLine(String.Format("{0}={1}", lvRow.id, String.Empty))
                                Else
                                    lvWriter.WriteLine(String.Format("{0}={1}", lvRow.id, CStr(lvRow(lvEnglishColumn))))
                                End If
                            Else
                                lvWriter.WriteLine(String.Format("{0}={1}", lvRow.id, lvRow._Default))
                            End If
                        End If
                    Next
                Else
                    For Each lvRow As NewDataSet.ResxDataRow In lvRows
                        If lvRow(lvColumn) Is DBNull.Value Then
                            ' do nothing
                        Else
                            lvWriter.WriteLine(String.Format("{0}={1}", lvRow.id, CStr(lvRow(lvColumn))))
                        End If

                    Next
                End If
                lvWriter.Close()
            End Using
        Next
        Process.Start(lvRootPath)
    End Sub

    Private Sub ExportToWxlFiles(ByVal piRootName As String, ByVal piNullOnly As Boolean)
        Dim lvDataset As NewDataSet = Me.CurrentDataSet
        Debug.Assert(Not IsNothing(lvDataset), "No dataset defined")
        Dim lvRootPath As String = ForceWorkingDirectory()
        Dim lvTable As NewDataSet.ResxDataDataTable = lvDataset.ResxData
        Dim lvRows = lvTable.Select(String.Empty, "Id")
        Dim lvEnglishColumn As DataColumn = lvTable.Columns("en-us")
        For Each lvColumn As DataColumn In lvTable.Columns
            If lvColumn Is lvTable.idColumn Then
                Continue For
            End If
            If lvColumn Is lvTable.VersionColumn Then
                Continue For
            End If
            Dim lvResxFile As String
            If lvColumn Is lvTable.DefaultColumn Then
                lvResxFile = String.Format("{0}.wxl", piRootName)
            Else
                lvResxFile = String.Format("{0}.{1}.wxl", piRootName, lvColumn.ColumnName)
            End If
            lvResxFile = IO.Path.Combine(lvRootPath, lvResxFile)
            Dim lvWriter As New System.Xml.XmlDocument
            Const cNS As String = "http://schemas.microsoft.com/wix/2006/localization"
            Dim lvRootNode As System.Xml.XmlElement = lvWriter.CreateElement("WixLocalization", cNS)
            lvRootNode.SetAttribute("Codepage", "1252")
            If lvColumn Is lvTable.DefaultColumn Then
                ' do nothing
            Else
                lvRootNode.SetAttribute("Culture", lvColumn.ColumnName)
            End If
            lvWriter.AppendChild(lvRootNode)
            If piNullOnly Then
                For Each lvRow As NewDataSet.ResxDataRow In lvRows
                    If DBNull.Value Is lvRow(lvColumn) Then
                        Dim lvNode As System.Xml.XmlElement = lvWriter.CreateElement("String", cNS)
                        lvNode.SetAttribute("Id", lvRow.id)
                        'lvNode.SetAttribute("Overridable", "Yes")
                        lvRootNode.AppendChild(lvNode)
                        If lvRow.Is_DefaultNull Then
                            If IsNothing(lvEnglishColumn) Then
                                ' do nothing
                            ElseIf lvRow(lvEnglishColumn) Is DBNull.Value Then
                                ' do nothing
                            Else
                                lvNode.InnerText = CStr(lvRow(lvEnglishColumn))
                            End If
                        Else
                            lvNode.InnerText = lvRow._Default
                        End If
                    End If
                Next
            Else
                For Each lvRow As NewDataSet.ResxDataRow In lvRows
                    If lvRow(lvColumn) Is DBNull.Value Then
                        ' do nothing
                    Else
                        Dim lvNode As System.Xml.XmlElement = lvWriter.CreateElement("String", cNS)
                        lvNode.SetAttribute("Id", lvRow.id)
                        'lvNode.SetAttribute("Overridable", "Yes")
                        lvRootNode.AppendChild(lvNode)
                        Dim lvValue As String = CStr(lvRow(lvColumn))
                        If String.IsNullOrEmpty(lvValue) Then
                            ' leave the inner text null
                        Else
                            lvNode.InnerText = lvValue
                        End If
                    End If

                Next
            End If
            lvWriter.Save(lvResxFile)
        Next
        Process.Start(lvRootPath)
    End Sub

    '﻿var localizedStrings = {
    '    "Core": {
    '        "BasicMessage" : "Hello, this is a localized string",
    '        "QuotedLabel" : "This isn't supposed to fail",
    '        "EscapedQuotesLabel" : "This has \"Escaped Quotes\""
    '    },
    '    "AssetManagement.Street" : {
    '        "Prompt" : "Are you sure?",
    '        "NameLabel" : "Name",
    '        "AnotherLabel" : "LocalizedLabel"
    '    }
    '}
    Private Sub ExportToJsonFiles(ByVal piRootName As String, ByVal piNullOnly As Boolean)
        If (piNullOnly) Then Return
        Dim lvDataset As NewDataSet = Me.CurrentDataSet
        Debug.Assert(Not IsNothing(lvDataset), "No dataset defined")
        Dim lvRootPath As String = ForceWorkingDirectory()
        Dim lvTable As NewDataSet.ResxDataDataTable = lvDataset.ResxData
        Dim lvRows = lvTable.Select(String.Empty, "Id")
        Dim lvEnglishColumn As DataColumn = lvTable.Columns("en-us")
        For Each lvColumn As DataColumn In lvTable.Columns
            If lvColumn Is lvTable.idColumn Then
                Continue For
            End If
            If lvColumn Is lvTable.VersionColumn Then
                Continue For
            End If
            Dim lvResxFile As String
            If lvColumn Is lvTable.DefaultColumn Then
                lvResxFile = String.Format("{0}.js", piRootName)
            Else
                lvResxFile = String.Format("{0}.{1}.js", piRootName, lvColumn.ColumnName)
            End If
            lvResxFile = IO.Path.Combine(lvRootPath, lvResxFile)

            Using lvWriter = IO.File.CreateText(lvResxFile)
                lvWriter.WriteLine("resx[{0}] = {{", QuoteIt(piRootName))
                For Each lvRow As NewDataSet.ResxDataRow In lvRows
                    If lvRow(lvColumn) Is DBNull.Value Then
                        ' do nothing
                    Else
                        ' "test": "test",
                        lvWriter.WriteLine("{0} : {1},", QuoteIt(lvRow.id), QuoteIt(CStr(lvRow(lvColumn))))
                    End If
                Next
                lvWriter.WriteLine("}")
                lvWriter.Close()
            End Using
        Next
        Process.Start(lvRootPath)
    End Sub

    Private Shared Function QuoteIt(ByVal piValue As String) As String
        Return String.Format("""{0}""", piValue.Replace("""", "\"""))
    End Function


    Private Sub ExportToResxFiles()
        Me.QuickInputToolStripTextBox.Visible = True
        Me.QuickInputToolStripTextBox.Text = Me.m_ExportToResxText
        Me.QuickInputToolStripTextBox.TextBox.SelectAll()
        Me.QuickInputToolStripTextBox.Focus()
        Me.m_QuickInputDelegate = AddressOf SetExportToResxText
    End Sub

    Private Sub ExportToResourceFiles()
        Me.QuickInputToolStripTextBox.Visible = True
        Me.QuickInputToolStripTextBox.Text = Me.m_ExportToResxText
        Me.QuickInputToolStripTextBox.TextBox.SelectAll()
        Me.QuickInputToolStripTextBox.Focus()
        Me.m_QuickInputDelegate = AddressOf SetExportToResourceText
    End Sub

    Private Sub ExportToPropertiesFiles()
        Me.QuickInputToolStripTextBox.Visible = True
        Me.QuickInputToolStripTextBox.Text = Me.m_ExportToResxText
        Me.QuickInputToolStripTextBox.TextBox.SelectAll()
        Me.QuickInputToolStripTextBox.Focus()
        Me.m_QuickInputDelegate = AddressOf SetExportToPropertiesText
    End Sub

    Private Sub ExportToWxlFiles()
        Me.QuickInputToolStripTextBox.Visible = True
        Me.QuickInputToolStripTextBox.Text = Me.m_ExportToWxlText
        Me.QuickInputToolStripTextBox.TextBox.SelectAll()
        Me.QuickInputToolStripTextBox.Focus()
        Me.m_QuickInputDelegate = AddressOf SetExportToWxlText
    End Sub
    Private m_ExportToWxlText As String = String.Empty

    Private Sub ExportToJsonFiles()
        Me.QuickInputToolStripTextBox.Visible = True
        Me.QuickInputToolStripTextBox.Text = Me.m_ExportToJsonText
        Me.QuickInputToolStripTextBox.TextBox.SelectAll()
        Me.QuickInputToolStripTextBox.Focus()
        Me.m_QuickInputDelegate = AddressOf SetExportToJsonText
    End Sub
    Private m_ExportToJsonText As String = String.Empty


    Private Sub GenerateResxFilesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenerateResxFilesToolStripMenuItem.Click
        Me.ExportToResxFiles()
    End Sub
#End Region

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Dim lvAbout As New AboutBox1
        lvAbout.ShowDialog()
    End Sub

    Private Sub OpenWorkingDirectoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenWorkingDirectoryToolStripMenuItem.Click
        Process.Start(Me.ForceWorkingDirectory)
    End Sub

    Private Sub RowsWithBlanksToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RowsWithBlanksToolStripMenuItem.Click
        Dim lvFilter = FilterUtility.FindFilter(Me.FastFilterTextBox)
        If Not IsNothing(lvFilter) Then
            lvFilter.HasNullRow = Not lvFilter.HasNullRow
        End If
    End Sub

    Private Sub ToolStripMenuItem1_DropDownOpening(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.DropDownOpening
        Me.RefreshViewMenuItems()
    End Sub

    Private Sub GenerateResourceFilesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenerateResourceFilesToolStripMenuItem.Click
        Me.ExportToResourceFiles()
    End Sub

    Private Sub GeneratePropertiesFilesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GeneratePropertiesFilesToolStripMenuItem.Click
        Me.ExportToPropertiesFiles()
    End Sub

    Private Sub GenerateWiXLocalizationFilesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenerateWiXLocalizationFilesToolStripMenuItem.Click
        Me.ExportToWxlFiles()
    End Sub

    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        Me.ExportToJsonFiles()
    End Sub

    Private Sub DataGridView1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        Dim lvGrid = TryCast(sender, DataGridView)
        If e.Control And e.KeyCode = Keys.V Then
            Dim lvData = Clipboard.GetDataObject()
            Dim lvData1 = Clipboard.GetData(DataFormats.Text)
            Dim lvData2 = CStr(Clipboard.GetData(DataFormats.CommaSeparatedValue))
            Dim lvData3 = Clipboard.GetData(DataFormats.Html)
            If (Not IsNothing(lvData2)) Then
                Dim lvFirstCell = lvGrid.SelectedCells(0)
                Dim lvColumnIndex = lvFirstCell.ColumnIndex
                Dim lvRowIndex = lvFirstCell.RowIndex
                Dim lvReader = New IO.StringReader(lvData2)
                Dim lvRow1 = lvReader.ReadLine()
                While Not String.IsNullOrEmpty(lvRow1)
                    Dim lvCells = lvRow1.Split(","c)
                    For x = 0 To lvCells.Length - 1
                        Dim lvCell = lvGrid(lvColumnIndex + x, lvRowIndex)
                        lvCell.Value = lvCells(x)
                    Next
                    lvRow1 = lvReader.ReadLine()
                    lvRowIndex += 1
                End While
            End If
        End If
    End Sub
End Class
