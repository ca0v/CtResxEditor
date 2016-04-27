Option Infer On
Option Explicit On
Option Strict On

'***************************************************
'* Paste this code into the caller class
'***************************************************
'#Region "Auto Update"
'Private m_AutoUpdate As New AutoUpdate([TODO: enter name of the msi file])
'Private m_UpdateAvailableHandlerReturned As Boolean = False
'Private m_InstallOnExit As Boolean = False

'Private Sub UpdateAvailableHandler(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs)
'    m_UpdateAvailableHandlerReturned = True
'    m_InstallOnExit = CBool(e.Result)
'End Sub

'Private Sub ApplicationExitHandler(ByVal sender As Object, ByVal e As EventArgs)
'    Do While Not m_UpdateAvailableHandlerReturned
'        MsgBox("Waiting for the installer download to complete")
'        Threading.Thread.Sleep(1000)
'    Loop
'    If (m_InstallOnExit) Then
'        m_AutoUpdate.StartInstaller()
'    End If
'End Sub
'#End Region

'***************************************************
'* Paste this code into the caller startup routine
'***************************************************
'Try
'    m_AutoUpdate.StartIsUpdateAvailable(AddressOf UpdateAvailableHandler)
'    AddHandler Application.ApplicationExit, AddressOf ApplicationExitHandler
'Catch ex As Exception
'    My.Application.Log.WriteException(ex)
'End Try

Imports System.ComponentModel

Friend Class AutoUpdate
    Private ReadOnly m_InstallerName As String
    Friend Sub New(ByVal piInstallerName As String)
        Me.m_InstallerName = piInstallerName
    End Sub

    Private Shared InfoRow As UpdateDataSet.UpdateInfoRow

    Private Shared m_DB As UpdateDataSet = Nothing
    Private Shared Function ForceDB() As UpdateDataSet
        If Not IsNothing(m_DB) Then
            Return m_DB
        End If
        m_DB = New UpdateDataSet
        Dim lvLocalFile = AllUsers("UpdateDataSet.xml")
        If IO.File.Exists(lvLocalFile) Then
            If 1 > Now.Subtract(IO.File.GetLastWriteTime(lvLocalFile)).TotalHours Then
                m_DB.ReadXml(lvLocalFile, XmlReadMode.IgnoreSchema)
                Return m_DB
            End If
        End If
        Dim lvURL = My.Settings.UpdateURL
        If Not String.IsNullOrEmpty(lvURL) Then
            Try
                m_DB.ReadXml(lvURL, XmlReadMode.IgnoreSchema)
            Catch ex As Exception
                My.Application.Log.WriteException(ex)
            End Try
        End If
        m_DB.WriteXml(lvLocalFile)
        Return m_DB
    End Function

    Private Function MsiFile() As String
        Return TempPath(Me.m_InstallerName)
    End Function

    Private Shared Function AllUsers(ByVal ParamArray piArgs() As String) As String
        Dim lvResult = My.Computer.FileSystem.SpecialDirectories.AllUsersApplicationData
        For Each lvItem In piArgs
            lvResult = IO.Path.Combine(lvResult, lvItem)
        Next
        Return lvResult
    End Function

    Private Shared Function TempPath(ByVal ParamArray piArgs() As String) As String
        Dim lvResult = IO.Path.GetTempPath
        For Each lvItem In piArgs
            lvResult = IO.Path.Combine(lvResult, lvItem)
        Next
        Return lvResult
    End Function

    Private Shared Function IsUpdateAvailable() As Boolean
        InfoRow = Nothing
        Dim lvRows = ForceDB.UpdateInfo.Select(String.Format("ProductName='{0}'", My.Application.Info.ProductName), "ProductVersion DESC")
        Dim lvThisVersion = My.Application.Info.Version
        For Each lvRow As UpdateDataSet.UpdateInfoRow In lvRows
            Dim lvVersion = New Version(lvRow.ProductVersion)
            If lvVersion > lvThisVersion Then
                InfoRow = lvRow
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub DeleteLocalInstaller()
        Dim lvFileInfo = New System.IO.FileInfo(MsiFile())
        If lvFileInfo.Exists Then
            lvFileInfo.Delete()
        End If
    End Sub

    Private Sub DownloadUpdate()
        Dim lvTarget = MsiFile()
        If Not IO.File.Exists(lvTarget) Then
            Dim lvURL = InfoRow.URL
            Using lvRequest = New System.Net.WebClient()
                lvRequest.DownloadFile(lvURL, lvTarget)
            End Using
        End If
    End Sub

    Friend Sub StartInstaller()
        Dim lvTarget = MsiFile()
        If IO.File.Exists(lvTarget) Then
            Process.Start("MSIEXEC.EXE", String.Format("/I ""{0}"" REINSTALL=ALL REINSTALLMODE=vomus", lvTarget))
        End If
    End Sub

    Friend Sub StartIsUpdateAvailable(ByVal piCallback As RunWorkerCompletedEventHandler)
        Dim lvWorker As New System.ComponentModel.BackgroundWorker()
        AddHandler lvWorker.DoWork, AddressOf DoStartIsUpdateAvailable
        AddHandler lvWorker.RunWorkerCompleted, piCallback
        lvWorker.RunWorkerAsync(Nothing)
    End Sub

    Private Sub DoStartIsUpdateAvailable(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        If AutoUpdate.IsUpdateAvailable() Then
            Me.DownloadUpdate()
            e.Result = True
        Else
            Me.DeleteLocalInstaller()
            e.Result = False
        End If
    End Sub

End Class

