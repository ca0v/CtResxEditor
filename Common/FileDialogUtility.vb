Option Explicit On
Option Strict On

Friend Class FileDialogUtility

    Private Structure FileDescriptionFileExtPair
        Public ReadOnly FileDescription As String
        Public ReadOnly FileExt As String
        Public Sub New(ByVal piDescription As String, ByVal piExt As String)
            Me.FileDescription = piDescription
            Me.FileExt = piExt
        End Sub
    End Structure

    Friend Class FileEventArgs
        Inherits System.ComponentModel.CancelEventArgs

        Private m_FileName As String
        Property FileName() As String
            Get
                Return Me.m_FileName
            End Get
            Set(ByVal value As String)
                Me.m_FileName = value
            End Set
        End Property

        Public Sub New(ByVal piFileName As String)
            Me.m_FileName = piFileName
        End Sub
    End Class

    Private m_FileName As String
    Private m_DefaultFileExt As String
    Private m_FileTypes As New List(Of FileDescriptionFileExtPair)
    Friend Sub New(ByVal piInitialFileName As String)
        Me.m_FileName = piInitialFileName
    End Sub

    Friend Sub AddFileExtension(ByVal piFileDescription As String, ByVal piFileExt As String)
        If String.IsNullOrEmpty(Me.m_DefaultFileExt) Then
            Me.m_DefaultFileExt = piFileExt
        End If
        Me.m_FileTypes.Add(New FileDescriptionFileExtPair(piFileDescription, piFileExt))
    End Sub

    Private Function BuildFilter() As String
        ' String.Format("Service Collection (*{0})|*{0}|WSDL (*.wsdl)|*.wsdl", cEamslFileExt)
        Dim lvResult As New System.Text.StringBuilder
        For Each lvItem As FileDescriptionFileExtPair In Me.m_FileTypes
            If (0 < lvResult.Length) Then
                lvResult.Append("|"c)
            End If
            lvResult.Append(String.Format("{0} (*{1})|*{1}", lvItem.FileDescription, lvItem.FileExt))
        Next
        Return lvResult.ToString
    End Function

    Private Function GetFilterIndex(ByVal piFileExt As String) As Int32
        For lvResult As Int32 = 0 To Me.m_FileTypes.Count - 1
            If Me.m_FileTypes(lvResult).FileExt = piFileExt Then
                Return lvResult
            End If
        Next
        Return -1
    End Function

    Friend Event FileOpening(ByVal sender As Object, ByVal e As FileEventArgs)
    Friend Event FileSaving(ByVal sender As Object, ByVal e As FileEventArgs)

    Private Function Save(ByVal piFileName As String) As Boolean
        Dim lvEventArgs As New FileEventArgs(piFileName)
        RaiseEvent FileSaving(Me, lvEventArgs)
        If Not lvEventArgs.Cancel Then
            Me.m_FileName = lvEventArgs.FileName
        End If
        Return lvEventArgs.Cancel
    End Function

    Friend Sub FileOpen(ByVal piFileName As String)
        Dim lvEventArgs As New FileDialogUtility.FileEventArgs(piFileName)
        RaiseEvent FileOpening(Me, lvEventArgs)
        If Not lvEventArgs.Cancel Then
            Me.m_FileName = lvEventArgs.FileName
        End If
    End Sub

    Friend Function FileSave() As Boolean
        If String.IsNullOrEmpty(Me.m_FileName) Then
            Return Me.FileSaveAs()
        Else
            Return Me.Save(Me.m_FileName)
        End If
    End Function

    Friend Function FileSaveAs() As Boolean
        Dim lvSaveFileDialog As New System.Windows.Forms.SaveFileDialog
        lvSaveFileDialog.FileName = Me.m_FileName
        lvSaveFileDialog.Filter = Me.BuildFilter()
        lvSaveFileDialog.FilterIndex = Me.GetFilterIndex(IO.Path.GetExtension(Me.m_FileName))
        lvSaveFileDialog.DefaultExt = m_DefaultFileExt
        lvSaveFileDialog.AddExtension = True
        lvSaveFileDialog.CheckFileExists = False
        lvSaveFileDialog.CreatePrompt = True
        lvSaveFileDialog.OverwritePrompt = True
        lvSaveFileDialog.SupportMultiDottedExtensions = False
        lvSaveFileDialog.ValidateNames = True
        Select Case lvSaveFileDialog.ShowDialog
            Case Windows.Forms.DialogResult.OK
                Return Me.Save(lvSaveFileDialog.FileName)
        End Select
        Return False
    End Function

    Friend Sub FileOpen()
        Dim lvOpenFileDialog As New Windows.Forms.OpenFileDialog
        lvOpenFileDialog.Filter = Me.BuildFilter
        lvOpenFileDialog.FilterIndex = Me.GetFilterIndex(IO.Path.GetExtension(Me.m_FileName))
        lvOpenFileDialog.DefaultExt = Me.m_DefaultFileExt
        lvOpenFileDialog.CheckFileExists = True
        lvOpenFileDialog.CheckPathExists = True
        lvOpenFileDialog.FileName = Me.m_FileName
        Select Case lvOpenFileDialog.ShowDialog
            Case Windows.Forms.DialogResult.OK
                Me.FileOpen(lvOpenFileDialog.FileName)
        End Select
    End Sub

End Class
