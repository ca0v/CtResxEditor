Option Explicit On
Option Strict On
Option Infer On

Friend Class ResUtility

    Friend Shared Sub ExtractCultureInfo(ByVal piFileName As String, ByRef poFileName As String, ByRef poCountry As String, ByRef poCulture As String, ByRef poExt As String)
        Dim lvRegEx = New System.Text.RegularExpressions.Regex("(?<file>.*)\.(?<locale>\w\w)-(?<country>\w\w)\.(?<ext>.*)")
        If lvRegEx.IsMatch(piFileName) Then
            Dim lvMatch = lvRegEx.Match(piFileName)
            poFileName = lvMatch.Groups("file").Value
            poCulture = lvMatch.Groups("locale").Value
            poCountry = lvMatch.Groups("country").Value
            poExt = lvMatch.Groups("ext").Value
        Else
            lvRegEx = New System.Text.RegularExpressions.Regex("(?<file>.*)\.(?<locale>\w\w)\.(?<ext>.*)")
            If lvRegEx.IsMatch(piFileName) Then
                Dim lvMatch = lvRegEx.Match(piFileName)
                poFileName = lvMatch.Groups("file").Value
                poCulture = lvMatch.Groups("locale").Value
                poCountry = String.Empty
                poExt = lvMatch.Groups("ext").Value
            Else
                lvRegEx = New System.Text.RegularExpressions.Regex("(?<file>.*)\.(?<ext>.*)")
                If lvRegEx.IsMatch(piFileName) Then
                    Dim lvMatch = lvRegEx.Match(piFileName)
                    poFileName = lvMatch.Groups("file").Value
                    poCulture = String.Empty
                    poCountry = String.Empty
                    poExt = lvMatch.Groups("ext").Value
                Else
                    Debug.Fail(String.Format("Invalid file name: {0}", piFileName))
                End If
            End If
        End If

    End Sub

    Friend Shared Sub MergeDataSet(ByVal piPath As String, ByVal piName As String, ByVal piExt As String, ByVal piDataSet As NewDataSet)
        Dim lvRootDir As String = piPath
        Debug.Assert(IO.Directory.Exists(lvRootDir), String.Format("{0} does not exist", lvRootDir))
        Dim lvListOfFiles = New List(Of String)
        Dim lvSearchPattern As String = String.Format("{0}.{1}", piName, piExt)
        Dim lvFiles = IO.Directory.GetFiles(lvRootDir, lvSearchPattern, IO.SearchOption.TopDirectoryOnly)
        lvListOfFiles.AddRange(lvFiles)
        lvSearchPattern = String.Format("{0}.??.{1}", piName, piExt)
        lvFiles = IO.Directory.GetFiles(lvRootDir, lvSearchPattern, IO.SearchOption.TopDirectoryOnly)
        lvListOfFiles.AddRange(lvFiles)
        lvSearchPattern = String.Format("{0}.??-??.{1}", piName, piExt)
        lvFiles = IO.Directory.GetFiles(lvRootDir, lvSearchPattern, IO.SearchOption.TopDirectoryOnly)
        lvListOfFiles.AddRange(lvFiles)
        If "RESX".Equals(piExt, StringComparison.InvariantCultureIgnoreCase) Then
            For Each lvResxFile As String In lvListOfFiles
                MergeResXDataSet(lvResxFile, piDataSet, OverwriteOptions.AlwaysOverwrite)
            Next
        ElseIf "RESOURCES".Equals(piExt, StringComparison.InvariantCultureIgnoreCase) Then
            For Each lvResxFile As String In lvListOfFiles
                MergeResourceDataSet(lvResxFile, piDataSet)
            Next
        End If
    End Sub

    <Flags()> Friend Enum OverwriteOptions
        None = 0
        AllowOverwrite = 1
        PromptForOverwrite = 2
        AlwaysOverwrite = AllowOverwrite Or PromptForOverwrite
    End Enum

    Friend Shared Sub MergeResXDataSet(ByVal piFileName As String, ByVal piDataSet As NewDataSet, ByVal piAllowOverwrite As OverwriteOptions)
        Debug.Assert(".RESX".Equals(IO.Path.GetExtension(piFileName), StringComparison.InvariantCultureIgnoreCase), "Expecting a .RESX file")
        Dim lvFileName As String = String.Empty
        Dim lvCountry As String = String.Empty
        Dim lvCulture As String = String.Empty
        Dim lvExt As String = String.Empty
        ExtractCultureInfo(IO.Path.GetFileName(piFileName), lvFileName, lvCountry, lvCulture, lvExt)
        Dim lvCultureCode As String = lvCulture
        If Not String.IsNullOrEmpty(lvCountry) Then
            lvCultureCode &= "-" & lvCountry
        End If
        Using lvResxReader As New System.Resources.ResXResourceReader(piFileName)
            Dim lvColumn As DataColumn = ForceCultureColumn(lvCultureCode, piDataSet.ResxData)
            Dim lvEnumerator = lvResxReader.GetEnumerator
            Select Case piAllowOverwrite
                Case OverwriteOptions.None ' never overwrite
                    Do While lvEnumerator.MoveNext
                        Dim lvRow As NewDataSet.ResxDataRow = ForceResxRow(lvEnumerator.Entry.Key.ToString, piDataSet.ResxData)
                        If lvRow(lvColumn) Is DBNull.Value Then
                            lvRow(lvColumn) = lvEnumerator.Entry.Value.ToString
                            lvRow.Version = Now
                        Else
                            ' do nothing
                        End If
                    Loop
                Case OverwriteOptions.AlwaysOverwrite
                    Do While lvEnumerator.MoveNext
                        Dim lvRow As NewDataSet.ResxDataRow = ForceResxRow(lvEnumerator.Entry.Key.ToString, piDataSet.ResxData)
                        If lvRow(lvColumn) Is DBNull.Value OrElse Not String.Equals(lvRow(lvColumn), lvEnumerator.Entry.Value) Then
                            lvRow(lvColumn) = lvEnumerator.Entry.Value.ToString
                            lvRow.Version = Now
                        Else
                            ' do nothing
                        End If
                    Loop
                Case OverwriteOptions.PromptForOverwrite
                    Dim lvSkip = New Generic.SortedList(Of String, DialogResult)
                    Do While lvEnumerator.MoveNext
                        Dim lvRow As NewDataSet.ResxDataRow = ForceResxRow(lvEnumerator.Entry.Key.ToString, piDataSet.ResxData)
                        If lvRow(lvColumn) Is DBNull.Value Then
                            lvRow(lvColumn) = lvEnumerator.Entry.Value.ToString
                            lvRow.Version = Now
                        ElseIf Not String.Equals(lvRow(lvColumn), lvEnumerator.Entry.Value) Then
                            ' prompt to overwrite
                            Dim lvResult As DialogResult
                            If Not lvSkip.TryGetValue(lvRow.id, lvResult) Then
                                Dim lvText = String.Format("{2}: Replace ""{0}"" with ""{1}""?", lvRow(lvColumn), lvEnumerator.Entry.Value, lvRow.id)
                                Dim lvCaption = "Okay to overwrite?"
                                lvResult = System.Windows.Forms.MessageBox.Show(lvText, lvCaption, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                                lvSkip.Add(lvRow.id, lvResult)
                            End If
                            Select Case lvResult
                                Case DialogResult.Yes
                                    lvRow(lvColumn) = lvEnumerator.Entry.Value.ToString
                                    lvRow.Version = Now
                                Case DialogResult.No
                                Case DialogResult.Cancel
                                    Exit Do
                            End Select
                        Else
                            ' do nothing
                        End If
                    Loop
            End Select
            lvResxReader.Close()
        End Using
    End Sub

    Friend Shared Sub MergeResourceDataSet(ByVal piFileName As String, ByVal piDataSet As NewDataSet)
        Debug.Assert(".RESOURCES".Equals(IO.Path.GetExtension(piFileName), StringComparison.InvariantCultureIgnoreCase), "Expecting a .RESOURCES file")
        Dim lvFileName As String = String.Empty
        Dim lvCountry As String = String.Empty
        Dim lvCulture As String = String.Empty
        Dim lvExt As String = String.Empty
        ExtractCultureInfo(IO.Path.GetFileName(piFileName), lvFileName, lvCountry, lvCulture, lvExt)
        Dim lvCultureCode As String = lvCulture
        If Not String.IsNullOrEmpty(lvCountry) Then
            lvCultureCode &= "-" & lvCountry
        End If
        Using lvResxReader As New System.Resources.ResourceReader(piFileName)
            Dim lvColumn As DataColumn = ForceCultureColumn(lvCultureCode, piDataSet.ResxData)
            Dim lvEnumerator = lvResxReader.GetEnumerator
            Do While lvEnumerator.MoveNext
                Dim lvNewValue = lvEnumerator.Entry.Value
                Dim lvRow As NewDataSet.ResxDataRow = ForceResxRow(lvEnumerator.Entry.Key.ToString, piDataSet.ResxData)
                If IsNothing(lvNewValue) Then
                    If lvRow(lvColumn) Is DBNull.Value Then
                        ' nothing to do
                    Else
                        lvRow(lvColumn) = DBNull.Value
                    End If
                ElseIf Not String.Equals(lvRow(lvColumn), lvNewValue) Then
                    lvRow(lvColumn) = lvNewValue.ToString
                    lvRow.Version = Now
                Else
                    ' do nothing
                End If
            Loop
            lvResxReader.Close()
        End Using
    End Sub

    Friend Shared Sub MergeNameValuePairDataSet(ByVal piFileName As String, ByVal piDataSet As NewDataSet, ByVal piCultureCode As String)
        Debug.Assert(".PROPERTIES".Equals(IO.Path.GetExtension(piFileName), StringComparison.InvariantCultureIgnoreCase), "Expecting a .PROPERTIES file")
        Dim lvCollection As New Generic.Dictionary(Of String, String)
        Using lvResxReader = IO.File.OpenText(piFileName)
            Dim lvColumn As DataColumn = ForceCultureColumn(piCultureCode, piDataSet.ResxData)
            Do While Not lvResxReader.EndOfStream
                Dim lvLine As String = lvResxReader.ReadLine
                If Not String.IsNullOrEmpty(lvLine) AndAlso Not lvLine.StartsWith("#") Then
                    Dim lvIndex As Int32 = lvLine.IndexOf("="c)
                    If (0 < lvIndex) Then
                        Dim lvName As String = lvLine.Substring(0, lvIndex)
                        Dim lvValue As String = lvLine.Substring(lvIndex + 1)
                        Dim lvRow As NewDataSet.ResxDataRow = ForceResxRow(lvName, piDataSet.ResxData)
                        If lvRow(lvColumn) Is DBNull.Value OrElse Not String.Equals(lvRow(lvColumn), lvValue) Then
                            lvRow(lvColumn) = lvValue
                            lvRow.Version = Now
                        Else
                            ' do nothing
                        End If
                    End If
                End If
            Loop
            lvResxReader.Close()
        End Using
    End Sub

    Friend Shared Function ForceCultureColumn(ByVal piCultureCode As String, ByVal piResxData As NewDataSet.ResxDataDataTable) As DataColumn
        If String.IsNullOrEmpty(piCultureCode) Then
            Return piResxData.DefaultColumn
        End If
        Dim lvResult As DataColumn = piResxData.Columns(piCultureCode)
        If IsNothing(lvResult) Then
            lvResult = piResxData.Columns.Add(piCultureCode)
        End If
        Return lvResult
    End Function

    Private Shared Function ForceResxRow(ByVal piResourceId As String, ByVal piResxData As NewDataSet.ResxDataDataTable) As NewDataSet.ResxDataRow
        Dim lvResult As NewDataSet.ResxDataRow = piResxData.FindByid(piResourceId)
        If IsNothing(lvResult) Then
            lvResult = piResxData.NewResxDataRow
            lvResult.id = piResourceId
            piResxData.AddResxDataRow(lvResult)
        Else
            ' nothing to do
        End If
        Return lvResult
    End Function

End Class
