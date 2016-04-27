Option Explicit On
Option Strict On

Friend Class CommandLineArgAttribute
    Inherits Attribute
    Public ReadOnly Switch As String
    Public ReadOnly Description As String
    Public ReadOnly Required As Boolean
    Public Sub New(ByVal piSwitch As String, ByVal piDescription As String, ByVal piRequired As Boolean)
        Me.Switch = piSwitch
        Me.Description = piDescription
        Me.Required = piRequired
    End Sub
End Class

Public MustInherit Class BaseCommandArgs

    ' not public to defeat serialization
    <CommandLineArg("-config=", "Get settings from this configuration file", False)> Friend ConfigFile As String
    <CommandLineArg("-help=", "Show help", False)> Friend ShowHelp As Boolean

    Friend Function IsValid() As Boolean
        Try
            Me.Validate()
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return False
        End Try
    End Function

    Protected Overridable Sub Validate()
        For Each lvField As Generic.KeyValuePair(Of CommandLineArgAttribute, Reflection.FieldInfo) In Me.GetCommandLineFields()
            If lvField.Key.Required Then
                Dim lvValue As Object = lvField.Value.GetValue(Me)
                If IsNothing(lvValue) Then
                    Throw New System.ApplicationException(String.Format("value for {0} is required", lvField.Key.Switch))
                End If
                If lvValue.GetType Is GetType(String) Then
                    If String.IsNullOrEmpty(CStr(lvValue)) Then
                        Throw New System.ApplicationException(String.Format("non-trivial value for {0} is required", lvField.Key.Switch))
                    End If
                End If
            End If
        Next
    End Sub

    Private Shared Sub Write(ByVal piFileName As String, ByVal piValue As String)
        Using lvReader As New System.IO.StringReader(piValue)
            Using lvFileStream As New System.IO.FileStream(piFileName, IO.FileMode.Create)
                Dim lvWriter As New System.IO.StreamWriter(lvFileStream)
                lvWriter.Write(piValue)
                lvWriter.Close()
            End Using
            lvReader.Close()
        End Using
    End Sub

    Private Shared Function Serialize(ByVal piObject As Object) As String
        Dim lvSerializer As New System.Xml.Serialization.XmlSerializer(piObject.GetType)
        Using lvWriter As New System.IO.StringWriter
            lvSerializer.Serialize(lvWriter, piObject)
            lvWriter.Close()
            Return lvWriter.ToString
        End Using
    End Function

    Private Shared Function Deserialize(ByVal piFileName As String, ByVal piType As System.Type) As Object
        Dim lvSerializer As New System.Xml.Serialization.XmlSerializer(piType)
        Using lvFileStream As New IO.FileStream(piFileName, IO.FileMode.Open, IO.FileAccess.Read)
            Using lvReader As New System.IO.StreamReader(lvFileStream, System.Text.Encoding.UTF8, False)
                Return lvSerializer.Deserialize(lvReader)
            End Using
        End Using
    End Function

    Public Sub SaveToFile(ByVal piFileName As String)
        Write(piFileName, Serialize(Me))
    End Sub

    Public Sub LoadFromFile(ByVal piFileName As String)
        Dim lvBaseCommandArgs As BaseCommandArgs = DirectCast(Deserialize(piFileName, Me.GetType), BaseCommandArgs)
        Me.Initialize(lvBaseCommandArgs)
        Me.ConfigFile = piFileName ' overwrite the configfile value in the config file
    End Sub

    Private Shared Sub ExtractValue(ByVal piArg As String, ByRef poValue As Int16)
        Dim lvIndex As Int32 = piArg.IndexOf("=")
        If (0 > lvIndex) Then
            poValue = -1
        Else
            poValue = Int16.Parse(piArg.Substring(lvIndex + 1))
        End If
    End Sub

    Private Shared Sub ExtractValue(ByVal piArg As String, ByRef poValue As String)
        Dim lvIndex As Int32 = piArg.IndexOf("=")
        If (0 > lvIndex) Then
            poValue = String.Empty
        Else
            poValue = piArg.Substring(lvIndex + 1)
        End If
    End Sub

    Private Shared Sub ExtractValue(ByVal piArg As String, ByRef poValue As Boolean)
        Dim lvIndex As Int32 = piArg.IndexOf("=")
        If (0 > lvIndex) Then
            poValue = True
        Else
            poValue = Boolean.Parse(piArg.Substring(lvIndex + 1))
        End If
    End Sub

    Private Shared Sub ExtractValue(ByVal piArg As String, ByRef poValue As String())
        Dim lvValue As String = Nothing
        ExtractValue(piArg, lvValue)
        poValue = lvValue.Split(","c)
        For i As Int32 = 0 To poValue.Length - 1
            poValue(i) = poValue(i).Trim
        Next
    End Sub

    Private Function GetCommandLineFields() As Dictionary(Of CommandLineArgAttribute, Reflection.FieldInfo)
        Dim lvType As System.Type = Me.GetType
        Dim lvResult As New Dictionary(Of CommandLineArgAttribute, Reflection.FieldInfo)
        Dim lvFields() As Reflection.FieldInfo = lvType.GetFields(Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Public Or Reflection.BindingFlags.Instance Or Reflection.BindingFlags.Static)
        For Each lvField As Reflection.FieldInfo In lvFields
            Dim lvAttributes() As Object = lvField.GetCustomAttributes(GetType(CommandLineArgAttribute), True)
            If (0 < lvAttributes.Length) Then
                Dim lvCommandLineArg As CommandLineArgAttribute = DirectCast(lvAttributes(0), CommandLineArgAttribute)
                lvResult.Add(lvCommandLineArg, lvField)
            End If
        Next
        Return lvResult
    End Function

    Private Shared Function QuoteIt(ByVal piValue As String, ByVal piQuote As String) As String
        Return String.Concat(piQuote, piValue.Replace(piQuote, piQuote & piQuote), piQuote)
    End Function

    Friend Function ShowUsage() As String
        Dim lvType As System.Type = Me.GetType
        Dim lvExe As String = IO.Path.GetFileName(lvType.Assembly.Location)
        Dim lvFields As Dictionary(Of CommandLineArgAttribute, Reflection.FieldInfo) = GetCommandLineFields()
        Dim lvResult As String = "Descriptions: " & vbCrLf
        For Each lvField As Generic.KeyValuePair(Of CommandLineArgAttribute, Reflection.FieldInfo) In lvFields
            lvResult = String.Concat(lvResult, lvField.Key.Switch, IIf(lvField.Key.Switch.Length >= 8, vbTab, vbTab & vbTab), lvField.Key.Description, vbCrLf)
        Next
        Return lvResult
    End Function

    Friend Function ShowSettings() As String
        Dim lvType As System.Type = Me.GetType
        Dim lvResult As New System.Text.StringBuilder
        lvResult.Append("Default settings: ")
        lvResult.Append(vbCrLf)
        Dim lvFields As Dictionary(Of CommandLineArgAttribute, Reflection.FieldInfo) = GetCommandLineFields()
        For Each lvField As Generic.KeyValuePair(Of CommandLineArgAttribute, Reflection.FieldInfo) In lvFields
            Dim lvValue As Object = lvField.Value.GetValue(Me)
            If IsNothing(lvValue) Then
                Continue For
            End If
            If False Then
            ElseIf lvValue.GetType Is GetType(String()) Then
                lvValue = String.Join(","c, DirectCast(lvValue, String()))
            ElseIf lvValue.GetType Is GetType(Boolean) Then
                If Not CBool(lvValue) Then
                    Continue For
                End If
            End If
            lvResult.Append(" " & lvField.Key.Switch)
            lvResult.Append(QuoteIt(lvValue.ToString, """"))
            lvResult.Append(vbCrLf)
        Next
        Return lvResult.ToString
    End Function

    Protected Overridable Sub AssignValue(ByVal piFieldInfo As System.Reflection.FieldInfo, ByVal piValue As Object)
        piFieldInfo.SetValue(Me, piValue)
    End Sub

    Public Sub Initialize(ByVal piArgs() As String)
        ' apply the settings in the config file first so they can be overwritten by the command line args
        For Each lvItem As String In piArgs
            If lvItem.StartsWith("-config=") Then
                Dim lvValueStr As String = lvItem.Substring("-config=".Length)
                If IO.File.Exists(lvValueStr) Then
                    Me.LoadFromFile(lvValueStr)
                End If
                Exit For
            End If
        Next
        For Each lvItem As String In piArgs
            Dim lvAssigned As Boolean = False
            For Each lvField As Generic.KeyValuePair(Of CommandLineArgAttribute, Reflection.FieldInfo) In Me.GetCommandLineFields()
                If lvItem.StartsWith(lvField.Key.Switch) Then
                    Dim lvValueStr As String = lvItem.Substring(lvField.Key.Switch.Length)
                    ' convert to the desired type
                    Dim lvValueObj As Object = System.Convert.ChangeType(lvValueStr, lvField.Value.FieldType)
                    AssignValue(lvField.Value, lvValueObj)
                    lvAssigned = True
                    Exit For
                End If
            Next
            Debug.Assert(lvAssigned, String.Format("{0} is an invalid switch (ignored)", lvItem))
            If Not lvAssigned Then
                Console.WriteLine("{0} ignored", lvItem)
            End If
        Next
    End Sub

    Private Sub Initialize(ByVal piCommandArgs As BaseCommandArgs)
        Debug.Assert(Me.GetType.IsAssignableFrom(piCommandArgs.GetType), "Assumes piCommandArgs is a subclass")
        For Each lvField As Generic.KeyValuePair(Of CommandLineArgAttribute, Reflection.FieldInfo) In Me.GetCommandLineFields()
            AssignValue(lvField.Value, lvField.Value.GetValue(piCommandArgs))
        Next
    End Sub

End Class