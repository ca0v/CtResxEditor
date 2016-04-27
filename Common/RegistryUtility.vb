Option Explicit On
Option Strict On
Imports Microsoft.Win32

Friend Class RegistryUtility

    Friend Shared Sub AssociateUserCommandToFolder(ByVal piAssembly As Reflection.Assembly)
        Dim lvApplicationName As String = GetApplicationName(piAssembly)
        Dim lvRootKey As RegistryKey = My.Computer.Registry.CurrentUser
        Dim lvSubKey As RegistryKey = ForceRegistryKey(lvRootKey, String.Format("Software\Classes\Directory\shell\{0}\command", lvApplicationName))
        Try
            ForceKeyValue(lvSubKey, String.Empty, String.Format("{0} ""%1""", piAssembly.Location))
        Finally
            lvSubKey.Close()
        End Try
    End Sub

    Friend Shared Sub AssociateUserCommandToDrive(ByVal piAssembly As Reflection.Assembly)
        Dim lvApplicationName As String = GetApplicationName(piAssembly)
        Dim lvRootKey As RegistryKey = My.Computer.Registry.CurrentUser
        Dim lvSubKey As RegistryKey = ForceRegistryKey(lvRootKey, String.Format("Software\Classes\Drive\shell\{0}\command", lvApplicationName))
        Try
            ForceKeyValue(lvSubKey, String.Empty, String.Format("{0} ""%1""", piAssembly.Location))
        Finally
            lvSubKey.Close()
        End Try
    End Sub

    Friend Shared Sub AssociateUserCommandToFileExtension(ByVal piAssembly As Reflection.Assembly, ByVal piFileExt As String)
        Dim lvRootKey As RegistryKey = ForceRegistryKey(My.Computer.Registry.CurrentUser, "Software\Classes")
        Try
            AssociateCommandToFileExtension(lvRootKey, piAssembly, piFileExt)
        Finally
            lvRootKey.Close()
        End Try
    End Sub

    Friend Shared Sub AssociateCommandToFileExtension(ByVal piAssembly As Reflection.Assembly, ByVal piFileExt As String)
        AssociateCommandToFileExtension(My.Computer.Registry.ClassesRoot, piAssembly, piFileExt)
    End Sub

    Private Shared Sub AssociateCommandToFileExtension(ByVal piRootKey As RegistryKey, ByVal piAssembly As Reflection.Assembly, ByVal piFileExt As String)
        Dim lvCompanyName As String = GetCompanyName(piAssembly)
        Dim lvApplicationName As String = GetApplicationName(piAssembly)
        Dim lvSubKey As RegistryKey
        lvSubKey = ForceRegistryKey(piRootKey, String.Format("{0}", piFileExt))
        Try
            ForceKeyValue(lvSubKey, String.Empty, String.Format("{0}\{1}", lvCompanyName, lvApplicationName))
        Finally
            lvSubKey.Close()
        End Try
        lvSubKey = ForceRegistryKey(piRootKey, String.Format("{0}\{1}\shell\open\command", lvCompanyName, lvApplicationName))
        Try
            ForceKeyValue(lvSubKey, String.Empty, String.Format("{0} %1", piAssembly.Location))
        Finally
            lvSubKey.Close()
        End Try
    End Sub

    Private Shared Sub ForceKeyValue(ByVal piKey As RegistryKey, ByVal piName As String, ByVal piValue As String)
        If Not piValue.Equals(piKey.GetValue(piName)) Then
            piKey.SetValue(piName, piValue)
        End If
    End Sub

    Private Shared Function ForceRegistryKey(ByVal piRoot As RegistryKey, ByVal piName As String) As RegistryKey
        Dim lvResult As RegistryKey = piRoot.OpenSubKey(piName, True)
        If IsNothing(lvResult) Then
            lvResult = piRoot.CreateSubKey(piName)
        End If
        Return lvResult
    End Function

    Private Shared Function GetApplicationName(ByVal piAssembly As Reflection.Assembly) As String
        Dim lvAssemblies() As Object = piAssembly.GetCustomAttributes(GetType(System.Reflection.AssemblyProductAttribute), False)
        If (0 < lvAssemblies.Length) Then
            Return DirectCast(lvAssemblies(0), Reflection.AssemblyProductAttribute).Product
        End If
        lvAssemblies = piAssembly.GetCustomAttributes(GetType(System.Reflection.AssemblyTitleAttribute), False)
        If (0 < lvAssemblies.Length) Then
            Return DirectCast(lvAssemblies(0), Reflection.AssemblyTitleAttribute).Title
        End If
        Return IO.Path.GetFileNameWithoutExtension(piAssembly.Location)
    End Function

    Private Shared Function GetCompanyName(ByVal piAssembly As Reflection.Assembly) As String
        Dim lvAssemblies() As Object = piAssembly.GetCustomAttributes(GetType(System.Reflection.AssemblyCompanyAttribute), False)
        If (0 < lvAssemblies.Length) Then
            Return DirectCast(lvAssemblies(0), Reflection.AssemblyCompanyAttribute).Company
        End If
        Return IO.Path.GetFileNameWithoutExtension(piAssembly.Location)
    End Function

End Class