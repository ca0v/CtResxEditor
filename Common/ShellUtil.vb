Option Strict On
Option Explicit On

Friend Class ShellUtil

    Private Sub New()
        ' this is a module
    End Sub

    Friend Shared Sub AddShortcutToStartMenu(ByVal piAssembly As Reflection.Assembly)
        Dim lvStartPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Programs)
        lvStartPath = IO.Path.Combine(lvStartPath, Windows.Forms.Application.CompanyName)
        If Not IO.Directory.Exists(lvStartPath) Then
            IO.Directory.CreateDirectory(lvStartPath)
        End If
        Dim lvProductName As String = IO.Path.Combine(lvStartPath, Windows.Forms.Application.ProductName)
        Dim lvProductDescription As String = String.Empty
        If Not TryGetProductDescription(piAssembly, lvProductDescription) Then
            lvProductDescription = Windows.Forms.Application.ProductName
        End If
        ShellUtilHelper.CreateShortcut(lvStartPath, lvProductName, lvProductDescription, piAssembly.Location)
    End Sub

    Friend Shared Function TryGetProductDescription(ByVal piAssembly As System.Reflection.Assembly, ByRef poDescription As String) As Boolean
        For Each lvAttribute As System.Reflection.AssemblyDescriptionAttribute In piAssembly.GetCustomAttributes(GetType(System.Reflection.AssemblyDescriptionAttribute), False)
            poDescription = lvAttribute.Description
            Return True
        Next
        Return False
    End Function

End Class