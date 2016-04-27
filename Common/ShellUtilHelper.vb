Option Strict Off
Option Explicit On

Friend Module ShellUtilHelper
    Friend Sub CreateShortcut(ByVal piStartPath As String, ByVal piProductName As String, ByVal piProductDescription As String, ByVal piLocation As String)
        Dim lvShell As Object = CreateObject("WScript.Shell")
        Dim lvShortcut As Object = lvShell.CreateShortcut(IO.Path.Combine(piStartPath, piProductName & ".lnk"))
        lvShortcut.TargetPath = piLocation
        lvShortcut.Description = piProductDescription
        lvShortcut.WorkingDirectory = IO.Path.GetDirectoryName(piLocation)
        lvShortcut.IconLocation = String.Format("{0}, 0", piLocation)
        lvShortcut.Save()
    End Sub
End Module

