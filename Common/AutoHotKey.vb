Option Explicit On
Option Strict On

Friend Class AutoHotKey

    Private Class HotKeyPool
        Inherits Collections.Hashtable ' Generic.SortedList(Of Char, System.ComponentModel.Component)
    End Class

    Private Class HotKeyControls
        Inherits Collections.ArrayList ' Generic.List(Of System.ComponentModel.Component)
    End Class

    Private Shared Function CreateHotKeyPool() As HotKeyPool
        Dim lvResult As New HotKeyPool
        Static svKeys As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        For Each lvChar As Char In svKeys
            lvResult.Add(lvChar, Nothing)
        Next
        Return lvResult
    End Function

    Private Shared Sub GetControls(ByVal piContainer As System.ComponentModel.Component, ByVal piControls As HotKeyControls, ByVal piControlWithHotkeys As HotKeyControls)
        If False Then
        ElseIf GetType(Windows.Forms.MenuStrip).IsAssignableFrom(piContainer.GetType) Then
            For Each lvItem As Windows.Forms.ToolStripItem In DirectCast(piContainer, Windows.Forms.MenuStrip).Items
                Dim lvText As String = lvItem.Text
                If False Then
                ElseIf AutoHotKey.HasHotKey(lvText) Then
                    piControlWithHotkeys.Add(lvItem)
                ElseIf IsHotKeyAssignable(lvText) Then
                    piControls.Add(lvItem)
                End If
                ' recursively hotkey sub-menu items using a new pool
                AssignHotKeys(lvItem)
            Next
        ElseIf GetType(Control).IsAssignableFrom(piContainer.GetType) Then
            For Each lvItem As Control In DirectCast(piContainer, Control).Controls
                If IsHotkeyControl(lvItem) Then
                    Dim lvText As String = lvItem.Text
                    If False Then
                    ElseIf AutoHotKey.HasHotKey(lvText) Then
                        piControlWithHotkeys.Add(lvItem)
                    ElseIf IsHotKeyAssignable(lvText) Then
                        piControls.Add(lvItem)
                    End If
                End If
                GetControls(lvItem, piControls, piControlWithHotkeys)
            Next
        ElseIf GetType(Windows.Forms.ToolStripMenuItem).IsAssignableFrom(piContainer.GetType) Then
            For Each lvItem As Windows.Forms.ToolStripItem In DirectCast(piContainer, Windows.Forms.ToolStripMenuItem).DropDownItems
                Dim lvText As String = lvItem.Text
                If False Then
                ElseIf AutoHotKey.HasHotKey(lvText) Then
                    piControlWithHotkeys.Add(lvItem)
                ElseIf IsHotKeyAssignable(lvText) Then
                    piControls.Add(lvItem)
                End If
                ' recursively hotkey sub-menu items using a new pool
                AssignHotKeys(lvItem)
            Next
        End If
    End Sub

    Private Shared Sub AssignHotKeys(ByVal piPool As HotKeyPool, ByVal piControls As HotKeyControls)
        For Each lvControl As System.ComponentModel.Component In piControls
            Dim lvText As String = GetText(lvControl)
            If AutoHotKey.HasHotKey(lvText) Then
                Debug.Fail(String.Format("Hotkey already assigned to {0}", lvText))
            ElseIf Not IsHotKeyAssignable(lvText) Then
                Debug.Fail(String.Format("Hotkey not assignable to {0}", lvText))
            Else
                Dim lvHotKeyAssigned As Boolean = TryAssignHotKey(piPool, lvControl)
                If Not lvHotKeyAssigned Then
                    For Each lvChar As Char In lvText.ToUpper
                        If piPool.ContainsKey(lvChar) Then
                            Dim lvPriorControl As System.ComponentModel.Component = DirectCast(piPool.Item(lvChar), System.ComponentModel.Component)
                            If TryAssignHotKey(piPool, lvPriorControl) Then
                                piPool.Item(lvChar) = lvControl
                                lvHotKeyAssigned = True
                                Exit For
                            End If
                        End If
                    Next
                End If
                If Not lvHotKeyAssigned Then
                    Debug.Fail(String.Format("Cannot assign a hotkey to {0}", lvText))
                End If
            End If
        Next
    End Sub

    Private Shared Sub RemoveHotKeys(ByVal piPool As HotKeyPool, ByVal piControls As HotKeyControls)
        For Each lvControl As System.ComponentModel.Component In piControls
            Dim lvText As String = GetText(lvControl)
            Dim lvHotKey As Char = GetHotKey(lvText)
            If piPool.ContainsKey(lvHotKey) Then
                piPool.Remove(lvHotKey)
            End If
        Next
    End Sub

    Private Shared Function TryAssignHotKey(ByVal piPool As HotKeyPool, ByVal piControl As System.ComponentModel.Component) As Boolean
        Dim lvText As String = GetText(piControl)
        For Each lvChar As Char In lvText.ToUpper
            If piPool.ContainsKey(lvChar) Then
                If IsNothing(piPool.Item(lvChar)) Then
                    piPool.Item(lvChar) = piControl
                    Return True
                End If
            End If
        Next
        Return False
    End Function

    Private Shared Function IsNullOrEmpty(ByVal piText As String) As Boolean
        Return piText Is Nothing OrElse 0 = piText.Length
    End Function

    Friend Shared Function HasHotKey(ByVal piText As String) As Boolean
        Return Not IsNullOrEmpty(piText) AndAlso (0 <= piText.IndexOf("&"c, 0, piText.Length - 1))
    End Function

    Private Shared Function IsHotKeyAssignable(ByVal piText As String) As Boolean
        Return Not IsNullOrEmpty(piText) AndAlso (0 > piText.IndexOf("&"c))
    End Function

    Friend Shared Function GetHotKey(ByVal piText As String) As Char
        Dim lvIndex As Int32 = piText.IndexOf("&"c)
        'If (0 > lvIndex) OrElse (lvIndex >= piText.Length) Then Return Nothing
        Return Char.ToUpper(piText.Chars(lvIndex + 1))
    End Function

    Private Shared Function GetText(ByVal piContainer As System.ComponentModel.Component) As String
        If GetType(Control).IsAssignableFrom(piContainer.GetType) Then
            Return DirectCast(piContainer, Control).Text
        Else
            Dim lvTextFieldInfo As Reflection.PropertyInfo = piContainer.GetType().GetProperty("Text")
            If Not IsNothing(lvTextFieldInfo) Then
                Return CStr(lvTextFieldInfo.GetValue(piContainer, Nothing))
            Else
                Debug.Fail(String.Format("{0}.Text does not exist", piContainer.GetType.Name))
                Return Nothing
            End If
        End If
    End Function

    Private Shared Sub SetText(ByVal piContainer As System.ComponentModel.Component, ByVal piText As String)
        If GetType(Control).IsAssignableFrom(piContainer.GetType) Then
            DirectCast(piContainer, Control).Text = piText
        Else
            Dim lvTextFieldInfo As Reflection.PropertyInfo = piContainer.GetType().GetProperty("Text")
            If Not IsNothing(lvTextFieldInfo) Then
                lvTextFieldInfo.SetValue(piContainer, piText, Nothing)
            Else
                Debug.Fail(String.Format("{0}.Text does not exist", piContainer.GetType.Name))
            End If
        End If
    End Sub

    Private Shared Function IsHotkeyControl(ByVal piControl As Object) As Boolean
        Dim piType As System.Type = piControl.GetType
        If False Then
        ElseIf GetType(Label).IsAssignableFrom(piType) Then
            Return DirectCast(piControl, Label).UseMnemonic
        ElseIf GetType(Button).IsAssignableFrom(piType) Then
            Return True
        ElseIf GetType(CheckBox).IsAssignableFrom(piType) Then
            Return True
        ElseIf GetType(MenuItem).IsAssignableFrom(piType) Then
            Return True
        End If
        Return False
    End Function

    Private Shared Function TryIndexOf(ByVal piText As String, ByVal piChar As Char, ByRef poIndex As Int32) As Boolean
        For lvIndex As Int32 = 0 To piText.Length - 1
            If piText.Chars(lvIndex) = piChar Then
                poIndex = lvIndex
                Return True
            End If
        Next
        Return False
    End Function

    Friend Shared Sub AssignHotKeys(ByVal piContainer As System.ComponentModel.Component)
        Dim lvPool As HotKeyPool = CreateHotKeyPool()
        Dim lvControlsWithoutHotkeys As New HotKeyControls
        Dim lvControlsWithHotkeys As New HotKeyControls
        GetControls(piContainer, lvControlsWithoutHotkeys, lvControlsWithHotkeys)
        RemoveHotKeys(lvPool, lvControlsWithHotkeys)
        AssignHotKeys(lvPool, lvControlsWithoutHotkeys)
        For Each lvChar As Char In lvPool.Keys
            Dim lvControl As System.ComponentModel.Component = DirectCast(lvPool.Item(lvChar), System.ComponentModel.Component)
            If Not IsNothing(lvControl) Then
                Dim lvText As String = GetText(lvControl)
                Dim lvCharIndex As Int32 = -1
                If TryIndexOf(lvText.ToUpper, Char.ToUpper(lvChar), lvCharIndex) Then
                    lvText = lvText.Substring(0, lvCharIndex) & "&" & lvText.Substring(lvCharIndex)
                    SetText(lvControl, lvText)
                Else
                    Debug.Fail(String.Format("Hotkey {0} not found in {1}", lvChar, lvText))
                End If
            Else
                ' unassigned hotkey
            End If
        Next
    End Sub

End Class