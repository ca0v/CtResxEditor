Option Explicit On
Option Strict On
Option Infer On

Public Interface IFastFilter
    Property HasNullRow() As Boolean
End Interface

Friend Class FilterUtility

    Friend Shared Function LevenshteinDistance(ByVal s As String, ByVal t As String) As Integer
        Dim lvBuffer() As Integer = Nothing
        Return LevenshteinDistance(s, t, lvBuffer)
    End Function

    Friend Shared Function LevenshteinDistance(ByVal s As String, ByVal t As String, ByRef poBuffer() As Integer) As Integer
        Dim sLength As Integer = s.Length ' length of s
        Dim tLength As Integer = t.Length ' length of t

        ' Step 1
        If tLength = 0 Then
            Return sLength
        ElseIf sLength = 0 Then
            Return tLength
        End If

        Dim lvMatrixSize As Integer = (1 + sLength) * (1 + tLength)
        If IsNothing(poBuffer) OrElse (poBuffer.Length < lvMatrixSize) Then
            poBuffer = New Integer(0 To lvMatrixSize - 1) {}
        End If

        ' fill first row
        For lvIndex As Integer = 0 To sLength
            poBuffer(lvIndex) = lvIndex
        Next
        'fill first column
        For lvIndex As Integer = 1 To tLength
            poBuffer(lvIndex * (sLength + 1)) = lvIndex
        Next

        Dim lvCost As Integer ' cost
        Dim lvDistance As Integer = 0
        For lvRowIndex As Integer = 0 To sLength - 1
            Dim s_i As Char = s(lvRowIndex)
            For lvColIndex As Integer = 0 To tLength - 1
                Dim tchar As Char = t(lvColIndex)
                Select Case tchar
                    Case s_i
                        lvCost = 0
                    Case ChrW((Asc(s_i) Xor 32))
                        lvCost = 1
                    Case "."c, "-"c, "_"c
                        lvCost = 1
                    Case Else
                        lvCost = 2
                End Select
                ' Step 6
                Dim lvTopLeftIndex As Integer = lvColIndex * (sLength + 1) + lvRowIndex
                Dim lvTopLeft As Integer = poBuffer(lvTopLeftIndex)
                Dim lvTop As Integer = poBuffer(lvTopLeftIndex + 1)
                Dim lvLeft As Integer = poBuffer(lvTopLeftIndex + (sLength + 1))
                lvDistance = Math.Min(lvTopLeft + lvCost, Math.Min(lvLeft, lvTop) + 2)
                poBuffer(lvTopLeftIndex + sLength + 2) = lvDistance
            Next
        Next
        Return lvDistance
    End Function

    Friend Shared Function EscapeForLike(ByVal piFilter As String) As String
        Dim lvEscapedFilter As New System.Text.StringBuilder(piFilter.Length)
        Static svSpecialCharacters() As Char = {"["c, "]"c, "*"c, "%"c}
        For Each lvChar As Char In piFilter.ToCharArray
            If "'"c = lvChar Then
                lvEscapedFilter.Append("''")
            ElseIf (0 <= Array.IndexOf(svSpecialCharacters, lvChar)) Then
                lvEscapedFilter.Append("[")
                lvEscapedFilter.Append(lvChar)
                lvEscapedFilter.Append("]")
            Else
                lvEscapedFilter.Append(lvChar)
            End If
        Next
        Return lvEscapedFilter.ToString
    End Function

    Friend Shared Function BuildFilter(ByVal piTable As Data.DataTable, ByVal piFilter As String) As System.Text.StringBuilder
        Dim lvFilter As New System.Text.StringBuilder
        If Not String.IsNullOrEmpty(piFilter) Then
            Dim lvFilters() As String = piFilter.Split(" "c)
            For lvIndex As Int32 = 0 To lvFilters.Length - 1
                If (0 < lvFilter.Length) Then
                    lvFilter.Append(" AND ")
                End If
                lvFilter.Append("(")
                Dim lvValue = EscapeForLike(lvFilters(lvIndex))
                Dim lvNegate = lvValue.StartsWith("!")
                If lvNegate Then
                    lvValue = lvValue.Substring(1)
                End If
                Dim lvFilterCount As Int32 = 0
                Dim lvOriginalValue = lvValue
                For Each lvColumn As DataColumn In piTable.Columns
                    lvValue = lvOriginalValue ' restore the value so side-effects don't occur
                    If GetType(String).IsAssignableFrom(lvColumn.DataType) Then
                        Dim lvOp = "LIKE"
                        If (1 < lvValue.Length) Then
                            If (0 <= "=".IndexOf(lvValue.Substring(0, 1))) Then
                                lvOp = lvValue.Substring(0, 1)
                                lvValue = lvValue.Substring(1)
                            End If
                        End If
                        If (0 < lvFilterCount) Then
                            If lvNegate Then
                                lvFilter.Append(" AND ")
                            Else
                                lvFilter.Append(" OR ")
                            End If
                        End If
                        If lvNegate Then
                            lvFilter.Append("NOT ")
                        End If
                        If lvOp = "LIKE" Then
                            lvValue = String.Format("%{0}%", lvValue)
                            lvFilter.Append(String.Format("[{0}] {2} '{1}'", lvColumn.ColumnName, lvValue, lvOp))
                        Else
                            ' =A becomes ='A' OR LIKE '=A'
                            lvFilter.Append(String.Format("(([{0}] {2} '{1}') OR ([{0}] LIKE '%{2}{1}%'))", lvColumn.ColumnName, lvValue, lvOp))
                        End If
                        lvFilterCount += 1
                    ElseIf GetType(System.DateTime).IsAssignableFrom(lvColumn.DataType) Then
                        Dim lvDateValue As DateTime
                        If DateTime.TryParse(lvValue, lvDateValue) Then
                            If (0 < lvFilterCount) Then
                                If lvNegate Then
                                    lvFilter.Append(" AND ")
                                Else
                                    lvFilter.Append(" OR ")
                                End If
                            End If
                            If lvNegate Then
                                lvFilter.Append("NOT ")
                            End If
                            lvFilter.Append(String.Format("([{0}] >= #{1}#) AND ([{0}] < #{2}#)", lvColumn.ColumnName, lvDateValue, lvDateValue.AddDays(1)))
                            lvFilterCount += 1
                        End If
                    ElseIf GetType(System.Decimal).IsAssignableFrom(lvColumn.DataType) _
                    OrElse GetType(System.Double).IsAssignableFrom(lvColumn.DataType) _
                    OrElse GetType(System.Int64).IsAssignableFrom(lvColumn.DataType) _
                    OrElse GetType(System.Int32).IsAssignableFrom(lvColumn.DataType) _
                    OrElse GetType(System.Int16).IsAssignableFrom(lvColumn.DataType) Then
                        Dim lvOp = "="
                        If (1 < lvValue.Length) Then
                            If (0 <= "><=".IndexOf(lvValue.Substring(0, 1))) Then
                                lvOp = lvValue.Substring(0, 1)
                                lvValue = lvValue.Substring(1)
                            End If
                        End If
                        Dim lvNumValue As Double
                        If Double.TryParse(lvValue, lvNumValue) Then
                            If (0 < lvFilterCount) Then
                                If lvNegate Then
                                    lvFilter.Append(" AND ")
                                Else
                                    lvFilter.Append(" OR ")
                                End If
                            End If
                            If lvNegate Then
                                lvFilter.Append("NOT ")
                            End If
                            lvFilter.Append(String.Format("[{0}]{2}{1}", lvColumn.ColumnName, lvNumValue, lvOp))
                            lvFilterCount += 1
                        End If
                    End If
                Next
                lvFilter.Append(")")
            Next
        End If
        Return lvFilter
    End Function

#Region "Fast Filter"

    Private Class FastFilter
        Implements IFastFilter

        Private ReadOnly m_Control As Windows.Forms.Control
        Private ReadOnly m_Table As DataTable
        Friend Sub New(ByVal piControl As Windows.Forms.Control, ByVal piTable As Data.DataTable)
            Me.m_Control = piControl
            Me.m_Table = piTable
            AddHandler piControl.TextChanged, AddressOf Me.FilterTextChangedHandler
            AddHandler piControl.DoubleClick, AddressOf Me.FilterDoubleClickHandler
        End Sub

        Friend Sub Unregister()
            RemoveHandler Me.m_Control.TextChanged, AddressOf Me.FilterTextChangedHandler
            RemoveHandler Me.m_Control.DoubleClick, AddressOf Me.FilterDoubleClickHandler
            If Not IsNothing(Me.m_Timer) Then
                Me.m_Timer.Stop()
                Me.m_Timer.Dispose()
                Me.m_Timer = Nothing
            End If
        End Sub

        Private Sub FilterTextChangedHandler(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Me.StartDelayFilter()
        End Sub

        Private Sub FilterDoubleClickHandler(ByVal sender As Object, ByVal e As EventArgs)
            Me.ShowClosestMatches(Me.m_Control.Text)
        End Sub

        Private Sub StartDelayFilter()
            If IsNothing(m_Timer) Then
                Me.m_Timer = New Timers.Timer(500)
                AddHandler m_Timer.Elapsed, AddressOf TimerElapsedHandler
            End If
            Me.m_Timer.Stop()
            Me.m_Timer.Start()
        End Sub
        Private m_Timer As System.Timers.Timer

        Private Sub TimerElapsedHandler(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs)
            If Me.m_Control.InvokeRequired() Then
                Me.m_Control.Invoke(New EventHandler(Of Timers.ElapsedEventArgs)(AddressOf Me.TimerElapsedHandler), sender, e)
            Else
                Me.m_Timer.Stop()
                FilterResults()
            End If
        End Sub

        Private m_HasNullRow As Boolean
        Private Property HasNullRow() As Boolean Implements IFastFilter.HasNullRow
            Get
                Return Me.m_HasNullRow
            End Get
            Set(ByVal value As Boolean)
                Me.m_HasNullRow = value
                Me.FilterResults()
            End Set
        End Property

        Private Sub FilterResults()
            Dim lvTable As DataTable = Me.m_Table
            If Not IsNothing(lvTable) Then
                Dim lvFilter As System.Text.StringBuilder = BuildFilter(lvTable, Me.m_Control.Text)
                If Me.m_HasNullRow Then
                    If (0 < lvFilter.Length) Then
                        lvFilter.Append(" AND ")
                    End If
                    lvFilter.Append(BuildNullFilter(lvTable))
                End If
                Dim lvArgs As New FilterEventArgs(lvFilter.ToString)
                RaiseEvent Filtering(Me, lvArgs)
                lvTable.DefaultView.RowFilter = lvArgs.Filter
            End If
        End Sub

        Private Shared Function BuildNullFilter(ByVal piTable As Data.DataTable) As System.Text.StringBuilder
            Dim lvFilter As New System.Text.StringBuilder
            Dim lvFilterCount As Int32 = 0
            For Each lvColumn As DataColumn In piTable.Columns
                If GetType(String).IsAssignableFrom(lvColumn.DataType) Then
                    If (0 < lvFilterCount) Then
                        lvFilter.Append(" OR ")
                    End If
                    lvFilter.Append(String.Format("[{0}] IS NULL", lvColumn.ColumnName))
                    lvFilterCount += 1
                End If
            Next
            Return lvFilter
        End Function

        Friend Event Filtering(ByVal sender As Object, ByVal e As FilterEventArgs)

        Private Sub ShowClosestMatches(ByVal piFilter As String)
            Dim lvMatches As New Generic.List(Of String)
            Dim lvColumns As New Generic.List(Of DataColumn)
            For Each lvColumn As DataColumn In Me.m_Table.Columns
                If lvColumn.DataType Is GetType(System.String) Then
                    lvColumns.Add(lvColumn)
                End If
            Next
            Dim lvMinDistance As Int32 = 10
            For Each lvRow As DataRow In Me.m_Table.Rows
                For Each lvColumn As DataColumn In lvColumns
                    Dim lvCellValue As Object = lvRow(lvColumn)
                    If lvCellValue IsNot DBNull.Value Then
                        Dim lvDistance As Int32 = FilterUtility.LevenshteinDistance(piFilter, CStr(lvCellValue))
                        If lvDistance <= lvMinDistance Then
                            If 2 * lvDistance < piFilter.Length + CStr(lvCellValue).Length Then
                                lvMinDistance = lvDistance
                                lvMatches.Add(String.Format("{0} LIKE '{1}'", lvColumn.ColumnName, CStr(lvCellValue)))
                                Exit For
                            End If
                        End If
                    End If
                Next
            Next
            If (0 < lvMatches.Count) Then
                Dim lvFilter As New System.Text.StringBuilder
                For lvIndex As Int32 = 0 To lvMatches.Count - 1
                    If (0 < lvIndex) Then
                        lvFilter.Append(" OR ")
                    End If
                    lvFilter.Append(lvMatches(lvIndex))
                Next
                Me.m_Table.DefaultView.RowFilter = lvFilter.ToString
            End If
        End Sub
    End Class

    Friend Shared Event Filtering(ByVal sender As Object, ByVal e As FilterEventArgs)

    Friend Class FilterEventArgs
        Inherits EventArgs
        Friend Filter As String
        Friend Sub New(ByVal piFilter As String)
            Me.Filter = piFilter
        End Sub
    End Class

    Private Shared Sub FilteringHandler(ByVal sender As Object, ByVal e As FilterEventArgs)
        RaiseEvent Filtering(sender, e)
    End Sub

    Private Shared m_FastFilters As New Generic.Dictionary(Of Control, FastFilter)
    Friend Shared Sub RegisterFilter(ByVal piControl As Windows.Forms.Control, ByVal piTable As Data.DataTable)
        Dim lvFastFilter As FastFilter = Nothing
        If m_FastFilters.ContainsKey(piControl) Then
            lvFastFilter = m_FastFilters(piControl)
            lvFastFilter.Unregister()
            m_FastFilters.Remove(piControl)
            RemoveHandler lvFastFilter.Filtering, AddressOf FilteringHandler
        End If
        lvFastFilter = New FastFilter(piControl, piTable)
        m_FastFilters.Add(piControl, lvFastFilter)
        AddHandler lvFastFilter.Filtering, AddressOf FilteringHandler
    End Sub

    Friend Shared Function FindFilter(ByVal piControl As Windows.Forms.Control) As IFastFilter
        If m_FastFilters.ContainsKey(piControl) Then
            Return m_FastFilters(piControl)
        Else
            Return Nothing
        End If
    End Function

#End Region

End Class
