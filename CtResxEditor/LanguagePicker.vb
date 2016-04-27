Option Explicit On
Option Strict On

Imports System.Globalization

Friend Class LanguagePicker

    Private ReadOnly m_Cultures As CultureInfo()
    Private ReadOnly m_SelectedCultures As Generic.List(Of CultureInfo)

    Friend Sub New()
        Me.m_Cultures = System.Globalization.CultureInfo.GetCultures(Globalization.CultureTypes.FrameworkCultures)
        Me.m_SelectedCultures = New Generic.List(Of CultureInfo)
    End Sub

    Friend Sub SelectCulture(ByVal piCulture As CultureInfo)
        Debug.Assert(Not Me.m_SelectedCultures.Contains(piCulture), "Already selected")
        Me.m_SelectedCultures.Add(piCulture)
    End Sub

    Friend Function DataTable() As DataTable
        Dim lvResult As New DataTable
        Dim lvCultureColumn As DataColumn = lvResult.Columns.Add("Culture")
        Dim lvDescriptionColumn As DataColumn = lvResult.Columns.Add("Description")
        Dim lvSelectedColumn As DataColumn = lvResult.Columns.Add("Selected", GetType(System.Boolean))
        For Each lvCulture As CultureInfo In Me.m_Cultures
            Dim lvRow As DataRow = lvResult.NewRow
            lvRow(lvCultureColumn) = lvCulture.Name
            lvRow(lvDescriptionColumn) = lvCulture.EnglishName
            lvRow(lvSelectedColumn) = Me.m_SelectedCultures.Contains(lvCulture)
            lvResult.Rows.Add(lvRow)
        Next
        Return lvResult
    End Function

    Friend Function PickCultures() As Generic.List(Of CultureInfo)
        Dim lvFilterText As New TextBox
        lvFilterText.Dock = DockStyle.Top
        Dim lvDataGridView As New DataGridView
        Dim lvDataTable As DataTable = Me.DataTable
        lvDataGridView.DataSource = lvDataTable
        lvDataGridView.AllowUserToAddRows = False
        lvDataGridView.AllowUserToDeleteRows = False
        Using lvForm As New Form
            lvForm.Controls.Add(lvDataGridView)
            lvForm.Controls.Add(lvFilterText)
            lvDataGridView.Dock = DockStyle.Fill
            lvForm.Size = New Drawing.Size(400, 380)
            lvForm.Text = "Cultures"
            FilterUtility.RegisterFilter(lvFilterText, lvDataTable)
            lvDataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells)
            Select Case lvForm.ShowDialog
                Case DialogResult.OK, DialogResult.Cancel
                    Dim lvResult As New Generic.List(Of CultureInfo)
                    Dim lvSelectedColumn As DataColumn = lvDataTable.Columns("Selected")
                    Dim lvCultureColumn As DataColumn = lvDataTable.Columns("Culture")
                    For Each lvRow As DataRow In lvDataTable.Rows
                        If Not DBNull.Value Is lvRow(lvSelectedColumn) Then
                            If CBool(lvRow(lvSelectedColumn)) Then
                                lvResult.Add(New CultureInfo(CStr(lvRow(lvCultureColumn))))
                            End If
                        End If
                    Next
                    Return lvResult
            End Select
        End Using
        Return Nothing
    End Function

End Class
