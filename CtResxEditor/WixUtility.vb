Friend Class WixUtility

    Friend Shared Sub MergeWxlDataSet(ByVal piFileName As String, ByVal piDataSet As NewDataSet)
        Debug.Assert(".WXL".Equals(IO.Path.GetExtension(piFileName), StringComparison.InvariantCultureIgnoreCase), "Expecting a .WXL file")
        Dim lvDom As New System.Xml.XmlDocument
        lvDom.Load(piFileName)
        Dim lvNode As System.Xml.XmlElement = lvDom.DocumentElement
        Dim lvCultureCode As String = lvNode.GetAttribute("Culture")
        Dim lvColumn As DataColumn = ResUtility.ForceCultureColumn(lvCultureCode, piDataSet.ResxData)
        Dim lvNsManager As Xml.XmlNamespaceManager = New Xml.XmlNamespaceManager(lvDom.NameTable)
        lvNsManager.AddNamespace("wix", "http://schemas.microsoft.com/wix/2006/localization")
        Dim lvEnumerator = lvNode.SelectNodes("wix:String[@Id]", lvNsManager)
        For Each lvNode In lvEnumerator
            Dim lvRow As NewDataSet.ResxDataRow = ForceWxlRow(lvNode.GetAttribute("Id"), piDataSet.ResxData)
            If lvRow(lvColumn) Is DBNull.Value OrElse Not String.Equals(lvRow(lvColumn), lvNode.InnerText) Then
                lvRow(lvColumn) = lvNode.InnerText
                lvRow.Version = Now
            Else
                ' do nothing
            End If
        Next
    End Sub

    Private Shared Function ForceWxlRow(ByVal piResourceId As String, ByVal piResxData As NewDataSet.ResxDataDataTable) As NewDataSet.ResxDataRow
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
