Attribute VB_Name = "Module1"
Public Sub ShowData1(Rs As ADODB.Recordset, Dgrid As MSFlexGrid)
     Rs.MoveFirst
     Dim i As Integer
     i = 0
     Dgrid.Rows = Rs.RecordCount + 1
     Dgrid.Cols = Rs.Fields.Count
     For j = 0 To Rs.Fields.Count - 1
        Dgrid.TextMatrix(0, j) = Rs.Fields(j).Name
     Next j
     Do While Not Rs.EOF
     i = i + 1
     For j = 0 To Rs.Fields.Count - 1
        If Not IsNull(Rs.Fields(j).Value) Then
         Dgrid.TextMatrix(i, j) = Rs.Fields(j).Value
        End If
     Next j
     Rs.MoveNext
     Loop
End Sub

