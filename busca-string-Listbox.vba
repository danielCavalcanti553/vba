Private Sub btnBusca_Click()

Dim i As Integer
Dim busca As String
i = 1

busca = txBusca.Value

With Worksheets(1).Range("b1:b19")
    Set c = .Find(busca, LookIn:=xlValues)
    
    ListBox1.ColumnCount = 3
    ListBox1.AddItem
    ListBox1.List(0, 0) = "ID"
    ListBox1.List(0, 1) = "NOME"
    ListBox1.List(0, 2) = "VENDA"
    
    If Not c Is Nothing Then
        firstAddress = c.Address
        Do
            'c.Value = c.Column
            'ListBox1.AddItem c.Value
            
            
            ListBox1.AddItem
            ListBox1.List(i, 0) = Cells(c.Row, 1)
            ListBox1.List(i, 1) = Cells(c.Row, 2)
            ListBox1.List(i, 2) = Cells(c.Row, 3)
            i = i + 1
            
            Set c = .FindNext(c)
         Loop While Not c Is Nothing And c.Address <> firstAddress
    End If
End With
End Sub
