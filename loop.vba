#LOOP

Sub vetor()
    Dim compras(1 To 5) As String
    
    compras(1) = "Arroz1"
    compras(2) = "Arroz2"
    compras(3) = "Arroz3"
    compras(4) = "Arroz4"
    compras(5) = "Arroz5"
    
    MsgBox compras(2)
    
End Sub

Sub matriz()
    Dim notas(1 To 2, 1 To 2) As String
    
    notas(1, 1) = 20
    notas(1, 2) = 30
    notas(2, 1) = 40
    notas(2, 2) = 50
    
    
    MsgBox notas(2, 2)
    
End Sub

Sub loopFor()

For i = 1 To 10
    Cells(i, 1) = i
Next i

End Sub

Sub loopFor2()
'Soma coluna A até linha 10

Dim tot As Integer
tot = 0
For i = 1 To 10
    tot = Cells(i, 1) + tot
Next i

MsgBox "Total " & tot

End Sub

Sub loopFor3()
Dim vazio As Integer
vazio = 0
For cont = 1 To 10
    If Cells(cont, 1).Value = "" Then
        vazio = vazio + 1
    End If
    
    
Next cont

MsgBox "Vazias " & vazio

End Sub

Sub loopEach()
    Dim rg As Range
    Set rg = Range("b1:c10")
    
    For Each cel In rg.Cells
        Debug.Print cel.Address
    Next cel
End Sub

Sub loopEach2()
    Dim rg As Range
    Set rg = Range("a1:c10")
    
    For Each cel In rg.Cells
        If cel.Value > 5 Then
            cel.Font.Color = vbRed
        End If
    Next cel
End Sub

Sub nomes()
    MsgBox Application.WorksheetFunction.Sum(Range("DATA"))
End Sub

Sub pegarValoresArea()
    Dim cels, cel As Range
    Set cels = Application.Selection
    For Each cel In cels.Cells
        Debug.Print cel.Value
    Next
End Sub


Sub somaSelectedValues()
    Dim tot As Integer
    Dim cels, cel As Range
    Set cels = Application.Selection
    tot = 0
    For Each cel In cels.Cells
        tot = tot + cel.Value
    Next
    Debug.Print tot
End Sub

Sub imprimeLocalAddress()
    Dim rg, area As Range
    Set rg = Application.Selection
    For Each area In rg.areas
        Debug.Print area.Address
    Next
    
End Sub
