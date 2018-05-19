Sub varArray()
    Dim nomes(5) As String
    nomes(0) = "Daniel"
    nomes(1) = "Fabio"
    nomes(2) = "Ana"
    nomes(3) = "Renata"
    nomes(5) = "Pedro"
    'nomes(6) = "Mariana" Erro
    MsgBox nomes(2)
    MsgBox nomes(4)
End Sub

Sub varArrayMulti()
 Dim nota(4, 3) As Double
 nota(1, 1) = 20
 nota(1, 2) = 40
 nota(1, 3) = 20

MsgBox nota(1, 2)
End Sub
