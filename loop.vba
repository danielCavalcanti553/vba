Sub vetor()
    
    Dim compras(1 To 5) As String
    compras(1) = "Arroz"
    compras(2) = "Macarrão"
    compras(3) = "Feijão"
    compras(4) = "Leite"
    compras(5) = "Biscoito"
    
    Debug.Print compras(1)
    ' Exibir > Janela Verificação Imediata
End Sub

Sub matriz()

    Dim notas(1 To 2, 1 To 3) As Double
    notas(1, 1) = 55.5
    notas(1, 2) = 70
    notas(1, 3) = 77
    notas(2, 1) = 10
    notas(2, 2) = 90
    notas(3, 3) = 75
    
    Debug.Print notas(1, 3)
    Debug.Print notas(2, 3)
End Sub

Sub loopFor()

 For i = 1 To 10
 ' Rodar 10 vezes
    Debug.Print i
    Cells(i, 1).Value = i
    ' FIm
 Next
 
End Sub

Sub forVetor()

    Dim compras(1 To 5) As String
    compras(1) = "Arroz"
    compras(2) = "Macarrão"
    compras(3) = "Feijão"
    compras(4) = "Leite"
    compras(5) = "Biscoito"
    
    For i = 1 To 5
        Debug.Print compras(i)
        Cells(i, 2).Value = compras(i)
    Next
End Sub

Sub valCells()
    Cells(2, 3).Value = "Valor"
End Sub

Sub arrayVar()
    ' Armazenar tipos aleatórios (String, Integer, etc)
    Dim vet(1 To 3) As Variant
    vet(1) = 20
    vet(2) = "Daniel"
    Debug.Print vet(1)
    Debug.Print vet(2)
End Sub

Sub forMatriz()

    Dim notas(1 To 2, 1 To 3) As Double
    notas(1, 1) = 55.5
    notas(1, 2) = 70
    notas(1, 3) = 77
    notas(2, 1) = 10
    notas(2, 2) = 90
    notas(2, 3) = 75
    
    For x = 1 To 2
        For y = 1 To 3
            Debug.Print x & "-" & y & " - " & notas(x, y)
        Next
    Next
End Sub

Sub ex1()
    Dim numeros(1 To 7) As Integer
    numeros(1) = 20
    numeros(2) = 50
    numeros(3) = 80
    numeros(4) = 40
    numeros(5) = 80
    numeros(6) = 60
    numeros(7) = 65
    
    For x = 1 To 7
        If numeros(x) < 50 Then
            Debug.Print numeros(x)
        End If
    Next
    
End Sub

Sub ex2()
    Dim soma As Integer
    Dim numeros(1 To 7) As Integer
    numeros(1) = 20
    numeros(2) = 50
    numeros(3) = 80
    numeros(4) = 40
    numeros(5) = 80
    numeros(6) = 60
    numeros(7) = 65
    
    ' Retornem a soma do vetor
    For abc = 1 To 7
        soma = soma + numeros(abc)
    Next
    
    Debug.Print "Total: " & soma
End Sub

Sub ex3()

    Dim soma As Integer
    
    For i = 1 To 8
        soma = Cells(i, 1) + soma
    Next
    
    Debug.Print soma

End Sub


Sub ex4()

    Dim soma As Integer
    Dim i As Integer
    i = 1
    
    Do While Cells(i, 1) <> ""
        soma = soma + Cells(i, 1)
        i = i + 1
    Loop
    
    Debug.Print soma
    
End Sub

Sub vetorUbound()
    Dim compras(1 To 5) As String
    compras(1) = "Arroz"
    compras(2) = "Macarrão"
    compras(3) = "Feijão"
    compras(4) = "Leite"
    compras(5) = "Biscoito"
    
    ' Ubound retorna o último índice de um vetor
    For i = 1 To UBound(compras)
        Debug.Print compras(i)
        Cells(i, 2).Value = compras(i)
    Next
End Sub

Sub loopWhile()
    
    Dim total, n As Integer
    n = 1
    Do While n <> 0
        n = InputBox("Digite um número, 0 para sair")
        total = total + n
    Loop
    
    MsgBox total - 1
    
End Sub

Sub loopDoWhile()
    Dim total, n As Integer
    
    Do
        n = InputBox("Digite um número, 0 para sair")
        total = total + n
    Loop Until n = 0
    MsgBox total
    
End Sub

Sub ex5()
    
    ' Armazenar 5 números em vetor
    Dim numeros(1 To 5) As Integer
    
    
    'Criar um loop para armazenar 5 números
    For i = 1 To UBound(numeros)
        ' armazenar no vetor
        numeros(i) = InputBox("Digite um número")
        ' Testar
        ' Debug.Print numeros(i)
    Next
    
    ' Pegar os dados do vetor e colocar na planilha
    For i = 1 To UBound(numeros)
        ' Cells na posição i e coluna A (1)
        Cells(i, 1) = numeros(i)
        
        ' Verificar se o valor da celula é menor que 50
        If Cells(i, 1) < 50 Then
            Cells(i, 1).Font.Color = vbRed ' Se sim, céluna cor vermelho
        Else
                Cells(i, 1).Font.Color = vbBlue ' Se não, céluna cor vermelho
        End If
        
    Next
    
End Sub

Sub ex5v2()
    
    'Criar um loop para armazenar 5 números
    For i = 1 To 5
        ' armazenar no vetor
        Cells(i, 1) = InputBox("Digite um número")
        
        If Cells(i, 1) < 50 Then
            Cells(i, 1).Font.Color = vbRed ' Se sim, céluna cor vermelho
        Else
                Cells(i, 1).Font.Color = vbBlue ' Se não, céluna cor vermelho2
        End If
    Next
End Sub

Sub loopEach()

    Dim celulas, celula As Range
    Set celulas = Range("a1:a5")
    
    For Each celula In celulas
        
        celula.Font.Color = vbGreen
        celula.Value = 10
        
    Next celula
    
End Sub


Sub loopEachIntervalo()

    Dim celulas, celula As Range
    Set celulas = Range("DADOS")
    
    For Each celula In celulas
        
        celula.Font.Color = vbRed
        celula.Value = 10
        
    Next celula
    
End Sub

Sub loopEachIntervalo2()

    Dim celulas, celula As Range
    Set celulas = Application.Selection
    Dim tot As Integer
    
    For Each celula In celulas
        
            tot = tot + celula
        
    Next celula
    MsgBox "Total: " & tot
End Sub

Sub imprimeLocalAddress()
    Dim rg, area As Range
    Set rg = Application.Selection
    For Each area In rg.areas
        Debug.Print area.Address
    Next
    
End Sub
