
Sub estrutControle()

    Dim a As Integer
    a = CDbl(InputBox("Digite um número"))
    
    If a > 100 Then
        MsgBox "Seu número é maior que 100"
    Else
        MsgBox "Seu número é menor ou igual a 100"
    End If

End Sub

Sub estrutIfElse()

    Dim cod As Integer
    cod = CDbl(InputBox("Digite um número"))
    
    If cod = 1 Then
        MsgBox "Zona Sul"
    ElseIf cod = 2 Then
        MsgBox "Zona Norte"
    ElseIf cod = 3 Then
        MsgBox "Zona Oeste"
    Else
        MsgBox "Código inválido"
    End If

End Sub

Sub estrutControleAnd()
    Dim salarioBruto As Double
  salarioBruto = CDbl(InputBox("Digite Salário", "Encontre Alíquota INSS"))
  
  If salarioBruto <= 1693.72 Then
        MsgBox "Alíquota de 8%"
    End If

    If salarioBruto >= 1693.73 And salarioBruto <= 2822.9 Then
        MsgBox "Alíquota de 9%"
    End If
    
  
    If salarioBruto >= 2822.91 And salarioBruto <= 5645.8 Then
        MsgBox "Alíquota de 11%"
    End If
    
    If salarioBruto > 5645.8 Then
        MsgBox "Desconto máximo R$ 621.00"
    End If
    
End Sub

Sub estrutControleOr()

    ' Estacionamento gratuito se ficou ate 15 minutos ou
    ' gastou acima de R$ 50,00
    
    Dim minutos As Integer
    Dim gasto As Double
    
    minutos = CInt(InputBox("Digite número de minutos", "Estacionamento"))
    gasto = CDbl(InputBox("Digite o Gasto", "Estacionamento"))
  
    
    If minutos < 15 Or gasto >= 50 Then
        MsgBox "Estacionamento Gratuito"
    Else
        MsgBox "Pague o estacionamento"
    End If

End Sub
