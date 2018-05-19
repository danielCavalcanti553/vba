
Sub estrutCase()
    Dim horas As Integer
     horas = CInt(InputBox("Digite número de horas", "Estacionamento"))
   
   Select Case horas
    Case Is <= 1
        MsgBox "R$ 6,00"
    Case 2
        MsgBox "R$ 8,00"
    Case 3
        MsgBox "R$ 9,00"
    Case Else
        MsgBox "Diária - R$ 12,00"
   End Select

End Sub

Sub estrutCase2()
    Dim salarioBruto As Double
  salarioBruto = CDbl(InputBox("Digite Salário", "Encontre Alíquota INSS"))
    
   Select Case salarioBruto
    Case 0 To 1693.72
        MsgBox "Alíquota 8%"
    Case 1693.73 To 2822.9
        MsgBox "Alíquota 9%"
    Case 2822.91 To 5645.8
        MsgBox "Alíquota 11%"
    Case Else
        MsgBox "Desconto máximo R$ 621.00"
   End Select

End Sub

Sub calcDesconto()

    Dim tipoPag, msg As String
    
    msg = vbCrLf + "Crédito" + vbCrLf + "Débito" + vbCrLf + "Dinheiro" + vbCrLf + "Cheque"
    
    
    tipoPag = InputBox("Digite Forma Pagamento:" + msg, "Calcular Desconto")
    
   Select Case salarioBruto
    Case 0 To 1693.72
        MsgBox "Alíquota 8%"
    Case 1693.73 To 2822.9
        MsgBox "Alíquota 9%"
    Case 2822.91 To 5645.8
        MsgBox "Alíquota 11%"
    Case Else
        MsgBox "Desconto máximo R$ 621.00"
   End Select
   
End Sub

