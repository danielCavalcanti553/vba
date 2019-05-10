Sub funcaoSeAnd()
    
    If Range("numero1") > Range("valor") And Range("numero2") > Range("valor") Then
        Range("h9") = "Os dois números são maiores que " & Range("valor")
    Else
        Range("h9") = "Algum ou nenhum número é menor que " & Range("valor")
    End If
    
End Sub

Sub funcaoSeOr()
    
    If Range("numero1") > Range("valor") Or Range("numero2") > Range("valor") Then
        Range("i9") = "Pelo menos 1 número é maior que " & Range("valor")
    Else
        Range("i9") = "Nenhum número é maior que " & Range("valor")
    End If
    
End Sub
    
Sub funcaoSeXor()
    
    If Range("numero1") < Range("valor") Xor Range("numero2") < Range("valor") Then
        Range("j9") = "Somente 1 dos números é maior que " & Range("valor")
    Else
        Range("j9") = "Os números são menores maiores ou menores que " & Range("valor")
    End If
    
End Sub

Sub funcaoSeAll()
 Call funcaoSeAnd
 Call funcaoSeOr
 Call funcaoSeXor
End Sub




