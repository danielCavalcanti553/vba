'' FUNÇÃO SE
' Com dados digitados na planilha, prencha os dados aplicando a fórmula SOMA

Sub funcaoSOMA()
    'Range("E2") = Range("a2") + Range("b2") + Range("c2") + Range("d2")
    Range("E2").Formula = "=sum(A2:D2)"
    Range("F2") = WorksheetFunction.Sum(Range("a1:d2"))

End Sub


