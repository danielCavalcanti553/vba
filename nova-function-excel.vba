Public Function NovaFuncion(valor As Integer, valor2 As Integer) As Integer
    NovaFuncion = valor * valor2
End Function

'Define a categoria da função
Sub definirCategoria()
    Application.MacroOptions Macro:="NovaFuncion", Category:=7
End Sub

'0-Sem categoria
'1-Financeira
'2-Data e Hora
'3-Matemática e Trigonometria
'4-Estatística
'5-Procura e Referência
'6-Banco de Dados
'7-Texto
'8-Lógico
'9-Informação
