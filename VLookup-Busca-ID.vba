Private Sub btnBusca_Click()

    On Error GoTo trataErro

    REG = Range("TOT_REG") + 1
    Set intervalo = Sheets("Planilha1").Range("A1:C" & REG)
    
    codigo = CInt(txCodigo.Value)
      
    txNome.Value = Application.WorksheetFunction.VLookup(codigo, intervalo, 2, False)
    txVendas.Value = Application.WorksheetFunction.VLookup(codigo, intervalo, 3, False)
    
    
Exit Sub

trataErro:
   texto = "Produto não localizado!"
   mensagem = MsgBox(texto, vbOKOnly + vbInformation)
End Sub
