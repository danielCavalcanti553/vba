Dim i, linhaAtual As Integer

Private Sub btnExcluir_Click()
    Sheets("DADOS").Rows(linhaAtual).Delete
    
    Call proximo
    
End Sub

Private Sub btnProximo_Click()
Call proximo

End Sub



Sub proximo()
    
    If i < 1 Then
        i = 1
    End If
    

    
    i = i + 1
    
    If Sheets("DADOS").Cells(i, 1) = "" Then
        i = 2
    End If
    
    linhaAtual = Sheets("DADOS").Cells(i, 1).Row
    
    ID = Sheets("DADOS").Cells(i, 1)
    strCaminhoDestino = CStr(ThisWorkbook.Path) & "\imagens\CAR" & ID & ".jpg"
    Image.Picture = LoadPicture(strCaminhoDestino)
    
    txID.Value = Sheets("DADOS").Cells(i, 1)
    txModelo.Value = Sheets("DADOS").Cells(i, 2)
     txPlaca.Value = Sheets("DADOS").Cells(i, 3)
      txMarca.Value = Sheets("DADOS").Cells(i, 4)
       txCor.Value = Sheets("DADOS").Cells(i, 5)
       
  
       
    If Sheets("DADOS").Cells(i, 6).Value = "Sim" Then
        seguroSim.Value = True
    Else
        txSeguroNao.Value = True
    
    End If
    
    txAcessorios.Value = Sheets("DADOS").Cells(i, 7).Value
    
    'Dim LArray() As String
    'LArray = Split(Sheets("DADOS").Cells(i, 7).Value, ",")
    'Debug.Print LArray(1)

End Sub

Private Sub UserForm_Initialize()
Call proximo
End Sub
