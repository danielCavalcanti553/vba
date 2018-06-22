Dim i, linhaAtual As Integer

Private Sub btnAtualizar_Click()
    
    Sheets("DADOS").Cells(linhaAtual, 2) = txModelo.Value
    Sheets("DADOS").Cells(linhaAtual, 3) = txPlaca.Value
    Sheets("DADOS").Cells(linhaAtual, 4) = txMarca.Value
    Sheets("DADOS").Cells(linhaAtual, 5) = txCor.Value
    Sheets("DADOS").Cells(linhaAtual, 6) = obterSeguro
    Sheets("DADOS").Cells(linhaAtual, 7) = txAcessorios.Value
    
    On Error GoTo TratarErro
    
    Call copiarImagem(lblCaminho.Caption)
    
TratarErro:
    
End Sub

Sub copiarImagem(strCaminhoImagem As String)
    ID = txID.Value
    strCaminhoDestino = CStr(ThisWorkbook.Path) & "\imagens\CAR" & ID & ".jpg"
    FileCopy strCaminhoImagem, strCaminhoDestino
End Sub

Function obterSeguro() As String
Dim seguro As String


If seguroSim.Value = True Then
    seguro = "SIM"
Else
    seguro = "NAO"
End If
obterSeguro = seguro

End Function

Private Sub btnExcluir_Click()

    confirm = MsgBox("Deseja realmente excluir", vbYesNo)
    
    If confirm = 6 Then
        Sheets("DADOS").Rows(linhaAtual).Delete
        Call proximo
    End If
    
End Sub

Private Sub btnImagem_Click()

 Dim strCaminhoImagem As String
    
    strCaminhoImagem = Application.GetOpenFilename("*.jpg,*.jpg,*.bmp,*.bmp")
    lblCaminho.Caption = strCaminhoImagem
    Image.Picture = LoadPicture(strCaminhoImagem)


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
    lblCaminho.Caption = CStr(ThisWorkbook.Path) & "\imagens\CAR" & ID & ".jpg"
    
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
