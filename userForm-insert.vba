Private Sub btnCadastrar_Click()
    
    ID = Range("PROX_REG")
    R = ID + 1
    
    Call copiarImagem(lblCaminho.Caption)
    
    Cells(R, 1) = ID
    Cells(R, 2) = txModelo.Value
    Cells(R, 3) = txPlaca.Value
    Cells(R, 4) = txMarca.Value
    Cells(R, 5) = txCor.Value
    Cells(R, 6) = obterSeguro
    Cells(R, 7) = obterAcessorios
    
    Call removeFields

End Sub

Private Sub CarregarImagem_Click()

    Dim strCaminhoImagem, strCaminhoDestino As String
    
    ID = Range("PROX_REG")
    
    strCaminhoDestino = CStr(ThisWorkbook.Path) & "\imagens\" & ID & ".jpg"
    strCaminhoImagem = Application.GetOpenFilename("*.jpg,*.jpg,*.bmp,*.bmp")
    
    Image1.Picture = LoadPicture(strCaminhoImagem)
    
    lblCaminho.Caption = strCaminhoImagem


End Sub

Sub copiarImagem(strCaminhoImagem As String)
    ID = Range("PROX_REG")
    strCaminhoDestino = CStr(ThisWorkbook.Path) & "\imagens\" & ID & ".jpg"
    FileCopy strCaminhoImagem, strCaminhoDestino
End Sub

Private Sub UserForm_Initialize()
    
    Call carregaMarcas
    Call carregaCor
    
    
    strCaminhoDestino = CStr(ThisWorkbook.Path) & "\config\noimage.jpg"
    Image1.Picture = LoadPicture(strCaminhoDestino)
    Debug.Print strCaminhoDestino

End Sub

Sub carregaMarcas()
    Dim marcas, marca As Range
    Set marcas = Sheets("MARCA").Range("A1:A100")
    
    For Each marca In marcas
        If marca <> "" Then
            'Exit For
             txMarca.AddItem marca
        End If
    Next marca
End Sub

Sub carregaCor()
    Dim cores, cor As Range
    Set cores = Sheets("COR").Range("A1:A100")
    
    For Each cor In cores
        If cor <> "" Then
            'Exit For
             txCor.AddItem cor
        End If
    Next cor
End Sub

Sub removeMarcas()
    txMarca.Clear
End Sub

Function obterAcessorios() As String


    acessorios = ""
    
    For i = 0 To txAcessorios.ListCount - 1
    
        If i = 0 Then
            acessorios = acessorios + "" + txAcessorios.Column(0, i)
        ElseIf txAcessorios.Selected(i) = True Then
            acessorios = acessorios + ", " + txAcessorios.Column(0, i)
        End If
    Next
    obterAcessorios = acessorios
    
End Function


Function obterSeguro() As String
Dim seguro As String


If txSeguroSim.Value = True Then
    seguro = "SIM"
Else
    seguro = "NAO"
End If
obterSeguro = seguro

End Function

Sub removeFields()

    txModelo.Value = ""
    txPlaca.Value = ""
    txMarca.Value = ""
    txCor.Value = ""
    
    txSeguroSim.Value = True
    
    For i = 0 To txAcessorios.ListCount - 1
    
        txAcessorios.Selected(i) = False
        
    Next
    
End Sub
