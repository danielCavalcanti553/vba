' https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/format-function-visual-basic-for-applications

Function formatDate()
' DATA

    tdate = "25/01/1989"
    nDate = CDate(tdate)
    
    Debug.Print Format(tdate, "d")
    Debug.Print Format(tdate, "dd")
    Debug.Print Format(tdate, "ddd")
    Debug.Print Format(tdate, "dddd")
    Debug.Print
    Debug.Print Format(tdate, "m")
    Debug.Print Format(tdate, "mm")
    Debug.Print Format(tdate, "mmm")
    Debug.Print Format(tdate, "mmmm")
    Debug.Print
    Debug.Print Format(tdate, "yy")
    Debug.Print Format(tdate, "yyyy")

End Function


Function formatHora()
' HORA

    tdate = "3:25:05"
    nDate = CDate(tdate)
    
    Debug.Print nDate
    Debug.Print Format(nDate, "h")
    Debug.Print Format(nDate, "hh")
    Debug.Print Format(nDate, "n")
    Debug.Print Format(nDate, "nn")
    Debug.Print Format(nDate, "s")
    Debug.Print Format(nDate, "ss")

End Function

Function formatDataHora()
' HORA

    tdate = "25/01/2018 3:25:05"
    nDate = CDate(tdate)
    
    Debug.Print nDate
    Debug.Print Format(nDate, "dd ""de"" mmmm ""de"" yyyy")


End Function

Function calcDateDiff()
    
    inicioDate = CDate("07/02/2018")
    fimDate = CDate("21/03/2018")
    
    nDate = dateDiff("d", inicioDate, fimDate)
    Debug.Print "Dias " & nDate
    
    nDate = dateDiff("m", inicioDate, fimDate)
    Debug.Print "Meses " & nDate
    
    nDate = dateDiff("yyyy", inicioDate, fimDate)
    Debug.Print "Anos " & nDate
    
End Function





' Não realizar
Function calcDateDiffNow()
    
    dataNascimento = #6/1/1989#
    dataHoje = Date
    
    mNasc = CInt(Format(dataNascimento, "m"))
    mDate = CInt(Format(dataHoje, "m"))
    
    dNasc = CInt(Format(dataNascimento, "d"))
    dDate = CInt(Format(dataHoje, "d"))
    
    idade = dateDiff("yyyy", dataNascimento, dataHoje)
    
    If (mNasc <= mDate) And (dNasc > dDate) Then
            Debug.Print "V - Dia"
            idade = idade - 1
    End If
    
    Debug.Print idade

End Function



Sub defined_form()

' User-defined formats.
MyStr = Format(5459.4, "##,##0.00")    ' Returns "5,459.40".
MyStr = Format(334.9, "###0.00")    ' Returns "334.90".
MyStr = Format(5, "0.00%")    ' Returns "500.00%".
MyStr = Format("HELLO", "<")    ' Returns "hello".
MyStr = Format("This is it", ">")    ' Returns "THIS IS IT".



End Sub
