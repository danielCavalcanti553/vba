Sub objRange()
    Dim intervalo As Range
    Set intervalo = Range("A1:a20")
    intervalo.Interior.ColorIndex = 6
End Sub

Sub objSheet()
    Dim plan As Worksheet
    Set plan2 = Sheets("planilha2")
    plan2.Select
End Sub

Sub objAplication()
    Application.DisplayFormulaBar = True

End Sub


Sub objActiveWindow()
    ActiveWindow.DisplayHeadings = True
End Sub

Sub inputApp()

    result = MsgBox("Digite texto", vbYesNo)
    MsgBox result

End Sub

Sub workFunction()
    
    Set myRange = Worksheets("Planilha1").Range("c6:c10")
    tot = Application.WorksheetFunction.Sum(myRange)
    MsgBox tot

End Sub
