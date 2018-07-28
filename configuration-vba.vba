Sub INIT_CONFIG_INTERFACE()
    
    ' DESATIVA RIBBON
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",false)"
    
    ' DESATIVA GRIDS
    ActiveWindow.DisplayGridlines = False
    
    ' DESATIVA BARRA DE FORMULAS
    Application.DisplayFormulaBar = False
    
    'DESATIVA BARRA DE STATUS
    Application.DisplayStatusBar = False
    
    'DESATIVA TABS PLANILHAS
    ActiveWindow.DisplayWorkbookTabs = False
    
    'DESATIVA TÍTULOS (LINHAS E COLUNAS)
    ActiveWindow.DisplayHeadings = False
    
    ActiveSheet.DisplayPageBreaks = False
    
End Sub


Sub OUT_CONFIG_INTERFACE()

    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
    ActiveWindow.DisplayGridlines = True
    Application.DisplayFormulaBar = True
    Application.DisplayStatusBar = True
    ActiveWindow.DisplayHeadings = True
    ActiveWindow.DisplayWorkbookTabs = True
    'ActiveWindow.DisplayWorkbookTabs = True
    ActiveSheet.DisplayPageBreaks = False
End Sub




Public Sub OTIMIZATION_INIT()

    'Desabilita atualização de tela
    Application.ScreenUpdating = False
    
    'Desativa modo automático de cálculo
    Application.Calculation = xlManual
    
    'Desativa os eventos no excel
    Application.EnableEvents = False
    
    'Desabilita as quebras de páginas
    ActiveSheet.DisplayPageBreaks = False
    
End Sub


Public Sub OTIMIZATION_OUT()

    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    
End Sub


' OUTRAS FUNÇÕES
Sub PASS(namePlan)
    Dim plan As Worksheet
    Set plan = Sheets(namePlan)
    plan.Protect '123456'
End Sub

Sub UNPASS(namePlan)
    Dim plan As Worksheet
    Set plan = Sheets(namePlan)
    plan.Unprotect '123456'
End Sub

Sub total()

    Call UNPASS("Home")
    Call OTIMIZATION_INIT
    
    Dim n1, n2 As Double
    n1 = Range("CEL_N1")
    n2 = Range("CEL_N2")
    Range("CEL_TOTAL") = n1 + n2
    
    Call OTIMIZATION_OUT
    Call PASS("Home")
    
    
End Sub
