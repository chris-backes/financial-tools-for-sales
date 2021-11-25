Attribute VB_Name = "Module1"
Sub BR_TotalCost_MarkUp()
Attribute BR_TotalCost_MarkUp.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BR_TotalCost_MarkUp Macro
'

'
    Range("C8").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((RC[-2]*R[-4]C[4])-(R[-4]C[4]))/(R[-4]C[3]*R[-4]C[2]))"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN((R[-4]C[3]*RC[-3])-(R[-4]C[3]),2)"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN(RC[-5]*R[-4]C[1],2)"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = _
        "=ROUND(((R[-7]C[4]+(R[-7]C[-1]*RC[-2]/RC[-1]))/(R[-7]C[4])),2)"
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN((R[-7]C[3]*RC[-1])-(R[-7]C[3]),2)"
    Range("F11").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN(R[-7]C[1]*RC[-3],2)"
    Range("C14").Select
    ActiveCell.FormulaR1C1 = _
        "=MROUND((RC[-1]*R[-10]C[-1])/((R[-10]C[4]*RC[-2])-(R[-10]C[4])),40)"
    Range("D14").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN((R[-10]C[3]*RC[-3])-(R[-10]C[3]),2)"
    Range("F14").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN(RC[-5]*R[-10]C[1],2)"
    Range("F15").Select
End Sub
Sub BR_Normalized_MarkUp()
Attribute BR_Normalized_MarkUp.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BR_Normalized_MarkUp Macro
'

'
    Range("C8").Select
    ActiveCell.FormulaR1C1 = _
        "=(RC[-1]*((RC[-2]*R[-4]C[3])-(R[-4]C[4]))/(R[-4]C[3]*R[-4]C[2]))"
    Range("D8").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN((R[-4]C[2]*RC[-3])-(R[-4]C[3]),2)"
    Range("F8").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN(RC[-5]*R[-4]C[0],2)"
    Range("D11").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN((R[-7]C[2]*RC[-1])-(R[-7]C[3]),2)"
    Range("C11").Select
    ActiveCell.FormulaR1C1 = _
        "=ROUND(((R[-7]C[4]+(R[-7]C[-1]*RC[-2]/RC[-1]))/(R[-7]C[3])),2)"
    Range("F11").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN(R[-7]C*RC[-3],2)"
    Range("F14").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN(RC[-5]*R[-10]C,2)"
    Range("D14").Select
    ActiveCell.FormulaR1C1 = "=ROUNDDOWN((R[-10]C[2]*RC[-3])-(R[-10]C[3]),2)"
    Range("C14").Select
    ActiveCell.FormulaR1C1 = _
        "=MROUND((RC[-1]*R[-10]C[-1])/((R[-10]C[3]*RC[-2])-(R[-10]C[4])),40)"
    Range("C15").Select
End Sub
Sub PY_Hourly()
Attribute PY_Hourly.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PY_Hourly Macro
'

'
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "45"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]*RC[3]"
    Range("B5").Select
    
    Range("A3:A4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("B3:B4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
Sub PY_Salary()
Attribute PY_Salary.VB_ProcData.VB_Invoke_Func = " \n14"
'
' PY_Salary Macro
'

'
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "100000"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "=RC[1]/RC[4]"
    Range("A5").Select
    
    Range("B3:B4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Range("A3:A4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub
