VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AIntegration()
'
' AIntegration Macro
'
'There are various variables which determine the profit of billed labor, principally spread (the difference between the hourly cost of employing and the amount
'billed) and the hours worked. The ecel sheet used formulas to manipulate these three variables (profit, spread, and hours billed) such that, by holding one of
'those constant, how the one variable changes when another does. Profit in this instance is a function of the yearly compensation (% of salary).
'
    Range("C15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-7]C[-2]"
    Range("D15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-7]C[-1]"
    Range("E15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-7]C[-3]"
    Range("F15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-7]C"
    Range("G7:G8").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("G10:G11").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("G13:G14").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
End Sub
Sub BIntegration()
'
' BIntegration Macro

    Range("C15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-4]C"
    Range("D15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-4]C[-3]"
    Range("E15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-4]C[-3]"
    Range("F15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-4]C"
    Range("G7:G8").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("G13:G14").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("G10:G11").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("H16").Select
End Sub
Sub CIntegration()
'
' CIntegration Macro

    Range("C15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C[-2]"
    Range("D15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C[-2]"
    Range("E15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C[-2]"
    Range("F15").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-1]C"
    Range("G13:G14").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    Range("G10:G11").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("G7:G8").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub
