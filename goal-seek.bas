Attribute VB_Name = "Goal_Seek"
Sub GS_Hours_Percent()
Attribute GS_Hours_Percent.VB_ProcData.VB_Invoke_Func = "t\n14"

    Range("C14").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("C14").GoalSeek Goal:=Range("J15"), ChangingCell:=Range("B14")
    
End Sub
Sub GS_Hours_Mark_Up()

    Range("C14").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("C14").GoalSeek Goal:=Range("I15"), ChangingCell:=Range("A14")
    
End Sub
Sub GS_Percent_Markup()

    Range("C8").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("C8").GoalSeek Goal:=Range("I9"), ChangingCell:=Range("A8")

End Sub
Sub GS_Markup_Percent()

    Range("C11").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("C11").GoalSeek Goal:=Range("I12"), ChangingCell:=Range("A11")

End Sub
Sub GS_Percent_Hours()

    Range("C8").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("C8").GoalSeek Goal:=Range("J9"), ChangingCell:=Range("B8")

End Sub
Sub GS_Markup_Hours()

    Range("C11").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    Range("C11").GoalSeek Goal:=Range("J12"), ChangingCell:=Range("B11")

End Sub
