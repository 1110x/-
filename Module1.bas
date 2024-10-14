Attribute VB_Name = "Module1"
attribute VB_Name = "Module1"
Sub 野1()
Attribute 野1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 野1 野
'

For r = 14 To 75

    Sheets("").Select
    Rows("1:1").Select
    A = Sheets("TESTS").Cells(r, "H").text
    B = Sheets("TESTS").Cells(r, "I").text
    
    Selection.Replace what:=A, Replacement:=B, lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
Next r
Debug.Print ""
End Sub
Sub 野2()
Attribute 野2.VB_ProcData.VB_Invoke_Func = " \n14"

X = 2
For c = 3 To 63
Sheets("TESTS").Cells(X, "H") = Sheets("萸살궄").Cells(1, c)
X = X + 1
Next c



End Sub

Sub 野3()

X = 2
For c = 13 To 196 Step (3)
Sheets("TESTS").Cells(X, "I") = Sheets("").Cells(1, c)
X = X + 1
Next c



End Sub
