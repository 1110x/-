Attribute VB_Name = "Module1"
Sub 매크로1()
Attribute 매크로1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로1 매크로
'

For r = 14 To 75

    Sheets("견적발행정보").Select
    Rows("1:1").Select
    A = Sheets("TESTS").Cells(r, "H").text
    B = Sheets("TESTS").Cells(r, "I").text
    
    Selection.Replace what:=A, Replacement:=B, lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
Next r
Debug.Print "끝"
End Sub
Sub 매크로2()
Attribute 매크로2.VB_ProcData.VB_Invoke_Func = " \n14"

X = 2
For c = 3 To 63
Sheets("TESTS").Cells(X, "H") = Sheets("분석결과자료").Cells(1, c)
X = X + 1
Next c



End Sub

Sub 매크로3()

X = 2
For c = 13 To 196 Step (3)
Sheets("TESTS").Cells(X, "I") = Sheets("견적발행정보").Cells(1, c)
X = X + 1
Next c



End Sub
