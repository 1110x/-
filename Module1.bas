Attribute VB_Name = "Module1"
Sub ��ũ��1()
Attribute ��ũ��1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��ũ��1 ��ũ��
'

For r = 14 To 75

    Sheets("������������").Select
    Rows("1:1").Select
    A = Sheets("TESTS").Cells(r, "H").text
    B = Sheets("TESTS").Cells(r, "I").text
    
    Selection.Replace what:=A, Replacement:=B, lookat:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
Next r
Debug.Print "��"
End Sub
Sub ��ũ��2()
Attribute ��ũ��2.VB_ProcData.VB_Invoke_Func = " \n14"

X = 2
For c = 3 To 63
Sheets("TESTS").Cells(X, "H") = Sheets("�м�����ڷ�").Cells(1, c)
X = X + 1
Next c



End Sub

Sub ��ũ��3()

X = 2
For c = 13 To 196 Step (3)
Sheets("TESTS").Cells(X, "I") = Sheets("������������").Cells(1, c)
X = X + 1
Next c



End Sub
