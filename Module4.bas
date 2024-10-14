Attribute VB_Name = "Module4"
Dim c As Integer

Sub 의뢰일자_일괄()
For c = Columns("B").Column To Columns("AY").Column
Sheets("의뢰입력").Cells(2, c) = Sheets("의뢰입력").Cells(2, "B")
Next c
End Sub
Sub 채취일자_일괄()
For c = Columns("B").Column To Columns("AY").Column
Sheets("의뢰입력").Cells(3, c) = Sheets("의뢰입력").Cells(3, "B")
Next c
End Sub
Sub 채취시간_일괄()
    Dim ST As Date
    Dim AT As Date
'    Dim c As Integer

    ST = Sheets("의뢰입력").Cells(4, "B").Value
    For c = Columns("C").Column To Columns("AY").Column
        AT = AT + DateAdd("n", 10, TimeValue("00:00"))
        Sheets("의뢰입력").Cells(4, c).Value = ST + AT
    Next c
End Sub
Sub 의뢰사업장_일괄()
For c = Columns("B").Column To Columns("AY").Column
 Sheets("의뢰입력").Cells(5, c) = Sheets("의뢰입력").Cells(5, "B")
Next c
End Sub

Sub 입회자_일괄()

For c = Columns("B").Column To Columns("AY").Column
Sheets("의뢰입력").Cells(8, c) = Sheets("의뢰입력").Cells(8, "B")
Next c

End Sub
Sub 시료채취자_일괄()
For c = Columns("B").Column To Columns("AY").Column
Sheets("의뢰입력").Cells(9, c) = Sheets("의뢰입력").Cells(9, "B")
Sheets("의뢰입력").Cells(10, c) = Sheets("의뢰입력").Cells(10, "B")
Next c
End Sub
Sub 정도보증_일괄()
For c = Columns("B").Column To Columns("AY").Column
Sheets("의뢰입력").Cells(12, c) = Sheets("의뢰입력").Cells(12, "B")
Next c
End Sub
Sub 분석완료_일괄()
For c = Columns("B").Column To Columns("AY").Column
Sheets("의뢰입력").Cells(13, c) = Sheets("의뢰입력").Cells(13, "B")
Next c
End Sub
Sub 견적구분_일괄()
For c = Columns("B").Column To Columns("AY").Column
Sheets("의뢰입력").Cells(14, c) = Sheets("의뢰입력").Cells(14, "B")
Next c
End Sub
Sub 의뢰항목_일괄()

For c = Columns("B").Column To Columns("AY").Column
    For r = 15 To 75
        Sheets("의뢰입력").Cells(r, c) = Sheets("의뢰입력").Cells(r, "B")
    Next r
Next c

End Sub

Sub ClearX()
Sheets("의뢰입력").Range("B2:AY5,B7:AY100") = ""
End Sub
Sub 의뢰입력_진행()



For c = Columns("B").Column To Columns("AY").Column
X = Sheets("의뢰정보").Cells(Sheets("의뢰정보").Rows.Count, "A").End(xlUp).row + 1
 For r = 2 To 75
    Sheets("의뢰정보").Cells(X, r - 1) = Sheets("의뢰입력").Cells(r, c)
 Next r

Next c

End Sub
