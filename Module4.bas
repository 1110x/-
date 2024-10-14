Attribute VB_Name = "Module4"
癤풞ttribute VB_Name = "Module4"
Dim c As Integer

Sub 퓐_構()
For c = Columns("B").Column To Columns("AY").Column
Sheets("퓐韜").Cells(2, c) = Sheets("퓐韜").Cells(2, "B")
Next c
End Sub
Sub 채_構()
For c = Columns("B").Column To Columns("AY").Column
Sheets("퓐韜").Cells(3, c) = Sheets("퓐韜").Cells(3, "B")
Next c
End Sub
Sub 채챨_構()
    Dim ST As Date
    Dim AT As Date
'    Dim c As Integer

    ST = Sheets("퓐韜").Cells(4, "B").Value
    For c = Columns("C").Column To Columns("AY").Column
        AT = AT + DateAdd("n", 10, TimeValue("00:00"))
        Sheets("퓐韜").Cells(4, c).Value = ST + AT
    Next c
End Sub
Sub 퓐迷_構()
For c = Columns("B").Column To Columns("AY").Column
 Sheets("퓐韜").Cells(5, c) = Sheets("퓐韜").Cells(5, "B")
Next c
End Sub

Sub 회_構()

For c = Columns("B").Column To Columns("AY").Column
Sheets("퓐韜").Cells(8, c) = Sheets("퓐韜").Cells(8, "B")
Next c

End Sub
Sub 첨채_構()
For c = Columns("B").Column To Columns("AY").Column
Sheets("퓐韜").Cells(9, c) = Sheets("퓐韜").Cells(9, "B")
Sheets("퓐韜").Cells(10, c) = Sheets("퓐韜").Cells(10, "B")
Next c
End Sub
Sub _構()
For c = Columns("B").Column To Columns("AY").Column
Sheets("퓐韜").Cells(12, c) = Sheets("퓐韜").Cells(12, "B")
Next c
End Sub
Sub 劇狗_構()
For c = Columns("B").Column To Columns("AY").Column
Sheets("퓐韜").Cells(13, c) = Sheets("퓐韜").Cells(13, "B")
Next c
End Sub
Sub _構()
For c = Columns("B").Column To Columns("AY").Column
Sheets("퓐韜").Cells(14, c) = Sheets("퓐韜").Cells(14, "B")
Next c
End Sub
Sub 퓐琉_構()

For c = Columns("B").Column To Columns("AY").Column
    For r = 15 To 75
        Sheets("퓐韜").Cells(r, c) = Sheets("퓐韜").Cells(r, "B")
    Next r
Next c

End Sub

Sub ClearX()
Sheets("퓐韜").Range("B2:AY5,B7:AY100") = ""
End Sub
Sub 퓐韜_()



For c = Columns("B").Column To Columns("AY").Column
X = Sheets("퓐").Cells(Sheets("퓐").Rows.Count, "A").End(xlUp).row + 1
 For r = 2 To 75
    Sheets("퓐").Cells(X, r - 1) = Sheets("퓐韜").Cells(r, c)
 Next r

Next c

End Sub
