Attribute VB_Name = "Module4"
Dim c As Integer

Sub Ƿ_ϰ()
For c = Columns("B").Column To Columns("AY").Column
Sheets("ǷԷ").Cells(2, c) = Sheets("ǷԷ").Cells(2, "B")
Next c
End Sub
Sub ä_ϰ()
For c = Columns("B").Column To Columns("AY").Column
Sheets("ǷԷ").Cells(3, c) = Sheets("ǷԷ").Cells(3, "B")
Next c
End Sub
Sub äð_ϰ()
    Dim ST As Date
    Dim AT As Date
'    Dim c As Integer

    ST = Sheets("ǷԷ").Cells(4, "B").Value
    For c = Columns("C").Column To Columns("AY").Column
        AT = AT + DateAdd("n", 10, TimeValue("00:00"))
        Sheets("ǷԷ").Cells(4, c).Value = ST + AT
    Next c
End Sub
Sub Ƿڻ_ϰ()
For c = Columns("B").Column To Columns("AY").Column
 Sheets("ǷԷ").Cells(5, c) = Sheets("ǷԷ").Cells(5, "B")
Next c
End Sub

Sub ȸ_ϰ()

For c = Columns("B").Column To Columns("AY").Column
Sheets("ǷԷ").Cells(8, c) = Sheets("ǷԷ").Cells(8, "B")
Next c

End Sub
Sub ÷ä_ϰ()
For c = Columns("B").Column To Columns("AY").Column
Sheets("ǷԷ").Cells(9, c) = Sheets("ǷԷ").Cells(9, "B")
Sheets("ǷԷ").Cells(10, c) = Sheets("ǷԷ").Cells(10, "B")
Next c
End Sub
Sub _ϰ()
For c = Columns("B").Column To Columns("AY").Column
Sheets("ǷԷ").Cells(12, c) = Sheets("ǷԷ").Cells(12, "B")
Next c
End Sub
Sub мϷ_ϰ()
For c = Columns("B").Column To Columns("AY").Column
Sheets("ǷԷ").Cells(13, c) = Sheets("ǷԷ").Cells(13, "B")
Next c
End Sub
Sub _ϰ()
For c = Columns("B").Column To Columns("AY").Column
Sheets("ǷԷ").Cells(14, c) = Sheets("ǷԷ").Cells(14, "B")
Next c
End Sub
Sub Ƿ׸_ϰ()

For c = Columns("B").Column To Columns("AY").Column
    For r = 15 To 75
        Sheets("ǷԷ").Cells(r, c) = Sheets("ǷԷ").Cells(r, "B")
    Next r
Next c

End Sub

Sub ClearX()
Sheets("ǷԷ").Range("B2:AY5,B7:AY100") = ""
End Sub
Sub ǷԷ_()



For c = Columns("B").Column To Columns("AY").Column
X = Sheets("Ƿ").Cells(Sheets("Ƿ").Rows.Count, "A").End(xlUp).row + 1
 For r = 2 To 75
    Sheets("Ƿ").Cells(X, r - 1) = Sheets("ǷԷ").Cells(r, c)
 Next r

Next c

End Sub
