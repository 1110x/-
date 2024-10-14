Attribute VB_Name = "법정양식거시기"
Sub 법정양식()
On Error Resume Next

If ActiveSheet.Name = "수질측정기록부" Then
    SHN = "수질측정기록부"
    '=-=-=-=-==--=-=-=-=-=-=-=
    X = UserForm1.ListView1.ListItems(1).ListSubItems(2)
    xR = Sheets("계약정보").Columns("H").Find(what:=X, lookat:=xlWhole).row
    
    Sheets(SHN).Cells(2, "D") = Sheets("계약정보").Cells(xR, "B") '상호명
    Sheets(SHN).Cells(2, "I") = Sheets("계약정보").Cells(xR, "E") '시설별
    
    Sheets(SHN).Cells(3, "D") = Sheets("계약정보").Cells(xR, "C") '소재지
    Sheets(SHN).Cells(3, "I") = Sheets("계약정보").Cells(xR, "F") '종류별
    
    Sheets(SHN).Cells(4, "D") = Sheets("계약정보").Cells(xR, "D") '대표자
    Sheets(SHN).Cells(4, "I") = Sheets("계약정보").Cells(xR, "G") '생산품
    
    Sheets(SHN).Cells(5, "D") = UserForm1.ListView1.ListItems(1).ListSubItems(4) '환경기술인=입회자
    Sheets(SHN).Cells(6, "D") = "제출 또는 보고용"
    Sheets(SHN).Cells(7, "D") = UserForm1.ListView1.ListItems(1).ListSubItems(3)
    Sheets(SHN).Cells(8, "D") = UserForm1.ListView3.ListItems(1).text & "외 " & ListView3.ListItems.Count - 1 & "건" & "(아래 ⑤측정분석 결과의 항목과 같음)"
    Sheets(SHN).Cells(9, "D") = "P:4L G:4L"
    '======================================================= 수소이온 농도 있는지 확인
    Dim itemExists As Boolean
    itemExists = False
    Dim index As Long
    Dim item As ListItem
    For Each item In UserForm1.ListView3.ListItems
        index = index + 1
        If item.text = "수소이온농도(pH)" Then
            itemExists = True
            Exit For
        End If
    Next item
    
    If itemExists Then
       Sheets(SHN).Cells(10, "D") = "현장측정항목 : pH" & UserForm1.ListView3.ListItems(index).ListSubItems(1).text
    Else
       Sheets(SHN).Cells(10, "D") = ""
    End If
    '======================================================= 수소이온 농도 있는지 확인
    Sheets(SHN).Cells(11, "D") = UserForm1.ListView1.ListItems(1).text
    
    If UserForm1.ListView2.ListItems(1).text <> "" Then
    Sheets(SHN).Cells(11, "I") = UserForm1.ListView2.ListItems(1).text & ", " & UserForm1.ListView2.ListItems(1).ListSubItems(1).text
    Else
    Sheets(SHN).Cells(11, "I") = ""
    
    
    
    
    End If
    
    Sheets(SHN).Range("B13:J72") = ""
    
    For Each Data In UserForm1.ListView3.ListItems
    r = r + 1
    
    거시기 = Sheets("측정DB").Columns("s").Find(what:=UserForm1.ListView3.ListItems(r).text, lookat:=xlWhole).row
    X = Sheets("측정DB").Cells(거시기, "T")



    Sheets(SHN).Cells(r + 12, "B") = Data
    Sheets(SHN).Cells(r + 12, "D") = UserForm1.ListView3.ListItems(r).ListSubItems(4)
    
    If Not UserForm1.ListView3.ListItems(r).ListSubItems(1) = "불검출" Then

    If X = 0 Then
        Sheets(SHN).Cells(r + 12, "F").NumberFormatLocal = "0"
    ElseIf X = 1 Then
        Sheets(SHN).Cells(r + 12, "F").NumberFormatLocal = "0.0"
    ElseIf X = 2 Then
        Sheets(SHN).Cells(r + 12, "F").NumberFormatLocal = "0.00"
    ElseIf X = 3 Then
        Sheets(SHN).Cells(r + 12, "F").NumberFormatLocal = "0.000"
    ElseIf X = 4 Then
        Sheets(SHN).Cells(r + 12, "F").NumberFormatLocal = "0.0000"
    End If
    

     Sheets(SHN).Cells(r + 12, "F") = Val(UserForm1.ListView3.ListItems(r).ListSubItems(1).text)



    Else
       Sheets(SHN).Cells(r + 12, "F") = UserForm1.ListView3.ListItems(r).ListSubItems(1)
    End If
    
    Sheets(SHN).Cells(r + 12, "H") = UserForm1.ListView3.ListItems(r).ListSubItems(2)
    Next Data
    
    Sheets(SHN).Cells(73, "D") = UserForm1.ListView1.ListItems(1).ListSubItems(1) & " ~ " & UserForm1.ListView2.ListItems(1).ListSubItems(4)
    Sheets(SHN).Cells(77, "A") = Format(CDate(UserForm1.ListView2.ListItems(1).ListSubItems(4)), "YYYY년 MM월 DD일")
    
    If UserForm1.ListView3.ListItems.Count >= 23 Then
    Sheets(SHN).Rows("35:72").Hidden = False
    Else
    Sheets(SHN).Rows("35:72").Hidden = True
    
    End If
'=-=-=-=-==--=-=-=-=-=-=-=
End If

End Sub

