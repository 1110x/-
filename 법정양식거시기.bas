Attribute VB_Name = "İŽñ"
Sub ()
On Error Resume Next

If ActiveSheet.Name = "Ϻ" Then
    SHN = "Ϻ"
    '=-=-=-=-==--=-=-=-=-=-=-=
    X = UserForm1.ListView1.ListItems(1).ListSubItems(2)
    xR = Sheets("").Columns("H").Find(what:=X, lookat:=xlWhole).row
    
    Sheets(SHN).Cells(2, "D") = Sheets("").Cells(xR, "B") 'ȣ
    Sheets(SHN).Cells(2, "I") = Sheets("").Cells(xR, "E") 'ü
    
    Sheets(SHN).Cells(3, "D") = Sheets("").Cells(xR, "C") '
    Sheets(SHN).Cells(3, "I") = Sheets("").Cells(xR, "F") '
    
    Sheets(SHN).Cells(4, "D") = Sheets("").Cells(xR, "D") 'ǥ
    Sheets(SHN).Cells(4, "I") = Sheets("").Cells(xR, "G") 'ǰ
    
    Sheets(SHN).Cells(5, "D") = UserForm1.ListView1.ListItems(1).ListSubItems(4) 'ȯ=ȸ
    Sheets(SHN).Cells(6, "D") = " Ǵ "
    Sheets(SHN).Cells(7, "D") = UserForm1.ListView1.ListItems(1).ListSubItems(3)
    Sheets(SHN).Cells(8, "D") = UserForm1.ListView3.ListItems(1).text & " " & ListView3.ListItems.Count - 1 & "" & "(Ʒ м  ׸ )"
    Sheets(SHN).Cells(9, "D") = "P:4L G:4L"
    '======================================================= ̿  ִ Ȯ
    Dim itemExists As Boolean
    itemExists = False
    Dim index As Long
    Dim item As ListItem
    For Each item In UserForm1.ListView3.ListItems
        index = index + 1
        If item.text = "̿³(pH)" Then
            itemExists = True
            Exit For
        End If
    Next item
    
    If itemExists Then
       Sheets(SHN).Cells(10, "D") = "׸ : pH" & UserForm1.ListView3.ListItems(index).ListSubItems(1).text
    Else
       Sheets(SHN).Cells(10, "D") = ""
    End If
    '======================================================= ̿  ִ Ȯ
    Sheets(SHN).Cells(11, "D") = UserForm1.ListView1.ListItems(1).text
    
    If UserForm1.ListView2.ListItems(1).text <> "" Then
    Sheets(SHN).Cells(11, "I") = UserForm1.ListView2.ListItems(1).text & ", " & UserForm1.ListView2.ListItems(1).ListSubItems(1).text
    Else
    Sheets(SHN).Cells(11, "I") = ""
    
    
    
    
    End If
    
    Sheets(SHN).Range("B13:J72") = ""
    
    For Each Data In UserForm1.ListView3.ListItems
    r = r + 1
    
    Žñ = Sheets("DB").Columns("s").Find(what:=UserForm1.ListView3.ListItems(r).text, lookat:=xlWhole).row
    X = Sheets("DB").Cells(Žñ, "T")



    Sheets(SHN).Cells(r + 12, "B") = Data
    Sheets(SHN).Cells(r + 12, "D") = UserForm1.ListView3.ListItems(r).ListSubItems(4)
    
    If Not UserForm1.ListView3.ListItems(r).ListSubItems(1) = "Ұ" Then

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
    Sheets(SHN).Cells(77, "A") = Format(CDate(UserForm1.ListView2.ListItems(1).ListSubItems(4)), "YYYY MM DD")
    
    If UserForm1.ListView3.ListItems.Count >= 23 Then
    Sheets(SHN).Rows("35:72").Hidden = False
    Else
    Sheets(SHN).Rows("35:72").Hidden = True
    
    End If
'=-=-=-=-==--=-=-=-=-=-=-=
End If

End Sub

