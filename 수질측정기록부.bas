Attribute VB_Name = "Ϻ"
Sub Ϻڷ()
 Dim Node As MSComctlLib.Node 'TreeView Ʈ 带 Ÿ ü
    
    
    For X = 1 To UserForm1.TreeView1.Nodes.Count
     If UserForm1.TreeView1.Nodes(X).ForeColor = RGB(255, 0, 0) Then
        s = X
        Exit For
     End If
    Next X

    For X = UserForm1.TreeView1.Nodes.Count To 1 Step (-1)
     If UserForm1.TreeView1.Nodes(X).ForeColor = RGB(255, 0, 0) Then
        E = X
        Exit For
     End If
    Next X
    
    
For xxx = E To s Step (-1)
    Set Node = UserForm1.TreeView1.Nodes(xxx) ' 2 մϴ.
    Node.EnsureVisible '尡 ̵ ũմϴ.
    Node.Selected = True '带 մϴ.


    Dim ws As Worksheet
If Not UserForm1.TreeView1.SelectedItem Is Nothing Then
    If Not UserForm1.TreeView1.SelectedItem.Parent Is Nothing Then
'        Debug.Print "ڽĳ: " & TreeView1.SelectedItem.Parent.text
        Set ws = ThisWorkbook.Sheets("Ƿ")
        lastRow = Sheets("Ƿ").Cells(Sheets("Ƿ").Rows.Count, "A").End(xlUp).row
        For r = 2 To lastRow
         X = "" & Sheets("Ƿ").Cells(r, "E").text & "" & Sheets("Ƿ").Cells(r, "F").Value '  ws.Cells(i, "C")  & ws.Cells(i, 4).Value)
         If Sheets("Ƿ").Cells(r, "A") = UserForm1.TreeView1.SelectedItem.Parent.text And X = UserForm1.TreeView1.SelectedItem.text Then

          '===============================================================================================
         UserForm1.ListView1.ListItems.Clear
         With UserForm1.ListView1
        .ColumnHeaders.Clear '  ÷ 
        .Gridlines = True
        .View = lvwReport ' Report  

        '  ÷ ʺ Ϸ ʿ信  Width Ӽ   ֽϴ.
        .ColumnHeaders.Add , , "Ƿ", 100
        .ColumnHeaders.Add , , "ä", 100
        .ColumnHeaders.Add , , "Ƿڻ", 100
        .ColumnHeaders.Add , , "÷", 120
        .ColumnHeaders.Add , , "ȸ", 100

         End With

         UserForm1.ListView2.ListItems.Clear
         With UserForm1.ListView2
        .ColumnHeaders.Clear '  ÷ 
        .Gridlines = True
        .View = lvwReport ' Report  
        '  ÷ ʺ Ϸ ʿ信  Width Ӽ   ֽϴ.
        .ColumnHeaders.Add , , "÷ä1", 100
        .ColumnHeaders.Add , , "÷ä2", 100
        .ColumnHeaders.Add , , "", 120
        .ColumnHeaders.Add , , "", 100
        .ColumnHeaders.Add , , "м", 100
         End With

         UserForm1.ListView3.ListItems.Clear
         With UserForm1.ListView3
        .ColumnHeaders.Clear '  ÷ 
        .Gridlines = True
        .View = lvwReport ' Report  
        '  ÷ ʺ Ϸ ʿ信  Width Ӽ   ֽϴ.
        .ColumnHeaders.Add , , "Ƿ׸", 130
        .ColumnHeaders.Add , , "м", 70
        .ColumnHeaders.Add , , "м", 170
        .ColumnHeaders.Add , , "м", 100
        .ColumnHeaders.Add , , "", 50
        .ColumnHeaders.Add , , "Method NO", 20
        .ColumnHeaders.Add , , "instrument NO", 20
        .ColumnHeaders.Add , , "м", 20
         End With


         Set item = UserForm1.ListView1.ListItems.Add(1, , Sheets("Ƿ").Cells(r, "A").Value) 'Ƿ
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "B").Value                'ä
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "E").Value                'Ƿڻ 'Ī
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "F").Value                '÷
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "G").Value                'ȸ

         Set item = UserForm1.ListView2.ListItems.Add(1, , Sheets("Ƿ").Cells(r, "H").Value) '÷ä-1
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "I").Value                '÷ä-2
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "J").Value                '
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "K").Value                '
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "L").Value                'м

         For T = Range("N1").Column To Range("BT1").Column
         If Sheets("Ƿ").Cells(r, T) <> "" Then
         Set Titem = UserForm1.ListView3.ListItems.Add(UserForm1.ListView3.ListItems.Count + 1, , Sheets("Ƿ").Cells(1, T))
                     Titem.ListSubItems.Add , , "-"             'Subitem-1 м
                     Titem.ListSubItems.Add , , "-"             'Subitem-2 м
                     Titem.ListSubItems.Add , , "-"             'Subitem-3 м
                     Titem.ListSubItems.Add , , "-"             'Subitem-4 
                     Titem.ListSubItems.Add , , "-"             'Subitem-5 Method NO
                     Titem.ListSubItems.Add , , "-"             'Subitem-6 instrument NO
                     Titem.ListSubItems.Add , , "-"             'Subitem-7 м
         End If
         Next T
         '===============================================================================================

         End If

        Next r

        Call ã
        UserForm1.ListView3.ColumnHeaders(1).text = "Ƿ׸ ( " & UserForm1.ListView3.ListItems.Count & ") "
    Else
        UserForm1.Label1.Caption = "ãνϴ"
    End If

мҷ
ã

UserForm1.TreeView1.Nodes(xxx).ForeColor = RGB(0, 0, 128)
ActiveSheet.PrintOut
End If


Next xxx

End Sub
   
