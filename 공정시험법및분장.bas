Attribute VB_Name = "׺"
Sub ()
    Dim X As Integer
    Dim XT As Range
    Dim DT As Long
    Dim DR As Long
    Dim FoundCell As Range
    
    For X = 1 To UserForm1.ListView3.ListItems.Count
        ' "DB" Ʈ ش ؽƮ ãϴ.
        Set XT = Sheets("DB").Columns(3).Find(what:=UserForm1.ListView3.ListItems(X).text, lookat:=xlWhole)
        
        ' ã 쿡 ۾ մϴ.
        If Not XT Is Nothing Then
            ' ش   ListView Ҵմϴ.
            UserForm1.ListView3.ListItems(X).ListSubItems(2).text = Sheets("DB").Cells(XT.row, "E").Value                       'м
            UserForm1.ListView3.ListItems(X).ListSubItems(3).text = Sheets("DB").Cells(XT.row, "G").Value                       'м
            UserForm1.ListView3.ListItems(X).ListSubItems(4).text = "-"                                                             '  ʿ
            UserForm1.ListView3.ListItems(X).ListSubItems(5).text = Format(Sheets("DB").Cells(XT.row, "I").Value, "0000")       'Method NO
            UserForm1.ListView3.ListItems(X).ListSubItems(6).text = Sheets("DB").Cells(XT.row, "H").Value                       'м NO
            
            ' "" Ʈ ¥ ãϴ.
            Set FoundCell = Sheets("").Columns(1).Find(what:=CDate(UserForm1.ListView1.ListItems(1)), lookat:=xlWhole)
            If Not FoundCell Is Nothing Then
                DT = FoundCell.row
                ' "" Ʈ ش ؽƮ ãϴ.
                DR = Sheets("").Rows(1).Find(what:=UserForm1.ListView3.ListItems(X).text, lookat:=xlWhole).Column
                UserForm1.ListView3.ListItems(X).ListSubItems(7).text = Sheets("").Cells(DT, DR).text                           'м NO
            Else
                ' DT ã  쿡  ó
                UserForm1.ListView3.ListItems(X).ListSubItems(7).text = "Not Found"
            End If
        Else
            ' XT ã  쿡  ó
            ' MsgBox "ã   ׸: " & ListView3.ListItems(X).Text
        End If
    Next X
End Sub
