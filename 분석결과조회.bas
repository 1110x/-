Attribute VB_Name = "мȸ"
Sub мҷ()
    Dim targetDate As Date
    Dim targetObj As String
    Dim ws As Worksheet
    Dim FoundCell As Range
    Dim currentCell As Range
    Dim 
    

   Application.ScreenUpdating = False

   

    ' ڷκ ¥ ÷ 
    targetDate = DateValue(UserForm1.ListView1.ListItems(1).ListSubItems(1).text)
    targetObj = UserForm1.ListView1.ListItems(1).ListSubItems(3).text

    ' ϴ ۾  Ʈ 
    Set ws = ThisWorkbook.Sheets("мڷ") ' Ʈ ̸ ڽ Ʈ ̸ ° 

    ' Find ޼带 Ͽ ġϴ  ã
    Set FoundCell = ws.Columns(1).Find(what:=targetDate, LookIn:=xlValues, lookat:=xlWhole)


    ' ã  ְ, ÷ ġϸ ش   ȣ 
    Do While Not FoundCell Is Nothing

        If FoundCell.Offset(0, 1).Value = targetObj Then
            X = FoundCell.row
            Sheets("輺").Cells(1, "C") = "RENEWUS-WAC-" & Format(FoundCell.Offset(0, 0).Value, "YYYY") & "-" & FoundCell.row & "-A"
            Sheets("輺").Cells(1, "K") = "RENEWUS-WAC-" & Format(FoundCell.Offset(0, 0).Value, "YYYY") & "-" & FoundCell.row & "-B"
            For c = 1 To UserForm1.ListView3.ListItems.Count
              TC = Sheets("мڷ").Rows(1).Find(what:=UserForm1.ListView3.ListItems(c).text, lookat:=1).Column
              UserForm1.ListView3.ListItems(c).ListSubItems(1).text = Sheets("мڷ").Cells(X, TC).Value
               = Sheets("DB").Columns(3).Find(what:=UserForm1.ListView3.ListItems(c).text, lookat:=xlWhole).row
              If c <= 32 Then
                  If Sheets("мڷ").Cells(X, TC).Value = "" Then
                  Sheets("輺").Cells(9 + c, "F") = "м"
                  Sheets("輺").Cells(9 + c, "G") = ""

                  ElseIf Sheets("мڷ").Cells(X, TC).Value = "Ұ" Then
                  Sheets("輺").Cells(9 + c, "F") = "м"
                  Sheets("輺").Cells(9 + c, "G") = ""
                  Else
                  Sheets("輺").Cells(9 + c, "F") = Sheets("мڷ").Cells(X, TC).Value
                  Sheets("輺").Cells(9 + c, "F").NumberFormatLocal = Sheets("DB").Cells(, "A")
                  Sheets("輺").Cells(9 + c, "G") = Sheets("DB").Cells(, "B").Value
                  End If
              Else
                  If Sheets("мڷ").Cells(X, TC).Value = "" Then
                  Sheets("輺").Cells(9 + c - 32, "N") = "м"
                  Sheets("輺").Cells(9 + c - 32, "O") = ""
                  Else
                  Sheets("輺").Cells(9 + c - 32, "N") = Sheets("мڷ").Cells(X, TC).Value
                  Sheets("輺").Cells(9 + c - 32, "N").NumberFormatLocal = Sheets("DB").Cells(, "A")
                  Sheets("輺").Cells(9 + c - 32, "O") = Sheets("DB").Cells(, "B").Value
                  End If
                  
              End If
              
            Next c

            Exit Do ' ġϴ  ãǷ ݺ 
        End If

        ' ġ   ġϴ  ã    ˻
        Set FoundCell = ws.Columns(1).FindNext(FoundCell)
    Loop

    '   Ȯ ġϴ  ã   ޽ 
    If FoundCell Is Nothing Then
        Application.ScreenUpdating = True
        Exit Sub
        Debug.Print "ġϴ ¥ ã  ų ÷ ġ ʽϴ."
    End If
    
    If Sheets("輺").CheckBoxes("Ȯζ 5").Value = 1 Then
     ã
    End If
    
        Application.ScreenUpdating = True
End Sub
