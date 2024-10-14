Attribute VB_Name = "мԷ_"
Sub мԷ¸Ʈǹ()

    Dim targetDate As Date
    Dim targetObj As String
    Dim ws As Worksheet
    Dim FoundCell As Range
    Dim currentCell As Range
    Dim 
    Dim FullText As String
    Dim ExtractedText As String
    Dim BracketPosition As Integer

   Application.ScreenUpdating = True

   


    ' ϴ ۾  Ʈ 
    Set ws = ThisWorkbook.Sheets("Ƿ") ' Ʈ ̸ ڽ Ʈ ̸ ° 

    ' Find ޼带 Ͽ ġϴ  ã
    Set FoundCell = ws.Columns(1).Find(what:=targetDate, LookIn:=xlValues, lookat:=xlWhole)
    Set RealSampList = UserForm1.ListView4.ListItems(UserForm1.ListView4.ListItems.Count).ListSubItems(2)
    BracketPosition = InStr(RealSampList, "")
    
    If BracketPosition > 0 Then
       ExtractedText = Mid(RealSampList, BracketPosition + 1)
    Else
       ExtractedText = "ش ڰ ϴ."
    End If

    
    
    Set ws = ThisWorkbook.Sheets("Ƿ") ' Ʈ ̸ ڽ Ʈ ̸ ° 

    searchValue = UserForm1.ListView4.ListItems(UserForm1.ListView4.ListItems.Count).ListSubItems(1).text
    Set FoundCell = ws.Columns(1).Find(what:=CDate(searchValue), LookIn:=xlValues, lookat:=xlWhole)



    ' ã  ְ, ÷ ġϸ ش   ȣ 
    Do While Not FoundCell Is Nothing

        If FoundCell.Offset(0, 5).Value = ExtractedText Then
            
           UserForm1.ListView4.ListItems(UserForm1.ListView4.ListItems.Count).ListSubItems(3).text = FoundCell.Offset(0, 9)
            Exit Do ' ġϴ  ãǷ ݺ 
        End If

        ' ġ   ġϴ  ã    ˻
        Set FoundCell = ws.Columns(1).FindNext(FoundCell)
        
        
        
    Loop
    
End Sub
Sub ۾() '÷ Ʈ ̵
Application.ScreenUpdating = False
Dim TX As Worksheet
Set TX = Sheets("м Է")
TX.Activate
TX.Range("A3:BM100") = ""
If UserForm1.ListView4.ListItems.Count > 0 Then

For r = 1 To UserForm1.ListView4.ListItems.Count
 TX.Cells(r + 2, 1) = UserForm1.ListView4.ListItems(r)
 TX.Cells(r + 2, 2) = UserForm1.ListView4.ListItems(r).ListSubItems(1)
 TX.Cells(r + 2, 3) = UserForm1.ListView4.ListItems(r).ListSubItems(2)
 TX.Cells(r + 2, 4) = UserForm1.ListView4.ListItems(r).ListSubItems(3)
Next r

End If
Application.ScreenUpdating = True
End Sub
Sub ۾2() 'м׸
Application.ScreenUpdating = False

Dim TX As Worksheet
Set TX = Sheets("м Է")
TX.Activate
TX.Range("E2:BM2") = " "

c = 5

For N = 1 To UserForm1.TreeView7.Nodes.Count
 
 If Not UserForm1.TreeView7.Nodes(N).Parent Is Nothing And UserForm1.TreeView7.Nodes(N).ForeColor = RGB(255, 123, 0) Then
 TX.Cells(2, c) = UserForm1.TreeView7.Nodes(N).text
 c = c + 1
 
 End If
 
 
Next N


Application.ScreenUpdating = True
End Sub
Sub мڷԷ()
    Dim TX As Worksheet
    Dim r As Long
    Dim text As String
    Dim startPos As Long
    Dim endPos As Long
    Dim textToExtract As String
    Dim lengthTextToExtract As Long
    
    '  ũƮ 
    Set TX = Sheets("м Է")
    
    ' Ͱ ִ   ݺ (100  ͱ)
    For r = 2 To TX.Cells(100, "C").End(xlUp).row
        text = TX.Cells(r, "C").Value
        
        ' ȣ ۰  ġ ã
        startPos = InStr(text, "")
        endPos = InStr(text, "")
        
        If startPos > 0 And endPos > startPos Then
            textToExtract = Mid(text, endPos + 1)
            
            '   ۼ
''''            TX.Cells(r, "D").Value = Trim(textToExtract) '  D Է

                 

            
            Debug.Print Trim(textToExtract)
        Else
            ' ȣ  
'''            TX.Cells(r, "D").Value = "ȣ "
        End If
    Next r
End Sub

Function ExtractDesiredText(text As String) As String
    ' ϴ ؽƮ ϴ  ߰մϴ.
    '  , Ư ڿ Ե ؽƮ ȯϵ   ֽϴ.
    
    '  ÷ ü ؽƮ ״ ȯϰ ֽϴ.
    ExtractDesiredText = text
    
    ' ߰  ʿ信  ۼϼ.
End Function

' ListView4 ߺ ׸ ˻
Function IsInListView(ByVal NodeText As String, ByVal parentNodeText As String) As Boolean
    Dim i As Integer
    Dim ListItem As ListItem
    
    IsInListView = False
    For i = 1 To UserForm1.ListView4.ListItems.Count
        Set ListItem = UserForm1.ListView4.ListItems(i)
        If ListItem.SubItems(1) = NodeText And ListItem.SubItems(2) = parentNodeText Then
            IsInListView = True
            Exit Function
        End If
    Next i
End Function

Sub мXԷϱ()
    Dim targetDate As Date
    Dim targetObj As String
    Dim FoundCell As Range
    Dim currentCell As Range
    Dim 
    Dim RX As Worksheet
    Dim TX As Worksheet
    Dim XX As Range
    
    
    Set RX = Sheets("м Է")
    Set TX = ThisWorkbook.Sheets("мڷ") ' Ʈ ̸ ڽ Ʈ ̸ ° 
    
   Application.ScreenUpdating = False

   For r = 3 To Sheets("м Է").Cells(2, 3).End(xlDown).row
   

    ' ڷκ ¥ ÷ 
    targetDate = Format(Sheets("м Է").Cells(r, "B"), "YYYY-MM-DD")
    targetObj = Right(RX.Cells(r, "C"), Len(RX.Cells(r, "C")) - InStr(RX.Cells(r, "C"), ""))
       

    ' ϴ ۾  Ʈ 


    ' Find ޼带 Ͽ ġϴ  ã
    Set FoundCell = TX.Columns(1).Find(what:=targetDate, LookIn:=xlValues, lookat:=xlWhole)


    ' ã  ְ, ÷ ġϸ ش   ȣ 
    Do While Not FoundCell Is Nothing

        If FoundCell.Offset(0, 1).Value = targetObj Then
            X = FoundCell.row

            For H = Range("E1").Column To Range("BM1").Column
              If Not IsNumeric(RX.Cells(2, H).Value) Then
                Set XX = TX.Rows(1).Find(what:=RX.Cells(2, H).Value, lookat:=xlWhole)
                 
                 If XX <> "" Then
                   A = MsgBox(XX.Value & " Էµ  [" & TX.Cells(X, XX.Column) & "]  [" & RX.Cells(r, H) & "]  Ͻðڽϱ", vbYesNo, RX.Cells(r, "C") & "ڷ Ȯ ")
                    If A = vbYes Then
                    TX.Cells(X, XX.Column) = RX.Cells(r, H)
                    End If
                 Else
                    TX.Cells(X, XX.Column) = RX.Cells(r, H)
                 End If
                 
                 
              Else

              End If
              
            Next H

            Exit Do ' ġϴ  ãǷ ݺ 
        End If

        ' ġ   ġϴ  ã    ˻
        Set FoundCell = TX.Columns(1).FindNext(FoundCell)
    Loop

    '   Ȯ ġϴ  ã   ޽ 
    If FoundCell Is Nothing Then
        Application.ScreenUpdating = True
        Debug.Print "ġϴ ¥ ã  ų ÷ ġ ʽϴ."
    End If
    

  Next r
  
        Application.ScreenUpdating = True
End Sub
