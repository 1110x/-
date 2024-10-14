Attribute VB_Name = "Combo"
Sub Combo1()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim comboBox As MSForms.comboBox
    Dim lastColumn As Long
    Dim firstRow As Long
    
    ' "ܰ" Ʈ 
    Set ws = ThisWorkbook.Sheets("ܰ")
    
    ' E    ã  ù °    ã
    firstRow = 1
    lastColumn = ws.Cells(firstRow, ws.Columns.Count).End(xlToLeft).Column
    
    ' E1    
    Set rng = ws.Range(ws.Cells(firstRow, 5), ws.Cells(firstRow, lastColumn))
    
    ' UserForm1 ComboBox1 
    Set comboBox = UserForm1.ComboBox1
    
    ' ComboBox ʱȭ (  )
    comboBox.Clear
    
    '     ComboBox ߰
    For Each cell In rng
        If cell.Value <> "" Then
            comboBox.AddItem cell.Value
        End If
    Next cell
    
    ' ù ° ׸ 
    If comboBox.ListCount > 0 Then
        comboBox.ListIndex = 0
    End If
End Sub
Sub Combob2()
'Real Combobox2   Ʒ ...
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim selectedValue As String
    
   Set ws = ThisWorkbook.Sheets("ü")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    For r = 3 To lastRow
     
     If UserForm1.ComboBox2.Value = ws.Cells(r, "C") & " " & ws.Cells(r, "E") & " " & ws.Cells(r, "D") Then
      UserForm1.TextBox4 = ws.Cells(r, "G")
      UserForm1.TextBox5 = ws.Cells(r, "F")
      Exit Sub
      
     End If
     
    Next r
End Sub
Sub Combo2()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim comboBox As MSForms.comboBox
    Dim lastColumn As Long
    Dim firstRow As Long
    
    ' "ܰ" Ʈ 
    Set ws = ThisWorkbook.Sheets("")
    
    ' E    ã  ù °    ã
    firstRow = 2
    lastRow = ws.Cells(firstRow, 2).End(xlDown).row

    
    ' J1    
    Set rng = ws.Range(ws.Cells(firstRow, 2), ws.Cells(lastRow, 2))
    '================================================================= COMBOBOX3
    ' UserForm1 ComboBox1 
    Set comboBox = UserForm1.ComboBox3
    
    ' ComboBox ʱȭ (  )
    comboBox.Clear
    
    '     ComboBox ߰
    For Each cell In rng
        If cell.Value <> "" Then
            comboBox.AddItem cell.Value
        End If
    Next cell
    
    ' ù ° ׸ 
    If comboBox.ListCount > 0 Then
        comboBox.ListIndex = 0
    End If
   '================================================================= COMBOBOX5
        ' UserForm1 ComboBox1 
    Set comboBox = UserForm1.ComboBox5
    
    ' ComboBox ʱȭ (  )
    comboBox.Clear
    
    '     ComboBox ߰
    For Each cell In rng
        If cell.Value <> "" Then
            comboBox.AddItem cell.Value
        End If
    Next cell
    
    ' ù ° ׸ 
    If comboBox.ListCount > 0 Then
        comboBox.ListIndex = 0
    End If
    
    
End Sub

Sub combo4()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim comboBox As MSForms.comboBox
    Dim lastColumn As Long
    Dim firstRow As Long
        Set ws = ThisWorkbook.Sheets("")
        Set comboBox = UserForm1.ComboBox4
        comboBox.Clear

        For r = 2 To 20

         If ws.Cells(r, "A") <> "" Then
           comboBox.AddItem ws.Cells(r, "A").Value
         Else
           Exit For
         End If

        Next r



    If comboBox.ListCount > 0 Then
        comboBox.ListIndex = 2
    End If


End Sub

Sub combo6()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim comboBox As MSForms.comboBox
    Dim lastColumn As Long
    
       Set comboBox = UserForm1.ComboBox6
        comboBox.Clear
        
           comboBox.AddItem "üⰣ -_-'"
       For c = Format(Now(), "YYYY") To 2021 Step (-1)
         

           comboBox.AddItem c
         
       Next c
       
       
    If comboBox.ListCount > 0 Then
        comboBox.ListIndex = 1
    End If
                 

End Sub
Sub combo7()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim dict As Object
    Dim startDate As Date
    Dim row As Long

    ' Scripting.Dictionary 
    Set dict = CreateObject("Scripting.Dictionary")

    ' ũƮ 
    Set ws = ThisWorkbook.Sheets("DB") ' Ʈ ̸ Ʈ 

    ' B2 BH     
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).row

    '  ¥ 
    startDate = DateSerial(2024, 1, 1)

    ' ¥ ǿ ´ ͸ ComboBox7 ߰
    For row = 2 To lastRow
        ' A ¥ 2024 1 1  Ȯ
'        If ws.Cells(row, 1).Value >= startDate Then
            ' ش  B BH ͸ ȸ
            For Each cell In ws.Range(ws.Cells(row, "N"), ws.Cells(row, "N"))
                If cell.Value <> "" And Not dict.exists(cell.Value) Then
                    dict.Add cell.Value, Nothing
                    UserForm1.ComboBox7.AddItem cell.Value
                End If
            Next cell
'        End If
    Next row

    ' ComboBox7 ù °  ⺻ 
    If UserForm1.ComboBox7.ListCount > 0 Then
        UserForm1.ComboBox7.ListIndex = 1
    End If
    
End Sub
Sub ׸Combo()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim X As Long
    Dim ׸ As Variant
    Dim cell As Range
    Dim TT As String
    Dim r As Long

    ' ũƮ 
    Set ws = ThisWorkbook.Sheets("") '  Ʈ 
    
    ' Ʈ ʱȭ
    For r = 1 To UserForm1.TreeView7.Nodes.Count
        UserForm1.TreeView7.Nodes(r).ForeColor = RGB(0, 0, 0)
        If Not UserForm1.TreeView7.Nodes(r).Parent Is Nothing Then
            UserForm1.TreeView7.Nodes(r).Parent.Expanded = False
        End If
    Next r
    
    '  ¥ ġϴ  ã
    ' X = ws.Columns(1).Find(what:=CDate(Format(Now(), "YYYY-MM-DD")), lookat:=xlWhole).Row
    X = ws.Columns(1).Find(what:=CDate("2024-07-22"), lookat:=xlWhole).row '׽Ʈ  ӽ÷ 

    '   B    ݺ
    For Each cell In ws.Range(ws.Cells(X, 2), ws.Cells(X, ws.Cells(X, ws.Columns.Count).End(xlToLeft).Column))
        ׸ = cell.Value
        ' ǿ   
        If ׸ = UserForm1.ComboBox7.Value Then
            TT = ws.Cells(1, cell.Column).Value
            
            For r = 1 To UserForm1.TreeView7.Nodes.Count
                If UserForm1.TreeView7.Nodes(r).text = TT Then
                    UserForm1.TreeView7.Nodes(r).ForeColor = RGB(255, 123, 0)
                    If Not UserForm1.TreeView7.Nodes(r).Parent Is Nothing Then
                        UserForm1.TreeView7.Nodes(r).Parent.ForeColor = RGB(255, 0, 0)
                        UserForm1.TreeView7.Nodes(r).Parent.Expanded = True
                    End If
                    Exit For
                End If
            Next r
            
        End If
    Next cell
    
    
End Sub

