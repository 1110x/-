Attribute VB_Name = "Ư"
Sub ƯϰԷ()
    Dim ws As Worksheet
    Dim cell As Range
    Dim item As String
    Dim IsDuplicate As Boolean
    Dim parentNode As String
    Dim amount As Variant
    Dim listBox As MSForms.listBox
    Dim comboBox As MSForms.comboBox
    Dim i As Long
    Dim lastRow As Long
    
    ' UserForm1 ListBox1  ComboBox1 
    Set listBox = UserForm1.ListBox1
    Set comboBox = UserForm1.ComboBox1
    
    ' "ܰ" Ʈ 
    Set ws = ThisWorkbook.Sheets("ܰ")
    
    ' ComboBox  ׸ شϴ ÷ ã
    Dim selectedColumn As Integer
    selectedColumn = 0
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(1, i).Value = comboBox.Value Then
            selectedColumn = i
            Exit For
        End If
    Next i
    
    If selectedColumn = 0 Then
        MsgBox "ComboBox  ׸ شϴ ÷ ã  ϴ.", vbExclamation
        Exit Sub
    End If
    
    ' D2:D33    ݺ
    For Each cell In ws.Range("D2:D33")
        item = cell.Value
        
        ' ListBox1 ߺ Ȯ
        IsDuplicate = False
        For i = 0 To listBox.ListCount - 1
            If listBox.List(i, 1) = item Then
                IsDuplicate = True
                Exit For
            End If
        Next i
        
        ' ߺ ƴϸ ListBox1 ߰
        If Not IsDuplicate And item <> "" Then
            ' θ   ݾ  
            parentNode = ws.Cells(cell.row, 3).Value ' θ 
            amount = ws.Cells(cell.row, selectedColumn).Value '  ܰ
            
            ' ListBox1 ׸ ߰
            listBox.AddItem
            listBox.List(listBox.ListCount - 1, 0) = parentNode
            listBox.List(listBox.ListCount - 1, 1) = item
            listBox.List(listBox.ListCount - 1, 2) = UserForm1.TextBox3.Value
            listBox.List(listBox.ListCount - 1, 3) = Format(amount, "#,###")
            listBox.List(listBox.ListCount - 1, 4) = Format(amount * UserForm1.TextBox3.Value, "#,###")
        End If
    Next cell
End Sub
