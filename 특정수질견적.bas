Attribute VB_Name = "특정수질견적"
Sub 특정수질견적일괄입력()
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
    
    ' UserForm1의 ListBox1 및 ComboBox1 참조
    Set listBox = UserForm1.ListBox1
    Set comboBox = UserForm1.ComboBox1
    
    ' "견적단가" 시트 참조
    Set ws = ThisWorkbook.Sheets("견적단가")
    
    ' ComboBox에서 선택한 항목에 해당하는 컬럼 찾기
    Dim selectedColumn As Integer
    selectedColumn = 0
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(1, i).Value = comboBox.Value Then
            selectedColumn = i
            Exit For
        End If
    Next i
    
    If selectedColumn = 0 Then
        MsgBox "ComboBox에서 선택한 항목에 해당하는 컬럼을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' D2:D33 범위의 각 셀을 반복
    For Each cell In ws.Range("D2:D33")
        item = cell.Value
        
        ' ListBox1에서 중복 확인
        IsDuplicate = False
        For i = 0 To listBox.ListCount - 1
            If listBox.List(i, 1) = item Then
                IsDuplicate = True
                Exit For
            End If
        Next i
        
        ' 중복이 아니면 ListBox1에 추가
        If Not IsDuplicate And item <> "" Then
            ' 부모 노드 및 금액을 가져오는 예제
            parentNode = ws.Cells(cell.row, 3).Value ' 부모 노드
            amount = ws.Cells(cell.row, selectedColumn).Value ' 견적 단가
            
            ' ListBox1에 항목 추가
            listBox.AddItem
            listBox.List(listBox.ListCount - 1, 0) = parentNode
            listBox.List(listBox.ListCount - 1, 1) = item
            listBox.List(listBox.ListCount - 1, 2) = UserForm1.TextBox3.Value
            listBox.List(listBox.ListCount - 1, 3) = Format(amount, "#,###")
            listBox.List(listBox.ListCount - 1, 4) = Format(amount * UserForm1.TextBox3.Value, "#,###")
        End If
    Next cell
End Sub
