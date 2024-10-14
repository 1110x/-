
Attribute VB_Name = "Combo"
Sub Combo1()
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim comboBox As MSForms.comboBox
    Dim lastColumn As Long
    Dim firstRow As Long
    'test 00::
    ' "견적단가" 시트 참조
    Set ws = ThisWorkbook.Sheets("견적단가")
    
    ' E 열의 마지막 열을 찾기 위해 첫 번째 행의 마지막 열을 찾음
    firstRow = 1
    lastColumn = ws.Cells(firstRow, ws.Columns.Count).End(xlToLeft).Column
    
    ' E1부터 마지막 열까지의 범위 참조
    Set rng = ws.Range(ws.Cells(firstRow, 5), ws.Cells(firstRow, lastColumn))
    
    ' UserForm1의 ComboBox1 참조
    Set comboBox = UserForm1.ComboBox1
    
    ' ComboBox 초기화 (기존 아이템 제거)
    comboBox.Clear
    
    ' 범위의 각 셀 값을 ComboBox에 추가
    For Each cell In rng
        If cell.Value <> "" Then
            comboBox.AddItem cell.Value
        End If
    Next cell
    
    ' 첫 번째 항목을 선택
    If comboBox.ListCount > 0 Then
        comboBox.ListIndex = 0
    End If
End Sub
Sub Combob2()
'Real Combobox2 값 설정임 아래것은 ...
    Dim ws As Worksheet

    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim selectedValue As String
    
   Set ws = ThisWorkbook.Sheets("업체담당자")
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
    
    ' "견적단가" 시트 참조
    Set ws = ThisWorkbook.Sheets("계약정보")
    
    ' E 열의 마지막 열을 찾기 위해 첫 번째 행의 마지막 열을 찾음
    firstRow = 2
    lastRow = ws.Cells(firstRow, 2).End(xlDown).row

    
    ' J1부터 마지막 열까지의 범위 참조
    Set rng = ws.Range(ws.Cells(firstRow, 2), ws.Cells(lastRow, 2))
    '================================================================= COMBOBOX3
    ' UserForm1의 ComboBox1 참조
    Set comboBox = UserForm1.ComboBox3
    
    ' ComboBox 초기화 (기존 아이템 제거)
    comboBox.Clear
    
    ' 범위의 각 셀 값을 ComboBox에 추가
    For Each cell In rng
        If cell.Value <> "" Then
            comboBox.AddItem cell.Value
        End If
    Next cell
    
    ' 첫 번째 항목을 선택
    If comboBox.ListCount > 0 Then
        comboBox.ListIndex = 0
    End If
   '================================================================= COMBOBOX5
        ' UserForm1의 ComboBox1 참조
    Set comboBox = UserForm1.ComboBox5
    
    ' ComboBox 초기화 (기존 아이템 제거)
    comboBox.Clear
    
    ' 범위의 각 셀 값을 ComboBox에 추가
    For Each cell In rng
        If cell.Value <> "" Then
            comboBox.AddItem cell.Value
        End If
    Next cell
    
    ' 첫 번째 항목을 선택
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
        Set ws = ThisWorkbook.Sheets("담당자정보")
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
        
           comboBox.AddItem "전체기간 -_-'"
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

    ' Scripting.Dictionary 생성
    Set dict = CreateObject("Scripting.Dictionary")

    ' 워크시트 지정
    Set ws = ThisWorkbook.Sheets("측정DB") ' 분장시트라는 이름의 시트로 지정

    ' B2에서 BH 열의 마지막 행까지의 범위 선택
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).row

    ' 시작 날짜 설정
    startDate = DateSerial(2024, 1, 1)

    ' 날짜 조건에 맞는 데이터만 ComboBox7에 추가
    For row = 2 To lastRow
        ' A열의 날짜가 2024년 1월 1일 이후인지 확인
'        If ws.Cells(row, 1).Value >= startDate Then
            ' 해당 행의 B열부터 BH열까지의 데이터를 순회
            For Each cell In ws.Range(ws.Cells(row, "N"), ws.Cells(row, "N"))
                If cell.Value <> "" And Not dict.exists(cell.Value) Then
                    dict.Add cell.Value, Nothing
                    UserForm1.ComboBox7.AddItem cell.Value
                End If
            Next cell
'        End If
    Next row

    ' ComboBox7의 첫 번째 아이템을 기본 선택
    If UserForm1.ComboBox7.ListCount > 0 Then
        UserForm1.ComboBox7.ListIndex = 1
    End If
    
End Sub
Sub 분장항목Combo()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim X As Long
    Dim 항목 As Variant
    Dim cell As Range
    Dim TT As String
    Dim r As Long

    ' 워크시트 지정
    Set ws = ThisWorkbook.Sheets("분장") ' 분장 시트로 지정
    
    ' 트리뷰 초기화
    For r = 1 To UserForm1.TreeView7.Nodes.Count
        UserForm1.TreeView7.Nodes(r).ForeColor = RGB(0, 0, 0)
        If Not UserForm1.TreeView7.Nodes(r).Parent Is Nothing Then
            UserForm1.TreeView7.Nodes(r).Parent.Expanded = False
        End If
    Next r
    
    ' 현재 날짜와 일치하는 행 찾기
    ' X = ws.Columns(1).Find(what:=CDate(Format(Now(), "YYYY-MM-DD")), lookat:=xlWhole).Row
    X = ws.Columns(1).Find(what:=CDate("2024-07-22"), lookat:=xlWhole).row '테스트를 위해 임시로 사용

    ' 지정된 행의 B열부터 마지막 열까지의 값들을 반복
    For Each cell In ws.Range(ws.Cells(X, 2), ws.Cells(X, ws.Cells(X, ws.Columns.Count).End(xlToLeft).Column))
        항목 = cell.Value
        ' 조건에 따른 디버그 출력
        If 항목 = UserForm1.ComboBox7.Value Then
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

