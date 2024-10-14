Attribute VB_Name = "방류기준거시기"
Sub 방류기준찾기()
    Dim X As Integer
    Dim XT As Range
    Dim T As Range, TR As Range
    Dim TZ As Long
    Dim 기준 As String
    
    ' ListView2에서 기준 값을 가져옴
    기준 = UserForm1.ListView2.ListItems(1).ListSubItems(2).text
    
    ' 방류기준표 시트의 2행에서 기준 값을 찾음
    Set T = Sheets("방류기준표").Rows(2).Find(what:=기준, lookat:=xlWhole)
    
    ' 기준 값을 찾은 경우에만 작업 수행
    If Not T Is Nothing Then
        For r = 1 To UserForm1.ListView3.ListItems.Count
            ' 방류기준표 시트의 1열에서 ListView3 항목을 찾음
            Set TR = Sheets("방류기준표").Columns(1).Find(what:=UserForm1.ListView3.ListItems(r).text, lookat:=xlWhole)
         
            ' 항목을 찾은 경우에만 작업 수행
            If Not TR Is Nothing Then
                UserForm1.ListView3.ListItems(r).ListSubItems(4).text = Sheets("방류기준표").Cells(TR.row, T.Column).Value
            End If
        Next r
    End If
    
    ' "확인란 5" 체크박스가 선택된 경우
    If Sheets("시험성적서").CheckBoxes("확인란 5").Value = 1 Then
        ' H10:H41 범위에 대한 작업 수행
        For Each XT In Sheets("시험성적서").Range("H10:H41").Cells
            If Sheets("시험성적서").Cells(XT.row, "D").Value <> "" Then
                Set TR = Sheets("방류기준표").Columns(1).Find(what:=Sheets("시험성적서").Cells(XT.row, "D").Value, lookat:=xlWhole)
                If Not TR Is Nothing Then
                    XT.Value = Sheets("방류기준표").Cells(TR.row, T.Column).Value
                End If
            End If
        Next XT
        
        ' P10:P41 범위에 대한 작업 수행
        For Each XT In Sheets("시험성적서").Range("P10:P41").Cells
            If Sheets("시험성적서").Cells(XT.row, "L").Value <> "" Then
                Set TR = Sheets("방류기준표").Columns(1).Find(what:=Sheets("시험성적서").Cells(XT.row, "L").Value, lookat:=xlWhole)
                If Not TR Is Nothing Then
                    XT.Value = Sheets("방류기준표").Cells(TR.row, T.Column).Value
                End If
            End If
        Next XT
        
    Else
        For Each XT In Sheets("시험성적서").Range("H10:H41").Cells
            If Sheets("시험성적서").Cells(XT.row, "D").Value <> "" Then

                    XT.Value = ""
            End If
        Next XT
        
        ' P10:P41 범위에 대한 작업 수행
        For Each XT In Sheets("시험성적서").Range("P10:P41").Cells
            If Sheets("시험성적서").Cells(XT.row, "L").Value <> "" Then

                    XT.Value = ""
            End If
        Next XT
    End If
End Sub


