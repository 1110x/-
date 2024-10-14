Attribute VB_Name = "분석결과조회"
Sub 분석결과불러오기()
    Dim targetDate As Date
    Dim targetObj As String
    Dim ws As Worksheet
    Dim FoundCell As Range
    Dim currentCell As Range
    Dim 단위
    

   Application.ScreenUpdating = False

   

    ' 사용자로부터 날짜와 시료명을 얻어옴
    targetDate = DateValue(UserForm1.ListView1.ListItems(1).ListSubItems(1).text)
    targetObj = UserForm1.ListView1.ListItems(1).ListSubItems(3).text

    ' 원하는 작업을 할 시트를 지정
    Set ws = ThisWorkbook.Sheets("분석결과자료") ' 시트 이름을 자신의 시트 이름에 맞게 수정

    ' Find 메서드를 사용하여 일치하는 셀을 찾음
    Set FoundCell = ws.Columns(1).Find(what:=targetDate, LookIn:=xlValues, lookat:=xlWhole)


    ' 찾은 셀이 있고, 시료명이 일치하면 해당 셀의 행 번호를 출력
    Do While Not FoundCell Is Nothing

        If FoundCell.Offset(0, 1).Value = targetObj Then
            X = FoundCell.row
            Sheets("시험성적서").Cells(1, "C") = "RENEWUS-WAC-" & Format(FoundCell.Offset(0, 0).Value, "YYYY") & "-" & FoundCell.row & "-A"
            Sheets("시험성적서").Cells(1, "K") = "RENEWUS-WAC-" & Format(FoundCell.Offset(0, 0).Value, "YYYY") & "-" & FoundCell.row & "-B"
            For c = 1 To UserForm1.ListView3.ListItems.Count
              TC = Sheets("분석결과자료").Rows(1).Find(what:=UserForm1.ListView3.ListItems(c).text, lookat:=1).Column
              UserForm1.ListView3.ListItems(c).ListSubItems(1).text = Sheets("분석결과자료").Cells(X, TC).Value
              단위 = Sheets("측정DB").Columns(3).Find(what:=UserForm1.ListView3.ListItems(c).text, lookat:=xlWhole).row
              If c <= 32 Then
                  If Sheets("분석결과자료").Cells(X, TC).Value = "" Then
                  Sheets("시험성적서").Cells(9 + c, "F") = "분석전"
                  Sheets("시험성적서").Cells(9 + c, "G") = ""

                  ElseIf Sheets("분석결과자료").Cells(X, TC).Value = "불검출" Then
                  Sheets("시험성적서").Cells(9 + c, "F") = "분석전"
                  Sheets("시험성적서").Cells(9 + c, "G") = ""
                  Else
                  Sheets("시험성적서").Cells(9 + c, "F") = Sheets("분석결과자료").Cells(X, TC).Value
                  Sheets("시험성적서").Cells(9 + c, "F").NumberFormatLocal = Sheets("측정DB").Cells(단위, "A")
                  Sheets("시험성적서").Cells(9 + c, "G") = Sheets("측정DB").Cells(단위, "B").Value
                  End If
              Else
                  If Sheets("분석결과자료").Cells(X, TC).Value = "" Then
                  Sheets("시험성적서").Cells(9 + c - 32, "N") = "분석전"
                  Sheets("시험성적서").Cells(9 + c - 32, "O") = ""
                  Else
                  Sheets("시험성적서").Cells(9 + c - 32, "N") = Sheets("분석결과자료").Cells(X, TC).Value
                  Sheets("시험성적서").Cells(9 + c - 32, "N").NumberFormatLocal = Sheets("측정DB").Cells(단위, "A")
                  Sheets("시험성적서").Cells(9 + c - 32, "O") = Sheets("측정DB").Cells(단위, "B").Value
                  End If
                  
              End If
              
            Next c

            Exit Do ' 일치하는 셀을 찾았으므로 반복문 종료
        End If

        ' 일치하지 않으면 다음 일치하는 셀을 찾기 위해 다음 셀 검색
        Set FoundCell = ws.Columns(1).FindNext(FoundCell)
    Loop

    ' 모든 셀을 확인했지만 일치하는 셀을 찾지 못한 경우 메시지 출력
    If FoundCell Is Nothing Then
        Application.ScreenUpdating = True
        Exit Sub
        Debug.Print "일치하는 날짜를 찾을 수 없거나 시료명이 일치하지 않습니다."
    End If
    
    If Sheets("시험성적서").CheckBoxes("확인란 5").Value = 1 Then
     방류기준찾기
    End If
    
        Application.ScreenUpdating = True
End Sub
