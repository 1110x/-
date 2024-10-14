Attribute VB_Name = "분석결과입력_세팅"
Sub 분석결과입력리스트의법적기준()

    Dim targetDate As Date
    Dim targetObj As String
    Dim ws As Worksheet
    Dim FoundCell As Range
    Dim currentCell As Range
    Dim 단위
    Dim FullText As String
    Dim ExtractedText As String
    Dim BracketPosition As Integer

   Application.ScreenUpdating = True

   


    ' 원하는 작업을 할 시트를 지정
    Set ws = ThisWorkbook.Sheets("의뢰정보") ' 시트 이름을 자신의 시트 이름에 맞게 수정

    ' Find 메서드를 사용하여 일치하는 셀을 찾음
    Set FoundCell = ws.Columns(1).Find(what:=targetDate, LookIn:=xlValues, lookat:=xlWhole)
    Set RealSampList = UserForm1.ListView4.ListItems(UserForm1.ListView4.ListItems.Count).ListSubItems(2)
    BracketPosition = InStr(RealSampList, "】")
    
    If BracketPosition > 0 Then
       ExtractedText = Mid(RealSampList, BracketPosition + 1)
    Else
       ExtractedText = "해당 문자가 없습니다."
    End If

    
    
    Set ws = ThisWorkbook.Sheets("의뢰정보") ' 시트 이름을 자신의 시트 이름에 맞게 수정

    searchValue = UserForm1.ListView4.ListItems(UserForm1.ListView4.ListItems.Count).ListSubItems(1).text
    Set FoundCell = ws.Columns(1).Find(what:=CDate(searchValue), LookIn:=xlValues, lookat:=xlWhole)



    ' 찾은 셀이 있고, 시료명이 일치하면 해당 셀의 행 번호를 출력
    Do While Not FoundCell Is Nothing

        If FoundCell.Offset(0, 5).Value = ExtractedText Then
            
           UserForm1.ListView4.ListItems(UserForm1.ListView4.ListItems.Count).ListSubItems(3).text = FoundCell.Offset(0, 9)
            Exit Do ' 일치하는 셀을 찾았으므로 반복문 종료
        End If

        ' 일치하지 않으면 다음 일치하는 셀을 찾기 위해 다음 셀 검색
        Set FoundCell = ws.Columns(1).FindNext(FoundCell)
        
        
        
    Loop
    
End Sub
Sub 작업시작() '시료명 리스트 이동
Application.ScreenUpdating = False
Dim TX As Worksheet
Set TX = Sheets("분석결과 입력")
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
Sub 작업시작2() '분석항목
Application.ScreenUpdating = False

Dim TX As Worksheet
Set TX = Sheets("분석결과 입력")
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
Sub 분석결과자료최종입력()
    Dim TX As Worksheet
    Dim r As Long
    Dim text As String
    Dim startPos As Long
    Dim endPos As Long
    Dim textToExtract As String
    Dim lengthTextToExtract As Long
    
    ' 현재 워크시트 설정
    Set TX = Sheets("분석결과 입력")
    
    ' 데이터가 있는 셀 범위 반복 (100행부터 마지막 데이터까지)
    For r = 2 To TX.Cells(100, "C").End(xlUp).row
        text = TX.Cells(r, "C").Value
        
        ' 대괄호의 시작과 끝 위치 찾기
        startPos = InStr(text, "【")
        endPos = InStr(text, "】")
        
        If startPos > 0 And endPos > startPos Then
            textToExtract = Mid(text, endPos + 1)
            
            ' 결과를 셀에 작성
''''            TX.Cells(r, "D").Value = Trim(textToExtract) ' 결과를 D열에 입력

                 

            
            Debug.Print Trim(textToExtract)
        Else
            ' 대괄호가 없는 경우
'''            TX.Cells(r, "D").Value = "대괄호가 없음"
        End If
    Next r
End Sub

Function ExtractDesiredText(text As String) As String
    ' 원하는 텍스트를 추출하는 로직을 추가합니다.
    ' 예를 들어, 특정 문자열이 포함된 텍스트를 반환하도록 수정할 수 있습니다.
    
    ' 현재는 예시로 전체 텍스트를 그대로 반환하고 있습니다.
    ExtractDesiredText = text
    
    ' 추가 로직 필요에 따라 작성하세요.
End Function

' ListView4에서 중복 항목 검사
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

Sub 분석X결과입력하기()
    Dim targetDate As Date
    Dim targetObj As String
    Dim FoundCell As Range
    Dim currentCell As Range
    Dim 단위
    Dim RX As Worksheet
    Dim TX As Worksheet
    Dim XX As Range
    
    
    Set RX = Sheets("분석결과 입력")
    Set TX = ThisWorkbook.Sheets("분석결과자료") ' 시트 이름을 자신의 시트 이름에 맞게 수정
    
   Application.ScreenUpdating = False

   For r = 3 To Sheets("분석결과 입력").Cells(2, 3).End(xlDown).row
   

    ' 사용자로부터 날짜와 시료명을 얻어옴
    targetDate = Format(Sheets("분석결과 입력").Cells(r, "B"), "YYYY-MM-DD")
    targetObj = Right(RX.Cells(r, "C"), Len(RX.Cells(r, "C")) - InStr(RX.Cells(r, "C"), "】"))
       

    ' 원하는 작업을 할 시트를 지정


    ' Find 메서드를 사용하여 일치하는 셀을 찾음
    Set FoundCell = TX.Columns(1).Find(what:=targetDate, LookIn:=xlValues, lookat:=xlWhole)


    ' 찾은 셀이 있고, 시료명이 일치하면 해당 셀의 행 번호를 출력
    Do While Not FoundCell Is Nothing

        If FoundCell.Offset(0, 1).Value = targetObj Then
            X = FoundCell.row

            For H = Range("E1").Column To Range("BM1").Column
              If Not IsNumeric(RX.Cells(2, H).Value) Then
                Set XX = TX.Rows(1).Find(what:=RX.Cells(2, H).Value, lookat:=xlWhole)
                 
                 If XX <> "" Then
                   A = MsgBox(XX.Value & "의 기존입력된 결과값 [" & TX.Cells(X, XX.Column) & "] 에서 [" & RX.Cells(r, H) & "] 로 변경하시겠습니까", vbYesNo, RX.Cells(r, "C") & "기존자료 확인 ♡")
                    If A = vbYes Then
                    TX.Cells(X, XX.Column) = RX.Cells(r, H)
                    End If
                 Else
                    TX.Cells(X, XX.Column) = RX.Cells(r, H)
                 End If
                 
                 
              Else

              End If
              
            Next H

            Exit Do ' 일치하는 셀을 찾았으므로 반복문 종료
        End If

        ' 일치하지 않으면 다음 일치하는 셀을 찾기 위해 다음 셀 검색
        Set FoundCell = TX.Columns(1).FindNext(FoundCell)
    Loop

    ' 모든 셀을 확인했지만 일치하는 셀을 찾지 못한 경우 메시지 출력
    If FoundCell Is Nothing Then
        Application.ScreenUpdating = True
        Debug.Print "일치하는 날짜를 찾을 수 없거나 시료명이 일치하지 않습니다."
    End If
    

  Next r
  
        Application.ScreenUpdating = True
End Sub
