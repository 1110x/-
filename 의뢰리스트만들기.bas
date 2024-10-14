Attribute VB_Name = "의뢰리스트만들기"
'Global SR As Integer

Sub 의뢰리스트이동()
    Dim X As String
    Dim i As Integer
    Dim IsDuplicate As Boolean
    
    X = UserForm1.TreeView4.SelectedItem.text
    
    ' 중복 여부를 확인합니다.
    IsDuplicate = False
    For i = 0 To UserForm1.ListBox2.ListCount - 1
        If UserForm1.ListBox2.List(i, 1) = X Then
            IsDuplicate = True
            Exit For
        End If
    Next i
    
    ' 중복되지 않는 경우에만 항목을 추가합니다.
    If Not IsDuplicate Then
        UserForm1.ListBox2.AddItem "" ' 빈 항목 추가하여 새로운 행 생성
        ' ListBox의 마지막 추가된 행의 특정 열에 값을 설정
        UserForm1.ListBox2.List(UserForm1.ListBox2.ListCount - 1, 0) = Format(UserForm1.ListBox2.ListCount, "00")
        UserForm1.ListBox2.List(UserForm1.ListBox2.ListCount - 1, 1) = X
    Else
        Exit Sub
        
    End If
End Sub

Sub 의뢰항목체크()
Application.ScreenUpdating = False
        Set ws = ThisWorkbook.Sheets("견적발행정보")
        lastRow = Sheets("견적발행정보").Cells(Sheets("견적발행정보").Rows.Count, "A").End(xlUp).row
        For r = 2 To lastRow
         Z = "【" & Sheets("견적발행정보").Cells(r, "C").text & "】" & Sheets("견적발행정보").Cells(r, "H").Value '  ws.Cells(i, "C")  & ws.Cells(i, 4).Value)
         
         
                If Sheets("견적발행정보").Cells(r, "A") = Sheets("분석의뢰 입력").Cells(1, "F") And Z = Sheets("분석의뢰 입력").Cells(1, "B") Then
                      Debug.Print Z
                       Sheets("분석의뢰 입력").Cells(1, "A") = Sheets("견적발행정보").Cells(r, "C")
                       Set 약칭 = Sheets("계약정보").Columns(8).Find(what:=Sheets("견적발행정보").Cells(r, "C").text, lookat:=xlWhole)
                       
                       If Not 약칭 Is Nothing Then
                         Sheets("분석의뢰 입력").Cells(1, "A") = Sheets("계약정보").Cells(약칭.row, "H")
                       End If
                                     Total = 0
                                     For X = 13 To 193 Step (3)
                                             If Sheets("견적발행정보").Cells(r, X) <> "" Then
                                                 For TR = 3 To Cells(2, "B").End(xlDown).row
                                                 
                                                  TX = Sheets("분석의뢰 입력").Rows(2).Find(what:=Sheets("견적발행정보").Cells(1, X), lookat:=xlWhole).Column
                                                  Sheets("분석의뢰 입력").Cells(TR, TX) = "O"
                                                 Next TR
                                             End If
                                     Next X
                       Application.ScreenUpdating = True
                       Exit Sub
                End If
        Next r
Application.ScreenUpdating = True
End Sub
Sub 분석의뢰입력TO의뢰정보()
    Dim TS As Worksheet
    Dim SS As Worksheet
    Dim FoundCell As Range
    Dim FirstFound As Range ' 첫 번째로 찾은 셀을 기억하는 변수
    Dim targetDate As String
    Dim 의뢰사업자 As String
    Dim 약칭 As String
    Dim 채취자1 As String, 채취자2 As String, 입회자 As String
    Dim SR As Long, TR As Long, SC As Long
    Dim AN As Integer
    Dim isFound As Boolean ' 찾았는지 여부를 기록하는 변수
    Dim XE As Integer ' 디버그용 변수
    
    Set TS = Sheets("의뢰정보")
    Set SS = Sheets("분석의뢰 입력")

    ' 시료리스트가 존재하지 않으면 즉시 서브 종료
    If SS.Cells(50, "B").End(xlUp).row < 3 Then
        Debug.Print "종료"
        Exit Sub
    End If

    ' 입력 값 설정
    의뢰사업자 = Sheets("계약정보").Cells(Sheets("계약정보").Columns("H").Find(what:=SS.Cells(1, 1), lookat:=xlWhole).row, "B")
    약칭 = SS.Cells(1, 1)
    채취자1 = SS.Cells(1, "P")
    채취자2 = SS.Cells(1, "S")
    입회자 = SS.Cells(1, "X")

    targetDate = Format(SS.Cells(1, "F"), "YYYY-MM-DD")

    ' 분석의뢰 입력 데이터 처리
    For SR = 3 To SS.Cells(50, "B").End(xlUp).row
        TR = TS.Cells(2, "A").End(xlDown).row + 1
        Set FoundCell = TS.Columns(1).Find(what:=CDate(targetDate), LookIn:=xlValues, lookat:=xlWhole)

        ' 초기 상태로 찾은 값 없음
        isFound = False

        ' 찾는 값이 없으면 바로 새 데이터 입력으로 이동
        If FoundCell Is Nothing Then
            GoTo NoMatchFound
        Else
            ' 첫 번째로 찾은 셀을 기억
            Set FirstFound = FoundCell
        End If

        ' 데이터 중복 확인 및 처리
        Do
            ' 동일한 셀로 돌아왔을 경우 루프 종료
            If FoundCell.Address = FirstFound.Address And isFound Then
                Exit Do
            End If

            ' 중복된 데이터 확인 및 사용자에게 확인 메시지
            If FoundCell.Offset(0, 5).Value = SS.Cells(SR, "B").Value Then
                AN = MsgBox(SS.Cells(1, "F") & "의" & vbCrLf & " 시료가 이미 입력되어 있습니다." & vbCrLf & " 재입력하시겠습니까?", vbYesNo, SS.Cells(SR, "B") & " 의뢰 중복입력♡")
                If AN = vbYes Then
                    ' 중복 데이터 덮어쓰기
                    TS.Cells(FoundCell.row, "A") = SS.Cells(1, "F") ' 날짜
                    TS.Cells(FoundCell.row, "B") = SS.Cells(1, "K") ' 기타 정보
                    TS.Cells(FoundCell.row, "D") = 의뢰사업자
                    TS.Cells(FoundCell.row, "E") = 약칭
                    TS.Cells(FoundCell.row, "F") = SS.Cells(SR, "B") ' 시료명
                    TS.Cells(FoundCell.row, "G") = 입회자
                    TS.Cells(FoundCell.row, "H") = 채취자1
                    TS.Cells(FoundCell.row, "I") = 채취자2

                    ' 분석 항목 입력
                    For SC = 4 To 64 ' D3부터 BK1까지
                        If SS.Cells(SR, SC).Value <> "" Then
                            TS.Cells(FoundCell.row, 10 + SC) = "O"
                        Else
                            TS.Cells(FoundCell.row, 10 + SC) = ""
                        End If
                    Next SC
                End If
                isFound = True
                Exit Do
            End If

            ' 다음 찾기
            Set FoundCell = TS.Columns(1).FindNext(FoundCell)

            ' 다시 첫 번째로 찾은 셀로 돌아오면 종료
            If FoundCell Is Nothing Or FoundCell.Address = FirstFound.Address Then Exit Do

        Loop

NoMatchFound:
        ' 중복 데이터가 없을 경우 새로 데이터 입력
        If Not isFound Then
            TR = TS.Cells(2, "A").End(xlDown).row + 1
            TS.Cells(TR, "A") = SS.Cells(1, "F") ' 날짜
            TS.Cells(TR, "B") = SS.Cells(1, "K") ' 기타 정보
            TS.Cells(TR, "D") = 의뢰사업자
            TS.Cells(TR, "E") = 약칭
            TS.Cells(TR, "F") = SS.Cells(SR, "B") ' 시료명
            TS.Cells(TR, "G") = 입회자
            TS.Cells(TR, "H") = 채취자1
            TS.Cells(TR, "I") = 채취자2

            ' 분석 항목 입력
            For SC = 4 To 64 ' D3부터 BK1까지
                If SS.Cells(SR, SC).Value <> "" Then
                    TS.Cells(TR, 10 + SC) = "O"
                Else
                    TS.Cells(TR, 10 + SC) = ""
                End If
            Next SC
        End If
    Call 분석항목적용(SR)
    Next SR

    
End Sub

Sub 분석항목적용(ByVal SR As Long)
    Dim TS As Worksheet
    Dim SS As Worksheet
    Dim FirstFound As Range ' 첫 번째로 찾은 셀을 기억하는 변수
    Set TS = Sheets("분석결과자료")
    Set SS = Sheets("분석의뢰 입력")
    
    Debug.Print SR
    
    
    targetDate = Format(SS.Cells(1, "F"), "YYYY-MM-DD")
    
Set FoundCell = TS.Columns(1).Find(what:=CDate(targetDate), LookIn:=xlValues, lookat:=xlWhole)
        ' 초기 상태로 찾은 값 없음
        isFound = False

        ' 찾는 값이 없으면 바로 새 데이터 입력으로 이동
        If FoundCell Is Nothing Then
            GoTo NoMatchFound
        Else
            ' 첫 번째로 찾은 셀을 기억
            Set FirstFound = FoundCell
        End If
        
        Do
            ' 동일한 셀로 돌아왔을 경우 루프 종료
            If FoundCell.Address = FirstFound.Address And isFound Then
                Exit Do
            End If

            ' 중복된 데이터 확인 및 사용자에게 확인 메시지
            If FoundCell.Offset(0, 1).Value = SS.Cells(SR, "B").Value Then
                AN = MsgBox(SS.Cells(1, "F") & "의 해당시료가 분석결과표에" & vbCrLf & " 이미존재합니다." & vbCrLf & " 재입력하시겠습니까?", vbYesNo, SS.Cells(SR, "B") & " 의뢰항목 중복☆")
                If AN = vbYes Then

                    ' 분석(의뢰)항목 입력
                    For SC = 4 To 64 ' D3부터 BK1까지
                        If SS.Cells(SR, SC).Value <> "" Then
                            TS.Cells(FoundCell.row, SC - 1).Interior.Pattern = -4142
                        Else
                            TS.Cells(FoundCell.row, SC - 1).Interior.Pattern = -4124
                        End If
                    Next SC
                End If
                isFound = True
                Exit Do
            End If

            ' 다음 찾기
            Set FoundCell = TS.Columns(1).FindNext(FoundCell)

            ' 다시 첫 번째로 찾은 셀로 돌아오면 종료
            If FoundCell Is Nothing Or FoundCell.Address = FirstFound.Address Then Exit Do

        Loop

NoMatchFound:
        ' 중복 데이터가 없을 경우 새로 데이터 입력
        If Not isFound Then
            TR = TS.Cells(2, "A").End(xlDown).row + 1
            ' 분석 항목 입력
            For SC = 4 To 64 ' D3부터 BK1까지
               TS.Cells(TR, 1) = SS.Cells(1, "K")
               TS.Cells(TR, "B") = SS.Cells(SR, "B")
               
                If SS.Cells(SR, SC).Value <> "" Then
                    TS.Cells(TR, SC - 1).Interior.Pattern = -4142
                Else
                    TS.Cells(TR, SC - 1).Interior.Pattern = -4124
                End If
                
            Next SC
        End If


End Sub
