Attribute VB_Name = "견적서입력하기"
Sub 견적서입력()
Application.ScreenUpdating = False
On Error Resume Next

Sheets("견적서").Cells(3, "C") = UserForm1.ComboBox3.Value '견적대상 업체명
Sheets("견적서").Cells(5, "C") = UserForm1.ComboBox2.Value
Sheets("견적서").Cells(6, "C") = UserForm1.TextBox5.Value
Sheets("견적서").Cells(7, "C") = UserForm1.TextBox4.Value
Sheets("견적서").Cells(8, "C") = UserForm1.TextBox2.Value
Sheets("견적서").Cells(3, "H") = Format(UserForm1.TextBox1.Value, "YYYYMMDD") & "-" & Format(Now(), "HHMMSS")
Sheets("견적서").Cells(4, "H") = UserForm1.TextBox1
Sheets("견적서").Cells(6, "H") = UserForm1.ComboBox4.Value & "/061-685-8186"

Set 담당자 = Sheets("담당자정보").Rows(100).Find(what:=UserForm1.ComboBox4.Value, lookat:=xlWhole)

If Not 담당자 Is Nothing Then
    Sheets("견적서").Cells(7, "H") = Sheets("담당자정보").Cells(100, 담당자.Column + 3)
End If

Sheets("견적서").Range("A11:I43,K4:S36") = ""

For r = 0 To UserForm1.ListBox1.ListCount - 1

'    Sheets("견적서").Cells(r + 11, "A") = r + 1
'    Sheets("견적서").Cells(r + 11, "B") = UserForm1.ListBox1.List(r, 0)
'    Sheets("견적서").Cells(r + 11, "D") = UserForm1.ListBox1.List(r, 1)
'    Set 실험항목 = Sheets("측정DB").Columns(3).Find(WHAT:=UserForm1.ListBox1.List(r, 1), lookat:=xlWhole)
'
'    If Not 실험항목 Is Nothing Then
'    Sheets("견적서").Cells(r + 11, "F") = Sheets("측정DB").Cells(실험항목.row, "D")
'    End If
'    Sheets("견적서").Cells(r + 11, "H") = UserForm1.TextBox3.Value
'    Sheets("견적서").Cells(r + 11, "I") = UserForm1.ListBox1.List(r, 3)
    
    
                        x1 = Sheets("견적서").Cells(100, "D").End(xlUp).row + 1
                        x2 = Sheets("견적서").Cells(100, "N").End(xlUp).row + 1
                        
                        If Sheets("견적서").Cells(100, "D").End(xlUp).row + 1 < 44 Then
                             X = Sheets("견적서").Cells(100, "D").End(xlUp).row + 1
                             what = 0
                             Sheets("견적서").Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
                             N = 0
                             Sheets("견적서").Range("A9").Font.Color = 16777215
                             Sheets("견적서").Range("A45").Font.Color = 8421504
                             Sheets("견적서").Range("A44").Font.Color = 0
                          Else
                             X = Sheets("견적서").Cells(100, "N").End(xlUp).row + 1
                             what = 10
                             Sheets("견적서").Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
                             N = 40
                             Sheets("견적서").Range("A9").Font.Color = 8421504
                             Sheets("견적서").Range("A45,H44").Font.Color = 16777215
                        End If
                        
                        Sheets("견적서").Cells(X, what + 1) = X - 10 + N
                        Sheets("견적서").Cells(X, what + 2) = UserForm1.ListBox1.List(r, 0)
                        Sheets("견적서").Cells(X, what + 4) = UserForm1.ListBox1.List(r, 1)
                        amount = ws.Cells(Application.Match(childNode.text, ws.Columns(4), 0), selectedColumn).Value
                        
                        Set ES = Sheets("측정DB").Columns(3).Find(what:=UserForm1.ListBox1.List(r, 1), lookat:=xlWhole)
                        If Not ES Is Nothing Then
                        Sheets("견적서").Cells(X, what + 6) = Sheets("측정DB").Cells(ES.row, "D")
                        End If
                        
'                        If Sheets("견적서").Cells(X - 1, WHAT + 8) = "수량" Then
                         Sheets("견적서").Cells(X, what + 8) = UserForm1.TextBox3.Value
'                        Else
'                         Sheets("견적서").Cells(X, WHAT + 8) = Sheets("견적서").Cells(X - 1, WHAT + 8)
'                        End If
                        
                        Sheets("견적서").Cells(X, what + 9) = UserForm1.ListBox1.List(r, 3)
                        
                        
                        
    
    
Next r


Application.ScreenUpdating = True

End Sub

Sub 견적발행_정보입력()
    Dim TX As Worksheet
    Dim RX As Worksheet
    Dim SX As Worksheet
    Dim 약칭 As Long
    Dim RTX As Long
    Dim RSX As Long
    Dim oldlist As Range
    Dim IsDuplicate As Boolean

    Set RX = ThisWorkbook.Sheets("견적서")
    Set TX = ThisWorkbook.Sheets("견적발행정보")
    Set SX = ThisWorkbook.Sheets("의뢰정보")

    RTX = TX.Cells(2, "A").End(xlDown).row + 1
    RSX = SX.Cells(2, "A").End(xlDown).row + 1
    RRX = RX.Cells(11, "A").End(xlDown).row
    RRXX = RX.Cells(4, "K").End(xlDown).row
    ' 중복 확인을 위한 초기값
    IsDuplicate = False

    ' TX 시트에서 I열의 기존 항목들 확인
    For Each oldlist In TX.Range(TX.Cells(3, "I"), TX.Cells(RTX - 1, "I"))
        If oldlist.Value = RX.Cells(3, "H").Value Then
            IsDuplicate = True
            RTX = oldlist.row
            Exit For
        End If
    Next oldlist

    ' 중복이 확인되면 사용자에게 물어보기
    If IsDuplicate Then
        A = MsgBox(RX.Cells(3, "H").Value & "의 견적서는 이미 입력되어 있습니다." & vbCrLf & "다시(수정) 입력하시겠습니까?", vbYesNo)
        If A = vbNo Then Exit Sub
    End If

    ' 데이터 입력
    TX.Cells(RTX, "A").Value = RX.Cells(4, "H").Value ' 견적발행일자
    TX.Cells(RTX, "B").Value = RX.Cells(3, "C").Value ' 견적 요청업체명
    약칭 = Sheets("계약정보").Columns(2).Find(what:=Sheets("견적서").Cells(3, "C").Value, lookat:=xlWhole).row
    TX.Cells(RTX, "C").Value = Sheets("계약정보").Cells(약칭, "H").Value ' 약칭
    TX.Cells(RTX, "D").Value = Sheets("계약정보").Cells(약칭, "D").Value ' 대표자
    TX.Cells(RTX, "E").Value = RX.Cells(5, "C").Value ' 견적요청자 성명
    TX.Cells(RTX, "F").Value = RX.Cells(6, "C").Value ' 견적요청자 전화번호
    TX.Cells(RTX, "G").Value = RX.Cells(7, "C").Value ' 견적요청자 이메일
    TX.Cells(RTX, "H").Value = RX.Cells(8, "C").Value ' 견적 시료명
    TX.Cells(RTX, "I").Value = RX.Cells(3, "H").Value ' 견적번호
    
    For Each 의뢰항목 In RX.Range(RX.Cells(11, "D"), RX.Cells(RRX, "D"))
        If 의뢰항목 <> "" Then
            TC = TX.Rows(1).Find(what:=의뢰항목, lookat:=xlWhole).Column
            TX.Cells(RTX, TC) = RX.Cells(의뢰항목.row, "H")
            TX.Cells(RTX, TC + 1) = RX.Cells(의뢰항목.row, "I")
            TX.Cells(RTX, TC + 2) = RX.Cells(의뢰항목.row, "J")
        End If
    Next 의뢰항목
    
    If Application.Count(RX.Range("K4:K36")) > 0 Then
        For Each 의뢰항목 In RX.Range(RX.Cells(4, "N"), RX.Cells(RRXX, "N"))
            If 의뢰항목 <> "" Then
                TC = TX.Rows(1).Find(what:=의뢰항목, lookat:=xlWhole).Column
                TX.Cells(RTX, TC) = RX.Cells(의뢰항목.row, "R")
                TX.Cells(RTX, TC + 1) = RX.Cells(의뢰항목.row, "S")
                TX.Cells(RTX, TC + 2) = RX.Cells(의뢰항목.row, "T")
            End If
        Next 의뢰항목
    End If
    
    
    MsgBox RX.Cells(3, "C") & " " & RX.Cells(8, "C") & vbCrLf & " 견적서 입력 완료"
End Sub

