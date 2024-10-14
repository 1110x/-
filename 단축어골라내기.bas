Attribute VB_Name = "단축어골라내기"
Sub 단축어찾기()
    Dim text As String
    Dim result As String
    Dim startPos As Long
    Dim endPos As Long
    Dim ExtractedText As String

    ' TextBox8의 텍스트 가져오기
    text = UserForm1.TextBox8.text
    result = ""

    ' 대괄호 안의 텍스트 추출
    Do
        ' 대괄호의 시작과 끝 위치 찾기
        startPos = InStr(text, "【")
        endPos = InStr(startPos + 1, text, "】")

        ' 대괄호가 모두 존재하는 경우
        If startPos > 0 And endPos > startPos Then
            ' 대괄호 안의 텍스트 추출
            ExtractedText = Mid(text, startPos + 1, endPos - startPos - 1)
            result = result & ExtractedText & " "

            ' 다음 대괄호를 찾기 위해 텍스트 갱신
            text = Mid(text, endPos + 1)
        Else
            Exit Do
        End If
    Loop

    ' 결과 출력
    약어 = Trim(result)

    
UserForm1.ComboBox5.ListIndex = Sheets("계약정보").Columns("H").Find(what:=약어, lookat:=xlWhole).row - 2

    
End Sub
