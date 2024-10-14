Attribute VB_Name = "공정시험법및분장"
Sub 시험법()
    Dim X As Integer
    Dim XT As Range
    Dim DT As Long
    Dim DR As Long
    Dim FoundCell As Range
    
    For X = 1 To UserForm1.ListView3.ListItems.Count
        ' "측정DB" 시트에서 해당 텍스트를 찾습니다.
        Set XT = Sheets("측정DB").Columns(3).Find(what:=UserForm1.ListView3.ListItems(X).text, lookat:=xlWhole)
        
        ' 찾은 경우에만 작업을 수행합니다.
        If Not XT Is Nothing Then
            ' 해당 셀의 값을 ListView에 할당합니다.
            UserForm1.ListView3.ListItems(X).ListSubItems(2).text = Sheets("측정DB").Cells(XT.row, "E").Value                       '분석방법
            UserForm1.ListView3.ListItems(X).ListSubItems(3).text = Sheets("측정DB").Cells(XT.row, "G").Value                       '분석장비
            UserForm1.ListView3.ListItems(X).ListSubItems(4).text = "-"                                                             '법적기준 ☆ 수정필요
            UserForm1.ListView3.ListItems(X).ListSubItems(5).text = Format(Sheets("측정DB").Cells(XT.row, "I").Value, "0000")       'Method NO
            UserForm1.ListView3.ListItems(X).ListSubItems(6).text = Sheets("측정DB").Cells(XT.row, "H").Value                       '분석장비 NO
            
            ' "분장" 시트에서 날짜를 찾습니다.
            Set FoundCell = Sheets("분장").Columns(1).Find(what:=CDate(UserForm1.ListView1.ListItems(1)), lookat:=xlWhole)
            If Not FoundCell Is Nothing Then
                DT = FoundCell.row
                ' "분장" 시트에서 해당 텍스트를 찾습니다.
                DR = Sheets("분장").Rows(1).Find(what:=UserForm1.ListView3.ListItems(X).text, lookat:=xlWhole).Column
                UserForm1.ListView3.ListItems(X).ListSubItems(7).text = Sheets("분장").Cells(DT, DR).text                           '분석장비 NO
            Else
                ' DT를 찾지 못한 경우에 대한 처리
                UserForm1.ListView3.ListItems(X).ListSubItems(7).text = "Not Found"
            End If
        Else
            ' XT를 찾지 못한 경우에 대한 처리
            ' MsgBox "찾을 수 없는 항목: " & ListView3.ListItems(X).Text
        End If
    Next X
End Sub
