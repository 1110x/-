Attribute VB_Name = "수질측정기록부"
Sub 수질측정기록부자료()
 Dim Node As MSComctlLib.Node 'TreeView 컨트롤의 노드를 나타내는 개체
    
    
    For X = 1 To UserForm1.TreeView1.Nodes.Count
     If UserForm1.TreeView1.Nodes(X).ForeColor = RGB(255, 0, 0) Then
        s = X
        Exit For
     End If
    Next X

    For X = UserForm1.TreeView1.Nodes.Count To 1 Step (-1)
     If UserForm1.TreeView1.Nodes(X).ForeColor = RGB(255, 0, 0) Then
        E = X
        Exit For
     End If
    Next X
    
    
For xxx = E To s Step (-1)
    Set Node = UserForm1.TreeView1.Nodes(xxx) '노드 2를 선택합니다.
    Node.EnsureVisible '노드가 보이도록 스크롤합니다.
    Node.Selected = True '노드를 선택합니다.


    Dim ws As Worksheet
If Not UserForm1.TreeView1.SelectedItem Is Nothing Then
    If Not UserForm1.TreeView1.SelectedItem.Parent Is Nothing Then
'        Debug.Print "자식노드: " & TreeView1.SelectedItem.Parent.text
        Set ws = ThisWorkbook.Sheets("의뢰정보")
        lastRow = Sheets("의뢰정보").Cells(Sheets("의뢰정보").Rows.Count, "A").End(xlUp).row
        For r = 2 To lastRow
         X = "【" & Sheets("의뢰정보").Cells(r, "E").text & "】" & Sheets("의뢰정보").Cells(r, "F").Value '  ws.Cells(i, "C")  & ws.Cells(i, 4).Value)
         If Sheets("의뢰정보").Cells(r, "A") = UserForm1.TreeView1.SelectedItem.Parent.text And X = UserForm1.TreeView1.SelectedItem.text Then

          '===============================================================================================
         UserForm1.ListView1.ListItems.Clear
         With UserForm1.ListView1
        .ColumnHeaders.Clear ' 기존 컬럼 제거
        .Gridlines = True
        .View = lvwReport ' Report 모드로 설정

        ' 각 컬럼의 너비를 조절하려면 필요에 따라 Width 속성을 설정할 수 있습니다.
        .ColumnHeaders.Add , , "의뢰일자", 100
        .ColumnHeaders.Add , , "채취일자", 100
        .ColumnHeaders.Add , , "의뢰사업장", 100
        .ColumnHeaders.Add , , "시료명", 120
        .ColumnHeaders.Add , , "입회자", 100

         End With

         UserForm1.ListView2.ListItems.Clear
         With UserForm1.ListView2
        .ColumnHeaders.Clear ' 기존 컬럼 제거
        .Gridlines = True
        .View = lvwReport ' Report 모드로 설정
        ' 각 컬럼의 너비를 조절하려면 필요에 따라 Width 속성을 설정할 수 있습니다.
        .ColumnHeaders.Add , , "시료채취자1", 100
        .ColumnHeaders.Add , , "시료채취자2", 100
        .ColumnHeaders.Add , , "방류허용기준", 120
        .ColumnHeaders.Add , , "정도보증유무", 100
        .ColumnHeaders.Add , , "분석종료일", 100
         End With

         UserForm1.ListView3.ListItems.Clear
         With UserForm1.ListView3
        .ColumnHeaders.Clear ' 기존 컬럼 제거
        .Gridlines = True
        .View = lvwReport ' Report 모드로 설정
        ' 각 컬럼의 너비를 조절하려면 필요에 따라 Width 속성을 설정할 수 있습니다.
        .ColumnHeaders.Add , , "의뢰항목", 130
        .ColumnHeaders.Add , , "분석결과", 70
        .ColumnHeaders.Add , , "분석방법", 170
        .ColumnHeaders.Add , , "분석장비", 100
        .ColumnHeaders.Add , , "법적기준", 50
        .ColumnHeaders.Add , , "Method NO", 20
        .ColumnHeaders.Add , , "instrument NO", 20
        .ColumnHeaders.Add , , "분석담당자", 20
         End With


         Set item = UserForm1.ListView1.ListItems.Add(1, , Sheets("의뢰정보").Cells(r, "A").Value) '의뢰일자
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "B").Value                '채취일자
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "E").Value                '의뢰사업장 '약칭
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "F").Value                '시료명
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "G").Value                '입회자

         Set item = UserForm1.ListView2.ListItems.Add(1, , Sheets("의뢰정보").Cells(r, "H").Value) '시료채취자-1
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "I").Value                '시료채취자-2
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "J").Value                '방류허용기준
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "K").Value                '정도보증유무
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "L").Value                '분석종료일

         For T = Range("N1").Column To Range("BT1").Column
         If Sheets("의뢰정보").Cells(r, T) <> "" Then
         Set Titem = UserForm1.ListView3.ListItems.Add(UserForm1.ListView3.ListItems.Count + 1, , Sheets("의뢰정보").Cells(1, T))
                     Titem.ListSubItems.Add , , "-"             'Subitem-1 분석결과
                     Titem.ListSubItems.Add , , "-"             'Subitem-2 분석방법
                     Titem.ListSubItems.Add , , "-"             'Subitem-3 분석장비
                     Titem.ListSubItems.Add , , "-"             'Subitem-4 법적기준
                     Titem.ListSubItems.Add , , "-"             'Subitem-5 Method NO
                     Titem.ListSubItems.Add , , "-"             'Subitem-6 instrument NO
                     Titem.ListSubItems.Add , , "-"             'Subitem-7 분석담당자
         End If
         Next T
         '===============================================================================================

         End If

        Next r

        Call 시험법찾기
        UserForm1.ListView3.ColumnHeaders(1).text = "의뢰항목 (총 " & UserForm1.ListView3.ListItems.Count & "건) "
    Else
        UserForm1.Label1.Caption = "못찾겄습니다"
    End If

분석결과불러오기
방류기준찾기
법정양식
UserForm1.TreeView1.Nodes(xxx).ForeColor = RGB(0, 0, 128)
ActiveSheet.PrintOut
End If


Next xxx

End Sub
   
