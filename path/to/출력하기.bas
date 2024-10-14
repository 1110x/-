Attribute VB_Name = "출력하기"
Sub SetPrintPageSettings()
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim RowsPerPage As Long
    Dim TotalPages As Long

    ' 마지막 행과 열을 찾습니다.
    lastRow = Cells.Find(what:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    lastColumn = Cells.Find(what:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' 한 페이지당 출력될 행 수를 계산합니다.
    RowsPerPage = 42 ' 예시로 42행을 한 페이지로 설정했습니다.

    ' 총 페이지 수를 계산합니다.
    TotalPages = Application.WorksheetFunction.Ceiling(lastRow / RowsPerPage, 1)

    ' 페이지 설정을 합니다.
    With ActiveSheet.PageSetup
        .Orientation = xlPortrait '페이지 방향 설정 (수직 방향)
        .Zoom = False '축소/확대 옵션 해제
        .FitToPagesWide = 1 '페이지 너비에 맞추기
        .FitToPagesTall = TotalPages '총 페이지 수로 페이지 높이 설정
        '추가적인 페이지 설정을 원하는 경우 여기에 추가
    End With
End Sub


Sub 붉은색으로물들여라()
UserForm1.TreeView1.SelectedItem.ForeColor = RGB(255, 0, 0)
UserForm1.TreeView1.SelectedItem.Selected = False



For c = 1 To UserForm1.TreeView1.Nodes.Count
 If UserForm1.TreeView1.Nodes(c).ForeColor = RGB(255, 0, 0) Then
    r = r + 1
 End If
Next c

If r > 1 Then

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
    
For RED = s To E
    If Not UserForm1.TreeView1.Nodes(RED).Parent Is Nothing Then ' 부모 노드가 없는 경우에만 작업 수행
        If Not UserForm1.TreeView1.Nodes(RED).Children Then ' 자식 노드가 없는 경우에만 작업 수행
            UserForm1.TreeView1.Nodes(RED).ForeColor = RGB(255, 0, 0)
        End If
    End If
Next RED

End If
End Sub
