Attribute VB_Name = "ϱ"
Sub SetPrintPageSettings()
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim RowsPerPage As Long
    Dim TotalPages As Long

    '    ãϴ.
    lastRow = Cells.Find(what:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    lastColumn = Cells.Find(what:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    '   µ   մϴ.
    RowsPerPage = 42 ' ÷ 42   ߽ϴ.

    '    մϴ.
    TotalPages = Application.WorksheetFunction.Ceiling(lastRow / RowsPerPage, 1)

    '   մϴ.
    With ActiveSheet.PageSetup
        .Orientation = xlPortrait '   ( )
        .Zoom = False '/Ȯ ɼ 
        .FitToPagesWide = 1 ' ʺ ߱
        .FitToPagesTall = TotalPages '     
        '߰   ϴ  ⿡ ߰
    End With
End Sub


Sub ι鿩()
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
    If Not UserForm1.TreeView1.Nodes(RED).Parent Is Nothing Then ' θ 尡  쿡 ۾ 
        If Not UserForm1.TreeView1.Nodes(RED).Children Then ' ڽ 尡  쿡 ۾ 
            UserForm1.TreeView1.Nodes(RED).ForeColor = RGB(255, 0, 0)
        End If
    End If
Next RED

End If
End Sub
