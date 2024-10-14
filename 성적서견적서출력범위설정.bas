Attribute VB_Name = "성적서견적서출력범위설정"
Sub 매크로5()
Attribute 매크로5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로5 매크로
'

'
End Sub
Sub 매크로6()
Attribute 매크로6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 매크로6 매크로
'

'
    ActiveWindow.SmallScroll Down:=0
End Sub
Sub OpenGroups()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 그룹을 열기 위해 OutlineLevel을 설정합니다.
    ws.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
End Sub

Sub CloseGroups()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' 그룹을 닫기 위해 OutlineLevel을 설정합니다.
    ws.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
End Sub

