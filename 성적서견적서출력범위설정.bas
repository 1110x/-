Attribute VB_Name = "��������������¹�������"
Sub ��ũ��5()
Attribute ��ũ��5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��ũ��5 ��ũ��
'

'
End Sub
Sub ��ũ��6()
Attribute ��ũ��6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��ũ��6 ��ũ��
'

'
    ActiveWindow.SmallScroll Down:=0
End Sub
Sub OpenGroups()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' �׷��� ���� ���� OutlineLevel�� �����մϴ�.
    ws.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
End Sub

Sub CloseGroups()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' �׷��� �ݱ� ���� OutlineLevel�� �����մϴ�.
    ws.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
End Sub

