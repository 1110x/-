Attribute VB_Name = "UPDATE"
Sub UpdateAllModulesAndFormsFromGitHub()
    Dim fileList As Variant
    Dim i As Integer
    
    ' ������Ʈ�� ���� ���
    fileList = Array( _
        "Combo.bas", "Module1.bas", "Module2.bas", "UPDATE.bas", "Module4.bas", "TESTModule.bas", _
        "UserForm1.frm", "UserForm1.frx", "�������Է��ϱ�.bas", "��������.bas", "����������׺���.bas", _
        "�⺻����.bas", "������󳻱�.bas", "����׼���.bas", "����Ʈ��_����.bas", "�������_ã�ư�.bas", _
        "������ذŽñ�.bas", "������ظ�����ӽ�.bas", "������İŽñ�.bas", "�м�����Է�_����.bas", _
        "�м������ȸ.bas", "��������������¹�������.bas", "�Ҽ���.bas", "����������Ϻ�.bas", _
        "�Ƿڸ���Ʈ�����.bas", "����ϱ�.bas", "Ʈ����_����.bas", "Ư����������.bas", "�հ�.bas")
    
    ' ��� ������ �ٿ�ε��ϰ� ������Ʈ
    For i = LBound(fileList) To UBound(fileList)
        Call DownloadAndUpdateComponent(fileList(i), Left(fileList(i), InStrRev(fileList(i), ".") - 1), Right(fileList(i), 3))
    Next i
End Sub

