Attribute VB_Name = "Module2"
Sub TEST20240819()
    Dim X As Range
    
    ' G10:G15 ���� ���� �� ���� ���� �ݺ�
    For Each X In Range("H10:H41,P10:P41").Cells
        X.Value
    Next X

End Sub
