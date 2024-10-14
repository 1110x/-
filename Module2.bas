Attribute VB_Name = "Module2"
Sub TEST20240819()
    Dim X As Range
    
    ' G10:G15 범위 내의 각 셀에 대해 반복
    For Each X In Range("H10:H41,P10:P41").Cells
        X.Value
    Next X

End Sub
