Attribute VB_Name = "UPDATE"
Sub UpdateAllModulesAndFormsFromGitHub()
    Dim fileList As Variant
    Dim i As Integer
    
    ' Ʈ  
    fileList = Array( _
        "Combo.bas", "Module1.bas", "Module2.bas", "UPDATE.bas", "Module4.bas", "TESTModule.bas", _
        "UserForm1.frm", "UserForm1.frx", "Էϱ.bas", ".bas", "׺.bas", _
        "⺻.bas", "󳻱.bas", "׼.bas", "Ʈ_.bas", "_ãư.bas", _
        "ذŽñ.bas", "ظӽ.bas", "İŽñ.bas", "мԷ_.bas", _
        "мȸ.bas", "¹.bas", "Ҽ.bas", "Ϻ.bas", _
        "ǷڸƮ.bas", "ϱ.bas", "Ʈ_.bas", "Ư.bas", "հ.bas")
    
    '   ٿεϰ Ʈ
    For i = LBound(fileList) To UBound(fileList)
        Call DownloadAndUpdateComponent(fileList(i), Left(fileList(i), InStrRev(fileList(i), ".") - 1), Right(fileList(i), 3))
    Next i
End Sub

