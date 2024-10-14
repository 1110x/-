Attribute VB_Name = "UPDATE"
Sub ExportModulesAndUserForms()
    Dim vbComp As Object
    Dim exportPath As String
    Dim fileName As String
    
    '   
    exportPath = "C:\CENTER\" ' ϴ η ϼ

    ' ΰ ȿ Ȯ
    If Right(exportPath, 1) <> "\" Then
        exportPath = exportPath & "\"
    End If

    '   
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_Module
                fileName = exportPath & vbComp.Name & ".bas"
                vbComp.Export fileName
                MsgBox vbComp.Name & "  " & fileName & " ϴ."
            Case vbext_ct_MSForm
                fileName = exportPath & vbComp.Name & ".frm"
                vbComp.Export fileName
                MsgBox vbComp.Name & "  " & fileName & " ϴ."
        End Select
    Next vbComp

    MsgBox "   ϴ."
End Sub


Sub UpdateAllModulesAndFormsFromGitHub()
    Dim fileList As Variant
    Dim i As Integer
    '??
    ' Ʈ  
    fileList = Array( _
        "Combo.bas", "Module1.bas", "Module2.bas", "Module3.bas", "UPDATE.bas", "Module4.bas", "TESTModule.bas", _
        "UserForm1.frm", "UserForm1.frx", "Էϱ.bas", ".bas", "׺.bas", _
        "⺻.bas", "󳻱.bas", "׼.bas", "Ʈ_.bas", "_ãư.bas", _
        "ذŽñ.bas", "ظӽ.bas", "İŽñ.bas", "мԷ_.bas", _
        "мȸ.bas", "¹.bas", "Ҽ.bas", "Ϻ.bas", _
        "ǷڸƮ.bas", "ϱ.bas", "Ʈ_.bas", "Ư.bas", "հ.bas")
    
    '   ٿεϰ Ʈ
    For i = LBound(fileList) To UBound(fileList)
        DownloadAndUpdateComponentFromGitHub fileList(i)
    Next i
End Sub

Sub DownloadAndUpdateComponentFromGitHub(ByVal fileName As String)
    Dim http As Object
    Dim url As String
    Dim fileData As String
    Dim filePath As String
    Dim componentType As String
    Dim stream As Object
    
    '  Ȯ 
    componentType = Right(fileName, 3)
    
    ' 꿡   (Raw URL )
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://raw.githubusercontent.com/1110x/Center/main/" & fileName
    http.Open "GET", url, False
    http.Send
    
    ' HTTP û   ó
    If http.Status = 200 Then
        fileData = http.responseText
        filePath = "C:\Users\ironu\OneDrive\ ȭ\" & fileName ' θ  
        
        ' ADODB.Stream ü Ͽ UTF-8  
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2 ' ؽƮ
        stream.Charset = "utf-8" ' UTF-8 
        stream.Open
        stream.WriteText fileData
        stream.SaveToFile filePath, 2 ' 2 
        stream.Close
        
        '  VBA     
        On Error Resume Next
        If componentType = "bas" Or componentType = "frm" Then
            ' VBProject ϴ Ȯ
            If ThisWorkbook.VBProject Is Nothing Then
                MsgBox "VBProject ã  ϴ."
                Exit Sub
            End If
            
            ' Ұ ϴ Ȯ  
            Dim comp As Object
            Set comp = ThisWorkbook.VBProject.VBComponents(Left(fileName, Len(fileName) - 4))
            If Not comp Is Nothing Then
                ThisWorkbook.VBProject.VBComponents.Remove comp
            End If
        End If
        On Error GoTo 0
        
        '   ߰
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Import filePath
        If Err.Number <> 0 Then
            MsgBox fileName & "  : " & Err.Description
        Else
            MsgBox fileName & "  ƮǾϴ!"
        End If
        On Error GoTo 0
    Else
        MsgBox fileName & " Ʈ : " & http.Status
    End If
End Sub

