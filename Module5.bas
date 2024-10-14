Attribute VB_Name = "Module5"
癤풞ttribute VB_Name = "UPDATE"
Sub ExportModulesAndUserForms()
    Dim vbComp As Object
    Dim exportPath As String
    Dim fileName As String
    
    '
    exportPath = "C:\CENTER\" ' 求 管 究

    ' 寬 효 확
    If Right(exportPath, 1) <> "\" Then
        exportPath = exportPath & "\"
    End If

    '
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_Module
                fileName = exportPath & vbComp.Name & ".bas"
                vbComp.Export fileName
                MsgBox vbComp.Name & "  " & fileName & " 求."
            Case vbext_ct_MSForm
                fileName = exportPath & vbComp.Name & ".frm"
                vbComp.Export fileName
                MsgBox vbComp.Name & "  " & fileName & " 求."
        End Select
    Next vbComp

    MsgBox "   求."
End Sub


Sub UpdateAllModulesAndFormsFromGitHub()
    Dim fileList As Variant
    Dim i As Integer
    '??
    ' 트
    fileList = Array( _
        "Combo.bas", "Module1.bas", "Module2.bas", "Module3.bas", "UPDATE.bas", "Module4.bas", "TESTModule.bas", _
        "UserForm1.frm", "UserForm1.frx", "韜歐.bas", ".bas", "留.bas", _
        "羞.bas", "車뺑.bas", "硫.bas", "트_.bas", "_찾튼.bas", _
        "莫탐챰.bas", "挽擔.bas", "캅탐챰.bas", "劇韜_.bas", _
        "劇회.bas", "쨔.bas", "寗.bas", "瞿.bas", _
        "퓐美트.bas", "歐.bas", "트_.bas", "특.bas", "卵.bas")
    
    '   牟琯構 트
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
    
    '  확
    componentType = Right(fileName, 3)
    
    ' 轅   (Raw URL )
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "https://raw.githubusercontent.com/1110x/Center/main/" & fileName
    http.Open "GET", url, False
    http.Send
    
    ' HTTP 청   처
    If http.Status = 200 Then
        fileData = http.responseText
        filePath = "C:\Users\ironu\OneDrive\ 화\" & fileName ' 罐
        
        ' ADODB.Stream 체 臼 UTF-8
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2 ' 灣트
        stream.Charset = "utf-8" ' UTF-8
        stream.Open
        stream.WriteText fileData
        stream.SaveToFile filePath, 2 ' 2 杵
        stream.Close
        
        '  VBA
        On Error Resume Next
        If componentType = "bas" Or componentType = "frm" Then
            ' VBProject 求 확
            If ThisWorkbook.VBProject Is Nothing Then
                MsgBox "VBProject 찾  求."
                Exit Sub
            End If
            
            ' 柰 求 확
            Dim comp As Object
            Set comp = ThisWorkbook.VBProject.VBComponents(Left(fileName, Len(fileName) - 4))
            If Not comp Is Nothing Then
                ThisWorkbook.VBProject.VBComponents.Remove comp
            End If
        End If
        On Error GoTo 0
        
        '   煞
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Import filePath
        If Err.Number <> 0 Then
            MsgBox fileName & "  : " & Err.Description
        Else
            MsgBox fileName & "  트퓸求!"
        End If
        On Error GoTo 0
    Else
        MsgBox fileName & " 트 : " & http.Status
    End If
End Sub

