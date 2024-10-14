Attribute VB_Name = "트리뷰_세팅"
Sub LoadTreeViewWithData1()
    Dim ws As Worksheet
    Dim treeView As Object
    Dim lastRow As Long
    Dim i As Long
    Dim parentNode As Object
    Dim currentDate As Variant
    Application.ScreenUpdating = False
    
    ' "의뢰정보" 시트를 찾습니다. 시트 이름에 맞게 수정해주세요.
    Set ws = ThisWorkbook.Sheets("의뢰정보")
    
    ' 트리뷰를 추가합니다. "MSComctlLib.TreeCtrl"를 참조 추가해야 합니다.
    Set treeView = UserForm1.TreeView1

    ' 트리뷰 초기화
    treeView.Nodes.Clear
    
    ' 마지막 행을 찾습니다.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 날짜를 기준으로 내림차순 정렬
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Range("A1:BZ" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 각 날짜별로 부모 노드와 자식 노드를 추가
    For i = 2 To lastRow
        ' 현재 날짜를 가져옵니다.
        currentDate = ws.Cells(i, 1).Value
        
        ' 부모 노드가 없거나 현재 날짜와 이전 날짜가 다를 경우 새로운 부모 노드를 추가
        If parentNode Is Nothing Or currentDate <> ws.Cells(i - 1, 1).Value Then
            Set parentNode = treeView.Nodes.Add(, , "ParentKey" & i, currentDate)
            parentNode.ForeColor = RGB(0, 128, 0)
        End If
        
        Dim childNode As Object
        Set childNode = treeView.Nodes.Add(parentNode.index, tvwChild, "ChildKey" & i, "【" & ws.Cells(i, "E").text & "】" & ws.Cells(i, "F").Value)
        ' 자식 노드 텍스트의 색상을 파란색으로 설정
        childNode.ForeColor = RGB(0, 0, 128)
    Next i
    
    
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' 처음 20개의 부모 노드만 펼쳐진 상태로 설정
    For i = 1 To 60
        If treeView.Nodes(i).Children > 0 Then
            treeView.Nodes(i).Expanded = True
        End If
    Next i

    

Application.ScreenUpdating = True

        
End Sub


Sub LoadTreeViewWithData3()
    Dim ws As Worksheet
    Dim treeView As Object
    Dim lastRow As Long
    Dim i As Long
    Dim parentNode As Object
    Dim currentDate As Variant
    Application.ScreenUpdating = False
    

    
    ' TreeView2 컨트롤 참조
    Set treeView = UserForm1.TreeView2
    
    ' 트리뷰 초기화
    treeView.Nodes.Clear
    
    ' "견적단가" 시트 참조
    Set ws = ThisWorkbook.Sheets("견적단가")

    ' C 열의 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' 딕셔너리 초기화
    Set parentDict = CreateObject("Scripting.Dictionary")

    ' C2:C41의 부모 노드 추가 및 중복 처리
    For i = 2 To lastRow
        ParentItem = ws.Cells(i, 3).Value

        ' 부모 노드가 이미 딕셔너리에 있는지 확인
        If Not parentDict.exists(ParentItem) Then
            ' 부모 노드가 없으면 추가
            Set parentNode = treeView.Nodes.Add(, , "P" & i, ParentItem)
            parentDict.Add ParentItem, parentNode
        Else
            ' 부모 노드가 있으면 가져오기
            Set parentNode = parentDict(ParentItem)
        End If

        ' D2:D42의 하위 노드 추가
        childItem = ws.Cells(i, 4).Value
        If childItem <> "" Then
            treeView.Nodes.Add parentNode, tvwChild, "C" & i, childItem
        End If
    Next i

    ' "특정수질유해물질" 부모 노드 추가
    Dim specificNode As MSComctlLib.Node
    If Not parentDict.exists("특정수질유해물질") Then
        Set specificNode = treeView.Nodes.Add(, , "P_Special", "특정수질유해물질")
        parentDict.Add "특정수질유해물질", specificNode
    Else
        Set specificNode = parentDict("특정수질유해물질")
    End If
    
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' 처음 20개의 부모 노드만 펼쳐진 상태로 설정
    For i = 1 To 60
        If treeView.Nodes(i).Children > 0 Then
            treeView.Nodes(i).Expanded = True
        End If
    Next i

'
End Sub
Sub LoadTreeViewWithData2()
    Dim ws As Worksheet
    Dim treeView As Object
    Dim lastRow As Long
    Dim i As Long
    Dim parentNode As Object
    Dim currentDate As Variant
    Application.ScreenUpdating = False
    
    ' "의뢰정보" 시트를 찾습니다. 시트 이름에 맞게 수정해주세요.
    Set ws = ThisWorkbook.Sheets("견적발행정보")
    
    ' 트리뷰를 추가합니다. "MSComctlLib.TreeCtrl"를 참조 추가해야 합니다.
    Set treeView = UserForm1.TreeView3
    
    ' 트리뷰 초기화
    treeView.Nodes.Clear
    
    ' 마지막 행을 찾습니다.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 날짜를 기준으로 내림차순 정렬
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range("A3:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Range("A2:ZZ" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 각 날짜별로 부모 노드와 자식 노드를 추가
    For i = 3 To lastRow
        ' 현재 날짜를 가져옵니다.
        currentDate = ws.Cells(i, 1).Value
        
        ' 부모 노드가 없거나 현재 날짜와 이전 날짜가 다를 경우 새로운 부모 노드를 추가
        If parentNode Is Nothing Or currentDate <> ws.Cells(i - 1, 1).Value Then
            Set parentNode = treeView.Nodes.Add(, , "ParentKey" & i, currentDate)
            parentNode.ForeColor = RGB(0, 128, 0)
        End If
        
        Dim childNode As Object
        Set childNode = treeView.Nodes.Add(parentNode.index, tvwChild, "ChildKey" & i, "【" & ws.Cells(i, "C").text & "】" & ws.Cells(i, "H").Value)
        ' 자식 노드 텍스트의 색상을 파란색으로 설정
        childNode.ForeColor = RGB(0, 0, 128)
    Next i
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' 처음 20개의 부모 노드만 펼쳐진 상태로 설정
    For i = 1 To 60
        If treeView.Nodes(i).Children > 0 Then
            treeView.Nodes(i).Expanded = True
        End If
    Next i

   Application.ScreenUpdating = True
   
End Sub
Sub LoadTreeViewWithData4()
    Dim ws As Worksheet
    Dim treeView As Object
    Dim lastRow As Long
    Dim i As Long
    Dim parentNode As Object
    Dim currentDate As Variant
    Application.ScreenUpdating = False
    
    ' "의뢰정보" 시트를 찾습니다. 시트 이름에 맞게 수정해주세요.
    Set ws = ThisWorkbook.Sheets("의뢰정보")
    
    ' 트리뷰를 추가합니다. "MSComctlLib.TreeCtrl"를 참조 추가해야 합니다.
    Set treeView = UserForm1.TreeView4

    ' 트리뷰 초기화
    treeView.Nodes.Clear
    
    ' 마지막 행을 찾습니다.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 날짜를 기준으로 내림차순 정렬
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Range("A1:BZ" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 각 날짜별로 부모 노드와 자식 노드를 추가
    For i = 2 To lastRow
        ' 현재 날짜를 가져옵니다.
        currentDate = ws.Cells(i, 1).Value
        
        ' 부모 노드가 없거나 현재 날짜와 이전 날짜가 다를 경우 새로운 부모 노드를 추가
        If parentNode Is Nothing Or currentDate <> ws.Cells(i - 1, 1).Value Then
            Set parentNode = treeView.Nodes.Add(, , "ParentKey" & i, currentDate)
            parentNode.ForeColor = RGB(0, 128, 0)
        End If
        
        Dim childNode As Object
        Set childNode = treeView.Nodes.Add(parentNode.index, tvwChild, "ChildKey" & i, "【" & ws.Cells(i, "E").text & "】" & ws.Cells(i, "F").Value)
        ' 자식 노드 텍스트의 색상을 파란색으로 설정
        childNode.ForeColor = RGB(0, 0, 128)
    Next i
    
    
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' 처음 20개의 부모 노드만 펼쳐진 상태로 설정
    For i = 1 To 60
        If treeView.Nodes(i).Children > 0 Then
            treeView.Nodes(i).Expanded = True
        End If
    Next i

    

Application.ScreenUpdating = True

        
End Sub
Sub LoadTreeViewWithData5()
    Dim ws As Worksheet
    Dim treeView As Object
    Dim lastRow As Long
    Dim i As Long
    Dim parentNode As Object
    Dim currentDate As Variant
    Application.ScreenUpdating = False
    
    ' "의뢰정보" 시트를 찾습니다. 시트 이름에 맞게 수정해주세요.
    Set ws = ThisWorkbook.Sheets("의뢰정보")
    
    ' 트리뷰를 추가합니다. "MSComctlLib.TreeCtrl"를 참조 추가해야 합니다.
    Set treeView = UserForm1.TreeView5

    ' 트리뷰 초기화
    treeView.Nodes.Clear
    
    ' 마지막 행을 찾습니다.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 날짜를 기준으로 내림차순 정렬
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Range("A1:BZ" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 각 날짜별로 부모 노드와 자식 노드를 추가
    For i = 2 To lastRow
        ' 현재 날짜를 가져옵니다.
        currentDate = ws.Cells(i, 1).Value
        
        ' 부모 노드가 없거나 현재 날짜와 이전 날짜가 다를 경우 새로운 부모 노드를 추가
        If parentNode Is Nothing Or currentDate <> ws.Cells(i - 1, 1).Value Then
            Set parentNode = treeView.Nodes.Add(, , "ParentKey" & i, currentDate)
            parentNode.ForeColor = RGB(0, 128, 0)
        End If
        
        Dim childNode As Object
        Set childNode = treeView.Nodes.Add(parentNode.index, tvwChild, "ChildKey" & i, "【" & ws.Cells(i, "E").text & "】" & ws.Cells(i, "F").Value)
        ' 자식 노드 텍스트의 색상을 파란색으로 설정
        childNode.ForeColor = RGB(0, 0, 128)
    Next i
    
    
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' 처음 20개의 부모 노드만 펼쳐진 상태로 설정
    For i = 1 To 60
        If treeView.Nodes(i).Children > 0 Then
            treeView.Nodes(i).Expanded = True
        End If
    Next i

    

Application.ScreenUpdating = True

        
End Sub
Sub LoadTreeViewWithData6()
    Dim ws As Worksheet
    Dim treeView As Object
    Dim lastRow As Long
    Dim i As Long
    Dim parentNode As Object
    Dim currentDate As Variant
    Application.ScreenUpdating = False
    
    ' "의뢰정보" 시트를 찾습니다. 시트 이름에 맞게 수정해주세요.
    Set ws = ThisWorkbook.Sheets("견적발행정보")
    
    ' 트리뷰를 추가합니다. "MSComctlLib.TreeCtrl"를 참조 추가해야 합니다.
    Set treeView = UserForm1.TreeView6
    
    ' 트리뷰 초기화
    treeView.Nodes.Clear
    
    ' 마지막 행을 찾습니다.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' 날짜를 기준으로 내림차순 정렬
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=Range("A3:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ws.Sort
        .SetRange Range("A2:ZZ" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 각 날짜별로 부모 노드와 자식 노드를 추가
    For i = 3 To lastRow
        ' 현재 날짜를 가져옵니다.
        currentDate = ws.Cells(i, 1).Value
        
        ' 부모 노드가 없거나 현재 날짜와 이전 날짜가 다를 경우 새로운 부모 노드를 추가
        If parentNode Is Nothing Or currentDate <> ws.Cells(i - 1, 1).Value Then
            Set parentNode = treeView.Nodes.Add(, , "ParentKey" & i, currentDate)
            parentNode.ForeColor = RGB(0, 128, 0)
        End If
        
        Dim childNode As Object
        Set childNode = treeView.Nodes.Add(parentNode.index, tvwChild, "ChildKey" & i, "【" & ws.Cells(i, "C").text & "】" & ws.Cells(i, "H").Value)
        ' 자식 노드 텍스트의 색상을 파란색으로 설정
        childNode.ForeColor = RGB(0, 0, 128)
    Next i
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' 처음 20개의 부모 노드만 펼쳐진 상태로 설정
    For i = 1 To 60
        If treeView.Nodes(i).Children > 0 Then
            treeView.Nodes(i).Expanded = True
        End If
    Next i

   Application.ScreenUpdating = True
   
End Sub
Sub LoadTreeViewWithData7()
    Dim ws As Worksheet
    Dim treeView As Object
    Dim lastRow As Long
    Dim i As Long
    Dim parentNode As Object
    Dim currentDate As Variant
    Application.ScreenUpdating = False
    

    
    ' TreeView2 컨트롤 참조
    Set treeView = UserForm1.TreeView7
    
    ' 트리뷰 초기화
    treeView.Nodes.Clear
    
    ' "견적단가" 시트 참조
    Set ws = ThisWorkbook.Sheets("견적단가")

    ' C 열의 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' 딕셔너리 초기화
    Set parentDict = CreateObject("Scripting.Dictionary")

    ' C2:C41의 부모 노드 추가 및 중복 처리
    For i = 2 To lastRow
        ParentItem = ws.Cells(i, 3).Value

        ' 부모 노드가 이미 딕셔너리에 있는지 확인
        If Not parentDict.exists(ParentItem) Then
            ' 부모 노드가 없으면 추가
            Set parentNode = treeView.Nodes.Add(, , "P" & i, ParentItem)
            parentDict.Add ParentItem, parentNode
        Else
            ' 부모 노드가 있으면 가져오기
            Set parentNode = parentDict(ParentItem)
        End If

        ' D2:D42의 하위 노드 추가
        childItem = ws.Cells(i, 4).Value
        If childItem <> "" Then
            treeView.Nodes.Add parentNode, tvwChild, "C" & i, childItem
        End If
    Next i


    
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' 처음 20개의 부모 노드만 펼쳐진 상태로 설정
    For i = 1 To 60
        If treeView.Nodes(i).Children > 0 Then
            treeView.Nodes(i).Expanded = True
        End If
    Next i

'
End Sub



Sub SetInitialNodeStates1()
    Dim i As Long
    Dim treeView As MSComctlLib.treeView
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim parentDict As Object
    Dim parentNode As MSComctlLib.Node
    Dim childNode As MSComctlLib.Node
    Dim ParentItem As Variant
    Dim childItem As Variant
    Dim parentNodesCount As Integer
    
    Set treeView = UserForm1.TreeView1
    
    ' 처음부터 20개의 부모 노드는 펼쳐진 상태로, 그 이후의 노드는 닫힌 상태로 설정
    For i = 1 To treeView.Nodes.Count
        If i <= 20 Then
            treeView.Nodes(i).Expanded = True
        Else
            treeView.Nodes(i).Expanded = False
        End If
    Next i

    ' 모든 노드를 닫힌 상태로 설정
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' 처음 20개의 부모 노드만 펼쳐진 상태로 설정
    For i = 1 To 60
        If treeView.Nodes(i).Children > 0 Then
            treeView.Nodes(i).Expanded = True
        End If
    Next i






End Sub

Sub SetInitialNodeStates2()
    Dim i As Long
    Dim treeView As MSComctlLib.treeView
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim parentDict As Object
    Dim parentNode As MSComctlLib.Node
    Dim childNode As MSComctlLib.Node
    Dim ParentItem As Variant
    Dim childItem As Variant
    Dim parentNodesCount As Integer
    
    

    ' TreeView2 컨트롤 참조
    Set treeView = UserForm1.TreeView2

    ' "견적단가" 시트 참조
    Set ws = ThisWorkbook.Sheets("견적단가")

    ' C 열의 마지막 행 찾기
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' 딕셔너리 초기화
    Set parentDict = CreateObject("Scripting.Dictionary")

    ' C2:C41의 부모 노드 추가 및 중복 처리
    For i = 2 To lastRow
        ParentItem = ws.Cells(i, 3).Value

        ' 부모 노드가 이미 딕셔너리에 있는지 확인
        If Not parentDict.exists(ParentItem) Then
            ' 부모 노드가 없으면 추가
            Set parentNode = treeView.Nodes.Add(, , "P" & i, ParentItem)
            parentDict.Add ParentItem, parentNode
        Else
            ' 부모 노드가 있으면 가져오기
            Set parentNode = parentDict(ParentItem)
        End If

        ' D2:D42의 하위 노드 추가
        childItem = ws.Cells(i, 4).Value
        If childItem <> "" Then
            treeView.Nodes.Add parentNode, tvwChild, "C" & i, childItem
        End If
    Next i

    ' "특정수질유해물질" 부모 노드 추가
    Dim specificNode As MSComctlLib.Node
    If Not parentDict.exists("특정수질유해물질") Then
        Set specificNode = treeView.Nodes.Add(, , "P_Special", "특정수질유해물질")
        parentDict.Add "특정수질유해물질", specificNode
    Else
        Set specificNode = parentDict("특정수질유해물질")
    End If

    ' 특정 노드 아래에 하위 노드 추가 예시
    ' 여기에 하위 노드 추가 코드를 작성하세요
'    treeView.Nodes.Add specificNode, tvwChild, "C_Special1", "하위노드1"
'    treeView.Nodes.Add specificNode, tvwChild, "C_Special2", "하위노드2"

    ' 모든 노드를 닫힌 상태로 설정
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' 처음 20개의 부모 노드만 펼쳐진 상태로 설정
    parentNodesCount = 0
    For Each ParentItem In parentDict.keys
        If parentNodesCount < 20 Then
            parentDict(ParentItem).Expanded = True
            parentNodesCount = parentNodesCount + 1
        Else
            Exit For
        End If
    Next ParentItem

    
    Set treeView = UserForm1.TreeView3
    
        ' 모든 노드를 닫힌 상태로 설정
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' 처음 20개의 부모 노드만 펼쳐진 상태로 설정
    parentNodesCount = 0
    For Each ParentItem In parentDict.keys
        If parentNodesCount < 20 Then
            parentDict(ParentItem).Expanded = True
            parentNodesCount = parentNodesCount + 1
        Else
            Exit For
        End If
    Next ParentItem
    
End Sub
Sub AddTreeViewItem()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim comboValue As String
    Dim dict As Object
    Dim i As Long
    

    Application.ScreenUpdating = False

    ' TreeView를 초기화
    UserForm1.TreeView4.Nodes.Clear
    
    ' 시트와 COMBOBOX5의 값을 설정
    Set ws = ThisWorkbook.Sheets("의뢰정보")
    comboValue = UserForm1.ComboBox5.Value
    
    ' Dictionary 객체 초기화 (Scripting.Dictionary 사용)
    Set dict = CreateObject("Scripting.Dictionary")
    

    
    
    
    ' 의뢰정보 시트의 마지막 행을 찾습니다.
    If UserForm1.ComboBox6.Value <> "전체기간 -_-'" Then


       For Z = ws.Cells(ws.Rows.Count, 4).End(xlUp).row To 2 Step (-1)
         If Left(ws.Cells(Z, 1).Value, 4) = UserForm1.ComboBox6.Value Then
         lastRow = Z
         Exit For
         End If
       Next Z '
    Else

       lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).row
    End If
    
    
'    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    ' 콤보박스 값이 비어있는지 확인합니다.
    If comboValue = "" Then
        MsgBox "콤보박스의 값이 비어 있습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 의뢰정보 시트에서 콤보박스 값과 일치하는 항목을 찾습니다.
    For i = 2 To lastRow ' Assuming headers are in the first row
        If ws.Cells(i, 4).Value = comboValue Then
            ' Dictionary를 사용하여 이미 추가된 항목인지 확인
            If Not dict.exists(ws.Cells(i, 6).Value) Then
                ' 항목이 추가되지 않았으면 TreeView에 추가하고, Dictionary에 추가
                UserForm1.TreeView4.Nodes.Add , , , ws.Cells(i, 6).Value
                dict.Add ws.Cells(i, 6).Value, Nothing
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub 트리뷰3클릭체크()
    Dim r As Integer
    Dim X As Node
    Dim NodeText As String

    ' 모든 노드의 글자색을 검정으로 초기화
    For r = 1 To UserForm1.TreeView2.Nodes.Count
        UserForm1.TreeView2.Nodes(r).ForeColor = RGB(0, 0, 0)
    Next r

    ' ListBox1의 항목을 검사하여 해당하는 노드의 글자색을 빨간색으로 변경
    For r = 0 To UserForm1.ListBox1.ListCount - 1
        For Each X In UserForm1.TreeView2.Nodes
            NodeText = X.text
            If UserForm1.ListBox1.List(r, 1) = NodeText Then
                X.ForeColor = RGB(255, 0, 0)
                Exit For
            End If
        Next X
    Next r

End Sub
