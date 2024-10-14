Attribute VB_Name = "Ʈ_"
Sub LoadTreeViewWithData1()
    Dim ws As Worksheet
    Dim treeView As Object
    Dim lastRow As Long
    Dim i As Long
    Dim parentNode As Object
    Dim currentDate As Variant
    Application.ScreenUpdating = False
    
    ' "Ƿ" Ʈ ãϴ. Ʈ ̸ ° ּ.
    Set ws = ThisWorkbook.Sheets("Ƿ")
    
    ' Ʈ並 ߰մϴ. "MSComctlLib.TreeCtrl"  ߰ؾ մϴ.
    Set treeView = UserForm1.TreeView1

    ' Ʈ ʱȭ
    treeView.Nodes.Clear
    
    '   ãϴ.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' ¥   
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
    
    '  ¥ θ  ڽ 带 ߰
    For i = 2 To lastRow
        '  ¥ ɴϴ.
        currentDate = ws.Cells(i, 1).Value
        
        ' θ 尡 ų  ¥  ¥ ٸ  ο θ 带 ߰
        If parentNode Is Nothing Or currentDate <> ws.Cells(i - 1, 1).Value Then
            Set parentNode = treeView.Nodes.Add(, , "ParentKey" & i, currentDate)
            parentNode.ForeColor = RGB(0, 128, 0)
        End If
        
        Dim childNode As Object
        Set childNode = treeView.Nodes.Add(parentNode.index, tvwChild, "ChildKey" & i, "" & ws.Cells(i, "E").text & "" & ws.Cells(i, "F").Value)
        ' ڽ  ؽƮ  Ķ 
        childNode.ForeColor = RGB(0, 0, 128)
    Next i
    
    
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' ó 20 θ 常  · 
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
    

    
    ' TreeView2 Ʈ 
    Set treeView = UserForm1.TreeView2
    
    ' Ʈ ʱȭ
    treeView.Nodes.Clear
    
    ' "ܰ" Ʈ 
    Set ws = ThisWorkbook.Sheets("ܰ")

    ' C    ã
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' ųʸ ʱȭ
    Set parentDict = CreateObject("Scripting.Dictionary")

    ' C2:C41 θ  ߰  ߺ ó
    For i = 2 To lastRow
        ParentItem = ws.Cells(i, 3).Value

        ' θ 尡 ̹ ųʸ ִ Ȯ
        If Not parentDict.exists(ParentItem) Then
            ' θ 尡  ߰
            Set parentNode = treeView.Nodes.Add(, , "P" & i, ParentItem)
            parentDict.Add ParentItem, parentNode
        Else
            ' θ 尡  
            Set parentNode = parentDict(ParentItem)
        End If

        ' D2:D42   ߰
        childItem = ws.Cells(i, 4).Value
        If childItem <> "" Then
            treeView.Nodes.Add parentNode, tvwChild, "C" & i, childItem
        End If
    Next i

    ' "Ưع" θ  ߰
    Dim specificNode As MSComctlLib.Node
    If Not parentDict.exists("Ưع") Then
        Set specificNode = treeView.Nodes.Add(, , "P_Special", "Ưع")
        parentDict.Add "Ưع", specificNode
    Else
        Set specificNode = parentDict("Ưع")
    End If
    
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' ó 20 θ 常  · 
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
    
    ' "Ƿ" Ʈ ãϴ. Ʈ ̸ ° ּ.
    Set ws = ThisWorkbook.Sheets("")
    
    ' Ʈ並 ߰մϴ. "MSComctlLib.TreeCtrl"  ߰ؾ մϴ.
    Set treeView = UserForm1.TreeView3
    
    ' Ʈ ʱȭ
    treeView.Nodes.Clear
    
    '   ãϴ.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' ¥   
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
    
    '  ¥ θ  ڽ 带 ߰
    For i = 3 To lastRow
        '  ¥ ɴϴ.
        currentDate = ws.Cells(i, 1).Value
        
        ' θ 尡 ų  ¥  ¥ ٸ  ο θ 带 ߰
        If parentNode Is Nothing Or currentDate <> ws.Cells(i - 1, 1).Value Then
            Set parentNode = treeView.Nodes.Add(, , "ParentKey" & i, currentDate)
            parentNode.ForeColor = RGB(0, 128, 0)
        End If
        
        Dim childNode As Object
        Set childNode = treeView.Nodes.Add(parentNode.index, tvwChild, "ChildKey" & i, "" & ws.Cells(i, "C").text & "" & ws.Cells(i, "H").Value)
        ' ڽ  ؽƮ  Ķ 
        childNode.ForeColor = RGB(0, 0, 128)
    Next i
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' ó 20 θ 常  · 
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
    
    ' "Ƿ" Ʈ ãϴ. Ʈ ̸ ° ּ.
    Set ws = ThisWorkbook.Sheets("Ƿ")
    
    ' Ʈ並 ߰մϴ. "MSComctlLib.TreeCtrl"  ߰ؾ մϴ.
    Set treeView = UserForm1.TreeView4

    ' Ʈ ʱȭ
    treeView.Nodes.Clear
    
    '   ãϴ.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' ¥   
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
    
    '  ¥ θ  ڽ 带 ߰
    For i = 2 To lastRow
        '  ¥ ɴϴ.
        currentDate = ws.Cells(i, 1).Value
        
        ' θ 尡 ų  ¥  ¥ ٸ  ο θ 带 ߰
        If parentNode Is Nothing Or currentDate <> ws.Cells(i - 1, 1).Value Then
            Set parentNode = treeView.Nodes.Add(, , "ParentKey" & i, currentDate)
            parentNode.ForeColor = RGB(0, 128, 0)
        End If
        
        Dim childNode As Object
        Set childNode = treeView.Nodes.Add(parentNode.index, tvwChild, "ChildKey" & i, "" & ws.Cells(i, "E").text & "" & ws.Cells(i, "F").Value)
        ' ڽ  ؽƮ  Ķ 
        childNode.ForeColor = RGB(0, 0, 128)
    Next i
    
    
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' ó 20 θ 常  · 
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
    
    ' "Ƿ" Ʈ ãϴ. Ʈ ̸ ° ּ.
    Set ws = ThisWorkbook.Sheets("Ƿ")
    
    ' Ʈ並 ߰մϴ. "MSComctlLib.TreeCtrl"  ߰ؾ մϴ.
    Set treeView = UserForm1.TreeView5

    ' Ʈ ʱȭ
    treeView.Nodes.Clear
    
    '   ãϴ.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' ¥   
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
    
    '  ¥ θ  ڽ 带 ߰
    For i = 2 To lastRow
        '  ¥ ɴϴ.
        currentDate = ws.Cells(i, 1).Value
        
        ' θ 尡 ų  ¥  ¥ ٸ  ο θ 带 ߰
        If parentNode Is Nothing Or currentDate <> ws.Cells(i - 1, 1).Value Then
            Set parentNode = treeView.Nodes.Add(, , "ParentKey" & i, currentDate)
            parentNode.ForeColor = RGB(0, 128, 0)
        End If
        
        Dim childNode As Object
        Set childNode = treeView.Nodes.Add(parentNode.index, tvwChild, "ChildKey" & i, "" & ws.Cells(i, "E").text & "" & ws.Cells(i, "F").Value)
        ' ڽ  ؽƮ  Ķ 
        childNode.ForeColor = RGB(0, 0, 128)
    Next i
    
    
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' ó 20 θ 常  · 
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
    
    ' "Ƿ" Ʈ ãϴ. Ʈ ̸ ° ּ.
    Set ws = ThisWorkbook.Sheets("")
    
    ' Ʈ並 ߰մϴ. "MSComctlLib.TreeCtrl"  ߰ؾ մϴ.
    Set treeView = UserForm1.TreeView6
    
    ' Ʈ ʱȭ
    treeView.Nodes.Clear
    
    '   ãϴ.
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' ¥   
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
    
    '  ¥ θ  ڽ 带 ߰
    For i = 3 To lastRow
        '  ¥ ɴϴ.
        currentDate = ws.Cells(i, 1).Value
        
        ' θ 尡 ų  ¥  ¥ ٸ  ο θ 带 ߰
        If parentNode Is Nothing Or currentDate <> ws.Cells(i - 1, 1).Value Then
            Set parentNode = treeView.Nodes.Add(, , "ParentKey" & i, currentDate)
            parentNode.ForeColor = RGB(0, 128, 0)
        End If
        
        Dim childNode As Object
        Set childNode = treeView.Nodes.Add(parentNode.index, tvwChild, "ChildKey" & i, "" & ws.Cells(i, "C").text & "" & ws.Cells(i, "H").Value)
        ' ڽ  ؽƮ  Ķ 
        childNode.ForeColor = RGB(0, 0, 128)
    Next i
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' ó 20 θ 常  · 
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
    

    
    ' TreeView2 Ʈ 
    Set treeView = UserForm1.TreeView7
    
    ' Ʈ ʱȭ
    treeView.Nodes.Clear
    
    ' "ܰ" Ʈ 
    Set ws = ThisWorkbook.Sheets("ܰ")

    ' C    ã
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' ųʸ ʱȭ
    Set parentDict = CreateObject("Scripting.Dictionary")

    ' C2:C41 θ  ߰  ߺ ó
    For i = 2 To lastRow
        ParentItem = ws.Cells(i, 3).Value

        ' θ 尡 ̹ ųʸ ִ Ȯ
        If Not parentDict.exists(ParentItem) Then
            ' θ 尡  ߰
            Set parentNode = treeView.Nodes.Add(, , "P" & i, ParentItem)
            parentDict.Add ParentItem, parentNode
        Else
            ' θ 尡  
            Set parentNode = parentDict(ParentItem)
        End If

        ' D2:D42   ߰
        childItem = ws.Cells(i, 4).Value
        If childItem <> "" Then
            treeView.Nodes.Add parentNode, tvwChild, "C" & i, childItem
        End If
    Next i


    
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' ó 20 θ 常  · 
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
    
    ' ó 20 θ   ·,     · 
    For i = 1 To treeView.Nodes.Count
        If i <= 20 Then
            treeView.Nodes(i).Expanded = True
        Else
            treeView.Nodes(i).Expanded = False
        End If
    Next i

    '  带  · 
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' ó 20 θ 常  · 
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
    
    

    ' TreeView2 Ʈ 
    Set treeView = UserForm1.TreeView2

    ' "ܰ" Ʈ 
    Set ws = ThisWorkbook.Sheets("ܰ")

    ' C    ã
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' ųʸ ʱȭ
    Set parentDict = CreateObject("Scripting.Dictionary")

    ' C2:C41 θ  ߰  ߺ ó
    For i = 2 To lastRow
        ParentItem = ws.Cells(i, 3).Value

        ' θ 尡 ̹ ųʸ ִ Ȯ
        If Not parentDict.exists(ParentItem) Then
            ' θ 尡  ߰
            Set parentNode = treeView.Nodes.Add(, , "P" & i, ParentItem)
            parentDict.Add ParentItem, parentNode
        Else
            ' θ 尡  
            Set parentNode = parentDict(ParentItem)
        End If

        ' D2:D42   ߰
        childItem = ws.Cells(i, 4).Value
        If childItem <> "" Then
            treeView.Nodes.Add parentNode, tvwChild, "C" & i, childItem
        End If
    Next i

    ' "Ưع" θ  ߰
    Dim specificNode As MSComctlLib.Node
    If Not parentDict.exists("Ưع") Then
        Set specificNode = treeView.Nodes.Add(, , "P_Special", "Ưع")
        parentDict.Add "Ưع", specificNode
    Else
        Set specificNode = parentDict("Ưع")
    End If

    ' Ư  Ʒ   ߰ 
    ' ⿡   ߰ ڵ带 ۼϼ
'    treeView.Nodes.Add specificNode, tvwChild, "C_Special1", "1"
'    treeView.Nodes.Add specificNode, tvwChild, "C_Special2", "2"

    '  带  · 
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' ó 20 θ 常  · 
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
    
        '  带  · 
    For i = 1 To treeView.Nodes.Count
        treeView.Nodes(i).Expanded = False
    Next i

    ' ó 20 θ 常  · 
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

    ' TreeView ʱȭ
    UserForm1.TreeView4.Nodes.Clear
    
    ' Ʈ COMBOBOX5  
    Set ws = ThisWorkbook.Sheets("Ƿ")
    comboValue = UserForm1.ComboBox5.Value
    
    ' Dictionary ü ʱȭ (Scripting.Dictionary )
    Set dict = CreateObject("Scripting.Dictionary")
    

    
    
    
    ' Ƿ Ʈ   ãϴ.
    If UserForm1.ComboBox6.Value <> "üⰣ -_-'" Then


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
    ' ޺ڽ  ִ Ȯմϴ.
    If comboValue = "" Then
        MsgBox "޺ڽ   ֽϴ.", vbExclamation
        Exit Sub
    End If
    
    ' Ƿ Ʈ ޺ڽ  ġϴ ׸ ãϴ.
    For i = 2 To lastRow ' Assuming headers are in the first row
        If ws.Cells(i, 4).Value = comboValue Then
            ' Dictionary Ͽ ̹ ߰ ׸ Ȯ
            If Not dict.exists(ws.Cells(i, 6).Value) Then
                ' ׸ ߰ ʾ TreeView ߰ϰ, Dictionary ߰
                UserForm1.TreeView4.Nodes.Add , , , ws.Cells(i, 6).Value
                dict.Add ws.Cells(i, 6).Value, Nothing
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub Ʈ3Ŭüũ()
    Dim r As Integer
    Dim X As Node
    Dim NodeText As String

    '   ڻ  ʱȭ
    For r = 1 To UserForm1.TreeView2.Nodes.Count
        UserForm1.TreeView2.Nodes(r).ForeColor = RGB(0, 0, 0)
    Next r

    ' ListBox1 ׸ ˻Ͽ شϴ  ڻ  
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
