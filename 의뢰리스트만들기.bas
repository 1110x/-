Attribute VB_Name = "ǷڸƮ"
'Global SR As Integer

Sub ǷڸƮ̵()
    Dim X As String
    Dim i As Integer
    Dim IsDuplicate As Boolean
    
    X = UserForm1.TreeView4.SelectedItem.text
    
    ' ߺ θ Ȯմϴ.
    IsDuplicate = False
    For i = 0 To UserForm1.ListBox2.ListCount - 1
        If UserForm1.ListBox2.List(i, 1) = X Then
            IsDuplicate = True
            Exit For
        End If
    Next i
    
    ' ߺ ʴ 쿡 ׸ ߰մϴ.
    If Not IsDuplicate Then
        UserForm1.ListBox2.AddItem "" '  ׸ ߰Ͽ ο  
        ' ListBox  ߰  Ư   
        UserForm1.ListBox2.List(UserForm1.ListBox2.ListCount - 1, 0) = Format(UserForm1.ListBox2.ListCount, "00")
        UserForm1.ListBox2.List(UserForm1.ListBox2.ListCount - 1, 1) = X
    Else
        Exit Sub
        
    End If
End Sub

Sub Ƿ׸üũ()
Application.ScreenUpdating = False
        Set ws = ThisWorkbook.Sheets("")
        lastRow = Sheets("").Cells(Sheets("").Rows.Count, "A").End(xlUp).row
        For r = 2 To lastRow
         Z = "" & Sheets("").Cells(r, "C").text & "" & Sheets("").Cells(r, "H").Value '  ws.Cells(i, "C")  & ws.Cells(i, 4).Value)
         
         
                If Sheets("").Cells(r, "A") = Sheets("мǷ Է").Cells(1, "F") And Z = Sheets("мǷ Է").Cells(1, "B") Then
                      Debug.Print Z
                       Sheets("мǷ Է").Cells(1, "A") = Sheets("").Cells(r, "C")
                       Set Ī = Sheets("").Columns(8).Find(what:=Sheets("").Cells(r, "C").text, lookat:=xlWhole)
                       
                       If Not Ī Is Nothing Then
                         Sheets("мǷ Է").Cells(1, "A") = Sheets("").Cells(Ī.row, "H")
                       End If
                                     Total = 0
                                     For X = 13 To 193 Step (3)
                                             If Sheets("").Cells(r, X) <> "" Then
                                                 For TR = 3 To Cells(2, "B").End(xlDown).row
                                                 
                                                  TX = Sheets("мǷ Է").Rows(2).Find(what:=Sheets("").Cells(1, X), lookat:=xlWhole).Column
                                                  Sheets("мǷ Է").Cells(TR, TX) = "O"
                                                 Next TR
                                             End If
                                     Next X
                       Application.ScreenUpdating = True
                       Exit Sub
                End If
        Next r
Application.ScreenUpdating = True
End Sub
Sub мǷԷTOǷ()
    Dim TS As Worksheet
    Dim SS As Worksheet
    Dim FoundCell As Range
    Dim FirstFound As Range ' ù ° ã  ϴ 
    Dim targetDate As String
    Dim Ƿڻ As String
    Dim Ī As String
    Dim ä1 As String, ä2 As String, ȸ As String
    Dim SR As Long, TR As Long, SC As Long
    Dim AN As Integer
    Dim isFound As Boolean ' ãҴ θ ϴ 
    Dim XE As Integer ' ׿ 
    
    Set TS = Sheets("Ƿ")
    Set SS = Sheets("мǷ Է")

    ' ÷ḮƮ     
    If SS.Cells(50, "B").End(xlUp).row < 3 Then
        Debug.Print ""
        Exit Sub
    End If

    ' Է  
    Ƿڻ = Sheets("").Cells(Sheets("").Columns("H").Find(what:=SS.Cells(1, 1), lookat:=xlWhole).row, "B")
    Ī = SS.Cells(1, 1)
    ä1 = SS.Cells(1, "P")
    ä2 = SS.Cells(1, "S")
    ȸ = SS.Cells(1, "X")

    targetDate = Format(SS.Cells(1, "F"), "YYYY-MM-DD")

    ' мǷ Է  ó
    For SR = 3 To SS.Cells(50, "B").End(xlUp).row
        TR = TS.Cells(2, "A").End(xlDown).row + 1
        Set FoundCell = TS.Columns(1).Find(what:=CDate(targetDate), LookIn:=xlValues, lookat:=xlWhole)

        ' ʱ · ã  
        isFound = False

        ' ã   ٷ   Է ̵
        If FoundCell Is Nothing Then
            GoTo NoMatchFound
        Else
            ' ù ° ã  
            Set FirstFound = FoundCell
        End If

        '  ߺ Ȯ  ó
        Do
            '   ƿ   
            If FoundCell.Address = FirstFound.Address And isFound Then
                Exit Do
            End If

            ' ߺ  Ȯ  ڿ Ȯ ޽
            If FoundCell.Offset(0, 5).Value = SS.Cells(SR, "B").Value Then
                AN = MsgBox(SS.Cells(1, "F") & "" & vbCrLf & " ÷ᰡ ̹ ԷµǾ ֽϴ." & vbCrLf & " ԷϽðڽϱ?", vbYesNo, SS.Cells(SR, "B") & " Ƿ ߺԷ¢")
                If AN = vbYes Then
                    ' ߺ  
                    TS.Cells(FoundCell.row, "A") = SS.Cells(1, "F") ' ¥
                    TS.Cells(FoundCell.row, "B") = SS.Cells(1, "K") ' Ÿ 
                    TS.Cells(FoundCell.row, "D") = Ƿڻ
                    TS.Cells(FoundCell.row, "E") = Ī
                    TS.Cells(FoundCell.row, "F") = SS.Cells(SR, "B") ' ÷
                    TS.Cells(FoundCell.row, "G") = ȸ
                    TS.Cells(FoundCell.row, "H") = ä1
                    TS.Cells(FoundCell.row, "I") = ä2

                    ' м ׸ Է
                    For SC = 4 To 64 ' D3 BK1
                        If SS.Cells(SR, SC).Value <> "" Then
                            TS.Cells(FoundCell.row, 10 + SC) = "O"
                        Else
                            TS.Cells(FoundCell.row, 10 + SC) = ""
                        End If
                    Next SC
                End If
                isFound = True
                Exit Do
            End If

            '  ã
            Set FoundCell = TS.Columns(1).FindNext(FoundCell)

            ' ٽ ù ° ã  ƿ 
            If FoundCell Is Nothing Or FoundCell.Address = FirstFound.Address Then Exit Do

        Loop

NoMatchFound:
        ' ߺ Ͱ     Է
        If Not isFound Then
            TR = TS.Cells(2, "A").End(xlDown).row + 1
            TS.Cells(TR, "A") = SS.Cells(1, "F") ' ¥
            TS.Cells(TR, "B") = SS.Cells(1, "K") ' Ÿ 
            TS.Cells(TR, "D") = Ƿڻ
            TS.Cells(TR, "E") = Ī
            TS.Cells(TR, "F") = SS.Cells(SR, "B") ' ÷
            TS.Cells(TR, "G") = ȸ
            TS.Cells(TR, "H") = ä1
            TS.Cells(TR, "I") = ä2

            ' м ׸ Է
            For SC = 4 To 64 ' D3 BK1
                If SS.Cells(SR, SC).Value <> "" Then
                    TS.Cells(TR, 10 + SC) = "O"
                Else
                    TS.Cells(TR, 10 + SC) = ""
                End If
            Next SC
        End If
    Call м׸(SR)
    Next SR

    
End Sub

Sub м׸(ByVal SR As Long)
    Dim TS As Worksheet
    Dim SS As Worksheet
    Dim FirstFound As Range ' ù ° ã  ϴ 
    Set TS = Sheets("мڷ")
    Set SS = Sheets("мǷ Է")
    
    Debug.Print SR
    
    
    targetDate = Format(SS.Cells(1, "F"), "YYYY-MM-DD")
    
Set FoundCell = TS.Columns(1).Find(what:=CDate(targetDate), LookIn:=xlValues, lookat:=xlWhole)
        ' ʱ · ã  
        isFound = False

        ' ã   ٷ   Է ̵
        If FoundCell Is Nothing Then
            GoTo NoMatchFound
        Else
            ' ù ° ã  
            Set FirstFound = FoundCell
        End If
        
        Do
            '   ƿ   
            If FoundCell.Address = FirstFound.Address And isFound Then
                Exit Do
            End If

            ' ߺ  Ȯ  ڿ Ȯ ޽
            If FoundCell.Offset(0, 1).Value = SS.Cells(SR, "B").Value Then
                AN = MsgBox(SS.Cells(1, "F") & " ش÷ᰡ мǥ" & vbCrLf & " ̹մϴ." & vbCrLf & " ԷϽðڽϱ?", vbYesNo, SS.Cells(SR, "B") & " Ƿ׸ ߺ")
                If AN = vbYes Then

                    ' м(Ƿ)׸ Է
                    For SC = 4 To 64 ' D3 BK1
                        If SS.Cells(SR, SC).Value <> "" Then
                            TS.Cells(FoundCell.row, SC - 1).Interior.Pattern = -4142
                        Else
                            TS.Cells(FoundCell.row, SC - 1).Interior.Pattern = -4124
                        End If
                    Next SC
                End If
                isFound = True
                Exit Do
            End If

            '  ã
            Set FoundCell = TS.Columns(1).FindNext(FoundCell)

            ' ٽ ù ° ã  ƿ 
            If FoundCell Is Nothing Or FoundCell.Address = FirstFound.Address Then Exit Do

        Loop

NoMatchFound:
        ' ߺ Ͱ     Է
        If Not isFound Then
            TR = TS.Cells(2, "A").End(xlDown).row + 1
            ' м ׸ Է
            For SC = 4 To 64 ' D3 BK1
               TS.Cells(TR, 1) = SS.Cells(1, "K")
               TS.Cells(TR, "B") = SS.Cells(SR, "B")
               
                If SS.Cells(SR, SC).Value <> "" Then
                    TS.Cells(TR, SC - 1).Interior.Pattern = -4142
                Else
                    TS.Cells(TR, SC - 1).Interior.Pattern = -4124
                End If
                
            Next SC
        End If


End Sub
