Attribute VB_Name = "Էϱ"
Sub Է()
Application.ScreenUpdating = False
On Error Resume Next

Sheets("").Cells(3, "C") = UserForm1.ComboBox3.Value ' ü
Sheets("").Cells(5, "C") = UserForm1.ComboBox2.Value
Sheets("").Cells(6, "C") = UserForm1.TextBox5.Value
Sheets("").Cells(7, "C") = UserForm1.TextBox4.Value
Sheets("").Cells(8, "C") = UserForm1.TextBox2.Value
Sheets("").Cells(3, "H") = Format(UserForm1.TextBox1.Value, "YYYYMMDD") & "-" & Format(Now(), "HHMMSS")
Sheets("").Cells(4, "H") = UserForm1.TextBox1
Sheets("").Cells(6, "H") = UserForm1.ComboBox4.Value & "/061-685-8186"

Set  = Sheets("").Rows(100).Find(what:=UserForm1.ComboBox4.Value, lookat:=xlWhole)

If Not  Is Nothing Then
    Sheets("").Cells(7, "H") = Sheets("").Cells(100, .Column + 3)
End If

Sheets("").Range("A11:I43,K4:S36") = ""

For r = 0 To UserForm1.ListBox1.ListCount - 1

'    Sheets("").Cells(r + 11, "A") = r + 1
'    Sheets("").Cells(r + 11, "B") = UserForm1.ListBox1.List(r, 0)
'    Sheets("").Cells(r + 11, "D") = UserForm1.ListBox1.List(r, 1)
'    Set ׸ = Sheets("DB").Columns(3).Find(WHAT:=UserForm1.ListBox1.List(r, 1), lookat:=xlWhole)
'
'    If Not ׸ Is Nothing Then
'    Sheets("").Cells(r + 11, "F") = Sheets("DB").Cells(׸.row, "D")
'    End If
'    Sheets("").Cells(r + 11, "H") = UserForm1.TextBox3.Value
'    Sheets("").Cells(r + 11, "I") = UserForm1.ListBox1.List(r, 3)
    
    
                        x1 = Sheets("").Cells(100, "D").End(xlUp).row + 1
                        x2 = Sheets("").Cells(100, "N").End(xlUp).row + 1
                        
                        If Sheets("").Cells(100, "D").End(xlUp).row + 1 < 44 Then
                             X = Sheets("").Cells(100, "D").End(xlUp).row + 1
                             what = 0
                             Sheets("").Outline.ShowLevels RowLevels:=1, ColumnLevels:=1
                             N = 0
                             Sheets("").Range("A9").Font.Color = 16777215
                             Sheets("").Range("A45").Font.Color = 8421504
                             Sheets("").Range("A44").Font.Color = 0
                          Else
                             X = Sheets("").Cells(100, "N").End(xlUp).row + 1
                             what = 10
                             Sheets("").Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
                             N = 40
                             Sheets("").Range("A9").Font.Color = 8421504
                             Sheets("").Range("A45,H44").Font.Color = 16777215
                        End If
                        
                        Sheets("").Cells(X, what + 1) = X - 10 + N
                        Sheets("").Cells(X, what + 2) = UserForm1.ListBox1.List(r, 0)
                        Sheets("").Cells(X, what + 4) = UserForm1.ListBox1.List(r, 1)
                        amount = ws.Cells(Application.Match(childNode.text, ws.Columns(4), 0), selectedColumn).Value
                        
                        Set ES = Sheets("DB").Columns(3).Find(what:=UserForm1.ListBox1.List(r, 1), lookat:=xlWhole)
                        If Not ES Is Nothing Then
                        Sheets("").Cells(X, what + 6) = Sheets("DB").Cells(ES.row, "D")
                        End If
                        
'                        If Sheets("").Cells(X - 1, WHAT + 8) = "" Then
                         Sheets("").Cells(X, what + 8) = UserForm1.TextBox3.Value
'                        Else
'                         Sheets("").Cells(X, WHAT + 8) = Sheets("").Cells(X - 1, WHAT + 8)
'                        End If
                        
                        Sheets("").Cells(X, what + 9) = UserForm1.ListBox1.List(r, 3)
                        
                        
                        
    
    
Next r


Application.ScreenUpdating = True

End Sub

Sub _Է()
    Dim TX As Worksheet
    Dim RX As Worksheet
    Dim SX As Worksheet
    Dim Ī As Long
    Dim RTX As Long
    Dim RSX As Long
    Dim oldlist As Range
    Dim IsDuplicate As Boolean

    Set RX = ThisWorkbook.Sheets("")
    Set TX = ThisWorkbook.Sheets("")
    Set SX = ThisWorkbook.Sheets("Ƿ")

    RTX = TX.Cells(2, "A").End(xlDown).row + 1
    RSX = SX.Cells(2, "A").End(xlDown).row + 1
    RRX = RX.Cells(11, "A").End(xlDown).row
    RRXX = RX.Cells(4, "K").End(xlDown).row
    ' ߺ Ȯ  ʱⰪ
    IsDuplicate = False

    ' TX Ʈ I  ׸ Ȯ
    For Each oldlist In TX.Range(TX.Cells(3, "I"), TX.Cells(RTX - 1, "I"))
        If oldlist.Value = RX.Cells(3, "H").Value Then
            IsDuplicate = True
            RTX = oldlist.row
            Exit For
        End If
    Next oldlist

    ' ߺ ȮεǸ ڿ 
    If IsDuplicate Then
        A = MsgBox(RX.Cells(3, "H").Value & "  ̹ ԷµǾ ֽϴ." & vbCrLf & "ٽ() ԷϽðڽϱ?", vbYesNo)
        If A = vbNo Then Exit Sub
    End If

    '  Է
    TX.Cells(RTX, "A").Value = RX.Cells(4, "H").Value ' 
    TX.Cells(RTX, "B").Value = RX.Cells(3, "C").Value '  ûü
    Ī = Sheets("").Columns(2).Find(what:=Sheets("").Cells(3, "C").Value, lookat:=xlWhole).row
    TX.Cells(RTX, "C").Value = Sheets("").Cells(Ī, "H").Value ' Ī
    TX.Cells(RTX, "D").Value = Sheets("").Cells(Ī, "D").Value ' ǥ
    TX.Cells(RTX, "E").Value = RX.Cells(5, "C").Value ' û 
    TX.Cells(RTX, "F").Value = RX.Cells(6, "C").Value ' û ȭȣ
    TX.Cells(RTX, "G").Value = RX.Cells(7, "C").Value ' û ̸
    TX.Cells(RTX, "H").Value = RX.Cells(8, "C").Value '  ÷
    TX.Cells(RTX, "I").Value = RX.Cells(3, "H").Value ' ȣ
    
    For Each Ƿ׸ In RX.Range(RX.Cells(11, "D"), RX.Cells(RRX, "D"))
        If Ƿ׸ <> "" Then
            TC = TX.Rows(1).Find(what:=Ƿ׸, lookat:=xlWhole).Column
            TX.Cells(RTX, TC) = RX.Cells(Ƿ׸.row, "H")
            TX.Cells(RTX, TC + 1) = RX.Cells(Ƿ׸.row, "I")
            TX.Cells(RTX, TC + 2) = RX.Cells(Ƿ׸.row, "J")
        End If
    Next Ƿ׸
    
    If Application.Count(RX.Range("K4:K36")) > 0 Then
        For Each Ƿ׸ In RX.Range(RX.Cells(4, "N"), RX.Cells(RRXX, "N"))
            If Ƿ׸ <> "" Then
                TC = TX.Rows(1).Find(what:=Ƿ׸, lookat:=xlWhole).Column
                TX.Cells(RTX, TC) = RX.Cells(Ƿ׸.row, "R")
                TX.Cells(RTX, TC + 1) = RX.Cells(Ƿ׸.row, "S")
                TX.Cells(RTX, TC + 2) = RX.Cells(Ƿ׸.row, "T")
            End If
        Next Ƿ׸
    End If
    
    
    MsgBox RX.Cells(3, "C") & " " & RX.Cells(8, "C") & vbCrLf & "  Է Ϸ"
End Sub

