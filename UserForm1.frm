VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "-м"
   ClientHeight    =   10815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16620
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  ' 
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public driver As New WebDriver


Private Sub CheckBox62_Click()
If CheckBox62.Value = True Then
UserForm1.TextBox10 = UserForm1.TextBox9
Else
UserForm1.TextBox10 = ""
End If
End Sub

Private Sub ComboBox1_Change()

End Sub
Private Sub ComboBox2_Change()

    Combob2
        
End Sub

Private Sub ComboBox3_Change()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range
    Dim cell As Range
    Dim selectedValue As String
    On Error Resume Next
    
    ComboBox2.Clear ' ޺ڽ3 ʱȭ
    
    Set ws = ThisWorkbook.Sheets("ü")
    selectedValue = ComboBox3.Value
    
    ' õ  شϴ  E  ޺ڽ3 ߰
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    Set rng = ws.Range("A2:A" & lastRow)
    
    For Each cell In rng
        If cell.Value = selectedValue Then
            ComboBox2.AddItem cell.Offset(0, 2).Value & " " & cell.Offset(0, 4).Value & " " & cell.Offset(0, 3).Value ' E  ߰
        End If
    Next cell
    
    ComboBox2.ListIndex = 0
End Sub

Private Sub ComboBox5_Change()
AddTreeViewItem
End Sub

Private Sub ComboBox6_Change()
AddTreeViewItem
End Sub

Private Sub ComboBox7_Change()
׸Combo
End Sub

Private Sub CommandButton1_Click()   '====================-=-=-=-=-=-=-=-=-=-=-=-  
    driver.Start "edge"
    X = 13 - Sheets("DB").Cells(1, "R")
    ID = Sheets("DB").Cells(X, "P")
    PS = Sheets("DB").Cells(X, "Q")

    driver.Get "https://.kr/login.go"
    driver.FindElementById("user_email").SendKeys ID
    driver.FindElementById("login_pwd_confirm").SendKeys PS
    driver.FindElementById("login").Click
    driver.Get "https://.kr/ms/field_water.do" '
End Sub

Private Sub CommandButton13_Click()
For r = UserForm1.ListBox1.ListCount - 1 To 0 Step (-1)

UserForm1.ListBox1.RemoveItem (r)
Next r
հݾ
End Sub

Private Sub CommandButton14_Click()
Է
End Sub

Private Sub CommandButton15_Click()
Application.ScreenUpdating = False
Sheets("мǷ Է").Range("A3:BK100") = ""

Sheets("мǷ Է").Cells(1, "B") = UserForm1.TextBox8.text ' ÷
Sheets("мǷ Է").Cells(1, "F") = UserForm1.TextBox9.text ' 
Sheets("мǷ Է").Cells(1, "K") = UserForm1.TextBox10.text ' ÷ä

For r = 0 To UserForm1.ListBox2.ListCount - 1
 Sheets("мǷ Է").Cells(r + 3, "A") = UserForm1.ListBox2.List(r, 0) ' 
 Sheets("мǷ Է").Cells(r + 3, "B") = UserForm1.ListBox2.List(r, 1) '÷ Է
Next r
Ƿ׸üũ
Application.ScreenUpdating = True

End Sub

Private Sub CommandButton16_Click()
UserForm1.ListBox2.Clear
End Sub

Private Sub CommandButton17_Click()
۾

End Sub

Private Sub CommandButton18_Click()
۾2
End Sub

Private Sub CommandButton3_Click() ' ȯ() Է -------------------------PAGE1
If driver.FindElementById("ui-id-1").Attribute("aria-expanded") = True Then
driver.FindElementByXPath("//*[@id=""edit_env_psic_name""]").Clear.SendKeys ListView1.ListItems(1).ListSubItems(4).text
End If
End Sub
Private Sub CommandButton4_Click() 'м  -------------------------PAGE1
If driver.FindElementById("ui-id-1").Attribute("aria-expanded") = True Then
    Y = Left(ListView1.ListItems(1).ListSubItems(1).text, 4)
    M = Right(Left(ListView1.ListItems(1).ListSubItems(1).text, 7), 2)
    D = Right(Left(ListView1.ListItems(1).ListSubItems(1).text, 10), 2)
    DATEX = Format(Y, "0000") & "-" & Format(M, "00") & "-" & Format(D, "00")
    script1 = "var input = document.getElementById('edit_meas_start_dt');" & "input.value = '" & DATEX & "';" & "input.blur();"
    driver.ExecuteScript script1
End If
End Sub
Private Sub CommandButton5_Click() '༱     -------------------------PAGE1
    Dim keys As New Selenium.keys
    Dim X As Range
    Set X = Sheets("").Columns(8).Find(what:=ListView1.ListItems(1).ListSubItems(2).text, lookat:=xlWhole)
If driver.FindElementById("ui-id-1").Attribute("aria-expanded") = True Then
    If Not X Is Nothing Then
        ' ʿ  
        driver.FindElementByXPath("//*[@id=""wid-id-1""]/div/div[2]/div/fieldset/div[6]/section/label[2]/span").Click
        driver.FindElementByXPath("/html/body/span/span/span[1]/input").Click
        driver.FindElementByXPath("/html/body/span/span/span[1]/input").SendKeys Left(Sheets("").Cells(X.row, "J").text, 6)
        driver.FindElementByXPath("/html/body/span/span/span[1]/input").SendKeys (keys.Enter)
    End If
End If
End Sub

Private Sub CommandButton6_Click() 'м(Ƿ)׸ Է  -------------------------PAGE1
CommandButton7_Click

Dim jsScript As String
Dim i As Integer

jsScript = "var selectElement = document.getElementById('edit_meas_item');"

For i = 1 To ListView3.ListItems.Count
    Dim cellValue As Variant
    
    cellValue = Sheets("TESTS").Cells(i + 1, "A").Value
    X = Sheets("TESTS").Columns(2).Find(what:=ListView3.ListItems(i).text, lookat:=1).row

'    If Not IsEmpty(cellValue) Then
        jsScript = jsScript & "selectElement.options[" & X - 1 & "].selected = true;"
'    End If
Next i

jsScript = jsScript & "selectElement.dispatchEvent(new Event('change'));"

' JavaScript ڵ带 
driver.ExecuteScript jsScript

End Sub
Private Sub CommandButton7_Click()  'м(Ƿ)׸   -------------------------PAGE1
Dim itemCountScript As String
itemCountScript = "return document.querySelectorAll('#wid-id-1 div div:nth-child(2) fieldset div:nth-child(9) section div span span:nth-child(1) span ul li span').length;"

' ũƮ 
Dim itemCount As Integer
itemCount = driver.ExecuteScript(itemCountScript)
For i = itemCount To 1 Step (-1)
    clickScript = "document.querySelector('#wid-id-1 div div:nth-child(2) fieldset div:nth-child(9) section div span span:nth-child(1) span ul li:nth-child(" & i & ") span').click();"
    driver.ExecuteScript clickScript
Next i

driver.FindElementByXPath("//*[@id=""wid-id-1""]/div/div[2]/div/fieldset/div[9]/section/div/span/span[1]/span/ul").Click

End Sub
Private Sub CommandButton8_Click() 'мⰣ     -------------------------PAGE3
If driver.FindElementById("ui-id-3").Attribute("aria-expanded") = True Then
    Y = Left(ListView1.ListItems(1).ListSubItems(1).text, 4)
    M = Right(Left(ListView1.ListItems(1).ListSubItems(1).text, 6), 2)
    D = Right(Left(ListView1.ListItems(1).ListSubItems(1).text, 8), 2)
    DATES = Y & M & D
    DATEE = Year(Now) & "-" & Month(Now) & "-" & Day(Now)
    script1 = "var input = document.getElementById('anze_start_dt_1');" & "input.value = '" & DATES & "';" & "input.blur();"
    driver.ExecuteScript script1
    
    script2 = "var input = document.getElementById('anze_end_dt_1');" & "input.value = '" & DATEE & "';" & "input.blur();"
    driver.ExecuteScript script2
End If
End Sub

Private Sub CommandButton9_Click()  '==============================äο Է   PAGE1
    Dim jsScript As String
    Dim x1 As Range, x2 As Range

    ' "edit_emp_id" Ұ ǥõǴ Ȯ
If driver.FindElementById("ui-id-1").Attribute("aria-expanded") = True Then
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''  äο 
    Dim element As Object
    Dim liElements As Object
    Dim liCount As Integer

    ' XPath  ã
    Set element = driver.FindElementByXPath("//*[@id=""wid-id-4""]/div/div[2]/div/fieldset/div[2]/section[2]/span/span[1]/span/ul")

    ' ã ҿ li ±׵ ã
    Set liElements = element.FindElementsByTag("li")

    ' li ±  Ȯ
    liCount = liElements.Count

    If liCount > 1 Then
    For r = liCount To 2 Step (-1)
     driver.FindElementByXPath("//*[@id=""wid-id-4""]/div/div[2]/div/fieldset/div[2]/section[2]/span/span[1]/span/ul/li[" & r - 1 & "]/span").Click
    Next r
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ListView2  ְ 0 ƴ 쿡 ó
        If ListView2.ListItems(1).text <> "0" And ListView2.ListItems(1).ListSubItems(1).text <> "0" Then
            ' JavaScript ũƮ ʱȭ
            jsScript = "var selectElement = document.getElementById('edit_emp_id');"

            ' ListView2  شϴ  ã
            Set x1 = Sheets("DB").Columns(14).Find(what:=ListView2.ListItems(1).text, lookat:=xlWhole).Rows
            Set x2 = Sheets("DB").Columns(14).Find(what:=ListView2.ListItems(1).ListSubItems(1).text, lookat:=xlWhole).Rows

            ' JavaScript ڵ ۼ
            jsScript1 = jsScript & "selectElement.options[" & x1.row - 1 & "].selected = true;"
            jsScript2 = jsScript & "selectElement.options[" & x2.row - 1 & "].selected = true;"

            ' JavaScript ڵ忡 ̺Ʈ ġ ߰
            jsScript1 = jsScript1 & "selectElement.dispatchEvent(new Event('change'));"
            jsScript2 = jsScript2 & "selectElement.dispatchEvent(new Event('change'));"

            ' JavaScript ڵ 
            driver.ExecuteScript jsScript1
            driver.ExecuteScript jsScript2
        End If

        ' XPath Ͽ  Ŭ
        driver.FindElementByXPath("//*[@id=""wid-id-1""]/div/div[2]/div/fieldset/div[9]/section").Click
End If
End Sub
Private Sub CommandButton2_Click() 'ä  -------------------------PAGE2

If driver.FindElementById("ui-id-2").Attribute("aria-expanded") = True Then
driver.FindElementById("samp_vesl_desc").Clear.SendKeys "P:2, G:1"
End If

End Sub

Private Sub CommandButton10_Click() 'PAGE2 ÷ä   -------------------------PAGE2

If driver.FindElementById("ui-id-2").Attribute("aria-expanded") = True Then
 driver.FindElementByXPath("//*[@id=""edit_meas_loc_desc_1""]").Clear.SendKeys ListView1.ListItems(1).ListSubItems(3).text
End If

End Sub
Private Sub CommandButton11_Click() 'PAGE4 мڷ Է
    If driver.FindElementById("ui-id-4").Attribute("aria-expanded") = True Then
        Dim trElements As Object
        Dim desiredValue1 As String, desiredValue2 As String
        Dim script As String
        Dim startTime As Double
        Set trElements = driver.FindElementsByXPath("//*[@id='tbAnze']/tbody/tr[contains(@class, 'tr_')]")
        
        Dim trCount As Integer
        trCount = trElements.Count
        
'        Debug.Print " " & trCount & " tr ±׸ ãҽϴ."

        Y = Left(ListView1.ListItems(1).ListSubItems(1).text, 4)
        M = Right(Left(ListView1.ListItems(1).ListSubItems(1).text, 7), 2)
        D = Right(Left(ListView1.ListItems(1).ListSubItems(1).text, 10), 2)
        DATES = Format(Y, "0000") & "-" & Format(M, "00") & "-" & Format(D, "00")
        DATEE = Format(Year(Now), "0000") & "-" & Format(Month(Now), "00") & "-" & Format(Day(Now), "00")
            
        For i = 1 To trCount
            For X = 1 To ListView3.ListItems.Count
                If ListView3.ListItems(X).text = driver.FindElementById("meas_item_name_" & i).text Then
                  startTime = Timer
                  
                    driver.FindElementById("allow_val_" & i).Clear.SendKeys ListView3.ListItems(X).ListSubItems(4).text  'ġ
                    driver.FindElementById("anze_val_" & i).Clear.SendKeys ListView3.ListItems(X).ListSubItems(1).text   'м

                    desiredValue1 = ListView3.ListItems(X).ListSubItems(5).text                                          'м(Method)
                    desiredValue2 = ListView3.ListItems(X).ListSubItems(6).text                                          'м

                    Set selectElement1 = driver.FindElementById("anze_mthd_" & i)                                        'м Է
                    selectElement1.AsSelect.SelectByValue desiredValue1

                    Set selectElement2 = driver.FindElementById("anze_equip_no_1" & i)                                   'м Է
                    selectElement2.AsSelect.SelectByValue desiredValue2
                    
                    Set selectElement = driver.FindElementByName("anze_login_id_1" & i).AsSelect                         'мη (忡 ãƼ) Է
                    selectElement.SelectByValue Sheets("DB").Cells(Sheets("DB").Columns(14).Find(what:=ListView3.ListItems(X).ListSubItems(7).text, lookat:=xlWhole).row, "R")
                    

                    script1 = "var input = document.getElementById('anze_start_dt_" & i & "');" & "input.value = '" & DATES & "';" & "input.blur();"
                    driver.ExecuteScript script1
                    
                    driver.FindElementByName("anze_start_tm_" & i).SendKeys "0900"

                    script2 = "var input = document.getElementById('anze_end_dt_" & i & "');" & "input.value = '" & DATEE & "';" & "input.blur();"
                    driver.ExecuteScript script2
                    
                    driver.FindElementByName("anze_end_tm_" & i).SendKeys "1800"

                    Exit For
                     Do While Timer < startTime + 0.01
                       DoEvents
                    Loop
                End If
            Next X
        Next i
    End If
End Sub
Private Sub CommandButton12_Click() ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''TES

Sheets.PrintOut

  
End Sub
Sub TESTcheckbox()
Dim i As Integer

For i = 1 To 60 ' 3 üũڽ Դϴ. ʿ信  ϼ.
    Me.Controls("Checkbox" & i).Caption = Sheets("׸ڵ").Cells(i, "J")
Next i

End Sub



Private Sub Label11_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
sellist = UserForm1.ListBox1.ListIndex
UserForm1.ListBox1.RemoveItem (sellist)
    հݾ
End Sub





Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim selectedIndex As Integer
    Dim i As Integer
    
    ' õ ׸ ε ɴϴ.
    selectedIndex = UserForm1.ListBox2.ListIndex
    
    ' õ ׸ ִ 쿡 մϴ.
    If selectedIndex <> -1 Then
        ' ׸ 
        UserForm1.ListBox2.RemoveItem selectedIndex
        
        ' ׸   ȣ ٽ ű
        For i = 0 To UserForm1.ListBox2.ListCount - 1
            UserForm1.ListBox2.List(i, 0) = Format(i + 1, "00") ' ù °  ȣ ű
        Next i
    Else
        MsgBox "õ ׸ ϴ.", vbExclamation
    End If
End Sub




Private Sub MultiPage1_Change()

x1 = UserForm1.Left
x2 = UserForm1.Top

' UserForm (0,0) ġ ̵
UserForm1.Move 0, 0

'    ġ ̵
UserForm1.Move x1, x2
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox3_Change()
For r = o To UserForm1.ListBox1.ListCount - 1
UserForm1.ListBox1.List(r, 2) = TextBox3.text
Next r

    հݾ
End Sub

Private Sub TreeView1_DblClick()

ι鿩

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim ws As Worksheet
    Dim ü As Range
    
           Application.ScreenUpdating = False

 Sheets("輺").Range("C1:D7,F3:H7,K1:L7,N3:P7,A10:P41") = ""

   
If Not TreeView1.SelectedItem Is Nothing Then
    If Not TreeView1.SelectedItem.Parent Is Nothing Then
        Debug.Print "ڽĳ: " & TreeView1.SelectedItem.Parent.text
        Set ws = ThisWorkbook.Sheets("Ƿ")
        lastRow = Sheets("Ƿ").Cells(Sheets("Ƿ").Rows.Count, "A").End(xlUp).row
        For r = 2 To lastRow
         X = "" & Sheets("Ƿ").Cells(r, "E").text & "" & Sheets("Ƿ").Cells(r, "F").Value '  ws.Cells(i, "C")  & ws.Cells(i, 4).Value)
         If Sheets("Ƿ").Cells(r, "A") = TreeView1.SelectedItem.Parent.text And X = TreeView1.SelectedItem.text Then

          '===============================================================================================
         ListView1.ListItems.Clear
         With ListView1
        .ColumnHeaders.Clear '  ÷ 
        .Gridlines = True
        .View = lvwReport ' Report  

        '  ÷ ʺ Ϸ ʿ信  Width Ӽ   ֽϴ.
        .ColumnHeaders.Add , , "Ƿ", 100
        .ColumnHeaders.Add , , "ä", 100
        .ColumnHeaders.Add , , "Ƿڻ", 100
        .ColumnHeaders.Add , , "÷", 120
        .ColumnHeaders.Add , , "ȸ", 100

         End With

         ListView2.ListItems.Clear
         With ListView2
        .ColumnHeaders.Clear '  ÷ 
        .Gridlines = True
        .View = lvwReport ' Report  
        '  ÷ ʺ Ϸ ʿ信  Width Ӽ   ֽϴ.
        .ColumnHeaders.Add , , "÷ä1", 100
        .ColumnHeaders.Add , , "÷ä2", 100
        .ColumnHeaders.Add , , "", 120
        .ColumnHeaders.Add , , "", 100
        .ColumnHeaders.Add , , "м", 100
         End With

         ListView3.ListItems.Clear
         With ListView3
        .ColumnHeaders.Clear '  ÷ 
        .Gridlines = True
        .View = lvwReport ' Report  
        '  ÷ ʺ Ϸ ʿ信  Width Ӽ   ֽϴ.
        .ColumnHeaders.Add , , "Ƿ׸", 130
        .ColumnHeaders.Add , , "м", 70
        .ColumnHeaders.Add , , "м", 170
        .ColumnHeaders.Add , , "м", 100
        .ColumnHeaders.Add , , "", 50
        .ColumnHeaders.Add , , "Method NO", 20
        .ColumnHeaders.Add , , "instrument NO", 20
        .ColumnHeaders.Add , , "м", 20
         End With


         Set item = ListView1.ListItems.Add(1, , Sheets("Ƿ").Cells(r, "A").Value) 'Ƿ
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "B").Value                'ä
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "E").Value                'Ƿڻ 'Ī
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "F").Value                '÷
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "G").Value                'ȸ

         Set item = ListView2.ListItems.Add(1, , Sheets("Ƿ").Cells(r, "H").Value) '÷ä-1
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "I").Value                '÷ä-2
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "J").Value                '
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "K").Value                '
         item.ListSubItems.Add , , Sheets("Ƿ").Cells(r, "L").Value                'м

         For T = Range("N1").Column To Range("BT1").Column
'         Sheets("輺").Cells(10, "A") = 1
         
         If Sheets("Ƿ").Cells(r, T) <> "" Then

         
               G = Sheets("輺").Range("A42").End(xlUp).row + 1
               G2 = Sheets("輺").Range("I42").End(xlUp).row + 1
            If G <= 41 Then
               Sheets("輺").Cells(G, "A") = G - 9
               Sheets("輺").Cells(G, "D") = Sheets("Ƿ").Cells(1, T)
               GX = Sheets("ܰ").Columns(4).Find(what:=Sheets("Ƿ").Cells(1, T), lookat:=xlWhole).row
               Sheets("輺").Cells(G, "B") = Sheets("ܰ").Cells(GX, 3)
               Sheets("輺").Cells(G, "E") = Sheets("ܰ").Cells(GX, 2)
               
            Else
               Sheets("輺").Cells(G2, "I") = G2 - 9
               GX = Sheets("ܰ").Columns(4).Find(what:=Sheets("Ƿ").Cells(1, T), lookat:=xlWhole).row
               Sheets("輺").Cells(G2, "J") = Sheets("ܰ").Cells(GX, 3)
               Sheets("輺").Cells(G2, "L") = Sheets("Ƿ").Cells(1, T)
               Sheets("輺").Cells(G2, "M") = Sheets("ܰ").Cells(GX, 2)

            End If
            
         Set Titem = ListView3.ListItems.Add(ListView3.ListItems.Count + 1, , Sheets("Ƿ").Cells(1, T))
                     Titem.ListSubItems.Add , , "-"             'Subitem-1 м
                     Titem.ListSubItems.Add , , "-"             'Subitem-2 м
                     Titem.ListSubItems.Add , , "-"             'Subitem-3 м
                     Titem.ListSubItems.Add , , "-"             'Subitem-4 
                     Titem.ListSubItems.Add , , "-"             'Subitem-5 Method NO
                     Titem.ListSubItems.Add , , "-"             'Subitem-6 instrument NO
                     Titem.ListSubItems.Add , , "-"             'Subitem-7 м
         End If
         Next T
         '===============================================================================================

         End If

        Next r

        мҷ
        ã

If ActiveSheet.Name = "Ϻ" Then

End If

Call 


        ListView3.ColumnHeaders(1).text = "Ƿ׸ ( " & ListView3.ListItems.Count & ") "
    Else
        Label1.Caption = "ãνϴ"
    End If


End If
       Set ü = Sheets("").Columns("H").Find(what:=UserForm1.ListView1.ListItems(1).ListSubItems(2), lookat:=xlWhole)
       If Sheets("輺").Cells(10, "I") = "" Then
           Sheets("輺").PageSetup.PrintArea = "A1:H44"
           
           If Not ü Is Nothing Then
           
           
           Sheets("輺").Range("C3,K3") = Sheets("").Cells(ü.row, "B")
           Sheets("輺").Range("C4,K4") = Sheets("").Cells(ü.row, "D")
           Sheets("輺").Range("C5,K5") = UserForm1.ListView1.ListItems(1).ListSubItems(4)
           Sheets("輺").Range("C7,K7") = UserForm1.ListView1.ListItems(1).ListSubItems(3)
           
           
           Sheets("輺").Range("F3,N3") = UserForm1.ListView1.ListItems(1).ListSubItems(1)
           Sheets("輺").Range("F4,N4") = UserForm1.ListView2.ListItems(1).text & ", " & UserForm1.ListView2.ListItems(1).ListSubItems(1)
           
           If UserForm1.ListView2.ListItems(1).ListSubItems(3) <> "O" Then
           Sheets("輺").Range("F7,N7") = ""
           Else
           Sheets("輺").Range("F7,N7") = " "
           End If
           
           End If
           
         Else
           Sheets("輺").PageSetup.PrintArea = "A1:H44,I1:P44"
           
           If Not ü Is Nothing Then
           
           
           Sheets("輺").Range("C3,K3") = Sheets("").Cells(ü.row, "B")
           Sheets("輺").Range("C4,K4") = Sheets("").Cells(ü.row, "D")
           Sheets("輺").Range("C5,K5") = UserForm1.ListView1.ListItems(1).ListSubItems(4)
           Sheets("輺").Range("C7,K7") = UserForm1.ListView1.ListItems(1).ListSubItems(3)
           
           
           Sheets("輺").Range("F3,N3") = UserForm1.ListView1.ListItems(1).ListSubItems(1)
           Sheets("輺").Range("F4,N4") = UserForm1.ListView2.ListItems(1).text & ", " & UserForm1.ListView2.ListItems(1).ListSubItems(1)
           
           If UserForm1.ListView2.ListItems(1).ListSubItems(3) <> "O" Then
           Sheets("輺").Range("F7,N7") = ""
           Else
           Sheets("輺").Range("F7,N7") = " "
           End If
           
           End If
           
       End If
       
       Application.ScreenUpdating = True



End Sub


Private Sub TreeView2_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim childNode As MSComctlLib.Node
    Dim parentNode As MSComctlLib.Node
    Dim listBox As MSForms.listBox
    Dim comboBox As MSForms.comboBox
    Dim ws As Worksheet
    Dim i As Long
    Dim IsDuplicate As Boolean
    Dim selectedColumn As Integer
    Dim amount As Variant

    ' UserForm1 ListBox1  ComboBox1 
    Set listBox = Me.ListBox1
    Set comboBox = Me.ComboBox1
    Set ws = ThisWorkbook.Sheets("ܰ")
    
   If UserForm1.TreeView2.SelectedItem.text = "Ưع" Then
   ƯϰԷ
   End If
   
   
    
    ' ComboBox  ׸ شϴ ÷ ã
    selectedColumn = 0
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(1, i).Value = comboBox.Value Then
            selectedColumn = i
            Exit For
        End If
    Next i
    
    If selectedColumn = 0 Then
        MsgBox "ComboBox  ׸ شϴ ÷ ã  ϴ.", vbExclamation
        Exit Sub
    End If
    
    ' ListBox ʱȭ
'    listBox.Clear
    
    ' õ ׸ ִ Ȯ
    If Not UserForm1.TreeView2.SelectedItem Is Nothing Then
        With UserForm1.TreeView2.SelectedItem
            ' õ ׸ θ  Ȯ
            If .Children > 0 Then
                ' θ     
                Set childNode = .Child
                
                Do While Not childNode Is Nothing
                    ' ߺ Ȯ
                    IsDuplicate = False
                    For i = 0 To listBox.ListCount - 1
                        If listBox.List(i, 1) = childNode.text Then
                            IsDuplicate = True
                            Exit For
                        End If
                    Next i
                    
                    ' ߺ ƴϸ ListBox θ  ڽ   ݾ ߰
                    If Not IsDuplicate Then
                        listBox.AddItem
                        listBox.List(listBox.ListCount - 1, 0) = .text
                        listBox.List(listBox.ListCount - 1, 1) = childNode.text
                        amount = ws.Cells(Application.Match(childNode.text, ws.Columns(4), 0), selectedColumn).Value
                        listBox.List(listBox.ListCount - 1, 2) = TextBox3.Value
                        listBox.List(listBox.ListCount - 1, 3) = Format(amount, "#,###")
                        listBox.List(listBox.ListCount - 1, 4) = Format(amount * TextBox3.Value, "#,###")
                    End If
                    Set childNode = childNode.Next
                Loop
            Else
                ' ڽ   θ  õ ڽ   ݾ ߰
                If Not .Parent Is Nothing Then
                    Set parentNode = .Parent
                    ' ߺ Ȯ
                    IsDuplicate = False
                    For i = 0 To listBox.ListCount - 1
                        If listBox.List(i, 1) = .text Then
                            IsDuplicate = True
                            Exit For
                        End If
                    Next i
                    
                    ' ߺ ƴϸ ListBox θ  ڽ   ݾ ߰
                    If Not IsDuplicate Then
                        listBox.AddItem
                        listBox.List(listBox.ListCount - 1, 0) = parentNode.text
                        listBox.List(listBox.ListCount - 1, 1) = .text
                        amount = ws.Cells(Application.Match(.text, ws.Columns(4), 0), selectedColumn).Value
                        listBox.List(listBox.ListCount - 1, 2) = TextBox3.Value
                        listBox.List(listBox.ListCount - 1, 3) = Format(amount, "#,###")
                        listBox.List(listBox.ListCount - 1, 4) = Format(amount * TextBox3.Value, "#,###")
                    End If
                Else
                    Debug.Print "This node has no parent."
                End If
            End If
        End With
    Else
        Debug.Print "No item is selected in the TreeView."
    End If
    
    հݾ
    
End Sub

Private Sub TreeView3_Click()
 Dim ws As Worksheet
 Dim Z
 Dim CL As Integer
 
 UserForm1.TextBox8 = UserForm1.TreeView3.SelectedItem.text
 
If Not UserForm1.TreeView3.SelectedItem Is Nothing Then
    If Not UserForm1.TreeView3.SelectedItem.Parent Is Nothing Then
        UserForm1.TextBox9 = UserForm1.TreeView3.SelectedItem.Parent.text
    Else
        ' õ ׸ ֻ  , θ Ƿ Label12 ʱȭ
        UserForm1.TextBox9 = "No Parent"
    End If
Else
    ' õ ׸   Label12 ʱȭ
    UserForm1.TextBox9 = "No Item Selected"
End If
 
   
If Not TreeView3.SelectedItem Is Nothing Then
    If Not TreeView3.SelectedItem.Parent Is Nothing Then

        Set ws = ThisWorkbook.Sheets("")
        lastRow = Sheets("").Cells(Sheets("").Rows.Count, "A").End(xlUp).row
        For r = 2 To lastRow
         Z = "" & Sheets("").Cells(r, "C").text & "" & Sheets("").Cells(r, "H").Value '  ws.Cells(i, "C")  & ws.Cells(i, 4).Value)
         
         If Sheets("").Cells(r, "A") = TreeView3.SelectedItem.Parent.text And Z = TreeView3.SelectedItem.text Then

          
          Set listBox = UserForm1.ListBox1
          listBox.Clear
          
          UserForm1.TextBox1 = Sheets("").Cells(r, "A")
          
          Set Ī = Sheets("").Columns(8).Find(what:=Sheets("").Cells(r, "C").text, lookat:=xlWhole)
          If Not Ī Is Nothing Then
            UserForm1.ComboBox3.ListIndex = Ī.row - 2
          End If
          
          UserForm1.ComboBox4.Value = Sheets("").Cells(r, "K") '####  ؾ ..
          UserForm1.TextBox4.Value = Sheets("").Cells(r, "G")
          UserForm1.TextBox5.Value = Sheets("").Cells(r, "F")
          UserForm1.TextBox2.Value = Sheets("").Cells(r, "H")
          UserForm1.ComboBox2.Value = Sheets("").Cells(r, "E")
          
                        Total = 0
                        For X = 13 To 193 Step (3)

                        If Sheets("").Cells(r, X) <> "" Then
                        amount = 0
                       
                        Z = Sheets("ܰ").Columns(4).Find(what:=Sheets("").Cells(1, X), lookat:=xlWhole).row
                        
                        listBox.AddItem
                        listBox.List(listBox.ListCount - 1, 0) = Sheets("ܰ").Cells(Z, "C")    'Ƿ/м׸ 뱸
                        listBox.List(listBox.ListCount - 1, 1) = Sheets("").Cells(1, X)  'Ƿ/м׸ ұ
                        listBox.List(listBox.ListCount - 1, 2) = Sheets("").Cells(r, X)  'Ƿ/м׸ 
                        listBox.List(listBox.ListCount - 1, 3) = Format(Sheets("").Cells(r, X + 1), "#,###") 'Ƿ/м׸ ܰ
                        amount = Sheets("").Cells(r, X) * Sheets("").Cells(r, X + 1)
                        Total = amount + Total
                        
                        listBox.List(listBox.ListCount - 1, 4) = Format(amount, "#,###")
                        
                        End If
                        
                        
                        Next X
                         

         
         
          
         End If

        Next r


    Else
        Label1.Caption = "ãνϴ"
    End If


End If
'========================================================================================================================
X = 0
For r = 0 To UserForm1.ListBox1.ListCount - 1
     If UserForm1.ListBox1.List(r, 2) <> "" And UserForm1.ListBox1.List(r, 2) <> "-" And UserForm1.ListBox1.List(r, 2) <> 0 Then
         X = UserForm1.ListBox1.List(r, 2) + X
     End If
Next r

If Not UserForm1.ListBox1.ListCount = 0 And UserForm1.TextBox3 <> "" Then

  UserForm1.Label5.Caption = UserForm1.ListBox1.List(0, 1) & " " & UserForm1.ListBox1.ListCount & " " & Format(Total, "#,###")
Else
  UserForm1.Label5.Caption = "Ǽ/Ѿ"
End If
'=========================================================================================================================

Ʈ3Ŭüũ

End Sub
Private Sub TreeView4_Click()
UserForm1.TreeView4.SelectedItem.ForeColor = RGB(255, 0, 0)
ǷڸƮ̵
End Sub
Private Sub TreeView5_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    Dim ListItem As ListItem
    Dim NodeText As String
    Dim ParentText As String
    Dim childNode As MSComctlLib.Node
    Dim IsDuplicate As Boolean
    Dim ItemIndex As Integer
    
    

    
    
    '  ListView ִ    ο ȣ 
    ItemIndex = ListView4.ListItems.Count + 1
    
    ' θ   ڽ  ü ó
    If Node.Children > 0 Then
        ParentText = Node.text
        Set childNode = Node.Child
        
        Do While Not childNode Is Nothing
            NodeText = childNode.text
            IsDuplicate = False
            
            ' ߺ üũ
            For i = 1 To ListView4.ListItems.Count
                If ListView4.ListItems(i).SubItems(1) = ParentText And ListView4.ListItems(i).SubItems(2) = NodeText Then
                    IsDuplicate = True
                    Exit For
                End If
            Next i
            
            ' ߺ ʴ 쿡 ߰
            If Not IsDuplicate Then
                Set ListItem = ListView4.ListItems.Add(, , ItemIndex)
                ListItem.SubItems(1) = ParentText
                ListItem.SubItems(2) = NodeText
                ListItem.SubItems(3) = ""
                ItemIndex = ItemIndex + 1
                мԷ¸Ʈǹ
           End If
            
            '  ڽ  ̵
            Set childNode = childNode.Next
        Loop
    Else
        ' ڽ 尡   ( )
        If Not Node.Parent Is Nothing Then
            ParentText = Node.Parent.text
        Else
            ParentText = Node.text
        End If
        
        NodeText = Node.text
        IsDuplicate = False
        
        ' ߺ üũ
        For i = 1 To ListView4.ListItems.Count
            If ListView4.ListItems(i).SubItems(1) = ParentText And ListView4.ListItems(i).SubItems(2) = NodeText Then
                IsDuplicate = True
                Exit For
            End If
        Next i
        
        ' ߺ ʴ 쿡 ߰
        If Not IsDuplicate Then
            Set ListItem = ListView4.ListItems.Add(, , ItemIndex)
            ListItem.SubItems(1) = ParentText
            ListItem.SubItems(2) = NodeText
            ListItem.SubItems(3) = ""
            мԷ¸Ʈǹ
        End If
    End If
    
    
End Sub








Private Sub TreeView6_Click()
 Dim ws As Worksheet
 Dim Z
 Dim CL As Integer
 

 
If Not UserForm1.TreeView6.SelectedItem Is Nothing Then
    If Not UserForm1.TreeView6.SelectedItem.Parent Is Nothing Then
        UserForm1.TextBox8 = UserForm1.TreeView6.SelectedItem.text
        UserForm1.TextBox9 = UserForm1.TreeView6.SelectedItem.Parent.text
    Else
        ' õ ׸ ֻ  , θ Ƿ Label12 ʱȭ
        UserForm1.TextBox9 = UserForm1.TreeView6.SelectedItem.text

    End If
Else
    ' õ ׸   Label12 ʱȭ
    UserForm1.TextBox9 = "No Item Selected"
End If
 
ã

 
   
''If Not TreeView3.SelectedItem Is Nothing Then
''    If Not TreeView3.SelectedItem.Parent Is Nothing Then
''
''        Set ws = ThisWorkbook.Sheets("")
''        lastRow = Sheets("").Cells(Sheets("").Rows.Count, "A").End(xlUp).Row
''        For r = 2 To lastRow
''         Z = "" & Sheets("").Cells(r, "C").Text & "" & Sheets("").Cells(r, "H").Value '  ws.Cells(i, "C")  & ws.Cells(i, 4).Value)
''
''         If Sheets("").Cells(r, "A") = TreeView3.SelectedItem.Parent.Text And Z = TreeView3.SelectedItem.Text Then
''
''
''          Set listBox = UserForm1.ListBox1
''          listBox.Clear
''
''          UserForm1.TextBox1 = Sheets("").Cells(r, "A")
''
''          Set Ī = Sheets("").Columns(8).Find(what:=Sheets("").Cells(r, "C").Text, lookat:=xlWhole)
''          If Not Ī Is Nothing Then
''            UserForm1.ComboBox3.ListIndex = Ī.Row - 2
''          End If
''
''          UserForm1.ComboBox4.Value = Sheets("").Cells(r, "K") '####  ؾ ..
''          UserForm1.TextBox4.Value = Sheets("").Cells(r, "G")
''          UserForm1.TextBox5.Value = Sheets("").Cells(r, "F")
''          UserForm1.TextBox2.Value = Sheets("").Cells(r, "H")
''          UserForm1.ComboBox2.Value = Sheets("").Cells(r, "E")
''
''                        Total = 0
''                        For X = 13 To 193 Step (3)
''
''                        If Sheets("").Cells(r, X) <> "" Then
''                        amount = 0
''
''                        Z = Sheets("ܰ").Columns(4).Find(what:=Sheets("").Cells(1, X), lookat:=xlWhole).Row
''
''                        listBox.AddItem
''                        listBox.List(listBox.ListCount - 1, 0) = Sheets("ܰ").Cells(Z, "C")    'Ƿ/м׸ 뱸
''                        listBox.List(listBox.ListCount - 1, 1) = Sheets("").Cells(1, X)  'Ƿ/м׸ ұ
''                        listBox.List(listBox.ListCount - 1, 2) = Sheets("").Cells(r, X)  'Ƿ/м׸ 
''                        listBox.List(listBox.ListCount - 1, 3) = Format(Sheets("").Cells(r, X + 1), "#,###") 'Ƿ/м׸ ܰ
''                        amount = Sheets("").Cells(r, X) * Sheets("").Cells(r, X + 1)
''                        Total = amount + Total
''
''                        listBox.List(listBox.ListCount - 1, 4) = Format(amount, "#,###")
''
''                        End If
''
''
''                        Next X
''
''
''
''
''
''         End If
''
''        Next r
''
''
''    Else
''        Label1.Caption = "ãνϴ"
''    End If
''
''
''End If
'========================================================================================================================



End Sub



Private Sub TreeView7_Click()

If Not UserForm1.TreeView7.SelectedItem.Parent Is Nothing And UserForm1.TreeView7.SelectedItem.ForeColor <> RGB(255, 123, 0) Then
UserForm1.TreeView7.SelectedItem.ForeColor = RGB(255, 123, 0)
UserForm1.TreeView7.SelectedItem.Parent.ForeColor = RGB(255, 0, 0)
End If

End Sub

Private Sub TreeView7_DblClick()
    Dim childNode As Node
    Dim allChildrenBlack As Boolean
    Dim parentNode As Node
        ' õ   RGB(0, 0, 0) 
        UserForm1.TreeView7.SelectedItem.ForeColor = RGB(0, 0, 0)
        
        
    ' õ 尡 θ 带  ִ Ȯ
    If Not UserForm1.TreeView7.SelectedItem.Parent Is Nothing Then
        ' θ 带 
        Set parentNode = UserForm1.TreeView7.SelectedItem.Parent
        
        '  ڽ   RGB(0,0,0) Ȯ
        allChildrenBlack = True
        Set childNode = parentNode.Child
        Do While Not childNode Is Nothing
            If childNode.ForeColor <> RGB(0, 0, 0) Then
                allChildrenBlack = False
                Exit Do
            End If
            Set childNode = childNode.Next
        Loop
        
        ' ڽ 尡  RGB(0,0,0)  θ 嵵 RGB(0,0,0) 
        If allChildrenBlack Then
            parentNode.ForeColor = RGB(0, 0, 0)
            parentNode.Expanded = False
            
        End If


    End If
End Sub

Private Sub UserForm_Initialize()
    LoadTreeViewWithData1
    LoadTreeViewWithData2
    LoadTreeViewWithData5
    LoadTreeViewWithData6
    LoadTreeViewWithData7
    SetInitialNodeStates1
    SetInitialNodeStates2
    AddListView1Columns
    AddListView2Columns
    AddListView3Columns
    TESTcheckbox
    Combo1
    Combo2
    combo4
    combo6
    combo7
    TextBox1 = Format(Now(), "YYYY-MM-DD")
    L4_Set
End Sub




Private Sub AddListView1Columns()
    ' Ʈ ÷ ߰
    With ListView1
        .ColumnHeaders.Clear '  ÷ 
        .View = lvwReport ' Report  
        .Gridlines = True
        '  ÷ ʺ Ϸ ʿ信  Width Ӽ cn  ֽϴ.
        .ColumnHeaders.Add , , "Ƿ", 100
        .ColumnHeaders.Add , , "ä", 100
        .ColumnHeaders.Add , , "Ƿڻ", 100
        .ColumnHeaders.Add , , "÷", 100
        .ColumnHeaders.Add , , "ȸ", 100

    End With
End Sub
Private Sub AddListView2Columns()
    ' Ʈ ÷ ߰
    With ListView2
        .ColumnHeaders.Clear '  ÷ 
        .View = lvwReport ' Report  
        .Gridlines = True
        '  ÷ ʺ Ϸ ʿ信  Width Ӽ   ֽϴ.
        .ColumnHeaders.Add , , "÷ä1", 100
        .ColumnHeaders.Add , , "÷ä2", 100
        .ColumnHeaders.Add , , "", 100
        .ColumnHeaders.Add , , "", 100
        .ColumnHeaders.Add , , "м", 100

    End With
End Sub
Private Sub AddListView3Columns()
    ' Ʈ ÷ ߰
    With ListView3
        .ColumnHeaders.Clear '  ÷ 
        .View = lvwReport ' Report  
        .Gridlines = True
        '  ÷ ʺ Ϸ ʿ信  Width Ӽ   ֽϴ.
        .ColumnHeaders.Add , , "Ƿ׸", 130
        .ColumnHeaders.Add , , "м", 70
        .ColumnHeaders.Add , , "м", 170
        .ColumnHeaders.Add , , "м", 100
        .ColumnHeaders.Add , , "", 50
        .ColumnHeaders.Add , , "Method NO", 20
        .ColumnHeaders.Add , , "instrument NO", 20
        .ColumnHeaders.Add , , "м", 20
    End With
End Sub

Sub ã()
    Dim X As Integer
    Dim XT As Range
    Dim T As Range, TR As Range
    
     = ListView2.ListItems(1).ListSubItems(2).text
    Set T = Sheets("ǥ").Rows(2).Find(what:=, lookat:=xlWhole)
    
    If Not T Is Nothing Then
        For r = 1 To ListView3.ListItems.Count
         Set TR = Sheets("ǥ").Columns(1).Find(what:=ListView3.ListItems(r).text, lookat:=xlWhole)
         If Not TR Is Nothing Then
          ListView3.ListItems(r).ListSubItems(4).text = Sheets("ǥ").Cells(TR.row, T.Column)
         End If
        Next r
    End If
    

End Sub

Sub мã()
    Dim X As Integer
    Dim XT As Range
    Dim T As Range, TR As Range
    
     = ListView2.ListItems(1).ListSubItems(2).text
    Set T = Sheets("ǥ").Rows(2).Find(what:=, lookat:=xlWhole)
    
    If Not T Is Nothing Then
        For r = 1 To ListView3.ListItems.Count
         Set TR = Sheets("ǥ").Columns(1).Find(what:=ListView3.ListItems(r).text, lookat:=xlWhole)
         If Not TR Is Nothing Then
          ListView3.ListItems(r).ListSubItems(4).text = Sheets("ǥ").Cells(TR.row, T.Column)
         End If
        Next r
    End If
    

End Sub

Sub ()
On Error Resume Next

If ActiveSheet.Name = "Ϻ" Then
    SHN = "Ϻ"
    '=-=-=-=-==--=-=-=-=-=-=-=
    X = UserForm1.ListView1.ListItems(1).ListSubItems(2)
    xR = Sheets("").Columns("H").Find(what:=X, lookat:=xlWhole).row
    
    Sheets(SHN).Cells(2, "D") = Sheets("").Cells(xR, "B") 'ȣ
    Sheets(SHN).Cells(2, "I") = Sheets("").Cells(xR, "E") 'ü
    
    Sheets(SHN).Cells(3, "D") = Sheets("").Cells(xR, "C") '
    Sheets(SHN).Cells(3, "I") = Sheets("").Cells(xR, "F") '
    
    Sheets(SHN).Cells(4, "D") = Sheets("").Cells(xR, "D") 'ǥ
    Sheets(SHN).Cells(4, "I") = Sheets("").Cells(xR, "G") 'ǰ
    
    Sheets(SHN).Cells(5, "D") = UserForm1.ListView1.ListItems(1).ListSubItems(4) 'ȯ=ȸ
    Sheets(SHN).Cells(6, "D") = " Ǵ "
    Sheets(SHN).Cells(7, "D") = UserForm1.ListView1.ListItems(1).ListSubItems(3)
    Sheets(SHN).Cells(8, "D") = UserForm1.ListView3.ListItems(1).text & " " & ListView3.ListItems.Count - 1 & "" & "(Ʒ м  ׸ )"
    Sheets(SHN).Cells(9, "D") = "P:4L G:4L"
    '======================================================= ̿  ִ Ȯ
    Dim itemExists As Boolean
    itemExists = False
    Dim index As Long
    Dim item As ListItem
    For Each item In ListView3.ListItems
        index = index + 1
        If item.text = "̿³(pH)" Then
            itemExists = True
            Exit For
        End If
    Next item
    
    If itemExists Then
       Sheets(SHN).Cells(10, "D") = "׸ : pH" & ListView3.ListItems(index).ListSubItems(1).text
    Else
       Sheets(SHN).Cells(10, "D") = ""
    End If
    '======================================================= ̿  ִ Ȯ
    Sheets(SHN).Cells(11, "D") = UserForm1.ListView1.ListItems(1).text
    
    If UserForm1.ListView2.ListItems(1).text <> "" Then
    Sheets(SHN).Cells(11, "I") = UserForm1.ListView2.ListItems(1).text & ", " & UserForm1.ListView2.ListItems(1).ListSubItems(1).text
    Else
    Sheets(SHN).Cells(11, "I") = ""
    
    
    
    
    End If
    
    Sheets(SHN).Range("B13:J72") = ""
    
    For Each Data In ListView3.ListItems
    r = r + 1
    Sheets(SHN).Cells(r + 12, "B") = Data
    Sheets(SHN).Cells(r + 12, "D") = ListView3.ListItems(r).ListSubItems(4)
        X = Sheets("DB").Columns("s").Find(what:=UserForm1.ListView3.ListItems(r).text, lookat:=xlWhole).row

    If Not UserForm1.ListView3.ListItems(r).ListSubItems(1) = "Ұ" Then
       Sheets(SHN).Cells(r + 12, "F") = Round(UserForm1.ListView3.ListItems(r).ListSubItems(1), Sheets("DB").Cells(X, "T"))
    Else
       Sheets(SHN).Cells(r + 12, "F") = UserForm1.ListView3.ListItems(r).ListSubItems(1)
    End If
    
    Sheets(SHN).Cells(r + 12, "H") = ListView3.ListItems(r).ListSubItems(2)
    Next Data
    
    Sheets(SHN).Cells(73, "D") = ListView1.ListItems(1).ListSubItems(1) & " ~ " & ListView2.ListItems(1).ListSubItems(4)
    Sheets(SHN).Cells(77, "A") = Format(CDate(ListView2.ListItems(1).ListSubItems(4)), "YYYY MM DD")
    
    If UserForm1.ListView3.ListItems.Count >= 23 Then
    Sheets(SHN).Rows("35:72").Hidden = False
    Else
    Sheets(SHN).Rows("35:72").Hidden = True
    
    End If
'=-=-=-=-==--=-=-=-=-=-=-=
End If

End Sub


Function IsInListView(ByVal NodeText As String, ByVal parentNodeText As String) As Boolean
    Dim i As Integer
    Dim ListItem As ListItem
    
    IsInListView = False
    For i = 1 To ListView4.ListItems.Count
        Set ListItem = ListView4.ListItems(i)
        If ListItem.SubItems(1) = NodeText And ListItem.SubItems(2) = parentNodeText Then
            IsInListView = True
            Exit Function
        End If
    Next i
End Function

