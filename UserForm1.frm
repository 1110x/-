VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "리뉴어스-수질분석센터"
   ClientHeight    =   10815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16620
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '소유자 가운데
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
견적종류변경
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
    
    ComboBox2.Clear ' 콤보박스3 초기화
    
    Set ws = ThisWorkbook.Sheets("업체담당자")
    selectedValue = ComboBox3.Value
    
    ' 선택된 값에 해당하는 행의 E열 값들을 콤보박스3에 추가
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    Set rng = ws.Range("A2:A" & lastRow)
    
    For Each cell In rng
        If cell.Value = selectedValue Then
            ComboBox2.AddItem cell.Offset(0, 2).Value & " " & cell.Offset(0, 4).Value & " " & cell.Offset(0, 3).Value ' E열 값을 추가
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
분장항목Combo
End Sub

Private Sub CommandButton1_Click()   '====================-=-=-=-=-=-=-=-=-=-=-=- 측정인 시작
    driver.Start "edge"
    X = 13 - Sheets("측정DB").Cells(1, "R")
    ID = Sheets("측정DB").Cells(X, "P")
    PS = Sheets("측정DB").Cells(X, "Q")

    driver.Get "https://측정인.kr/login.go"
    driver.FindElementById("user_email").SendKeys ID
    driver.FindElementById("login_pwd_confirm").SendKeys PS
    driver.FindElementById("login").Click
    driver.Get "https://측정인.kr/ms/field_water.do" '
End Sub

Private Sub CommandButton13_Click()
For r = UserForm1.ListBox1.ListCount - 1 To 0 Step (-1)

UserForm1.ListBox1.RemoveItem (r)
Next r
합계금액
End Sub

Private Sub CommandButton14_Click()
견적서입력
End Sub

Private Sub CommandButton15_Click()
Application.ScreenUpdating = False
Sheets("분석의뢰 입력").Range("A3:BK100") = ""

Sheets("분석의뢰 입력").Cells(1, "B") = UserForm1.TextBox8.text ' 시료명
Sheets("분석의뢰 입력").Cells(1, "F") = UserForm1.TextBox9.text ' 견적발행일
Sheets("분석의뢰 입력").Cells(1, "K") = UserForm1.TextBox10.text ' 시료채취일

For r = 0 To UserForm1.ListBox2.ListCount - 1
 Sheets("분석의뢰 입력").Cells(r + 3, "A") = UserForm1.ListBox2.List(r, 0) '구분 순번
 Sheets("분석의뢰 입력").Cells(r + 3, "B") = UserForm1.ListBox2.List(r, 1) '시료명 입력
Next r
의뢰항목체크
Application.ScreenUpdating = True

End Sub

Private Sub CommandButton16_Click()
UserForm1.ListBox2.Clear
End Sub

Private Sub CommandButton17_Click()
작업시작

End Sub

Private Sub CommandButton18_Click()
작업시작2
End Sub

Private Sub CommandButton3_Click() ' 환경기술인(담당자) 입력 -------------------------PAGE1
If driver.FindElementById("ui-id-1").Attribute("aria-expanded") = True Then
driver.FindElementByXPath("//*[@id=""edit_env_psic_name""]").Clear.SendKeys ListView1.ListItems(1).ListSubItems(4).text
End If
End Sub
Private Sub CommandButton4_Click() '분석시작일자  -------------------------PAGE1
If driver.FindElementById("ui-id-1").Attribute("aria-expanded") = True Then
    Y = Left(ListView1.ListItems(1).ListSubItems(1).text, 4)
    M = Right(Left(ListView1.ListItems(1).ListSubItems(1).text, 7), 2)
    D = Right(Left(ListView1.ListItems(1).ListSubItems(1).text, 10), 2)
    DATEX = Format(Y, "0000") & "-" & Format(M, "00") & "-" & Format(D, "00")
    script1 = "var input = document.getElementById('edit_meas_start_dt');" & "input.value = '" & DATEX & "';" & "input.blur();"
    driver.ExecuteScript script1
End If
End Sub
Private Sub CommandButton5_Click() '계약선택     -------------------------PAGE1
    Dim keys As New Selenium.keys
    Dim X As Range
    Set X = Sheets("계약정보").Columns(8).Find(what:=ListView1.ListItems(1).ListSubItems(2).text, lookat:=xlWhole)
If driver.FindElementById("ui-id-1").Attribute("aria-expanded") = True Then
    If Not X Is Nothing Then
        ' 필요한 동작 수행
        driver.FindElementByXPath("//*[@id=""wid-id-1""]/div/div[2]/div/fieldset/div[6]/section/label[2]/span").Click
        driver.FindElementByXPath("/html/body/span/span/span[1]/input").Click
        driver.FindElementByXPath("/html/body/span/span/span[1]/input").SendKeys Left(Sheets("계약정보").Cells(X.row, "J").text, 6)
        driver.FindElementByXPath("/html/body/span/span/span[1]/input").SendKeys (keys.Enter)
    End If
End If
End Sub

Private Sub CommandButton6_Click() '분석(의뢰)항목 입력  -------------------------PAGE1
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

' JavaScript 코드를 실행
driver.ExecuteScript jsScript

End Sub
Private Sub CommandButton7_Click()  '분석(의뢰)항목 삭제동작  -------------------------PAGE1
Dim itemCountScript As String
itemCountScript = "return document.querySelectorAll('#wid-id-1 div div:nth-child(2) fieldset div:nth-child(9) section div span span:nth-child(1) span ul li span').length;"

' 스크립트 실행
Dim itemCount As Integer
itemCount = driver.ExecuteScript(itemCountScript)
For i = itemCount To 1 Step (-1)
    clickScript = "document.querySelector('#wid-id-1 div div:nth-child(2) fieldset div:nth-child(9) section div span span:nth-child(1) span ul li:nth-child(" & i & ") span').click();"
    driver.ExecuteScript clickScript
Next i

driver.FindElementByXPath("//*[@id=""wid-id-1""]/div/div[2]/div/fieldset/div[9]/section/div/span/span[1]/span/ul").Click

End Sub
Private Sub CommandButton8_Click() '분석기간     -------------------------PAGE3
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

Private Sub CommandButton9_Click()  '==============================채수인원 입력   PAGE1
    Dim jsScript As String
    Dim x1 As Range, x2 As Range

    ' "edit_emp_id" 요소가 표시되는지 확인
If driver.FindElementById("ui-id-1").Attribute("aria-expanded") = True Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' 기존 채수인원 삭제
    Dim element As Object
    Dim liElements As Object
    Dim liCount As Integer

    ' XPath로 요소 찾기
    Set element = driver.FindElementByXPath("//*[@id=""wid-id-4""]/div/div[2]/div/fieldset/div[2]/section[2]/span/span[1]/span/ul")

    ' 찾은 요소에서 li 태그들 찾기
    Set liElements = element.FindElementsByTag("li")

    ' li 태그의 개수 확인
    liCount = liElements.Count

    If liCount > 1 Then
    For r = liCount To 2 Step (-1)
     driver.FindElementByXPath("//*[@id=""wid-id-4""]/div/div[2]/div/fieldset/div[2]/section[2]/span/span[1]/span/ul/li[" & r - 1 & "]/span").Click
    Next r
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' ListView2에 값이 있고 0이 아닌 경우에만 처리
        If ListView2.ListItems(1).text <> "0" And ListView2.ListItems(1).ListSubItems(1).text <> "0" Then
            ' JavaScript 스크립트 초기화
            jsScript = "var selectElement = document.getElementById('edit_emp_id');"

            ' ListView2의 값에 해당하는 행 찾기
            Set x1 = Sheets("측정DB").Columns(14).Find(what:=ListView2.ListItems(1).text, lookat:=xlWhole).Rows
            Set x2 = Sheets("측정DB").Columns(14).Find(what:=ListView2.ListItems(1).ListSubItems(1).text, lookat:=xlWhole).Rows

            ' JavaScript 코드 작성
            jsScript1 = jsScript & "selectElement.options[" & x1.row - 1 & "].selected = true;"
            jsScript2 = jsScript & "selectElement.options[" & x2.row - 1 & "].selected = true;"

            ' JavaScript 코드에 이벤트 디스패치 추가
            jsScript1 = jsScript1 & "selectElement.dispatchEvent(new Event('change'));"
            jsScript2 = jsScript2 & "selectElement.dispatchEvent(new Event('change'));"

            ' JavaScript 코드 실행
            driver.ExecuteScript jsScript1
            driver.ExecuteScript jsScript2
        End If

        ' XPath를 사용하여 요소 클릭
        driver.FindElementByXPath("//*[@id=""wid-id-1""]/div/div[2]/div/fieldset/div[9]/section").Click
End If
End Sub
Private Sub CommandButton2_Click() '채취용기  -------------------------PAGE2

If driver.FindElementById("ui-id-2").Attribute("aria-expanded") = True Then
driver.FindElementById("samp_vesl_desc").Clear.SendKeys "P:2, G:1"
End If

End Sub

Private Sub CommandButton10_Click() 'PAGE2 시료채취 지점명 설정 -------------------------PAGE2

If driver.FindElementById("ui-id-2").Attribute("aria-expanded") = True Then
 driver.FindElementByXPath("//*[@id=""edit_meas_loc_desc_1""]").Clear.SendKeys ListView1.ListItems(1).ListSubItems(3).text
End If

End Sub
Private Sub CommandButton11_Click() 'PAGE4 분석자료 입력
    If driver.FindElementById("ui-id-4").Attribute("aria-expanded") = True Then
        Dim trElements As Object
        Dim desiredValue1 As String, desiredValue2 As String
        Dim script As String
        Dim startTime As Double
        Set trElements = driver.FindElementsByXPath("//*[@id='tbAnze']/tbody/tr[contains(@class, 'tr_')]")
        
        Dim trCount As Integer
        trCount = trElements.Count
        
'        Debug.Print "총 " & trCount & "개의 tr 태그를 찾았습니다."

        Y = Left(ListView1.ListItems(1).ListSubItems(1).text, 4)
        M = Right(Left(ListView1.ListItems(1).ListSubItems(1).text, 7), 2)
        D = Right(Left(ListView1.ListItems(1).ListSubItems(1).text, 10), 2)
        DATES = Format(Y, "0000") & "-" & Format(M, "00") & "-" & Format(D, "00")
        DATEE = Format(Year(Now), "0000") & "-" & Format(Month(Now), "00") & "-" & Format(Day(Now), "00")
            
        For i = 1 To trCount
            For X = 1 To ListView3.ListItems.Count
                If ListView3.ListItems(X).text = driver.FindElementById("meas_item_name_" & i).text Then
                  startTime = Timer
                  
                    driver.FindElementById("allow_val_" & i).Clear.SendKeys ListView3.ListItems(X).ListSubItems(4).text  '허용기준치
                    driver.FindElementById("anze_val_" & i).Clear.SendKeys ListView3.ListItems(X).ListSubItems(1).text   '분석결과

                    desiredValue1 = ListView3.ListItems(X).ListSubItems(5).text                                          '분석방법(Method)
                    desiredValue2 = ListView3.ListItems(X).ListSubItems(6).text                                          '분석장비

                    Set selectElement1 = driver.FindElementById("anze_mthd_" & i)                                        '분석방법 입력
                    selectElement1.AsSelect.SelectByValue desiredValue1

                    Set selectElement2 = driver.FindElementById("anze_equip_no_1" & i)                                   '분석장비 입력
                    selectElement2.AsSelect.SelectByValue desiredValue2
                    
                    Set selectElement = driver.FindElementByName("anze_login_id_1" & i).AsSelect                         '분석인력 (업무분장에서 찾아서) 입력
                    selectElement.SelectByValue Sheets("측정DB").Cells(Sheets("측정DB").Columns(14).Find(what:=ListView3.ListItems(X).ListSubItems(7).text, lookat:=xlWhole).row, "R")
                    

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

For i = 1 To 60 ' 3은 체크박스의 개수입니다. 필요에 따라 조절하세요.
    Me.Controls("Checkbox" & i).Caption = Sheets("항목코드").Cells(i, "J")
Next i

End Sub



Private Sub Label11_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
sellist = UserForm1.ListBox1.ListIndex
UserForm1.ListBox1.RemoveItem (sellist)
    합계금액
End Sub





Private Sub ListBox2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim selectedIndex As Integer
    Dim i As Integer
    
    ' 선택된 항목의 인덱스를 가져옵니다.
    selectedIndex = UserForm1.ListBox2.ListIndex
    
    ' 선택된 항목이 있는 경우에만 삭제합니다.
    If selectedIndex <> -1 Then
        ' 항목 삭제
        UserForm1.ListBox2.RemoveItem selectedIndex
        
        ' 항목 삭제 후 번호를 다시 매기기
        For i = 0 To UserForm1.ListBox2.ListCount - 1
            UserForm1.ListBox2.List(i, 0) = Format(i + 1, "00") ' 첫 번째 열에 번호 매기기
        Next i
    Else
        MsgBox "선택된 항목이 없습니다.", vbExclamation
    End If
End Sub




Private Sub MultiPage1_Change()

x1 = UserForm1.Left
x2 = UserForm1.Top

' UserForm을 (0,0) 위치로 이동
UserForm1.Move 0, 0

' 그 다음에 원래 위치로 이동
UserForm1.Move x1, x2
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub TextBox3_Change()
For r = o To UserForm1.ListBox1.ListCount - 1
UserForm1.ListBox1.List(r, 2) = TextBox3.text
Next r

    합계금액
End Sub

Private Sub TreeView1_DblClick()

붉은색으로물들여라

End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim ws As Worksheet
    Dim 업체명 As Range
    
           Application.ScreenUpdating = False

 Sheets("시험성적서").Range("C1:D7,F3:H7,K1:L7,N3:P7,A10:P41") = ""

   
If Not TreeView1.SelectedItem Is Nothing Then
    If Not TreeView1.SelectedItem.Parent Is Nothing Then
        Debug.Print "자식노드: " & TreeView1.SelectedItem.Parent.text
        Set ws = ThisWorkbook.Sheets("의뢰정보")
        lastRow = Sheets("의뢰정보").Cells(Sheets("의뢰정보").Rows.Count, "A").End(xlUp).row
        For r = 2 To lastRow
         X = "【" & Sheets("의뢰정보").Cells(r, "E").text & "】" & Sheets("의뢰정보").Cells(r, "F").Value '  ws.Cells(i, "C")  & ws.Cells(i, 4).Value)
         If Sheets("의뢰정보").Cells(r, "A") = TreeView1.SelectedItem.Parent.text And X = TreeView1.SelectedItem.text Then

          '===============================================================================================
         ListView1.ListItems.Clear
         With ListView1
        .ColumnHeaders.Clear ' 기존 컬럼 제거
        .Gridlines = True
        .View = lvwReport ' Report 모드로 설정

        ' 각 컬럼의 너비를 조절하려면 필요에 따라 Width 속성을 설정할 수 있습니다.
        .ColumnHeaders.Add , , "의뢰일자", 100
        .ColumnHeaders.Add , , "채취일자", 100
        .ColumnHeaders.Add , , "의뢰사업장", 100
        .ColumnHeaders.Add , , "시료명", 120
        .ColumnHeaders.Add , , "입회자", 100

         End With

         ListView2.ListItems.Clear
         With ListView2
        .ColumnHeaders.Clear ' 기존 컬럼 제거
        .Gridlines = True
        .View = lvwReport ' Report 모드로 설정
        ' 각 컬럼의 너비를 조절하려면 필요에 따라 Width 속성을 설정할 수 있습니다.
        .ColumnHeaders.Add , , "시료채취자1", 100
        .ColumnHeaders.Add , , "시료채취자2", 100
        .ColumnHeaders.Add , , "방류허용기준", 120
        .ColumnHeaders.Add , , "정도보증유무", 100
        .ColumnHeaders.Add , , "분석종료일", 100
         End With

         ListView3.ListItems.Clear
         With ListView3
        .ColumnHeaders.Clear ' 기존 컬럼 제거
        .Gridlines = True
        .View = lvwReport ' Report 모드로 설정
        ' 각 컬럼의 너비를 조절하려면 필요에 따라 Width 속성을 설정할 수 있습니다.
        .ColumnHeaders.Add , , "의뢰항목", 130
        .ColumnHeaders.Add , , "분석결과", 70
        .ColumnHeaders.Add , , "분석방법", 170
        .ColumnHeaders.Add , , "분석장비", 100
        .ColumnHeaders.Add , , "법적기준", 50
        .ColumnHeaders.Add , , "Method NO", 20
        .ColumnHeaders.Add , , "instrument NO", 20
        .ColumnHeaders.Add , , "분석담당자", 20
         End With


         Set item = ListView1.ListItems.Add(1, , Sheets("의뢰정보").Cells(r, "A").Value) '의뢰일자
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "B").Value                '채취일자
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "E").Value                '의뢰사업장 '약칭
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "F").Value                '시료명
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "G").Value                '입회자

         Set item = ListView2.ListItems.Add(1, , Sheets("의뢰정보").Cells(r, "H").Value) '시료채취자-1
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "I").Value                '시료채취자-2
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "J").Value                '방류허용기준
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "K").Value                '정도보증유무
         item.ListSubItems.Add , , Sheets("의뢰정보").Cells(r, "L").Value                '분석종료일

         For T = Range("N1").Column To Range("BT1").Column
'         Sheets("시험성적서").Cells(10, "A") = 1
         
         If Sheets("의뢰정보").Cells(r, T) <> "" Then

         
               G = Sheets("시험성적서").Range("A42").End(xlUp).row + 1
               G2 = Sheets("시험성적서").Range("I42").End(xlUp).row + 1
            If G <= 41 Then
               Sheets("시험성적서").Cells(G, "A") = G - 9
               Sheets("시험성적서").Cells(G, "D") = Sheets("의뢰정보").Cells(1, T)
               GX = Sheets("견적단가").Columns(4).Find(what:=Sheets("의뢰정보").Cells(1, T), lookat:=xlWhole).row
               Sheets("시험성적서").Cells(G, "B") = Sheets("견적단가").Cells(GX, 3)
               Sheets("시험성적서").Cells(G, "E") = Sheets("견적단가").Cells(GX, 2)
               
            Else
               Sheets("시험성적서").Cells(G2, "I") = G2 - 9
               GX = Sheets("견적단가").Columns(4).Find(what:=Sheets("의뢰정보").Cells(1, T), lookat:=xlWhole).row
               Sheets("시험성적서").Cells(G2, "J") = Sheets("견적단가").Cells(GX, 3)
               Sheets("시험성적서").Cells(G2, "L") = Sheets("의뢰정보").Cells(1, T)
               Sheets("시험성적서").Cells(G2, "M") = Sheets("견적단가").Cells(GX, 2)

            End If
            
         Set Titem = ListView3.ListItems.Add(ListView3.ListItems.Count + 1, , Sheets("의뢰정보").Cells(1, T))
                     Titem.ListSubItems.Add , , "-"             'Subitem-1 분석결과
                     Titem.ListSubItems.Add , , "-"             'Subitem-2 분석방법
                     Titem.ListSubItems.Add , , "-"             'Subitem-3 분석장비
                     Titem.ListSubItems.Add , , "-"             'Subitem-4 법적기준
                     Titem.ListSubItems.Add , , "-"             'Subitem-5 Method NO
                     Titem.ListSubItems.Add , , "-"             'Subitem-6 instrument NO
                     Titem.ListSubItems.Add , , "-"             'Subitem-7 분석담당자
         End If
         Next T
         '===============================================================================================

         End If

        Next r

        분석결과불러오기
        방류기준찾기

If ActiveSheet.Name = "수질측정기록부" Then
법정양식
End If

Call 시험법


        ListView3.ColumnHeaders(1).text = "의뢰항목 (총 " & ListView3.ListItems.Count & "건) "
    Else
        Label1.Caption = "못찾겄습니다"
    End If


End If
       Set 업체명 = Sheets("계약정보").Columns("H").Find(what:=UserForm1.ListView1.ListItems(1).ListSubItems(2), lookat:=xlWhole)
       If Sheets("시험성적서").Cells(10, "I") = "" Then
           Sheets("시험성적서").PageSetup.PrintArea = "A1:H44"
           
           If Not 업체명 Is Nothing Then
           
           
           Sheets("시험성적서").Range("C3,K3") = Sheets("계약정보").Cells(업체명.row, "B")
           Sheets("시험성적서").Range("C4,K4") = Sheets("계약정보").Cells(업체명.row, "D")
           Sheets("시험성적서").Range("C5,K5") = UserForm1.ListView1.ListItems(1).ListSubItems(4)
           Sheets("시험성적서").Range("C7,K7") = UserForm1.ListView1.ListItems(1).ListSubItems(3)
           
           
           Sheets("시험성적서").Range("F3,N3") = UserForm1.ListView1.ListItems(1).ListSubItems(1)
           Sheets("시험성적서").Range("F4,N4") = UserForm1.ListView2.ListItems(1).text & ", " & UserForm1.ListView2.ListItems(1).ListSubItems(1)
           
           If UserForm1.ListView2.ListItems(1).ListSubItems(3) <> "O" Then
           Sheets("시험성적서").Range("F7,N7") = "참고용"
           Else
           Sheets("시험성적서").Range("F7,N7") = "정도보증 적용"
           End If
           
           End If
           
         Else
           Sheets("시험성적서").PageSetup.PrintArea = "A1:H44,I1:P44"
           
           If Not 업체명 Is Nothing Then
           
           
           Sheets("시험성적서").Range("C3,K3") = Sheets("계약정보").Cells(업체명.row, "B")
           Sheets("시험성적서").Range("C4,K4") = Sheets("계약정보").Cells(업체명.row, "D")
           Sheets("시험성적서").Range("C5,K5") = UserForm1.ListView1.ListItems(1).ListSubItems(4)
           Sheets("시험성적서").Range("C7,K7") = UserForm1.ListView1.ListItems(1).ListSubItems(3)
           
           
           Sheets("시험성적서").Range("F3,N3") = UserForm1.ListView1.ListItems(1).ListSubItems(1)
           Sheets("시험성적서").Range("F4,N4") = UserForm1.ListView2.ListItems(1).text & ", " & UserForm1.ListView2.ListItems(1).ListSubItems(1)
           
           If UserForm1.ListView2.ListItems(1).ListSubItems(3) <> "O" Then
           Sheets("시험성적서").Range("F7,N7") = "참고용"
           Else
           Sheets("시험성적서").Range("F7,N7") = "정도보증 적용"
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

    ' UserForm1의 ListBox1 및 ComboBox1 참조
    Set listBox = Me.ListBox1
    Set comboBox = Me.ComboBox1
    Set ws = ThisWorkbook.Sheets("견적단가")
    
   If UserForm1.TreeView2.SelectedItem.text = "특정수질유해물질" Then
   특정수질견적일괄입력
   End If
   
   
    
    ' ComboBox에서 선택한 항목에 해당하는 컬럼 찾기
    selectedColumn = 0
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If ws.Cells(1, i).Value = comboBox.Value Then
            selectedColumn = i
            Exit For
        End If
    Next i
    
    If selectedColumn = 0 Then
        MsgBox "ComboBox에서 선택한 항목에 해당하는 컬럼을 찾을 수 없습니다.", vbExclamation
        Exit Sub
    End If
    
    ' ListBox 초기화
'    listBox.Clear
    
    ' 선택된 항목이 있는지 확인
    If Not UserForm1.TreeView2.SelectedItem Is Nothing Then
        With UserForm1.TreeView2.SelectedItem
            ' 선택된 항목이 부모 노드인지 확인
            If .Children > 0 Then
                ' 부모 노드의 모든 하위 노드 출력
                Set childNode = .Child
                
                Do While Not childNode Is Nothing
                    ' 중복 확인
                    IsDuplicate = False
                    For i = 0 To listBox.ListCount - 1
                        If listBox.List(i, 1) = childNode.text Then
                            IsDuplicate = True
                            Exit For
                        End If
                    Next i
                    
                    ' 중복이 아니면 ListBox에 부모 노드와 자식 노드 및 금액 추가
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
                ' 자식 노드일 경우 부모 노드와 선택된 자식 노드 및 금액 추가
                If Not .Parent Is Nothing Then
                    Set parentNode = .Parent
                    ' 중복 확인
                    IsDuplicate = False
                    For i = 0 To listBox.ListCount - 1
                        If listBox.List(i, 1) = .text Then
                            IsDuplicate = True
                            Exit For
                        End If
                    Next i
                    
                    ' 중복이 아니면 ListBox에 부모 노드와 자식 노드 및 금액 추가
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
    
    합계금액
    
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
        ' 선택된 항목이 최상위 노드인 경우, 부모가 없으므로 Label12를 초기화
        UserForm1.TextBox9 = "No Parent"
    End If
Else
    ' 선택된 항목이 없을 경우 Label12를 초기화
    UserForm1.TextBox9 = "No Item Selected"
End If
 
   
If Not TreeView3.SelectedItem Is Nothing Then
    If Not TreeView3.SelectedItem.Parent Is Nothing Then

        Set ws = ThisWorkbook.Sheets("견적발행정보")
        lastRow = Sheets("견적발행정보").Cells(Sheets("견적발행정보").Rows.Count, "A").End(xlUp).row
        For r = 2 To lastRow
         Z = "【" & Sheets("견적발행정보").Cells(r, "C").text & "】" & Sheets("견적발행정보").Cells(r, "H").Value '  ws.Cells(i, "C")  & ws.Cells(i, 4).Value)
         
         If Sheets("견적발행정보").Cells(r, "A") = TreeView3.SelectedItem.Parent.text And Z = TreeView3.SelectedItem.text Then

          
          Set listBox = UserForm1.ListBox1
          listBox.Clear
          
          UserForm1.TextBox1 = Sheets("견적발행정보").Cells(r, "A")
          
          Set 약칭 = Sheets("계약정보").Columns(8).Find(what:=Sheets("견적발행정보").Cells(r, "C").text, lookat:=xlWhole)
          If Not 약칭 Is Nothing Then
            UserForm1.ComboBox3.ListIndex = 약칭.row - 2
          End If
          
          UserForm1.ComboBox4.Value = Sheets("견적발행정보").Cells(r, "K") '#### 요고 수정해야 함..ㅋㅋ
          UserForm1.TextBox4.Value = Sheets("견적발행정보").Cells(r, "G")
          UserForm1.TextBox5.Value = Sheets("견적발행정보").Cells(r, "F")
          UserForm1.TextBox2.Value = Sheets("견적발행정보").Cells(r, "H")
          UserForm1.ComboBox2.Value = Sheets("견적발행정보").Cells(r, "E")
          
                        Total = 0
                        For X = 13 To 193 Step (3)

                        If Sheets("견적발행정보").Cells(r, X) <> "" Then
                        amount = 0
                       
                        Z = Sheets("견적단가").Columns(4).Find(what:=Sheets("견적발행정보").Cells(1, X), lookat:=xlWhole).row
                        
                        listBox.AddItem
                        listBox.List(listBox.ListCount - 1, 0) = Sheets("견적단가").Cells(Z, "C")    '의뢰/분석항목 대구분
                        listBox.List(listBox.ListCount - 1, 1) = Sheets("견적발행정보").Cells(1, X)  '의뢰/분석항목 소구분
                        listBox.List(listBox.ListCount - 1, 2) = Sheets("견적발행정보").Cells(r, X)  '의뢰/분석항목 수량
                        listBox.List(listBox.ListCount - 1, 3) = Format(Sheets("견적발행정보").Cells(r, X + 1), "#,###") '의뢰/분석항목 단가
                        amount = Sheets("견적발행정보").Cells(r, X) * Sheets("견적발행정보").Cells(r, X + 1)
                        Total = amount + Total
                        
                        listBox.List(listBox.ListCount - 1, 4) = Format(amount, "#,###")
                        
                        End If
                        
                        
                        Next X
                         

         
         
          
         End If

        Next r


    Else
        Label1.Caption = "못찾겄습니다"
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

  UserForm1.Label5.Caption = UserForm1.ListBox1.List(0, 1) & "포함 " & UserForm1.ListBox1.ListCount & "종 【" & Format(Total, "#,###원】")
Else
  UserForm1.Label5.Caption = "견적건수/총액"
End If
'=========================================================================================================================

트리뷰3클릭체크

End Sub
Private Sub TreeView4_Click()
UserForm1.TreeView4.SelectedItem.ForeColor = RGB(255, 0, 0)
의뢰리스트이동
End Sub
Private Sub TreeView5_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    Dim ListItem As ListItem
    Dim NodeText As String
    Dim ParentText As String
    Dim childNode As MSComctlLib.Node
    Dim IsDuplicate As Boolean
    Dim ItemIndex As Integer
    
    

    
    
    ' 현재 ListView에 있는 아이템 수에 따라 새로운 번호 지정
    ItemIndex = ListView4.ListItems.Count + 1
    
    ' 부모 노드인 경우 자식 노드 전체를 처리
    If Node.Children > 0 Then
        ParentText = Node.text
        Set childNode = Node.Child
        
        Do While Not childNode Is Nothing
            NodeText = childNode.text
            IsDuplicate = False
            
            ' 중복 체크
            For i = 1 To ListView4.ListItems.Count
                If ListView4.ListItems(i).SubItems(1) = ParentText And ListView4.ListItems(i).SubItems(2) = NodeText Then
                    IsDuplicate = True
                    Exit For
                End If
            Next i
            
            ' 중복되지 않는 경우에만 추가
            If Not IsDuplicate Then
                Set ListItem = ListView4.ListItems.Add(, , ItemIndex)
                ListItem.SubItems(1) = ParentText
                ListItem.SubItems(2) = NodeText
                ListItem.SubItems(3) = "기준"
                ItemIndex = ItemIndex + 1
                분석결과입력리스트의법적기준
           End If
            
            ' 다음 자식 노드로 이동
            Set childNode = childNode.Next
        Loop
    Else
        ' 자식 노드가 없는 경우 (리프 노드)
        If Not Node.Parent Is Nothing Then
            ParentText = Node.Parent.text
        Else
            ParentText = Node.text
        End If
        
        NodeText = Node.text
        IsDuplicate = False
        
        ' 중복 체크
        For i = 1 To ListView4.ListItems.Count
            If ListView4.ListItems(i).SubItems(1) = ParentText And ListView4.ListItems(i).SubItems(2) = NodeText Then
                IsDuplicate = True
                Exit For
            End If
        Next i
        
        ' 중복되지 않는 경우에만 추가
        If Not IsDuplicate Then
            Set ListItem = ListView4.ListItems.Add(, , ItemIndex)
            ListItem.SubItems(1) = ParentText
            ListItem.SubItems(2) = NodeText
            ListItem.SubItems(3) = "기준"
            분석결과입력리스트의법적기준
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
        ' 선택된 항목이 최상위 노드인 경우, 부모가 없으므로 Label12를 초기화
        UserForm1.TextBox9 = UserForm1.TreeView6.SelectedItem.text

    End If
Else
    ' 선택된 항목이 없을 경우 Label12를 초기화
    UserForm1.TextBox9 = "No Item Selected"
End If
 
단축어찾기

 
   
''If Not TreeView3.SelectedItem Is Nothing Then
''    If Not TreeView3.SelectedItem.Parent Is Nothing Then
''
''        Set ws = ThisWorkbook.Sheets("견적발행정보")
''        lastRow = Sheets("견적발행정보").Cells(Sheets("견적발행정보").Rows.Count, "A").End(xlUp).Row
''        For r = 2 To lastRow
''         Z = "【" & Sheets("견적발행정보").Cells(r, "C").Text & "】" & Sheets("견적발행정보").Cells(r, "H").Value '  ws.Cells(i, "C")  & ws.Cells(i, 4).Value)
''
''         If Sheets("견적발행정보").Cells(r, "A") = TreeView3.SelectedItem.Parent.Text And Z = TreeView3.SelectedItem.Text Then
''
''
''          Set listBox = UserForm1.ListBox1
''          listBox.Clear
''
''          UserForm1.TextBox1 = Sheets("견적발행정보").Cells(r, "A")
''
''          Set 약칭 = Sheets("계약정보").Columns(8).Find(what:=Sheets("견적발행정보").Cells(r, "C").Text, lookat:=xlWhole)
''          If Not 약칭 Is Nothing Then
''            UserForm1.ComboBox3.ListIndex = 약칭.Row - 2
''          End If
''
''          UserForm1.ComboBox4.Value = Sheets("견적발행정보").Cells(r, "K") '#### 요고 수정해야 함..ㅋㅋ
''          UserForm1.TextBox4.Value = Sheets("견적발행정보").Cells(r, "G")
''          UserForm1.TextBox5.Value = Sheets("견적발행정보").Cells(r, "F")
''          UserForm1.TextBox2.Value = Sheets("견적발행정보").Cells(r, "H")
''          UserForm1.ComboBox2.Value = Sheets("견적발행정보").Cells(r, "E")
''
''                        Total = 0
''                        For X = 13 To 193 Step (3)
''
''                        If Sheets("견적발행정보").Cells(r, X) <> "" Then
''                        amount = 0
''
''                        Z = Sheets("견적단가").Columns(4).Find(what:=Sheets("견적발행정보").Cells(1, X), lookat:=xlWhole).Row
''
''                        listBox.AddItem
''                        listBox.List(listBox.ListCount - 1, 0) = Sheets("견적단가").Cells(Z, "C")    '의뢰/분석항목 대구분
''                        listBox.List(listBox.ListCount - 1, 1) = Sheets("견적발행정보").Cells(1, X)  '의뢰/분석항목 소구분
''                        listBox.List(listBox.ListCount - 1, 2) = Sheets("견적발행정보").Cells(r, X)  '의뢰/분석항목 수량
''                        listBox.List(listBox.ListCount - 1, 3) = Format(Sheets("견적발행정보").Cells(r, X + 1), "#,###") '의뢰/분석항목 단가
''                        amount = Sheets("견적발행정보").Cells(r, X) * Sheets("견적발행정보").Cells(r, X + 1)
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
''        Label1.Caption = "못찾겄습니다"
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
        ' 선택된 노드의 색상을 RGB(0, 0, 0)으로 변경
        UserForm1.TreeView7.SelectedItem.ForeColor = RGB(0, 0, 0)
        
        
    ' 선택된 노드가 부모 노드를 가지고 있는지 확인
    If Not UserForm1.TreeView7.SelectedItem.Parent Is Nothing Then
        ' 부모 노드를 가져옴
        Set parentNode = UserForm1.TreeView7.SelectedItem.Parent
        
        ' 모든 자식 노드의 색상이 RGB(0,0,0)인지 확인
        allChildrenBlack = True
        Set childNode = parentNode.Child
        Do While Not childNode Is Nothing
            If childNode.ForeColor <> RGB(0, 0, 0) Then
                allChildrenBlack = False
                Exit Do
            End If
            Set childNode = childNode.Next
        Loop
        
        ' 자식 노드가 전부 RGB(0,0,0)인 경우 부모 노드도 RGB(0,0,0)으로 변경
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
    ' 리스트뷰의 컬럼 추가
    With ListView1
        .ColumnHeaders.Clear ' 기존 컬럼 제거
        .View = lvwReport ' Report 모드로 설정
        .Gridlines = True
        ' 각 컬럼의 너비를 조절하려면 필요에 따라 Width 속성을 설cn정할 수 있습니다.
        .ColumnHeaders.Add , , "의뢰일자", 100
        .ColumnHeaders.Add , , "채취일자", 100
        .ColumnHeaders.Add , , "의뢰사업장", 100
        .ColumnHeaders.Add , , "시료명", 100
        .ColumnHeaders.Add , , "입회자", 100

    End With
End Sub
Private Sub AddListView2Columns()
    ' 리스트뷰의 컬럼 추가
    With ListView2
        .ColumnHeaders.Clear ' 기존 컬럼 제거
        .View = lvwReport ' Report 모드로 설정
        .Gridlines = True
        ' 각 컬럼의 너비를 조절하려면 필요에 따라 Width 속성을 설정할 수 있습니다.
        .ColumnHeaders.Add , , "시료채취자1", 100
        .ColumnHeaders.Add , , "시료채취자2", 100
        .ColumnHeaders.Add , , "방류허용기준", 100
        .ColumnHeaders.Add , , "정보보증유무", 100
        .ColumnHeaders.Add , , "분석종료일", 100

    End With
End Sub
Private Sub AddListView3Columns()
    ' 리스트뷰의 컬럼 추가
    With ListView3
        .ColumnHeaders.Clear ' 기존 컬럼 제거
        .View = lvwReport ' Report 모드로 설정
        .Gridlines = True
        ' 각 컬럼의 너비를 조절하려면 필요에 따라 Width 속성을 설정할 수 있습니다.
        .ColumnHeaders.Add , , "의뢰항목", 130
        .ColumnHeaders.Add , , "분석결과", 70
        .ColumnHeaders.Add , , "분석방법", 170
        .ColumnHeaders.Add , , "분석장비", 100
        .ColumnHeaders.Add , , "법적기준", 50
        .ColumnHeaders.Add , , "Method NO", 20
        .ColumnHeaders.Add , , "instrument NO", 20
        .ColumnHeaders.Add , , "분석담당자", 20
    End With
End Sub

Sub 방류기준찾기()
    Dim X As Integer
    Dim XT As Range
    Dim T As Range, TR As Range
    
    기준 = ListView2.ListItems(1).ListSubItems(2).text
    Set T = Sheets("방류기준표").Rows(2).Find(what:=기준, lookat:=xlWhole)
    
    If Not T Is Nothing Then
        For r = 1 To ListView3.ListItems.Count
         Set TR = Sheets("방류기준표").Columns(1).Find(what:=ListView3.ListItems(r).text, lookat:=xlWhole)
         If Not TR Is Nothing Then
          ListView3.ListItems(r).ListSubItems(4).text = Sheets("방류기준표").Cells(TR.row, T.Column)
         End If
        Next r
    End If
    

End Sub

Sub 분석담당찾기()
    Dim X As Integer
    Dim XT As Range
    Dim T As Range, TR As Range
    
    기준 = ListView2.ListItems(1).ListSubItems(2).text
    Set T = Sheets("방류기준표").Rows(2).Find(what:=기준, lookat:=xlWhole)
    
    If Not T Is Nothing Then
        For r = 1 To ListView3.ListItems.Count
         Set TR = Sheets("방류기준표").Columns(1).Find(what:=ListView3.ListItems(r).text, lookat:=xlWhole)
         If Not TR Is Nothing Then
          ListView3.ListItems(r).ListSubItems(4).text = Sheets("방류기준표").Cells(TR.row, T.Column)
         End If
        Next r
    End If
    

End Sub

Sub 법정양식()
On Error Resume Next

If ActiveSheet.Name = "수질측정기록부" Then
    SHN = "수질측정기록부"
    '=-=-=-=-==--=-=-=-=-=-=-=
    X = UserForm1.ListView1.ListItems(1).ListSubItems(2)
    xR = Sheets("계약정보").Columns("H").Find(what:=X, lookat:=xlWhole).row
    
    Sheets(SHN).Cells(2, "D") = Sheets("계약정보").Cells(xR, "B") '상호명
    Sheets(SHN).Cells(2, "I") = Sheets("계약정보").Cells(xR, "E") '시설별
    
    Sheets(SHN).Cells(3, "D") = Sheets("계약정보").Cells(xR, "C") '소재지
    Sheets(SHN).Cells(3, "I") = Sheets("계약정보").Cells(xR, "F") '종류별
    
    Sheets(SHN).Cells(4, "D") = Sheets("계약정보").Cells(xR, "D") '대표자
    Sheets(SHN).Cells(4, "I") = Sheets("계약정보").Cells(xR, "G") '생산품
    
    Sheets(SHN).Cells(5, "D") = UserForm1.ListView1.ListItems(1).ListSubItems(4) '환경기술인=입회자
    Sheets(SHN).Cells(6, "D") = "제출 또는 보고용"
    Sheets(SHN).Cells(7, "D") = UserForm1.ListView1.ListItems(1).ListSubItems(3)
    Sheets(SHN).Cells(8, "D") = UserForm1.ListView3.ListItems(1).text & "외 " & ListView3.ListItems.Count - 1 & "건" & "(아래 ⑤측정분석 결과의 항목과 같음)"
    Sheets(SHN).Cells(9, "D") = "P:4L G:4L"
    '======================================================= 수소이온 농도 있는지 확인
    Dim itemExists As Boolean
    itemExists = False
    Dim index As Long
    Dim item As ListItem
    For Each item In ListView3.ListItems
        index = index + 1
        If item.text = "수소이온농도(pH)" Then
            itemExists = True
            Exit For
        End If
    Next item
    
    If itemExists Then
       Sheets(SHN).Cells(10, "D") = "현장측정항목 : pH" & ListView3.ListItems(index).ListSubItems(1).text
    Else
       Sheets(SHN).Cells(10, "D") = ""
    End If
    '======================================================= 수소이온 농도 있는지 확인
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
        X = Sheets("측정DB").Columns("s").Find(what:=UserForm1.ListView3.ListItems(r).text, lookat:=xlWhole).row

    If Not UserForm1.ListView3.ListItems(r).ListSubItems(1) = "불검출" Then
       Sheets(SHN).Cells(r + 12, "F") = Round(UserForm1.ListView3.ListItems(r).ListSubItems(1), Sheets("측정DB").Cells(X, "T"))
    Else
       Sheets(SHN).Cells(r + 12, "F") = UserForm1.ListView3.ListItems(r).ListSubItems(1)
    End If
    
    Sheets(SHN).Cells(r + 12, "H") = ListView3.ListItems(r).ListSubItems(2)
    Next Data
    
    Sheets(SHN).Cells(73, "D") = ListView1.ListItems(1).ListSubItems(1) & " ~ " & ListView2.ListItems(1).ListSubItems(4)
    Sheets(SHN).Cells(77, "A") = Format(CDate(ListView2.ListItems(1).ListSubItems(4)), "YYYY년 MM월 DD일")
    
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

