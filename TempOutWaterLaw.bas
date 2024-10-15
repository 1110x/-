Attribute VB_Name = "방류기준만들기임시"
Sub 항목별방류기준찾기()
Dim TX As Worksheet
Dim SX As Worksheet
Dim TR As Range

 Set TX = Sheets("의뢰정보")
 Set SX = Sheets("방류기준표 정리")
 
For c = 1 To 19 Step (2)
        For r = 2 To 242
         
         If Cells(r, c) <> "" Then
         
          Set TR = TX.Columns(6).Find(what:=SX.Cells(r, c), lookat:=xlWhole)
          
          
           If Not TR Is Nothing Then
                SX.Cells(r, c + 1) = TX.Cells(TR.row, "J")
           Else
                SX.Cells(r, c + 1) = "못찾긋다 ㅠㅠ"
           End If
           
         End If
         
        Next r
Next c

End Sub
