Attribute VB_Name = "ظӽ"
Sub ׸񺰹ã()
Dim TX As Worksheet
Dim SX As Worksheet
Dim TR As Range

 Set TX = Sheets("Ƿ")
 Set SX = Sheets("ǥ ")
 
For c = 1 To 19 Step (2)
        For r = 2 To 242
         
         If Cells(r, c) <> "" Then
         
          Set TR = TX.Columns(6).Find(what:=SX.Cells(r, c), lookat:=xlWhole)
          
          
           If Not TR Is Nothing Then
                SX.Cells(r, c + 1) = TX.Cells(TR.row, "J")
           Else
                SX.Cells(r, c + 1) = "ãߴ Ф"
           End If
           
         End If
         
        Next r
Next c

End Sub
