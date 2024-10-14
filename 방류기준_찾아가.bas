Attribute VB_Name = "방류기준_찾아가"
Sub 단추3_Click()
    Dim TX As Worksheet
    Dim SX As Worksheet
    Dim TR As Range
    Dim r As Long
    
    Set TX = Sheets("방류기준표 정리")
    Set SX = Sheets("분석의뢰 입력")

    ' Loop through each row starting from row 3
    For r = 3 To SX.Cells(2, 2).End(xlDown).row
        ' Find the value in TX that matches the value in SX
        Set TR = TX.Cells.Find(what:=SX.Cells(r, "B"), lookat:=xlWhole)

        ' Check if TR is found
        If Not TR Is Nothing Then
            ' Correctly retrieve the value from the next column in TX
            SX.Cells(r, "C").Value = TX.Cells(TR.row, TR.Column + 1).Value
        End If
    Next r
End Sub
