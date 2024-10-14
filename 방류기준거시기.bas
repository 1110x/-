Attribute VB_Name = "ذŽñ"
Sub ã()
    Dim X As Integer
    Dim XT As Range
    Dim T As Range, TR As Range
    Dim TZ As Long
    Dim  As String
    
    ' ListView2   
     = UserForm1.ListView2.ListItems(1).ListSubItems(2).text
    
    ' ǥ Ʈ 2࿡   ã
    Set T = Sheets("ǥ").Rows(2).Find(what:=, lookat:=xlWhole)
    
    '   ã 쿡 ۾ 
    If Not T Is Nothing Then
        For r = 1 To UserForm1.ListView3.ListItems.Count
            ' ǥ Ʈ 1 ListView3 ׸ ã
            Set TR = Sheets("ǥ").Columns(1).Find(what:=UserForm1.ListView3.ListItems(r).text, lookat:=xlWhole)
         
            ' ׸ ã 쿡 ۾ 
            If Not TR Is Nothing Then
                UserForm1.ListView3.ListItems(r).ListSubItems(4).text = Sheets("ǥ").Cells(TR.row, T.Column).Value
            End If
        Next r
    End If
    
    ' "Ȯζ 5" üũڽ õ 
    If Sheets("輺").CheckBoxes("Ȯζ 5").Value = 1 Then
        ' H10:H41   ۾ 
        For Each XT In Sheets("輺").Range("H10:H41").Cells
            If Sheets("輺").Cells(XT.row, "D").Value <> "" Then
                Set TR = Sheets("ǥ").Columns(1).Find(what:=Sheets("輺").Cells(XT.row, "D").Value, lookat:=xlWhole)
                If Not TR Is Nothing Then
                    XT.Value = Sheets("ǥ").Cells(TR.row, T.Column).Value
                End If
            End If
        Next XT
        
        ' P10:P41   ۾ 
        For Each XT In Sheets("輺").Range("P10:P41").Cells
            If Sheets("輺").Cells(XT.row, "L").Value <> "" Then
                Set TR = Sheets("ǥ").Columns(1).Find(what:=Sheets("輺").Cells(XT.row, "L").Value, lookat:=xlWhole)
                If Not TR Is Nothing Then
                    XT.Value = Sheets("ǥ").Cells(TR.row, T.Column).Value
                End If
            End If
        Next XT
        
    Else
        For Each XT In Sheets("輺").Range("H10:H41").Cells
            If Sheets("輺").Cells(XT.row, "D").Value <> "" Then

                    XT.Value = ""
            End If
        Next XT
        
        ' P10:P41   ۾ 
        For Each XT In Sheets("輺").Range("P10:P41").Cells
            If Sheets("輺").Cells(XT.row, "L").Value <> "" Then

                    XT.Value = ""
            End If
        Next XT
    End If
End Sub


