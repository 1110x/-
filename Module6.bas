Attribute VB_Name = "Module6"
癤풞ttribute VB_Name = "TESTModule"
Sub TEST01()

Dim r As Integer

For r = 1 To 10
    Columns(r + 10).ColumnWidth = Columns(r).ColumnWidth
Next r


End Sub

Sub ()


For r = 2 To 1819


If Cells(r, "E") <> False Then
Cells(r, "C") = Cells(r, "E")
End If


If Cells(r, "D") <> False Then
Cells(r, "B") = Cells(r, "D")
End If


Next r


End Sub

Sub 2()


For r = 3 To 368
  Sheets("").Cells(r, "B") = Left(Sheets("").Cells(r, "I"), 4) & "-" & Mid(Sheets("").Cells(r, "I"), 5, 2) & "-" & Mid(Sheets("").Cells(r, "I"), 7, 2)
Next r


End Sub


Sub 釜獺_寗()

For Each X In Range("Z2:Z1814") '
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("Z1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AA2:AA1814") '
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AA1").Column) = Round(X, 2)
   End If
 End If
End If
Next X

For Each X In Range("AB2:AB1814") '
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AB1").Column) = Round(X, 2)
   End If
 End If
End If
Next X

For Each X In Range("AC2:AC1814") '
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AC1").Column) = Round(X, 4)
   End If
 End If
End If
Next X

For Each X In Range("AD2:AD1814") '6크
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AD1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AD2:AD1814") '6크
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AD1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AE2:AE1814") '카
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AE1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AF2:AF1814") '
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AF1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AG2:AG1814") '티
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AG1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AH2:AH1814") '크
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AH1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AI2:AI1814") '철
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AI1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AJ2:AJ1814") '틸
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AJ1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AK2:AK1814") '
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AK1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AL2:AL1814") '慕
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AL1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AM2:AM1814") '
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AM1").Column) = Round(X, 3)
   End If
 End If
End If
Next X

For Each X In Range("AO2:AO1814") '寗
If X <> "" Then
 If X <> "柰" Then
   If X <> "劇柰" Then
   Cells(X.row, Range("AO1").Column) = Round(X, 2)
   End If
 End If
End If
Next X


End Sub
