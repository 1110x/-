Attribute VB_Name = "⺻"
Sub 8_Click()
Attribute 8_Click.VB_ProcData.VB_Invoke_Func = "Z\n14"
UserForm1.Show 0
Sheets("TESTS").Cells(1, "B") = UserForm1.Left



End Sub

Sub ()
    Dim s As Long
    Dim E As Long
    
        Sheets("").Rows(1 & ":" & 1000).EntireRow.Hidden = flase
    s = Sheets("").Range("A3:A100").End(xlDown).row + 5
    E = 99
    
    ' Hide the rows from s to E
    Sheets("").Rows(s & ":" & E).EntireRow.Hidden = True
End Sub
