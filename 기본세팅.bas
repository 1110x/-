Attribute VB_Name = "기본세팅"
Sub 단추8_Click()
Attribute 단추8_Click.VB_ProcData.VB_Invoke_Func = "Z\n14"
UserForm1.Show 0
Sheets("TESTS").Cells(1, "B") = UserForm1.Left



End Sub

Sub 세팅()
    Dim s As Long
    Dim E As Long
    
        Sheets("담당자정보").Rows(1 & ":" & 1000).EntireRow.Hidden = flase
    s = Sheets("담당자정보").Range("A3:A100").End(xlDown).row + 5
    E = 99
    
    ' Hide the rows from s to E
    Sheets("담당자정보").Rows(s & ":" & E).EntireRow.Hidden = True
End Sub
