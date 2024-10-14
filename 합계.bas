Attribute VB_Name = "հ"
Sub հݾ()
X = 0


If Not UserForm1.ListBox1.ListCount = 0 And UserForm1.TextBox3 <> "" Then
  H = UserForm1.TextBox3.Value
  
  For r = 0 To UserForm1.ListBox1.ListCount - 1
     If UserForm1.ListBox1.List(r, 3) <> "" And UserForm1.ListBox1.List(r, 3) <> "-" And UserForm1.ListBox1.List(r, 3) <> 0 Then
         UserForm1.ListBox1.List(r, 4) = Format(UserForm1.ListBox1.List(r, 2) * UserForm1.ListBox1.List(r, 3), "#,###")
         
         X = UserForm1.ListBox1.List(r, 2) * UserForm1.ListBox1.List(r, 3) + X
        Else
         UserForm1.ListBox1.List(r, 4) = UserForm1.ListBox1.List(r, 2) * 0
         X = UserForm1.ListBox1.List(r, 2) * 0 + X
     End If
  Next r

  UserForm1.Label5.Caption = UserForm1.ListBox1.List(0, 1) & " " & UserForm1.ListBox1.ListCount & " " & Format(X, "#,###")
Else
  UserForm1.Label5.Caption = "Ǽ/Ѿ"
End If


End Sub
