Attribute VB_Name = ""
Sub ()
On Error Resume Next

X = Sheets("ܰ").Rows(1).Find(what:=UserForm1.ComboBox1.text, lookat:=xlWhole).Column

For r = 0 To UserForm1.ListBox1.ListCount - 1
TG = Sheets("ܰ").Columns(4).Find(what:=UserForm1.ListBox1.List(r, 1), lookat:=xlWhole).row

UserForm1.ListBox1.List(r, 3) = Format(Sheets("ܰ").Cells(TG, X), "#,###")

Next r
հݾ
End Sub
