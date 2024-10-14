Attribute VB_Name = "견적종류"
Sub 견적종류변경()
On Error Resume Next

X = Sheets("견적단가").Rows(1).Find(what:=UserForm1.ComboBox1.text, lookat:=xlWhole).Column

For r = 0 To UserForm1.ListBox1.ListCount - 1
TG = Sheets("견적단가").Columns(4).Find(what:=UserForm1.ListBox1.List(r, 1), lookat:=xlWhole).row

UserForm1.ListBox1.List(r, 3) = Format(Sheets("견적단가").Cells(TG, X), "#,###")

Next r
합계금액
End Sub
