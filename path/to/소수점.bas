Attribute VB_Name = "소수점"
Sub 소수점자리맞춤()

For r = 1 To UserForm1.ListView3.ListItems.Count

X = Sheets("측정DB").Columns("s").Find(what:=UserForm1.ListView3.ListItems(r).text, lookat:=xlWhole).row
UserForm1.ListView3.ListItems(r).ListSubItems(1) = Round(UserForm1.ListView3.ListItems(r).ListSubItems(1), Sheets("측정DB").Cells(X, "T"))


Next r

End Sub
