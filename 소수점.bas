Attribute VB_Name = "Ҽ"
Sub Ҽڸ()

For r = 1 To UserForm1.ListView3.ListItems.Count

X = Sheets("DB").Columns("s").Find(what:=UserForm1.ListView3.ListItems(r).text, lookat:=xlWhole).row
UserForm1.ListView3.ListItems(r).ListSubItems(1) = Round(UserForm1.ListView3.ListItems(r).ListSubItems(1), Sheets("DB").Cells(X, "T"))


Next r

End Sub
