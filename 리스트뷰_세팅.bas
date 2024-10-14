Attribute VB_Name = "Ʈ_"

Sub L4_Set()
    ' ListView4 ʱȭ  
    With UserForm1.ListView4
        '  ¸ Report  (÷  ǥ)
        .View = lvwReport
        
        ' ü   ϵ 
        .FullRowSelect = True
        
        ' ׸ ǥ
        .Gridlines = True
        
        '  ÷  (缳 )
        .ColumnHeaders.Clear
        
        ' ù° ÷: 
        .ColumnHeaders.Add , , "", 40
        .ColumnHeaders.Add , , "Ƿ", 65
        .ColumnHeaders.Add , , "÷̸", 180
        .ColumnHeaders.Add , , "", 70
    End With
End Sub
