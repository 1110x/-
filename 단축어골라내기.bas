Attribute VB_Name = "󳻱"
Sub ã()
    Dim text As String
    Dim result As String
    Dim startPos As Long
    Dim endPos As Long
    Dim ExtractedText As String

    ' TextBox8 ؽƮ 
    text = UserForm1.TextBox8.text
    result = ""

    ' ȣ  ؽƮ 
    Do
        ' ȣ ۰  ġ ã
        startPos = InStr(text, "")
        endPos = InStr(startPos + 1, text, "")

        ' ȣ  ϴ 
        If startPos > 0 And endPos > startPos Then
            ' ȣ  ؽƮ 
            ExtractedText = Mid(text, startPos + 1, endPos - startPos - 1)
            result = result & ExtractedText & " "

            '  ȣ ã  ؽƮ 
            text = Mid(text, endPos + 1)
        Else
            Exit Do
        End If
    Loop

    '  
     = Trim(result)

    
UserForm1.ComboBox5.ListIndex = Sheets("").Columns("H").Find(what:=, lookat:=xlWhole).row - 2

    
End Sub
