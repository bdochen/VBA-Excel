Sub 自动复制工作表模板()

Dim i, j As Integer

i = 1

j = 1

For i = 1 To 30   '循环30次，相当于复制30个工作表

j = j + 1

    Sheets("模板").Copy After:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = j
    
Next
End Sub
