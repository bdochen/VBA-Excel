Sub 复制2至12月的11张工作表()
    '将当前工作表复制11份
    Dim i As Long
    For i = 2 To 12
        ActiveSheet.Copy after:=ActiveSheet
        ActiveSheet.Name = i & "月"
    Next
    ActiveWorkbook.Save
End Sub



Sub 复制2至19日的工作表()

    Dim i As Long
    For i = 2 To 19
        ActiveSheet.Copy after:=ActiveSheet
        ActiveSheet.Name = i
    Next
    ActiveWorkbook.Save
End Sub

