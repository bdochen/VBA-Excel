Sub 复制2至12月的12张工作表()
    '将当前工作表复制11份
    Dim i As Long
    For i = 2 To 12
        ActiveSheet.Copy after:=ActiveSheet
        ActiveSheet.Name = i & "月"
    Next
    ActiveWorkbook.Save
End Sub
