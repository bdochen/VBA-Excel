Sub 新建1至12月的12张工作表()
    Dim i As Long
    Worksheets.Add
    ActiveSheet.Name = "1月"
    For i = 2 To 12
        ActiveSheet.Copy after:=ActiveSheet
        ActiveSheet.Name = i & "月"
    Next
    ActiveWorkbook.Save
End Sub
