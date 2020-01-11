Sub newfolder()

    Dim i&, p$

    On Error Resume Next

'将文件名事先填入第一列
    For i = 1 To Cells(Rows.Count, 1).End(xlUp).Row

        MkDir ThisWorkbook.Path & "\" & Cells(i, 1)

    Next

End Sub
