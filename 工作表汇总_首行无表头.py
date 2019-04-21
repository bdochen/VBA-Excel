Sub 工作表汇总()
    Dim DestSheet As Worksheet
    Dim EverySheet As Worksheet
    Dim CopyRange As Range
    Dim LastRow As Long
    
    On Error Resume Next
    ActiveWorkbook.Worksheets("超级汇总表VBA").Delete
    On Error GoTo 0
    Set DestSheet = ActiveWorkbook.Worksheets.Add
    DestSheet.Name = "超级汇总表VBA"
    
    For Each EverySheet In ActiveWorkbook.Worksheets
        If EverySheet.Name <> DestSheet.Name Then
            LastRow = DestSheet.UsedRange.Rows.Count
           'MsgBox LastRow
            Set CopyRange = EverySheet.UsedRange
            CopyRange.Copy
            If LastRow = 1 Then
                LastRow = 0
            Else
                LastRow = LastRow
            End If
            DestSheet.Cells(LastRow + 1, 1).PasteSpecial
         End If
    Next
    DestSheet.Columns.AutoFit
End Sub
