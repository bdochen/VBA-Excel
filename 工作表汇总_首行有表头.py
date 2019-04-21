Sub 工作表汇总_首行有表头()
    Dim DestSheet As Worksheet
    Dim EverySheet As Worksheet
    Dim CopyRange As Range
    Dim LastRow As Long
    Dim UseRow As Long
    Dim RowRange As String
    
    Application.DisplayAlerts = False
    On Error Resume Next
    ActiveWorkbook.Worksheets("超级汇总表VBA").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Set DestSheet = ActiveWorkbook.Worksheets.Add
    DestSheet.Name = "超级汇总表VBA"
    
    For Each EverySheet In ActiveWorkbook.Worksheets
        If EverySheet.Name <> DestSheet.Name Then
            LastRow = DestSheet.UsedRange.Rows.Count
            If LastRow = 1 Then
                LastRow = 0
                Set CopyRange = EverySheet.UsedRange
                CopyRange.Copy
            Else
                LastRow = LastRow
                UseRow = EverySheet.UsedRange.Rows.Count
                
                
                If UseRow = 1 Then
                    Set CopyRange = EverySheet.UsedRange
                    CopyRange.Copy
                Else
                    RowRange = "2:" & UseRow
                    'MsgBox RowRange
                    EverySheet.Select
                    Rows(RowRange).Select
                    Selection.Copy
                End If
            
            End If
            DestSheet.Cells(LastRow + 1, 1).PasteSpecial
         End If
    Next
    DestSheet.Columns.AutoFit
    DestSheet.Select
    ActiveWorkbook.Save
End Sub
