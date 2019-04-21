Sub 合并工作薄到一张工作表()
    
    Dim SummarySheet As Worksheet
    Dim SelectedFiles() As Variant
    Dim Nrow As Long
    Dim FileName As String
    Dim NFile As Long
    Dim WorkBk As Workbook
    Dim SourceRange As Range
    Dim DestRange As Range
    
    
    Set SummarySheet = ThisWorkbook.Worksheets(1)
    
    SelectedFiles = Application.GetOpenFilename(filefilter:="Excel 文件(*.xl*),*.xl*", MultiSelect:=True)
    
    Nrow = 1
    
    For NFile = LBound(SelectedFiles) To UBound(SelectedFiles)
        
        FileName = SelectedFiles(NFile)
        
        Set WorkBk = Workbooks.Open(FileName)
    
        Set SourceRange = WorkBk.Worksheets(1).UsedRange
        Set DestRange = SummarySheet.Range("A" & Nrow)
        Set DestRange = DestRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
        
        DestRange.Value = SourceRange.Value
        Nrow = Nrow + DestRange.Rows.Count
        
        WorkBk.Close savechanges:=False
    
    Next
        
        SummarySheet.Columns.AutoFit
        
        
      
        
End Sub
