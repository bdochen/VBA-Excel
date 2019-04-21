Sub 合并工作薄()
    
    Dim SummarySheet As Worksheet
    Dim SelectedFiles() As Variant
    Dim Nrow As Long
    Dim FileName As String
    Dim NFile As Long
    Dim WorkBk As Workbook
    Dim SourceRange As Range
    Dim DestRange As Range
    Dim SheetName_temp As String
    Dim i As Long
    Dim WN As String
    

    
    SelectedFiles = Application.GetOpenFilename(filefilter:="Excel 文件(*.xl*),*.xl*", MultiSelect:=True)
    
    Nrow = 1
    i = 1
    
    For NFile = LBound(SelectedFiles) To UBound(SelectedFiles)
    
        
        FileName = SelectedFiles(NFile)
        
        
        Set WorkBk = Workbooks.Open(FileName)

        Set SourceRange = WorkBk.Worksheets(1).UsedRange
        
        SheetName_temp = WorkBk.Worksheets(1).Name
        
        WN = Workbooks(1).Name
        Windows(WN).Activate
        Worksheets.Add
        SheetName_temp = SheetName_temp & "-" & i
        ActiveSheet.Name = SheetName_temp
        i = i + 1
        Set SummarySheet = ThisWorkbook.Worksheets(SheetName_temp)
        
        Set DestRange = SummarySheet.Range("A" & Nrow)
        Set DestRange = DestRange.Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
        
        DestRange.Value = SourceRange.Value

        
        WorkBk.Close savechanges:=False
    
    Next
        
        SummarySheet.Columns.AutoFit
    
End Sub


