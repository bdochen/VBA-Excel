Sub 数据透视VBA()

'数据透视VBA中工作表名称不能有空格，因此进行了重命名

    ActiveSheet.Name = "数据透视"
    Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Crland_row = Selection.Rows.Count
    
    Crland_SourceData = ActiveSheet.Name & "!R1C1:R" & Crland_row & "C2"
'   MsgBox Crland_SourceData
    
    Crland_TableDestination = ActiveSheet.Name & "!R1C7"
'   MsgBox Crland_TableDestination

    Range("A1:B6").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Crland_SourceData, Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=Crland_TableDestination, TableName:="数据透视表crland", DefaultVersion:= _
        xlPivotTableVersion15
    Sheets("数据透视").Select
    Cells(1, 7).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("数据透视表crland").PivotFields("型号")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("数据透视表crland").AddDataField ActiveSheet.PivotTables("数据透视表crland" _
        ).PivotFields("价格"), "求和项:价格", xlSum
End Sub


'********************************************************************

Sub 数据透视VBA()

'数据透视VBA中工作表名称不能有空格，因此进行了重命名
' 自定义首行数据
'定义首行数据
    LX = Cells(1, 1).Value
    HZ = Cells(1, 2).Value
    
    ActiveSheet.Name = "数据透视"
    Range("A1:B1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Crland_row = Selection.Rows.Count
    
    Crland_SourceData = ActiveSheet.Name & "!R1C1:R" & Crland_row & "C2"
'   MsgBox Crland_SourceData
    
    Crland_TableDestination = ActiveSheet.Name & "!R1C7"
'   MsgBox Crland_TableDestination

    Range("A1:B6").Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Crland_SourceData, Version:=xlPivotTableVersion15).CreatePivotTable _
        TableDestination:=Crland_TableDestination, TableName:="数据透视表crland", DefaultVersion:= _
        xlPivotTableVersion15
        
        
'   Sheets("数据透视").Select
    Cells(1, 7).Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    With ActiveSheet.PivotTables("数据透视表crland").PivotFields(LX)
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("数据透视表crland").AddDataField ActiveSheet.PivotTables("数据透视表crland" _
        ).PivotFields(HZ), "求和项:" & HZ, xlSum
End Sub
