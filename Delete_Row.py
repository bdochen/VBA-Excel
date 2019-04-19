Sub 筛选特定单元格删除其所在行()
    
    '计算F列数据占用行数
    Dim VRow As Long
    VRow = Cells(Rows.Count, 6).End(xlUp).Row
    Dim rng As Range
    Dim rng1 As Range
    Dim rng2 As Range
    Dim b As Long
    
    '定义数组
    Dim List()
    ReDim List(0 To VRow)
    b = 0
    List(1) = 0
    
    '将值等于某个特定字符的单元格行号写入数组List()
    For i = 1 To VRow
        If Cells(i, 6) = "失效" Then
            List(b) = i
            b = b + 1
        End If
    Next

    '穷举List()中的数据并选定其对应的单元格
    For bb = 0 To b
        If List(bb) <> 0 Then
            If bb = 0 Then
                Set rng1 = Cells(List(bb), 6)
                Set rng = Union(rng1, rng1)
            Else
                Set rng1 = Cells(List(bb), 6)
                Set rng = Union(rng, rng1)
            End If
        End If
    Next
    
    rng.Select
    Selection.EntireRow.Delete
    

    ActiveWorkbook.Save
    Range("A1").Select
    '全选所有存在字符的单元格以便数据处理
    'Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select

End Sub
