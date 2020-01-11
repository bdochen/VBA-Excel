Function PPH(str As String) As String
    Set regEx = CreateObject("vbscript.regexp")
    With regEx
        .Global = 1
        .Pattern = "[Function RemoveNarrow(str As String) As String"
    Set regEx = CreateObject("vbscript.regexp")
    With regEx
        .Global = 1
        
        .Pattern = "[^\u4e00-\u9fff]"
       PPH = .Replace(str, "")
    End With
    End With
End Function


Sub 逐个工作表添加行和值()
Dim Sheet_Name, Content As String
Dim Col_num As Long

For Each sht In Worksheets
    sht.Activate
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "被审计单位：company"
    
    Range("A3").Select
    Sheet_Name = ActiveSheet.Name
    Content = "项目：" & Sheet_Name
    ActiveCell.FormulaR1C1 = Content
    
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "财务报表截止日/期间：2019年12月31日"
    
    
    Range("A5").Select
    Col_num = Selection.Columns.Count
    'MsgBox (Row_num)
    If Col_num > 4 Then
        Col_num = Col_num
    Else
        Col_num = 5
    End If
    Cells(2, Col_num - 2).Select
    ActiveCell.FormulaR1C1 = "索引号："
    
    Cells(2, Col_num).Select
    ActiveCell.FormulaR1C1 = "页次："
    
    Cells(3, Col_num - 2).Select
    ActiveCell.FormulaR1C1 = "编制人： user"
    
    
    Cells(3, Col_num).Select
    ActiveCell.FormulaR1C1 = "日期：2020-1-6"
    
    
    Cells(4, Col_num - 2).Select
    ActiveCell.FormulaR1C1 = "复核人： user"
    
    
    Cells(4, Col_num).Select
    ActiveCell.FormulaR1C1 = "日期：2020-1-21"



Next

End Sub
