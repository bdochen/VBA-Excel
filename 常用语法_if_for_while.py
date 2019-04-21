Sub Doloop循环()
    
    Dim Score As Integer
    Dim Grade As String
    
    
    Range("A3").Select
    
    Do While ActiveCell.Value <> ""
        Score = ActiveCell.Offset(0, 3).Value
            
            If Score >= 90 Then
                Grade = "优秀"
            ElseIf Score >= 80 Then
                Grade = "普通"
            Else
                Grade = "合格"
            End If
        
        ActiveCell.Offset(0, 4).Value = Grade
            
        
        ActiveCell.Offset(1, 0).Select
    Loop

  
End Sub


'******************************************************************


Sub foreach语句()

    Dim SingleCell As Range
    
    Dim ListofCells As Range
    
    ThisWorkbook.Activate
    Worksheets("员工信息").Activate
    
    
    Set ListofCells = Range("A3", Range("A2").End(xlDown))
    
    Workbooks.Add
    
    Range("A1").Value = "优秀员工列表"
    Range("A2").Select
    
    For Each SingleCell In ListofCells

        If SingleCell.Offset(0, 3).Value >= 90 Then
            
            ActiveCell.Value = SingleCell.Offset(0, 1).Value
            ActiveCell.Offset(1, 0).Select
        End If
    
    Next

End Sub


'******************************************************************



