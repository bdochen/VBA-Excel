Sub 复制行中有指定数值的行()
  'LW update @ 20190801
  Dim a, b, c As Range, firstAddress, sheetname As String, DestSheet As Worksheet
  sheetname = "Sheet1"
  Sheets(sheetname).Select
  Set DestSheet = ActiveWorkbook.Worksheets.Add
  '产生随机表名
  DestSheet.Name = "LW" & Int(Replace(Rnd(), ".", "") / 100)
  
  On Error Resume Next
  With Worksheets(sheetname).Range("a1:p1")
    Set c = .Find(2, LookIn:=xlValues, lookat:=xlWhole)
    '下面为部分匹配
    'Set c = .Find(2, LookIn:=xlValues, lookat:=xlPart)
    If Not c Is Nothing Then
      Set a = c
      firstAddress = c.Address
      Do
        Set c = .FindNext(c)
        If Not c Is Nothing Then
            Set a = Union(a, c)
        End If
      Loop While Not c Is Nothing And c.Address <> firstAddress
    End If
    a.EntireColumn.Select
    a.EntireRow.Copy Destination:=DestSheet.Range(firstAddress)
    'a.EntireRow.Copy Destination:=Worksheets("A").Range(firstAddress)
    'a.EntireRow.Copy Destination:=Worksheets("A").Range("A2")
  End With
  DestSheet.Activate
  'ActiveWorkbook.Save
End Sub
