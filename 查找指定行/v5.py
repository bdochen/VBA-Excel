Sub 复制行中有指定数值的行()
  'LW update @ 20190801
  Dim a, b, c As Range, firstAddress, Sheetname, Key As String, DestSheet As Worksheet
  
  '定义需要取值的表名
  Sheetname = "Sheet1"
  '定义需要取值的字符
  Key = 2
  
  Sheets(Sheetname).Select
  Set DestSheet = ActiveWorkbook.Worksheets.Add
  '产生随机表名
  DestSheet.Name = "LW" & Int(Replace(Rnd(), ".", "") / 1)
  Sheets(Sheetname).Select
  On Error Resume Next
  With Worksheets(Sheetname).Range("a1:H1")
    Set c = .Find(Key, LookIn:=xlValues, lookat:=xlWhole)
    '下面为部分匹配
    'Set c = .Find(Key, LookIn:=xlValues, lookat:=xlPart)
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
    a.EntireColumn.Copy Destination:=DestSheet.Range(firstAddress)
    'a.EntireRow.Copy Destination:=Worksheets("A").Range(firstAddress)
    'a.EntireRow.Copy Destination:=Worksheets("A").Range("A2")
  End With
  DestSheet.Activate
  'ActiveWorkbook.Save
End Sub
