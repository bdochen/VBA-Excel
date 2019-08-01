Sub 复制行中有指定数值的行()
  'LW update @ 20190801
  Dim a, b, c As Range, firstAddress As String, DestSheet As Worksheet
  Set DestSheet = ActiveWorkbook.Worksheets.Add
  '产生随机表名
  DestSheet.Name = "LW" & Int(Replace(Rnd(), ".", "") / 10000)
  On Error Resume Next
  With Worksheets("Sheet4").Range("a1:a15")
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
    a.EntireRow.Select
    a.EntireRow.Copy Destination:=DestSheet.Range(firstAddress)
    'a.EntireRow.Copy Destination:=Worksheets("A").Range(firstAddress)
    'a.EntireRow.Copy Destination:=Worksheets("A").Range("A2")
  End With
  DestSheet.Activate
  ActiveWorkbook.Save
End Sub
