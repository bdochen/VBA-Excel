Sub test1()
  Dim a, b, c As Range, firstAddress As String ,DestSheet As Worksheet
  Set DestSheet = ActiveWorkbook.Worksheets.Add
  DestSheet.Name = rand()
  On Error Resume Next
  With Worksheets("Sheet4").Range("a1:a15")
    Set c = .Find(2, LookIn:=xlValues)
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
    a.EntireRow.Copy Destination:=Worksheets("A").Range(firstAddress)
    'a.EntireRow.Copy Destination:=Worksheets("A").Range("A2")
  End With
End Sub
