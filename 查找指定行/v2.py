Sub test1()
  Dim a, b, c As Range, firstAddress As String
  On Error Resume Next
  With Worksheets(1).Range("a1:a15")
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
    'Sheets("Sheet2").Select
    a.EntireRow.Copy
    Range("A18").Select
    ActiveSheet.Paste
    
  End With
End Sub
