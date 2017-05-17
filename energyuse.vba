Option Explicit

Sub energuse()

Dim s1n As Long
Dim n As Long
Dim random As Long
Dim sheet1 As Worksheet

Set sheet1 = Worksheets("Sheet1")

sheet2.Cells(1, 93).Value = "Archetype"
sheet1.Range("A1").EntireColumn.Insert
For n = 2 To 3345
    If sheet1.Cells(n, 22).Value = 3 And _
        sheet1.Cells(n, 24).Value = 6 Then
            sheet2.Cells(n, 94).Value = "1b"
    ElseIf sheet1.Cells(n, 22).Value = 3 And _
        sheet1.Cells(n, 24).Value = 7 Then
            sheet2.Cells(n, 94).Value = "1c"
    ElseIf sheet1.Cells(n, 22).Value = 4 And _
        sheet1.Cells(n, 24).Value = 7 Then
            sheet2.Cells(n, 94).Value = "2a"
    ElseIf sheet1.Cells(n, 22).Value = 6 And _
        sheet1.Cells(n, 24).Value = 5 Then
            sheet2.Cells(n, 94).Value = "4a"
    ElseIf sheet1.Cells(n, 22).Value = 6 And _
        sheet1.Cells(n, 24).Value = 7 Then
            sheet2.Cells(n, 94).Value = "4b"
    ElseIf sheet1.Cells(n, 22).Value = 7 And _
        sheet1.Cells(n, 24).Value = 7 Then
            sheet2.Cells(n, 94).Value = "5a"
    ElseIf sheet1.Cells(n, 22).Value = 3 And _
        sheet1.Cells(n, 30).Value = 3 Then
            sheet2.Cells(n, 94).Value = "1a"
    ElseIf sheet1.Cells(n, 22).Value = 5 And _
        sheet1.Cells(n, 30).Value = 1 Then
            sheet2.Cells(n, 94).Value = "3a"
    ElseIf sheet1.Cells(n, 22).Value = 6 And _
        sheet1.Cells(n, 30).Value = 1 Then
            sheet2.Cells(n, 94).Value = "4c"
    ElseIf sheet1.Cells(n, 22).Value = 7 And _
        sheet1.Cells(n, 30).Value = 1 Then
            sheet2.Cells(n, 94).Value = "5b"
    End If
Next n
sheet1.Range("A1").EntireColumn.Delete

For s1n = 2 To 3345
    If sheet1.Cells(s1n, 93).Value = "1b" Then
      random = ((918 - 2 + 1) * Rnd) + 2
      sheet1.Cells(s1n, 92).Value = _
      (Sheets("Sheet2").Cells(random, 2).Value) * (sheet1.Cells(s1n, 91).Value)
    ElseIf sheet1.Cells(s1n, 93).Value = "1c" Then
      random = ((655 - 2 + 1) * Rnd) + 2
      sheet1.Cells(s1n, 92).Value = _
      (Sheets("Sheet2").Cells(random, 3).Value) * (sheet1.Cells(s1n, 91).Value)
    ElseIf sheet1.Cells(s1n, 93).Value = "2a" Then
      random = ((78 - 2 + 1) * Rnd) + 2
      sheet1.Cells(s1n, 92).Value = _
      (Sheets("Sheet2").Cells(random, 4).Value) * (sheet1.Cells(s1n, 91).Value)
    ElseIf sheet1.Cells(s1n, 93).Value = "4b" Then
      random = ((71 - 2 + 1) * Rnd) + 2
      sheet1.Cells(s1n, 92).Value = _
      (Sheets("Sheet2").Cells(random, 7).Value) * (sheet1.Cells(s1n, 91).Value)
    ElseIf sheet1.Cells(s1n, 93).Value = "5a" Then
      random = ((18 - 2 + 1) * Rnd) + 2
      sheet1.Cells(s1n, 92).Value = _
      (Sheets("Sheet2").Cells(random, 9).Value) * (sheet1.Cells(s1n, 91).Value)
    ElseIf sheet1.Cells(s1n, 93).Value = "1a" Then
      random = ((194 - 2 + 1) * Rnd) + 2
      sheet1.Cells(s1n, 92).Value = _
      (Sheets("Sheet2").Cells(random, 1).Value) * (sheet1.Cells(s1n, 91).Value)
    ElseIf sheet1.Cells(s1n, 93).Value = "3a" Then
      random = ((14 - 2 + 1) * Rnd) + 2
      sheet1.Cells(s1n, 92).Value = _
      (Sheets("Sheet2").Cells(random, 5).Value) * (sheet1.Cells(s1n, 91).Value)
    ElseIf sheet1.Cells(s1n, 93).Value = "4c" Then
      random = ((5 - 2 + 1) * Rnd) + 2
      sheet1.Cells(s1n, 92).Value = _
      (Sheets("Sheet2").Cells(random, 8).Value) * (sheet1.Cells(s1n, 91).Value) 'Not enough data values for 4a and 4b
    End If
Next s1n

'For s1n = 2 To 3345
'    If sheet1.Cells(s1n, 21).Value = 3 And _
'    sheet1.Cells(s1n, 23).Value = 6 Then
'      random = ((918 - 2 + 1) * Rnd) + 2
'      sheet1.Cells(s1n, 92).Value = _
'      (Sheets("Sheet2").Cells(random, 2).Value) * (sheet1.Cells(s1n, 91).Value)
'    ElseIf sheet1.Cells(s1n, 21).Value = 3 And _
'      sheet1.Cells(s1n, 23).Value = 7 Then
'      random = ((655 - 2 + 1) * Rnd) + 2
'      sheet1.Cells(s1n, 92).Value = _
'      (Sheets("Sheet2").Cells(random, 3).Value) * (sheet1.Cells(s1n, 91).Value)
'    ElseIf sheet1.Cells(s1n, 21).Value = 4 And _
'      sheet1.Cells(s1n, 23).Value = 7 Then
'      random = ((78 - 2 + 1) * Rnd) + 2
'      sheet1.Cells(s1n, 92).Value = _
'      (Sheets("Sheet2").Cells(random, 4).Value) * (sheet1.Cells(s1n, 91).Value)
'    ElseIf sheet1.Cells(s1n, 21).Value = 6 And _
'      sheet1.Cells(s1n, 23).Value = 7 Then
'      random = ((71 - 2 + 1) * Rnd) + 2
'      sheet1.Cells(s1n, 92).Value = _
'      (Sheets("Sheet2").Cells(random, 7).Value) * (sheet1.Cells(s1n, 91).Value)
'    ElseIf sheet1.Cells(s1n, 21).Value = 7 And _
'      sheet1.Cells(s1n, 23).Value = 7 Then
'      random = ((18 - 2 + 1) * Rnd) + 2
'      sheet1.Cells(s1n, 92).Value = _
'      (Sheets("Sheet2").Cells(random, 9).Value) * (sheet1.Cells(s1n, 91).Value)
'    ElseIf sheet1.Cells(s1n, 21).Value = 3 And _
'      sheet1.Cells(s1n, 23).Value = 3 Then
'      random = ((194 - 2 + 1) * Rnd) + 2
'      sheet1.Cells(s1n, 92).Value = _
'      (Sheets("Sheet2").Cells(random, 1).Value) * (sheet1.Cells(s1n, 91).Value)
'    ElseIf sheet1.Cells(s1n, 21).Value = 5 And _
'      sheet1.Cells(s1n, 23).Value = 1 Then
'      random = ((14 - 2 + 1) * Rnd) + 2
'      sheet1.Cells(s1n, 92).Value = _
'      (Sheets("Sheet2").Cells(random, 5).Value) * (sheet1.Cells(s1n, 91).Value)
'    ElseIf sheet1.Cells(s1n, 21).Value = 6 And _
'      sheet1.Cells(s1n, 23).Value = 1 Then
'      random = ((5 - 2 + 1) * Rnd) + 2
'      sheet1.Cells(s1n, 92).Value = _
'      (Sheets("Sheet2").Cells(random, 8).Value) * (sheet1.Cells(s1n, 91).Value)
'    End If
'Next s1n

For s1n = 2 To 3345
    If sheet1.Cells(s1n, 92).Value = "" Then
        sheet1.Cells(s1n, 92).Value = 0
    End If
Next s1n

End Sub
