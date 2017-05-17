Option Explicit

Sub preprocessing2()

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets.Add(After:= _
ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws.Name = "postCodeSortEPC"

Set ws = ThisWorkbook.Sheets.Add(After:= _
ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws.Name = "propertyNameSort"

Set ws = ThisWorkbook.Sheets.Add(After:= _
ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws.Name = "SUSDEMtext"

Set ws = ThisWorkbook.Sheets.Add(After:= _
ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws.Name = "Sheet1"

Set ws = ThisWorkbook.Sheets.Add(After:= _
ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws.Name = "SUSDEMinput"

sortPostCode
extractEPCaddress
extractHPaddress
sortAddress
DateSorter

Worksheets("propertyNameSort").Copy _
After:=Worksheets("propertyNameSort")
Sheets("propertyNameSort (2)").Name = "Haringey_All"
Sheets("Haringey_All").Range("A1").EntireColumn.Delete

Dim LastRowHA As Long
Dim LastRowPNS As Long
Dim HAsheet As Worksheet
Dim STsheet As Worksheet
Dim PNSsheet As Worksheet

Set HAsheet = Worksheets("Haringey_All")
Set STsheet = Worksheets("SUSDEMtext")
Set PNSsheet = Worksheets("propertyNameSort")

LastRowHA = lRow(HAsheet)
LastRowPNS = lRow(PNSsheet)

'Subs take values from HAsheet which contains UKMaps, UKBuildings and EPC data for all 1304 houses in Haringey that contain EPC data
'For each UKMap and UKBuildings entry, there may be several EPC entries, this could be due to there being several different flats within one house
'Multiple EPC certificates taken, but previous ones for the same resisdence ignored to take most recent in previous code
'If there are multiple EPC certificates from the same date for the same building, where possible, average values are taken

archetype HAsheet, STsheet, LastRowPNS 'Defines archetype depending on RBCA and RBCT UKBuildings classification
DwellingType HAsheet, STsheet, LastRowPNS 'Defines dwelling type using EPC data
DwellingPosition HAsheet, STsheet, LastRowPNS 'Defines dwelling position, taking into account dwellying type previously defined and EPC data
averageEPC "NoOfRooms", 5, "E", 135, "EE" 'Finds average number of rooms from EPC data
averageEPC "DoubleGlazingInsulation", 4, "D", 131, "EA" 'Finds average double glazing insulation percentage from EPC data
averageEPCenergy "EnergyConsumption", 27, "AA", 112, "DH"
NoOfStoreys 'Finds number of storeys by applying rules for different archetypes and using the height given by UKBuildings
BuildingHeight HAsheet, STsheet, LastRowPNS ' Takes building height from UKBuildings data
fid HAsheet, STsheet, LastRowPNS 'Extracts fid
Address HAsheet, STsheet, LastRowPNS 'Takes address from previously combined data, concatenates post code and city
areas HAsheet, STsheet, LastRowPNS 'Defines areas of different floors of residences using UKBuildings data
floorHeights HAsheet, STsheet, LastRowPNS 'Defines floor heights by using previously defined building height and number of floors, different formulas
'for different arhcetypes taking into account different construction methods
perimeters HAsheet, STsheet, LastRowPNS 'Takes perimeter from UKBuilding data
ageBand HAsheet, STsheet, LastRowPNS 'Age catagory for SUSDEM different to that defined in UKBuildings, overalap between age bands, weighted random number assigns age band
externalwallconstruction HAsheet, STsheet, LastRowPNS 'Takes qualitative EPC certificate despriptor of wall type and applies rules to fit into appropriate SUSDEM classification
floorConstruction HAsheet, STsheet, LastRowPNS 'Takes qualitative EPC certificate descprptor of floor consturction and applies rules to fit into appropriate SUSDEM classification
wwr
doubleglazing
averageEPCroofing "RoofInsulation", 28, "AB", 1, "A"

'Stopping Application Alerts
Application.DisplayAlerts = False
'Deleting sheets no longer needed
Sheets("postCodeSortEPC").Delete
Sheets("propertyNameSort").Delete
Sheets("Haringey_All").Delete
'Enabling Application alerts
Application.DisplayAlerts = True
Sheets("Haringey_Processed").Range("A1").EntireColumn.Delete

SUSDEMin

End Sub

Sub sortPostCode()

Sheets("Haringey_Processed").Columns(62).Copy Destination:=Sheets("Sheet1").Columns(1)

' hiker95, 07/26/2012
' http://www.mrexcel.com/forum/showthread.php?649576-Extract-unique-values-from-one-column-using-VBA
Dim d As Object, c As Variant, i As Long, lr As Long
Set d = CreateObject("Scripting.Dictionary")
lr = Worksheets("Sheet1").Cells(Rows.count, 1).End(xlUp).Row
c = Worksheets("Sheet1").Range("A2:A" & lr)
For i = 1 To UBound(c, 1)
  d(c(i, 1)) = 1
Next i
Worksheets("Sheet1").Range("B2").Resize(d.count) = Application.Transpose(d.keys)

Sheets("Sheet1").Columns(1).Delete

Dim EPCn As Long
Dim HPn As Long
Dim S1n As Long
Dim HPsheet As Worksheet
Dim EPCsheet As Worksheet
Dim PCSsheet As Worksheet
Dim S1sheet As Worksheet
Dim LastRowHP As Long
Dim LastRowEPC As Long
Dim LastRowS1 As Long
Dim n As Long

Set HPsheet = Worksheets("Haringey_Processed")
Set EPCsheet = Worksheets("DomesticBulkDataAB")
Set PCSsheet = Worksheets("postCodeSortEPC")
Set S1sheet = Worksheets("Sheet1")
'Haringey_Processed comes from the CiMo M0.py script, it contains UKMaps and UKbuildings merged
'DomesticBulkDataAB contains EPC data for the whole of london
'PostCodeSortEPC is the location for EPC data with a Tottenham Borough post-code
LastRowHP = lRow(HPsheet)
LastRowEPC = lRow(EPCsheet)
LastRowS1 = lRow(S1sheet)
PCSsheet.Rows(1).Value = EPCsheet.Rows(1).Value
n = 2
For EPCn = 2 To LastRowEPC
        For S1n = 2 To LastRowS1
            If EPCsheet.Cells(EPCn, 5).Value = S1sheet.Cells(S1n, 1).Value Then 'If the EPC post-code matches the Haringey proccessed post-code then
            ' put this value in the post code sort sheet
                PCSsheet.Rows(n).Value = EPCsheet.Rows(EPCn).Value
                n = n + 1
                Exit For
            End If
        Next S1n
Next EPCn

Application.DisplayAlerts = False
Sheets("Sheet1").Delete
Application.DisplayAlerts = True

End Sub

Sub extractEPCaddress()

Dim PCSn As Long
Dim PCSsheet As Worksheet
Dim LastRowPCS As Long
Dim a As Long
Dim b As Long
Dim c As Long
Dim d As Long
Dim e As Long
Dim f As Long
Dim m As Long
Dim str As String
Dim txt As String
Dim num As String
Dim Utxt As String

Set PCSsheet = Worksheets("postCodeSortEPC")

LastRowPCS = lRow(PCSsheet)

PCSsheet.Range("A1").EntireColumn.Insert
PCSsheet.Range("A1").EntireColumn.Insert
PCSsheet.Range("A1").EntireColumn.Insert
PCSsheet.Range("A1").EntireColumn.Insert
PCSsheet.Range("A1").EntireColumn.Insert
PCSsheet.Range("A1").EntireColumn.Insert
PCSsheet.Range("A1").EntireColumn.Insert
PCSsheet.Range("A1").EntireColumn.Insert
PCSsheet.Range("A1").EntireColumn.Insert
'Adds extra columns to allow for address formatting
'Different variables set for individual loops to make code more consise
b = 1
c = 11
d = 3
e = 2
f = 4

For a = 1 To 2
For PCSn = 2 To LastRowPCS
    PCSsheet.Cells(PCSn, b).Value = PCSsheet.Cells(PCSn, c).Value 'Takes first line of Adress from EPC data
    str = PCSsheet.Cells(PCSn, b).Value 'Assigns this to string
    If Left(str, 4) = "Flat" Then 'If the string begins with "Flat"
        PCSsheet.Cells(PCSn, b).Value = Right(str, Len(str) - 4) 'Remove the word Flat
    End If
    str = PCSsheet.Cells(PCSn, b).Value 'Get text from string by removing all numbers and symbols
    txt = Replace(str, "-", "")
    txt = Replace(txt, ",", "")
    txt = Replace(txt, "0", "")
    txt = Replace(txt, "1", "")
    txt = Replace(txt, "2", "")
    txt = Replace(txt, "3", "")
    txt = Replace(txt, "4", "")
    txt = Replace(txt, "5", "")
    txt = Replace(txt, "6", "")
    txt = Replace(txt, "7", "")
    txt = Replace(txt, "8", "")
    txt = Replace(txt, "9", "")
    If Left(txt, 2) = "a " Or Left(txt, 2) = "b " _
    Or Left(txt, 2) = "c " Or Left(txt, 2) = "d " Then
        txt = Right(txt, Len(txt) - 2) 'Gets rid of letters after letter number adresses, ie 36a
    End If
    txt = Trim(txt)
    If Len(txt) = 1 Then
    txt = "" 'Removes one letter text
    End If
    PCSsheet.Cells(PCSn, d).Value = txt
    num = Left(str, 4)
    For m = 1 To 4
        If Right(num, 1) = "1" Or Right(num, 1) = "2" Or Right(num, 1) = "3" Or _
        Right(num, 1) = "4" Or Right(num, 1) = "5" Or Right(num, 1) = "6" _
        Or Right(num, 1) = "0" Or Right(num, 1) = "7" Or Right(num, 1) = "8" Or _
        Right(num, 1) = "9" Then
            num = num
        Else
            If Len(num) > 0 Then
                num = Left(num, Len(num) - 1)
            End If 'Gets rid of blan spaces at the end of number string
        End If
    Next m
    PCSsheet.Cells(PCSn, e).Value = num
    Utxt = UCase(txt) 'Sets to capitals to make the same as Haringey_Processed
    PCSsheet.Cells(PCSn, f).Value = num & " " & Utxt 'Concatenates number and address
Next PCSn
b = 5
c = 12
d = 7
e = 6
f = 8
Next a 'Repeats for second line of address

For PCSn = 2 To LastRowPCS
    If IsEmpty(PCSsheet.Cells(PCSn, 2).Value) = True Or _
    IsEmpty(PCSsheet.Cells(PCSn, 3).Value) = True Then 'If first line of address empty
        PCSsheet.Cells(PCSn, 9).Value = Trim(PCSsheet.Cells(PCSn, 8).Value)
    Else 'Take the second line of the address
        PCSsheet.Cells(PCSn, 9).Value = Trim(PCSsheet.Cells(PCSn, 4).Value)
    End If
Next PCSn

PCSsheet.Range("A1").EntireColumn.Delete
PCSsheet.Range("A1").EntireColumn.Delete
PCSsheet.Range("A1").EntireColumn.Delete
PCSsheet.Range("A1").EntireColumn.Delete
PCSsheet.Range("A1").EntireColumn.Delete
PCSsheet.Range("A1").EntireColumn.Delete
PCSsheet.Range("A1").EntireColumn.Delete
PCSsheet.Range("A1").EntireColumn.Delete

End Sub

Sub extractHPaddress()

Dim HPn As Long
Dim HPsheet As Worksheet
Dim LastRowHP As Long
Dim txt As String
Dim num As String

Set HPsheet = Worksheets("Haringey_Processed")
LastRowHP = lRow(HPsheet)
HPsheet.Range("A1").EntireColumn.Insert

For HPn = 2 To LastRowHP
    txt = HPsheet.Cells(HPn, 57).Value 'Sets address txt
    num = HPsheet.Cells(HPn, 51).Value 'Sets address number
    HPsheet.Cells(HPn, 1).Value = Trim(num & " " & txt)
    If HPsheet.Cells(HPn, 1).Value = 0 Then
        HPsheet.Cells(HPn, 1).Value = ""
    End If
Next HPn

End Sub

Sub sortAddress()

Dim HPn As Long
Dim PCSn As Long
Dim PNSn As Long
Dim HPsheet As Worksheet
Dim PNSsheet As Worksheet
Dim PCSsheet As Worksheet
Dim LastRowHP As Long
Dim LastRowPCS As Long
Dim n As Long
Dim c As Integer

Set HPsheet = Worksheets("Haringey_Processed")
Set PNSsheet = Worksheets("propertyNamesort")
Set PCSsheet = Worksheets("postCodeSortEPC")

LastRowHP = lRow(HPsheet)
LastRowPCS = lRow(PCSsheet)
PNSsheet.Rows(1).Value = HPsheet.Rows(1).Value
PNSsheet.Range("CP" & 1 & ":FS" & 1).Value = _
PCSsheet.Range("A" & 1 & ":CD" & 1).Value
n = 2
For HPn = 2 To LastRowHP
    c = 1
    For PCSn = 2 To LastRowPCS
        If PCSsheet.Cells(PCSn, 1).Value = HPsheet.Cells(HPn, 1).Value _
        And c = 1 Then
            PNSsheet.Rows(n).Value = HPsheet.Rows(HPn).Value
            PNSsheet.Range("CP" & n & ":FS" & n).Value = _
            PCSsheet.Range("A" & PCSn & ":CD" & PCSn).Value
            n = n + 1
            c = 2
        ElseIf PCSsheet.Cells(PCSn, 1).Value = HPsheet.Cells(HPn, 1).Value _
        And c = 2 Then
            PNSsheet.Range("CP" & n & ":FS" & n).Value = _
            PCSsheet.Range("A" & PCSn & ":CD" & PCSn).Value
            n = n + 1
        End If
    Next PCSn
Next HPn

End Sub

Sub DateSorter()

Dim d1 As String
Dim d2 As String
Dim d3 As String
Dim d4 As String
Dim d5 As String
Dim d6 As String
Dim d7 As String
Dim date1 As Date
Dim date2 As Date
Dim date3 As Date
Dim date4 As Date
Dim date5 As Date
Dim date6 As Date
Dim date7 As Date
Dim dateMax As Date
Dim Match As Boolean
Dim MatchNum As Boolean
Dim a As Long
Dim b As Integer
Dim n As Long
Dim bMax As Integer
Dim MaxLocation As Integer
Dim PNSsheet As Worksheet
Dim LastRowPNS As Long
Set PNSsheet = Worksheets("propertyNameSort")
PNSsheet.Range("A1").EntireColumn.Insert
LastRowPNS = lRow(PNSsheet)
a = 2
For n = 1 To LastRowPNS
b = 1
If IsEmpty(PNSsheet.Cells(a + 1, 5)) = True Then
    b = 2
    d1 = PNSsheet.Cells(a, 104).Value
    If d1 = "" Then
        d1 = "01/01/1700"
    End If
    date1 = CDate(d1)
    d2 = PNSsheet.Cells(a + 1, 104).Value
    If d2 = "" Then
        d2 = "01/01/1700"
    End If
    date2 = CDate(d2)
    If IsEmpty(PNSsheet.Cells(a + 2, 5)) = True Then
        b = 3
        d3 = PNSsheet.Cells(a + 2, 104).Value
        If d3 = "" Then
        d3 = "01/01/1700"
        End If
        date3 = CDate(d3)
        If IsEmpty(PNSsheet.Cells(a + 3, 5)) = True Then
            b = 4
            d4 = PNSsheet.Cells(a + 3, 104)
            If d4 = "" Then
            d4 = "01/01/1700"
            End If
            date4 = CDate(d4)
            If IsEmpty(PNSsheet.Cells(a + 4, 5)) = True Then
                b = 5
                d5 = PNSsheet.Cells(a + 4, 104).Value
                If d5 = "" Then
                d5 = "01/01/1700"
                End If
                date5 = CDate(d5)
                If IsEmpty(PNSsheet.Cells(a + 5, 5)) = True Then
                    b = 6
                    d6 = PNSsheet.Cells(a + 5, 104).Value
                    If d6 = "" Then
                    d6 = "01/01/1700"
                    End If
                    date6 = CDate(d6)
                    If IsEmpty(PNSsheet.Cells(a + 6, 5)) = True Then
                        b = 7
                        d7 = PNSsheet.Cells(a + 6, 104).Value
                        If d7 = "" Then
                        d7 = "01/01/1700"
                        End If
                        date7 = CDate(d7)
                    End If
                End If
            End If
        End If
    End If
End If

If b = 2 Or b = 3 Or b = 4 Or b = 5 Or b = 6 Or b = 7 Then
    If PNSsheet.Cells(a + 1, 97).Value = PNSsheet.Cells(a, 97).Value Then
        If date1 < date2 Then
            PNSsheet.Range("CP" & a & ":FS" & a).Clear
        ElseIf date2 < date1 Then
            PNSsheet.Range("CP" & a + 1 & ":FS" & a + 1).Clear
        End If
    End If
End If

If b = 3 Or b = 4 Or b = 5 Or b = 6 Or b = 7 Then
    If PNSsheet.Cells(a + 2, 97).Value = PNSsheet.Cells(a + 0, 97).Value Then
        If date1 < date3 Then
            PNSsheet.Range("CP" & a & ":FS" & a).Clear
        ElseIf date3 < date1 Then
            PNSsheet.Range("CP" & a + 2 & ":FS" & a + 2).Clear
        End If
    ElseIf PNSsheet.Cells(a + 2, 97).Value = PNSsheet.Cells(a + 1, 97).Value Then
        If date2 < date3 Then
            PNSsheet.Range("CP" & a + 1 & ":FS" & a + 1).Clear
        ElseIf date3 < date2 Then
            PNSsheet.Range("CP" & a + 2 & ":FS" & a + 2).Clear
        End If
    End If
End If

If b = 4 Or b = 5 Or b = 6 Or b = 7 Then
    If PNSsheet.Cells(a + 3, 97).Value = PNSsheet.Cells(a, 97).Value Then
        If date1 < date4 Then
            PNSsheet.Range("CP" & a & ":FS" & a).Clear
        ElseIf date4 < date1 Then
            PNSsheet.Range("CP" & a + 3 & ":FS" & a + 3).Clear
        End If
    ElseIf PNSsheet.Cells(a + 3, 97).Value = PNSsheet.Cells(a + 1, 97).Value Then
        If date2 < date4 Then
            PNSsheet.Range("CP" & a + 1 & ":FS" & a + 1).Clear
        ElseIf date4 < date2 Then
            PNSsheet.Range("CP" & a + 3 & ":FS" & a + 3).Clear
        End If
    ElseIf PNSsheet.Cells(a + 3, 97).Value = PNSsheet.Cells(a + 2, 97).Value Then
        If date3 < date4 Then
            PNSsheet.Range("CP" & a + 2 & ":FS" & a + 2).Clear
        ElseIf date4 < date3 Then
            PNSsheet.Range("CP" & a + 3 & ":FS" & a + 3).Clear
        End If
    End If
End If

If b = 5 Or b = 6 Or b = 7 Then
    If PNSsheet.Cells(a + 4, 97).Value = PNSsheet.Cells(a, 97).Value Then
        If date1 < date5 Then
            PNSsheet.Range("CP" & a & ":FS" & a).Clear
        ElseIf date5 < date1 Then
            PNSsheet.Range("CP" & a + 4 & ":FS" & a + 4).Clear
        End If
    ElseIf PNSsheet.Cells(a + 4, 97).Value = PNSsheet.Cells(a + 1, 97).Value Then
        If date2 < date5 Then
            PNSsheet.Range("CP" & a + 1 & ":FS" & a + 1).Clear
        ElseIf date5 < date2 Then
            PNSsheet.Range("CP" & a + 4 & ":FS" & a + 4).Clear
        End If
    ElseIf PNSsheet.Cells(a + 4, 97).Value = PNSsheet.Cells(a + 2, 97).Value Then
        If date3 < date5 Then
            PNSsheet.Range("CP" & a + 2 & ":FS" & a + 2).Clear
        ElseIf date5 < date3 Then
            PNSsheet.Range("CP" & a + 4 & ":FS" & a + 4).Clear
        End If
     ElseIf PNSsheet.Cells(a + 4, 97).Value = PNSsheet.Cells(a + 3, 97).Value Then
        If date4 < date5 Then
            PNSsheet.Range("CP" & a + 3 & ":FS" & a + 3).Clear
        ElseIf date5 < date4 Then
            PNSsheet.Range("CP" & a + 4 & ":FS" & a + 4).Clear
        End If
    End If
End If

If b = 6 Or b = 7 Then
    If PNSsheet.Cells(a + 5, 97).Value = PNSsheet.Cells(a, 97).Value Then
        If date1 < date6 Then
            PNSsheet.Range("CP" & a & ":FS" & a).Clear
        ElseIf date6 < date1 Then
            PNSsheet.Range("CP" & a + 5 & ":FS" & a + 5).Clear
        End If
    ElseIf PNSsheet.Cells(a + 5, 97).Value = PNSsheet.Cells(a + 1, 97).Value Then
        If date2 < date6 Then
            PNSsheet.Range("CP" & a + 1 & ":FS" & a + 1).Clear
        ElseIf date6 < date2 Then
            PNSsheet.Range("CP" & a + 5 & ":FS" & a + 5).Clear
        End If
    ElseIf PNSsheet.Cells(a + 5, 97).Value = PNSsheet.Cells(a + 2, 97).Value Then
        If date3 < date6 Then
            PNSsheet.Range("CP" & a + 2 & ":FS" & a + 2).Clear
        ElseIf date6 < date3 Then
            PNSsheet.Range("CP" & a + 5 & ":FS" & a + 5).Clear
        End If
     ElseIf PNSsheet.Cells(a + 5, 97).Value = PNSsheet.Cells(a + 3, 97).Value Then
        If date4 < date6 Then
            PNSsheet.Range("CP" & a + 3 & ":FS" & a + 3).Clear
        ElseIf date6 < date4 Then
            PNSsheet.Range("CP" & a + 5 & ":FS" & a + 5).Clear
        End If
     ElseIf PNSsheet.Cells(a + 5, 97).Value = PNSsheet.Cells(a + 4, 97).Value Then
        If date5 < date6 Then
            PNSsheet.Range("CP" & a + 4 & ":FS" & a + 4).Clear
        ElseIf date6 < date5 Then
            PNSsheet.Range("CP" & a + 5 & ":FS" & a + 5).Clear
        End If
    End If
End If

If b = 7 Then
    If PNSsheet.Cells(a + 6, 97).Value = PNSsheet.Cells(a, 97).Value Then
        If date1 < date7 Then
            PNSsheet.Range("CP" & a & ":FS" & a).Clear
        ElseIf date7 < date1 Then
            PNSsheet.Range("CP" & a + 6 & ":FS" & a + 6).Clear
        End If
    ElseIf PNSsheet.Cells(a + 6, 97).Value = PNSsheet.Cells(a + 1, 97).Value Then
        If date2 < date7 Then
            PNSsheet.Range("CP" & a + 1 & ":FS" & a + 1).Clear
        ElseIf date7 < date2 Then
            PNSsheet.Range("CP" & a + 6 & ":FS" & a + 6).Clear
        End If
    ElseIf PNSsheet.Cells(a + 6, 97).Value = PNSsheet.Cells(a + 2, 97).Value Then
        If date3 < date7 Then
            PNSsheet.Range("CP" & a + 2 & ":FS" & a + 2).Clear
        ElseIf date7 < date3 Then
            PNSsheet.Range("CP" & a + 6 & ":FS" & a + 6).Clear
        End If
     ElseIf PNSsheet.Cells(a + 6, 97).Value = PNSsheet.Cells(a + 3, 97).Value Then
        If date4 < date7 Then
            PNSsheet.Range("CP" & a + 3 & ":FS" & a + 3).Clear
        ElseIf date7 < date4 Then
            PNSsheet.Range("CP" & a + 6 & ":FS" & a + 6).Clear
        End If
     ElseIf PNSsheet.Cells(a + 6, 97).Value = PNSsheet.Cells(a + 4, 97).Value Then
        If date5 < date7 Then
            PNSsheet.Range("CP" & a + 4 & ":FS" & a + 4).Clear
        ElseIf date7 < date5 Then
            PNSsheet.Range("CP" & a + 6 & ":FS" & a + 6).Clear
        End If
    ElseIf PNSsheet.Cells(a + 6, 97).Value = PNSsheet.Cells(a + 5, 97).Value Then
        If date6 < date7 Then
            PNSsheet.Range("CP" & a + 5 & ":FS" & a + 5).Clear
        ElseIf date7 < date6 Then
            PNSsheet.Range("CP" & a + 6 & ":FS" & a + 6).Clear
        End If
    End If
End If

a = a + b
If a = LastRowPNS Or a > LastRowPNS Then
    Exit Sub
End If
'If a > 86 Then
'a = "q"
'End If
Next n
End Sub

Sub archetype(sheet1 As Worksheet, sheet2 As Worksheet, LastRowPNS As Long)

Dim m As Long
Dim n As Long

sheet2.Cells(1, 3).Value = "Archetype"
m = 2
For n = 2 To LastRowPNS
If IsEmpty(sheet1.Cells(n, 4)) = False Then
    If sheet1.Cells(n, 22).Value = 3 And _
        sheet1.Cells(n, 24).Value = 6 Then
            sheet2.Cells(m, 3).Value = "1b"
    ElseIf sheet1.Cells(n, 22).Value = 3 And _
        sheet1.Cells(n, 24).Value = 7 Then
            sheet2.Cells(m, 3).Value = "1c"
    ElseIf sheet1.Cells(n, 22).Value = 4 And _
        sheet1.Cells(n, 24).Value = 7 Then
            sheet2.Cells(m, 3).Value = "2a"
    ElseIf sheet1.Cells(n, 22).Value = 6 And _
        sheet1.Cells(n, 24).Value = 5 Then
            sheet2.Cells(m, 3).Value = "4a"
    ElseIf sheet1.Cells(n, 22).Value = 6 And _
        sheet1.Cells(n, 24).Value = 7 Then
            sheet2.Cells(m, 3).Value = "4b"
    ElseIf sheet1.Cells(n, 22).Value = 7 And _
        sheet1.Cells(n, 24).Value = 7 Then
            sheet2.Cells(m, 3).Value = "5a"
    ElseIf sheet1.Cells(n, 22).Value = 3 And _
        sheet1.Cells(n, 30).Value = 3 Then
            sheet2.Cells(m, 3).Value = "1a"
    ElseIf sheet1.Cells(n, 22).Value = 5 And _
        sheet1.Cells(n, 30).Value = 1 Then
            sheet2.Cells(m, 3).Value = "3a"
    ElseIf sheet1.Cells(n, 22).Value = 6 And _
        sheet1.Cells(n, 30).Value = 1 Then
            sheet2.Cells(m, 3).Value = "4c"
    ElseIf sheet1.Cells(n, 22).Value = 7 And _
        sheet1.Cells(n, 30).Value = 1 Then
            sheet2.Cells(m, 3).Value = "5b"
    End If
    m = m + 1
End If
Next n

End Sub

Sub DwellingType(sheet1 As Worksheet, sheet2 As Worksheet, LastRowPNS As Long)
Dim n As Integer
Dim m As Integer
Dim b As Integer

sheet2.Cells(1, 1).Value = "DwellingType"
m = 2
For n = 2 To LastRowPNS
If IsEmpty(sheet1.Cells(n, 4)) = False Then
        If IsEmpty(sheet1.Cells(n, 95)) = False Then
            If sheet1.Cells(n, 102).Value = "Flat" Then
                sheet2.Cells(m, 1).Value = 1
            ElseIf sheet1.Cells(n, 102).Value = "Maisonette" Then
                sheet2.Cells(m, 1).Value = 3
            ElseIf sheet1.Cells(n, 102).Value = "House" Then
                sheet2.Cells(m, 1).Value = 2
            ElseIf sheet1.Cells(n, 102).Value = "Bungalow" Then
                sheet2.Cells(m, 1).Value = 3
            Else
                sheet2.Cells(m, 1).Value = "NaN"
            End If
        ElseIf IsEmpty(sheet1.Cells(n, 95)) = True Then
            b = blankspace(n, 95)
            If sheet1.Cells(n + b + 1, 102).Value = "Flat" Then
                sheet2.Cells(m, 1).Value = 1
            ElseIf sheet1.Cells(n + b + 1, 102).Value = "Maisonette" Then
                sheet2.Cells(m, 1).Value = 3
            ElseIf sheet1.Cells(n + b + 1, 102).Value = "House" Then
                sheet2.Cells(m, 1).Value = 2
            ElseIf sheet1.Cells(n + b + 1, 102).Value = "Bungalow" Then
                sheet2.Cells(m, 1).Value = 3
            Else
                sheet2.Cells(m, 1).Value = "NaN"
            End If
        End If
    m = m + 1
End If
Next n

End Sub

Sub DwellingPosition(sheet1 As Worksheet, sheet2 As Worksheet, LastRowPNS As Long)
Dim n As Integer
Dim m As Integer
Dim b As Integer

sheet2.Cells(1, 2).Value = "DwellingPosition"
m = 2
For n = 2 To LastRowPNS
If IsEmpty(sheet1.Cells(n, 4)) = False Then
        If IsEmpty(sheet1.Cells(n, 95)) = False Then
            If sheet2.Cells(m, 1).Value = 2 Or sheet2.Cells(m, 1).Value = 0 Then
                If sheet1.Cells(n, 129).Value = "Detached" Then
                    sheet2.Cells(m, 2).Value = 0
                ElseIf sheet1.Cells(n, 129).Value = "End-Terrace" Then
                    sheet2.Cells(m, 2).Value = 1
                ElseIf sheet1.Cells(n, 129).Value = "Mid-Terrace" Then
                    sheet2.Cells(m, 2).Value = 2
                ElseIf sheet1.Cells(n, 129).Value = "Semi-Detached" Then
                    sheet2.Cells(m, 2).Value = 3
                End If
            ElseIf sheet2.Cells(m, 1).Value = 1 Or sheet2.Cells(m, 1).Value = 3 Then
                If sheet1.Cells(n, 128).Value = "Y" Then
                    sheet2.Cells(m, 2).Value = 6
                ElseIf sheet1.Cells(n, 128).Value = "N" _
                And sheet1.Cells(n, 127).Value < 3 Then
                    sheet2.Cells(m, 2).Value = 4
                ElseIf sheet1.Cells(n, 128).Value = "N" _
                And sheet1.Cells(n, 127).Value > 2 And _
                sheet1.Cells(n, 126).Value <> "Ground" Then
                    sheet2.Cells(m, 2).Value = 5
                ElseIf sheet1.Cells(n, 128).Value = "N" _
                And sheet1.Cells(n, 127).Value > 2 And _
                sheet1.Cells(n, 126).Value = "Ground" Then
                    sheet2.Cells(m, 2).Value = 4
                End If
            End If
        ElseIf IsEmpty(sheet1.Cells(n, 95)) = True Then
            b = blankspace(n, 95) + 1
            If sheet2.Cells(m, 1).Value = 2 Or sheet2.Cells(m, 1).Value = 0 Then
                If sheet1.Cells(n + b, 129).Value = "Detached" Then
                    sheet2.Cells(m, 2).Value = 0
                ElseIf sheet1.Cells(n + b, 129).Value = "End-Terrace" Then
                    sheet2.Cells(m, 2).Value = 1
                ElseIf sheet1.Cells(n + b, 129).Value = "Mid-Terrace" Then
                    sheet2.Cells(m, 2).Value = 2
                ElseIf sheet1.Cells(n + b, 129).Value = "Semi-Detached" Then
                    sheet2.Cells(m, 2).Value = 3
                End If
            ElseIf sheet2.Cells(m, 1).Value = 1 Or sheet2.Cells(m, 1).Value = 3 Then
                If sheet1.Cells(n + b, 128).Value = "Y" Then
                    sheet2.Cells(m, 2).Value = 6
                ElseIf sheet1.Cells(n + b, 128).Value = "N" _
                And sheet1.Cells(n + b, 127).Value < 3 Then
                    sheet2.Cells(m, 2).Value = 4
                ElseIf sheet1.Cells(n + b, 128).Value = "N" _
                And sheet1.Cells(n + b, 127).Value > 2 And _
                sheet1.Cells(n + b, 126).Value <> "Ground" Then
                    sheet2.Cells(m, 2).Value = 5
                ElseIf sheet1.Cells(n + b, 128).Value = "N" _
                And sheet1.Cells(n + b, 127).Value > 2 And _
                sheet1.Cells(n + b, 126).Value = "Ground" Then
                    sheet2.Cells(m, 2).Value = 4
                End If
            End If
        End If
    m = m + 1
End If
Next n

For m = 2 To m - 1
If IsEmpty(sheet2.Cells(m, 2)) = True Then
    If sheet2.Cells(m, 1).Value = 2 Then
        sheet2.Cells(m, 2).Value = 2
    ElseIf sheet2.Cells(m, 1).Value = 1 Or sheet2.Cells(m, 1).Value = 3 Then
        sheet2.Cells(m, 2).Value = 4
    ElseIf sheet2.Cells(m, 1).Value = 0 Then
        sheet2.Cells(m, 2).Value = 0
    End If
End If
Next m

End Sub

Sub averageEPC(variable As String, columnSTnum As Integer, columnSTstring As String, _
columnHAnum As Integer, columnHAstring As String)

Dim LastRowHA As Long
Dim LastRowPNS As Long
Dim LastRowST As Long
Dim sheet1 As Worksheet
Dim sheet2 As Worksheet
Dim PNSsheet As Worksheet
Dim n As Integer
Dim m As Integer
Dim b As Integer
Dim rng As Range
Dim av As Double
Dim av1a As Double, av1b As Double, av1c As Double, av2a As Double, av3a As Double, _
av4a As Double, av4b As Double, av4c As Double, av5a As Double, av5b As Double
Dim avAll As Double

Set sheet1 = Worksheets("Haringey_All")
Set sheet2 = Worksheets("SUSDEMtext")
Set PNSsheet = Worksheets("propertyNameSort")

LastRowHA = lRow(sheet1)
LastRowPNS = lRow(PNSsheet)

sheet2.Cells(1, columnSTnum).Value = variable
m = 2
For n = 2 To LastRowPNS
If IsEmpty(sheet1.Cells(n, 4)) = False Then
    b = blankspace(n, 4) 'Finds number of rows with EPC data
    Set rng = sheet1.Range(columnHAstring & n & ":" & columnHAstring & n + b) 'Set range to include all EPC data for residence
    av = customAverageRooms(rng, 1, 100, False, "NaN", b, sheet1, sheet2, False, _
    columnHAnum, columnSTnum, n) 'Find average of desired variable
    sheet2.Cells(m, columnSTnum).Value = av 'Sets SUSDEMtext cell to this average
    m = m + 1
End If
Next n
LastRowST = lRow(sheet2)
Set rng = sheet2.Range(columnSTstring & "2:" & columnSTstring & LastRowST)
avAll = customAverageRooms(rng, 0, 100, False, "1a", b, sheet1, sheet2, True, columnHAnum, columnSTnum, n) 'Finds averages for all residences
av1a = customAverageRooms(rng, 0, 100, True, "1a", b, sheet1, sheet2, False, columnHAnum, columnSTnum, n) 'Finds averages for all different archetypes
av1b = customAverageRooms(rng, 0, 100, True, "1b", b, sheet1, sheet2, False, columnHAnum, columnSTnum, n)
av1c = customAverageRooms(rng, 0, 100, True, "1c", b, sheet1, sheet2, False, columnHAnum, columnSTnum, n)
av2a = customAverageRooms(rng, 0, 100, True, "2a", b, sheet1, sheet2, False, columnHAnum, columnSTnum, n)
av3a = customAverageRooms(rng, 0, 100, True, "3a", b, sheet1, sheet2, False, columnHAnum, columnSTnum, n)
av4a = customAverageRooms(rng, 0, 100, True, "4a", b, sheet1, sheet2, False, columnHAnum, columnSTnum, n)
av4b = customAverageRooms(rng, 0, 100, True, "4b", b, sheet1, sheet2, False, columnHAnum, columnSTnum, n)
av4c = customAverageRooms(rng, 0, 100, True, "4c", b, sheet1, sheet2, False, columnHAnum, columnSTnum, n)
av5a = customAverageRooms(rng, 0, 100, True, "5a", b, sheet1, sheet2, False, columnHAnum, columnSTnum, n)
av5b = customAverageRooms(rng, 0, 100, True, "5b", b, sheet1, sheet2, False, columnHAnum, columnSTnum, n)

For m = 2 To LastRowST 'If number of rooms is not available, then assign the average depending on the archetype
If sheet2.Cells(m, 3).Value = "1a" And sheet2.Cells(m, columnSTstring).Value = 20000 Then
    sheet2.Cells(m, columnSTstring).Value = av1a
ElseIf sheet2.Cells(m, 3).Value = "1b" And sheet2.Cells(m, columnSTstring).Value = 20000 Then
    sheet2.Cells(m, columnSTstring).Value = av1b
ElseIf sheet2.Cells(m, 3).Value = "1c" And sheet2.Cells(m, columnSTstring).Value = 20000 Then
    sheet2.Cells(m, columnSTstring).Value = av1c
ElseIf sheet2.Cells(m, 3).Value = "2a" And sheet2.Cells(m, columnSTstring).Value = 20000 Then
    sheet2.Cells(m, columnSTstring).Value = av2a
ElseIf sheet2.Cells(m, 3).Value = "3a" And sheet2.Cells(m, columnSTstring).Value = 20000 Then
    sheet2.Cells(m, columnSTstring).Value = av3a
ElseIf sheet2.Cells(m, 3).Value = "4a" And sheet2.Cells(m, columnSTstring).Value = 20000 Then
    sheet2.Cells(m, columnSTstring).Value = av4a
ElseIf sheet2.Cells(m, 3).Value = "4b" And sheet2.Cells(m, columnSTstring).Value = 20000 Then
    sheet2.Cells(m, columnSTstring).Value = av4b
ElseIf sheet2.Cells(m, 3).Value = "4c" And sheet2.Cells(m, columnSTstring).Value = 20000 Then
    sheet2.Cells(m, columnSTstring).Value = av4c
ElseIf sheet2.Cells(m, 3).Value = "5a" And sheet2.Cells(m, columnSTstring).Value = 20000 Then
    sheet2.Cells(m, columnSTstring).Value = av5a
ElseIf sheet2.Cells(m, 3).Value = "5b" And sheet2.Cells(m, columnSTstring).Value = 20000 Then
    sheet2.Cells(m, columnSTstring).Value = av5b
ElseIf IsEmpty(sheet2.Cells(m, 3)) And sheet2.Cells(m, columnSTstring).Value = 20000 Then 'If reisdence is not contained in an archetype, then apply average of all residences
    sheet2.Cells(m, columnSTstring).Value = avAll
End If
Next m

End Sub

Sub averageEPCenergy(variable As String, columnSTnum As Integer, columnSTstring As String, _
columnHAnum As Integer, columnHAstring As String)

Dim LastRowHA As Long
Dim LastRowPNS As Long
Dim LastRowST As Long
Dim sheet1 As Worksheet
Dim sheet2 As Worksheet
Dim PNSsheet As Worksheet
Dim n As Integer
Dim m As Integer
Dim b As Integer
Dim rng As Range
Dim av As Double
Dim av1a As Double, av1b As Double, av1c As Double, av2a As Double, av3a As Double, _
av4a As Double, av4b As Double, av4c As Double, av5a As Double, av5b As Double
Dim avAll As Double

Set sheet1 = Worksheets("Haringey_All")
Set sheet2 = Worksheets("SUSDEMtext")
Set PNSsheet = Worksheets("propertyNameSort")

LastRowHA = lRow(sheet1)
LastRowPNS = lRow(PNSsheet)

sheet2.Cells(1, columnSTnum).Value = variable
m = 2
For n = 2 To LastRowPNS
If IsEmpty(sheet1.Cells(n, 4)) = False Then
    b = blankspace(n, 4) 'Finds number of rows with EPC data
    Set rng = sheet1.Range(columnHAstring & n & ":" & columnHAstring & n + b) 'Set range to include all EPC data for residence
    av = customAverageRooms(rng, 1, 20000, False, "NaN", b, sheet1, sheet2, False, _
    columnHAnum, columnSTnum, n) 'Find average of desired variable
    sheet2.Cells(m, columnSTnum).Value = av 'Sets SUSDEMtext cell to this average
    m = m + 1
End If
Next n
For n = 2 To LastRowPNS
If sheet2.Cells(n, columnSTnum).Value = 20000 Then
sheet2.Cells(n, columnSTnum).Value = 0
End If
Next n
End Sub

Sub NoOfStoreys()

Dim LastRowHA As Long
Dim LastRowPNS As Long
Dim LastRowST As Long
Dim sheet1 As Worksheet
Dim sheet2 As Worksheet
Dim PNSsheet As Worksheet
Dim n As Integer
Dim m As Integer
Dim b As Integer

Set sheet1 = Worksheets("Haringey_All")
Set sheet2 = Worksheets("SUSDEMtext")
Set PNSsheet = Worksheets("propertyNameSort")
LastRowHA = lRow(sheet1)
LastRowPNS = lRow(PNSsheet)
sheet2.Cells(1, 6).Value = "NoOfStoreys"
m = 2
For n = 2 To LastRowPNS
    If IsEmpty(sheet1.Cells(n, 4)) = False Then
    If sheet2.Cells(m, 3).Value = "1a" Or sheet2.Cells(m, 3).Value = "1b" _
    Or sheet2.Cells(m, 3).Value = "1c" Or sheet2.Cells(m, 3).Value = "2a" _
    Or sheet2.Cells(m, 3).Value = "" Then
        If sheet1.Cells(n, 19).Value >= 12 Then
            sheet2.Cells(m, 6).Value = 3
        ElseIf sheet1.Cells(n, 19).Value < 12 Then
            sheet2.Cells(m, 6).Value = 2
        End If
    ElseIf sheet2.Cells(m, 3).Value = "3a" Then
        If sheet1.Cells(n, 19).Value <= 11.4 Then
            sheet2.Cells(m, 6).Value = 2
        ElseIf sheet1.Cells(n, 19).Value > 13 And _
        sheet1.Cells(n, 19).Value < 13.8 Then
            sheet2.Cells(m, 6).Value = 3
        ElseIf sheet1.Cells(n, 19).Value >= 13.8 Then
            sheet2.Cells(m, 6).Value = 3
        End If
    ElseIf sheet2.Cells(m, 3).Value = "4a" Then
        sheet2.Cells(m, 6).Value = 3
    ElseIf sheet2.Cells(m, 3).Value = "4b" Then
        If sheet1.Cells(n, 19).Value < 10.3 Then
            sheet2.Cells(m, 6).Value = 2
        ElseIf sheet1.Cells(n, 19).Value >= 10.3 Then
            sheet2.Cells(m, 6).Value = 3
        End If
    ElseIf sheet2.Cells(m, 3).Value = "4c" Then
        If sheet1.Cells(n, 19).Value <= 10.8 Then
            sheet2.Cells(m, 6).Value = 3
        ElseIf sheet1.Cells(n, 19).Value > 10.8 Then
            sheet2.Cells(m, 6).Value = 4
        End If
    ElseIf sheet2.Cells(m, 3).Value = "5a" Then
        sheet2.Cells(m, 6).Value = 2
    ElseIf sheet2.Cells(m, 3).Value = "5b" Then
        If sheet1.Cells(n, 19).Value <= 14.4 Then
            sheet2.Cells(m, 6).Value = 3
        ElseIf sheet1.Cells(n, 19).Value > 14.5 Then
            sheet2.Cells(m, 6).Value = 4
        End If
    End If
    m = m + 1
    End If
Next n

End Sub

Sub BuildingHeight(HAsheet As Worksheet, STsheet As Worksheet, LastRowPNS As Long)
Dim m As Long
Dim n As Long

STsheet.Cells(1, 7).Value = "BuildingHeight"
m = 2
For n = 2 To LastRowPNS
    If IsEmpty(HAsheet.Cells(n, 4)) = False Then
        STsheet.Cells(m, 7).Value = HAsheet.Cells(n, 19)
        m = m + 1
    End If
Next n

End Sub

Sub fid(HAsheet As Worksheet, STsheet As Worksheet, LastRowPNS As Long)
Dim m As Long
Dim n As Long

STsheet.Cells(1, 26).Value = "FID"
m = 2
For n = 2 To LastRowPNS
    If IsEmpty(HAsheet.Cells(n, 4)) = False Then
        STsheet.Cells(m, 26).Value = HAsheet.Cells(n, 3)
        m = m + 1
    End If
Next n

End Sub

Sub Address(HAsheet As Worksheet, STsheet As Worksheet, LastRowPNS As Long)
Dim m As Long
Dim n As Long

STsheet.Cells(1, 8).Value = "Address"
m = 2
For n = 2 To LastRowPNS
    If IsEmpty(HAsheet.Cells(n, 4)) = False Then
        STsheet.Cells(m, 8).Value = HAsheet.Cells(n, 1).Value & ", " & _
        HAsheet.Cells(n, 63).Value & ", London"
        m = m + 1
    End If
Next n

End Sub

Sub areas(HAsheet As Worksheet, STsheet As Worksheet, LastRowPNS As Long)
Dim m As Long
Dim n As Long
Dim nfloors As Integer

STsheet.Cells(1, 9).Value = "GroundFloorArea"
STsheet.Cells(1, 10).Value = "FirstFloorArea"
STsheet.Cells(1, 11).Value = "SecondFloorArea"

m = 2
For n = 2 To LastRowPNS
    If IsEmpty(HAsheet.Cells(n, 4)) = False Then
        nfloors = STsheet.Cells(m, 6).Value
        If nfloors > 0 Then
            STsheet.Cells(m, 9).Value = HAsheet.Cells(n, 16).Value
        End If
        If nfloors > 1 Then
            STsheet.Cells(m, 10).Value = HAsheet.Cells(n, 16).Value
        End If
        If nfloors = 3 Then
            STsheet.Cells(m, 11).Value = HAsheet.Cells(n, 16).Value
        End If
    m = m + 1
    End If
Next n

End Sub

Sub ageBand(HAsheet As Worksheet, STsheet As Worksheet, LastRowPNS As Long)
Dim m As Long
Dim n As Long
Dim randomValue As Single

STsheet.Cells(1, 18).Value = "AgeBandCode"
m = 2
For n = 2 To LastRowPNS
    If IsEmpty(HAsheet.Cells(n, 4)) = False Then
        randomValue = CInt(Rnd()) 'Random number and weighting used to assign age band
        If IsEmpty(HAsheet.Cells(n, 22)) = True Then
            STsheet.Cells(m, 18).Value = ""
        ElseIf HAsheet.Cells(n, 22) = 3 Then
            If randomValue <= 0.68 Then
                STsheet.Cells(m, 18).Value = 1
            ElseIf randomValue > 0.68 Then
                STsheet.Cells(m, 18).Value = 2
            End If
        ElseIf HAsheet.Cells(n, 22) = 4 Then
            If randomValue <= 0.57 Then
                STsheet.Cells(m, 18).Value = 2
            ElseIf randomValue > 0.57 Then
                STsheet.Cells(m, 18).Value = 3
            End If
        ElseIf HAsheet.Cells(n, 22) = 5 Then
            If randomValue <= 0.29 Then
                STsheet.Cells(m, 18).Value = 3
            ElseIf randomValue > 0.29 Then
                STsheet.Cells(m, 18).Value = 4
            End If
        ElseIf HAsheet.Cells(n, 22) = 6 Then
            If randomValue <= 0.35 Then
                STsheet.Cells(m, 18).Value = 4
            ElseIf randomValue > 0.35 And randomValue <= 0.45 Then
                STsheet.Cells(m, 18).Value = 5
            ElseIf randomValue > 0.45 Then
                STsheet.Cells(m, 18).Value = 6
            End If
        ElseIf HAsheet.Cells(n, 22) = 7 Then
            If randomValue <= 0.1875 Then
                STsheet.Cells(m, 18).Value = 7
            ElseIf randomValue > 0.5 And randomValue <= 0.45 Then
                STsheet.Cells(m, 18).Value = 8
            ElseIf randomValue > 0.3125 Then
                STsheet.Cells(m, 18).Value = 9
            End If
        End If
        m = m + 1
    End If
Next n

End Sub

Sub externalwallconstruction(HAsheet As Worksheet, STsheet As Worksheet, LastRowPNS As Long)
Dim m As Long
Dim n As Integer
Dim b As Integer
Dim wallType As String

STsheet.Cells(1, 19).Value = "ExternalWall1"
STsheet.Cells(1, 20).Value = "ExternalWall2"

m = 2
For n = 2 To LastRowPNS
    If IsEmpty(STsheet.Cells(m, 4)) = False Then
        If IsEmpty(HAsheet.Cells(n, 95)) = False Then
            b = -1
        ElseIf IsEmpty(HAsheet.Cells(n, 95)) = True Then
            b = blankspace(n, 95)
        End If
        wallType = HAsheet.Cells(n + b + 1, 150).Value
        If Left(wallType, 7) = "Average" Then
            STsheet.Cells(m, 19).Value = 0
            STsheet.Cells(m, 20).Value = 3 'Closest to high thermal transmitance in HA all
        ElseIf Trim(wallType) = "Cavity wall, as built, insulated (assumed)" Then
            STsheet.Cells(m, 19).Value = 0
            STsheet.Cells(m, 20).Value = 3
        ElseIf Trim(wallType) = "Cavity wall, as built, no insulation (assumed)" Then
            STsheet.Cells(m, 19).Value = 0
            STsheet.Cells(m, 20).Value = 1
        ElseIf Trim(wallType) = "Cavity wall, as built, partial insulation (assumed)" Then
            STsheet.Cells(m, 19).Value = 0
            STsheet.Cells(m, 20).Value = 3
        ElseIf Trim(wallType) = "Cavity wall, filled cavity" Then
            STsheet.Cells(m, 19).Value = 0
            STsheet.Cells(m, 20).Value = 3
        ElseIf Trim(wallType) = "Cavity wall, with external insulation" Then
            STsheet.Cells(m, 19).Value = 0
            STsheet.Cells(m, 20).Value = 2
        ElseIf Trim(wallType) = "Solid brick, as built, insulated (assumed)" Then
            STsheet.Cells(m, 19).Value = 1
            STsheet.Cells(m, 20).Value = 3
        ElseIf Trim(wallType) = "Solid brick, as built, no insulation (assumed)" Then
            STsheet.Cells(m, 19).Value = 1
            STsheet.Cells(m, 20).Value = 1
        ElseIf Trim(wallType) = "Solid brick, as built, partial insulation (assumed)" Then
            STsheet.Cells(m, 19).Value = 1
            STsheet.Cells(m, 20).Value = 3
        ElseIf Trim(wallType) = "Solid brick, with internal insulation" Then
            STsheet.Cells(m, 19).Value = 1
            STsheet.Cells(m, 20).Value = 4
        ElseIf Trim(wallType) = "System built, as built, insulated (assumed)" Then
            STsheet.Cells(m, 19).Value = 2
            STsheet.Cells(m, 20).Value = 3
        ElseIf Trim(wallType) = "System built, as built, no insulation (assumed)" Then
            STsheet.Cells(m, 19).Value = 2
            STsheet.Cells(m, 20).Value = 1
        ElseIf Trim(wallType) = "System built, as built, insulated (assumed)" Then
            STsheet.Cells(m, 19).Value = 2
            STsheet.Cells(m, 20).Value = 3
        ElseIf Trim(wallType) = "System built, as built, partial insulation (assumed)" Then
            STsheet.Cells(m, 19).Value = 2
            STsheet.Cells(m, 20).Value = 3
        ElseIf Trim(wallType) = "Timber frame, as built, insulated (assumed)" Then
            STsheet.Cells(m, 19).Value = 3
            STsheet.Cells(m, 20).Value = 3
        ElseIf Trim(wallType) = "Timber frame, as built, partial insulation (assumed)" Then
            STsheet.Cells(m, 19).Value = 3
            STsheet.Cells(m, 20).Value = 3
        Else
            STsheet.Cells(m, 19).Value = 1
            STsheet.Cells(m, 20).Value = 3 'Most typical, very few buildings don't have a catagory
        End If

'Cavity wall, as built, insulated (assumed), _
'Cavity wall, as built, no insulation (assumed) _
'Cavity wall, as built, partial insulation (assumed) _
'Cavity wall, filled cavity _
'Cavity wall, with external insulation _
'Solid brick, as built, insulated (assumed) _
'Solid brick, as built, no insulation (assumed) _
'Solid brick, as built, partial insulation (assumed) _
'Solid brick, with internal insulation _
'System built, as built, insulated (assumed) _
'System built, as built, no insulation (assumed) _
'System built, as built, partial insulation (assumed) _
'System built, as built, insulated (assumed) _
'System built, as built, no insulation (assumed) _
'System built, as built, partial insulation (assumed) _
'Timber frame, as built, insulated (assumed) _
'Timber frame, as built, partial insulation (assumed) _
'Wall
' Include calculations made earlier to do with building type ect

    End If
    m = m + 1
Next n

End Sub

Sub floorConstruction(HAsheet As Worksheet, STsheet As Worksheet, LastRowPNS As Long)
Dim m As Long
Dim n As Integer
Dim b As Integer
Dim floorType As String

STsheet.Cells(1, 21).Value = "FloorConstruction"

m = 2
For n = 2 To LastRowPNS
    If IsEmpty(STsheet.Cells(m, 4)) = False Then
        If IsEmpty(HAsheet.Cells(n, 95)) = False Then
            b = -1
        ElseIf IsEmpty(HAsheet.Cells(n, 95)) = True Then
            b = blankspace(n, 95)
        End If
        floorType = HAsheet.Cells(n + b + 1, 144).Value
        If Left(Trim(floorType), 5) = "Solid" Then
            STsheet.Cells(m, 21).Value = 1
        ElseIf Left(Trim(floorType), 9) = "Suspended" Then
            STsheet.Cells(m, 21).Value = 2
        Else
            STsheet.Cells(m, 21).Value = 2 'Make this more accurate
        End If
    End If
    m = m + 1
Next n

End Sub

Sub wwr()

Dim LastRowHA As Long
Dim LastRowPNS As Long
Dim sheet1 As Worksheet
Dim sheet2 As Worksheet
Dim PNSsheet As Worksheet
Dim m As Long
Dim n As Long

Set sheet1 = Worksheets("Haringey_All")
Set sheet2 = Worksheets("SUSDEMtext")
Set PNSsheet = Worksheets("propertyNameSort")

LastRowHA = lRow(sheet1)
LastRowPNS = lRow(PNSsheet)

sheet2.Cells(1, 22).Value = "WWR"
m = 2
For n = 2 To LastRowHA
If IsEmpty(sheet1.Cells(n, 4)) = False Then
    If sheet1.Cells(n, 3).Value = "1a" Or sheet1.Cells(n, 3).Value = "1b" Or _
    sheet1.Cells(n, 3).Value = "1c" Or sheet1.Cells(n, 3).Value = "3a" Or _
    sheet1.Cells(n, 3).Value = "4a" Or sheet1.Cells(n, 3).Value = "5b" Or _
    IsEmpty(sheet1.Cells(n, 3)) Then
        sheet2.Cells(m, 22).Value = Application.WorksheetFunction.NormInv(Rnd(), 0.25, 0.001)
        m = m + 1
    ElseIf sheet2.Cells(m, 3).Value = "2a" Or sheet2.Cells(m, 3).Value = "4b" Or _
    sheet2.Cells(m, 3).Value = "4c" Or sheet2.Cells(m, 3).Value = "5b" Then
        sheet2.Cells(m, 22).Value = Application.WorksheetFunction.NormInv(Rnd(), 0.4, 0.001)
        m = m + 1
    Else
        sheet2.Cells(m, 22).Value = Application.WorksheetFunction.NormInv(Rnd(), 0.25, 0.001)
        m = m + 1
    End If
End If
Next n

End Sub

Sub doubleglazing()

Dim LastRowHA As Long
Dim LastRowPNS As Long
Dim sheet1 As Worksheet
Dim sheet2 As Worksheet
Dim PNSsheet As Worksheet
Dim m As Long
Dim n As Long

Set sheet1 = Worksheets("Haringey_All")
Set sheet2 = Worksheets("SUSDEMtext")
Set PNSsheet = Worksheets("propertyNameSort")

LastRowHA = lRow(sheet1)
LastRowPNS = lRow(PNSsheet)

sheet2.Cells(1, 23).Value = "DoubleGlazingInstallation"
m = 2
For n = 2 To LastRowHA
If IsEmpty(sheet1.Cells(n, 4)) = False Then
    If sheet1.Cells(n, 3).Value = "1a" Or sheet1.Cells(n, 3).Value = "1b" Or _
    sheet1.Cells(n, 3).Value = "1c" Or sheet1.Cells(n, 3).Value = "2a" Or _
    sheet1.Cells(n, 3).Value = "3a" Or sheet1.Cells(n, 3).Value = "4a" Or _
    sheet1.Cells(n, 3).Value = "4b" Or sheet1.Cells(n, 3).Value = "4c" Or IsEmpty(sheet1.Cells(n, 3)) Then
        sheet2.Cells(m, 23).Value = Application.WorksheetFunction.NormInv(Rnd(), 65, 5)
        m = m + 1
    ElseIf sheet2.Cells(m, 3).Value = "5a" Or sheet2.Cells(m, 3).Value = "5b" Then
        sheet2.Cells(m, 23).Value = 100
        m = m + 1
    Else
        sheet2.Cells(m, 23).Value = Application.WorksheetFunction.NormInv(Rnd(), 65, 5)
        m = m + 1
    End If
End If
Next n

End Sub

Sub averageEPCroofing(variable As String, columnSTnum As Integer, columnSTstring As String, _
columnHAnum As Integer, columnHAstring As String)

Dim LastRowHA As Long
Dim LastRowPNS As Long
Dim LastRowST As Long
Dim sheet1 As Worksheet
Dim sheet2 As Worksheet
Dim PNSsheet As Worksheet
Dim n As Integer
Dim m As Integer
Dim b As Integer
Dim rng As Range
Dim s As String
Dim ss As Double
Dim av As Double
Dim av1a As Double, av1b As Double, av1c As Double, av2a As Double, av3a As Double, _
av4a As Double, av4b As Double, av4c As Double, av5a As Double, av5b As Double
Dim avAll As Double

Set sheet1 = Worksheets("Haringey_All")
Set sheet2 = Worksheets("SUSDEMtext")
Set PNSsheet = Worksheets("propertyNameSort")

LastRowHA = lRow(sheet1)
LastRowPNS = lRow(PNSsheet)
LastRowST = lRow(sheet2)

sheet1.Range("A1").EntireColumn.Insert
For n = 2 To LastRowPNS
s = onlyDigits(sheet1.Range("FA" & n).Value)
ss = Val(s)
sheet1.Cells(n, 1).Value = ss
Next n

sheet2.Cells(1, columnSTnum).Value = variable
m = 2
For n = 2 To LastRowPNS
If IsEmpty(sheet1.Cells(n, 4)) = False Then
    b = blankspace(n, 4) 'Finds number of rows with EPC data
    Set rng = sheet1.Range(columnHAstring & n & ":" & columnHAstring & n + b) 'Set range to include all EPC data for residence
    av = customAverageRooms(rng, 1, 20000, False, "NaN", b, sheet1, sheet2, False, _
    columnHAnum, columnSTnum, n) 'Find average of desired variable
    sheet2.Cells(m, columnSTnum).Value = av 'Sets SUSDEMtext cell to this average
    m = m + 1
End If
Next n

For n = 2 To LastRowST
If sheet2.Cells(n, columnSTnum).Value = 20000 Then
sheet2.Cells(n, columnSTnum).Value = 0
End If
Next n

sheet1.Range("A1").EntireColumn.Delete
For n = 2 To LastRowST
If sheet2.Cells(n, 28).Value = 0 Then
    If sheet2.Cells(n, 3).Value = "1a" Or sheet2.Cells(n, 3).Value = "1b" Or _
    sheet2.Cells(n, 3).Value = "1c" Or sheet2.Cells(n, 3).Value = "2a" Or _
    sheet2.Cells(n, 3).Value = "3a" Or sheet2.Cells(n, 3).Value = "4a" Or _
    sheet2.Cells(n, 3).Value = "4b" Or sheet2.Cells(n, 3).Value = "4c" Or IsEmpty(sheet2.Cells(n, 3)) Then
        sheet2.Cells(n, 28).Value = Application.WorksheetFunction.NormInv(Rnd(), 100, 10)
    ElseIf sheet2.Cells(n, 3).Value = "5a" Or sheet2.Cells(n, 3).Value = "5b" Then
        sheet2.Cells(n, 28).Value = 300
    Else
        sheet2.Cells(n, 28).Value = Application.WorksheetFunction.NormInv(Rnd(), 100, 10)
    End If
End If
Next n

End Sub

Sub floorHeights(HAsheet As Worksheet, STsheet As Worksheet, LastRowPNS As Long)
Dim m As Long
Dim n As Long
Dim hgt As Double
Dim floorhgt As Double
Dim nfloors As Integer

STsheet.Cells(1, 12).Value = "GroundFloorHeight"
STsheet.Cells(1, 13).Value = "FirstFloorHeight"
STsheet.Cells(1, 14).Value = "SecondFloorHeight"

m = 2
For n = 2 To LastRowPNS
    If IsEmpty(HAsheet.Cells(n, 4)) = False Then
        hgt = STsheet.Cells(m, 7).Value
        nfloors = STsheet.Cells(m, 6).Value
        If STsheet.Cells(m, 3).Value = "1a" Or STsheet.Cells(m, 3).Value = "1b" Or _
        STsheet.Cells(m, 3).Value = "1c" Or STsheet.Cells(m, 3).Value = "2a" Then
            If nfloors = 0 Then
                nfloors = 1
            Else
            floorhgt = (hgt - 2.4 - ((nfloors + 1) * 0.25)) / nfloors
            End If
            If nfloors > 0 Then
                STsheet.Cells(m, 12).Value = floorhgt
            End If
            If nfloors > 1 Then
                STsheet.Cells(m, 13).Value = floorhgt
            End If
            If nfloors = 3 Then
                STsheet.Cells(m, 14).Value = floorhgt
            End If
        Else
            If nfloors = 0 Then
                nfloors = 1
            Else
            floorhgt = (hgt - 1 - ((nfloors + 1) * 0.25)) / nfloors
            End If
            If nfloors > 0 Then
                STsheet.Cells(m, 12).Value = floorhgt
            End If
            If nfloors > 1 Then
                STsheet.Cells(m, 13).Value = floorhgt
            End If
            If nfloors = 3 Then
                STsheet.Cells(m, 14).Value = floorhgt
            End If
        End If
    m = m + 1
    End If
Next n

End Sub

Sub perimeters(HAsheet As Worksheet, STsheet As Worksheet, LastRowPNS As Long)
Dim m As Long
Dim n As Long
Dim nfloors As Integer

STsheet.Cells(1, 15).Value = "GroundFloorPerimeter"
STsheet.Cells(1, 16).Value = "FirstFloorPerimeter"
STsheet.Cells(1, 17).Value = "SecondFloorPerimeter"

m = 2
For n = 2 To LastRowPNS
    If IsEmpty(HAsheet.Cells(n, 4)) = False Then
        nfloors = STsheet.Cells(m, 6).Value
        If nfloors > 0 Then
            STsheet.Cells(m, 15).Value = HAsheet.Cells(n, 91).Value
        End If
        If nfloors > 1 Then
            STsheet.Cells(m, 16).Value = HAsheet.Cells(n, 91).Value
        End If
        If nfloors = 3 Then
            STsheet.Cells(m, 17).Value = HAsheet.Cells(n, 91).Value
        End If
    m = m + 1
    End If
Next n

End Sub

Sub SUSDEMin()

Dim STsheet As Worksheet
Set STsheet = Worksheets("SUSDEMtext")
Dim SIsheet As Worksheet
Set SIsheet = Worksheets("SUSDEMinput")
Dim LastRowST As Long
LastRowST = lRow(STsheet)
Dim STn As Long
Dim SIn As Long
SIn = 2
SIsheet.Rows(1).Value = STsheet.Rows(1).Value
For STn = 2 To LastRowST
    If IsEmpty(STsheet.Cells(STn, 3)) = False Then
        SIsheet.Rows(SIn).Value = STsheet.Rows(STn).Value
        SIn = SIn + 1
    End If
Next STn

End Sub

Function onlyDigits(s As String) As String
    ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next

    ' Then return the return string.                          '
    onlyDigits = retval
End Function

Function customAverageRooms(rng As Range, lower As Integer, upper As Integer, _
arch As Boolean, code As String, b As Integer, _
sheet1 As Worksheet, sheet2 As Worksheet, all As Boolean, columnHAnum As Integer, _
columnSTnum, n As Integer) As Double
'Allows average value of a range to be determined
Dim cell As Range
Dim total As Long
Dim count As Long
Dim t As Boolean
Dim q As Long

If arch = False And all = False Then 'Finding averages for neither archetypes nor all
    t = False
    For Each cell In rng
        If cell.Value >= lower And cell.Value <= upper And IsEmpty(cell) = False Then 'Removes values that are out of range
            total = total + cell.Value
            count = count + 1
            t = True
        End If
    Next cell
    If IsEmpty(sheet1.Cells(n, columnHAnum)) And b = 0 Then
        customAverageRooms = 20000
    ElseIf t = True Then
        customAverageRooms = total / count
    ElseIf t = False Then
        customAverageRooms = 20000
    End If
ElseIf arch = True And all = False Then 'Finding averages for archetypes
    q = 2
    For Each cell In rng
        If cell.Value >= lower And cell.Value <= upper And IsEmpty(cell) = False _
        And sheet2.Cells(q, 3).Value = code And sheet2.Cells(q, columnSTnum) <> 20000 Then
            total = total + cell.Value
            count = count + 1
        End If
    q = q + 1
    Next cell
    If count = 0 Then
        count = 1
    End If
    customAverageRooms = total / count
ElseIf arch = False And all = True Then 'Finding averages for all
    q = 2
    For Each cell In rng
        If cell.Value >= lower And cell.Value <= upper And IsEmpty(cell) = False _
        And sheet2.Cells(q, columnSTnum) <> 20000 Then
            total = total + cell.Value
            count = count + 1
        End If
    q = q + 1
    Next cell
    customAverageRooms = total / count
End If
End Function

Function blankspace(a As Integer, b As Integer) As Integer
Dim HAsheet As Worksheet
Set HAsheet = Worksheets("Haringey_All")

blankspace = 0
If IsEmpty(HAsheet.Cells(a + 1, b)) = True Then
    blankspace = 1
    If IsEmpty(HAsheet.Cells(a + 2, b)) = True Then
        blankspace = 2
        If IsEmpty(HAsheet.Cells(a + 3, b)) = True Then
            blankspace = 3
            If IsEmpty(HAsheet.Cells(a + 4, b)) = True Then
                blankspace = 4
                If IsEmpty(HAsheet.Cells(a + 5, b)) = True Then
                    blankspace = 5
                    If IsEmpty(HAsheet.Cells(a + 6, b)) = True Then
                        blankspace = 6
                    End If
                End If
            End If
        End If
    End If
End If

End Function

Function lRow(sheet As Worksheet) As Long
'Finds the last non-blank cell on a sheet/range.

Dim lCol As Long

    lRow = sheet.Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row

End Function
