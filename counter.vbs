Option Explicit

Sub counter()

'Primary workbook
Dim wkb1 As Excel.Workbook
Dim wks1 As Excel.Worksheet
Set wkb1 = Excel.Workbooks("Pricing for Week 09 - 2020-2.xlsm")
Set wks1 = wkb1.Worksheets("buying worksheet")

'Second workbook
Dim wkb2 As Excel.Workbook
Dim wks2 As Excel.Worksheet
Set wkb2 = Excel.Workbooks("temp.xlsx")
Set wks2 = wkb2.Worksheets("Sheet1")

Dim rowRange1 As Range
Dim lastRow1 As Long

Dim rowRange2 As Range
Dim lastRow2 As Long

'Primary workbook
lastRow1 = wks1.Cells(wks1.Rows.count, "C").End(xlUp).Row
'Secondary workbook
lastRow2 = wks2.Cells(wks2.Rows.count, "A").End(xlUp).Row

'Set range for primary & secondary rows
Set rowRange1 = wks1.Range("C1:C" & lastRow1)
Set rowRange2 = wks2.Range("A1:A" & lastRow2)

'primary & secondary empty values
Dim myValue1
Dim myValue2

Dim rrow1
Dim rrow2
Dim wipe

Dim count1 As Integer
Dim count2 As Integer
Dim clean1 As Integer

count1 = 0
clean1 = 0

'Set interior color for all cells in range to blank
For Each wipe In rowRange1
    clean1 = clean1 + 1
    wks1.Range("B" & clean1 & ":T" & clean1).Interior.ColorIndex = 0
    wks1.Range("W" & clean1 & ":AF" & clean1).Interior.ColorIndex = 0
Next wipe

'Set interior color for cells in range whos pm_id matches secondary workbook column A
For Each rrow1 In rowRange1
    count1 = count1 + 1
    Set myValue1 = wks1.Range("C" & count1)
    'Debug.Print (myValue1)
    count2 = 0
    For Each rrow2 In rowRange2
        count2 = count2 + 1
        'Debug.Print (count2)
        Set myValue2 = wks2.Range("A" & count2)
        'Debug.Print (myValue2)
        
        If myValue1.Value = myValue2.Value Then
            Debug.Print (myValue1)
            Debug.Print ("B" & count1 & ":T" & count1)
            wks1.Range("B" & count1 & ":T" & count1).Interior.ColorIndex = 3
            wks1.Range("W" & count1 & ":AF" & count1).Interior.ColorIndex = 3
        End If
        
    Next rrow2
Next rrow1
    
End Sub
