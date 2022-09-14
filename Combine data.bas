Attribute VB_Name = "Combine"
Sub execute()
Call CombineData
Call data_stack
Call prepare_criteria
End Sub

Sub CombineData()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
'On Error GoTo no_Summary
Sheets("Summary").Delete
'no_Summary:
Sheets.Add before:=Sheets(1)
ActiveSheet.Name = "Summary"

Dim get_header As Boolean, sht As Worksheet
Dim last_row As Integer, last_col As Integer, start_cell As Range, copy_rng As Range
Dim Brand As Range
get_header = True
For Each sht In Sheets
    sht.Activate
    If sht.Name <> "Summary" And sht.Name <> "SKU" Then
        Set start_cell = Cells(1, 1)
        Set Brand = Cells.Find("Brand")
        If Not Brand Is Nothing Then
            last_row = Cells(Rows.Count, Brand.Column).End(xlUp).Row
            last_col = Cells(start_cell.Row, Columns.Count).End(xlToLeft).Column
            '' get header
            If get_header Then
                Rows(1).Copy
                Sheets("Summary").Cells(1, 1).PasteSpecial Paste:=xlFormulas
                get_header = False
            End If
            '' copy data
            Set copy_rng = Range(start_cell.Offset(1, 0), Cells(last_row, last_col))
            copy_rng.Copy
            Sheets("Summary").Cells(Sheets("Summary").UsedRange.Rows.Count + 1, 1).PasteSpecial Paste:=xlFormulas
        End If
    End If
Next

Sheets("Summary").Activate
'' rearrange
Dim Country As Range
Set Country = Cells.Find("Country")
Range(Cells(1, 1), Cells(1, Country.Column - 1)).EntireColumn.Delete

'Cells.Find("ABV").EntireColumn.Delete

'Dim Jan As Range
'Set Jan = Cells.Find("Jan")
'Jan.Offset(0, -1).Value = "Criteria"
'last_row = Cells(Rows.Count, Jan.Column).End(xlUp).Row
'Range(Jan.Offset(1, -1), Cells(last_row, Jan.Column - 1)).FormulaR1C1 = "=substitute(RC[-7]&RC[-6]&RC[-5]&RC[-4]&RC[-3],"" "", """")"
End Sub

Sub data_stack()
ThisWorkbook.Sheets("Summary").Activate
Dim arr, brr()
Dim a_row As Integer, a_col As Integer, b_row As Integer, b_col As Integer
Dim col_header As Integer, i As Integer


' until column Jan are column header
Dim Jan As Range
Set Jan = Cells.Find("ABV").Offset(0, 1)
col_header = Jan.Column - 1

arr = Cells(1, 1).CurrentRegion
ReDim brr(1 To UBound(arr, 1) * UBound(arr, 2), 1 To col_header + 2)

b_row = 1
'' set header name
For b_col = 1 To col_header
    brr(1, b_col) = arr(1, b_col)
Next
brr(1, col_header + 1) = "Date"
brr(1, col_header + 2) = "Price"

b_row = 1
'' start stacking data
For a_row = 2 To UBound(arr, 1) ' iteration of row
    For a_col = col_header + 1 To UBound(arr, 2) ' iteratiion of column
        b_row = b_row + 1
        For i = 1 To col_header
            brr(b_row, i) = arr(a_row, i) ' column 1 to 8 remains unchanged
        Next
        brr(b_row, col_header + 1) = arr(1, a_col) ' Date_rng data stack
        brr(b_row, col_header + 2) = arr(a_row, a_col) ' Case data stack
    Next
Next

ActiveSheet.Cells.Clear
ActiveSheet.Cells(1, 1).Resize(b_row, col_header + 2) = brr
End Sub
Sub prepare_criteria()
ThisWorkbook.Sheets("Summary").Activate
Dim Date_rng As Range
Dim last_row As Integer
Set Date_rng = Cells.Find("Date")
Date_rng.EntireColumn.NumberFormat = "mmm-yy"
last_row = Cells(Rows.Count, Date_rng.Column).End(xlUp).Row
'Date_rng.EntireColumn.Resize(, 2).Insert
'Date_rng.Offset(0, -1).Value = "Year"
'Range(Date_rng.Offset(1, -1), Cells(last_row, Date_rng.Column - 1)).Value = Year(Now())

Date_rng.EntireColumn.Insert
Date_rng.Offset(0, -1).Value = "Criteria"
Range(Date_rng.Offset(1, -1), Cells(last_row, Date_rng.Column - 1)).FormulaR1C1 = "=SUBSTITUTE(TEXT(RC[1],""MMM/YY"")&RC[-8]&RC[-7]&RC[-6]&RC[-5]&RC[-4],"" "","""")"



End Sub

Sub trim_rng()
''Trim every cell
Dim sht As Worksheet, rng As Range
For Each sht In Sheets
    For Each rng In sht.UsedRange
        rng.Value = Trim(rng.Value)
    Next
Next
End Sub

