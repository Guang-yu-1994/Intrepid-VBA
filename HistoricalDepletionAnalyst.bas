Attribute VB_Name = "HistoricalDepletionAnalyst"
Option Explicit
Sub import_ledger()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim sht As Worksheet, tool_sheet As Worksheet
    

'' Add toolsheet and depletions
Dim sheet_exist As Boolean
sheet_exist = False
For Each sht In ThisWorkbook.Sheets
    If sht.Name = "ToolSheet" Then
        sheet_exist = True
        Sheets("ToolSheet").Cells.Clear
    End If
Next
If sheet_exist = False Then
    ThisWorkbook.Sheets.Add
    ActiveSheet.Name = "ToolSheet"
End If


For Each sht In ThisWorkbook.Sheets
    If sht.Name <> "ToolSheet" And sht.Name <> "Assumptions" Then
        sht.Delete
    End If
Next



'' import data source
Dim last_row As Long, get_header As Boolean
Dim ledger_source_name As String, file_name As Variant, f As Variant, ledger_source As Workbook
Set tool_sheet = ThisWorkbook.Sheets("ToolSheet")

file_name = Application.GetOpenFilename("excels,*xls*", 1, "select file", , True)
If IsArray(file_name) Then
    For Each f In file_name
        Set ledger_source = Workbooks.Open(f)
        ThisWorkbook.Sheets("ToolSheet").Cells.Clear
        Call clean_up_sales_ledger(ledger_source)
        
        ledger_source.Close
    Next
Else
    End
End If
End Sub
Sub clean_up_sales_ledger(ledger_source As Workbook)


ledger_source.Sheets(1).Activate
'' find the ledger time (per Month)
Dim ledger_time As String
ledger_time = Cells.Find("Sales Per Region And Brand").Offset(1, 0).Value

Dim last_row As Integer, last_col As Integer
last_col = ActiveSheet.UsedRange.Columns.Count

Dim FinancialRow As Range, Brand_cell As Range, Amount_cell As Range
Set FinancialRow = Cells.Find("Financial Row")
Set Brand_cell = Cells.Find("Brand", lookat:=xlWhole)
Set Amount_cell = Cells.Find("Amount", lookat:=xlWhole)


'filter out unnecessary rows in the original sales ledger
Range(FinancialRow, Cells(ActiveSheet.UsedRange.Rows.Count, last_col)).AutoFilter FinancialRow.Column, "<>"
last_row = Cells(Rows.Count, Amount_cell.Column).End(xlUp).Row ' update last row after filter
Range(FinancialRow.Offset(1, 0), Cells(last_row, Amount_cell.Column)).EntireRow.Delete
Range(FinancialRow, Cells(ActiveSheet.UsedRange.Rows.Count, last_col)).AutoFilter

'' step 1 remove internal records
Dim Name_cell As Range, company_name_path As String, company_name_wb As Workbook
Set Name_cell = Cells.Find("Name")
company_name_path = "F:\Intrepid Spirits\Consolidation\HistoricalDepletionAnalyst\CompanyName.xlsx"
Set company_name_wb = GetObject(company_name_path)
Application.CutCopyMode = False
Range(FinancialRow, Cells(ActiveSheet.UsedRange.Rows.Count, last_col)).AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=company_name_wb.Sheets(1).UsedRange, Unique:=False
last_row = Cells(Rows.Count, Amount_cell.Column).End(xlUp).Row ' update last row after filter
Range(FinancialRow.Offset(1, 0), Cells(last_row, Amount_cell.Column)).EntireRow.Delete
Application.CutCopyMode = True
Range(FinancialRow, Cells(ActiveSheet.UsedRange.Rows.Count, last_col)).AutoFilter
company_name_wb.Close



'' step 2 minus number and positive number with the same abs value should be the same brand

' filter the blank brand rows
Range(FinancialRow, Cells(ActiveSheet.UsedRange.Rows.Count, last_col)).AutoFilter Brand_cell.Column, "="
last_row = Cells(Rows.Count, Amount_cell.Column).End(xlUp).Row ' update last row after filter


'copy the blank brand rows to thiswork book tool sheet, then delete them in the original ledger
Range(FinancialRow, Cells(last_row, last_col)).Copy ThisWorkbook.Sheets("ToolSheet").Cells(1, 1)
Range(FinancialRow.Offset(1, 0), Cells(last_row, Amount_cell.Column)).EntireRow.Delete
Range(FinancialRow, Cells(last_row, last_col)).AutoFilter
last_row = Cells(Rows.Count, Amount_cell.Column).End(xlUp).Row ' update last row after delete

' and "-" in front of each amount
ThisWorkbook.Sheets("ToolSheet").Activate
Dim t_Amount_cell As Range, t_Amount_rng As Range, t_Amount_arr, i
Dim t_last_row As Integer
Set t_Amount_cell = Cells.Find("Amount", lookat:=xlWhole)
t_last_row = Cells(Rows.Count, t_Amount_cell.Column).End(xlUp).Row
Set t_Amount_rng = Range(t_Amount_cell.Offset(1, 0), Cells(t_last_row, t_Amount_cell.Column))
t_Amount_arr = t_Amount_rng
' use array to accelerate
For i = 1 To UBound(t_Amount_arr, 1)
    t_Amount_arr(i, 1) = -1 * t_Amount_arr(i, 1)
Next
t_Amount_rng = t_Amount_arr

t_Amount_rng.Offset(0, 2).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],'[" & ledger_source.Name & "]" & ledger_source.Sheets(1).Name & "'!C8:C10,3,0),"""")"
Cells.Copy
Cells.PasteSpecial xlPasteValues
' recover the number for vlookup temp use
For i = 1 To UBound(t_Amount_arr, 1)
    t_Amount_arr(i, 1) = -1 * t_Amount_arr(i, 1)
Next
t_Amount_rng = t_Amount_arr

t_last_row = Cells(Rows.Count, t_Amount_cell.Column).End(xlUp).Row - 1 ' later would delete row 1 so minus 1
ThisWorkbook.Sheets("ToolSheet").Rows(1).Delete
Range(Cells(1, 1), Cells(t_last_row, 1)).EntireRow.Copy ledger_source.Sheets(1).Cells(last_row + 1, 1)
Cells.Clear

'' step 3  EMEA sample has the region of other EMEA,APAC as other APAC, America as other Americas
Dim region_to_replace, r
region_to_replace = Array("EMEA", "APAC", "America")
For Each r In region_to_replace
    ledger_source.Sheets(1).Activate
    Range(FinancialRow, Cells(ActiveSheet.UsedRange.Rows.Count, last_col)).AutoFilter Name_cell.Column, "=*" & r & "*"
    last_row = Cells(Rows.Count, Amount_cell.Column).End(xlUp).Row ' update last row after filter
    ' copy to ToolSheet
    Range(FinancialRow, Cells(last_row, last_col)).Copy ThisWorkbook.Sheets("ToolSheet").Cells(1, 1)
    Range(FinancialRow.Offset(1, 0), Cells(last_row, Amount_cell.Column)).EntireRow.Delete
    Range(FinancialRow, Cells(last_row, last_col)).AutoFilter
    last_row = Cells(Rows.Count, Amount_cell.Column).End(xlUp).Row ' update last row after delete
    
    ThisWorkbook.Sheets("ToolSheet").Activate
    Dim t_Region_cell As Range
    Set t_Region_cell = Cells.Find("Region (GL)")
    t_last_row = Cells(Rows.Count, t_Region_cell.Column).End(xlUp).Row - 1
    Range(t_Region_cell.Offset(1, 0), Cells(t_last_row + 1, t_Region_cell.Column)).Value = "other " & r
    ThisWorkbook.Sheets("ToolSheet").Rows(1).Delete
    Range(Cells(1, 1), Cells(t_last_row, 1)).EntireRow.Copy ledger_source.Sheets(1).Cells(last_row + 1, 1)
    Cells.Clear
Next

'' step 4 replace Europe as other EMEA in Sales Ledger
ledger_source.Sheets(1).Activate
Cells.Replace "Europe", "other EMEA"
Cells.Replace "APAC", "other APAC"
Cells.Replace "Americas", "other Americas"
Cells.Replace "other America", "other Americas"
Cells.Replace "United States", "USA"
Cells.Replace "United Kingdom", "UK"
' sort by Date
Dim Date_cell As Range
Set Date_cell = Cells.Find("Date")

Range(FinancialRow, Cells(ActiveSheet.UsedRange.Rows.Count, last_col)).AutoFilter
ActiveSheet.AutoFilter.Sort.SortFields.Clear
ActiveSheet.AutoFilter.Sort.SortFields. _
        Add2 Key:=Date_cell, SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
        
'' step 5 add region
Call add_region

'' save as xlsx to use pivot table and power pivot
ledger_source.SaveAs ledger_source.Path & "\" & Split(ledger_source.Name, ".")(0) & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
End Sub


