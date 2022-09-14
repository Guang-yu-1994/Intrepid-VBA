Attribute VB_Name = "PriceUSDtoEUR"
Sub USD_to_EUR()
ThisWorkbook.Sheets("Summary").Activate
Dim FX_path As String, Date_cell As Range, wb_FX_name As String
Dim last_row As Long, last_col As Integer
Dim wb_FX As Workbook
Dim sht As Worksheet
Dim FC_name

'' new currency path
Dim new_currency_path As String
new_currency_path = "F:\Intrepid Spirits\Budget\MultiCurrency\"
'' how many currencies in the FX wb need to translate
FX_path = "F:\Intrepid Spirits\Budget\Budet Restructure\Replacement\"
wb_FX_name = "FX.xlsx"
Set wb_FX = GetObject(FX_path & wb_FX_name)
FC_name = "USD"
Dim USD_col As Integer
USD_col = wb_FX.Sheets(1).Cells.Find("USD", lookat:=xlPart).Column
wb_FX.Close


ThisWorkbook.Sheets("Tool").Cells.Clear
ActiveSheet.Cells.Copy ThisWorkbook.Sheets("Tool").Cells(1, 1)

'' Prepare for vlookup the FX values on Date
ThisWorkbook.Sheets("Tool").Activate
Set Date_cell = Cells.Find("Date")
last_row = ActiveSheet.UsedRange.Rows.Count
last_col = ActiveSheet.UsedRange.Columns.Count



'' set data range needs translation
Dim base_currency_rng As Range
Dim FX_rng As Range
Dim currency_data_header As Range
'        Dim base_currency_arr, fx_arr
If Date_cell.Offset(0, 1).Value Like "*Case*" Then
    Set base_currency_rng = Range(Date_cell.Offset(1, 2), Cells(last_row, last_col))
    Set currency_data_header = Range(Date_cell.Offset(0, 2), Cells(1, last_col))
Else
    Set base_currency_rng = Range(Date_cell.Offset(1, 1), Cells(last_row, last_col))
    Set currency_data_header = Range(Date_cell.Offset(0, 1), Cells(1, last_col))
End If




'' iteration of vlookup FX

'' copy the EUR version to new wb at first
ThisWorkbook.Sheets("Tool").Activate
Set FX_rng = Range(Cells(2, last_col + 1), Cells(last_row, last_col + 1))
'' start translating

Cells(1, last_col + 1).Value = "FX" & FC_name
FX_rng.FormulaR1C1 = "=VLOOKUP(RC[" & Date_cell.Column - last_col - 1 & "],'" & FX_path & "[" & wb_FX_name & "]" & "FX'!C1:C3," & USD_col & ",0)"
base_currency_rng.Offset(0, base_currency_rng.Columns.Count + 1).FormulaArray = "=" & base_currency_rng.Address(ReferenceStyle:=xlR1C1) & "/" & FX_rng.Address(ReferenceStyle:=xlR1C1)
currency_data_header.Offset(0, currency_data_header.Columns.Count + 1).Value = "PriceEUR"
Cells.Copy
Cells.PasteSpecial xlPasteValuesAndNumberFormats
ActiveSheet.Copy
'' must delete in new wb, because the original tool sheet still needs for later use, if delete in tool, base_currency rng would be nothing
'' use union to delete because if delete base currency range at frist, columns reduce cause range address change
Union(Range(base_currency_rng.Address), Range(FX_rng.Address)).EntireColumn.Delete
ActiveSheet.Name = "Summary EUR"
ActiveWorkbook.SaveAs "F:\Intrepid Spirits\Budget\Budet Restructure\ProductDetailStructure\PriceDataStructured\PriceData Americas (EUR).xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
ActiveWorkbook.Close

ThisWorkbook.Sheets("Tool").Cells.Clear
End Sub
