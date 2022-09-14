Attribute VB_Name = "MultiCurrencyTranslator"
Option Explicit
Sub execute()
Call import
Call translate
End Sub
Sub import()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim sht As Worksheet, tool_sheet As Worksheet
    

'' Add toolsheet and depletions
Dim sheet_exist As Boolean
sheet_exist = False
For Each sht In ThisWorkbook.Sheets
    If sht.Name = "ToolSheet" Then
        sheet_exist = True
        ThisWorkbook.Sheets("ToolSheet").Cells.Clear
    End If
Next
If sheet_exist = False Then
    ThisWorkbook.Sheets.Add
    ActiveSheet.Name = "ToolSheet"
End If

'' delete other sheets
For Each sht In Sheets
    If sht.Name <> "ToolSheet" Then
        sht.Delete
    End If
Next

'' import data source
Dim last_row As Long, get_header As Boolean
Dim wb_source_name As String, file_name As Variant, f As Variant, wb_source As Workbook
Set tool_sheet = ThisWorkbook.Sheets("ToolSheet")

file_name = Application.GetOpenFilename("excels,*xls*", 1, "select file", , True)
If IsArray(file_name) Then
    For Each f In file_name
        Set wb_source = Workbooks.Open(f, 3)
            For Each sht In wb_source.Sheets
                Call copy_data(sht)
            Next
        wb_source.Close
    Next
Else
    End
End If

End Sub

Sub copy_data(sht As Worksheet)
Dim Criteria_cell As Range
sht.Activate
Set Criteria_cell = Cells.Find("Criteria", lookat:=xlPart)
If Criteria_cell Is Nothing Then
    GoTo Exit_sub
Else
    ThisWorkbook.Sheets.Add after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = sht.Name
    sht.Cells.Copy
    ThisWorkbook.Sheets(sht.Name).Cells.PasteSpecial xlPasteValuesAndNumberFormats
    
End If
Exit_sub:
End Sub

Sub translate(Optional base_currency As String = "EUR")
'' to decide multiply or divided by FX rate
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
Set wb_FX = Workbooks.Open(FX_path & wb_FX_name, 3)
FC_name = wb_FX.Sheets(1).Range(wb_FX.Sheets(1).Cells(1, 2), wb_FX.Sheets(1).Cells(1, wb_FX.Sheets(1).UsedRange.Columns.Count))
wb_FX.Close

For Each sht In ThisWorkbook.Sheets
    If sht.Name <> "ToolSheet" Then
        ThisWorkbook.Sheets("ToolSheet").Cells.Clear
        sht.Cells.Copy ThisWorkbook.Sheets("ToolSheet").Cells(1, 1)
        
        '' Prepare for vlookup the FX values on Date
        ThisWorkbook.Sheets("ToolSheet").Activate
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
        


'        base_currency_arr = base_currency_rng
        
        '' iteration of vlookup FX
        Dim i As Integer
        Dim new_wb As Workbook
        Set new_wb = Workbooks.Add
        '' copy the EUR version to new wb at first
        ThisWorkbook.Sheets("ToolSheet").Cells.Copy
        new_wb.Sheets(1).Cells(1, 1).PasteSpecial xlPasteValuesAndNumberFormats
        new_wb.Sheets(1).Name = "EUR" & sht.Name
        
        ThisWorkbook.Sheets("ToolSheet").Activate
        Set FX_rng = Range(Cells(2, last_col + 1), Cells(last_row, last_col + 1))
        '' start translating
        For i = 1 To UBound(FC_name, 2)
            Cells(1, last_col + 1).Value = FC_name(1, i)
            FX_rng.FormulaR1C1 = "=VLOOKUP(RC[" & Date_cell.Column - last_col - 1 & "],'" & FX_path & "[" & wb_FX_name & "]" & "FX'!C1:C3," & i + 1 & ",0)"
            base_currency_rng.Offset(0, base_currency_rng.Columns.Count + 1).FormulaArray = "=" & base_currency_rng.Address(ReferenceStyle:=xlR1C1) & "*" & FX_rng.Address(ReferenceStyle:=xlR1C1)
            currency_data_header.Offset(0, currency_data_header.Columns.Count + 1).FormulaArray = "=" & currency_data_header.Address(ReferenceStyle:=xlR1C1) & "&""" & Right(FC_name(1, i), Len(FC_name(1, i)) - Len("EUR")) & """"
            Cells.Copy
            Cells.PasteSpecial xlPasteValuesAndNumberFormats

            ActiveSheet.Copy after:=new_wb.Sheets(new_wb.Sheets.Count)
            '' must delete in new wb, because the original tool sheet still needs for later use, if delete in tool, base_currency rng would be nothing
            '' use union to delete because if delete base currency range at frist, columns reduce cause range address change
            Union(Range(base_currency_rng.Address), Range(FX_rng.Address)).EntireColumn.Delete
            
            new_wb.Sheets(new_wb.Sheets.Count).Name = Right(FC_name(1, i), Len(FC_name(1, i)) - Len("EUR")) & sht.Name
            ThisWorkbook.Sheets("ToolSheet").Activate
        Next
        new_wb.SaveAs new_currency_path & sht.Name & "MultiCurrency.xlsx"
        new_wb.Close
    End If
Next





End Sub
