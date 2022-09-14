Attribute VB_Name = "JPCustomerAnalyst"
Option Explicit
Public header As Variant
Sub execute()
Call import_plan
Call process_data

End Sub
Sub import_plan()
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

sheet_exist = False
For Each sht In ThisWorkbook.Sheets
    If sht.Name = "Depletions" Then
        sheet_exist = True
        Sheets("Depletions").Cells.Clear
    End If
Next
If sheet_exist = False Then
    ThisWorkbook.Sheets.Add
    ActiveSheet.Name = "Depletions"
End If

For Each sht In ThisWorkbook.Sheets
    If sht.Name <> "ToolSheet" And sht.Name <> "Depletions" Then
        sht.Delete
    End If
Next



'' import data source
Dim last_row As Long, get_header As Boolean
Dim file_name As Variant, f As Variant, dep_source As Workbook
Set tool_sheet = ThisWorkbook.Sheets("ToolSheet")

file_name = Application.GetOpenFilename("excels,*xls*", 1, "select file", , True)
If IsArray(file_name) Then
    For Each f In file_name
        Set dep_source = Workbooks.Open(f, 3)
        For Each sht In dep_source.Sheets
            If Len(sht.Name) < 6 Then
                sht.Cells.Copy
                ThisWorkbook.Sheets.Add after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = Left(dep_source.Name, 4) & sht.Name
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Cells(1, 1).PasteSpecial xlPasteValuesAndNumberFormats
            End If
        Next
        dep_source.Close
    Next
Else
    End
End If
End Sub

Sub process_data()
Dim sht As Worksheet
For Each sht In ThisWorkbook.Sheets
    If sht.Name <> "ToolSheet" And sht.Name <> "Depletions" Then
        sht.Activate
        Call replace_SKU_name
        Call copy_data
    End If
Next

Call rearrange_data


End Sub
Sub copy_data()
Dim start_cells As Range, start_cell As Range
Dim ML As Integer, SKU_variant As String, t_last_row As Long, d_last_row As Long
Dim Case_cell As Range, copy_data_rng As Range
Dim get_header As Boolean
Dim t_start_cell As Range
get_header = True

Set start_cells = get_start_cell()
For Each start_cell In start_cells

    ML = get_ml(start_cell.Offset(-1, 0))
    SKU_variant = extract_variant(start_cell.Offset(-1, 0))
    start_cell.CurrentRegion.Copy Sheets("ToolSheet").Cells(1, 1)
    Sheets("ToolSheet").Activate
    Set t_start_cell = Cells.Find("Sales Figures")
    If t_start_cell.Column <> 1 Then
        Range(Cells(1, 1), Cells(1, t_start_cell.Column - 1)).EntireColumn.Delete
    End If
    '' delete the first two rows to leave the header as the first row in tool sheet
    Rows("1:2").Delete
    t_last_row = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(1, 1).EntireColumn.Insert
    Cells(1, 1).Value = "Variant"
    Set Case_cell = Cells.Find("case")
    Case_cell.Offset(0, 1).Value = "ML"
    
    '' then store the header as public var
    header = Range(Cells(1, 1), Cells(1, 6))
    ' fill in the ml and variant info

    Range(Case_cell.Offset(1, 1), Cells(t_last_row, Case_cell.Column + 1)).Value = ML
    Range(Cells(2, 1), Cells(t_last_row, 1)).Value = SKU_variant
    '' copy the data (no header, header could be added later)
    Range(Cells(2, 1), Cells(t_last_row, Case_cell.Column + 1)).Copy
    Sheets("Depletions").Activate
    d_last_row = ActiveSheet.UsedRange.Rows.Count
    Cells(d_last_row + 1, 1).PasteSpecial xlPasteValuesAndNumberFormats
    Sheets("ToolSheet").Cells.Clear
Next


End Sub
Sub rearrange_data()
Dim sht As Worksheet
For Each sht In ThisWorkbook.Sheets
    If sht.Name <> "ToolSheet" And sht.Name <> "Depletions" Then
        sht.Delete
    End If
Next

Dim last_row As Long, last_col As Integer
'insert the header to sheets Depletions
Sheets("Depletions").Activate
last_col = Cells(2, Columns.Count).End(xlToLeft).Column
'Rows(1).EntireRow.Insert
Cells(1, 1).Resize(, last_col) = header


'' delete the subtotals
Dim Sold_to As Range
Set Sold_to = Cells.Find("Sold to", lookat:=xlPart)
ActiveSheet.UsedRange.AutoFilter Field:=Sold_to.Column, Criteria1:="="
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells(1, 1).EntireRow.Insert
Cells(1, 1).Resize(, last_col) = header
Cells.Replace "date", "Date"

'' Calculate 9L
last_row = ActiveSheet.UsedRange.Rows.Count
Cells(1, last_col + 1).Value = "9LCase"
Range(Cells(2, last_col + 1), Cells(last_row, last_col + 1)).FormulaR1C1 = "=RC[-1]*RC[-2]/9000"

'' extract customer Column
Set Sold_to = Cells.Find("Sold to", lookat:=xlPart)
Sold_to.EntireColumn.Insert
Sold_to.Offset(0, -1).Value = "Customer"
Range(Sold_to.Offset(1, -1), Cells(last_row, Sold_to.Column - 1)).FormulaR1C1 = "=extract_customer(RC[1])"

'' set date formate
Dim Date_cell As Range, Date_rng As Range, rng As Range
Dim a, i As Long, arr, brr
Set Date_cell = Cells.Find("Date")
Set Date_rng = Range(Date_cell.Offset(1, 0), Cells(last_row, Date_cell.Column))
Date_rng.NumberFormat = "dd/mm/yyyy;@"
arr = Date_rng
ReDim brr(1 To UBound(arr, 1), 1 To UBound(arr, 2))
For i = LBound(arr, 1) To UBound(arr, 1)
    brr(i, 1) = CDate(arr(i, 1))
Next
Date_rng = brr

'' autofit
Cells.EntireColumn.AutoFit

'' paste as value
Cells.Copy
Cells.PasteSpecial xlPasteValuesAndNumberFormats
End Sub

Sub replace_SKU_name()
Dim replacement_path As String, repl_wb As Workbook
Dim replacements, i As Integer
replacement_path = "F:\Intrepid Spirits\Arnout\LanguageTranslation\replacement for JP Customer.xlsx"
Set repl_wb = GetObject(replacement_path)
replacements = repl_wb.Sheets(1).Cells(1, 1).CurrentRegion
For i = LBound(replacements, 1) To UBound(replacements, 1)
    Cells.Replace replacements(i, 1), replacements(i, 2), xlWhole
Next
repl_wb.Close
End Sub
Function get_start_cell()
Dim start_cell As Range, start_cells As Range, rng As Range
Dim goal_string As String
Dim origin_col As Integer
goal_string = "Sales Figures"
Set start_cells = Cells.Find(goal_string, after:=Cells(1, 1), searchdirection:=xlNext, MatchCase:=False, LookIn:=xlValues, SearchOrder:=xlByColumns)
Set start_cell = Cells.FindNext(after:=start_cells)

' if only one start_cell
If start_cells.Address = start_cell.Address Then
    Set get_start_cell = start_cells
Else
    origin_col = start_cells.Column
    Do While start_cell.Column <> origin_col
        Set start_cells = Union(start_cells, start_cell)
        Set start_cell = Rows(start_cell.Row).FindNext(after:=start_cell)
    Loop
    Set get_start_cell = start_cells
End If
End Function


Function get_ml(rng As Range)
Dim reg As Object
Set reg = CreateObject("vbscript.regexp")
With reg
    .Global = True
    .Pattern = "\d+"
    On Error GoTo error_exit:
    get_ml = CInt(.execute(Trim(rng.Value))(0))
    Exit Function
error_exit:
    get_ml = 0
End With
End Function
Function extract_variant(rng As Range)
Dim i As Integer
Dim r As Integer, c As Integer
Dim r_result(), r_rng
Dim reg As Object
ReDim r_result(1 To rng.Rows.Count, 1 To rng.Columns.Count)
Set reg = CreateObject("vbscript.regexp")
r_rng = rng ' store rng in array
If IsArray(r_rng) Then
    With reg
        .Global = True
        .Pattern = "\d+\s*[mc]l"
        For r = 1 To UBound(r_rng, 1)
            For c = 1 To UBound(r_rng, 2)
                r_result(r, c) = Trim(.Replace(Trim(r_rng(r, c)), ""))
            Next
        Next
    End With
    extract_variant = r_result
Else
    With reg
        .Global = True
        .Pattern = "\d+\s*[mc]l"
        extract_variant = Trim(.Replace(Trim(r_rng), ""))
    End With
End If
End Function
Function extract_customer(rng As Range)
Dim i As Integer
Dim r As Integer, c As Integer
Dim r_result(), r_rng
Dim reg As Object
Dim mats
ReDim r_result(1 To rng.Rows.Count, 1 To rng.Columns.Count)
Set reg = CreateObject("vbscript.regexp")
r_rng = rng ' store rng in array

With reg
On Error GoTo another_pattern:
    .Global = True
    .Pattern = "(.*:.*?)(\d+[Co., Ltd]?)(.+?(?=\s*Distribution|\s*Co|\s*Log|r\d+|$))(\d+)?"
     extract_customer = Trim(.execute(Trim(r_rng))(0).SubMatches(2))
     If Len(extract_customer) < 2 Then GoTo another_pattern
     Exit Function
End With
another_pattern:
With reg
    .Global = True
    .Pattern = "(.*:.*?)(\d+[Co., Ltd]?|\s*)(.+?(?=\s*Distribution|\s*Co|\s*Log|r\d+|$))(\d+)?"
     extract_customer = Trim(.execute(Trim(r_rng))(0).SubMatches(2))
End With

End Function
