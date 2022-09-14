Attribute VB_Name = "GlobalDepletionAnalyst"
Option Explicit
'Public start_time, end_time
Sub execute()
Dim start_time, end_time

Call import_plan
Call process_data
'Call sum_plan
'Call Export
'end_time = Timer
'MsgBox end_time - start_time
End Sub
Sub import_plan()
'start_time = Timer
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim sht As Worksheet, tool_sheet As Worksheet
    

'' Add toolsheet and depletions
Dim sheet_exist As Boolean
sheet_exist = False
For Each sht In ThisWorkbook.Sheets
    If sht.Name = "ToolSheet" Then sheet_exist = True
Next
If sheet_exist = False Then
    ThisWorkbook.Sheets.Add
    ActiveSheet.Name = "ToolSheet"
End If

sheet_exist = False
For Each sht In ThisWorkbook.Sheets
    If sht.Name = "Depletions" Then sheet_exist = True
Next
If sheet_exist = False Then
    ThisWorkbook.Sheets.Add
    ActiveSheet.Name = "Depletions"
End If

For Each sht In ThisWorkbook.Sheets
    If sht.Name <> "ToolSheet" Or sht.Name = "Depletions" Then
        sht.Delete
    End If
Next



'' import data source
Dim last_row As Integer, get_header As Boolean
Dim plan_source_name As String, file_name As Variant, f As Variant, plan_source As Workbook
Set tool_sheet = ThisWorkbook.Sheets("ToolSheet")

file_name = Application.GetOpenFilename("excels,*xls*", 1, "select file", , True)
If IsArray(file_name) Then
    For Each f In file_name
        Set plan_source = Workbooks.Open(f)
        plan_source_name = Trim(Split(plan_source.Name, "-")(1))
        For Each sht In plan_source.Sheets
            If UCase(sht.Name) Like UCase("*Plan*History*") Then
                sht.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = plan_source_name & "HistoryAndPlan"
            ElseIf UCase(sht.Name) Like UCase("*Actual*") Then
                sht.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = plan_source_name & "Actual"
            ElseIf UCase(sht.Name) Like UCase("*LE*") Then
                sht.Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = plan_source_name & "LE"
            End If
        Next
        plan_source.Close
    Next
Else
    End
End If

Dim product_detail As String, detail_wb As Workbook, rng As Range
product_detail = Dir(ThisWorkbook.Path & "\ProductDetail\*.xls*")
Do While product_detail <> ""
    Set detail_wb = Workbooks.Open(ThisWorkbook.Path & "\ProductDetail\" & product_detail)
    detail_wb.Sheets(1).Copy after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    For Each rng In Range(Cells(1, 1), Cells(1, 1).Offset(0, 3).End(xlDown))
        rng.Value = remove_blanks(rng)
    Next
    detail_wb.Close
    product_detail = Dir
Loop
End Sub
Sub process_data()
Dim sht As Worksheet

For Each sht In ThisWorkbook.Sheets
    If UCase(sht.Name) Like UCase("*HistoryAndPlan*") Then
        Call copy_data(sht.Name)
    End If
Next

For Each sht In ThisWorkbook.Sheets
    If UCase(sht.Name) Like UCase("*[a-z]*Actual*") Then
        Call copy_data(sht.Name)
    End If
Next

For Each sht In ThisWorkbook.Sheets
    If UCase(sht.Name) Like UCase("*LE*") Then
        Call copy_data(sht.Name)
    End If
Next
Call sum_depletion
Call data_stack
'
'' vlookup product details
Dim sheet_names, sheet_name
sheet_names = Array("Price", "Cost", "AMP")
For Each sheet_name In sheet_names
    Call map_product_detail(sheet_name)
Next
Call get_profit


'' calculate summary
Call calculate_summary
Call add_region
Call add_date
Call nine_l
End Sub
Sub copy_data(ByVal sheet_name As String)
Sheets(sheet_name).Activate
Dim start_cell As Range, start_cells As Range, to_sheet As String
Set start_cells = get_start_cell(sheet_name)
For Each start_cell In start_cells
    If Not start_cell.Offset(0, 1).Value Like "A&P*" Then
        start_cell.EntireColumn.Find("Market").CurrentRegion.EntireColumn.Copy
        Sheets("ToolSheet").Cells(1, 1).PasteSpecial Paste:=xlPasteValues
        
        '' create new spreadsheet for storing the data
        '' the new sheet name depends on budget or actual or LE
        
        to_sheet = get_new_sheet_name(sheet_name, start_cell.Offset(0, 1).Value)
        
        
        
        Dim sheet_exist As Boolean, sht As Worksheet
        sheet_exist = False
        For Each sht In ThisWorkbook.Sheets
            If sht.Name = to_sheet Then
                sheet_exist = True
            End If
        Next
        If sheet_exist = False Then
            ThisWorkbook.Sheets.Add
            ActiveSheet.Name = to_sheet
        End If
        
        Call filter_data
        Call prepare_criteria
        
        '' Add year column to the spreadsheet
        ThisWorkbook.Sheets("ToolSheet").Activate
        Columns("A:A").Insert
        Cells(1, 1).Value = "Year"
        Range(Cells(2, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 1)).Value = extract_year(to_sheet)
        
        '' add depletion catagory to the spreadsheet
        Columns("A:A").Insert
        Cells(1, 1).Value = "Catagory"
        Range(Cells(2, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 1)).Value = extract_catagory(to_sheet)
        
        
        Dim last_row As Integer, last_row_tool As Integer
        last_row_tool = ThisWorkbook.Sheets("ToolSheet").Cells(Rows.Count, 1).End(xlUp).Row
        last_row = ThisWorkbook.Sheets(to_sheet).Cells(Rows.Count, 1).End(xlUp).Row
        If ThisWorkbook.Sheets(to_sheet).Cells(1, 1) = "" Then
            ThisWorkbook.Sheets("ToolSheet").UsedRange.Copy ThisWorkbook.Sheets(to_sheet).Cells(last_row, 1)
        Else
            ThisWorkbook.Sheets("ToolSheet").UsedRange.Rows("2:" & last_row_tool).Copy ThisWorkbook.Sheets(to_sheet).Cells(last_row + 1, 1)
        End If
        Sheets("ToolSheet").Cells.Clear
    End If
    
Next


End Sub
Sub filter_data()
Dim tool_sheet As Worksheet, Cases_cell As Range
Set tool_sheet = ThisWorkbook.Sheets("ToolSheet")
tool_sheet.Activate
Set Cases_cell = tool_sheet.Cells.Find("Cases", lookat:=xlWhole)
Cases_cell.Value = "TotalCases"
Range(Cases_cell.Offset(0, 1), Cases_cell.End(xlToRight)).EntireColumn.Delete

'' add country name
Columns("A:A").Insert
Dim mkt_cell As Range, f_mkt_cell As Range, fill_until_row As Integer
Set mkt_cell = Cells.Find("Market", after:=Cells(1, 1), SearchDirection:=xlPrevious, lookat:=xlWhole)
Set f_mkt_cell = mkt_cell
fill_until_row = Cells(Rows.Count, 2).End(xlUp).Row
Do While mkt_cell.Row <= f_mkt_cell.Row
    mkt_cell.Offset(0, -1).Value = mkt_cell.Offset(0, 1).Value
    mkt_cell.Offset(0, -1).AutoFill Range(mkt_cell.Offset(0, -1), Cells(fill_until_row, mkt_cell.Offset(0, -1).Column))
    Set f_mkt_cell = mkt_cell
    fill_until_row = mkt_cell.Row - 1
    Set mkt_cell = Cells.FindPrevious(after:=f_mkt_cell)
Loop

'' delete market cell rows and total physical case row
Dim last_cell As Range, filter_rng As Range
Dim last_row As Integer, last_col As Integer
last_row = Cells(Rows.Count, mkt_cell.Column).End(xlUp).Row
last_col = Cases_cell.Column
Set filter_rng = Range(Cells(1, 1), Cells(last_row, last_col))

Set mkt_cell = Cells.Find("Market", after:=Cells(1, 1), SearchDirection:=xlNext, lookat:=xlWhole)
filter_rng.AutoFilter Field:=mkt_cell.Column, Criteria1:="=*Market*", Operator:=xlOr, Criteria2:="=*Total*"
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete

'' Delete 0 cases columns
filter_rng.AutoFilter Field:=last_col, Criteria1:="<0.5", Criteria2:="=", Operator:=xlOr
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells.Find("Brand").Offset(0, -1).Value = "Country"


'' leave only one header
Dim Expression_cell As Range
'' delete Expression cell rows except for the first one (header)
Set Expression_cell = Cells.Find("Expression", after:=Cells(1, 1), SearchDirection:=xlNext)
Set filter_rng = tool_sheet.UsedRange
filter_rng.AutoFilter Field:=Expression_cell.Column, Criteria1:="Expression"
'' add temp sheet
ThisWorkbook.Sheets.Add after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = "Temp"
tool_sheet.Activate
filter_rng.Rows(1).Copy Sheets("Temp").Cells(1, 1)
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Rows(1).Insert
Sheets("Temp").Rows(1).Copy Cells(1, 1)
Sheets("Temp").Cells.Clear

'' delete the rows for subtotal each region
Dim Region_cell As Range, subtotal As String
Set Region_cell = Cells.Find("Country")
Set filter_rng = tool_sheet.UsedRange
subtotal = Cells(Rows.Count, Region_cell.Column).End(xlUp).Value
filter_rng.AutoFilter Field:=Region_cell.Column, Criteria1:="=" & subtotal
filter_rng.Rows(1).Copy Sheets("Temp").Cells(1, 1)
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Rows(1).Insert
Sheets("Temp").Rows(1).Copy Cells(1, 1)
Sheets("Temp").Cells.Clear
Sheets("Temp").Delete

End Sub
Sub prepare_criteria()
Sheets("ToolSheet").Activate
'' remove blanks for headers and region, brands, expression
Dim last_row As Integer
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Dim rng As Range
For Each rng In Range(Cells(1, 1), Cells(1, 1).End(xlToRight))
    rng.Value = remove_blanks(rng)
Next

For Each rng In Range(Cells(1, 1), Cells(last_row, 3))
    rng.Value = remove_blanks(rng)
Next

'' inset criteria columns for later join
Dim Jan_cell As Range, BrandExpression_cell As Range, Country_cell As Range, Brand_cell As Range, Expression_cell As Range
Dim RegionBrandExpression_cell As Range
Dim offset_col_b As Integer, offset_col_e As Integer, offset_col_r As Integer
Set Country_cell = Cells.Find("Country")
Set Brand_cell = Cells.Find("Brand")
Set Jan_cell = Cells.Find("Jan")
Set Expression_cell = Cells.Find("Expression")
Jan_cell.EntireColumn.Insert
Set BrandExpression_cell = Jan_cell.Offset(0, -1)
BrandExpression_cell.Value = "BrandExpression"


offset_col_b = Brand_cell.Column - BrandExpression_cell.Column
offset_col_e = Expression_cell.Column - BrandExpression_cell.Column

Range(BrandExpression_cell, Cells(last_row, BrandExpression_cell.Column)).FormulaR1C1 = "=RC[" & offset_col_b & "]&RC[" & offset_col_e & "]"

BrandExpression_cell.EntireColumn.Insert
Set RegionBrandExpression_cell = BrandExpression_cell.Offset(0, -1)
offset_col_r = Country_cell.Column - RegionBrandExpression_cell.Column
Range(RegionBrandExpression_cell, Cells(last_row, RegionBrandExpression_cell.Column)).FormulaR1C1 = "=RC[" & offset_col_r & "]&RC[1]"
End Sub

Sub sum_depletion()
ThisWorkbook.Sheets.Add
ActiveSheet.Name = "SumDepletion"
Sheets("SumDepletion").Activate
'' copy all depletions to one spreadsheet
Dim sht As Worksheet, get_header As Boolean, last_row As Integer
get_header = True
For Each sht In ThisWorkbook.Sheets
    If sht.Name Like "20*" Or sht.Name Like "19*" Then
        If get_header Then
            sht.UsedRange.Copy Cells(1, 1)
            get_header = False
        Else
            last_row = Cells(Rows.Count, 1).End(xlUp).Row
            sht.Rows("2:" & sht.UsedRange.Rows.Count).Copy Cells(last_row + 1, 1)
        End If
    End If
Next

'' sort the sum depletion, on Year, Region and Brand
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Worksheets("SumDepletion").Sort.SortFields.Clear
Worksheets("SumDepletion").Sort.SortFields.Add2 Key:=Range( _
    "A2:A" & last_row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
ActiveWorkbook.Worksheets("SumDepletion").Sort.SortFields.Add2 Key:=Range( _
    "B2:B" & last_row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
ActiveWorkbook.Worksheets("SumDepletion").Sort.SortFields.Add2 Key:=Range( _
    "C2:C" & last_row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortNormal
With ActiveWorkbook.Worksheets("SumDepletion").Sort
    .SetRange ActiveSheet.UsedRange
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'delete other sheets other than sum depletion
Dim left_sheets, left_sheets_str
left_sheets = Array("SumDepletion", "Price", "Cost", "AMP")
left_sheets_str = Join(left_sheets, "|")
For Each sht In Sheets
    If VBA.InStr(left_sheets_str, sht.Name) <= 0 Then
        sht.Delete
    End If
Next

Cells.Find("Expression").Offset(0, 1).Value = "Factor"
Cells.Find("TotalCases").EntireColumn.Delete
End Sub
'Sub data_stack()
'Sheets("SumDepletion").Activate
''' Insert Column for Month
'Dim Jan_cell As Range
'Set Jan_cell = Cells.Find("Jan")
'Jan_cell.EntireColumn.Insert
'Jan_cell.Offset(0, -1).Value = "Month"
'
''' set the columns for stack
'Dim stack_col_count As Integer, start_stack_row As Integer, start_stack_col As Integer, insert_rows As Integer
'Dim last_row As Integer, last_col As Integer, i As Integer
'Dim Month_row_rng As Range, Month_col_rng As Range, origin_col_rng As Range, origin_data_rng As Range, unchange_rng As Range
'last_row = Cells(Rows.Count, 1).End(xlUp).Row
'last_col = Cells(1, Columns.Count).End(xlToLeft).Column
'Set Month_row_rng = Range(Jan_cell, Jan_cell.End(xlToRight))
'start_stack_col = Jan_cell.Column
'
'stack_col_count = 12
'insert_rows = stack_col_count
'start_stack_row = 2
'For i = 2 To last_row
'    Rows(start_stack_row).Resize(stack_col_count).Insert
'    '' column 1 to column 7 value copy
'    Set origin_col_rng = Range(Cells(start_stack_row + insert_rows, 1), Cells(start_stack_row + insert_rows, start_stack_col - 1))
'    Set origin_data_rng = Range(Cells(start_stack_row + insert_rows, start_stack_col), Cells(start_stack_row + insert_rows, last_col))
'    Set unchange_rng = Range(Cells(start_stack_row, 1), Cells(start_stack_row + insert_rows - 1, start_stack_col - 1))
'    Set Month_col_rng = Range(Cells(start_stack_row, start_stack_col - 1), Cells(start_stack_row + insert_rows - 1, start_stack_col - 1))
'    unchange_rng.Value = origin_col_rng.Value
'    Month_row_rng.Copy
'    Month_col_rng.PasteSpecial Transpose:=True
'    origin_data_rng.Copy
'    Month_col_rng.Offset(0, 1).PasteSpecial Transpose:=True
'    origin_data_rng.EntireRow.Delete
'    start_stack_row = start_stack_row + 12
'Next
'Month_row_rng.Clear
'Jan_cell.Value = "Case"
'End Sub
Sub data_stack()
ThisWorkbook.Sheets("SumDepletion").Activate
Dim arr, brr()
Dim a_row As Integer, a_col As Integer, b_row As Integer, b_col As Integer
Dim col_header As Integer, i As Integer
arr = Cells(1, 1).CurrentRegion
ReDim brr(1 To UBound(arr, 1) * UBound(arr, 2), 1 To 10)
' until column 8 are column header
col_header = 8
b_row = 1
'' set header name
For b_col = 1 To col_header
    brr(1, b_col) = arr(1, b_col)
Next
brr(1, 9) = "Month"
brr(1, 10) = "Case"

b_row = 1
'' start stacking data
For a_row = 2 To UBound(arr, 1) ' iteration of row
    For a_col = col_header + 1 To UBound(arr, 2) ' iteratiion of column
        b_row = b_row + 1
        For i = 1 To col_header
            brr(b_row, i) = arr(a_row, i) ' column 1 to 8 remains unchanged
        Next
        brr(b_row, col_header + 1) = arr(1, a_col) ' Month data stack
        brr(b_row, col_header + 2) = arr(a_row, a_col) ' Case data stack
    Next
Next

ActiveSheet.Cells.Clear
ActiveSheet.Cells(1, 1).Resize(b_row, col_header + 2) = brr
End Sub
Sub map_product_detail(ByVal sheet_name As Variant)
Sheets("SumDepletion").Activate
Dim CountryBrandExpression_cell As Range, find_rng As Range, find_start_cell As Range
Dim last_col As Integer, last_row As Integer, offset_col As Integer, offset_col_y As Integer


last_col = Cells(1, Columns.Count).End(xlToLeft).Column
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Set CountryBrandExpression_cell = Cells.Find("CountryBrandExpression")


Set find_start_cell = Sheets(sheet_name).Cells.Find("CountryBrandExpression")
Set find_rng = Range(find_start_cell, find_start_cell.End(xlToRight)).EntireColumn
'' start mapping
Dim Year_cell As Range
Set Year_cell = Cells.Find("Year")

offset_col = CountryBrandExpression_cell.Column - last_col - 1
offset_col_y = Year_cell.Column - last_col - 1
Cells(1, last_col + 1).Value = sheet_name & "PerCase"
Range(Cells(2, last_col + 1), Cells(last_row, last_col + 1)).FormulaR1C1 = "=iferror(vlookup(RC[" & offset_col & "]," & sheet_name & "!" & find_rng.Address(ReferenceStyle:=xlR1C1) & "," & "RC[" & offset_col_y & "]- 2016" & ",0),0)"
End Sub

Sub get_profit()
Sheets("SumDepletion").Activate
Dim Profit_cell As Range, Profit_rng As Range, last_row As Integer
Set Profit_cell = Cells.Find("AMP", lookat:=xlPart).Offset(0, 1)
Profit_cell.Value = "ProfitPerCase"
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Set Profit_rng = Range(Profit_cell.Offset(1, 0), Cells(last_row, Profit_cell.Column))
Profit_rng.FormulaR1C1 = "=RC[-3]-RC[-2]-RC[-1]"
End Sub
Sub calculate_summary()
Sheets("SumDepletion").Activate
Dim last_col As Integer, last_row As Integer
Dim source_header As Range, source_rng As Range
Dim Case_cell As Range, Case_rng As Range
Dim Total_header As Range, Total_rng As Range
last_col = Cells(1, Columns.Count).End(xlToLeft).Column
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Set Case_cell = Cells.Find("Case")
Set Case_rng = Range(Case_cell.Offset(1, 0), Cells(last_row, Case_cell.Column))
Set source_header = Range(Case_cell.Offset(0, 1), Cells(1, last_col))
Set source_rng = Range(Case_cell.Offset(1, 1), Cells(last_row, last_col))
Set Total_header = source_header.Offset(0, source_header.Count)
Set Total_rng = source_rng.Offset(0, source_header.Count)
'' Calculate Total
Total_header.FormulaArray = "=""Total"" &" & "Left(" & source_header.Address(ReferenceStyle:=xlR1C1) & ", Len(" & source_header.Address(ReferenceStyle:=xlR1C1) & ") - Len(""PerCase""))"
Total_rng.FormulaArray = "=IFERROR(" & source_rng.Address(ReferenceStyle:=xlR1C1) & "*" & Case_rng.Address(ReferenceStyle:=xlR1C1) & ",0)"

End Sub
Sub add_region()
Dim APAC, EMEA, Americas
APAC = Array("Australia", "China", "Japan", "Korea", "HongKong", "India", "Taiwan", "Vietnam", "Cambodia")
EMEA = Array("Denmark", "Diplomatic", "France", "Germany", "Ireland", _
"Italy", "Jade", "Netherlands", "Poland", "UK", "Portugal", "Slovenia", _
"SouthAfrica", "UrbanDrinks", "France/Monaco", "Norway", "Russia", "Nigeria")
Americas = Array("Bolivia", "Canada", "Panama", "USA", "Mexico", "Baltics-Latvia/Estonia/Lithuania", "Caribbean", "UrbanDrinks&Drinks.Ch")


Dim APAC_str As String, EMEA_str As String, Americas_str As String
APAC_str = Join(APAC, "|")
EMEA_str = Join(EMEA, "|")
Americas_str = Join(Americas, "|")

Sheets("SumDepletion").Activate
Dim Country_cell As Range, Region_cell As Range
Dim last_row As Integer
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Set Country_cell = Cells.Find("Country")
Country_cell.EntireColumn.Insert
Set Region_cell = Country_cell.Offset(0, -1)
Region_cell.Value = "Region"
'' add regions for countries
Dim rng As Range
For Each rng In Range(Region_cell, Cells(last_row, Region_cell.Column))
    If VBA.InStr(APAC_str, rng.Offset(0, 1).Value) > 0 Then
        rng.Value = "APAC"
    ElseIf VBA.InStr(EMEA_str, rng.Offset(0, 1).Value) > 0 Then
        rng.Value = "EMEA"
    ElseIf VBA.InStr(Americas_str, rng.Offset(0, 1).Value) > 0 Then
        rng.Value = "Americas"
    End If
Next



End Sub
Sub nine_l()
'' Create 9L version spreadsheet

Sheets("SumDepletion").Copy after:=Sheets("SumDepletion")
ActiveSheet.Name = "9LSumDepletion"

Dim last_col As Integer, last_row As Integer
Dim source_rng As Range
Dim Case_cell As Range, Case_rng As Range
Dim Factor_cell As Range, Factor_rng As Range
Dim TotalPrice_cell As Range
Set TotalPrice_cell = Cells.Find("TotalPrice", lookat:=xlPart)

last_row = Cells(Rows.Count, 1).End(xlUp).Row

Set Case_cell = Cells.Find("Case")
Set Case_rng = Range(Case_cell.Offset(1, 0), Cells(last_row, Case_cell.Column))

Set source_rng = Range(Case_cell.Offset(1, 1), Cells(last_row, TotalPrice_cell.Column - 1))
Set Factor_cell = Cells.Find("Factor")
Set Factor_rng = Range(Factor_cell.Offset(1, 0), Cells(last_row, Factor_cell.Column))
'' use the same range in the sumdepletion manipulated by factor column
Case_rng.FormulaArray = "=iferror(SumDepletion!" & Case_rng.Address(ReferenceStyle:=xlR1C1) & "*" & Factor_rng.Address(ReferenceStyle:=xlR1C1) & "/9,0)"
source_rng.FormulaArray = "=iferror(SumDepletion!" & source_rng.Address(ReferenceStyle:=xlR1C1) & "/" & Factor_rng.Address(ReferenceStyle:=xlR1C1) & "*9,0)"

Cells.Replace "Case", "9LCase"

End Sub
Function extract_year(ByVal year_string As String)
Dim a As Object
Set a = CreateObject("vbscript.regexp")
With a
    .Pattern = "\d{4}"
    .Global = True
'    extract_year = CInt(Trim(.Replace(year_string, "")))
    extract_year = CInt(.execute(year_string)(0))
End With
End Function
Function extract_catagory(ByVal catagory_string As String)
Dim a As Object
Set a = CreateObject("vbscript.regexp")
With a
    .Pattern = "\d+"
    .Global = True
    extract_catagory = Trim(.Replace(catagory_string, ""))
End With
End Function
Function remove_blanks(ByRef rng As Range)
Dim reg As Object
Set reg = CreateObject("vbscript.regexp")
With reg
    .Global = True
    .Pattern = "\s"
    remove_blanks = Trim(.Replace(rng.Value, ""))
End With
End Function
Function get_start_cell(ByVal sheet_name As String)
ThisWorkbook.Sheets(sheet_name).Activate
Dim region_name, region, regions, start_cells As Range, start_cell As Range, origin_col As Integer, rng As Range
Dim LE_start_cell As Range '' for LE current region use

regions = Array("Americas", "EMEA", "APAC")
For Each region In regions
    If sheet_name Like "*" & region & "*" Then
        region_name = region
    End If
Next
Set start_cells = Rows(1).Find(region_name, after:=Cells(1, 1), SearchDirection:=xlNext, MatchCase:=False, LookIn:=xlValues, SearchOrder:=xlByRows)
Set start_cell = Rows(1).FindNext(after:=start_cells)

'' both LE and history+plan have actual depletions, to avoid duplication, only historyAndPlan copy actuals
If Not sheet_name Like "*LE*" Then

    ' if only one start_cell
    If start_cells.Address = start_cell.Address Then
        start_cells.Offset(0, 1).Value = Trim(UCase(start_cells.Offset(0, 1).Value))
        Set get_start_cell = start_cells
    
    Else
        origin_col = start_cells.Column
        Do While start_cell.Column <> origin_col
            Set start_cells = Union(start_cells, start_cell)
            Set start_cell = Rows(1).FindNext(after:=start_cell)
        Loop
        For Each rng In start_cells
            rng.Offset(0, 1).Value = Trim(UCase(rng.Offset(0, 1).Value))
        Next
        Set get_start_cell = start_cells
    End If
Else
    ' if only one start_cell
    If start_cells.Address = start_cell.Address Then
        start_cells.Offset(0, 1).Value = Trim(UCase(start_cells.Offset(0, 1).Value))
        Set get_start_cell = start_cells
    
    Else
        origin_col = start_cells.Column
        Do While start_cell.Column <> origin_col
            Set start_cells = Union(start_cells, start_cell)
            Set start_cell = Rows(1).FindNext(after:=start_cell)
        Loop
        
    '' find the start cell belongs to LE to avoid copy actuals in LE spread sheet AGAIN
        For Each rng In start_cells
            If extract_year(rng.Offset(0, 1).Value) >= Year(Now()) Then
                If LE_start_cell Is Nothing Then
                    Set LE_start_cell = rng
                Else
                    Set LE_start_cell = Union(LE_start_cell, rng)
                End If
            End If
        Next
        Set get_start_cell = LE_start_cell
    End If
    
End If
End Function
Function get_new_sheet_name(ByVal sheet_name As String, ByVal judge_value As Variant)
Dim current_year, data_year, appendix As String
current_year = Year(Now())
data_year = extract_year(judge_value)
If data_year < current_year Or UCase(sheet_name) Like UCase("*[a-z]*Actual*") Then
    appendix = "Actual"
ElseIf data_year >= current_year And UCase(sheet_name) Like UCase("*LE*") Then
    appendix = "LE"
Else
    appendix = "Budget"
End If
get_new_sheet_name = data_year & appendix
End Function

Sub add_date()
Sheets("SumDepletion").Activate
Dim Region_cell As Range, Month_cell As Range, Date_rng As Range
Set Region_cell = Cells.Find("Region")
Set Month_cell = Cells.Find("Month")
Region_cell.EntireColumn.Resize(, 2).Insert
Month_cell.EntireColumn.Copy Region_cell.Offset(0, -2).EntireColumn
Month_cell.EntireColumn.Delete
Region_cell.Offset(0, -1).Value = "Date"
'' Create Date
Dim rng As Range, last_row As Integer
last_row = Cells(Rows.Count, Region_cell.Column).End(xlUp).Row
Set Date_rng = Range(Region_cell.Offset(1, -1), Cells(last_row, Region_cell.Column - 1))
Dim month_days As String
Dim large_months, small_months
large_months = Array("Jan", "Mar", "May", "Jul", "Aug", "Dec")
small_months = Array("Apr", "Jun", "Sep", "Oct", "Nov")
For Each rng In Date_rng
    If InStr(Join(large_months, "|"), rng.Offset(0, -1)) > 0 Then
        month_days = "31"
    ElseIf InStr(Join(small_months, "|"), rng.Offset(0, -1)) > 0 Then
        month_days = "30"
    ElseIf rng.Offset(0, -2).Value Like "*00" And rng.Offset(0, -2).Value Mod 400 = 0 Then
        month_days = "29"
    ElseIf rng.Offset(0, -2).Value Mod 4 = 0 Then
        month_days = "29"
    Else
        month_days = "28"
    End If
    rng.Value = CDate(rng.Offset(0, -2) & " " & rng.Offset(0, -1) & " " & month_days)

    
Next
End Sub


