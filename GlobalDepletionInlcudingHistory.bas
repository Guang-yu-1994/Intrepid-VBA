Attribute VB_Name = "GlobalDepletionInlcudingHistory"
Option Explicit
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
    If sht.Name <> "ToolSheet" Or sht.Name = "Depletions" Then
        sht.Delete
    End If
Next



'' import data source
Dim last_row As Long, get_header As Boolean
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
'' trim value before replace
'' trim first cost less time because stack could create more row data
Call trim_rng(1, 5)
Call replace_name("BOD")
Call data_stack
Call add_region
Call add_date
Call reformat_depletion
Call map_product_detail
Call calculate_total
'Call concat_history
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
        '' delete columns after the 1st Cases
        Sheets("ToolSheet").Activate
        Dim Cases As Range
        Set Cases = Cells.Find("Cases", after:=Cells(1, 1), searchdirection:=xlNext)
        Range(Cases.Offset(0, 1), Cells(Cases.Row, Columns.Count)).EntireColumn.Delete
        
        '' delete columns before market if market is not in the 1st col
        Dim MKT As Range
        Set MKT = Cells.Find("Market", after:=Cells(1, 1), searchdirection:=xlNext)
        If MKT.Column <> 1 Then
            Range(Cells(MKT.Row, 1), MKT.Offset(0, -1)).EntireColumn.Delete
        End If
        
        '' create new spreadsheet for storing the data
        '' the new sheet name depends on budget or actual or LE1 or LE2
        Sheets(sheet_name).Activate
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
'        Call prepare_criteria
        
        '' Add year column to the spreadsheet
        ThisWorkbook.Sheets("ToolSheet").Activate
        Columns("A:A").Insert
        Cells(1, 1).Value = "Year"
        Range(Cells(2, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 1)).Value = extract_year(to_sheet)
        
        '' add depletion catagory to the spreadsheet
        Columns("A:A").Insert
        Cells(1, 1).Value = "Category"
        Range(Cells(2, 1), Cells(ActiveSheet.UsedRange.Rows.Count, 1)).Value = extract_category(to_sheet)
        
        
        Dim last_row As Long, last_row_tool As Long
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
Dim mkt_cell As Range, f_mkt_cell As Range, fill_until_row As Long
Set mkt_cell = Cells.Find("Market", after:=Cells(1, 1), searchdirection:=xlPrevious, lookat:=xlWhole)
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
Dim last_row As Long, last_col As Long
last_row = Cells(Rows.Count, mkt_cell.Column).End(xlUp).Row
last_col = Cases_cell.Column
Set filter_rng = Range(Cells(1, 1), Cells(last_row, last_col))

Set mkt_cell = Cells.Find("Market", after:=Cells(1, 1), searchdirection:=xlNext, lookat:=xlWhole)
filter_rng.AutoFilter Field:=mkt_cell.Column, Criteria1:="=*Market*", Operator:=xlOr, Criteria2:="=*Total*"
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete

'' Delete 0 cases columns
filter_rng.AutoFilter Field:=last_col, Criteria1:="<0.5", Criteria2:="=", Operator:=xlOr
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells.Find("Brand").Offset(0, -1).Value = "Country"


'' leave only one header
Dim Expression_cell As Range
'' delete Expression cell rows except for the first one (header)
Set Expression_cell = Cells.Find("Expression", after:=Cells(1, 1), searchdirection:=xlNext)
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

Dim Date_rng As Range
Dim last_row As Long
Set Date_rng = Cells.Find("Date")
Date_rng.EntireColumn.NumberFormat = "mmm-yy"
last_row = Cells(Rows.Count, Date_rng.Column).End(xlUp).Row
Dim join_string As String


'' Some cost items has GEXP, so it is necessary to create another criteria for this situation

Date_rng.EntireColumn.Insert
Date_rng.Offset(0, -1).Value = "CriteriaForCost1"
join_string = "RC[-8]&RC[-7]&RC[-6]&RC[-5]"
Range(Date_rng.Offset(1, -1), Cells(last_row, Date_rng.Column - 1)).FormulaR1C1 = "=SUBSTITUTE(TEXT(RC[1],""MMM/YYYY"")&" & join_string & ","" "","""")"

Date_rng.EntireColumn.Insert
Date_rng.Offset(0, -1).Value = "CriteriaForCost2"
join_string = """GEXP""&RC[-8]&RC[-7]&RC[-6]"
Range(Date_rng.Offset(1, -1), Cells(last_row, Date_rng.Column - 1)).FormulaR1C1 = "=SUBSTITUTE(TEXT(RC[1],""MMM/YYYY"")&" & join_string & ","" "","""")"

Date_rng.EntireColumn.Insert
Date_rng.Offset(0, -1).Value = "CriteriaForPrice"
join_string = "RC[-10]&RC[-9]&RC[-8]&RC[-7]&RC[-6]"
Range(Date_rng.Offset(1, -1), Cells(last_row, Date_rng.Column - 1)).FormulaR1C1 = "=SUBSTITUTE(TEXT(RC[1],""MMM/YYYY"")&" & join_string & ","" "","""")"



End Sub

Sub sum_depletion()
ThisWorkbook.Sheets.Add
ActiveSheet.Name = "SumDepletion"
Sheets("SumDepletion").Activate
'' copy all depletions to one spreadsheet
Dim sht As Worksheet, get_header As Boolean, last_row As Long
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
Cells.Find("Factor").EntireColumn.Delete

End Sub

Sub data_stack()
Dim sht As Worksheet

Sheets("SumDepletion").Activate
Dim arr, brr()
Dim a_row As Long, a_col As Long, b_row As Long, b_col As Long
Dim col_header As Long, i As Long


' until column Jan are column header
Dim Jan As Range
Dim last_col As Long
' find Jan cell
last_col = ActiveSheet.UsedRange.Columns.Count
'For i = 1 To last_col
'    If IsDate(Cells(1, i).Value) Then
'        Set Jan = Cells(1, i)
'        Exit For
'    End If
'Next
Set Jan = Cells.Find("Jan")
col_header = Jan.Column - 1

arr = Cells(1, 1).CurrentRegion
ReDim brr(1 To UBound(arr, 1) * UBound(arr, 2), 1 To col_header + 2)

b_row = 1
'' set header name
For b_col = 1 To col_header
    brr(1, b_col) = arr(1, b_col)
Next
brr(1, col_header + 1) = "Month"

brr(1, col_header + 2) = "Case"

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

Sheets("SumDepletion").Activate
Sheets("SumDepletion").Cells.Clear
ActiveSheet.Cells(1, 1).Resize(b_row, col_header + 2) = brr

End Sub


'Sub add_region()
'Dim APAC, EMEA, Americas
'APAC = Array("Australia", "China", "Japan", "Korea", "Hong Kong", "India", "Taiwan", "Vietnam", "Cambodia")
'
'EMEA = Array("Denmark", "Diplomatic", "Germany", "Ireland", _
'"Italy", "Jade", "Netherlands", "Poland", "UK", "Portugal", "Slovenia", _
'"South Africa", "Urban Drinks", "France", "Norway", "Russia", "Nigeria", "Baltics", "Dublin Airport", "Northern Ireland")
'
'Americas = Array("Bolivia", "Canada", "Panama", "USA", "Mexico", "Caribbean", "UrbanDrinks&Drinks.Ch")
'
'
'Dim APAC_str As String, EMEA_str As String, Americas_str As String
'APAC_str = Join(APAC, "|")
'EMEA_str = Join(EMEA, "|")
'Americas_str = Join(Americas, "|")
'
'Sheets("SumDepletion").Activate
'Dim Country_cell As Range, Region_cell As Range
'Dim last_row As Long
'last_row = Cells(Rows.Count, 1).End(xlUp).Row
'Set Country_cell = Cells.Find("Country")
'Country_cell.EntireColumn.Insert
'Set Region_cell = Country_cell.Offset(0, -1)
'Region_cell.Value = "Region"
''' add regions for countries
'Dim rng As Range
'For Each rng In Range(Region_cell, Cells(last_row, Region_cell.Column))
'    If VBA.InStr(APAC_str, rng.Offset(0, 1).Value) > 0 Then
'        rng.Value = "APAC"
'    ElseIf VBA.InStr(EMEA_str, rng.Offset(0, 1).Value) > 0 Then
'        rng.Value = "EMEA"
'    ElseIf VBA.InStr(Americas_str, rng.Offset(0, 1).Value) > 0 Then
'        rng.Value = "Americas"
'    End If
'Next
'End Sub

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
Function trim_blanks(ByRef rng As Range)
'Dim reg As Object
'Set reg = CreateObject("vbscript.regexp")
'With reg
'    .Global = True
'    .Pattern = "\s"
'    trim_blanks = Trim(.Replace(rng.Value, ""))
'End With
trim_blanks = Trim(rng.Value)
End Function
Function get_start_cell(ByVal sheet_name As String)
ThisWorkbook.Sheets(sheet_name).Activate
Dim region_name, region, regions, start_cells As Range, start_cell As Range, origin_col As Long, rng As Range
Dim LE_start_cell As Range '' for LE current region use

regions = Array("Americas", "EMEA", "APAC")
For Each region In regions
    If sheet_name Like "*" & region & "*" Then
        region_name = region
    End If
Next
Set start_cells = Rows(1).Find(region_name, after:=Cells(1, 1), searchdirection:=xlNext, MatchCase:=False, LookIn:=xlValues, SearchOrder:=xlByRows)
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
        If extract_year(start_cell.Offset(0, 1).Value) >= Year(Now()) And is_leTwo(sheet_name, start_cell) Then
            start_cells.Offset(0, 1).Value = Trim(UCase(start_cells.Offset(0, 1).Value))
            Set get_start_cell = start_cells
        End If
    Else
        origin_col = start_cells.Column
        Do While start_cell.Column <> origin_col
            Set start_cells = Union(start_cells, start_cell)
            Set start_cell = Rows(1).FindNext(after:=start_cell)
        Loop
        
    '' find the start cell belongs to LE to avoid copy actuals in LE spread sheet AGAIN
        For Each rng In start_cells
            If extract_year(rng.Offset(0, 1).Value) >= Year(Now()) And is_leTwo(sheet_name, rng) Then
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
Dim rng As Range, last_row As Long
last_row = Cells(Rows.Count, Region_cell.Column).End(xlUp).Row
Set Date_rng = Range(Region_cell.Offset(1, -1), Cells(last_row, Region_cell.Column - 1))
Dim month_days As String
Dim large_months, small_months
large_months = Array("Jan", "Mar", "May", "Jul", "Aug", "Oct""Dec")
small_months = Array("Apr", "Jun", "Sep", "Nov")
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

Sub reformat_depletion()
ThisWorkbook.Sheets("SumDepletion").Activate
Dim Expression_cell As Range, Date_cell As Range
Set Expression_cell = Cells.Find("Expression")
Expression_cell.EntireColumn.Resize(, 6).Insert
Expression_cell.Offset(0, -6).Resize(, 6) = Array("Variant", "Case Config", "DutyStatus", "ABV", "BottlesPerCase", "MLPerBottle")
Cells.Find("Year").EntireColumn.Delete
Cells.Find("Month").EntireColumn.Delete
Expression_cell.Offset(0, 1).EntireColumn.Insert

Set Date_cell = Cells.Find("Date")
Date_cell.EntireColumn.Copy
Expression_cell.Offset(0, 1).EntireColumn.PasteSpecial xlPasteValues
Expression_cell.Offset(0, 1).EntireColumn.NumberFormat = "m/d/yyyy"
Date_cell.EntireColumn.Delete


'' remove 0 case data
Dim header_arr, criteria_col As Integer
header_arr = Cells(1, 1).CurrentRegion.Rows(1)
Dim Case_cell As Range
Set Case_cell = Cells.Find("Case", lookat:=xlWhole)
criteria_col = Case_cell.Column
ActiveSheet.UsedRange.AutoFilter Field:=criteria_col, Criteria1:="<0.5", _
            Operator:=xlOr, Criteria2:="="
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete

Cells(1, 1).EntireRow.Insert
Cells(1, 1).Resize(, criteria_col) = header_arr


Dim last_row As Long
last_row = ActiveSheet.UsedRange.Rows.Count

''' trim value before replace
'Call trim_rng(6, 8)
'Call replace_name

Set Expression_cell = Cells.Find("Expression")
Range(Expression_cell.Offset(1, -6), Cells(last_row, Expression_cell.Column - 6)).Formula = "=get_variant(RC[6])"
Range(Expression_cell.Offset(1, -5), Cells(last_row, Expression_cell.Column - 5)).Formula = "=get_case_config(RC[5])"
Range(Expression_cell.Offset(1, -2), Cells(last_row, Expression_cell.Column - 2)).Formula = "=get_bottles_per_case(RC[-3])"
Range(Expression_cell.Offset(1, -1), Cells(last_row, Expression_cell.Column - 1)).Formula = "=get_ml_per_bottle(RC[-4])"

Cells.Copy
Cells.PasteSpecial xlPasteValues
Expression_cell.EntireColumn.Delete


Call prepare_criteria
End Sub


Sub calculate_total()
ThisWorkbook.Sheets("SumDepletion").Activate
Dim Case_cell As Range, Case_rng As Range, Detail_header As Range, detail_rng As Range
Dim last_row As Long
Set Case_cell = Cells.Find("Case", lookat:=xlWhole)
last_row = Cells(Rows.Count, Case_cell.Column).End(xlUp).Row
Set Case_rng = Range(Case_cell.Offset(1, 0), Cells(last_row, Case_cell.Column))

Set Detail_header = Range(Case_cell.Offset(0, 1), Case_cell.Offset(0, 3))
Set detail_rng = Range(Case_cell.Offset(1, 1), Cells(last_row, Case_cell.Column + 3))

Detail_header.Offset(0, 3).FormulaArray = "=""Total"" & " & Detail_header.Address(ReferenceStyle:=xlR1C1)
detail_rng.Offset(0, 3).FormulaArray = "=" & detail_rng.Address(ReferenceStyle:=xlR1C1) & "*" & Case_rng.Address(ReferenceStyle:=xlR1C1)
End Sub

Sub nine_l()
'' Create 9L version spreadsheet

Sheets("SumDepletion").Copy after:=Sheets("SumDepletion")
ActiveSheet.Name = "9LSumDepletion"

Dim last_col As Long, last_row As Long
Dim detail_rng As Range
Dim Case_cell As Range, Case_rng As Range
Dim Factor_cell As Range, Factor_rng As Range
Dim TotalPrice_cell As Range
Set TotalPrice_cell = Cells.Find("TotalPrice", lookat:=xlPart, LookIn:=xlValues)
Dim BottlesPerCase_cell As Range, BottlesPerCase_rng As Range
Dim ML_cell As Range, ML_rng As Range
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Set BottlesPerCase_cell = Cells.Find("BottlesPerCase")
Set BottlesPerCase_rng = Range(BottlesPerCase_cell.Offset(1, 0), Cells(last_row, BottlesPerCase_cell.Column))

Set ML_cell = Cells.Find("MLPerBottle")
Set ML_rng = Range(ML_cell.Offset(1, 0), Cells(last_row, ML_cell.Column))





Set Case_cell = Cells.Find("Case", lookat:=xlWhole)
Set Case_rng = Range(Case_cell.Offset(1, 0), Cells(last_row, Case_cell.Column))

Set detail_rng = Range(Case_cell.Offset(1, 1), Cells(last_row, TotalPrice_cell.Column - 1))


'' use the same range in the sumdepletion manipulated by factor column
Case_rng.FormulaArray = "=iferror(SumDepletion!" & Case_rng.Address(ReferenceStyle:=xlR1C1) & "*" & ML_rng.Address(ReferenceStyle:=xlR1C1) & "*" & BottlesPerCase_rng.Address(ReferenceStyle:=xlR1C1) & "/9000,0)"
detail_rng.FormulaArray = "=iferror(SumDepletion!" & detail_rng.Address(ReferenceStyle:=xlR1C1) & "/(" & ML_rng.Address(ReferenceStyle:=xlR1C1) & "*" & BottlesPerCase_rng.Address(ReferenceStyle:=xlR1C1) & ")" & "*9000,0)"

Cells.Replace "Case", "9LCase"

End Sub

Function is_leTwo(sheet_name As String, start_cell As Range)
Dim cat_str As String
Dim if_le_two As Boolean, year_str As String
year_str = Trim(Str(extract_year(start_cell.Offset(0, 1).Value)))
If sheet_name Like "*LE2*" Then
    cat_str = Trim(Replace(start_cell.Offset(0, 1).Value, year_str, ""))
    is_leTwo = (cat_str = "LE2")
Else
    is_leTwo = True
End If
End Function


