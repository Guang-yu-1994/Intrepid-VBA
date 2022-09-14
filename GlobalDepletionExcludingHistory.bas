Attribute VB_Name = "GlobalDepletionExcludingHistory"
Option Explicit
Public header, current_year
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
Dim plan_source_name As String, file_name As Variant, f As Variant, plan_source As Workbook
Set tool_sheet = ThisWorkbook.Sheets("ToolSheet")

file_name = Application.GetOpenFilename("excels,*xls*", 1, "select file", , True)
If IsArray(file_name) Then
    For Each f In file_name
        Set plan_source = Workbooks.Open(f)
        plan_source_name = Trim(Split(plan_source.Name, "-")(1))
        For Each sht In plan_source.Sheets
            If UCase(sht.Name) Like UCase("*Plan*History*") Then
                sht.Cells.Copy
                ThisWorkbook.Sheets.Add after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Cells(1, 1).PasteSpecial xlPasteValues
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = plan_source_name & "HistoryAndPlan"
            ElseIf UCase(sht.Name) Like UCase("*Actual*") Then
                sht.Cells.Copy
                ThisWorkbook.Sheets.Add after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Cells(1, 1).PasteSpecial xlPasteValues
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = plan_source_name & "Actual"
            ElseIf UCase(sht.Name) Like UCase("*LE1*") Then
                sht.Cells.Copy
                ThisWorkbook.Sheets.Add after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Cells(1, 1).PasteSpecial xlPasteValues
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = plan_source_name & "LE1"
            ElseIf UCase(sht.Name) Like UCase("*LE2*") Then
                sht.Cells.Copy
                ThisWorkbook.Sheets.Add after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Cells(1, 1).PasteSpecial xlPasteValues
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = plan_source_name & "LE2"
            ElseIf UCase(sht.Name) Like UCase("*LE3*") Then
                sht.Cells.Copy
                ThisWorkbook.Sheets.Add after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Cells(1, 1).PasteSpecial xlPasteValues
                ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = plan_source_name & "LE3"
            End If
        Next
        plan_source.Close
    Next
    current_year = InputBox("is the workbook year current year? if no enter the year")
    If current_year = "" Or IsEmpty(current_year) Or current_year = False Then
        current_year = Year(Now())
    End If

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
    If UCase(sht.Name) Like UCase("*LE1*") Then
        Call copy_data(sht.Name)
    End If
Next

For Each sht In ThisWorkbook.Sheets
    If UCase(sht.Name) Like UCase("*LE2*") Then
        Call copy_data(sht.Name)
    End If
Next

For Each sht In ThisWorkbook.Sheets
    If UCase(sht.Name) Like UCase("*LE3*") Then
        Call copy_data(sht.Name)
    End If
Next


Call sum_depletion
Call preprocess_data
Call standardize_data
Call concate_history_usa_Third
Call process_detail_data

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



'' leave only one header
Dim Expression_cell As Range
'' delete Expression cell rows except for the first one (header)
Set Expression_cell = Cells.Find("Expression", after:=Cells(1, 1), searchdirection:=xlNext)
Expression_cell.Offset(0, -2).Value = "Country"
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
Sub sum_depletion()
ThisWorkbook.Sheets.Add
ActiveSheet.Name = "SumDepletion"
Sheets("SumDepletion").Activate
'' copy all depletions to one spreadsheet
Dim sht As Worksheet, get_header As Boolean, last_row As Long
get_header = True
For Each sht In ThisWorkbook.Sheets
    If sht.Name Like "20*" Then
        If get_header Then
            sht.UsedRange.Copy Cells(1, 1)
            get_header = False
        Else
            last_row = Cells(Rows.Count, 1).End(xlUp).Row
            sht.Rows("2:" & sht.UsedRange.Rows.Count).Copy Cells(last_row + 1, 1)
        End If
    End If
Next

''' sort the sum depletion, on Year, Region and Brand
'last_row = Cells(Rows.Count, 1).End(xlUp).Row
'Worksheets("SumDepletion").Sort.SortFields.Clear
'Worksheets("SumDepletion").Sort.SortFields.Add2 Key:=Range( _
'    "A2:A" & last_row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'    xlSortNormal
'ActiveWorkbook.Worksheets("SumDepletion").Sort.SortFields.Add2 Key:=Range( _
'    "B2:B" & last_row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'    xlSortNormal
'ActiveWorkbook.Worksheets("SumDepletion").Sort.SortFields.Add2 Key:=Range( _
'    "C2:C" & last_row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
'    xlSortNormal
'With ActiveWorkbook.Worksheets("SumDepletion").Sort
'    .SetRange ActiveSheet.UsedRange
'    .header = xlYes
'    .MatchCase = False
'    .Orientation = xlTopToBottom
'    .SortMethod = xlPinYin
'    .Apply
'End With

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

Sub preprocess_data()
Call replace_name("A")
Dim month_days As String
Dim feb_days, date_header
Dim large_months, small_months

If current_year Like "*00" And current_year Mod 400 = 0 Then
    feb_days = "29"
ElseIf current_year Mod 4 = 0 Then
    feb_days = "29"
Else
    feb_days = "28"
End If

'' create header date
Dim Jan As Range
Set Jan = Cells.Find("Jan")
date_header = Array("31/01/" & current_year _
                    , feb_days & "/02/" & current_year _
                    , "31/03/" & current_year _
                    , "30/04/" & current_year _
                    , "31/05/" & current_year _
                    , "30/06/" & current_year _
                    , "31/07/" & current_year _
                    , "31/08/" & current_year _
                    , "30/09/" & current_year _
                    , "31/10/" & current_year _
                    , "30/11/" & current_year _
                    , "31/12/" & current_year)
Range(Cells(1, Jan.Column), Cells(1, Jan.Column + 11)).Value = date_header
Range(Cells(1, Jan.Column), Cells(1, Jan.Column + 11)).NumberFormat = "mmm-yy"

Dim rng As Range
For Each rng In Range(Cells(1, Jan.Column), Cells(1, Jan.Column + 11))
    rng.Value = CDate(rng.Value)
Next

Cells(1, Jan.Column + 12).Value = "Case"
header = Range(Cells(1, 1), Cells(1, Jan.Column + 12))

''' Calculate Total
Dim last_row As Long, last_col As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row
last_col = Cells(1, Columns.Count).End(xlToLeft).Column

Range(Cells(2, last_col), Cells(last_row, last_col)).FormulaR1C1 = "=IFERROR(sum(RC[-12]:RC[-1]),0)"

'' FILTER OUT 0
ActiveSheet.UsedRange.AutoFilter Field:=Jan.Column + 12, Criteria1:="<0.5", _
    Operator:=xlOr, Criteria2:="="
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells(1, 1).EntireRow.Insert
Range(Cells(1, 1), Cells(1, last_col)) = header
Range(Cells(1, 1), Cells(1, last_col)).NumberFormat = "mmm-yy"

'' delete the total case column
Cells(1, last_col).EntireColumn.Delete
last_row = Cells(Rows.Count, 1).End(xlUp).Row
last_col = Cells(1, Columns.Count).End(xlToLeft).Column

'' get variant and case config from expression
Dim Expression_cell As Range, Date_cell As Range
Set Expression_cell = Cells.Find("Expression")
Expression_cell.EntireColumn.Resize(, 2).Insert
Expression_cell.Offset(0, -2).Value = "Variant"
Range(Expression_cell.Offset(1, -2), Cells(last_row, Expression_cell.Column - 2)).Formula = "=get_variant(RC[2])"
Range(Expression_cell.Offset(1, -2), Cells(last_row, Expression_cell.Column - 2)).Value = Range(Expression_cell.Offset(1, -2), Cells(last_row, Expression_cell.Column - 2)).Value

Expression_cell.Offset(0, -1).Value = "Case Config"
Range(Expression_cell.Offset(1, -1), Cells(last_row, Expression_cell.Column - 1)).Formula = "=get_case_config(RC[1])"
Range(Expression_cell.Offset(1, -1), Cells(last_row, Expression_cell.Column - 1)).Value = Range(Expression_cell.Offset(1, -1), Cells(last_row, Expression_cell.Column - 1)).Value

'' create duty status column
Expression_cell.Value = "DutyStatus"
Range(Expression_cell.Offset(1, 0), Cells(last_row, Expression_cell.Column)).Clear


End Sub


Function extract_year(ByVal year_string As String)
Dim a As Object
Set a = CreateObject("vbscript.regexp")
With a
    .Pattern = "\d{4}"
    .Global = True
'    extract_year = CInt(Trim(.Replace(year_string, "")))
On Error GoTo no_year:
    extract_year = CInt(.execute(year_string)(0))
    Exit Function
End With
no_year:
    extract_year = 0
End Function
Function extract_category(ByVal category_string As String)
Dim a As Object
Set a = CreateObject("vbscript.regexp")
With a
    .Pattern = "\d{4}"
    .Global = True
    extract_category = Trim(.Replace(category_string, ""))
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
Dim region_name, region, regions, origin_col As Long, rng As Range


regions = Array("Americas", "EMEA", "APAC")
For Each region In regions
    If sheet_name Like "*" & region & "*" Then
        region_name = region
    End If
Next


If sheet_name Like "*LE1*" Then
    Set get_start_cell = Rows(1).Find(current_year & " LE1*").Offset(0, -1)
ElseIf sheet_name Like "*LE2*" Then
    Set get_start_cell = Rows(1).Find(current_year & " LE2*").Offset(0, -1)
ElseIf sheet_name Like "*LE3*" Then
    Set get_start_cell = Rows(1).Find(current_year & " LE3*").Offset(0, -1)
ElseIf sheet_name Like "*Actual*" Then
    Set get_start_cell = Rows(1).Find(current_year & " ACTUALS*").Offset(0, -1)
ElseIf sheet_name Like "*HistoryAndPlan*" Then
    Set get_start_cell = Rows(1).Find(current_year & " PLAN*").Offset(0, -1)
End If

End Function
Function get_new_sheet_name(ByVal sheet_name As String, ByVal judge_value As Variant)
Dim current_year, data_year, appendix As String
current_year = Year(Now())
data_year = extract_year(judge_value)
If data_year < current_year Or UCase(sheet_name) Like UCase("*[a-z]*Actual*") Then
    appendix = "Actual"
ElseIf data_year >= current_year And UCase(sheet_name) Like UCase("*LE1*") Then
    appendix = "LE1"
ElseIf data_year >= current_year And UCase(sheet_name) Like UCase("*LE2*") Then
    appendix = "LE2"
ElseIf data_year >= current_year And UCase(sheet_name) Like UCase("*LE3*") Then
    appendix = "LE3"
ElseIf data_year >= current_year And UCase(sheet_name) Like UCase("*Plan*") Then
    appendix = "Budget"
End If
get_new_sheet_name = data_year & appendix
End Function
'Sub add_date()
'Sheets("SumDepletion").Activate
'Dim Region_cell As Range, Month_cell As Range, Date_rng As Range
'Set Region_cell = Cells.Find("Region")
'Set Month_cell = Cells.Find("Month")
'Region_cell.EntireColumn.Resize(, 2).Insert
'Month_cell.EntireColumn.Copy Region_cell.Offset(0, -2).EntireColumn
'Month_cell.EntireColumn.Delete
'Region_cell.Offset(0, -1).Value = "Date"
''' Create Date
'Dim rng As Range, last_row As Long
'last_row = Cells(Rows.Count, Region_cell.Column).End(xlUp).Row
'Set Date_rng = Range(Region_cell.Offset(1, -1), Cells(last_row, Region_cell.Column - 1))
'Dim month_days As String
'Dim large_months, small_months
'large_months = Array("Jan", "Mar", "May", "Jul", "Aug", "Oct", "Dec")
'small_months = Array("Apr", "Jun", "Sep", "Nov")
'For Each rng In Date_rng
'    If InStr(Join(large_months, "|"), rng.Offset(0, -1)) > 0 Then
'        month_days = "31"
'    ElseIf InStr(Join(small_months, "|"), rng.Offset(0, -1)) > 0 Then
'        month_days = "30"
'    ElseIf rng.Offset(0, -2).Value Like "*00" And rng.Offset(0, -2).Value Mod 400 = 0 Then
'        month_days = "29"
'    ElseIf rng.Offset(0, -2).Value Mod 4 = 0 Then
'        month_days = "29"
'    Else
'        month_days = "28"
'    End If
'    rng.Value = CDate(rng.Offset(0, -2).Value & " " & rng.Offset(0, -1).Value & " " & month_days)
'Next
'End Sub
'Sub calculate_total()
'ThisWorkbook.Sheets("SumDepletion").Activate
'Dim Case_cell As Range, Case_rng As Range, Detail_header As Range, detail_rng As Range
'Dim last_row As Long
'Set Case_cell = Cells.Find("Case", lookat:=xlWhole)
'last_row = Cells(Rows.Count, Case_cell.Column).End(xlUp).Row
'Set Case_rng = Range(Case_cell.Offset(1, 0), Cells(last_row, Case_cell.Column))
'
'Set Detail_header = Range(Case_cell.Offset(0, 1), Case_cell.Offset(0, 3))
'Set detail_rng = Range(Case_cell.Offset(1, 1), Cells(last_row, Case_cell.Column + 3))
'
'Detail_header.Offset(0, 3).FormulaArray = "=""Total"" & " & Detail_header.Address(ReferenceStyle:=xlR1C1)
'detail_rng.Offset(0, 3).FormulaArray = "=" & detail_rng.Address(ReferenceStyle:=xlR1C1) & "*" & Case_rng.Address(ReferenceStyle:=xlR1C1)
'End Sub



Sub change_to_value()
Dim sht As Worksheet
For Each sht In ThisWorkbook.Sheets
    sht.Cells.Copy
    sht.Cells(1, 1).PasteSpecial xlPasteValuesAndNumberFormats
Next
End Sub

'Function is_leTwo(sheet_name As String, start_cell As Range)
'Dim cat_str As String
'Dim if_le_two As Boolean, year_str As String
'year_str = Trim(Str(extract_year(start_cell.Offset(0, 1).Value)))
'If sheet_name Like "*LE2*" Then
'    '' judge
'    cat_str = Trim(Replace(start_cell.Offset(0, 1).Value, year_str, ""))
'    is_leTwo = (cat_str = "LE2")
'Else
'    is_leTwo = True
'End If
'End Function


