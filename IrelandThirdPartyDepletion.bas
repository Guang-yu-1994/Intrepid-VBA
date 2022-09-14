Attribute VB_Name = "IrelandThirdPartyDepletion"
Sub execute()
Call import_plan
Call pre_process
Call data_stack
'Call reframe_data
End Sub

Sub import_plan()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim sht As Worksheet, tool_sheet As Worksheet
    

'' Add Summary and depletions
Dim sheet_exist As Boolean
sheet_exist = False
For Each sht In ThisWorkbook.Sheets
    If sht.Name = "Summary" Then
        sheet_exist = True
        Sheets("Summary").Cells.Clear
    End If
Next
If sheet_exist = False Then
    ThisWorkbook.Sheets.Add
    ActiveSheet.Name = "Summary"
End If



For Each sht In ThisWorkbook.Sheets
    If sht.Name <> "Summary" Then
        sht.Delete
    End If
Next



'' import data source
Dim last_row As Long, get_header As Boolean
Dim plan_source_name As String, file_name As Variant, f As Variant, plan_source As Workbook
Dim Goal_cell As Range

file_name = Application.GetOpenFilename("excels,*xls*", 1, "select file", , True)
If IsArray(file_name) Then
    For Each f In file_name
        Set plan_source = GetObject(f)
        plan_source.Sheets("Intrepid IRE Depletions").Cells.Copy
        ThisWorkbook.Sheets("Summary").Cells(1, 1).PasteSpecial xlPasteValues
        plan_source.Close
    Next
Else
    End
End If

End Sub
Sub pre_process()
ThisWorkbook.Sheets("Summary").Activate

'' replace name
Call replace_name("Crevan")

'' only current year
Dim Dec_cell As Range
Set Dec_cell = Cells.Find("Dec", lookat:=xlPart, searchdirection:=xlNext, after:=Cells(1, 1))
Range(Dec_cell.Offset(0, 1), Cells(Dec_cell.Row, Columns.Count)).EntireColumn.Delete

'' only third party
Dim third_party_start As Range
Set third_party_start = Cells.Find("Little & Green", after:=Cells(1, 1), searchdirection:=xlNext, searchorder:=xlByRows)
Range(Cells(1, 1), third_party_start.Offset(-1, 0)).EntireRow.Delete

'' no totals
Dim total_cell As Range
Set total_cell = third_party_start.EntireColumn.Find("Total", lookat:=xlPart, searchdirection:=xlNext)
Rows(total_cell.Row & ":" & Rows.Count).Delete

'' new header
Dim header_array
header_array = Array("Brand", "Expression", "Factor")
Rows(1).EntireRow.Insert
Range(third_party_start.Offset(-1, 0), third_party_start.Offset(-1, UBound(header_array))) = header_array
Range(third_party_start.Offset(-1, UBound(header_array) + 1), third_party_start.Offset(-1, UBound(header_array) + 12)) = get_date_header
Range(third_party_start.Offset(-1, UBound(header_array) + 1), third_party_start.Offset(-1, UBound(header_array) + 12)).NumberFormat = "mmm-yy"

'' add country
Dim last_row As Integer, Brand_cell As Range
last_row = ActiveSheet.UsedRange.Rows.Count
Set Brand_cell = Cells.Find("Brand")
Brand_cell.Offset(0, -1).Value = "Country"
Range(Brand_cell.Offset(1, -1), Cells(last_row, Brand_cell.Column - 1)).Value = "Ireland"

''' add region
'Brand_cell.Offset(0, -2).Value = "Region"
'Range(Brand_cell.Offset(1, -2), Cells(last_row, Brand_cell.Column - 2)).Value = "EMEA"

''add category
Brand_cell.Offset(0, -2).Value = "Category"
Range(Brand_cell.Offset(1, -2), Cells(last_row, Brand_cell.Column - 2)).Value = "Budget"

'' delete factor column and the 1st empty column
Cells.Find("Factor").EntireColumn.Delete
Cells(1, 1).EntireColumn.Delete

'' add variant and Case config
Dim Jan_cell As Range
Set Jan_cell = Cells.Find("Jan", LookIn:=xlValues)
Jan_cell.Resize(, 2).EntireColumn.Insert
Jan_cell.Offset(0, -1).Value = "CaseConfig"
Range(Jan_cell.Offset(1, -1), Cells(last_row, Jan_cell.Column - 1)).FormulaR1C1 = "=get_case_config(RC[-2])"

Jan_cell.Offset(0, -2).Value = "Variant"
Range(Jan_cell.Offset(1, -2), Cells(last_row, Jan_cell.Column - 2)).FormulaR1C1 = "=get_variant(RC[-1])"

'' copy as value and delete expression
Cells.Copy
Cells(1, 1).PasteSpecial xlPasteValues
Cells.Find("Expression").EntireColumn.Delete

'' choose unitl which month could be regarded as Actual
Dim Actual_month As String, Actual_month_cell As Range
ActualMonth.Show
Actual_month = ActualMonth.ActualMonthList.SelText

Set Actual_month_cell = Cells.Find(Actual_month, LookIn:=xlValues)
Range(Cells(2, 1), Cells(last_row, Actual_month_cell.Column)).Copy Cells(last_row + 1, 1)
Range(Cells(last_row + 1, 1), Cells(last_row + last_row - 1, 1)).Value = "Actual"


End Sub



Sub data_stack()
Dim sht As Worksheet, goal_sheet As Worksheet
Set goal_sheet = Sheets("Summary")

goal_sheet.Activate
Dim arr, brr()
Dim a_row As Long, a_col As Long, b_row As Long, b_col As Long
Dim col_header As Long, i As Long
Dim source_type

' until column Jan are column header
Dim Jan As Range
Dim last_col As Long
' find Jan cell
last_col = ActiveSheet.UsedRange.Columns.Count
For i = 1 To last_col
    If IsDate(Cells(1, i).Value) Then
        Set Jan = Cells(1, i)
        Exit For
    End If
Next
col_header = Jan.Column - 1

arr = Cells(1, 1).CurrentRegion
ReDim brr(1 To UBound(arr, 1) * UBound(arr, 2), 1 To col_header + 2)

b_row = 1
'' set header name
For b_col = 1 To col_header
    brr(1, b_col) = arr(1, b_col)
Next
brr(1, col_header + 1) = "Date"
source_type = "Case"
brr(1, col_header + 2) = source_type

b_row = 1
'' start stacking data
For a_row = 2 To UBound(arr, 1) ' iteration of row
    For a_col = col_header + 1 To UBound(arr, 2) ' iteratiion of column
        b_row = b_row + 1
        For i = 1 To col_header
            brr(b_row, i) = arr(a_row, i) ' column 1 to5 remains unchanged
        Next
        brr(b_row, col_header + 1) = arr(1, a_col) ' Date_rng data stack
        brr(b_row, col_header + 2) = arr(a_row, a_col) ' Case data stack
    Next
Next

Sheets.Add
ActiveSheet.Name = "Stacked"
Cells.Clear
ActiveSheet.Cells(1, 1).Resize(b_row, col_header + 2) = brr



'' FILTER OUT 0

last_col = Cells(1, Columns.Count).End(xlToLeft).Column
Dim header
header = Range(Cells(1, 1), Cells(1, last_col))

ActiveSheet.UsedRange.AutoFilter Field:=last_col, Criteria1:="<0.5", _
    Operator:=xlOr, Criteria2:="="
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells(1, 1).EntireRow.Insert
Range(Cells(1, 1), Cells(1, last_col)) = header

End Sub
'Sub reframe_data()
'
'Dim Brand_cell As Range, last_row As Long
'last_row = Cells(Rows.Count, 1).End(xlUp).Row
'Set Brand_cell = Cells.Find("Brand")
'
'
'
'' case detial
'Call get_case_detail
'
''' prepare criteria
''Call prepare_criteria
'
'''AVB price and cost
''Call map_product_detail
''
''' total price cost
''Call calculate_total
'
'End Sub

Function get_date_header()
Dim current_year, feb_days
current_year = Year(Now())

If current_year Like "*00" And current_year Mod 400 = 0 Then
    feb_days = "29"
ElseIf current_year Mod 4 = 0 Then
    feb_days = "29"
Else
    feb_days = "28"
End If

'' create header date
date_header = Array(CDate("31/01/" & current_year) _
                    , CDate(feb_days & "/02/" & current_year) _
                    , CDate("31/03/" & current_year) _
                    , CDate("30/04/" & current_year) _
                    , CDate("31/05/" & current_year) _
                    , CDate("30/06/" & current_year) _
                    , CDate("31/07/" & current_year) _
                    , CDate("31/08/" & current_year) _
                    , CDate("30/09/" & current_year) _
                    , CDate("31/10/" & current_year) _
                    , CDate("30/11/" & current_year) _
                    , CDate("31/12/" & current_year))
get_date_header = date_header
End Function
Sub get_case_detail()


Dim Date_cell As Range, Date_rng As Range
Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Set Date_cell = Cells.Find("Date")
Set Date_rng = Range(Date_cell.Offset(1, 0), Cells(last_row, Date_cell.Column))
Date_cell.EntireColumn.Resize(, 2).Insert
Date_cell.Offset(0, -1).Value = "MLPerBottle"
Date_cell.Offset(0, -2).Value = "BottlesPerCase"
Date_rng.Offset(0, -1).FormulaR1C1 = "=iferror(get_ml_per_bottle(RC[-2]),0)"
Date_rng.Offset(0, -2).FormulaR1C1 = "=iferror(get_bottles_per_case(RC[-1]),0)"

Date_cell.Offset(0, -2).EntireColumn.Resize(, 2).Insert
Date_cell.Offset(0, -4).Value = "DutyStatus"
Date_cell.Offset(0, -3).Value = "ABV"
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





Sub calculate_total()
'ThisWorkbook.Sheets("StackedShipmentPlan").Activate
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
