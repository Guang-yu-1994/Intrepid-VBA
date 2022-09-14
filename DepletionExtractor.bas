Attribute VB_Name = "DepletionExtractor"
Option Explicit
Sub get_shipment_plan()
Call import
Call reframe_shipmennt_plan
'must trip before replace, trim before stack comsume less time
Call trim_rng(1, 5)
Call data_stack
Call replace_name("Operating Depart")
Call add_region
Call get_case_detail
Call prepare_criteria
Call map_product_detail
Call calculate_total
Call nine_l
Call export
End Sub


Sub import()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim sht As Worksheet, tool_sheet As Worksheet
    

'' Add toolsheet and ShipmentPlan
Dim sheet_exist As Boolean
sheet_exist = False
For Each sht In ThisWorkbook.Sheets
    If sht.Name = "ToolSheet" Then sheet_exist = True
    Sheets("ToolSheet").Cells.Clear
Next
If sheet_exist = False Then
    ThisWorkbook.Sheets.Add
    ActiveSheet.Name = "ToolSheet"
End If

sheet_exist = False
For Each sht In ThisWorkbook.Sheets
    If sht.Name = "ShipmentPlan" Then sheet_exist = True
    Sheets("ShipmentPlan").Cells.Clear
Next
If sheet_exist = False Then
    ThisWorkbook.Sheets.Add
    ActiveSheet.Name = "ShipmentPlan"
End If

For Each sht In ThisWorkbook.Sheets
    If sht.Name <> "ToolSheet" And sht.Name <> "ShipmentPlan" Then
        sht.Delete
    End If
Next


'' import data source
Dim last_row As Long, get_header As Boolean
Dim plan_source_name As String, file_name As Variant, f As Variant, plan_source As Workbook
Dim SHIPMENT_PLAN As Range
Set tool_sheet = ThisWorkbook.Sheets("ToolSheet")

file_name = Application.GetOpenFilename("excels,*xls*", 1, "select file", , True)
If IsArray(file_name) Then
    For Each f In file_name
        Set plan_source = Workbooks.Open(f)
        For Each sht In plan_source.Sheets
            sht.Activate
            Set SHIPMENT_PLAN = sht.Cells.Find("SHIPMENT PLAN", SearchDirection:=xlPrevious, MatchCase:=True)
            If Not SHIPMENT_PLAN Is Nothing Then
                Dim s
                s = sht.Name
                Call copy_shipment_plan(SHIPMENT_PLAN, sht)
            End If
        Next
        plan_source.Close
    Next
Else
    End
End If
End Sub

Sub copy_shipment_plan(SHIPMENT_PLAN As Range, sht As Worksheet)
ThisWorkbook.Sheets("ToolSheet").Cells.Clear
SHIPMENT_PLAN.CurrentRegion.Copy
ThisWorkbook.Sheets("ToolSheet").Cells(1, 1).PasteSpecial Paste:=xlPasteValues

'' ready to copy
ThisWorkbook.Sheets("ToolSheet").Activate
Dim t_last_row As Long, t_col As Long, t_last_col As Long
Dim s_last_row As Long
Dim index_rng As Range, Actual_rng As Range, Budget_rng As Range
t_last_row = Cells(Rows.Count, 1).End(xlUp).Row
t_last_col = Cells(2, Columns.Count).End(xlToLeft).Column
Set index_rng = Range(Cells(4, 1), Cells(t_last_row, 3))


'' Copy Actual
s_last_row = ThisWorkbook.Sheets("ShipmentPlan").Cells(Rows.Count, 3).End(xlUp).Row + 1


'' distinguish Acutal and Budget Columns

For t_col = 4 To t_last_col
    If t_col Mod 2 = 1 Then
        If Actual_rng Is Nothing Then
            Set Actual_rng = Range(Cells(4, t_col), Cells(t_last_row, t_col))
        Else
            Set Actual_rng = Union(Actual_rng, Range(Cells(4, t_col), Cells(t_last_row, t_col)))
        End If
    Else
        If Budget_rng Is Nothing Then
            Set Budget_rng = Range(Cells(4, t_col), Cells(t_last_row, t_col))
        Else
            Set Budget_rng = Union(Budget_rng, Range(Cells(4, t_col), Cells(t_last_row, t_col)))
        End If
    End If
Next


ThisWorkbook.Sheets("ShipmentPlan").Activate
Range(Cells(s_last_row, 1), Cells(s_last_row + index_rng.Rows.Count - 1, 1)).Value = "Actual"
Range(Cells(s_last_row, 2), Cells(s_last_row + index_rng.Rows.Count - 1, 2)).Value = sht.Name
'index_rng.Copy
'ThisWorkbook.Sheets("ShipmentPlan").Cells(s_last_row, 3).PasteSpecial Paste:=xlValues
index_rng.Copy ThisWorkbook.Sheets("ShipmentPlan").Cells(s_last_row, 3)

Actual_rng.Copy Cells(s_last_row, 6)

'' Copy Budget
s_last_row = s_last_row + index_rng.Rows.Count + 1
Range(Cells(s_last_row, 1), Cells(s_last_row + index_rng.Rows.Count - 1, 1)).Value = "Budget"
Range(Cells(s_last_row, 2), Cells(s_last_row + index_rng.Rows.Count - 1, 2)).Value = sht.Name
'index_rng.Copy
'ThisWorkbook.Sheets("ShipmentPlan").Cells(s_last_row, 3).PasteSpecial Paste:=xlValues
index_rng.Copy ThisWorkbook.Sheets("ShipmentPlan").Cells(s_last_row, 3)

Budget_rng.Copy Cells(s_last_row, 6)

ThisWorkbook.Sheets("ToolSheet").Activate
Cells.Clear

End Sub

Sub reframe_shipmennt_plan()
Sheets("ShipmentPlan").Activate
''' create header
Dim header
Cells(1, 1).Value = "Category"
Cells(1, 2).Value = "Country"
Cells(1, 3).Value = "Brand"
Cells(1, 4).Value = "Variant"
Cells(1, 5).Value = "Case Config"

Dim month_days As String
Dim current_year, feb_days, date_header
Dim large_months, small_months
current_year = Year(Now())

If current_year Like "*00" And current_year Mod 400 = 0 Then
    feb_days = "29"
ElseIf current_year Mod 4 = 0 Then
    feb_days = "29"
Else
    feb_days = "28"
End If

'' create header date
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
Range(Cells(1, 6), Cells(1, 17)).Value = date_header
Dim rng As Range
For Each rng In Range(Cells(1, 6), Cells(1, 17))
    rng.Value = CDate(rng.Value)
Next

Cells(1, 18).Value = "Case"
header = Range(Cells(1, 1), Cells(1, 18))

''' Calculate Total
Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row
'Range(Cells(2, 18), Cells(last_row, 18)).FormulaR1C1 = "=sum(RC[-12]:RC[-1])"
Range(Cells(2, 18), Cells(last_row, 18)).FormulaR1C1 = "=IFERROR(sum(RC[-12]:RC[-1]),0)"

'' FILTER OUT 0
ActiveSheet.UsedRange.AutoFilter Field:=18, Criteria1:="<0.5", _
    Operator:=xlOr, Criteria2:="="
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells(1, 1).EntireRow.Insert
Range(Cells(1, 1), Cells(1, 18)) = header
Range(Cells(1, 6), Cells(1, 17)).NumberFormat = "mmm-yy"

'' delete the total case column
Cells(1, 18).EntireColumn.Delete

'' Delete total shipment rows
ActiveSheet.UsedRange.AutoFilter Field:=3, Criteria1:="*Total*", _
    Operator:=xlOr, Criteria2:="="
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells(1, 1).EntireRow.Insert
Range(Cells(1, 1), Cells(1, 18)) = header

End Sub

Sub data_stack()
Dim sht As Worksheet, goal_sheet As Worksheet

Set goal_sheet = Sheets("ShipmentPlan")

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
On Error Resume Next
ActiveSheet.Name = "StackedShipmentPlan"
Sheets("StackedShipmentPlan").Activate
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

Sub add_region()
Dim APAC, EMEA, Americas
APAC = Array("Australia", "China", "Japan", "Korea", "Hong Kong", "India", "Taiwan", "Vietnam", "Cambodia")

EMEA = Array("Denmark", "Diplomatic", "Germany", "Ireland", _
"Italy", "Jade", "Netherlands", "Poland", "UK", "Portugal", "Slovenia", _
"South Africa", "Urban Drinks", "France", "Norway", "Russia", "Nigeria", "Baltics", "Dublin Airport")

Americas = Array("Bolivia", "Canada", "Panama", "USA", "Mexico", "Caribbean", "UrbanDrinks&Drinks.Ch")


Dim APAC_str As String, EMEA_str As String, Americas_str As String
APAC_str = Join(APAC, "|")
EMEA_str = Join(EMEA, "|")
Americas_str = Join(Americas, "|")

Sheets("StackedShipmentPlan").Activate
Dim Country_cell As Range, Region_cell As Range
Dim last_row As Long
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

Sub get_case_detail()
Sheets("StackedShipmentPlan").Activate
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
Sub map_product_detail()
Sheets("StackedShipmentPlan").Activate
Dim detail_path As String, Case_cell As Range, last_row As Long
detail_path = "F:\Intrepid Spirits\Budget\Budet Restructure\ProductDetailStructure\ProductDetail\"
Set Case_cell = Cells.Find("Case", lookat:=xlWhole)
last_row = Cells(Rows.Count, Case_cell.Column).End(xlUp).Row
Case_cell.Offset(0, 1).Value = "Price"
'=VLOOKUP(M2,'F:\Intrepid Spirits\Budget\Budet Restructure\ProductDetailStructure\ProductDetail\[PriceData.xlsx]Sheet1'!$I:$K,3,0)
Range(Case_cell.Offset(1, 1), Cells(last_row, Case_cell.Column + 1)).FormulaR1C1 = "=VLOOKUP(RC[-3],'" & detail_path & "[PriceData.xlsx]Sheet1'!C9:C11,3,0)"

Case_cell.Offset(0, 2).Value = "Cost"
Range(Case_cell.Offset(1, 2), Cells(last_row, Case_cell.Column + 2)).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],'" & detail_path & "[CostData.xlsm]CostPivot'!C6:C7,2,0),VLOOKUP(RC[-5],'" & detail_path & "[CostData.xlsm]CostPivot'!C6:C7,2,0))"

Case_cell.Offset(0, 3).Value = "Margin"
Range(Case_cell.Offset(1, 3), Cells(last_row, Case_cell.Column + 3)).FormulaR1C1 = "=RC[-2]-RC[-1]"

'' get ABV
Dim ABV_cell As Range, Variant_cell As Range

Set ABV_cell = Cells.Find("ABV")
Set Variant_cell = Cells.Find("Variant")
Range(ABV_cell.Offset(1, 0), Cells(last_row, ABV_cell.Column)).FormulaR1C1 = "=VLOOKUP(RC[" & Variant_cell.Column - ABV_cell.Column & "],'" & detail_path & "[ABV.xlsx]ABV'!C2:C5,3,0)"
Range(ABV_cell.Offset(1, 0), Cells(last_row, ABV_cell.Column)).NumberFormat = "0.00%"



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

Sub nine_l()
'' Create 9L version spreadsheet

ActiveSheet.Copy after:=ActiveSheet
ActiveSheet.Name = "9LStackedShipmentPlan"

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
Case_rng.FormulaArray = "=iferror(StackedShipmentPlan!" & Case_rng.Address(ReferenceStyle:=xlR1C1) & "*" & ML_rng.Address(ReferenceStyle:=xlR1C1) & "*" & BottlesPerCase_rng.Address(ReferenceStyle:=xlR1C1) & "/9000,0)"
detail_rng.FormulaArray = "=iferror(StackedShipmentPlan!" & detail_rng.Address(ReferenceStyle:=xlR1C1) & "/(" & ML_rng.Address(ReferenceStyle:=xlR1C1) & "*" & BottlesPerCase_rng.Address(ReferenceStyle:=xlR1C1) & ")" & "*9000,0)"

Cells.Replace "Case", "9LCase"
End Sub

Sub export()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
ThisWorkbook.Sheets("StackedShipmentPlan").Activate
Dim new_wb As Workbook
Set new_wb = Workbooks.Add

new_wb.Sheets(1).Name = "ShipmentPlan"
'' paste as values, otherwise cannot delete because of array formula
ThisWorkbook.Sheets("StackedShipmentPlan").Cells.Copy
new_wb.Sheets("ShipmentPlan").Activate
Cells.PasteSpecial xlPasteValuesAndNumberFormats




'' delete budget in Sheet shipment plan
new_wb.Sheets("ShipmentPlan").Activate
Dim Category_cell As Range, header, Category_col As Integer
Dim last_row As Long, last_col As Integer
Dim Region_cell As Range
last_col = ActiveSheet.UsedRange.Columns.Count
header = ActiveSheet.UsedRange.Rows(1)

' delete the rows without region
Set Region_cell = Cells.Find("Region")
Set Category_cell = Cells.Find("Category")
Category_col = Category_cell.Column

ActiveSheet.UsedRange.AutoFilter Field:=Region_cell.Column, Criteria1:="="
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells(1, 1).EntireRow.Insert
Range(Cells(1, 1), Cells(1, last_col)) = header
ActiveSheet.Copy after:=ActiveSheet
ActiveSheet.Name = "Tool"

new_wb.Sheets("ShipmentPlan").Activate
ActiveSheet.UsedRange.AutoFilter Field:=Category_col, Criteria1:="Budget"
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells(1, 1).EntireRow.Insert
Range(Cells(1, 1), Cells(1, last_col)) = header


last_row = ActiveSheet.UsedRange.Rows.Count

' get lastest date in actual
Dim latest_date As Date
Dim Date_cell As Range, Date_col As Integer
Set Date_cell = Cells.Find("Date")
Date_col = Date_cell.Column
latest_date = Application.WorksheetFunction.Max(Range(Date_cell.Offset(1, 0), Cells(last_row, Date_cell.Column)))
' for before fileter, <= is not working in VBA date filtering
latest_date = DateAdd("d", 1, CStr(latest_date))

'' filter budget in Tool sheet whose date is after the date of latest actual, no need to keep header this time
new_wb.Sheets("Tool").Activate
header = ActiveSheet.UsedRange.Rows(1)

ActiveSheet.UsedRange.AutoFilter Field:=Category_col, Criteria1:="Actual"
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells(1, 1).EntireRow.Insert
Range(Cells(1, 1), Cells(1, last_col)) = header

ActiveSheet.UsedRange.AutoFilter Field:=Date_col, Criteria1:="<" & CDbl(latest_date)

Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete

ActiveSheet.UsedRange.Copy new_wb.Sheets("ShipmentPlan").Cells(last_row + 1, 1)
ActiveSheet.Delete

'' save as, dir first to see if the file exists, delete if exists
Dim save_as_path As String
save_as_path = "F:\Intrepid Spirits\Arnout\ScoreCard\Shipment Extractor\BudgetUseShipmentPlan.xlsx"
If Len(Dir(save_as_path)) = 0 Then
    new_wb.SaveCopyAs save_as_path
Else
    Kill save_as_path
    new_wb.SaveCopyAs save_as_path
End If
new_wb.Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
