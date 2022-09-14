Attribute VB_Name = "GlobalDepletionExtractor"
Option Explicit
Public wb_month As String, month_num
Public header, myPlanType, current_year
Public archive_path As String
Public is_all_Actual As Boolean
Sub execute()
Call import
Call preprocess
Call standardize_data
Call concat
Call process_detail_data
Call budget_use
End Sub
Sub concat()
Call export_current ' before concat export current
Call concat_previous_months
Call concat_other_depletion("Ireland")
Call concat_history("Depletion")
End Sub
Sub import()
Dim planType
planType = MsgBox("Depletion choose yes, Shipment choose no", vbYesNo)
If planType = vbYes Then
    myPlanType = "*DEPLETION*GOAL*"
Else
    myPlanType = "*SHIPMENT*PLAN*"
End If

Application.ScreenUpdating = False
Application.DisplayAlerts = False

Dim sht As Worksheet, tool_sheet As Worksheet
    

'' Add toolsheet and PlanSummary
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
    If sht.Name = Split(myPlanType, "*")(1) Then
        sheet_exist = True
        Sheets(Split(myPlanType, "*")(1)).Cells.Clear
    End If
Next
If sheet_exist = False Then
    ThisWorkbook.Sheets.Add
    ActiveSheet.Name = Split(myPlanType, "*")(1)
End If

For Each sht In ThisWorkbook.Sheets
    If sht.Name <> "ToolSheet" And sht.Name <> Split(myPlanType, "*")(1) Then
        sht.Delete
    End If
Next


'' import data source
Dim last_row As Long, get_header As Boolean
Dim plan_source_name As String, file_name As Variant, f As Variant, plan_source As Workbook
Dim Goal_cell As Range
Set tool_sheet = ThisWorkbook.Sheets("ToolSheet")

file_name = Application.GetOpenFilename("excels,*xls*", 1, "select file", , True)

If IsArray(file_name) Then
    For Each f In file_name
        Set plan_source = Workbooks.Open(f, 3)
        wb_month = get_wb_month(plan_source)
        is_all_Actual = False
        If wb_month = "" Then
            is_all_Actual = True
        End If
        For Each sht In plan_source.Sheets
            sht.Activate
            Set Goal_cell = Cells.Find(myPlanType, searchdirection:=xlNext, MatchCase:=True, after:=Cells(1, 1), lookat:=xlPart, LookIn:=xlValues)
            If Not Goal_cell Is Nothing And Not sht.Name Like "*Copy*" And Not sht.Name Like "*not*ready*" Then
                Call copy_goal_cell(Goal_cell, sht)
            End If
        Next
        plan_source.Close
    Next
Else
    End
End If
End Sub
Sub copy_goal_cell(Goal_cell As Range, sht As Worksheet)
ThisWorkbook.Sheets("ToolSheet").Cells.Clear
Goal_cell.CurrentRegion.Copy
ThisWorkbook.Sheets("ToolSheet").Cells(1, 1).PasteSpecial Paste:=xlPasteValues

'' remove extra rows if exist
ThisWorkbook.Sheets("ToolSheet").Activate
Dim Dec As Range
Set Dec = Cells.Find("Dec", after:=Cells(1, 1))
Range(Dec.Offset(0, 1), Cells(Dec.Row, Columns.Count)).EntireColumn.Delete


Dim t_goal_cell As Range
Set t_goal_cell = Cells.Find(Goal_cell.Value, after:=Cells(1, 1), lookat:=xlPart)
If t_goal_cell.Row <> 1 Then
    Rows("1:" & t_goal_cell.Row - 1).Delete
End If

Dim total_cell As Range
Set total_cell = t_goal_cell.EntireColumn.Find("Total", MatchCase:=False, searchdirection:=xlNext, lookat:=xlPart)
Rows(total_cell.Row & ":" & Rows.Count).Delete


'' delete the first row
Rows(1).Delete

'' ready to copy
Dim t_last_row As Long, t_col As Long, t_last_col As Long
Dim s_last_row As Long
t_last_row = Cells(Rows.Count, 1).End(xlUp).Row



'' add Country
Cells(1, 1).EntireColumn.Insert
Cells(1, 1).Value = "Country"
Range(Cells(2, 1), Cells(t_last_row, 1)).Value = sht.Name

'' add category
Dim Jan As Range
month_num = Month(wb_month & "-2022")
Cells(1, 1).EntireColumn.Insert
Cells(1, 1).Value = "Category"
Set Jan = Cells.Find("Jan")

If is_all_Actual Then
    month_num = 0
    Range(Cells(2, 1), Cells(t_last_row, 1)).Value = "Actual"
    'Only Dec Data, Jan to Nov would be done by concating, so hear delete
    Range(Jan.Offset(1, 0), Cells(t_last_row, Jan.Column + 10)).Clear
Else
    Range(Cells(2, 1), Cells(t_last_row, 1)).Value = "B" & month_num
    ' copy actuals
    Dim col_header As Range
    Set col_header = Range(Cells(2, 1), Cells(t_last_row, Jan.Column - 1))
    If wb_month <> "Jan" Then
        Dim divide_month As Range
        Set divide_month = Cells.Find(wb_month)
        Range(Cells(2, divide_month.Column - 1), Cells(t_last_row, divide_month.Column - 1)).Copy Cells(t_last_row + 1, divide_month.Column - 1) ' last months data are actual, copy them and append them
        col_header.Copy col_header.Offset(col_header.Rows.Count, 0) ' copy the col header after copying the actual month
        col_header.Columns(1).Offset(col_header.Rows.Count, 0).Value = "Actual" ' mark the actual data as Category Actual
        
        '' LE1 appears then Apr, LE2 appears when the wb is Jul
        If wb_month = "Apr" Or wb_month = "Jul" Or wb_month = "Oct" Then
            col_header.Copy col_header.Offset(col_header.Rows.Count * 2, 0) ' col header for LE
            Range(Cells(2, col_header.Columns.Count + 1), Cells(t_last_row, divide_month.Column - 1)).Copy Cells(2, col_header.Columns.Count + 1).Offset(col_header.Rows.Count * 2, 0) ' data for LE
            col_header.Columns(1).Offset(col_header.Rows.Count * 2, 0).Value = "LE" & (Month(wb_month & "-2022") - 1) / 3 ' LE for category
        End If
        
        Range(Cells(2, col_header.Columns.Count + 1), Cells(t_last_row, divide_month.Column - 1)).Clear 'the previou months data are actual instead of budget, so clear,all previous clear because historical actual would be done by concating
    End If
End If



'' insert duty status column
Jan.EntireColumn.Insert
Jan.Offset(0, -1).Value = "DutyStatus"
'' store the header
t_last_col = Cells(1, Columns.Count).End(xlToLeft).Column
header = Range(Cells(1, 1), Cells(1, t_last_col))

'' copy the data

s_last_row = ThisWorkbook.Sheets(Split(myPlanType, "*")(1)).Cells(Rows.Count, 1).End(xlUp).Row
t_last_row = Cells(Rows.Count, 1).End(xlUp).Row
Range(Cells(2, 1), Cells(t_last_row, t_last_col)).Copy ThisWorkbook.Sheets(Split(myPlanType, "*")(1)).Cells(s_last_row + 1, 1)


ThisWorkbook.Sheets("ToolSheet").Activate
Cells.Clear

End Sub
Sub preprocess()
Call replace_name("Operating Depart")
ThisWorkbook.Sheets(Split(myPlanType, "*")(1)).Activate
Cells(1, 1).Resize(, UBound(header, 2)) = header


Dim month_days As String
Dim feb_days, date_header
Dim large_months, small_months
current_year = InputBox("is the workbook year current year? if no enter the year")
If current_year = "" Or IsEmpty(current_year) Or current_year = False Then
    current_year = Year(Now())
End If

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
Range(Cells(1, 6), Cells(1, last_col)).NumberFormat = "mmm-yy"

'' delete the total case column
Cells(1, last_col).EntireColumn.Delete
End Sub
Sub export_current()


Dim plan_type As String
If myPlanType Like "*DEPLETION*" Then
    archive_path = "F:\Intrepid Spirits\Arnout\New Depletions\ArchivedDepletionData\"
    plan_type = "DEPLETION"
Else
    archive_path = "F:\Intrepid Spirits\Arnout\New Depletions\ArchivedShipmentData\"
    plan_type = "SHIPMENT"
End If
ThisWorkbook.Sheets(1).Copy

ActiveWorkbook.SaveCopyAs archive_path & plan_type & Format(CDate(wb_month & "-" & current_year), "yyyy-mm") & ".xlsx"
ActiveWorkbook.Close
End Sub
Sub budget_use()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
ThisWorkbook.Sheets(1).Activate
Dim new_wb As Workbook
Set new_wb = Workbooks.Add

new_wb.Sheets(1).Name = "ShipmentPlan"
'' paste as values, otherwise cannot delete because of array formula
ThisWorkbook.Sheets(1).Cells.Copy
new_wb.Sheets("ShipmentPlan").Activate
Cells.PasteSpecial xlPasteValuesAndNumberFormats

'' delete budget in Sheet shipment plan
new_wb.Sheets("ShipmentPlan").Activate
Dim Category_cell As Range, header, Category_col As Integer
Dim last_row As Long, last_col As Integer

last_col = ActiveSheet.UsedRange.Columns.Count
header = ActiveSheet.UsedRange.Rows(1)

Set Category_cell = Cells.Find("Category")
Category_col = Category_cell.Column


'' in the shipmentPlan sheet, only actuals and B & month_num
new_wb.Sheets("ShipmentPlan").Activate
If month_num <> 0 Then
    ActiveSheet.UsedRange.AutoFilter Field:=Category_col, Criteria1:="<>*Actual*", Operator:=xlAnd, Criteria2:="<>*B" & month_num & "*"
Else
    ActiveSheet.UsedRange.AutoFilter Field:=Category_col, Criteria1:="<>*Actual*"
End If
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells(1, 1).EntireRow.Insert
Range(Cells(1, 1), Cells(1, last_col)) = header

'' generate pivot table


'' save as, dir first to see if the file exists, delete if exists
Dim save_as_path As String
save_as_path = "F:\Intrepid Spirits\Arnout\New Depletions\BudgetUseShipmentPlan\" & wb_month & " " & current_year & ".xlsx"
If Len(Dir(save_as_path)) = 0 Then
    new_wb.SaveAs save_as_path
Else
    Kill save_as_path
    new_wb.SaveAs save_as_path
End If
new_wb.Close

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub
Sub get_pivot()
Sheets.Add
ActiveSheet.Name = "Pivot"
ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    "ShipmentPlan!" & Sheets("ShipmentPlan").UsedRange.Address(True, True, ReferenceStyle:=xlR1C1), Version:=8).CreatePivotTable TableDestination _
    :="Pivot!R1C1", TableName:="pivot_sale_and_cos", DefaultVersion:=8
'With ActiveSheet.PivotTables("pivot_sale_and_cos")
'    .ColumnGrand = True
'    .HasAutoFormat = True
'    .DisplayErrorString = False
'    .DisplayNullString = True
'    .EnableDrilldown = True
'    .ErrorString = ""
'    .MergeLabels = False
'    .NullString = ""
'    .PageFieldOrder = 2
'    .PageFieldWrapCount = 0
'    .PreserveFormatting = True
'    .RowGrand = True
'    .SaveData = True
'    .PrintTitles = False
'    .RepeatItemsOnEachPrintedPage = True
'    .TotalsAnnotation = False
'    .CompactRowIndent = 1
'    .InGridDropZones = False
'    .DisplayFieldCaptions = True
'    .DisplayMemberPropertyTooltips = False
'    .DisplayContextTooltips = True
'    .ShowDrillIndicators = True
'    .PrintDrillIndicators = False
'    .AllowMultipleFilters = False
'    .SortUsingCustomLists = True
'    .FieldListSortAscending = False
'    .ShowValuesRow = False
'    .CalculatedMembersInFilters = False
'    .RowAxisLayout xlCompactRow
'End With
With ActiveSheet.PivotTables("pivot_sale_and_cos").PivotCache
    .RefreshOnFileOpen = False
    .MissingItemsLimit = xlMissingItemsDefault
End With
ActiveSheet.PivotTables("pivot_sale_and_cos").RepeatAllLabels xlRepeatLabels
With ActiveSheet.PivotTables("pivot_sale_and_cos").PivotFields("Country")
    .Orientation = xlRowField
    .Position = 1
End With
ActiveSheet.PivotTables("pivot_sale_and_cos").AddDataField ActiveSheet.PivotTables( _
    "pivot_sale_and_cos").PivotFields("TotalPrice"), "Sum of TotalPrice", xlSum
ActiveSheet.PivotTables("pivot_sale_and_cos").AddDataField ActiveSheet.PivotTables( _
    "pivot_sale_and_cos").PivotFields("TotalCost"), "Sum of TotalCost", xlSum


End Sub
Sub concat_previous_months()
Dim archived_month, current_month, reg As Object
Dim wb_name As String, previous_month_wb As Workbook
Dim last_row As Long

Set reg = CreateObject("vbscript.regexp")
With reg
    .Global = True
    .Pattern = "\d{4}-\d{2}"
End With
wb_name = Dir(archive_path & "*xls*")
Do While wb_name <> ""
    archived_month = CDbl(CDate(reg.execute(wb_name)(0)))
    current_month = CDbl(CDate(wb_month & "-" & current_year))
    If is_all_Actual Then
        If wb_month <> "Dec" Then
            Set previous_month_wb = Workbooks.Open(archive_path & wb_name, 0)
            ThisWorkbook.Sheets(1).Activate
            last_row = ActiveSheet.UsedRange.Rows.Count
            previous_month_wb.Sheets(1).UsedRange.Copy Cells(last_row + 1, 1)
            Cells(last_row + 1, 1).EntireRow.Delete ' on need for extra header
            previous_month_wb.Close
        End If
    Else
        If archived_month < current_month Then
            Set previous_month_wb = Workbooks.Open(archive_path & wb_name, 0)
            ThisWorkbook.Sheets(1).Activate
            last_row = ActiveSheet.UsedRange.Rows.Count
            previous_month_wb.Sheets(1).UsedRange.Copy Cells(last_row + 1, 1)
            Cells(last_row + 1, 1).EntireRow.Delete ' on need for extra header
            previous_month_wb.Close
        End If
    End If

    wb_name = Dir
Loop
Set reg = Nothing
End Sub

Function get_wb_month(wb As Workbook)
Dim reg As Object
Set reg = CreateObject("vbscript.regexp")
With reg
    .Global = True
    .Pattern = "[JFMASOND][aepuco][nbrylgptvc]"
    On Error GoTo no_match:
    get_wb_month = .execute(wb.Name)(0)
    Exit Function
no_match:
get_wb_month = ""
End With
End Function



