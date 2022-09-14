Attribute VB_Name = "DepletionDataProcessor"
Sub standardize_data()
Call trim_rng(1, 5)
Call data_stack
End Sub
Sub concate_history_usa_Third()
Call concat_other_depletion("USA")
Call concat_other_depletion("Ireland")
Call concat_history("Depletion")
End Sub
Sub process_detail_data()
Call add_region
Call get_case_detail
Call prepare_criteria
Call map_product_detail
Call calculate_total
Call nine_l
End Sub
Sub trim_rng(from_col As Integer, ByVal until_col As Integer)
'On Error Resume Next
Dim arr, brr, r As Long, c As Long
arr = Range(Cells(1, from_col), Cells(ActiveSheet.UsedRange.Rows.Count, until_col))
ReDim brr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2))
For r = LBound(arr, 1) To UBound(arr, 1)
    For c = LBound(arr, 2) To UBound(arr, 2)
        brr(r, c) = Trim(arr(r, c))
    Next c
Next r
Range(Cells(1, from_col), Cells(ActiveSheet.UsedRange.Rows.Count, until_col)) = brr
End Sub
Sub replace_name(replace_source As String)
Dim wb_name As String
If replace_source = "Operating Depart" Then
    wb_name = "Replacement For Lindsey.xlsx"
Else
    wb_name = "Replacement For Arnout.xlsx"
End If
Set dic = CreateObject("Scripting.Dictionary")
dic.RemoveAll
Set rep = GetObject("F:\Intrepid Spirits\Budget\Budet Restructure\Replacement\" & wb_name)
Dim last_row As Integer, sht As Worksheet, i As Integer, k
Set sht = rep.Sheets("ReplacementAll")
last_row = sht.Cells(Rows.Count, 1).End(xlUp).Row
'' create dictionary for replacement
For i = 2 To last_row
    dic(sht.Cells(i, 1).Value) = sht.Cells(i, 2).Value
Next

rep.Close


'' start replacing
For Each sht In ThisWorkbook.Sheets
    For Each k In dic.keys()
        sht.Cells.Replace k, dic(k), xlPart
    Next
Next

Set dic = Nothing
End Sub

Sub data_stack()
Dim sht As Worksheet, goal_sheet As Worksheet
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

Cells.Clear
ActiveSheet.Cells(1, 1).Resize(b_row, col_header + 2) = brr


'' FILTER OUT 0
last_col = Cells(1, Columns.Count).End(xlToLeft).Column
header = Range(Cells(1, 1), Cells(1, last_col))
ActiveSheet.UsedRange.AutoFilter Field:=last_col, Criteria1:="<0.5", _
    Operator:=xlOr, Criteria2:="="
Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
Cells(1, 1).EntireRow.Insert
Range(Cells(1, 1), Cells(1, last_col)) = header
End Sub

Sub add_region()
Dim Brand_cell As Range
Dim last_row As Long
last_row = ActiveSheet.UsedRange.Rows.Count
Set Brand_cell = Cells.Find("Brand", searchdirection:=xlPrevious)
Brand_cell.Offset(0, -1).EntireColumn.Insert
Brand_cell.Offset(0, -2).Value = "Region"
''F:\Intrepid Spirits\Budget\Budet Restructure\Replacement\[CountryToRegion.xlsx]Sheet1'!$A:$B
Range(Brand_cell.Offset(1, -2), Cells(last_row, Brand_cell.Column - 2)).FormulaR1C1 = "=VLOOKUP(RC[1],'F:\Intrepid Spirits\Budget\Budet Restructure\Replacement\[CountryToRegion.xlsx]Sheet1'!C1:C2,2,0)"
End Sub
Sub get_case_detail()
Dim Date_cell As Range, Date_rng As Range
Dim Case_Config As Range
Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Set Date_cell = Cells.Find("Date")
Set Date_rng = Range(Date_cell.Offset(1, 0), Cells(last_row, Date_cell.Column))
Date_cell.EntireColumn.Resize(, 2).Insert

Set Case_Config = Cells.Find("Case Config")
Date_cell.Offset(0, -1).Value = "MLPerBottle"
Date_cell.Offset(0, -2).Value = "BottlesPerCase"
Date_rng.Offset(0, -1).FormulaR1C1 = "=iferror(get_ml_per_bottle(RC[" & Case_Config.Column - Date_cell.Column + 1 & "]),0)"
Date_rng.Offset(0, -1).Value = Date_rng.Offset(0, -1).Value

Date_rng.Offset(0, -2).FormulaR1C1 = "=iferror(get_bottles_per_case(RC[" & Case_Config.Column - Date_cell.Column + 2 & "]),0)"
Date_rng.Offset(0, -2).Value = Date_rng.Offset(0, -2).Value


End Sub

Sub prepare_criteria()

Dim Date_rng As Range
Dim last_row As Long
Set Date_rng = Cells.Find("Date")
Date_rng.EntireColumn.NumberFormat = "mmm-yy"
last_row = Cells(Rows.Count, Date_rng.Column).End(xlUp).Row
Dim join_string As String

Date_rng.Offset(0, -2).EntireColumn.Insert
Date_rng.Offset(0, -3).Value = "ABV"

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

Sub map_product_detail()

Dim detail_path As String, Case_cell As Range, last_row As Long
detail_path = "F:\Intrepid Spirits\Budget\Budet Restructure\ProductDetailStructure\ProductDetail\"
Set Case_cell = Cells.Find("Case", lookat:=xlWhole)
last_row = Cells(Rows.Count, Case_cell.Column).End(xlUp).Row
Case_cell.Offset(0, 1).Value = "Price"
'=VLOOKUP(M2,'F:\Intrepid Spirits\Budget\Budet Restructure\ProductDetailStructure\ProductDetail\[PriceData.xlsx]Sheet1'!$I:$K,3,0)
Range(Case_cell.Offset(1, 1), Cells(last_row, Case_cell.Column + 1)).FormulaR1C1 = "=VLOOKUP(RC[-3],'" & detail_path & "[PriceData.xlsx]Sheet1'!C9:C11,3,0)"

Case_cell.Offset(0, 2).Value = "Cost"
Range(Case_cell.Offset(1, 2), Cells(last_row, Case_cell.Column + 2)).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6],'" & detail_path & "[CostDataPivot.xlsx]Sheet1'!C6:C7,2,0),VLOOKUP(RC[-5],'" & detail_path & "[CostDataPivot.xlsx]Sheet1'!C6:C7,2,0))"

Case_cell.Offset(0, 3).Value = "Margin"
Range(Case_cell.Offset(1, 3), Cells(last_row, Case_cell.Column + 3)).FormulaR1C1 = "=RC[-2]-RC[-1]"

'' get ABV
Dim ABV_cell As Range, Variant_cell As Range

Set ABV_cell = Cells.Find("ABV")
Set Variant_cell = Cells.Find("Variant")
Range(ABV_cell.Offset(1, 0), Cells(last_row, ABV_cell.Column)).FormulaR1C1 = "=VLOOKUP(RC[" & Variant_cell.Column - ABV_cell.Column & "],'" & detail_path & "[ABV.xlsx]ABV'!C2:C5,3,0)"


End Sub

Sub calculate_total()
'ThisWorkbook.Sheets("StackedShipmentPlan").Activate
Dim Case_cell As Range, Case_rng As Range, Detail_header As Range, detail_rng As Range
Dim last_row As Long, last_col As Long
Set Case_cell = Cells.Find("Case", lookat:=xlWhole)
last_row = Cells(Rows.Count, Case_cell.Column).End(xlUp).Row
last_col = Cells(1, Columns.Count).End(xlToLeft).Column

Set Case_rng = Range(Case_cell.Offset(1, 0), Cells(last_row, Case_cell.Column))

Set Detail_header = Range(Case_cell.Offset(0, 1), Cells(1, last_col))
Set detail_rng = Range(Case_cell.Offset(1, 1), Cells(last_row, last_col))

Detail_header.Offset(0, Detail_header.Columns.Count).FormulaArray = "=""Total"" & " & Detail_header.Address(ReferenceStyle:=xlR1C1)
detail_rng.Offset(0, Detail_header.Columns.Count).FormulaArray = "=" & detail_rng.Address(ReferenceStyle:=xlR1C1) & "*" & Case_rng.Address(ReferenceStyle:=xlR1C1)
End Sub

Sub nine_l()
'' Create 9L version spreadsheet
Dim preSheet As Worksheet
Set preSheet = ActiveSheet
ActiveSheet.Copy after:=ActiveSheet
ActiveSheet.Name = "9L" & preSheet.Name

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
Case_rng.FormulaArray = "=iferror(" & preSheet.Name & "!" & Case_rng.Address(ReferenceStyle:=xlR1C1) & "*" & ML_rng.Address(ReferenceStyle:=xlR1C1) & "*" & BottlesPerCase_rng.Address(ReferenceStyle:=xlR1C1) & "/9000,0)"
detail_rng.FormulaArray = "=iferror(" & preSheet.Name & "!" & detail_rng.Address(ReferenceStyle:=xlR1C1) & "/(" & ML_rng.Address(ReferenceStyle:=xlR1C1) & "*" & BottlesPerCase_rng.Address(ReferenceStyle:=xlR1C1) & ")" & "*9000,0)"

Cells.Replace "Case", "9LCase"
End Sub
Sub concat_history(Optional history_type As String = "Depletion")
Dim history_folder As String, history_wb_name As String
Dim history As Workbook
history_folder = "F:\Intrepid Spirits\Budget\DataBase\HistoricalData\"
history_wb_name = Dir(history_folder & "*.xls*")
If history_type = "Depletion" Then
    Do While Not history_wb_name Like "*HistoricalDepletion*"
        history_wb_name = Dir
    Loop
ElseIf history_type Like "Price" Then
    Do While Not history_wb_name Like "*HistoricalPrice*"
        history_wb_name = Dir
    Loop
Else
    Do While Not history_wb_name Like "*HistoricalCost*"
        history_wb_name = Dir
    Loop
End If

'Set history = GetObject(history_folder & history_wb_name)
Set history = Workbooks.Open(history_folder & history_wb_name, UpdateLinks:=3)
Dim h_last_row As Long, h_last_col As Long
h_last_row = history.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row
h_last_col = history.Sheets(1).Cells(1, Columns.Count).End(xlToLeft).Column
history.Sheets(1).Range(history.Sheets(1).Cells(2, 1), history.Sheets(1).Cells(h_last_row, h_last_col)).Copy ThisWorkbook.Sheets(1).Cells(ThisWorkbook.Sheets(1).UsedRange.Rows.Count + 1, 1)
history.Close
End Sub

Sub concat_other_depletion(ByVal country_name As String)
Dim other_dep_folder As String, dep_wb_name As String
Dim dep As Workbook
other_dep_folder = "F:\Intrepid Spirits\Budget\DataBase\OtherDepletion\"
dep_wb_name = Dir(other_dep_folder & "*.xls*")
If country_name = "USA" Then
    Do While Not dep_wb_name Like "*USA*Depletion*"
        dep_wb_name = Dir
    Loop
Else
    Do While Not dep_wb_name Like "*IrelandThirdParty*Depletion*"
        dep_wb_name = Dir
    Loop
End If

'Set dep = GetObject(other_dep_folder & dep_wb_name)
Set dep = Workbooks.Open(other_dep_folder & dep_wb_name, 3)
Dim h_last_row As Long, h_last_col As Long
h_last_row = dep.Sheets("Stacked").Cells(Rows.Count, 1).End(xlUp).Row
h_last_col = dep.Sheets("Stacked").Cells(1, Columns.Count).End(xlToLeft).Column
dep.Sheets("Stacked").Range(dep.Sheets("Stacked").Cells(2, 1), dep.Sheets("Stacked").Cells(h_last_row, h_last_col)).Copy ThisWorkbook.Sheets(1).Cells(ThisWorkbook.Sheets(1).UsedRange.Rows.Count + 1, 1)
dep.Close
End Sub


