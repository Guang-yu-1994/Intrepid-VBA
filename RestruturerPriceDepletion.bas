Attribute VB_Name = "RestruturerPriceDepletion"
Option Explicit
Public source_type As String
Sub execute()
Call back_up
Call data_conbine
Call data_stack
Call prepare_criteria
Call remove_null
Call fit_columns
'Call split_case_config
End Sub
Sub data_conbine()
Dim sht As Worksheet, goal_sheet As Worksheet
For Each sht In Sheets
    If sht.Name Like "*Data*" Then
        source_type = Left(sht.Name, Len(sht.Name) - Len("Data"))
        If source_type <> "Cost" Then GoTo Exit_sub:

        Set goal_sheet = sht
        goal_sheet.Cells.Clear
    End If
Next

'start combining data
Dim get_header As Boolean
get_header = True
For Each sht In Sheets
    If Not sht.Name Like "*Data*" And Not sht.Name Like "*Summary*" And Not sht.Name Like "*Pivot*" Then
        If get_header Then
            sht.UsedRange.Copy goal_sheet.Cells(1, 1)
            get_header = False
        Else
            sht.Rows("2:" & sht.UsedRange.Rows(sht.UsedRange.Rows.Count).Row).Copy goal_sheet.Rows(goal_sheet.UsedRange.Rows.Count + 1)
        End If
    End If
Next
Exit_sub:
End Sub

Sub data_stack()
Dim sht As Worksheet, goal_sheet As Worksheet
For Each sht In Sheets
    If sht.Name Like "*Data*" Then
        Set goal_sheet = sht
        source_type = Left(sht.Name, Len(sht.Name) - Len("Data"))
    End If
Next
goal_sheet.Activate
Dim arr, brr()
Dim a_row As Long, a_col As Long, b_row As Long, b_col As Long
Dim col_header As Long, i As Long


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
If Not source_type Like "Price" And Not source_type Like "Cost" Then
    source_type = "Case"
End If
brr(1, col_header + 2) = source_type

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

Sheets("Summary").Activate
Sheets("Summary").Cells.Clear
ActiveSheet.Cells(1, 1).Resize(b_row, col_header + 2) = brr
End Sub
Sub prepare_criteria()

Dim Date_rng As Range
Dim last_row As Long
Set Date_rng = Cells.Find("Date")
Date_rng.EntireColumn.NumberFormat = "mmm-yy"
last_row = Cells(Rows.Count, Date_rng.Column).End(xlUp).Row
'Date_rng.EntireColumn.Resize(, 2).Insert
'Date_rng.Offset(0, -1).Value = "Year"
'Range(Date_rng.Offset(1, -1), Cells(last_row, Date_rng.Column - 1)).Value = Year(Now())

Date_rng.EntireColumn.Insert


Dim join_string As String

If source_type = "Cost" Then
    Date_rng.Offset(0, -1).Value = "CriteriaForCost"
    join_string = "RC[-8]&RC[-7]&RC[-6]&RC[-5]&RC[-4]&RC[-3]"
'ElseIf source_type = "Price" Then
'    join_string = "RC[-8]&RC[-7]&RC[-6]&RC[-5]&RC[-4]"
Else
    Date_rng.Offset(0, -1).Value = "CriteriaForPrice"
    join_string = "RC[-8]&RC[-7]&RC[-6]&RC[-5]&RC[-4]"
End If
Range(Date_rng.Offset(1, -1), Cells(last_row, Date_rng.Column - 1)).FormulaR1C1 = "=SUBSTITUTE(TEXT(RC[1],""MMM/YYYY"")&" & join_string & ","" "","""")"


'' Some cost items has GEXP, so it is necessary to create another criteria for this situation

If Not source_type Like "Cost" And Not source_type Like "Price" Then
    Date_rng.EntireColumn.Insert
    Date_rng.Offset(0, -1).Value = "CriteriaForCost1"
    join_string = "RC[-9]&RC[-8]&RC[-7]&RC[-6]"
    Range(Date_rng.Offset(1, -1), Cells(last_row, Date_rng.Column - 1)).FormulaR1C1 = "=SUBSTITUTE(TEXT(RC[1],""MMM/YYYY"")&" & join_string & ","" "","""")"
    
    Date_rng.EntireColumn.Insert
    Date_rng.Offset(0, -1).Value = "CriteriaForCost2"
    join_string = """GEXP""&RC[-9]&RC[-8]&RC[-7]"
    Range(Date_rng.Offset(1, -1), Cells(last_row, Date_rng.Column - 1)).FormulaR1C1 = "=SUBSTITUTE(TEXT(RC[1],""MMM/YYYY"")&" & join_string & ","" "","""")"
End If



End Sub
Sub fit_columns()
Dim sht As Worksheet
For Each sht In Sheets
    sht.Columns.AutoFit
Next
End Sub
Sub split_case_config()
Dim Case_Config As Range, Case_Config_rng As Range
Dim last_row As Long
last_row = Cells(Rows.Count, 1).End(xlUp).Row
Set Case_Config = Cells.Find("Case Config")
Set Case_Config_rng = Range(Case_Config.Offset(1, 0), Cells(last_row, Case_Config.Column))
Case_Config.Offset(0, 1).Resize(, 2).EntireColumn.Insert
Case_Config.Offset(0, 1).Value = "BottlesPerCase"
'Range(Case_Config.Offset(1, 1), Cells(last_row, Case_Config.Column + 1)).FormulaR1C1 = "=get_bottles_per_case(RC[-1])"
'Range(Case_Config.Offset(1, 1), Cells(last_row, Case_Config.Column + 1)).FormulaR1C1 = "=get_bottles_per_case(" & Case_Config_rng.Address(ReferenceStyle:=xlR1C1) & ")"
Case_Config_rng.Offset(0, 1).Value = get_bottles_per_case(Case_Config_rng)
Case_Config.Offset(0, 2).Value = "MLPerBottle"
'Range(Case_Config.Offset(1, 2), Cells(last_row, Case_Config.Column + 2)).FormulaR1C1 = "=get_ml_per_bottle(RC[-2])"
'Range(Case_Config.Offset(1, 2), Cells(last_row, Case_Config.Column + 2)).FormulaR1C1 = "=get_ml_per_bottle(" & Case_Config_rng.Address(ReferenceStyle:=xlR1C1) & ")"
Case_Config_rng.Offset(0, 2).Value = get_ml_per_bottle(Case_Config_rng)
End Sub


Sub remove_null()
If Not source_type Like "Cost" And Not source_type Like "Price" Then
    Dim Date_rng As Range
    Dim header_arr
    Dim criteria_col As Long
    Set Date_rng = Cells.Find("date")
    criteria_col = Date_rng.Column + 1
    header_arr = Cells(1, 1).CurrentRegion.Rows(1)
    
    ActiveSheet.UsedRange.AutoFilter Field:=criteria_col, Criteria1:="<0.5", _
            Operator:=xlOr, Criteria2:="="
    Cells.SpecialCells(xlCellTypeVisible).EntireRow.Delete
    
    Cells(1, 1).EntireRow.Insert
    Cells(1, 1).Resize(, criteria_col) = header_arr
End If
End Sub
Sub trim_rng()
''Trim every cell
Dim sht As Worksheet, rng As Range
For Each sht In Sheets
    For Each rng In sht.UsedRange
        rng.Value = Trim(rng.Value)
    Next
Next
End Sub
Sub back_up()
Dim back_up_path As String
back_up_path = "F:\Intrepid Spirits\Budget\Budet Restructure\BackUp\"
ThisWorkbook.SaveCopyAs back_up_path & ThisWorkbook.Name
End Sub

