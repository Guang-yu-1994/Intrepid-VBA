Attribute VB_Name = "PLPerRegionAndBrand"
Option Explicit
Public internal_fund_transfer As Double
Sub execute()
Call import_PL
Call get_company_names
Call filter_internal("SalesDetail")
Call filter_internal("COSDetail")
Call get_internal_fund_transfer
Call rearrange_InternalSalesPer("Brand")
Call rearrange_InternalSalesPer("Region")
Call sum_internal_transaction("InternalSalesPerBrand")
Call sum_internal_transaction("InternalSalesPerRegion")
Call rearrange_FS
Call export
End Sub

Sub import_PL()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
On Error Resume Next
ThisWorkbook.Sheets("SalesDetail").Delete
ThisWorkbook.Sheets("CompanyName").Delete
ThisWorkbook.Sheets("InternalSalesPerBrand").Delete
ThisWorkbook.Sheets("InternalSalesPerRegion").Delete


ThisWorkbook.Sheets.Add
ActiveSheet.Name = "InternalSalesPerBrand"
ThisWorkbook.Sheets.Add
ActiveSheet.Name = "InternalSalesPerRegion"


Dim sht As Worksheet
For Each sht In Sheets
    If sht.Name <> "SalesDetail" And sht.Name <> "CompanyName" And sht.Name <> "InternalSalesPerBrand" And sht.Name <> "InternalSalesPerRegion" And sht.Name <> "Instruction" Then
        sht.Delete
    End If
Next

Dim FS_name As Variant, current_FS As Variant, source_wb As Workbook
Dim APAC As Range, Americas As Range, EMEA As Range
FS_name = Application.GetOpenFilename("excels,*xls*", 1, "select current FR", , True)
If IsArray(FS_name) Then
    For Each current_FS In FS_name
        Set source_wb = Workbooks.Open(current_FS)
        source_wb.Sheets(1).Copy After:=ThisWorkbook.Sheets(Sheets.Count)
        If source_wb.Sheets(1).Name Like "*SalesPerRegionAndBrand*" Then
            ActiveSheet.Name = "SalesDetail"
            Cells.Replace "Intrepid Spirits Ireland Limited", "Intrepid Spirits Ireland Ltd." 'keep the name the same as PL
        ElseIf source_wb.Sheets(1).Name Like "*COSPerRegionAndBrand*" Then
            ActiveSheet.Name = "COSDetail"
            Cells.Replace "Intrepid Spirits Ireland Limited", "Intrepid Spirits Ireland Ltd." 'keep the name the same as PL
        ElseIf ActiveSheet.Name Like "*PLperSub*" Then
            ActiveSheet.Name = "ConsolidatedPLperSub"
        ElseIf source_wb.Name Like "*ProfitandLoss*" Then
            Set APAC = ActiveSheet.Cells.Find("APAC", lookat:=xlPart)
            Set Americas = ActiveSheet.Cells.Find("Americas", lookat:=xlPart)
            Set EMEA = ActiveSheet.Cells.Find("EMEA", lookat:=xlPart)

            If Not APAC Is Nothing Or Not Americas Is Nothing Or Not EMEA Is Nothing Then
                ActiveSheet.Name = "PLperRegion"
                Sheets("PLperRegion").Copy After:=Sheets("PLperRegion")
                ActiveSheet.Name = "ConsolidatedPLperRegion"
                Call reclassify_shippingAndHandling
            Else
                ActiveSheet.Name = "PLperBrand"
                Sheets("PLperBrand").Copy After:=Sheets("PLperBrand")
                ActiveSheet.Name = "ConsolidatedPLperBrand"
                Call reclassify_shippingAndHandling
            End If
        End If
        source_wb.Close
    Next
Else
    End
End If
End Sub
Sub get_company_names()
Sheets("ConsolidatedPLperSub").Activate
Dim PL_header As Range, IncomeOrExpense As Range, crng As Range, i As Integer
Dim non_company, non_company_str As String
Set IncomeOrExpense = Cells.Find("Ordinary Income/Expense")
Set PL_header = Range(IncomeOrExpense.Offset(-4, 0), Cells(IncomeOrExpense.Row - 1, ActiveSheet.UsedRange.Columns.Count))

'' get company name from PL
ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
ActiveSheet.Name = "CompanyName"
Cells(1, 1).Value = "Name"
non_company = Array(" ", "Parent Company", "Amount", "Total", "Adjustment")
non_company_str = Join(non_company, "|")
i = 2
For Each crng In PL_header
    If crng.Value <> "" And VBA.InStr(non_company_str, crng.Value) <= 0 Then
        Cells(i, 1).Value = crng.Value
        i = i + 1
    End If
Next

'' delete empty rows
Dim delete_rng As Range
For i = 1 To ActiveSheet.UsedRange.Rows.Count
    If Len(Cells(i, 1).Value) < 2 Then
        If delete_rng Is Nothing Then
            Set delete_rng = Cells(i, 1)
        Else
            Set delete_rng = Union(delete_rng, Cells(i, 1))
        End If
    End If
Next
delete_rng.EntireRow.Delete

End Sub
Sub filter_internal(ByVal ledger_name As String)
Sheets(ledger_name).Activate

Dim financial As Range, original_rng As Range
On Error Resume Next

' filter the internal transactions
Set financial = Cells.Find("Financial Row")
Range(Cells(1, 1), financial.Offset(-1, 0)).EntireRow.Delete
Set original_rng = Sheets(ledger_name).UsedRange
Application.CutCopyMode = False
original_rng.AdvancedFilter Action:=xlFilterInPlace, CriteriaRange:=Sheets("CompanyName").UsedRange, Unique:=False
Application.CutCopyMode = True

original_rng.Copy Cells(1, Sheets(ledger_name).UsedRange.Columns.Count + 1)
Cells.AutoFilter

original_rng.EntireColumn.Delete
End Sub
Sub get_internal_fund_transfer()
Sheets("ConsolidatedPLperSub").Activate
internal_fund_transfer = find_address("40010 - Sales", "Intrepid Spirits USA Inc", "ConsolidatedPLperSub").Value
End Sub


Sub rearrange_InternalSalesPer(ByVal brand_or_region As String)
'' the sales without brand means sales not related to inventory sale
Sheets("SalesDetail").Activate

Dim has_internal_transaction As Boolean
If ActiveSheet.UsedRange.Rows.Count = 1 Then
    has_internal_transaction = False
Else
    has_internal_transaction = True
End If

Dim filter_cell As Range, filter_rng As Range, rng As Range, col As Integer
Set filter_cell = Cells.Find(brand_or_region)
'' here cannot use endxl up because sometimes brand and regions are just empty, use usedrange.rows.count
Set filter_rng = Range(filter_cell.Offset(1, 0), Cells(ActiveSheet.UsedRange.Rows.Count, filter_cell.Column))

'' get the fileds name from internal sale
Dim internal_transaction_sheet_name As String
If brand_or_region Like "*Brand*" Then
    internal_transaction_sheet_name = "InternalSalesPerBrand"
Else
    internal_transaction_sheet_name = "InternalSalesPerRegion"
End If

Sheets(internal_transaction_sheet_name).Activate
Cells(1, 1).Value = "Account"
Cells(2, 1).Value = "40010 - Sales"
Cells(3, 1).Value = "50010 - Cost of Goods Sold"

col = 1
'' split no internal transaction and has internal transaction
If has_internal_transaction Then
    For Each rng In filter_rng
        If rng.Value = "" Then
            rng.Value = "- Unassigned -"
        End If
        
        If Cells.Find(rng.Value) Is Nothing Then
            col = col + 1
            Cells(1, col).Value = rng.Value
        End If
    Next
End If
End Sub
Sub sum_internal_transaction(ByVal internal_transaction_sheet As String)
Dim original_sht_name As String, condolidated_name As String, sumif_criteria As String
Sheets(internal_transaction_sheet).Activate
If internal_transaction_sheet Like "*Brand*" Then
    original_sht_name = "PLperBrand"
    condolidated_name = "ConsolidatedPLperBrand"
    sumif_criteria = "Brand"
Else
    original_sht_name = "PLperRegion"
    condolidated_name = "ConsolidatedPLperRegion"
    sumif_criteria = "region (GL)"
End If

Dim col As Integer, r As Integer, ledger_name As String
Dim ConsolidatedPLper_rng As Range
'Dim PLperBrand As Range
Dim offset_row As Integer, offset_col As Integer

For col = 2 To ActiveSheet.UsedRange.Columns.Count
    For r = 2 To ActiveSheet.UsedRange.Rows.Count
        If Cells(r, 1).Value Like "*Sales*" Then
            ledger_name = "SalesDetail"
        Else
            ledger_name = "COSDetail"
        End If
        
        Cells(r, col).Value = calculate_intenal_transaction(ledger_name, Cells(1, col).Value, sumif_criteria)

        Set ConsolidatedPLper_rng = find_address(account:=Cells(r, 1).Value, field_name:=Cells(1, col).Value, find_in_sheet:=condolidated_name)
'        Set PLperfield_rng = find_address(account:=Cells(r, 1).Value, field_name:=Cells(1, col).Value, find_in_sheet:="PLperBrand")
        offset_row = ConsolidatedPLper_rng.Row - r
        offset_col = ConsolidatedPLper_rng.Column - col
        
        ConsolidatedPLper_rng.FormulaR1C1 = "=" & original_sht_name & "!RC-" & internal_transaction_sheet & "!R[" & -offset_row & "]C[" & -offset_col & "]"
        ConsolidatedPLper_rng.Interior.color = vbYellow
        
        Sheets(internal_transaction_sheet).Activate


    Next r
Next col

'' eliminate mkt expense
Dim mkt_exp_cell As Range
Set mkt_exp_cell = find_address("65140 - General Marketing", "- Unassigned -", condolidated_name)
mkt_exp_cell.Value = mkt_exp_cell.Value - internal_fund_transfer
mkt_exp_cell.Interior.color = vbYellow

End Sub

Function calculate_intenal_transaction(ByVal ledger_name As String, ByVal field_name As String, sumif_criteria As String)
Dim field_rng As Range, amount_rng As Range

'' use sumif to calculate the internal transaction to save time

Set field_rng = Sheets(ledger_name).Cells.Find(sumif_criteria, searchformat:=False, lookat:=xlPart).EntireColumn
Set amount_rng = Sheets(ledger_name).Cells.Find("Amount", searchformat:=False, lookat:=xlPart).EntireColumn
calculate_intenal_transaction = Application.WorksheetFunction.SumIf(field_rng, field_name, amount_rng)

End Function


Function find_address(ByVal account As String, field_name As String, find_in_sheet As String)
Sheets(find_in_sheet).Activate
Dim row_num As Integer, col_num As Integer
row_num = Cells.Find(account, searchformat:=False).Row
col_num = Cells.Find(field_name, searchformat:=False).Column
Set find_address = ActiveSheet.Cells(row_num, col_num)
End Function

Sub rearrange_FS()
'sort the order of sheets
Dim FS_order, i As Integer
'Sheets("ConsolidatedPLperBrand").Cells.Replace "- Unassigned -", "InternalFundsTransfer"
'Sheets("InternalSalesPerBrand").Cells.Replace "- Unassigned -", "InternalFundsTransfer"
'Sheets("PLperBrand").Cells.Replace "- Unassigned -", "InternalFundsTransfer"

FS_order = Array("ConsolidatedPLperSub", "ConsolidatedPLperBrand", "ConsolidatedPLperRegion", "InternalSalesPerBrand", "InternalSalesPerRegion", "PLperBrand", "PLperRegion", "CompanyName", "SalesDetail", "COSDetail")
For i = LBound(FS_order) To UBound(FS_order)
    Sheets(FS_order(i)).Cells.EntireColumn.AutoFit
    Sheets(FS_order(i)).Move After:=Sheets(i + 1)
Next i
Sheets("ConsolidatedPLperBrand").Select
End Sub

Sub export()
Dim save_path As String
StartConsolidating.Hide
ThisWorkbook.Sheets(1).Activate
Dim export_date, date_str As String
date_str = Sheets("ConsolidatedPLperSub").Range("A4").Value

export_date = extract_wb_date(date_str)
save_path = ThisWorkbook.Path & "\" & export_date & "PLPerRegionAndBrand.xlsm"
ThisWorkbook.SaveCopyAs save_path
MsgBox "Consolidation Completed"

End Sub

Function extract_wb_date(date_str As String)
Dim reg As Object, match
Set reg = CreateObject("vbscript.regexp")
reg.Global = True
    
If date_str Like "*Q*" Then
    reg.Pattern = "Q(\d{1})\s(\d{4})"
    Set match = reg.execute(date_str)
    extract_wb_date = match(0).submatches(1) & match(0).submatches(0) * 3 - 2 & "-" & match(0).submatches(1) & match(0).submatches(0) * 3
ElseIf Len(date_str) <= 8 Then '' Jan yyyy
    extract_wb_date = Format(CDate(date_str), "yyyymm")
ElseIf date_str Like "*to*" Then '' From Jan yyyy to Apr yyyy
    date_str = Trim(Replace(date_str, "From ", ""))
    date_str = Replace(date_str, " to ", "-")
    extract_wb_date = extract_wb_date(CStr(Split(date_str, "-")(0))) & "-" & extract_wb_date(CStr(Split(date_str, "-")(1)))
End If

Set reg = Nothing
End Function

Sub reclassify_shippingAndHandling()
Dim shippingAndHandling As Range, COS As Range
Dim last_col As Integer, arr, a
Set shippingAndHandling = Cells.Find("40050 - Shipping and Handling")
'' no shippingAndHandling then no need to adjust
If shippingAndHandling Is Nothing Then GoTo no_shippingAndHandling

last_col = Cells(shippingAndHandling.Row, Columns.Count).End(xlToLeft).Column
arr = Range(shippingAndHandling.Offset(0, 1), Cells(shippingAndHandling.Row, last_col))
'' from sale to cost, the number should be reversed
For a = LBound(arr, 2) To UBound(arr, 2)
    arr(1, a) = -arr(1, a)
Next

'' write the reversed amounts for shipping and handling to the above row of 50010 - Cost of Goods Sold
Set COS = Cells.Find("50010 - Cost of Goods Sold")
COS.EntireRow.Insert
COS.Offset(-1, 0).Value = "40050 - Shipping and Handling"
Range(COS.Offset(-1, 1), Cells(COS.Row - 1, last_col)) = arr

'' keep the same format

Range(COS.Offset(-1, 0), Cells(COS.Row - 1, last_col)).Font.Bold = False

'' delete the original shipping and handling in the sales subitems
shippingAndHandling.EntireRow.Delete

'' reset the functions in the COS by adding the shipping and handling
Dim total_COS As Range
Set total_COS = Cells.Find("Total - 50000 - - COST OF GOODS SOLD -")
Range(total_COS.Offset(0, 1), Cells(total_COS.Row, last_col)).FormulaR1C1 = total_COS.Offset(0, 1).FormulaR1C1 & "+R[" & COS.Row - 1 - total_COS.Row & "]C"

no_shippingAndHandling:
End Sub





