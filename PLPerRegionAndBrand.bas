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
ActiveSheet.Name = "CompanyName"
ThisWorkbook.Sheets.Add
ActiveSheet.Name = "InternalSalesPerBrand"
ThisWorkbook.Sheets.Add
ActiveSheet.Name = "InternalSalesPerRegion"

Dim sht As Worksheet
For Each sht In Sheets
    If sht.Name <> "SalesDetail" And sht.Name <> "CompanyName" And sht.Name <> "InternalSalesPerBrand" And sht.Name <> "InternalSalesPerRegion" Then
        sht.Delete
    End If
Next

Dim FS_name As Variant, current_FS As Variant, source_wb As Workbook
Dim APAC As Range, Americas As Range, EMEA As Range
FS_name = Application.GetOpenFilename("excels,*xls*", 1, "select current FR", , True)
If IsArray(FS_name) Then
    For Each current_FS In FS_name
        Set source_wb = Workbooks.Open(current_FS)
        source_wb.Sheets(1).Copy after:=ThisWorkbook.Sheets(Sheets.Count)
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
                Sheets("PLperRegion").Copy after:=Sheets("PLperRegion")
                ActiveSheet.Name = "ConsolidatedPLperRegion"
            Else
                ActiveSheet.Name = "PLperBrand"
                Sheets("PLperBrand").Copy after:=Sheets("PLperBrand")
                ActiveSheet.Name = "ConsolidatedPLperBrand"

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
Sheets("CompanyName").Activate
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
Dim fielter_cell As Range, fielter_rng As Range, rng As Range, col As Integer
Set fielter_cell = Cells.Find(brand_or_region)
Set fielter_rng = Range(fielter_cell.Offset(1, 0), Cells(ActiveSheet.UsedRange.Rows.Count, fielter_cell.Column))

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
For Each rng In fielter_rng
    If rng.Value = "" Then
        rng.Value = "- Unassigned -"
    End If
    
    If Cells.Find(rng.Value) Is Nothing Then
        col = col + 1
        Cells(1, col).Value = rng.Value
    End If
Next
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
        ConsolidatedPLper_rng.Interior.Color = vbYellow
        
        Sheets(internal_transaction_sheet).Activate


    Next r
Next col

'' eliminate mkt expense
Dim mkt_exp_cell As Range
Set mkt_exp_cell = find_address("65140 - General Marketing", "- Unassigned -", condolidated_name)
mkt_exp_cell.Value = mkt_exp_cell.Value - internal_fund_transfer
mkt_exp_cell.Interior.Color = vbYellow

End Sub

Function calculate_intenal_transaction(ByVal ledger_name As String, ByVal field_name As String, sumif_criteria As String)
Dim field_rng As Range, amount_rng As Range

'' use sumif to calculate the internal transaction to save time

Set field_rng = Sheets(ledger_name).Cells.Find(sumif_criteria, SearchFormat:=False, lookat:=xlPart).EntireColumn
Set amount_rng = Sheets(ledger_name).Cells.Find("Amount", SearchFormat:=False, lookat:=xlPart).EntireColumn
calculate_intenal_transaction = Application.WorksheetFunction.SumIf(field_rng, field_name, amount_rng)

End Function


Function find_address(ByVal account As String, field_name As String, find_in_sheet As String)
Sheets(find_in_sheet).Activate
Dim row_num As Integer, col_num As Integer
row_num = Cells.Find(account, SearchFormat:=False).Row
col_num = Cells.Find(field_name, SearchFormat:=False).Column
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
    Sheets(FS_order(i)).Move after:=Sheets(i + 1)
Next i
Sheets("ConsolidatedPLperBrand").Select
End Sub

Sub export()
Dim save_path As String
save_path = ThisWorkbook.Path & "\" & Sheets("ConsolidatedPLperSub").Range("A4").Value & "PLPerRegionAndBrand.xlsm"
ThisWorkbook.SaveCopyAs save_path
End Sub



