Attribute VB_Name = "Consolidator"
         Option Explicit
Public header_color As Double
Sub execute()
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Call import
Call restructure_FS("BS")
Call restructure_FS("PL")
Call calculate_total_FS("BS")
Call calculate_total_FS("PL")
Call internal_sales_sum
Call adjust_CurrentPL
Call goodwill_summary
Call adjust_CurrentBS
Call inventory
Call goodwill_adjustment
Call reformat
Call export

MsgBox "Consolidation Complete"

Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub

Sub import()

Dim sht As Worksheet
Sheets(1).Name = "Instructions"
Call delete_sheet
Sheets("Instructions").Cells.Clear

' import original BS and PL
Dim source_wb As Workbook, i As Byte
Dim FS_name, current_FS, FS_order

On Error GoTo 0
FS_name = Application.GetOpenFilename("excels,*xls*", 1, "select current FR", , True)
If IsArray(FS_name) Then
    For Each current_FS In FS_name
        Set source_wb = Workbooks.Open(current_FS)
        source_wb.Sheets(1).Copy After:=ThisWorkbook.Sheets(Sheets.Count)
        If ActiveSheet.Range("A3").Value Like "*Balance Sheet*" Then
            ActiveSheet.Name = "BS"
        ElseIf ActiveSheet.Range("A3").Value Like "*Profit and Loss" Then
            ActiveSheet.Name = "PL"
        ElseIf ActiveSheet.Range("A9").Value Like "*Cost Of Sales*" Then
            ActiveSheet.Name = "Cost of Goods Sold"
            Cells.Replace "Intrepid Spirits Ireland Limited", "Intrepid Spirits Ireland Ltd." ' keep the name the same as BS PL
        ElseIf ActiveSheet.Range("A9").Value Like "*Income*" Then
            ActiveSheet.Name = "Sales"
            Cells.Replace "Intrepid Spirits Ireland Limited", "Intrepid Spirits Ireland Ltd." ' keep the name the same as BS PL
        ElseIf ActiveSheet.Range("A3").Value Like "*Inventory Valuation Summary*" Then
            ActiveSheet.Name = "Regal Rogue"
        End If
        source_wb.Close
    Next



    ' insert Goodwill row in BS
    Sheets("BS").Activate
    Dim total_fixed_asset As Range, used_col As Integer
    Set total_fixed_asset = Cells.Find("Total Fixed Assets", searchformat:=False)
    total_fixed_asset.EntireRow.Insert
    total_fixed_asset.Offset(-1, 0).Value = "Goodwill"
    used_col = Cells(total_fixed_asset.Row, Columns.Count).End(xlToLeft).Column
    Range(Cells(total_fixed_asset.Row, 2), Cells(total_fixed_asset.Row, used_col)).FormulaR1C1 = Cells(total_fixed_asset.Row, 2).FormulaR1C1 & "+ R[-1]C"

    ' insert Goodwill impairment in PL
    Sheets("PL").Activate
    Dim total_expense As Range
    Set total_expense = Cells.Find("Total - Expense", searchformat:=False, lookat:=xlPart)
    total_expense.EntireRow.Insert
    total_expense.Offset(-1, 0).Value = "Goodwill Impairment"
    used_col = Cells(total_expense.Row, Columns.Count).End(xlToLeft).Column
    Range(Cells(total_expense.Row, 2), Cells(total_expense.Row, used_col)).FormulaR1C1 = Cells(total_expense.Row, 2).FormulaR1C1 & "+ R[-1]C"
Else
    End
End If

' Create IntertalSales
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = "InternalSales"
With Sheets("InternalSales")
    .Cells(1, 1) = "CompanyName"
    .Cells(2, 1) = "40010 - Sales"
    .Cells(3, 1) = "50010 - Cost of Goods Sold"
    .Cells(4, 1) = "InventoryBalance"
    .Cells(5, 1) = "Margin"
    .Cells(6, 1) = "14020 - Finished Goods Inventory"
'    .Cells(6, 1).Comment = "this is the amount to be cancelled"
    .Cells(7, 1) = "Retained Earnings"
End With

Dim BS_header As Range, finance_row As Range, assets As Range, crng As Range
Dim non_company, non_company_str As String
Sheets("BS").Activate
Set finance_row = Cells.Find("Financial Row")
header_color = finance_row.Interior.color

Set assets = Cells.Find("ASSETS")
Set BS_header = Range(finance_row, Cells(assets.Row - 1, ActiveSheet.UsedRange.Columns.Count))

Sheets("InternalSales").Activate
non_company = Array(" ", "Financial Row", "Parent Company", "Amount", "Total")
non_company_str = Join(non_company, "|")
For Each crng In BS_header
    If crng.Value <> "" And VBA.InStr(non_company_str, crng.Value) <= 0 Then
        Cells(1, Columns.Count).End(xlToLeft).Offset(0, 1).Value = crng.Value
    End If
Next



' Create Goodwill sheet
Sheets.Add After:=Sheets(Sheets.Count)
ActiveSheet.Name = "Goodwill"
With Sheets("Goodwill")
    .Cells(1, 1) = "Goodwill Before Impairment"
    .Cells(2, 1) = "Goodwill Previously Impaired By"
    .Cells(3, 1) = "Goodwill Currently Impaired By"
    .Cells(4, 1) = "Goodwill After Impairment"
End With

'sort the order of sheets
FS_order = Array("Instructions", "BS", "PL", "Goodwill", "InternalSales", "Cost of Goods Sold", "Sales", "Regal Rogue")
For i = LBound(FS_order) To UBound(FS_order)
    Sheets(FS_order(i)).Move After:=Sheets(i + 1)
Next

End Sub

Sub delete_sheet()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim sht As Worksheet
On Error Resume Next
For Each sht In ThisWorkbook.Sheets
    If sht.Name <> "Instructions" Then sht.Delete
Next
End Sub
Sub restructure_FS(ByVal sheet_name As String)
ThisWorkbook.Sheets(sheet_name).Activate

' insert adjustment column
Dim header_row_up As Integer, header_row_down As Integer
Dim last_cell As Range, header As Range, i As Range, rng As Range
Dim last_col As Integer, last_row As Integer
Dim total_address As String
Dim adjustment_col() As Integer

header_row_up = Cells.Find(What:="Financial Row", lookat:=xlPart).Row + 1
header_row_down = Cells.Find(What:="Financial Row", lookat:=xlPart).Row + 4
Set last_cell = Cells.Find(What:="*", After:=Cells(1, 1), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
last_col = last_cell.Column
Set header = Range(Cells(header_row_up, 1), Cells(header_row_down, last_col))

' delete parent col, it is duplicated and the same as holding
last_cell.EntireColumn.Delete
If sheet_name = "BS" Then
    Cells(1, 2).EntireColumn.Delete
End If

header.Borders.LineStyle = xlContinuous
For Each i In header
    If i.Value = "Total" Then
        total_address = total_address & i.Address & ","
    End If
Next
total_address = Left(total_address, Len(total_address) - 1)
Range(total_address).EntireColumn.Insert

Set last_cell = Cells.Find(What:="*", After:=Cells(1, 1), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
last_col = last_cell.Column
last_row = last_cell.Row
For Each i In Range(Cells(header_row_down, 2), Cells(header_row_down, last_col))
    If Not i.MergeCells And i.Value = "" Then
        i.Value = "Adjustment"
        Range(i, Cells(last_row, i.Column)).Interior.color = vbYellow
        Range(Cells(header_row_down + 1, 2), Cells(last_row, 2)).Copy
        i.Offset(1, 0).PasteSpecial xlPasteFormulas
        For Each rng In Range(i.Offset(1, 0), Cells(last_row, i.Column))
            If Not rng.HasFormula Then
                rng.ClearContents
            End If
        Next
    End If
Next



Range(Rows(last_row + 1), Rows(Rows.Count)).Clear

' clear parent row
Cells(header_row_up - 1, 1).EntireRow.Delete

End Sub
Sub calculate_total_FS(ByVal sheet_name As String)

'total for each
Dim head_row As Integer, sum_col As Integer, account As Integer
Dim merged_company As Range, last_cell As Range
Dim last_col As Integer, last_row As Integer, start_account_row As Integer
Dim dict_num As Integer, c As Integer, merged_company_cols As Integer
Dim subtotal_dict As Object
Dim i As Variant, sum_string As String

ThisWorkbook.Sheets(sheet_name).Activate
Application.FindFormat.Interior.color = vbYellow
head_row = Cells.Find("Adjustment", SearchOrder:=xlByRows, SearchDirection:=xlNext, searchformat:=True).Row
sum_col = 5
Set last_cell = Cells.Find(What:="*", After:=Cells(1, 1), SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, searchformat:=False)
last_col = last_cell.Column
last_row = last_cell.Row
start_account_row = head_row + 3

Set subtotal_dict = CreateObject("Scripting.Dictionary")

dict_num = 0
For c = 2 To last_col
    If Cells(head_row, c).Value = "Adjustment" Then
        ' use subtotal address to record subtotals
        subtotal_dict(dict_num) = c + 1
        dict_num = dict_num + 1
'    ElseIf Cells(head_row - 2, c).Value = "Cocalero International HK Limited" Then
'        subtotal_dict(dict_num) = c
'        dict_num = dict_num + 1
        'see the how many companies are merged
        Set merged_company = Cells(head_row, c).Offset(-2, 0)
        merged_company_cols = merged_company.MergeArea.Columns.Count
        Range(Cells(start_account_row, c + 1), Cells(last_row, c + 1)).Formula2R1C1 = "=sum(RC[" & -merged_company_cols + 1 & "]:RC[-1])"
    End If
Next
' remove last item in dict
subtotal_dict.Remove (subtotal_dict.Count - 1)


For Each i In subtotal_dict.items()
    sum_string = sum_string & "RC[-" & last_col - i & "],"
Next
sum_string = "RC[" & -last_col + 2 & "]," & sum_string & "RC[-1]"
Range(Cells(start_account_row, last_col), Cells(last_row, last_col)).FormulaR1C1 = "=sum(" & sum_string & ")"
'' if BS should sum cocalero HK
Dim HK_offset_col As Integer, cocalero_HK As Range

Set cocalero_HK = Cells.Find("Cocalero International HK Limited", lookat:=xlPart, searchformat:=False)
If Not cocalero_HK Is Nothing Then
    HK_offset_col = last_col - cocalero_HK.Column
    Range(Cells(start_account_row, last_col), Cells(last_row, last_col)).FormulaR1C1 = Cells(start_account_row, last_col).FormulaR1C1 & "+RC[" & -HK_offset_col & "]"
End If
End Sub
Sub internal_sales_sum()


Dim margin_row As Integer, ISL_col As Integer, JP_col As Integer, JP_cell As Range, ISL_cell As Range
Dim IDL_col As Integer, IDL_cell As Range
Sheets("InternalSales").Activate
margin_row = Cells.Find("Margin", searchformat:=False, lookat:=xlPart).Row

Set ISL_cell = Cells.Find("Intrepid Spirits Limited", searchformat:=False, lookat:=xlPart)
ISL_col = ISL_cell.Column
Cells(margin_row, ISL_col).Value = 0.2



Set JP_cell = Cells.Find("Intrepid Japan", searchformat:=False, lookat:=xlPart)
If JP_cell Is Nothing Then GoTo no_JP
JP_col = JP_cell.Column
Cells(margin_row, JP_col).Value = 0.436665909879517
no_JP:

' add Intrepid Spirits Ireland Ltd if it dose not exist in the internalSales
Set IDL_cell = Cells.Find("Intrepid Spirits Ireland Ltd", searchformat:=False, lookat:=xlPart)
If IDL_cell Is Nothing Then GoTo no_IDL
IDL_col = IDL_cell.Column
Cells(margin_row, IDL_col).Value = 0
no_IDL:

' set the formula of inventory = balance * margin
Dim finish_good_row As Integer, InternalSales_col As Integer, RE_row As Integer
finish_good_row = Cells.Find("14020 - Finished Goods Inventory", searchformat:=False, lookat:=xlPart).Row
InternalSales_col = Cells(1, Columns.Count).End(xlToLeft).Column
Range(Cells(finish_good_row, 2), Cells(finish_good_row, InternalSales_col)).FormulaR1C1 = "=R[-1]C * R[-2]C"
RE_row = Cells.Find("Retained Earnings", searchformat:=False, lookat:=xlPart).Row
Range(Cells(RE_row, 2), Cells(RE_row, InternalSales_col)).FormulaR1C1 = "=R[-1]C"

'' delete empty cols
Dim i As Integer, delete_rng As Range
For i = 1 To ActiveSheet.UsedRange.Columns.Count
    If Len(Cells(1, i).Value) < 2 Then
        If delete_rng Is Nothing Then
            Set delete_rng = Cells(1, i)
        Else
            Set delete_rng = Union(delete_rng, Cells(1, i))
        End If
    End If
Next

delete_rng.EntireColumn.Delete

Call calculate_intenal_transaction("Cost of Goods Sold")
Call calculate_intenal_transaction("Sales")
End Sub

Sub calculate_intenal_transaction(ByVal ledger_name As String)
Dim account_row As Integer, crng As Range, name_rng As Range
Dim source_rng As Range, company_rng As Range
Dim start_row As Integer, end_row As Integer, end_col As Integer
Dim header As Range, name_cell As Range, nrng As Range, irng As Range
Dim internal_header As Range
Dim result As Variant



'' use sumif to calculate the internal transaction to save time
Sheets(ledger_name).Activate
Set name_rng = Cells.Find("Name", searchformat:=False, lookat:=xlPart).EntireColumn
Sheets("InternalSales").Activate
If Cells(1, 2).Value <> "" Then
    account_row = Cells.Find(ledger_name, searchformat:=False, lookat:=xlPart).Row
        Set company_rng = Range(Cells(1, 2), Cells(1, Columns.Count).End(xlToLeft))
        For Each crng In company_rng
            Cells(account_row, crng.Column) = Application.WorksheetFunction.SumIf(name_rng, crng.Value, name_rng.Offset(0, 3))
        Next
End If
End Sub


Sub adjust_CurrentPL()
Dim company_name As String, account_name As String
Dim arng As Range, account_rng As Range, company_rng As Range, crng As Range
Dim adj_address As Range
Dim result As String
Dim col_num As Integer, row_num As Integer
Dim head_row_down As Integer, head_row_up As Integer, head_col_right As Integer

'' eliminate sales and COS
Sheets("InternalSales").Activate
Set company_rng = Range(Cells(1, 2), Cells(1, Columns.Count).End(xlToLeft))
Set account_rng = Union(Cells.Find("40010 - Sales", searchformat:=False, lookat:=xlPart), Cells.Find("50010 - Cost of Goods Sold", searchformat:=False, lookat:=xlPart))

On Error Resume Next
For Each arng In account_rng
    For Each crng In company_rng
        Set adj_address = adjust_address(account_name:=arng.Value, company_name:=crng.Value, from_sheet:="InternalSales", to_sheet:="PL", total_level:=False)
        adj_address.Value = adj_address.Value - Sheets("InternalSales").Cells(arng.Row, crng.Column)
    Next
Next


'' eliminate marketing expense and sales of USA company
Dim market_exp As Range, market_exp_row As Integer, market_exp_col As Integer
Sheets("PL").Activate
market_exp_row = Cells.Find("40010 - Sales", searchformat:=False, lookat:=xlPart).Row
market_exp_col = Cells.Find("Intrepid Spirits USA", searchformat:=False, lookat:=xlPart).Column
'' eliminate internal mkt expents
Set adj_address = adjust_address(account_name:="65140 - General Marketing", company_name:="Intrepid Spirits USA", from_sheet:="PL", to_sheet:="PL")
adj_address.Value = adj_address.Value - Cells(market_exp_row, market_exp_col).Value



'' eliminate net income where internal sales happens
Dim net_income_adj_PL As Range, net_income_adj_BS As Range
Sheets("PL").Activate
Application.FindFormat.Interior.color = vbYellow
head_row_down = Cells.Find("Adjustment", After:=Cells(1, 1), SearchOrder:=xlByRows, SearchDirection:=xlNext, searchformat:=True).Row
head_row_up = head_row_down - 3
head_col_right = Cells(head_row_down, Columns.Count).End(xlToLeft).Column
Set company_rng = Range(Cells(head_row_up, 2), Cells(head_row_down, head_col_right))
For Each crng In company_rng
    If crng.Value <> "" And crng.Value <> "Adjustment" And crng.Value <> "Total" And crng.Value <> "Amount" And crng.MergeCells Then ' find company's name
        ' find the adjustment cell of net income in PL
        Set net_income_adj_PL = adjust_address(account_name:="Net Income", company_name:=crng.Value, from_sheet:="PL", to_sheet:="PL", total_level:=False)
        ' find the adjustment cell of net income in BS
        Set net_income_adj_BS = adjust_address(account_name:="Net Income", company_name:=crng.Value, from_sheet:="BS", to_sheet:="BS", total_level:=False)
        ' the adjusted net income should match between BS and PL
        net_income_adj_BS.Value = net_income_adj_PL.Value
    End If
Next

End Sub
Sub adjust_CurrentBS()
Sheets("BS").Activate
Dim interaccount_names() As Variant, interaccount_name As Variant, equity_names() As Variant, equity_name As Variant
Dim head_row_down As Integer, head_row_up As Integer, head_col_right As Integer
Dim header As Range, company_rng As Range, crng As Range, account_rng As Range, arng As Range
Dim result As String, adj_address As Range, adj_num_from_col As Integer
Application.FindFormat.Interior.color = vbYellow
head_row_down = Cells.Find("Adjustment", After:=Cells(1, 1), SearchOrder:=xlByRows, SearchDirection:=xlNext, searchformat:=True).Row
head_row_up = head_row_down - 3
head_col_right = Cells(head_row_down, Columns.Count).End(xlToLeft).Column
Set company_rng = Range(Cells(head_row_up, 2), Cells(head_row_down, head_col_right))
interaccount_names = Array("InterCo. Investments", "INTERCO INVESTMENT", "20110 - Intercompany Accounts Payable", _
                     "InterCo. Loan", "Intercompany Accounts Receivable")
equity_names = Array("Common Stock", "Preferred Stock", "Share Premium", "Opening Balance")

'' eliminate intercompany items
Set account_rng = Range(Cells(head_row_down + 1, 1), Cells(Rows.Count, 1).End(xlUp))
For Each interaccount_name In interaccount_names
    For Each arng In account_rng
        If arng.Value Like "*" & interaccount_name & "*" Then ' find the aimed account
            ' eliminate subgroup at first, in subgroup level not in total level
            For Each crng In company_rng
                If crng.Value <> "" And crng.Value <> "Adjustment" And crng.Value <> "Total" And crng.Value <> "Amount" Then ' find company's name
                    Set adj_address = adjust_address(account_name:=arng.Value, company_name:=crng.Value, from_sheet:="BS", to_sheet:="BS", total_level:=False)
                    If Not crng.MergeCells Or crng.Value Like "*Intrepid Spirits Holdings Inc*" Then
                        adj_address.Value = adj_address.Value - Sheets("BS").Cells(arng.Row, crng.Column)
                    End If
                End If
            Next
            ' eliminate whole group 2nd , in subgroup level not in total level
            For Each crng In company_rng
                If crng.Value <> "" And crng.Value <> "Adjustment" And crng.Value <> "Total" And crng.Value <> "Amount" Then ' find company's name
                    Set adj_address = adjust_address(account_name:=arng.Value, company_name:=crng.Value, from_sheet:="BS", to_sheet:="BS", total_level:=False)
                    If crng.MergeCells And Not crng.Value Like "*Intrepid Spirits Holdings Inc*" Then
                        adj_num_from_col = Range(Split(crng.MergeArea.Address(RowAbsolute:=False, ColumnAbsolute:=False), ":")(1)).Column
                        adj_address.Value = adj_address.Value - Sheets("BS").Cells(arng.Row, adj_num_from_col)
                    End If
                End If
            Next
        End If
    Next
Next




'' eliminate equity items other than net income , in subgroup level not in total level
For Each equity_name In equity_names
    For Each arng In account_rng
        If arng.Value Like "*" & equity_name & "*" Then ' find the aimed account
            ' eliminate subgroup at first
            For Each crng In company_rng
                If crng.Value <> "" And crng.Value <> "Adjustment" And crng.Value <> "Total" And crng.Value <> "Amount" Then ' find company's name
                    Set adj_address = adjust_address(account_name:=arng.Value, company_name:=crng.Value, from_sheet:="BS", to_sheet:="BS", total_level:=False)
                    If Not crng.MergeCells Then
                        adj_address.Value = adj_address.Value - Sheets("BS").Cells(arng.Row, crng.Column)
                    End If
                End If
            Next
            ' eliminate whole group 2nd
            For Each crng In company_rng
                If crng.Value <> "" And crng.Value <> "Adjustment" And crng.Value <> "Total" And crng.Value <> "Amount" Then ' find company's name
                    Set adj_address = adjust_address(account_name:=arng.Value, company_name:=crng.Value, from_sheet:="BS", to_sheet:="BS")
                    If crng.MergeCells And Not crng.Value Like "*Intrepid Spirits Holdings Inc*" Then
                        adj_num_from_col = Range(Split(crng.MergeArea.Address(RowAbsolute:=False, ColumnAbsolute:=False), ":")(1)).Column 'last column of mergearea
                        adj_address.Value = adj_address.Value - Sheets("BS").Cells(arng.Row, adj_num_from_col)
                    End If
                End If
            Next
        End If
    Next
Next




End Sub
Function adjust_address(ByVal account_name As String, ByVal company_name As String, ByVal from_sheet As String, ByVal to_sheet As String, Optional ByVal total_level As Boolean = True)

'' find the company name in the header range, otherwise the intrepid spirits would be A1
Sheets(to_sheet).Activate
Dim header As Range, header_start As Range, header_end As Range
Application.FindFormat.Interior.color = header_color
Set header_end = Cells.Find(What:="*", After:=Cells(1, 1), searchformat:=True, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
Set header_start = Cells.Find(What:="*", After:=Cells(1, 1), searchformat:=True, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
Set header = Range(header_start, header_end)
Application.FindFormat.Interior.color = xlNone

Dim row_num As Integer, col_num As Integer
Dim account_rng As Range, company_rng As Range, merge_address As String

Set account_rng = Cells.Find(account_name, searchformat:=False, After:=Cells(1, 1))
row_num = account_rng.Row

Set company_rng = header.Find(company_name, searchformat:=False)
If Not company_rng Is Nothing Then ' if find the company
    If Not total_level Then
        If Not company_rng.MergeCells Then
            Set company_rng = company_rng.Offset(-1, 0)
            merge_address = company_rng.MergeArea.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            col_num = Range(Split(merge_address, ":")(1)).Offset(0, -1).Column
        Else
            merge_address = company_rng.MergeArea.Address(RowAbsolute:=False, ColumnAbsolute:=False)
            col_num = Range(Split(merge_address, ":")(1)).Offset(0, -1).Column
        End If
    Else
        Set company_rng = company_rng.Offset(-1, 0)
        merge_address = company_rng.MergeArea.Address(RowAbsolute:=False, ColumnAbsolute:=False)
        col_num = Range(Split(merge_address, ":")(1)).Offset(0, -1).Column
    End If
    Set adjust_address = Sheets(to_sheet).Cells(row_num, col_num)
End If
End Function

Function find_address(ByVal account_name As String, ByVal company_name As String, ByVal from_sheet As String)
Dim row_num As Integer, col_num As Integer
Dim account_rng As Range, company_rng As Range, merge_address As String
Sheets(from_sheet).Activate
''' find the company name in the header range, otherwise the intrepid spirits would be A1
'Dim header As Range, header_start As Range, header_end As Range
'Application.FindFormat.Interior.color = header_color
'Set header_end = Cells.Find(What:="*", After:=Cells(1, 1), searchformat:=True, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
'Set header_start = Cells.Find(What:="*", After:=Cells(1, 1), searchformat:=True, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
'Set header = Range(header_start, header_end)
'Application.FindFormat.Interior.color = xlNone


Set account_rng = Cells.Find(account_name, searchformat:=False, After:=Cells(1, 1), lookat:=xlPart)
If account_rng Is Nothing Then
    Set find_address = Nothing
    GoTo nothing_exit
End If
row_num = account_rng.Row
Set company_rng = Cells.Find(company_name, searchformat:=False, After:=Cells(1, 1), lookat:=xlPart)
If company_rng Is Nothing Then
    Set find_address = Nothing
    GoTo nothing_exit
End If
col_num = company_rng.Column
Set find_address = Sheets(from_sheet).Cells(row_num, col_num)
nothing_exit:
End Function

Sub inventory()
Sheets("Regal Rogue").Activate
If ActiveSheet.AutoFilterMode = True Then
    ActiveSheet.AutoFilterMode = False
End If
Dim source_rng As Range, company_rng As Range, account_rng As Range, arng As Range, crng As Range, adj_address As Range
Dim s As String
Dim start_row As Integer, end_row As Integer, end_col As Integer, amount_col As Integer

start_row = Cells.Find("Item", Cells(1, 1), SearchOrder:=xlNext, searchformat:=False).Row
end_row = Cells(Rows.Count, 1).End(xlUp).Row
end_col = Cells(start_row, Columns.Count).End(xlToLeft).Column
Set source_rng = Range(Cells(start_row, 1), Cells(end_row, end_col))
s = source_rng.Address
source_rng.AutoFilter Field:=2, Criteria1:="=*Regal Rogue*"
source_rng.AutoFilter Field:=3, Criteria1:=">0"
source_rng.Copy Cells(end_row + 3, 1)
ActiveSheet.AutoFilterMode = False
Range(Cells(start_row, 1), Cells(end_row, end_col)).EntireRow.Delete


start_row = Cells.Find("Item", Cells(1, 1), SearchOrder:=xlNext, searchformat:=False).Row
end_row = Cells(Rows.Count, 1).End(xlUp).Row
end_col = Cells(start_row, Columns.Count).End(xlToLeft).Column
Set source_rng = Range(Cells(start_row, 1), Cells(end_row, end_col))
amount_col = Cells.Find("Inv. Value", searchformat:=False).Column
Cells(end_row + 1, amount_col).FormulaR1C1 = "=SUM(R[" & -end_row + start_row & "]C:R[-1]C)"
Cells(end_row + 1, 1) = "Total " & ActiveSheet.Name

'' get the inventory balance * margin of RR in internalSales
Dim internal_inventory_RR As Range, internal_RE_RR As Range
Sheets("InternalSales").Activate
Set internal_inventory_RR = find_address("InventoryBalance", "Intrepid Spirits Limited", "InternalSales")
internal_inventory_RR.Value = Sheets("Regal Rogue").Cells(end_row + 1, amount_col).Value


'' get the inventory balance * margin of Japan in internalSales
Dim JP_inventory As Range, internal_inventory_JP As Range, internal_RE_JP As Range
Set JP_inventory = find_address(account_name:="14020 - Finished Goods Inventory", company_name:="Intrepid Japan", from_sheet:="BS")
If Not JP_inventory Is Nothing Then
    Set internal_inventory_JP = find_address(account_name:="InventoryBalance", company_name:="Intrepid Japan", from_sheet:="InternalSales")
    internal_inventory_JP.Value = JP_inventory.Value
End If

'' get the inventory balance * margin of IDL in internalSales
Dim IDL_inventory As Range, internal_inventory_IDL As Range, internal_RE_IDL As Range
Set IDL_inventory = find_address(account_name:="14020 - Finished Goods Inventory", company_name:="Intrepid Spirits Ireland Ltd", from_sheet:="BS")
If Not IDL_inventory Is Nothing Then
    Set internal_inventory_IDL = find_address(account_name:="InventoryBalance", company_name:="Intrepid Spirits Ireland Ltd", from_sheet:="InternalSales")
    internal_inventory_IDL.Value = IDL_inventory.Value
End If

'' delete the columns without internal sales, COS and inventory balance
Dim delete_rng As Range, col As Integer
Dim sale_row As Integer, cos_row As Integer, finished_goods_row As Integer, i As Integer
Sheets("InternalSales").Activate
sale_row = Cells.Find("40010 - Sales").Row
cos_row = Cells.Find("50010 - Cost of Goods Sold").Row
finished_goods_row = Cells.Find("14020 - Finished Goods Inventory").Row

For i = 2 To ActiveSheet.UsedRange.Columns.Count
    If Cells(sale_row, i).Value = 0 And Cells(cos_row, i) = 0 And Cells(finished_goods_row, i) = 0 Then
        If delete_rng Is Nothing Then
            Set delete_rng = Cells(sale_row, i)
        Else
            Set delete_rng = Union(delete_rng, Cells(sale_row, i))
        End If
    End If
Next i
delete_rng.EntireColumn.Delete


'' eliminate inventory and RE, in total level
Dim offset_rows As Integer, offset_cols As Integer
Sheets("InternalSales").Activate
Set account_rng = Union(Cells.Find("14020 - Finished Goods Inventory", searchformat:=False, lookat:=xlPart), Cells.Find("Retained Earnings", searchformat:=False, lookat:=xlPart))
Set company_rng = Range(Cells(1, 2), Cells(1, Columns.Count).End(xlToLeft))
For Each arng In account_rng
    For Each crng In company_rng
        Set adj_address = adjust_address(account_name:=arng.Value, company_name:=crng.Value, from_sheet:="InternalSales", to_sheet:="BS", total_level:=True)
        offset_rows = arng.Row - adj_address.Row
        offset_cols = crng.Column - adj_address.Column
        If adj_address.HasFormula Then
            adj_address.FormulaR1C1 = adj_address.FormulaR1C1 & " - InternalSales!R[" & offset_rows & "]C[" & offset_cols & "]"
        Else
            adj_address.FormulaR1C1 = "=" & adj_address.Value & " - InternalSales!R[" & offset_rows & "]C[" & offset_cols & "]"
        End If

    Next
Next
End Sub
Sub export()
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Dim sht As Worksheet
Dim wb As Workbook
Dim my_path As String, response As String



my_path = ThisWorkbook.Path & "\" & Sheets("PL").Range("A4").Value & " consolidation.xls"
ThisWorkbook.SaveCopyAs my_path
'Call delete_sheet
'Set wb = Workbooks.Open(my_path)
'wb.Sheets("Instructions").Delete
'wb.Close

'' export PL to workbook for later use of consolidation of per brand and region and subsidiary
Dim PL_export_path As String
Dim goodwill_cell As Range
PL_export_path = "F:\Intrepid Spirits\Consolidation\Consolidation Per Brand & Region & Subsidiary\PLperSub" & Sheets("PL").Range("A4").Value & ".xlsx"
Sheets("PL").Copy
ActiveSheet.Name = "PLperSub"
Set goodwill_cell = ActiveSheet.Cells.Find("Goodwill Impairment", lookat:=xlPart)
goodwill_cell.EntireRow.Copy
goodwill_cell.PasteSpecial xlPasteValues
ActiveWorkbook.SaveAs PL_export_path
ActiveWorkbook.Close

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Sub interface()
Instructions.Show
End Sub

Sub goodwill_summary()
Sheets("BS").Activate
Dim interco_investment_row As Integer, BS_cols As Integer, total_equity_row As Integer
'interco_investment_row = Cells.Find("Total - 13000 - - INTERCOMPANY INVESTMENTS -", searchformat:=False, lookat:=xlPart).Row
'BS_cols = Cells(interco_investment_row, Columns.Count).End(xlToLeft).Column
'' sum the intercompany investment eliminated
'Sheets("Goodwill").Cells.Find("Eliminated Intercompany Investment").Offset(0, 1).Value = Cells(interco_investment_row, BS_cols).Value
'
'total_equity_row = Cells.Find("Total Equity", searchformat:=False, lookat:=xlPart).Row
'' sum the total equity of subsidiaries eliminated
'Sheets("Goodwill").Cells.Find("Eliminated Subsidiaries' Equity").Offset(0, 1).Value = Cells(total_equity_row, BS_cols).Value - Cells(total_equity_row, 2).Value
'
'Sheets("Goodwill").Cells.Find("Goodwill Before Impairment").Offset(0, 1).FormulaR1C1 = "=R[-2]C-R[-1]C"

Sheets("Goodwill").Cells.Find("Goodwill After Impairment").Offset(0, 1).FormulaR1C1 = "=R[-3]C-R[-2]C-R[-1]C"
End Sub

Sub goodwill_adjustment()
Sheets("BS").Activate
Dim goodwill_row As Integer, BS_cols As Integer
Dim goodwill_netincome_adj As Range, goodwill_currently_impair As Range
Dim offset_rows As Integer, offset_cols As Integer

' Goodwill after impairment show in the BS
goodwill_row = Cells.Find("Goodwill", lookat:=xlPart, searchformat:=False).Row
BS_cols = Cells(goodwill_row, Columns.Count).End(xlToLeft).Column
offset_rows = Sheets("Goodwill").Cells.Find("Goodwill After Impairment", searchformat:=False, lookat:=xlPart).Offset(0, 1).Row - goodwill_row
offset_cols = Sheets("Goodwill").Cells.Find("Goodwill After Impairment", searchformat:=False, lookat:=xlPart).Offset(0, 1).Column - BS_cols
Sheets("BS").Cells(goodwill_row, BS_cols).FormulaR1C1 = "=Goodwill!R[" & offset_rows & "]C[" & offset_cols & "]"

' adjust netincome in BS
Set goodwill_netincome_adj = adjust_address(account_name:="Net Income", company_name:="Intrepid Spirits Holdings Inc", from_sheet:="BS", to_sheet:="BS")
Set goodwill_currently_impair = Sheets("Goodwill").Cells.Find("Goodwill Currently Impaired By", lookat:=xlPart, searchformat:=False).Offset(0, 1)
offset_rows = goodwill_currently_impair.Row - goodwill_netincome_adj.Row
offset_cols = goodwill_currently_impair.Column - goodwill_netincome_adj.Column
If goodwill_netincome_adj.HasFormula Then
    goodwill_netincome_adj.FormulaR1C1 = goodwill_netincome_adj.FormulaR1C1 & " -Goodwill!R[" & offset_rows & "]C[" & offset_cols & "]"
Else
    goodwill_netincome_adj.FormulaR1C1 = "=" & goodwill_netincome_adj.Value & " -Goodwill!R[" & offset_rows & "]C[" & offset_cols & "]"
End If

' adjust goodwill cuurent impairment in PL
Set goodwill_netincome_adj = adjust_address(account_name:="Goodwill Impairment", company_name:="Intrepid Spirits Holdings Inc", from_sheet:="PL", to_sheet:="PL")
offset_rows = goodwill_currently_impair.Row - goodwill_netincome_adj.Row
offset_cols = goodwill_currently_impair.Column - goodwill_netincome_adj.Column
If goodwill_netincome_adj.HasFormula Then
    goodwill_netincome_adj.FormulaR1C1 = goodwill_netincome_adj.FormulaR1C1 & " -Goodwill!R[" & offset_rows & "]C[" & offset_cols & "]"
Else
    goodwill_netincome_adj.FormulaR1C1 = "=" & goodwill_netincome_adj.Value & " -Goodwill!R[" & offset_rows & "]C[" & offset_cols & "]"
End If

' adjust RE related to goodwill impairment in BS
Dim goodwill_RE_adj As Range, goodwill_previously_impair As Range
Set goodwill_RE_adj = adjust_address(account_name:="Retained Earnings", company_name:="Intrepid Spirits Holdings Inc", from_sheet:="BS", to_sheet:="BS")
Set goodwill_previously_impair = Sheets("Goodwill").Cells.Find("Goodwill Previously Impaired By", lookat:=xlPart, searchformat:=False).Offset(0, 1)
offset_rows = goodwill_previously_impair.Row - goodwill_RE_adj.Row
offset_cols = goodwill_previously_impair.Column - goodwill_RE_adj.Column
If goodwill_RE_adj.HasFormula Then
    goodwill_RE_adj.FormulaR1C1 = goodwill_RE_adj.FormulaR1C1 & " - Goodwill!R[" & offset_rows & "]C[" & offset_cols & "]"
Else
    goodwill_RE_adj.FormulaR1C1 = "=" & goodwill_RE_adj.Value & " -Goodwill!R[" & offset_rows & "]C[" & offset_cols & "]"
End If

End Sub

Sub reformat()
Dim sht As Worksheet
For Each sht In Sheets
    sht.Columns.EntireColumn.AutoFit
Next
End Sub







