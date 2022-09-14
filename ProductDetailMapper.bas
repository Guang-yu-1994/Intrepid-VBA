Attribute VB_Name = "ProductDetailMapper"
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
