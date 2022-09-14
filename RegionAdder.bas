Attribute VB_Name = "RegionAdder"
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
