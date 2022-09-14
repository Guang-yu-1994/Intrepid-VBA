Attribute VB_Name = "replace_name_module"
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
Sub trim_rng(from_col As Integer, ByVal until_col As Integer)
'On Error Resume Next
Dim rng As Range
For Each rng In Range(Cells(1, from_col), Cells(ActiveSheet.UsedRange.Rows.Count, until_col))
    rng.Value = Trim(rng.Value)
Next
End Sub

