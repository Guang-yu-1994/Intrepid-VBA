Attribute VB_Name = "HistoryConcator"
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
