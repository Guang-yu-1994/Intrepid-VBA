Attribute VB_Name = "CaseConfigSpliterRegVersion"
'Function get_bottles_per_case(rng As Range)
'Set reg = CreateObject("vbscript.regexp")
'With reg
'    .Global = True
'    .Pattern = "\D\d+\D+"
'    get_bottles_per_case = CInt(.Replace(Trim(rng(r, c).Value), ""))
'End With
'End Function

Function get_bottles_per_case(rng As Range)

Dim i As Integer
Dim r As Integer, c As Integer
Dim r_result(), r_rng
Set reg = CreateObject("vbscript.regexp")
r_rng = rng ' store rng in array
If IsArray(r_rng) Then
    ReDim r_result(1 To UBound(r_rng, 1), 1 To UBound(r_rng, 2))
    With reg
        .Global = True
        .Pattern = "\D\d+\D+"
        For r = 1 To UBound(r_rng, 1) ' iteration of row
            For c = 1 To UBound(r_rng, 2)  ' iteration of column
                r_result(r, c) = CInt(.Replace(Trim(r_rng(r, c)), ""))
            Next
        Next
    End With
    get_bottles_per_case = r_result
Else
    With reg
    .Global = True
    .Pattern = "\D\d+\D+"
        get_bottles_per_case = CInt(.Replace(Trim(r_rng), ""))
    End With
End If
End Function

Function get_ml_per_bottle(rng As Range)
Dim i As Integer
Dim r As Integer, c As Integer
Dim r_result(), r_rng
ReDim r_result(1 To rng.Rows.Count, 1 To rng.Columns.Count)
Set reg = CreateObject("vbscript.regexp")
r_rng = rng ' store rng in array
If IsArray(r_rng) Then
    With reg
        .Global = True
        .Pattern = "(\d+)(ml)"
        For r = 1 To UBound(r_rng, 1)
            For c = 1 To UBound(r_rng, 2)
                r_result(r, c) = CInt(.execute(Trim(r_rng(r, c)))(0).SubMatches(0))
            Next
        Next
    End With
    get_ml_per_bottle = r_result
Else
    With reg
        .Global = True
        .Pattern = "(\d+)(ml)"
        get_ml_per_bottle = CInt(.execute(Trim(r_rng))(0).SubMatches(0))
    End With
End If
End Function

Function get_case_config(rng As Range)
Dim i As Integer
Dim r As Integer, c As Integer
Dim r_result(), r_rng
ReDim r_result(1 To rng.Rows.Count, 1 To rng.Columns.Count)
Set reg = CreateObject("vbscript.regexp")
r_rng = rng ' store rng in array
If IsArray(r_rng) Then
    With reg
        .Global = True
        .Pattern = "\d+\s*x.*\d+\s*[mc]l"
        For r = 1 To UBound(r_rng, 1)
            For c = 1 To UBound(r_rng, 2)
                r_result(r, c) = .execute(Trim(r_rng(r, c)))(0)
            Next
        Next
    End With
    get_case_config = r_result
Else
    With reg
        .Global = True
        .Pattern = "\d+\s*x.*\d+\s*[mc]l"
        get_case_config = .execute(Trim(r_rng))(0)
    End With
End If
End Function
Function get_variant(rng As Range)
Dim i As Integer
Dim r As Integer, c As Integer
Dim r_result(), r_rng
ReDim r_result(1 To rng.Rows.Count, 1 To rng.Columns.Count)
Set reg = CreateObject("vbscript.regexp")
r_rng = rng ' store rng in array
If IsArray(r_rng) Then
    With reg
        .Global = True
        .Pattern = "\d+\s*x.*\d+\s*[mc]l"
        For r = 1 To UBound(r_rng, 1)
            For c = 1 To UBound(r_rng, 2)
                r_result(r, c) = .Replace(Trim(r_rng(r, c)), "")
            Next
        Next
    End With
    get_variant = r_result
Else
    With reg
        .Global = True
        .Pattern = "\d+\s*x.*\d+\s*[mc]l"
        get_variant = .Replace(Trim(r_rng), "")
    End With
End If
End Function
Sub trim_rng()
Dim sht As Worksheet
Dim rng As Range
For Each sht In Sheets
    For Each rng In sht.UsedRange
        rng.Value = Trim(rng(r, c).Value)
    Next
Next
End Sub

Sub replace_name()
Set dic = CreateObject("Scripting.Dictionary")
Set rep = GetObject("F:\Intrepid Spirits\Budget\Budet Restructure\Replacement\Reoplacement.xlsx")
Dim last_row As Integer, sht As Worksheet, i As Integer, k
Set sht = rep.Sheets("ReplacementAll")
last_row = sht.Cells(Rows.Count, 1).End(xlUp).Row
'' create dictionary for replacement
For i = 2 To last_row
    dic(sht.Cells(i, 1).Value) = sht.Cells(i, 2).Value
Next

rep.Close

For i = 0 To dic.Count
    
Next
'' start replacing
For Each sht In ThisWorkbook.Sheets
    For Each k In dic.keys()
        sht.Cells.Replace k, dic(k)
    Next
Next

Set dic = Nothing
End Sub

