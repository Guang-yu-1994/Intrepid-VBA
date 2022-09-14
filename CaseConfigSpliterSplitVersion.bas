Attribute VB_Name = "CaseConfigSpliterSplitVersion"
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
ReDim r_result(1 To rng.Rows.Count, 1 To rng.Columns.Count)
r_rng = rng
If IsArray(r_rng) Then
    For r = 1 To UBound(r_rng, 1) ' iteration of row
        For c = 1 To UBound(r_rng, 2)  ' iteration of column
            r_result(r, c) = CInt(Split(Left(Trim(r_rng(r, c)), Len(Trim(r_rng(r, c))) - Len("ml")), "x")(0))
        Next
    Next
    get_bottles_per_case = r_result
Else
    get_bottles_per_case = CInt(Split(Left(Trim(r_rng), Len(Trim(r_rng)) - Len("ml")), "x")(0))
End If
End Function

Function get_ml_per_bottle(rng As Range)
Dim i As Integer
Dim r As Integer, c As Integer
Dim r_result(), r_rng
ReDim r_result(1 To rng.Rows.Count, 1 To rng.Columns.Count)
r_rng = rng
If IsArray(r_rng) Then
    For r = 1 To UBound(r_rng, 1) ' iteration of row
        For c = 1 To UBound(r_rng, 2)  ' iteration of column
            r_result(r, c) = CInt(Split(Left(Trim(r_rng(r, c)), Len(Trim(r_rng(r, c))) - Len("ml")), "x")(1))
        Next
    Next
    get_ml_per_bottle = r_result
Else
    get_ml_per_bottle = CInt(Split(Left(Trim(r_rng), Len(Trim(r_rng)) - Len("ml")), "x")(1))
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
        .Pattern = "(\d+\.?\d+%)?\s*\d+\s*x.*\d+\s*[mc]l\s*(\d+\.?\d+%)?"
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
        .Pattern = "(\d+\.?\d+%)?\s*\d+\s*x.*\d+\s*[mc]l\s*(\d+\.?\d+%)?"
        get_variant = .Replace(Trim(r_rng), "")
    End With
End If
End Function
Function get_ABV(rng As Range)
Dim r_result(), r_rng
ReDim r_result(1 To rng.Rows.Count, 1 To rng.Columns.Count)
Set reg = CreateObject("vbscript.regexp")
r_rng = rng ' store rng in array
If IsArray(r_rng) Then
    With reg
        .Global = True
        .Pattern = "\d+\.?\d+%"
        For r = 1 To UBound(r_rng, 1)
            For c = 1 To UBound(r_rng, 2)
                r_result(r, c) = .execute(Trim(r_rng(r, c)))(0)
            Next
        Next
    End With
    get_ABV = r_result
Else
    With reg
        .Global = True
        .Pattern = "\d+\.?\d+%"
        On Error GoTo vlookup_ABV
        get_ABV = .execute(Trim(r_rng))(0)
        Exit Function

    End With

End If
vlookup_ABV:
End Function







