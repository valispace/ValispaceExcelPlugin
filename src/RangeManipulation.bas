Attribute VB_Name = "RangeManipulation"
Function subtractOneArea(Rng1 As Range, inRng2 As Range) As Range
    If Rng1.Areas.Count > 1 Then Exit Function
    If inRng2.Areas.Count > 1 Then Exit Function
    If Application.Intersect(Rng1, inRng2) Is Nothing Then
        Set subtractOneArea = Rng1
        Exit Function
        End If
    Dim Rng2 As Range
    Set Rng2 = Application.Intersect(Rng1, inRng2)
    Dim aRng As Range, OKRng As Range, Rslt As Range, WS As Worksheet
    Set WS = Rng1.Parent
    If Rng2.Row > Rng1.Row Then
        Set Rslt = WS.Range(Rng1.Rows(1), Rng1.Rows(Rng2.Row - Rng1.Row))
        End If
    If Rng2.Row + Rng2.Rows.Count < Rng1.Row + Rng1.Rows.Count Then
        Set Rslt = Union(Rslt, _
            WS.Range(Rng1.Rows(Rng2.Row - Rng1.Row + Rng2.Rows.Count + 1), _
                Rng1.Rows(Rng1.Rows.Count)))
        End If
    If Rng2.Column > Rng1.Column Then
        Set Rslt = Union(Rslt, WS.Range(WS.Cells(Rng2.Row, Rng1.Column), _
            WS.Cells(Rng2.Row + Rng2.Rows.Count - 1, Rng2.Column - 1)))
       End If
    If Rng2.Column + Rng2.Columns.Count < Rng1.Column + Rng1.Columns.Count Then
        Set Rslt = Union(Rslt, _
            WS.Range(WS.Cells(Rng2.Row, Rng2.Column + Rng2.Columns.Count), _
                WS.Cells(Rng2.Row + Rng2.Rows.Count - 1, _
                    Rng1.Column + Rng1.Columns.Count - 1)))
        End If
    Set subtractOneArea = Rslt
End Function
    
    
    
Function Subtract(Rng1 As Range, Rng2 As Range) As Range
    On Error Resume Next
    If Application.Intersect(Rng1, Rng2).Address <> Rng2.Address Then _
        Exit Function
    On Error GoTo 0
    Dim Rslt As Range, Rng1Rslt As Range, J As Integer, i As Integer
    For J = 1 To Rng1.Areas.Count
        Set Rslt = subtractOneArea(Rng1.Areas(J), Rng2.Areas(1))
        For i = 2 To Rng2.Areas.Count
            Set Rslt = Application.Intersect( _
                Rslt, subtractOneArea(Rng1.Areas(J), Rng2.Areas(i)))
            Next i
        If Rng1Rslt Is Nothing Then
            Set Rng1Rslt = Rslt
        ElseIf Rslt Is Nothing Then
            Rng1Rslt = Rng1Rslt
        Else
            Set Rng1Rslt = Union(Rng1Rslt, Rslt)
        End If
        Next J
    Set Subtract = Rng1Rslt
End Function
    

