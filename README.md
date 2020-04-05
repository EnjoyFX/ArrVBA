# ArrVBA
Class for operations with arrays in VBA

```vba
Sub Example()

    Dim A As New ArrVBA, B, C, D

    A.Based = 0

    A.Add 2020
    A.Add "This is"
    A.Add "short"
    A.Add "test"
    A.Add "of ArrVBA Class"

    A.PrintMe

    B = A.AsVariant
    C = A.AsVertical
    D = A.AsHorizontal

    A.OutHorizontal "B7", bold:=True

    A.RndFill elements:=800, lowerBound:=0, upperBound:=10000

    A.OutHorizontal "B9"

    A.OutVertical "B11"

    'A.OutDiagonal "D11"

    A.Sort

    A.OutVertical "C11"

    A.Sort Reverse:=True

    A.OutVertical "D11"
    
    A.Reverse
    
    A.OutVertical "E11"

End Sub

```
