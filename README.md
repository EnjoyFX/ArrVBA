# ArrVBA
Class for operations with arrays in VBA

```vba
Sub Example()

    Const num = 1000, lower = 0, upper = 1000000, reverse = False

    Dim A As New ArrVBA, B As New ArrVBA, C, D, E, src As Boolean

    src = Application.ScreenUpdating

    If src Then Application.ScreenUpdating = False


    A.Based = 0

    A.Add 2020
    A.Add "This is"
    A.Add "short"
    A.Add "test"
    A.Add "of ArrVBA Class"

    Debug.Print A.MaxValue
    
    Debug.Print A.MinValue

    A.PrintMe

    C = A.AsVariant
    D = A.AsVertical
    E = A.AsHorizontal

    A.OutHorizontal "B7", bold:=True

    A.RndFill elements:=num, lowerBound:=lower, upperBound:=upper

    A.OutHorizontal "B9"

    A.OutVertical "B11"

    'A.OutDiagonal "D11"

    A.Sort method:=SortMethod.Bubble, reverse:=reverse

    A.OutVertical "C11"


    A.RndFill elements:=num, lowerBound:=lower, upperBound:=upper

    A.OutVertical "E11"

    A.Sort method:=SortMethod.Insertion, reverse:=reverse
    
    A.OutVertical "F11"



    A.RndFill elements:=num, lowerBound:=lower, upperBound:=upper

    A.OutVertical "H11"

    A.Sort method:=SortMethod.Selection, reverse:=reverse

    A.OutVertical "I11"


    A.RndFill elements:=num, lowerBound:=lower, upperBound:=upper

    A.OutVertical "K11"

    A.Sort method:=SortMethod.Quick, reverse:=reverse

    A.OutVertical "L11"


'    G.RndFill elements:=num, lowerBound:=lower, upperBound:=upper
'
'    G.OutVertical "N11"
'
'    G.Sort method:=SortMethod.Heap ', Reverse:=True
'
'    G.OutVertical "O11"


    'A.Reverse

    'A.OutVertical "E11"

    If src Then Application.ScreenUpdating = True

End Sub


```
