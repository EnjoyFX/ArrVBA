Attribute VB_Name = "test"
Sub Example()

    Const Num = 1000, Lower = 0, Upper = 1000000, Reverse = False

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

    A.RndFill elements:=Num, lowerBound:=Lower, upperBound:=Upper

    A.OutHorizontal "B9"

    A.OutVertical "B11"

    'A.OutDiagonal "D11"

    A.Sort Method:=SortMethod.Bubble, Reverse:=Reverse

    A.OutVertical "C11"


    A.RndFill elements:=Num, lowerBound:=Lower, upperBound:=Upper

    A.OutVertical "E11"

    A.Sort Method:=SortMethod.Insertion, Reverse:=Reverse

    A.OutVertical "F11"



    A.RndFill elements:=Num, lowerBound:=Lower, upperBound:=Upper

    A.OutVertical "H11"

    A.Sort Method:=SortMethod.Selection, Reverse:=Reverse

    A.OutVertical "I11"


    A.RndFill elements:=Num, lowerBound:=Lower, upperBound:=Upper

    A.OutVertical "K11"

    A.Sort Method:=SortMethod.Quick, Reverse:=Reverse

    A.OutVertical "L11"


    A.RndFill elements:=Num, lowerBound:=Lower, upperBound:=Upper

    A.OutVertical "N11"
    
    Debug.Print "Array is sorted = "; A.isSorted
    
    A.Sort Method:=SortMethod.Heap, Reverse:=Reverse

    A.OutVertical "O11"

    Debug.Print "Array is sorted = "; A.isSorted
    
    'A.Reverse

    'A.OutVertical "E11"

    If src Then Application.ScreenUpdating = True

End Sub

Sub clsSheet()

    Dim src As Boolean

    src = Application.ScreenUpdating

    If src Then Application.ScreenUpdating = False

    ActiveSheet.Cells.Clear

    If src Then Application.ScreenUpdating = True

End Sub

