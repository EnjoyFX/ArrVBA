Attribute VB_Name = "test"
Sub Example()

    Const num = 1000, lower = 0, upper = 1000000

    Dim A As New ArrVBA, B, C, D, E As New ArrVBA, F As New ArrVBA, G As New ArrVBA


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

    A.RndFill elements:=num, lowerBound:=lower, upperBound:=upper

    A.OutHorizontal "B9"

    A.OutVertical "B11"

    'A.OutDiagonal "D11"

    A.Sort method:=SortMethod.Bubble

    A.OutVertical "C11"


    E.RndFill elements:=num, lowerBound:=lower, upperBound:=upper

    E.OutVertical "E11"

    E.Sort method:=SortMethod.Insertion

    E.OutVertical "F11"



    F.RndFill elements:=num, lowerBound:=lower, upperBound:=upper

    F.OutVertical "H11"

    F.Sort method:=SortMethod.Selection

    F.OutVertical "I11"
    
    
    G.RndFill elements:=num, lowerBound:=lower, upperBound:=upper
    
    G.OutVertical "K11"
            
    G.Sort method:=SortMethod.Quick

    G.OutVertical "L11"
    
    
    'A.Reverse
    
    'A.OutVertical "E11"


End Sub

Sub clsSheet()

    Dim src As Boolean

    src = Application.ScreenUpdating

    If src Then Application.ScreenUpdating = False

    ActiveSheet.Cells.Clear

    If src Then Application.ScreenUpdating = True

End Sub
