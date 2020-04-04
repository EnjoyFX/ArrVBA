# ArrVBA
Class for operations with arrays in VBA

```vba
Sub Example()

    Dim A As New ArrVBA, B, C, D

    A.Based = 0
    
    A.Add 1
    A.Add 2
    A.Add 10
    A.Add "Hello"
    
    A.PrintMe
    
    B = A.AsVariant
    C = A.AsVertical
    D = A.AsHorizontal
    
    A.OutHorizontal ("B7")
    
    Call A.RndFill(800, lowerBound:=0, upperBound:=100)
    
    A.OutHorizontal ("B9")
    
    A.OutVertical ("B11")

End Sub

```
