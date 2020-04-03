# ArrVBA
Class for operations with arrays in VBA

```vba
Sub tester()

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

End Sub

```
