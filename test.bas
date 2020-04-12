Attribute VB_Name = "test"
' Tests for ArrVBA methods
' call runTests() - to run tests
' is some error is present - tests will stop on line with error value

Option Explicit


Sub runTests()

    Dim A As New ArrVBA, toVar$, B As Variant

    ' test for Based 0
    
    A.Based = 0
    
    Debug.Assert A.Based = 0
    
    Debug.Assert A.Count = 0
    
    Debug.Assert A.AsString = vbNullString
    
    A.Add "one"
    
    Debug.Assert A.Arr(0) = "one"
    
    
    A.Add "two"
    
    Debug.Assert A.Arr(1) = "two"
    
    
    A.Add 2020
    
    Debug.Assert A.Arr(2) = 2020
    
    A.Add 777
    
    Debug.Assert A.Arr(3) = 777
    
    Debug.Assert A.Count = 4
    
    Debug.Assert A.MaxValue = 2020
    
    Debug.Assert A.MinValue = 777
    
    Debug.Assert A.AsString(",") = "one,two,2020,777"
    
    Call A.PrintMe(delim:=", ", toVar:=toVar)
    
    Debug.Assert toVar = "one, two, 2020, 777"
    
    
    A.Reverse
    
    Debug.Assert A.AsString(",") = "777,2020,two,one"
    
    A.Add "tWeNtY"
    
    B = A.FilterArr("tw")
    
    Debug.Assert Join(B, ",") = "two,tWeNtY"
    
    Debug.Assert A.isIncludeTemplate("tw") = True
    
    Debug.Assert A.isIncludeTemplate("ti") = False
    
    A.Clear
    
    
    ' test for Based 1
    
    A.Based = 1
    
    Debug.Assert A.Based = 1
    
    
    ' test for Based 10
    
    A.Based = 10
    
    Debug.Assert A.Based = 10
    
    
    Call MsgBox("Tests are passed")
    


End Sub



