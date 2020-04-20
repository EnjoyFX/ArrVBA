Attribute VB_Name = "test"
' Tests for ArrVBA methods
' Call runTests() - to run tests
' If some error is present - tests will stop on line with error value
' Otherwise - tests are passed

Option Explicit

Dim A As New ArrVBA, toVar$, B As Variant

Sub runTests()

' test for Based 0

    Call testsBased(0)


    ' test for Based 1

    Call testsBased(1)


    ' test for Based 10

    Call testsBased(10)


    Call MsgBox("Tests are passed")

End Sub



Sub testsBased(ByVal theBase%)

    A.Clear

    A.Based = theBase

    Debug.Assert A.Based = theBase

    Debug.Assert A.Count = 0

    Debug.Assert A.AsString = vbNullString

    A.Add "one"

    Debug.Assert A.Arr(theBase) = "one"


    A.Add "two"

    Debug.Assert A.Arr(theBase + 1) = "two"


    A.Add 2020

    Debug.Assert A.Arr(theBase + 2) = 2020

    A.Add 777

    Debug.Assert A.Arr(theBase + 3) = 777

    Debug.Assert A.Count = 4

    Debug.Assert A.MaxValue = 2020

    Debug.Assert A.MinValue = 777

    Debug.Assert A.AsString(",") = "one,two,2020,777"

    Call A.PrintMe(delim:=", ", toVar:=toVar)

    Debug.Assert toVar = "one, two, 2020, 777"

    A.ReverseArr

    Debug.Assert A.AsString(",") = "777,2020,two,one"

    A.Add "tWeNtY"

    B = A.FilterArr("tw")

    Debug.Assert Join(B, ",") = "two,tWeNtY"

    Debug.Assert A.isIncludeTemplate("tw") = True

    Debug.Assert A.isIncludeTemplate("ti") = False


    ' sort method tests

    Dim n%, m%, Reverse As Boolean, testElems, testSort$, testSortRe$

    testElems = Array(99, 10, 5, 3, 100, 1, -1, 7)

    testSort = "-1,1,3,5,7,10,99,100"

    testSortRe = "100,99,10,7,5,3,1,-1"


    For n = SortMethod.[_First] To SortMethod.[_Last]

        For m = 0 To 1

            A.Clear

            A.AddArr elems:=testElems

            Reverse = CBool(m)

            A.Sort Method:=n, Reverse:=Reverse

            Debug.Print "SortMethod = "; n; " Reverse = "; Reverse; " Base = "; theBase

            If Reverse = False Then

                Debug.Assert A.AsString = testSort

            Else

                Debug.Assert A.AsString = testSortRe

            End If

        Next m

    Next n


    Dim lowerBand&, upperBand&, Count&

    lowerBand = 1: upperBand = 10: Count = 3

    A.RndFill 3, lowerBound:=1, upperBound:=10

    Debug.Assert A.Count = Count

    Debug.Assert A.Arr(A.Based) >= lowerBand

    Debug.Assert A.Arr(A.Based) <= upperBand



    A.Clear

    A.ValueFill 10, "!"

    Debug.Assert A.AsString = "!,!,!,!,!,!,!,!,!,!"

    A.ValueFill 3, 999

    Debug.Assert A.AsString = "999,999,999"




End Sub


