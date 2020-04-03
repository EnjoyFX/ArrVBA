VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrVBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const className$ = "ArrVBA"    ' used in error handling process messages

Private Count As Long
Private theArr As Variant
Private t_Based As Integer


Private Sub Class_Initialize()

    theArr = Empty
    Count = 0

End Sub



Public Property Let Based(ByVal value%)

    t_Based = value

End Property


Public Property Get Based() As Integer

    Based = t_Based

End Property



Sub PrintMe(Optional ByVal delim = ", ")

    Dim txt$

    txt = Join(theArr, delim)

    Debug.Print txt

End Sub



Function Add(ByVal elem) As Variant    ' base 0+

    Const this_me$ = "Add"

    Dim remax&

    On Error GoTo errHandler

    If IsEmpty(theArr) Then

        ReDim theArr(t_Based To t_Based)
        theArr(t_Based) = elem
        Count = 1

    Else

        remax = UBound(theArr) + 1
        ReDim Preserve theArr(t_Based To remax)
        theArr(remax) = elem
        Count = Count + 1

    End If

    Add = theArr

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(Err, this_me)
    Err.Clear
    Resume Next

End Function


Function AsVariant() As Variant

    AsVariant = theArr

End Function


Function AsVertical() As Variant

    Const this_me$ = "AsVertical"

    Dim B, minB&, maxB&, n&, adder&

    On Error GoTo errHandler

    minB = LBound(theArr)

    maxB = UBound(theArr)

    If Based = 0 Then adder = 1 Else adder = 0

    ReDim B(1 To Count, 1 To 1)

    For n = minB To maxB

        B(n + adder, 1) = theArr(n)

    Next n

    AsVertical = B

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(Err, this_me)
    Err.Clear
    Resume Next

End Function


Function AsHorizontal() As Variant

    Const this_me$ = "AsHorizontal"

    Dim B, minB&, maxB&, n&, adder&

    On Error GoTo errHandler

    minB = LBound(theArr)

    maxB = UBound(theArr)

    If Based = 0 Then adder = 1 Else adder = 0

    ReDim B(1 To 1, 1 To Count)

    For n = minB To maxB

        B(1, n + adder) = theArr(n)

    Next n

    AsHorizontal = B

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(Err, this_me)
    Err.Clear
    Resume Next

End Function


Private Sub tinyHandler(ByRef ErrObj, ByVal this_me$)

    Dim txt$

    txt = "[" & className & "." & this_me & "]: " & ErrObj.Description & " (" & ErrObj.Number & ")"

    Debug.Print txt

End Sub
