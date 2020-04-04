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

#Const debug_me = True     ' Compiler directive  - if debug and timing needed

Const className$ = "ArrVBA"    ' Used in error handling process messages

Private Count As Long    ' Counter of elements in ArrVBA instance
Private theArr As Variant    ' Variant Array
Private t_Based As Integer    ' Base of array, default 0


Private Sub Class_Initialize()

    theArr = Empty
    Count = 0
    t_Based = 0

End Sub



Public Property Let Based(ByVal value%)

    t_Based = value

End Property


Public Property Get Based() As Integer

    Based = t_Based

End Property



Sub PrintMe(Optional ByVal delim = ", ")

    Const this_me$ = "PrintMe"

    Dim txt$

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    txt = Join(theArr, delim)

    Debug.Print txt

    #If debug_me = True Then
        Call tinyTiming(this_me, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Sub

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Sub



Function Add(ByVal elem) As Variant    ' base 0+

    Const this_me$ = "Add"

    Dim remax&

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

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

    #If debug_me = True Then
        Call tinyTiming(this_me, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(this_me, Err)
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

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    minB = LBound(theArr)

    maxB = UBound(theArr)

    If Based = 0 Then adder = 1 Else adder = 0

    ReDim B(1 To Count, 1 To 1)

    For n = minB To maxB

        B(n + adder, 1) = theArr(n)

    Next n

    AsVertical = B

    #If debug_me = True Then
        Call tinyTiming(this_me, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Function


Function AsHorizontal() As Variant

    Const this_me$ = "AsHorizontal"

    Dim B, minB&, maxB&, n&, adder&

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    minB = LBound(theArr)

    maxB = UBound(theArr)

    If Based = 0 Then adder = 1 Else adder = 0

    ReDim B(1 To 1, 1 To Count)

    For n = minB To maxB

        B(1, n + adder) = theArr(n)

    Next n

    AsHorizontal = B

    #If debug_me = True Then
        Call tinyTiming(this_me, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Function


Sub RndFill(ByVal elements&, _
            Optional ByVal lowerBound& = 0, _
            Optional ByVal upperBound& = 1)

    Const this_me$ = "RndFill"

    Dim B, minB&, maxB&, n&, bounds&, shifter%

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    If t_Based > elements Then

        elements = t_Based

        Debug.Print "Warning: number of elements less then base, used elements = "; elements

    End If

    If t_Based = 0 Then shifter = 1 Else shifter = 0

    elements = elements - shifter

    ReDim theArr(t_Based To elements)

    ' Abstract: Int ((upperBound - lowerBound + 1) * Rnd + lowerBound)

    bounds = upperBound - lowerBound + 1

    For n = t_Based To elements

        theArr(n) = Int(bounds * Rnd + lowerBound)

    Next

    Count = elements - t_Based + shifter

finish:

    #If debug_me = True Then
        Call tinyTiming(this_me & " for " & Count, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Sub

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Sub


Sub OutHorizontal(ByVal startAddr$, _
                  Optional ws As Worksheet, _
                  Optional ByRef rngOut As Range)

    Const this_me$ = "OutHorizontal"

    Dim B, minB&, maxB&, n&, adder&

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    minB = LBound(theArr)

    maxB = UBound(theArr)

    If Based = 0 Then adder = 1 Else adder = 0

    ReDim B(1 To 1, 1 To Count)

    For n = minB To maxB

        B(1, n + adder) = theArr(n)

    Next n

    If ws Is Nothing Then Set ws = ActiveSheet

    Set rngOut = Range(ws.Range(startAddr), ws.Range(startAddr).Offset(, maxB - minB + (1 - adder)))

    rngOut.Value2 = B

    #If debug_me = True Then
        Call tinyTiming(this_me, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Sub

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Sub


Private Sub tinyHandler(ByVal this_me$, ByRef ErrObj)

    Dim txt$

    txt = "[" & className & "." & this_me & "]: " & ErrObj.Description & " (" & ErrObj.Number & ")"

    Debug.Print txt

End Sub


Private Sub tinyTiming(ByVal this_me$, ByVal delta)

    Debug.Print "[" & className & "." & this_me & " timing] = " & Round(delta, 4); " s"

End Sub
