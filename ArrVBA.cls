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

Public Enum SortMethod
    BuiltIn = 0
    Insertion = 1
    Bubble = 2
    Selection = 3
    Quick = 4
End Enum


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



Sub PrintMe(Optional ByVal delim = ", ", _
            Optional ByRef toVar$ = "")

' Prints array to immendiate window [and to variable toVar]
' [delim] - delimiter for elements of array
' [toVar] - result also redirected to this variable

    Const this_me$ = "PrintMe"

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    toVar = Join(theArr, delim)

    Debug.Print toVar

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


Sub OutVertical(ByVal startAddr$, _
                Optional ws As Worksheet, _
                Optional ByRef rngOut As Range, _
                Optional ByVal bold As Boolean = False)

    Const this_me$ = "OutVertical"

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

    If ws Is Nothing Then Set ws = ActiveSheet

    Set rngOut = Range(ws.Range(startAddr), ws.Range(startAddr).Offset(maxB - minB + (1 - adder)))

    rngOut.Value2 = B
    If bold Then rngOut.Font.bold = True

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

Sub OutHorizontal(ByVal startAddr$, _
                  Optional ws As Worksheet, _
                  Optional ByRef rngOut As Range, _
                  Optional ByVal bold As Boolean = False)

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
    If bold Then rngOut.Font.bold = True

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


Sub OutDiagonal(ByVal startAddr$, _
                Optional ws As Worksheet, _
                Optional ByRef rngOut As Range, _
                Optional ByVal bold As Boolean = False)

    Const this_me$ = "OutDiagonal"

    Dim B, minB&, maxB&, n&, adder&

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    minB = LBound(theArr)

    maxB = UBound(theArr)

    If Based = 0 Then adder = 1 Else adder = 0

    ReDim B(1 To Count, 1 To Count)

    For n = minB To maxB

        B(n + adder, n + adder) = theArr(n)

    Next n

    If ws Is Nothing Then Set ws = ActiveSheet

    Set rngOut = Range(ws.Range(startAddr), _
                       ws.Range(startAddr).Offset(maxB - minB + (1 - adder), maxB - minB + (1 - adder)))

    rngOut.Value2 = B
    If bold Then rngOut.Font.bold = True

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

Sub Reverse(Optional ByVal method)

    Const this_me$ = "Reverse"

    Dim n&, m&, minN&, maxN&, B

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    minN = LBound(theArr)

    maxN = UBound(theArr)


    ReDim B(minN To maxN)


    For n = minN To maxN

        B(maxN - n) = theArr(n)

    Next n

    theArr = B

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



Sub Sort(Optional ByVal method% = SortMethod.Insertion, _
         Optional ByVal Reverse = False)

    Select Case method

    Case SortMethod.BuiltIn: theArr = sortExcel(theArr, Reverse:=Reverse)

    Case SortMethod.Insertion: theArr = sortInsertion(theArr, Reverse:=Reverse)

    Case SortMethod.Bubble: theArr = sortBubble(theArr, Reverse:=Reverse)

    Case SortMethod.Selection: theArr = sortSelection(theArr, Reverse:=Reverse)
    
    Case SortMethod.Quick: theArr = sortQuick(theArr, Reverse:=Reverse)

    Case Else: Debug.Print "Sort Method unknown!"

    End Select

End Sub

Private Function sortExcel(Arr, _
                           Optional ByVal Reverse = False) As Variant

End Function

Private Function sortInsertion(Arr, _
                               Optional ByVal Reverse = False) As Variant

    Const this_me$ = "sortInsertion"

    Dim n&, m&, minN&, maxN&, item_to_insert

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    minN = LBound(Arr)

    maxN = UBound(Arr)

    If Reverse = False Then

        If minN + 1 = maxN Then

            If Arr(minN) > Arr(maxN) Then GoSub swap_2elems

            GoTo finish

        End If


        For n = minN + 1 To maxN

            item_to_insert = Arr(n)

            m = n - 1

            Do While m >= 0
                If Arr(m) > item_to_insert Then
                    Arr(m + 1) = Arr(m)
                    m = m - 1
                Else
                    Exit Do
                End If
            Loop

            Arr(m + 1) = item_to_insert

        Next n

    Else

        If minN + 1 = maxN Then

            If Arr(minN) < Arr(maxN) Then GoSub swap_2elems

            GoTo finish

        End If


        For n = minN + 1 To maxN

            item_to_insert = Arr(n)

            m = n - 1

            Do While m >= 0
                If Arr(m) < item_to_insert Then
                    Arr(m + 1) = Arr(m)
                    m = m - 1
                Else
                    Exit Do
                End If
            Loop

            Arr(m + 1) = item_to_insert

        Next n

    End If

finish:

    sortInsertion = Arr

    #If debug_me = True Then
        Call tinyTiming(this_me, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Function

swap_2elems:

    Dim tmp

    tmp = Arr(minN)
    Arr(minN) = Arr(maxN)
    Arr(maxN) = tmp

    Return

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Function


Private Function sortBubble(Arr, _
                            Optional ByVal Reverse = False) As Variant

    Const this_me$ = "sortBubble"

    Dim n&, m&, minN&, maxN&, swapped As Boolean, tmp

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    minN = LBound(Arr)

    maxN = UBound(Arr)

    If Reverse = False Then

        If minN + 1 = maxN Then

            If Arr(minN) > Arr(maxN) Then GoSub swap_2elems

            GoTo finish

        End If

        swapped = True

        Do While swapped

            swapped = False

            For n = minN To maxN - 1

                If Arr(n) > Arr(n + 1) Then
                    tmp = Arr(n)
                    Arr(n) = Arr(n + 1)
                    Arr(n + 1) = tmp
                    swapped = True
                End If

            Next n

        Loop

    Else

        If minN + 1 = maxN Then

            If Arr(minN) < Arr(maxN) Then GoSub swap_2elems

            GoTo finish

        End If

        swapped = True

        Do While swapped

            swapped = False

            For n = minN To maxN - 1

                If Arr(n) < Arr(n + 1) Then
                    tmp = Arr(n)
                    Arr(n) = Arr(n + 1)
                    Arr(n + 1) = tmp
                    swapped = True
                End If

            Next n

        Loop

    End If

finish:

    sortBubble = Arr

    #If debug_me = True Then
        Call tinyTiming(this_me, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Function

swap_2elems:

    tmp = Arr(minN)
    Arr(minN) = Arr(maxN)
    Arr(maxN) = tmp

    Return

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Function


Private Function sortSelection(Arr, _
                               Optional ByVal Reverse = False) As Variant

    Const this_me$ = "sortSelection"

    Dim n&, m&, minN&, maxN&, lowest_ind&, tmp

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    minN = LBound(Arr)

    maxN = UBound(Arr)

    If Reverse = False Then

        If minN + 1 = maxN Then

            If Arr(minN) > Arr(maxN) Then GoSub swap_2elems

            GoTo finish

        End If

        For n = minN To maxN

            lowest_ind = n

            For m = n + 1 To maxN

                If Arr(m) < Arr(lowest_ind) Then lowest_ind = m

            Next m

            tmp = Arr(n)
            Arr(n) = Arr(lowest_ind)
            Arr(lowest_ind) = tmp

        Next n

    Else

        If minN + 1 = maxN Then

            If Arr(minN) < Arr(maxN) Then GoSub swap_2elems

            GoTo finish

        End If

        For n = minN To maxN

            lowest_ind = n

            For m = n + 1 To maxN

                If Arr(m) > Arr(lowest_ind) Then lowest_ind = m

            Next m

            tmp = Arr(n)
            Arr(n) = Arr(lowest_ind)
            Arr(lowest_ind) = tmp

        Next n

    End If

finish:

    sortSelection = Arr

    #If debug_me = True Then
        Call tinyTiming(this_me, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Function

swap_2elems:

    tmp = Arr(minN)
    Arr(minN) = Arr(maxN)
    Arr(maxN) = tmp

    Return

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Function


Private Function sortQuick(Arr, _
                           Optional ByVal Reverse = False, _
                           Optional ByVal minN& = -1, _
                           Optional ByVal maxN& = -1) As Variant

    Const this_me$ = "sortQuick"

    Dim n&, midVal, tmpMin, tmpMax, tmp, needShow As Boolean

    On Error GoTo errHandler
    
    needShow = False
    
    If minN = -1 Then
        #If debug_me = True Then
            Dim starts
            starts = Timer
        #End If
        minN = LBound(Arr)
        needShow = True
    End If

    tmpMin = minN

    If maxN = -1 Then maxN = UBound(Arr)
    
    
    tmpMax = maxN

    If Reverse = False Then

        midVal = Arr((minN + maxN) \ 2)

        While tmpMin <= tmpMax

            While (Arr(tmpMin) < midVal And tmpMin < maxN)

                tmpMin = tmpMin + 1

            Wend


            While (midVal < Arr(tmpMax) And tmpMax > minN)

                tmpMax = tmpMax - 1

            Wend


            If tmpMin <= tmpMax Then

                tmp = Arr(tmpMin)
                Arr(tmpMin) = Arr(tmpMax)
                Arr(tmpMax) = tmp
                tmpMin = tmpMin + 1
                tmpMax = tmpMax - 1

            End If

        Wend

        If minN < tmpMax Then Call sortQuick(Arr, minN:=minN, maxN:=tmpMax, Reverse:=Reverse)
        If tmpMin < maxN Then Call sortQuick(Arr, minN:=tmpMin, maxN:=maxN, Reverse:=Reverse)

    Else



    End If

finish:

    sortQuick = Arr

    #If debug_me = True Then
        If needShow Then Call tinyTiming(this_me, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Function


Private Sub tinyHandler(ByVal this_me$, ByRef ErrObj)

    Dim txt$

    txt = "[" & className & "." & this_me & "]: " & ErrObj.Description & " (" & ErrObj.Number & ")"

    Debug.Print txt

End Sub


Private Sub tinyTiming(ByVal this_me$, ByVal delta)

    Debug.Print "[" & className & "." & this_me & " timing] = " & Round(delta, 4); " s"

End Sub
