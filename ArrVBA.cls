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
'Option Base 0    ' 0 it's default, but added explicitly

#Const debug_me = False     ' Compiler directive  - if debug and timing needed

Const className$ = "ArrVBA"    ' Used in error handling process messages for this module

Public Enum SortMethod    ' sort methods allowed
    [_First] = 1    ' [_First]/[_Last] - for easy for..next
    Insertion = 1
    Bubble = 2
    Selection = 3
    Quick = 4
    Heap = 5    ' default
    [_Last] = 5
End Enum

' private class-level variables:
Private t_Count As Long     ' Counter of elements in ArrVBA instance
Private theArr As Variant    ' Variant Array
Private t_Based As Integer    ' Base of array, default 0


Private Sub Class_Initialize()

    Call Clear(keepBase:=False)

End Sub


Public Sub Clear(Optional ByVal keepBase As Boolean = True)

    theArr = Empty

    t_Count = 0    'CVErr(xlErrNA)

    If keepBase = False Then t_Based = 0

End Sub


Public Property Let Based(ByVal value%)

    If t_Based <> value Then

        Call reBaseArr(t_Based, value)

    End If

    t_Based = value

End Property


Public Property Get Based() As Integer

    Based = t_Based

End Property

Public Property Get count() As Long

    count = t_Count

End Property

Public Property Get Arr() As Variant

    Arr = theArr

End Property

Public Property Get MaxValue() As Variant

    MaxValue = getMaxValue

End Property

Public Property Get MinValue() As Variant

    MinValue = getMinValue

End Property

Private Function getMaxValue() As Variant

    Dim n&, needAnotherTry As Boolean, theMax, test

    On Error GoTo errHandler

    getMaxValue = Empty

    If IsEmpty(theArr) Then GoTo finish

    needAnotherTry = False

    getMaxValue = Application.WorksheetFunction.max(theArr)

    If needAnotherTry = True Then

        theMax = -9E+99

        For n = LBound(theArr) To UBound(theArr)

            test = theArr(n)

            If isOKNumeric(test) Then

                If theMax < test Then theMax = test

            End If

        Next n

        getMaxValue = theMax

    End If

finish:

    On Error GoTo 0

    Exit Function

errHandler:

    #If debug_me = True Then
        If Err.Number <> 1004 Then Call tinyHandler("_getMaxValue", Err)
    #End If
    needAnotherTry = True
    Err.Clear
    Resume Next

End Function


Private Function getMinValue() As Variant

    Dim n&, needAnotherTry As Boolean, theMin, test

    On Error GoTo errHandler

    getMinValue = Empty

    If IsEmpty(theArr) Then GoTo finish

    needAnotherTry = False

    getMinValue = Application.WorksheetFunction.Min(theArr)

    If needAnotherTry = True Then

        theMin = 9E+99

        For n = LBound(theArr) To UBound(theArr)

            test = theArr(n)

            If isOKNumeric(test) Then

                If theMin > test Then theMin = test

            End If

        Next n

        getMinValue = theMin

    End If

finish:

    On Error GoTo 0

    Exit Function

errHandler:

    #If debug_me = True Then
        If Err.Number <> 1004 Then Call tinyHandler("_getMinValue", Err)
    #End If
    needAnotherTry = True
    Err.Clear
    Resume Next

End Function



Private Function isOKNumeric(ByVal value As Variant) As Boolean

' internal tester for non-error and  numeric value

    If IsError(value) = False Then

        isOKNumeric = IsNumeric(value)

    End If

End Function

Public Sub PrintMe(Optional ByVal delim = ", ", _
                   Optional ByRef toVar$ = "")

' Prints array to immendiate window [and to variable toVar]
' [delim] - delimiter for elements of array
' [toVar] - result also redirected to this variable

    Const this_me$ = "PrintMe"

    Dim n&, needAnotherTry As Boolean, B, test

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    needAnotherTry = False

    toVar = Join(theArr, delim)

    If needAnotherTry = True Then

        B = copyArr(theArr)    ' working with copy to make no changes in main array

        For n = LBound(B) To UBound(B)
            test = B(n)
            If IsError(test) Then B(n) = CStr(test)    ' converting to string
        Next n

        toVar = Join(B, delim)

    End If

    Debug.Print toVar


    #If debug_me = True Then
        Call tinyTiming(this_me, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Sub

errHandler:

    #If debug_me = True Then
        If Err.Number <> 13 Then Call tinyHandler(this_me, Err)
    #End If
    needAnotherTry = True
    Err.Clear
    Resume Next

End Sub


Private Function copyArr(ByVal Arr) As Variant

    copyArr = Arr

End Function

Function Add(ByVal elem)

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

        t_Count = 1

    Else

        remax = UBound(theArr) + 1

        ReDim Preserve theArr(t_Based To remax)

        theArr(remax) = elem

        t_Count = t_Count + 1

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

Function AddArr(ByVal elems) As Variant    ' base 0+

    Const this_me$ = "AddArr"

    Dim remax&, minE&, maxE&, n&

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    minE = LBound(elems)

    maxE = UBound(elems)

    If IsEmpty(theArr) Then

        theArr = elems

        ReDim Preserve theArr(t_Based To t_Based + (maxE - minE))

        t_Count = maxE - minE + 1

    Else

        remax = UBound(theArr)

        ReDim Preserve theArr(t_Based To remax + maxE + 1)

        For n = minE To maxE

            theArr(remax + n + 1) = elems(n)
            t_Count = t_Count + 1

        Next n

    End If

    AddArr = theArr

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

Function AsString(Optional ByVal delim = ",") As String

    Const this_me$ = "AsString"

    On Error GoTo errHandler

    If IsEmpty(theArr) = True Then
        AsString = vbNullString
    Else
        AsString = Join(theArr, delim)
    End If

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Function

Function AsStringArr() As String()

    Const this_me$ = "AsStringArr"

    Dim B() As String, n&, minN&, maxN&

    On Error GoTo errHandler

    minN = LBound(theArr)

    maxN = UBound(theArr)

    ReDim B(minN To maxN)

    For n = minN To maxN
        B(n) = CStr(theArr(n))
    Next n

    AsStringArr = B

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

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

    adder = defineAdder(Based)

    ReDim B(1 To t_Count, 1 To 1)

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

    adder = defineAdder(Based)

    ReDim B(1 To 1, 1 To t_Count)

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

    Dim B, minB&, maxB&, n&, bounds&, adder%, lastElem

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    t_Count = elements
    
    adder = defineAdder(t_Based)
    
    lastElem = elements - adder
    
    ReDim theArr(t_Based To lastElem)

    ' Abstract: Int ((upperBound - lowerBound + 1) * Rnd + lowerBound)

    bounds = upperBound - lowerBound + 1

    For n = t_Based To lastElem

        theArr(n) = Int(bounds * Rnd + lowerBound)

    Next


finish:

    #If debug_me = True Then
        Call tinyTiming(this_me & " for " & t_Count, Timer - starts)
    #End If

    On Error GoTo 0

    Exit Sub

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Sub


Sub ValueFill(ByVal elements&, ByVal FillWith As Variant, _
              Optional ByVal lowerBound& = 0, _
              Optional ByVal upperBound& = 1)

    Const this_me$ = "ValueFill"

    Dim B, minB&, maxB&, n&, adder%, lastElem&

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    t_Count = elements
    
    adder = defineAdder(t_Based)
    
    lastElem = elements - adder
    
    ReDim theArr(t_Based To lastElem)

    For n = t_Based To lastElem

        theArr(n) = FillWith

    Next


finish:

    #If debug_me = True Then
        Call tinyTiming(this_me & " for " & t_Count, Timer - starts)
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

    adder = defineAdder(Based)

    ReDim B(1 To t_Count, 1 To 1)

    For n = minB To maxB

        B(n + adder, 1) = theArr(n)

    Next n

    If ws Is Nothing Then Set ws = ActiveSheet

    Set rngOut = Range(ws.Range(startAddr), ws.Range(startAddr).Offset(maxB - minB))

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

    adder = defineAdder(Based)

    ReDim B(1 To 1, 1 To t_Count)

    For n = minB To maxB

        B(1, n + adder) = theArr(n)

    Next n

    If ws Is Nothing Then Set ws = ActiveSheet

    Set rngOut = Range(ws.Range(startAddr), ws.Range(startAddr).Offset(, maxB - minB))

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

    adder = defineAdder(Based)

    ReDim B(1 To t_Count, 1 To t_Count)

    For n = minB To maxB

        B(n + adder, n + adder) = theArr(n)

    Next n

    If ws Is Nothing Then Set ws = ActiveSheet

    Set rngOut = Range(ws.Range(startAddr), _
                       ws.Range(startAddr).Offset(maxB - minB, maxB - minB))

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

Sub ReverseArr(Optional ByRef Arr)

    Const this_me$ = "ReverseArr"

    Dim n&, m&, minN&, maxN&, B, maxM%, reverseClass

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    reverseClass = False

    If IsMissing(Arr) Then
        Arr = theArr
        reverseClass = True
    End If

    minN = LBound(Arr)

    maxN = UBound(Arr)

    maxM = maxN

    ReDim B(minN To maxN)


    For n = minN To maxN

        B(maxM) = Arr(n)

        maxM = maxM - 1

    Next n

    Arr = B

    If reverseClass = True Then

        theArr = Arr

    End If

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



Sub Sort(Optional ByVal Method% = SortMethod.Heap, _
         Optional ByVal Reverse = False)

    Select Case Method

    Case SortMethod.Insertion: theArr = sortInsertion(theArr, Reverse:=Reverse)

    Case SortMethod.Bubble: theArr = sortBubble(theArr, Reverse:=Reverse)

    Case SortMethod.Selection: theArr = sortSelection(theArr, Reverse:=Reverse)

    Case SortMethod.Quick: theArr = sortQuick(theArr, Reverse:=Reverse)

    Case SortMethod.Heap: theArr = sortHeap(theArr, Reverse:=Reverse)

    Case Else: Debug.Print "Sort Method unknown!"

    End Select

End Sub



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

            Do While m >= t_Based
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

            Do While m >= t_Based
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

    midVal = Arr((minN + maxN) \ 2)

    If Reverse = False Then

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

        While tmpMin <= tmpMax

            While (Arr(tmpMin) > midVal And tmpMin < maxN)

                tmpMin = tmpMin + 1

            Wend


            While (midVal > Arr(tmpMax) And tmpMax > minN)

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

    End If


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


Private Function sortHeap(Arr, _
                          Optional ByVal Reverse = False) As Variant

    Const this_me$ = "sortHeap"

    Dim n&, m&, minN&, maxN&, i&, idxArr, needRebase As Boolean, oldBase%

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    needRebase = False

    If t_Based <> 0 Then

        oldBase = t_Based

        Call reBaseArr(t_Based, 0, Arr)

        needRebase = True

    End If


    minN = LBound(Arr)

    maxN = UBound(Arr)

    n = t_Count


    For i = n To minN Step -1

        Call heapify(Arr, n, i)

    Next i


    For i = n - 1 To minN + 1 Step -1

        Call swapPair(Arr, i, minN)

        Call heapify(Arr, i, minN)

    Next i

    If Reverse = True Then

        Call ReverseArr(Arr)

    End If


    If needRebase = True Then

        Call reBaseArr(0, oldBase, Arr)

    End If


    sortHeap = Arr

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


Private Function heapify(ByRef nums, ByVal heap_size&, ByVal root_index&):

    Dim largest&, left_child&, right_child&
    ' let index of max element - is root index
    largest = root_index
    left_child = (2 * root_index) + 1
    right_child = (2 * root_index) + 2

    ' updating left if acceptable
    If left_child < heap_size Then
        If nums(left_child) > nums(largest) Then
            largest = left_child
        End If
    End If

    ' updating right if acceptable
    If right_child < heap_size Then
        If nums(right_child) > nums(largest) Then
            largest = right_child
        End If
    End If

    ' If largest not root_index already - swap'em
    If largest <> root_index Then
        Call swapPair(nums, root_index, largest)
        ' Heapify the new root element to ensure it's the largest
        Call heapify(nums, heap_size, largest)
    End If

End Function



Function isSorted(Optional ByRef isAscending As Boolean) As Boolean

    Const this_me$ = "isSorted"

    Dim n&, minN&, maxN&

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    minN = LBound(theArr)

    maxN = UBound(theArr)

    If theArr(minN) < theArr(maxN) Then isAscending = True Else isAscending = False

    isSorted = True    ' optimistic start

    If isAscending = True Then

        For n = minN To maxN - 1

            If theArr(n) > theArr(n + 1) Then
                isSorted = False
                Exit For
            End If

        Next n

    Else

        For n = maxN To minN + 1 Step -1

            If theArr(n) > theArr(n - 1) Then
                isSorted = False
                Exit For
            End If

        Next n

    End If

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


Private Sub swapPair(idxArr, ByVal i&, ByVal j&)

    Dim tmp&

    tmp = idxArr(i)

    idxArr(i) = idxArr(j)

    idxArr(j) = tmp

End Sub

Private Sub reBaseArr(ByVal fromBase%, ByVal toBase%, Optional ByRef Arr)

    Const this_me$ = "reBaseArr"

    Dim n&, maxNew&, rebaseClass As Boolean

    On Error GoTo errHandler

    #If debug_me = True Then
        Dim starts
        starts = Timer
    #End If

    rebaseClass = False

    If IsMissing(Arr) Then
        Arr = theArr
        rebaseClass = True
    End If


    If IsEmpty(Arr) = True Then

        'nothing to do here

        GoTo finish

    Else

        maxNew = UBound(Arr) + (toBase - fromBase)

        ReDim Preserve Arr(toBase To maxNew)

    End If


    If rebaseClass = True Then

        theArr = Arr

    End If

finish:

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


Public Function FilterArr(ByVal searchFor$, _
                          Optional ByVal Include As Boolean = True, _
                          Optional ByVal Compare% = 1) As Variant

' Filtering value to string array

    Const this_me$ = "FilterArr"

    On Error GoTo errHandler

    FilterArr = Filter(theArr, searchFor, Include:=Include, Compare:=Compare)

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Function

Public Function isIncludeTemplate(ByVal template$, _
                                  Optional ByRef foundByTemplate, _
                                  Optional ByRef foundCounter& = 0) As Boolean

    Const this_me$ = "isIncludeTemplate"

    On Error GoTo errHandler

    isIncludeTemplate = False

    foundByTemplate = FilterArr(template)

    If UBound(foundByTemplate) <> -1 Then

        isIncludeTemplate = True

        foundCounter = UBound(foundByTemplate) - LBound(foundByTemplate) + 1

    Else

        foundByTemplate = Empty

    End If

    On Error GoTo 0

    Exit Function

errHandler:

    Call tinyHandler(this_me, Err)
    Err.Clear
    Resume Next

End Function


Private Function defineAdder(ByVal Based) As Integer

    Dim adder%

    If Based = 0 Then
        adder = 1
    Else
        If Based = 1 Then
            adder = 0
        Else
            adder = -Based + 1
        End If
    End If

    defineAdder = adder

End Function

Private Sub tinyHandler(ByVal this_me$, ByRef ErrObj)

    Dim txt$

    txt = "![" & className & "." & this_me & "]: " & ErrObj.Description & " (" & ErrObj.Number & ")"

    Debug.Print txt

End Sub


Private Sub tinyTiming(ByVal this_me$, ByVal delta)

    Debug.Print "[" & className & "." & this_me & " timing] = " & Round(delta, 4); " s"

End Sub
