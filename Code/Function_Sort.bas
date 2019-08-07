Attribute VB_Name = "Function_Sort"
Option Explicit
'http://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)&p=2909260#post2909260
' Sort a 2-dimensional array on either dimension
' Omit plngLeft & plngRight; they are used internally during recursion
' Sample usage to sort on column 4
' Dim MyArray(1 to 1000, 1 to 5) As Long
' QuickSort2 MyArray, 2, 4
' Dim MyArray(1 to 5, 1 to 1000) As Long
' QuickSort2 MyArray, 1, 4
Public Sub QuickSort2(ByRef pvarArray As Variant, plngDim As Long, plngCol As Long, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant
    Dim c As Long
    Dim cMin As Long
    Dim cMax As Long
    
    cMin = LBound(pvarArray, plngDim)
    cMax = UBound(pvarArray, plngDim)
    Select Case plngDim
        Case 1
            If plngRight = 0 Then
                plngLeft = LBound(pvarArray, 2)
                plngRight = UBound(pvarArray, 2)
            End If
            lngFirst = plngLeft
            lngLast = plngRight
            varMid = pvarArray(plngCol, (plngLeft + plngRight) \ 2)
            Do
                Do While pvarArray(plngCol, lngFirst) < varMid And lngFirst < plngRight
                    lngFirst = lngFirst + 1
                Loop
                Do While varMid < pvarArray(plngCol, lngLast) And lngLast > plngLeft
                    lngLast = lngLast - 1
                Loop
                If lngFirst <= lngLast Then
                    For c = cMin To cMax
                        varSwap = pvarArray(c, lngFirst)
                        pvarArray(c, lngFirst) = pvarArray(c, lngLast)
                        pvarArray(c, lngLast) = varSwap
                    Next
                    lngFirst = lngFirst + 1
                    lngLast = lngLast - 1
                End If
            Loop Until lngFirst > lngLast
            If plngLeft < lngLast Then QuickSort2 pvarArray, plngDim, plngCol, plngLeft, lngLast
            If lngFirst < plngRight Then QuickSort2 pvarArray, plngDim, plngCol, lngFirst, plngRight
        Case 2
            If plngRight = 0 Then
                plngLeft = LBound(pvarArray, 1)
                plngRight = UBound(pvarArray, 1)
            End If
            lngFirst = plngLeft
            lngLast = plngRight
            varMid = pvarArray((plngLeft + plngRight) \ 2, plngCol)
            Do
                Do While pvarArray(lngFirst, plngCol) < varMid And lngFirst < plngRight
                    lngFirst = lngFirst + 1
                Loop
                Do While varMid < pvarArray(lngLast, plngCol) And lngLast > plngLeft
                    lngLast = lngLast - 1
                Loop
                If lngFirst <= lngLast Then
                    For c = cMin To cMax
                        varSwap = pvarArray(lngFirst, c)
                        pvarArray(lngFirst, c) = pvarArray(lngLast, c)
                        pvarArray(lngLast, c) = varSwap
                    Next
                    lngFirst = lngFirst + 1
                    lngLast = lngLast - 1
                End If
            Loop Until lngFirst > lngLast
            If plngLeft < lngLast Then QuickSort2 pvarArray, plngDim, plngCol, plngLeft, lngLast
            If lngFirst < plngRight Then QuickSort2 pvarArray, plngDim, plngCol, lngFirst, plngRight
    End Select
End Sub

