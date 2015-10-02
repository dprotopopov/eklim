Attribute VB_Name = "QuickSortModule"
Option Explicit

'http://www.xtremevbtalk.com/showthread.php?t=150521
Public Sub QuickSort(vArray As Variant, l As Integer, r As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim x
    Dim Y

    i = l
    j = r
    x = vArray((l + r) / 2)

    While (i <= j)

        While (UCase(vArray(i)) < UCase(x) And i < r)
            i = i + 1
        Wend

        While (UCase(x) < UCase(vArray(j)) And j > l)
            j = j - 1
        Wend

        If (i <= j) Then
            Y = vArray(i)
            vArray(i) = vArray(j)
            vArray(j) = Y
            i = i + 1
            j = j - 1
        End If

    Wend

    If (l < j) Then QuickSort vArray, l, j
    If (i < r) Then QuickSort vArray, i, r

End Sub
