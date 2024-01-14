Attribute VB_Name = "DryAlgorithm"
Option Explicit

Sub Swap(ByRef A As Variant, ByRef B As Variant)
    Dim Tmp As Variant
    
    Tmp = A
    A = B
    B = Tmp
End Sub

Sub QuickSort(ByRef Items As Variant, LowerBound As Integer, UpperBound As Integer)
    Dim Pivot As Variant
    Dim Low As Integer
    Dim High As Integer

    If UpperBound <= LowerBound Then Exit Sub

    Pivot = Items(LowerBound + (UpperBound - LowerBound + 1) \ 2)
    Low = LowerBound
    High = UpperBound

    While Low <= High
        While Items(Low) < Pivot
            Low = Low + 1
        Wend
        While Items(High) > Pivot
            High = High - 1
        Wend
        If Low <= High Then
            Swap Items(Low), Items(High)
            Low = Low + 1
            High = High - 1
        End If
    Wend

    QuickSort Items, LowerBound, High
    QuickSort Items, Low, UpperBound
End Sub

