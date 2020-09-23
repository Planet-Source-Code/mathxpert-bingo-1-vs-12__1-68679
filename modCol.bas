Attribute VB_Name = "modCol"
Option Explicit

Public Function RandInt(ByVal Min As Double, ByVal Max As Double) As Double
    Randomize
    RandInt = Int(Rnd * (Max - Min + 1)) + Min
    DoEvents
    Randomize
End Function

Public Function NumberFound(ByVal Num As Double, ByVal Col As Collection) As Boolean
    Dim j As Long
    Dim bFound As Boolean
    
    bFound = False
    For j = 1 To Col.count
        If Num = Col(j) Then
            bFound = True
            Exit For
        End If
    Next
    
    NumberFound = bFound
End Function
