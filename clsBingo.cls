VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBingo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim i As Long
Dim i0 As Long

Dim UsedCol As Collection

Public Sub GoNext()
    i = i + 1
End Sub

Public Function RandomNumber() As Double
    Dim r As Double
    Dim count As Double
    Dim lMin As Long
    Dim lMax As Long
    Dim j As Long
    
    ' Remember, we are picking random numbers down the columns WITHOUT repeats!
    If i > 4 Or (UsedCol Is Nothing) Then
        i = 0
        If Not UsedCol Is Nothing Then i0 = i0 + 1
        Set UsedCol = New Collection
    End If
    
    lMin = i0 * 15 + 1
    lMax = (i0 + 1) * 15
    
    count = 0
    DoEvents
    r = RandInt(1, 15 - UsedCol.count)
    For j = lMin To lMax
        If Not NumberFound(j, UsedCol) Then
            count = count + 1
            If r = count Then
                RandomNumber = j
                UsedCol.Add j
                i = i + 1
                Exit For
            End If
        End If
    Next
End Function
