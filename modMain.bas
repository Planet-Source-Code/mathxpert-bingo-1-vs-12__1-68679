Attribute VB_Name = "modMain"
Private Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
    On Error Resume Next
    Dim iccex As tagInitCommonControlsEx
    ' Ensure CC available:
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    InitCommonControlsVB = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Sub Main()
    InitCommonControlsVB
    
    '
    ' Start your application here:
    '
    frmMain.Show
    
End Sub

