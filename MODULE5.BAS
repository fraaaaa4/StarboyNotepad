Attribute VB_Name = "Module5"




Private Const GWL_STYLE = (-16)
Private Const WS_BORDER = &H800000
Private Declare Function GetWindowLong Lib "user32" _
        Alias "GetWindowLongA" (ByVal hWnd As Long, _
        ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" _
        Alias "SetWindowLongA" (ByVal hWnd As Long, _
        ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
        ByVal X As Long, ByVal Y As Long, _
        ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub newborderstyle()
        Dim lStyle As Long
 
        lStyle = GetWindowLong(Form1.hWnd, GWL_STYLE)
        lStyle = lStyle And (Not WS_BORDER)
        SetWindowLong Form1.hWnd, GWL_STYLE, lStyle
        SetWindowPos Form1.hWnd, 0&, 0&, 0&, 0&, 0&, _
                SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER


End Sub


Public Sub oldborderstyle()
Dim lStyle As Long
 
        lStyle = GetWindowLong(Form1.hWnd, GWL_STYLE)
        lStyle = lStyle And (WS_BORDER)
        SetWindowPos Form1.hWnd, 0&, 0&, 0&, 0&, 0&, _
                SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER


End Sub
