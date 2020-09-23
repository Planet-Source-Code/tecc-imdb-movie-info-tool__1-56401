Attribute VB_Name = "modWindows"
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function SetFocusA Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type WINDOW_HANDLER
    pHwnd As Long
    pFrm As frmIMDB
    pID As Long
    
End Type
Public WNDWS() As WINDOW_HANDLER




Public Function SWF(hwns As Long)
    SetFocus hwns
End Function


Public Function NewWindow() As Long
Dim nn As Long
Set WNDWS(UBound(WNDWS)).pFrm = New frmIMDB
Load WNDWS(UBound(WNDWS)).pFrm
NewWindow = UBound(WNDWS)
'load a new one
nn = UBound(WNDWS) + 1
ReDim Preserve WNDWS(nn)
End Function
