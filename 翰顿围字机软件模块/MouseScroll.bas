Attribute VB_Name = "MouseScroll"
'==============MOUSESCROLL by errorcode100 September 2005==================
'This module allows you to add mouse-scroll-button capabilities to controls.
'See the comments in this module and the example form for info on how to use it
Option Explicit

Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type HOOKEDWND
ItsHWnd As Long
ItsOldProc As Long
ItsScrollMovement As Integer
End Type

Private Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_MOUSEMOVE = &H200

Private HookedWindows() As HOOKEDWND
Private UBoundHookedWindows As Long

'This sub should be called by Form_Load with the hWnd of any Controls
'That you want to be scrollable.
Public Sub AddScrollness(ByVal hwnd As Long)

'Expand the array of scrollable windows to include the new window
UBoundHookedWindows = UBoundHookedWindows + 1
ReDim Preserve HookedWindows(1 To UBoundHookedWindows) As HOOKEDWND
HookedWindows(UBoundHookedWindows).ItsHWnd = hwnd

'Set the Event handeler for the window to the WindowProc function below
HookedWindows(UBoundHookedWindows).ItsOldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

'This function recieves all events for the windows that have been
'scroll-enabled and converts a mouse scroll into a mouse move
'It then sends on the message to the default event handler
'It stores the mouse scroll as 1 or -1 depending on direction
Public Function WindowProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Find which of the scroll-enabled windows this hWnd belongs to
    Dim I As Long
    I = FindHookedWindow(hwnd)

    On Error GoTo Error_Handler
    
    If Msg = WM_MOUSEWHEEL Then
        'Get the cursor's co-ordinates so they can be sent in the mouse move event
        Dim cursor As POINTAPI
        Dim CombinedCo As Long
        GetCursorPos cursor
        CombinedCo = &HFF00 And (cursor.Y * &HFF)
        CombinedCo = CombinedCo Or cursor.X
        
        'Record which direction the scroll was in
        Dim delta As Long
        delta = wParam And &HFFFF0000
        If delta > 0 Then
            HookedWindows(I).ItsScrollMovement = 1
        Else
            HookedWindows(I).ItsScrollMovement = -1
        End If
            
        'Send the mouse move event
        WindowProc = CallWindowProc(HookedWindows(I).ItsOldProc, hwnd, WM_MOUSEMOVE, &O0, CombinedCo)
        
    Else 'Pass on the event to the default event handler
        WindowProc = CallWindowProc(HookedWindows(I).ItsOldProc, hwnd, Msg, wParam, lParam)
    End If
    Exit Function
    
Error_Handler:
    End
End Function

'Call this sub in Form_Unload for each scroll-enalbled control
'or suffer some kind of consequence....
'It puts the event handler back to the original one
'Note that the HookedWindow array still contains this item
Public Sub RemoveScrollness(hwnd As Long)
Dim I As Long
I = FindHookedWindow(hwnd)
    If I = 0 Then Exit Sub
SetWindowLong HookedWindows(I).ItsHWnd, GWL_WNDPROC, HookedWindows(I).ItsOldProc
End Sub

'Retrieve the Scroll Movement for a particular scroll-enabled window
'Call this at the begining of the Control_MouseMove event
'It sets itself to 0 after being called, so you must store the
'value in a variable.
Public Function GetScrollMovement(hwnd As Long) As Integer '0, 1 or -1
Dim I As Long
I = FindHookedWindow(hwnd)
    If I = 0 Then Exit Function
GetScrollMovement = HookedWindows(I).ItsScrollMovement
HookedWindows(I).ItsScrollMovement = 0
End Function

'Returns the index in the HookedWindow array of the window with the
'specified hWnd
Private Function FindHookedWindow(ByVal hwnd As Long) As Long
Dim I As Long
    For I = 1 To UBoundHookedWindows
        If HookedWindows(I).ItsHWnd = hwnd Then
        FindHookedWindow = I
        Exit Function
        End If
    Next I
End Function
