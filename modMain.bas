Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function MoveWindow& Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long)
Public Declare Function SetActiveWindow& Lib "user32" (ByVal hWnd As Long)
Public Declare Function SetWindowPos& Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10

Public Function GetWindowCaption(wHWnd As Long) As String

    Dim TxtLen  As Long
    Dim WinTxt  As String * 255
    
    TxtLen = GetWindowTextLength(wHWnd) + 1
    
    Call GetWindowText(wHWnd, WinTxt, TxtLen)
    
    GetWindowCaption = WinTxt
    
End Function

Public Sub SendToBottom(wHWnd As Long)

    Call SetWindowPos(wHWnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE)
    
End Sub

Public Sub BringToFront(wHWnd As Long)

    Call SetWindowPos(wHWnd, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE)
    
End Sub
