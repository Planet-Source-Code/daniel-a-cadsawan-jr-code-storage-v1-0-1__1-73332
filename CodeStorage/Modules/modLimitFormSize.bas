Attribute VB_Name = "modLimitFormSize"
Option Explicit

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc _
                Lib "user32" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hWnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long
Private Declare Function DefWindowProc _
                Lib "user32" _
                Alias "DefWindowProcA" (ByVal hWnd As Long, _
                                        ByVal wMsg As Long, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)
Private Const GWL_WNDPROC = (-4)
Private Const WM_SIZING = &H214
Private Const WMSZ_LEFT = 1
Private Const WMSZ_RIGHT = 2
Private Const WMSZ_TOP = 3
Private Const WMSZ_TOPLEFT = 4
Private Const WMSZ_TOPRIGHT = 5
Private Const WMSZ_BOTTOM = 6
Private Const WMSZ_BOTTOMLEFT = 7
Private Const WMSZ_BOTTOMRIGHT = 8
Private Const MIN_WIDTH = 700      ' The minimum width in pixels  '
Private Const MIN_HEIGHT = 500    ' The minimum height in pixels '
Private Const MAX_WIDTH = 2000      ' The maximum width in pixels  '
Private Const MAX_HEIGHT = 2000    ' The maximum height in pixels '
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type
Private mPrevProc As Long
Public Sub Hook(hWnd As Long)
    mPrevProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf NewWndProc)
End Sub
Public Sub Unhook(hWnd As Long)

    Call SetWindowLong(hWnd, GWL_WNDPROC, mPrevProc)
    mPrevProc = 0&
End Sub
Public Function NewWndProc(ByVal hWnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
    On Error Resume Next

    Dim r As RECT
    If uMsg = WM_SIZING Then
        Call CopyMemory(r, ByVal lParam, Len(r))

        ' Keep the form only at least as wide as MIN_WIDTH '
        If (r.Right - r.Left < MIN_WIDTH) Then
            Select Case wParam
                Case WMSZ_LEFT, WMSZ_BOTTOMLEFT, WMSZ_TOPLEFT
                    r.Left = r.Right - MIN_WIDTH
                Case WMSZ_RIGHT, WMSZ_BOTTOMRIGHT, WMSZ_TOPRIGHT
                    r.Right = r.Left + MIN_WIDTH
            End Select
        End If

        ' Keep the form only at least as tall as MIN_HEIGHT '
        If (r.Bottom - r.Top < MIN_HEIGHT) Then
            Select Case wParam
                Case WMSZ_TOP, WMSZ_TOPLEFT, WMSZ_TOPRIGHT
                    r.Top = r.Bottom - MIN_HEIGHT
                Case WMSZ_BOTTOM, WMSZ_BOTTOMLEFT, WMSZ_BOTTOMRIGHT
                    r.Bottom = r.Top + MIN_HEIGHT
            End Select
        End If

        ' Keep the form only as wide as MAX_WIDTH '
        If (r.Right - r.Left > MAX_WIDTH) Then
            Select Case wParam
                Case WMSZ_LEFT, WMSZ_BOTTOMLEFT, WMSZ_TOPLEFT
                    r.Left = r.Right - MAX_WIDTH
                Case WMSZ_RIGHT, WMSZ_BOTTOMRIGHT, WMSZ_TOPRIGHT
                    r.Right = r.Left + MAX_WIDTH
            End Select
        End If
        ' Keep the form only as tall as MAX_HEIGHT '
        If (r.Bottom - r.Top > MAX_HEIGHT) Then
            Select Case wParam
                Case WMSZ_TOP, WMSZ_TOPLEFT, WMSZ_TOPRIGHT
                    r.Top = r.Bottom - MAX_HEIGHT
                Case WMSZ_BOTTOM, WMSZ_BOTTOMLEFT, WMSZ_BOTTOMRIGHT
                    r.Bottom = r.Top + MAX_HEIGHT
            End Select
        End If
        Call CopyMemory(ByVal lParam, r, Len(r))
        NewWndProc = 0&
        Exit Function
    End If
    If mPrevProc > 0& Then
        NewWndProc = CallWindowProc(mPrevProc, hWnd, uMsg, wParam, lParam)
    Else
        NewWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
End Function



