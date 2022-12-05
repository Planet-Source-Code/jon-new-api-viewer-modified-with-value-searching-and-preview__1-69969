Attribute VB_Name = "modAPISubClass"
Option Explicit
'******************************************************************************
' Sizing restriction support module. Sizable forms are hooked to this module
'******************************************************************************

Public Const WinMinW As Long = 7200
#If ISADDIN = 1 Then
  Public Const WinMinH As Long = 7200
#Else
  Public Const WinMinH As Long = 6700
#End If

'*****************
' Local Storage
'*****************
Public m_hWnd As Long     'subclass hook for Main win

'*****************
' API Stuff
'*****************
Public Const GWL_WNDPROC As Long = (-4)
Public Const WM_GETMINMAXINFO As Long = &H24
'
' Screen Points in Pixels
'
Public Type POINTAPI
  X As Long
  Y As Long
End Type
'
' structure for window sizing
'
Type MINMAXINFO
  ptReserved As POINTAPI
  ptMaxSize As POINTAPI
  ptMaxPosition As POINTAPI
  ptMinTrackSize As POINTAPI
  ptMaxTrackSize As POINTAPI
End Type
'
' used to assign new WndProc
'
Public Declare Function SetWindowLong Lib "user32" _
       Alias "SetWindowLongA" _
      (ByVal hWnd As Long, _
       ByVal nIndex As Long, _
       ByVal dwNewLong As Long) As Long
'
' used to invoke old WndProc
'
Public Declare Function CallWindowProc Lib "user32" _
       Alias "CallWindowProcA" _
      (ByVal lpPrevWndFunc As Long, _
       ByVal hWnd As Long, _
       ByVal uMsg As Long, _
       ByVal wParam As Long, _
       ByVal lParam As Long) As Long
'
' used to copy MINMAXINFO structure
'
Public Declare Sub CopyMemory Lib "kernel32" _
      Alias "RtlMoveMemory" _
     (hpvDest As Any, _
      hpvSource As Any, _
      ByVal cbCopy As Long)

'*****************
' HookWin(): Subclass hwnd to VCWndProc
'*****************
Public Sub HookWin(ByVal hWnd As Long, PrvhWnd As Long)
  PrvhWnd = SetWindowLong( _
    hWnd, _
    GWL_WNDPROC, _
    AddressOf SzWndProc)
End Sub

'*****************
' UnhookWin(): remove subclass hook
'*****************
Public Sub UnhookWin(ByVal hWnd As Long, PrvhWnd As Long)
  Call SetWindowLong( _
    hWnd, _
    GWL_WNDPROC, _
    PrvhWnd)
  PrvhWnd = 0
End Sub

'********************************************************************************
'  Subclassing Form. Copy the following, and add handling for additional
'                    messages of interest.
'********************************************************************************
Private Function SzWndProc( _
        ByVal hWnd As Long, _
        ByVal uMsg As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long
  
  Static Calcs As Boolean '1-time flag
  Static SX As Long       'screenx
  Static SY As Long       'screeny
  Static STX As Long      'TwipsPerPixelX
  Static STY As Long      'TwipsPerPixelY
  
  Dim Result As Long
  Dim MnMxInfo As MINMAXINFO
'
' do this one time
'
  If Not Calcs Then
    STX = Screen.TwipsPerPixelX
    STY = Screen.TwipsPerPixelY
    SX = Screen.Width \ STX
    SY = Screen.Height \ STY
    Calcs = True
  End If
'
' handle messages
'
  Select Case hWnd            'check which form
    Case frmAPIViewer.hWnd   'is main VC form
      Select Case uMsg        'check message
        Case WM_GETMINMAXINFO 'sizing
          CopyMemory MnMxInfo, lParam, LenB(MnMxInfo)
          With MnMxInfo
            With .ptMinTrackSize  'set min size
              .X = WinMinW \ STX
              .Y = WinMinH \ STY
            End With
            With .ptMaxPosition
              .X = 0
              .Y = 0
            End With
            With .ptMaxTrackSize
              .X = SX
              .Y = SY
            End With
            With .ptMaxSize
              .X = SX
              .Y = SY
            End With
          End With
          CopyMemory ByVal lParam, MnMxInfo, LenB(MnMxInfo)
          SzWndProc = 0
        Case Else
          SzWndProc = CallWindowProc( _
            m_hWnd, _
            hWnd, _
            uMsg, _
            wParam, _
            lParam)
      End Select 'msg
  End Select 'hwnd
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

