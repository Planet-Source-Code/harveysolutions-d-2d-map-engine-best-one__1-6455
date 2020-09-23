Attribute VB_Name = "Module1"
Option Explicit
 
'COULD BE USEFULL
'Public Const MOUSEEVENTF_LEFTDOWN = &H2
'Public Const MOUSEEVENTF_LEFTUP = &H4
'Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
'Public Const MOUSEEVENTF_MIDDLEUP = &H40
'Public Const MOUSEEVENTF_RIGHTDOWN = &H8
'Public Const MOUSEEVENTF_RIGHTUP = &H10
'Public Const MOUSEEVENTF_MOVE = &H1

'PROGRAM CONSTANT
Public Const TILEWIDTH = 130
Public Const TILEHEIGHT = 130
Public Const SCROLLSPEED = 60

'VB/WINDOW CONSTANT
Public Const LR_LOADFROMFILE = &H10
Public Const LR_CREATEDIBSECTION = &H2000
Public Const SRCCOPY = &HCC0020

'VB/WINDOW TYPE
Public Type BITMAP
        bmType          As Long
        bmWidth         As Long
        bmHeight        As Long
        bmWidthBytes    As Long
        bmPlanes        As Integer
        bmBitsPixel     As Integer
        bmBits          As Long
End Type
'Public Type RECT
'        Left As Long
'        Top As Long
'        Right As Long
'        Bottom As Long
'End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type
'*************************************


Public Mousex As Long
Public Mousey As Long
Public g_Sensitivity
Public Const BufferSize = 10

Public EventHandle As Long
'Public Drawing As Boolean
Public Suspended As Boolean

Public procOld As Long

' Windows API declares and constants

Public Const GWL_WNDPROC = (-4)
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212
Public Const WM_SYSCOMMAND = &H112


'PROGRAM TYPE
Public Type MessageDisplay
        MessageSurface As DirectDrawSurface7
        MessageText As String
        Position As RECT
        DisplayTime As Long
        DisplayedTime As Long
End Type

' Various Standard API Functions for manipulation of DCs and Bitmaps
'*******************************************************************
'THOSE ONES COULD BE USEFULL
'Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
'Public Declare Function lstrcpy Lib "kernel32" (ByVal lpszDestinationString1 As Any, ByVal lpszSourceString2 As Any) As Long
'Public Declare Function timeGetTime Lib "winmm.dll" () As Long
'Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
'Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)
'Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
        


'************************************************************************************************
'function fit any type of surface
'Tiled or not Tiled
'************************************************************************************************
Public Function LoadBitmapIntoDXS(DXObject As DirectDraw7, ByVal BMPFile As String, ByVal NW As Integer, ByVal NH As Integer, ByVal StretchV) As DirectDrawSurface7
    Dim hBitmap As Long                 ' Handle on bitmap
    Dim dBitmap As BITMAP               ' Handle on bitmap descriptor
    Dim TempDXD As DDSURFACEDESC2       ' Surface description
    Dim TempDXS As DirectDrawSurface7  ' Created surface
    Dim dcBitmap As Long                ' Handle on image
    Dim dcDXS As Long                   ' Handle on surface context
    Dim ddck As DDCOLORKEY
    Dim i, i2
    ddck.low = 0
    ddck.high = 0
    
    ' Load bitmap
    hBitmap = LoadImage(ByVal 0&, BMPFile, 0, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
    ' Get bitmap descriptor
    GetObject hBitmap, Len(dBitmap), dBitmap
    ' Fill DX surface description
    With TempDXD
        '.dwSize = Len(TempDXD)
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lWidth = (dBitmap.bmWidth / StretchV) * NW
        .lHeight = (dBitmap.bmHeight / StretchV) * NH
    End With
    ' Create DX surface
    Set TempDXS = DXObject.CreateSurface(TempDXD)
    
    ' Create API memory DC
    dcBitmap = CreateCompatibleDC(ByVal 0&)
    ' Select the bitmap into API memory DC
    SelectObject dcBitmap, hBitmap
    ' Restore DX surface
    TempDXS.restore
    ' Get DX surface API DC
    dcDXS = TempDXS.GetDC()
    
    ' Blit BMP from API DC into DX DC using standard API BitBlt
    For i = 0 To NH
       For i2 = 0 To NW
        StretchBlt dcDXS, i2 * (dBitmap.bmWidth / StretchV), i * (dBitmap.bmHeight / StretchV), (dBitmap.bmWidth / StretchV), (dBitmap.bmHeight / StretchV), dcBitmap, 0, 0, dBitmap.bmWidth, dBitmap.bmHeight, SRCCOPY
       Next
     Next
    ' Cleanup
    TempDXS.ReleaseDC dcDXS
    DeleteDC dcBitmap
    DeleteObject hBitmap
    TempDXS.SetColorKey DDCKEY_SRCBLT, ddck
    ' Return created DX surface
    Set LoadBitmapIntoDXS = TempDXS
End Function
'function to scroll the map
Public Function GetNewViewPortPosition(ByVal direction, ByVal num, ByVal w1, ByVal w2) As Long
If direction > 0 Then
  If num + w1 < w2 - SCROLLSPEED + 1 Then
    num = num + direction * SCROLLSPEED
  End If
Else
  If num > SCROLLSPEED - 1 Then
     num = num + direction * SCROLLSPEED
  End If
End If
GetNewViewPortPosition = num
End Function


Public Function SysMenuProc(ByVal hWnd As Long, ByVal iMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

' This procedure intercepts Windows messages and looks for any that might encourage us
' to Unacquire the mouse.

  If iMsg = WM_ENTERMENULOOP Then
    objDIDev.Unacquire
    SetSystemCursor
  End If
  
  ' Call the default window procedure
  SysMenuProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)

End Function



    
