VERSION 5.00
Begin VB.Form FrmLunar 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TMR_Nuvens 
      Interval        =   10
      Left            =   120
      Top             =   1320
   End
   Begin VB.Timer TMR 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "FrmLunar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' API VARIABLES

' Some API Constant  and Type Declarations for use with API Functions
'Const IMAGE_BITMAP = 0
Const LR_LOADFROMFILE = &H10
Const LR_CREATEDIBSECTION = &H2000
Const SRCCOPY = &HCC0020

Private Type BITMAP
        bmType          As Long
        bmWidth         As Long
        bmHeight        As Long
        bmWidthBytes    As Long
        bmPlanes        As Integer
        bmBitsPixel     As Integer
        bmBits          As Long
End Type

' Various Standard API Functions for manipulation of DCs and Bitmaps
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal source As Long, ByVal length As Long)
Private Declare Function lstrcpy Lib "Kernel32" (ByVal lpszDestinationString1 As Any, ByVal lpszSourceString2 As Any) As Long

' DIRECTX DEFINITIONS

' DirectDraw Object
Dim DX As DirectDraw2

' Display capabilities
Dim DXCaps As DDSCAPS

' Description of DirectDraw Surface
Dim DXDFront As DDSURFACEDESC

' Front and Back Surface, for double Buffering
Dim DXSFront As DirectDrawSurface2
Dim DXSBack As DirectDrawSurface2

' Bitmap Resources sources, defined as DX-Surfaces
Dim DXSCenario1 As DirectDrawSurface2
Dim DXSCenario2 As DirectDrawSurface2
Dim DXSCarro As DirectDrawSurface2
Dim DXSNuvens As DirectDrawSurface2
Dim DXSEnemy01 As DirectDrawSurface2

Dim DXSBase As DirectDrawSurface2
Dim DXSBarra As DirectDrawSurface2


' Global Program Status
Dim DemoStarted As Boolean
Dim DemoRunning As Boolean
Dim DemoCounter As Long

Dim RS As RECT, Carro As RECT
Dim NuvemRect As RECT
Dim Enemy01Rect As RECT
Dim BaseRect As RECT
Dim BarraRect As RECT


Dim ddck As DDCOLORKEY
Dim cont
Dim PosicaoX

'
' Initialzie DX, Load graphics
'
Private Sub Form_Load()
     
     ' Don't execute load-code when called during execution
     '(just to be on the save side)
     
     If DemoStarted Then Exit Sub
     
     ' Set running status
     DemoRunning = True
     DemoStarted = True
     DemoCounter = 0
     
     IniciaDirectDraw
     
     Set DXSCenario1 = LoadBitmapIntoDXS(DX, App.Path + "\cenario1.bmp")
     Set DXSCenario2 = LoadBitmapIntoDXS(DX, App.Path + "\cenario2.bmp")
     Set DXSCarro = LoadBitmapIntoDXS(DX, App.Path + "\lunar.bmp")
     Set DXSNuvens = LoadBitmapIntoDXS(DX, App.Path + "\nuvens.bmp")
     
     Set DXSEnemy01 = LoadBitmapIntoDXS(DX, App.Path + "\enemy01.bmp")
     
     Set DXSBarra = LoadBitmapIntoDXS(DX, App.Path + "\barra.bmp")
     Set DXSBase = LoadBitmapIntoDXS(DX, App.Path + "\base.bmp")
     

     PosicaoX = 100
     ' Enable main loop
     Me.TMR_Nuvens.Enabled = True
     Me.TMR.Enabled = True

     
End Sub

Sub IniciaDirectDraw()
     
     ' Initialize DirectX-Object, set display mode
     DirectDrawCreate ByVal 0&, DX, Nothing
     DX.SetCooperativeLevel Me.hwnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
     
     'DX.SetDisplayMode 640, 480, 16, 0, 0
    
     ' Initialize front buffer description
     With DXDFront
         ' Get Structure size
         .dwSize = Len(DXDFront)
         ' Structure uses Surface Caps and count of BackBuffers
         .dwFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
         ' Structure describes a flippable (buffered) surface
         .DDSCAPS.dwCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX Or DDSCAPS_SYSTEMMEMORY
         ' Structure uses one BackBuffer
         .dwBackBufferCount = 1
     End With
    
     ' Create front buffer from structure
     DX.CreateSurface DXDFront, DXSFront, Nothing
     
     ' Create back buffer from front buffer
     DXCaps.dwCaps = DDSCAPS_BACKBUFFER
     DXSFront.GetAttachedSurface DXCaps, DXSBack
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    
    Case 37, 38
        If NumX > 15 Then
           NumX = NumX - 3 '5 '15
           If PosicaoX > 100 Then
              PosicaoX = PosicaoX - 3  '5 '15
           End If
        End If

        TMR.Enabled = True

    
    Case 39, 40
        If NumX < 630 Then
           NumX = NumX + 3 '5 '15
        Else
           PosicaoX = PosicaoX + 3 '5 '15
        End If
        
        TMR.Enabled = True
       
    Case 16, 17, 18
    Case 27
        TMR.Enabled = False
        DemoRunning = False
        Unload Me
End Select


End Sub

'
' Rest in peace, DirectX!
'
Private Sub Form_Unload(Cancel As Integer)

    'Flip from DX-Surface to standard GDI
    DX.FlipToGDISurface
    ' Restore old resolution and depth
    
    DX.RestoreDisplayMode
    ' Return control to windows
    DX.SetCooperativeLevel Me.hwnd, DDSCL_NORMAL
    
    ' !IMPORTANT! Clear all DX Objects

    Set DXSBack = Nothing
    Set DXSFront = Nothing
    Set DXSCenario1 = Nothing
    Set DXSCenario2 = Nothing
    Set DXSCarro = Nothing
    Set DXSNuvens = Nothing
    Set DXSEnemy01 = Nothing
    
    Set DXSBase = Nothing
    Set DXSBarra = Nothing
    
    Set DX = Nothing
    
End Sub

'
' DirectX Bitmap Loader
'
Private Function LoadBitmapIntoDXS(DXObject As DirectDraw2, ByVal BMPFile As String) As DirectDrawSurface2
    
    Dim hBitmap As Long                 ' Handle on bitmap
    Dim dBitmap As BITMAP               ' Handle on bitmap descriptor
    Dim TempDXD As DDSURFACEDESC        ' Surface description
    Dim TempDXS As DirectDrawSurface2   ' Created surface
    Dim dcBitmap As Long                ' Handle on image
    Dim dcDXS As Long                   ' Handle on surface context
    Dim ddck As DDCOLORKEY
    
    ddck.dwColorSpaceLowValue = 0
    ddck.dwColorSpaceHighValue = 0
    
    ' Load bitmap
    hBitmap = LoadImage(ByVal 0&, BMPFile, 0, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
    
    ' Get bitmap descriptor
    GetObject hBitmap, Len(dBitmap), dBitmap
    
    ' Fill DX surface description
    With TempDXD
        .dwSize = Len(TempDXD)
        .dwFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .DDSCAPS.dwCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .dwWidth = dBitmap.bmWidth
        .dwHeight = dBitmap.bmHeight
    End With
    ' Create DX surface
    DXObject.CreateSurface TempDXD, TempDXS, Nothing
    
    ' Create API memory DC
    dcBitmap = CreateCompatibleDC(ByVal 0&)
    ' Select the bitmap into API memory DC
    SelectObject dcBitmap, hBitmap
    
    ' Restore DX surface
    TempDXS.Restore
    ' Get DX surface API DC
    TempDXS.GetDC dcDXS
    ' Blit BMP from API DC into DX DC using standard API BitBlt
    StretchBlt dcDXS, 0, 0, TempDXD.dwWidth, TempDXD.dwHeight, dcBitmap, 0, 0, dBitmap.bmWidth, dBitmap.bmHeight, SRCCOPY
    
    ' Cleanup
    TempDXS.ReleaseDC dcDXS
    DeleteDC dcBitmap
    DeleteObject hBitmap
    
    TempDXS.SetColorKey DDCKEY_SRCBLT, ddck
    ' Return created DX surface
    Set LoadBitmapIntoDXS = TempDXS
    
End Function


Private Sub TMR_Lagarto_Timer()
    
    Dim LagartoRect As RECT
    Dim tmpRECT As RECT
    Static cont2
    Static cont3
    
''     If cont2 < 3 Then
'         cont2 = cont2 + 1
   '' Else
  ''      cont2 = 0
    ''End If

    'LagartoRect.Top = 0
   ' LagartoRect.Left = cont2 * 160.5
   ' LagartoRect.bottom = 100
   ' LagartoRect.Right = (cont2 + 1) * 160.5'

    'tmpRECT.Top = 0
    'tmpRECT.Left = 0
    'tmpRECT.Right = 160.5
    'tmpRECT.bottom = 100
       
    
'    DXSCenario1.BltFast 815, 319, DXSEnemy01, LagartoRect, DDBLTFAST_SRCCOLORKEY

'    RS.Top = 0
'    RS.Left = NumX
'    RS.Right = 640 + NumX
'    RS.bottom = 480
    
    'DXSBack.BltFast 0, 0, DXSCenario1, RS, DDBLTFAST_SRCCOLORKEY

'    If cont3 < 5 Then
'        cont3 = cont3 + 1
'    Else
'         cont3 = 0
''     End If

'    LiderRect.Top = 0
'    LiderRect.Left = cont3 * 209.6666666667
'    LiderRect.bottom = 200
'    LiderRect.Right = (cont3 + 1) * 209.6666666667


End Sub



Private Sub TMR_Nuvens_Timer()
    
    Dim NuvemRect As RECT
    Static cont
    
    NuvemRect.Top = 0
    NuvemRect.Left = 0
    NuvemRect.Right = 640
    NuvemRect.bottom = 480

    DXSBack.BltFast 0, 0, DXSNuvens, NuvemRect, 0

End Sub

'
' Contains main loop
'
Private Sub TMR_Timer()
    
    TMR.Enabled = False
    
    BaseRect.Top = 0
    BaseRect.Left = 0
    BaseRect.bottom = 150
    BaseRect.Right = 170

    DXSCenario2.BltFast 1100, 20, DXSBase, BaseRect, DDBLTFAST_SRCCOLORKEY
    
    RS.Top = 0
    RS.Left = NumX
    RS.Right = 640 + NumX
    RS.bottom = 230
    DXSBack.BltFast 0, 180, DXSCenario2, RS, DDBLTFAST_SRCCOLORKEY
    
    
    RS.bottom = 95
    
    DXSBack.BltFast 0, 385, DXSCenario1, RS, DDBLTFAST_SRCCOLORKEY
    
    Carro.Top = 0
    Carro.Left = 0
    Carro.bottom = 50
    Carro.Right = 110

    DXSBack.BltFast PosicaoX, 345, DXSCarro, Carro, DDBLTFAST_SRCCOLORKEY

    Enemy01Rect.Top = 0
    Enemy01Rect.Left = 0
    Enemy01Rect.bottom = 40
    Enemy01Rect.Right = 110

    DXSBack.BltFast 400, 50, DXSEnemy01, Enemy01Rect, DDBLTFAST_SRCCOLORKEY

    BarraRect.Top = 0
    BarraRect.Left = 0
    BarraRect.bottom = 50
    BarraRect.Right = 640

    DXSBack.BltFast 0, 0, DXSBarra, BarraRect, DDBLTFAST_SRCCOLORKEY

    ' Flip buffers
    On Error Resume Next
    Do
        DXSFront.Flip Nothing, 0
        'DXSFront.Flip DXSBack, 0
        If Err.Number = DDERR_SURFACELOST Then DXSFront.Restore
    Loop Until Err.Number = 0
    On Error GoTo 0
 
    ' Continue with demo
    DemoCounter = DemoCounter + 1
    DoEvents
    
End Sub


