VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1005
   FillColor       =   &H00FF00FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":000C
   MousePointer    =   99  'Custom
   ScaleHeight     =   32
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   67
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TMR_nuvens 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer TMR 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements DirectXEvent

' DirectDraw Object
Dim dx As DirectX7
Dim dd As DirectDraw7
'Dim ds As DirectSound
'Dim dp As DirectPlay4
'Dim di As IDirectInput2A
Public objDXEvent As DirectXEvent
Public objDI As DirectInput
Public objDIDev As DirectInputDevice

' Display capabilities
Dim DXCaps As DDSCAPS2
' Description of DirectDraw Surface
Dim DXDFront As DDSURFACEDESC2
' Front and Back Surface, for double Buffering
Dim DXSFront As DirectDrawSurface7
Dim DXSBack As DirectDrawSurface7
' Bitmap Resources sources, defined as DX-Surfaces
Dim DXSBase01 As DirectDrawSurface7
Dim DXSBackTiled As DirectDrawSurface7
Dim DXSBackTiled2 As DirectDrawSurface7
Dim DXSMapTiles() As DirectDrawSurface7
Dim DXSControlBar As DirectDrawSurface7
Dim DXSFont As DirectDrawSurface7
Dim DXSVLine As DirectDrawSurface7
Dim DXSHLine As DirectDrawSurface7
Dim DXSMouse As DirectDrawSurface7
Dim DXSSmMap As DirectDrawSurface7
Dim DXSSmScreen As DirectDrawSurface7
Dim BufferMouse As DirectDrawSurface7
' Global Program Variables
Dim Message1 As MessageDisplay 'Message to be display with special font
Dim Message2 As MessageDisplay 'not used
Dim MAPWidth
Dim MAPHeight
Dim MouseClick As Boolean
Dim MouseP As POINTAPI
Dim SMScreenX
Dim SMScreenY
Dim Mousex As Integer, Mousey As Integer
Dim ClickOrigineX As Integer, ClickOrigineY As Integer
Dim SelectSquare As Boolean
Dim SmallMapKliked As Boolean
Dim ViewPortX As Integer
Dim ViewPortY As Integer
Dim DemoStarted As Boolean
Dim DemoRunning As Boolean
Dim stoped As Boolean
Dim RS As RECT
Dim BufferMouseRS As RECT
Dim ddck As DDCOLORKEY
Dim DirectionX
Dim DirectionY
Dim ScrollMapHorizontal As Boolean
Dim ScrollMapVertical As Boolean
Dim TableMap() As Integer
Dim MapTableItems() As Integer
Dim DeskTopHWND As Long ' Desktop window handle
    Dim DeskTopHDC As Long ' Desktop device context handle
    Dim mouseHDC As Long
    Dim lTemp As Long ' temp variable
    
    
Dim PicBits(1 To 338000) As Byte, PicInfo As BITMAP, Cnt As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
 
Private Sub Form_Click()
Message1.DisplayedTime = 1
MouseClick = True
End Sub

'
' Initialzie DX, Load graphics
'
Private Sub Form_Load()
     Dim i, i2, TiledWidth As Long, TiledHeight As Long
     Dim file1, strtemp As String, strtemp1, strtemp2
     
     
    ' Retrieve the desktop window handle

    DeskTopHWND = GetDesktopWindow
    
    ' Find out it's HDC

    DeskTopHDC = GetWindowDC(hWnd)

 
     'assign a free file number
     file1 = FreeFile
     
     'open the map file
     Open App.Path & "\map1.txt" For Input As #file1
     
     'read width and height the first 2 lines
     Input #file1, strtemp1, strtemp2
       TiledWidth = Val(strtemp1)
       TiledHeight = Val(strtemp2)
     
     'redimension of the MapTable array according to the file
     ReDim TableMap(TiledWidth - 1, TiledHeight - 1)
     
     'counter for while loop
     i = 0
     
     'load the map array with the file values
     Do While Not EOF(file1)
       Input #file1, strtemp
       For i2 = 0 To Len(strtemp) - 1
         TableMap(i2, i) = Val(Mid(strtemp, i2 + 1, 1))
       Next
       i = i + 1
     Loop
     
     'Set the real width and heigth according to the file and Tile
     MAPWidth = i2 * TILEWIDTH
     MAPHeight = i * TILEHEIGHT
     ReDim MapTableItems(MAPWidth / 10, MAPHeight / 10)
     
     'Set left and top of the small map to be diplayed
     SMScreenX = 10
     SMScreenY = 390
     
     'This should be number of different tile in your map
     'and you should only load tiles utilized in the map
     ReDim DXSMapTiles(2)
     
     'set the message to be displayed when mouse is clicked
    BufferMouseRS.Right = 23
    BufferMouseRS.Bottom = 29
    BufferMouseRS.Left = Mousex
    BufferMouseRS.Top = Mousey
     
     Message1.DisplayTime = 55
     Message1.MessageText = "delaymessage"
     Message1.Position.Bottom = 15
     Message1.Position.Top = 0
     Message1.Position.Left = 0
     
End Sub

Sub DDInitiation()
     ' Initialize DirectX-Object, set display mode
     Set dx = New DirectX7
     Set dd = dx.DirectDrawCreate("")

     
     'dx DirectDrawCreate ByVal 0&, dx, Nothing
     dd.SetCooperativeLevel Me.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
     dd.SetDisplayMode 640, 480, 16, 0, 0
     'DirectSoundCreate ByVal 0&, ds, Nothing
     'DirectPlayCreate ByVal 0&, dp, Nothing
     
     ' Initialize front buffer description
     With DXDFront
        ' .dwSize = Len(DXDFront) ' Get Structure size
         .lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT ' Structure uses Surface Caps and count of BackBuffers
         .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX 'Or DDSCAPS_SYSTEMMEMORY' Structure describes a flippable (buffered) surface
         .lBackBufferCount = 1 ' Structure uses one BackBuffer
     End With
     
     ' Create front buffer from structure
     Set DXSFront = dd.CreateSurface(DXDFront)
     ' Create back buffer from front buffer
     DXCaps.lCaps = DDSCAPS_BACKBUFFER
     'attach both surface together
     Set DXSBack = DXSFront.GetAttachedSurface(DXCaps)
     
'**************************************************
'**************************************************
'**************************************************
'MOUSE
     
   '  procOld = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf SysMenuProc)

  ' Initialize our private cursor
  Mousex = 320
  Mousey = 240
'  g_Sensitivity = 1.2
  
  
  ' Create DirectInput and set up the mouse
 ' Set objDI = dx.DirectInputCreate
 ' Set objDIDev = objDI.CreateDevice("guid_SysMouse")
 ' Call objDIDev.SetCommonDataFormat(DIFORMAT_MOUSE)
 ' Call objDIDev.SetCooperativeLevel(hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)
  
  ' Set the buffer size
 ' Dim diProp As DIPROPLONG
 ' diProp.lHow = DIPH_DEVICE
 ' diProp.lObj = 0
 ' diProp.lData = BufferSize
 ' diProp.lSize = Len(diProp)
 ' Call objDIDev.SetProperty("DIPROP_BUFFERSIZE", diProp)

  ' Ask for notifications
  
 ' EventHandle = dx.CreateEvent(Me)
 ' Call objDIDev.SetEventNotification(EventHandle)
  
  ' Acquire the mouse
 ' AcquireMouse

     
End Sub

Private Sub DirectXEvent_DXCallback(ByVal eventid As Long)

' This is where we respond to any change in mouse state. Usually this will be an axis movement
' or button press or release, but it could also mean we've lost acquisition.
' Note: no event is signalled if we voluntarily Unacquire. Normally loss of acquisition will
' mean that the app window has lost the focus.

  Dim diDeviceData(1 To BufferSize) As DIDEVICEOBJECTDATA
  Dim NumItems As Integer
  Dim i, x As Integer
  Static OldSequence As Long
  
  ' Get data
  On Error GoTo INPUTLOST
  NumItems = objDIDev.GetDeviceData(diDeviceData, 0)
  On Error GoTo 0
  
 
 '***************************************
 'If Button = 1 Then SelectSquare = True
 '***************************************

  
  ' Process data
  For i = 1 To NumItems
    Select Case diDeviceData(i).lOfs
      Case DIMOFS_X
        Mousex = Mousex + diDeviceData(i).lData * g_Sensitivity
           
  '      If OldSequence <> diDeviceData(i).lSequence Then
   '        OldSequence = diDeviceData(i).lSequence
  '      Else
  '        OldSequence = 0
  '      End If
         
      Case DIMOFS_Y
        Mousey = Mousey + diDeviceData(i).lData * g_Sensitivity
 '       If OldSequence <> diDeviceData(i).lSequence Then
 '         OldSequence = diDeviceData(i).lSequence
 '       Else
 '         OldSequence = 0
 '       End If
        
      Case DIMOFS_BUTTON0
        If diDeviceData(i).lData And &H80 Then
   '
       ' Keep record for Line function
          CurrentX = Mousex
          CurrentY = Mousey
          ClickOrigineX = CurrentX
          ClickOrigineY = CurrentY
         'define square for the little map if clicked on
          SelectSquare = True
          If CurrentX > 10 And CurrentX < MAPWidth / 100 + 10 And CurrentY > 390 And CurrentY < MAPHeight / 100 + 390 Then SmallMapKliked = True
 
           
          ' Draw a point in case button-up follows immediately
   '       PSet (Mousex, Mousey)
        Else ' button up
          SelectSquare = False
           MouseClick = True
          'Drawing = False
        End If
           
      Case DIMOFS_BUTTON1
        If diDeviceData(i).lData = 0 Then
    '      Popup
        End If
        
    End Select
  Next i
 ' Label1.Caption = "X : " & Mousex & "     Y : " & Mousey
 'allow green select square
 Select Case Mousex 'axe
    Case Is <= 0: Mousex = 0: DirectionX = -1: ScrollMapHorizontal = True 'left 0
    Case Is >= 639: Mousex = 639: DirectionX = 1: ScrollMapHorizontal = True 'right 640 - 1
    Case Else: ScrollMapHorizontal = False 'no scroll
 End Select
 Select Case Mousey 'axe
    Case Is <= 0: Mousey = 0: DirectionY = -1: ScrollMapVertical = True 'top 0
    Case Is >= 479: Mousey = 479: DirectionY = 1: ScrollMapVertical = True 'bottom 480 -1
    Case Else: ScrollMapVertical = False 'no scroll
 End Select
 
 
 
  Exit Sub
  
INPUTLOST:
  ' Windows stole the mouse from us. DIERR_INPUTLOST is raised if the user switched to
  ' another app, but DIERR_NOTACQUIRED is raised if the Windows key was pressed.
  If (Err.Number = DIERR_INPUTLOST) Or (Err.Number = DIERR_NOTACQUIRED) Then
    SetSystemCursor
    Exit Sub
  End If
    
End Sub
Sub AcquireMouse()

  Dim CursorPoint As POINTAPI
  
  ' Move private cursor to system cursor.
  Call GetCursorPos(CursorPoint)  ' Get position before Windows loses cursor
  Call ScreenToClient(hWnd, CursorPoint)
  
  On Error GoTo CANNOTACQUIRE
  objDIDev.Acquire
  Mousex = CursorPoint.x
  Mousey = CursorPoint.y
  
  
  'frmCanvas.imgPencil.Visible = True
  On Error GoTo 0
  Exit Sub

CANNOTACQUIRE:
  Exit Sub
End Sub


Public Sub SetSystemCursor()

 ' Get the system cursor into the same position as the private cursor,
 ' and stop drawing
 
  Dim point As POINTAPI
  point.x = Mousex
  point.y = Mousey
  ClientToScreen hWnd, point
  SetCursorPos point.x, point.y

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 27 'Quit the GAME should ask before hein ? :)
        DemoRunning = False
End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'save the mouse click origin
CurrentX = x
CurrentY = y
ClickOrigineX = CurrentX
ClickOrigineY = CurrentY
'define square for the little map if clicked on
SelectSquare = True
If CurrentX > 10 And CurrentX < MAPWidth / 100 + 10 And CurrentY > 390 And CurrentY < MAPHeight / 100 + 390 Then
  SmallMapKliked = True
  ScrollMapHorizontal = True
  ScrollMapVertical = True
  SelectSquare = False
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  '****************************************
  ' Dim didevstate As DIMOUSESTATE
 If Not SmallMapKliked Then
 Select Case x 'axe
    Case 0: Mousex = 0: DirectionX = -1: ScrollMapHorizontal = True  'left 0
    Case 639: Mousex = 639: DirectionX = 1: ScrollMapHorizontal = True 'right 640 - 1
    Case Else: ScrollMapHorizontal = False 'no scroll
 End Select
 Select Case y 'axe
    Case 0: Mousey = 0: DirectionY = -1: ScrollMapVertical = True  'top 0
    Case 479: Mousey = 479: DirectionY = 1: ScrollMapVertical = True 'bottom 480 -1
    Case Else: ScrollMapVertical = False 'no scroll
 End Select
  
 End If
 
   'If CurrentX > 10 And CurrentX < MAPWidth / 100 + 10 And CurrentY > 390 And CurrentY < MAPHeight / 100 + 390 Then SmallMapKliked = True
Mousex = x
Mousey = y
  ' We want to force acquisition of the mouse whenever the context menu is closed,
  ' whenever we switch back to the application, or in any other circumstance where
  ' Windows is finished with the cursor. If a MouseMove event happens,
  ' we know the cursor is in our app window and Windows is generating mouse messages, therefore
  ' it's time to reacquire.

  ' Note: this event is triggered whenever the window gets the mouse, even when there's no mouse
  ' activity -- for example, when we have just Alt+Tabbed back, or cancelled out of the context
  ' menu with the Esc key.

   'If Suspended Then Exit Sub    ' Allow continued use of Windows cursor
  
  'MsgBox "dfgsdgf"
' lTemp = BitBlt(DeskTopHDC, 0, 0, 23, 29, mouseHDC, 0, 0, vbSrcCopy)
  
  ' This event gets called again after we acquire the mouse. In order to prevent the cursor
  ' position being set to the middle of the window, we check to see if we've already acquired,
  ' and if so, we don't reposition our private cursor. The only way to find out if the mouse
  ' is acquired is to try to retrieve data.
  
  'On Error GoTo NOTYETACQUIRED
  'Call objDIDev.GetDeviceStateMouse(didevstate)
  'On Error GoTo 0
  'Exit Sub
  
'NOTYETACQUIRED:
  'Call AcquireMouse
 End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'reset clicked surface
SelectSquare = False
SmallMapKliked = False
MouseClick = True
      
End Sub

Private Sub Form_Paint()
'DoEvents
 Dim i
 Dim hBmp As Long
 Dim shdc
 If DemoStarted Then Exit Sub
    DemoStarted = True
    
    'Call the Direct Draw obect initialization
    DDInitiation
    
    'Initilaize all the needed surface for the game
    Set DXSBackTiled = LoadBitmapIntoDXS(dd, App.Path + "\tile1.bmp", 6, 5, 1)
    
    Set DXSMapTiles(0) = LoadBitmapIntoDXS(dd, App.Path + "\tile1.bmp", 1, 1, 1)
    Set DXSMapTiles(1) = LoadBitmapIntoDXS(dd, App.Path + "\tile2.bmp", 1, 1, 1)
    Set DXSBase01 = LoadBitmapIntoDXS(dd, App.Path + "\Base1.bmp", 1, 1, 1)
    Set DXSControlBar = LoadBitmapIntoDXS(dd, App.Path + "\barra.bmp", 1, 1, 1)
    Set DXSFont = LoadBitmapIntoDXS(dd, App.Path + "\myfont.bmp", 1, 1, 1)
    Set DXSVLine = LoadBitmapIntoDXS(dd, App.Path + "\Vline.bmp", 640, 1, 1)
    Set DXSHLine = LoadBitmapIntoDXS(dd, App.Path + "\Vline.bmp", 1, 480, 1)
    Set DXSMouse = LoadBitmapIntoDXS(dd, App.Path + "\mouse.bmp", 1, 1, 1)
    Set BufferMouse = LoadBitmapIntoDXS(dd, App.Path + "\mouse.bmp", 1, 1, 1)
    Set DXSSmMap = LoadBitmapIntoDXS(dd, App.Path + "\SmallMap.bmp", 1, 1, 1)
    Set DXSSmScreen = LoadBitmapIntoDXS(dd, App.Path + "\smscreen.bmp", 1, 1, 1)
    'mouseHDC = DXSMouse.GetDC
    '**************************************************
    'HEY YOU !! ;)
    'DON'T MAKE A SCREEN FOR LOADING IMAGES BLABLABLA...
    'NO ONE WILL SEE IT EXEPCT FOR CPU ;)
    'UNLESS U HAVE 2 3 MG OF IMAGES TO LOAD CAPICH ;)
    'OR YOU HAVE A 486 OR COMODORE64 OR A NINT.. OUPS
    '***************************************************
    Dim rep, rep1, rep2 As Long
    'shdc = DXSBackTiled.GetDC()
    'hBmp = CreateCompatibleBitmap(shdc, 5 * 130, 4 * 130)
    'ReDim PicBits(1 To 640 * 480 * 4)
    'rep2 = SelectObject(shdc, hBmp)
    
    'GetBitmapBits rep2, UBound(PicBits), PicBits(1)
    
    'rep = Str(PicBits(1)) & Str(PicBits(2)) & Str(PicBits(3)) & Str(PicBits(4))
    
    'For i = 1 To UBound(PicBits)
    ' PicBits(i) = IIf((PicBits(i) - 50) < 0, 0, PicBits(i) - 50)
    'Next
    'SetBitmapBits rep2, UBound(PicBits), PicBits(1)
    'rep1 = Str(PicBits(1)) & Str(PicBits(2)) & Str(PicBits(3)) & Str(PicBits(4))
    'SetBitmapBits rep2, UBound(PicBits), PicBits(1)
    'DoEvents
    'Dim rs1 As RECT
    'rs1.Top = 0
    'rs1.Left = 0
    'rs1.Right = 5 * 130
    'rs1.Bottom = 4 * 130
    'DXSBackTiled.BltToDC shdc, rs1, rs1
    'DXSBackTiled.ReleaseDC shdc
    
    
    'DXSBackTiled.DrawText 0, 0, rep & "   DIFF   " & rep1, True
    
    
    
    'initialize global variables
    ShowCursor False
    ScrollMapHorizontal = False
    ScrollMapVertical = False
    ViewPortY = 0
    ViewPortX = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i
    'Flip from DX-Surface to standard GDI
    dd.FlipToGDISurface
    ' Restore old resolution and depth
    dd.RestoreDisplayMode
    ' Return control to windows
    dd.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
    
    ' !IMPORTANT! Clear all DX Objects
    Set DXSBackTiled = Nothing
    Set DXSMapTiles(0) = Nothing
    Set DXSMapTiles(1) = Nothing
    Set DXSBase01 = Nothing
    Set DXSControlBar = Nothing
    Set DXSFont = Nothing
    Set DXSVLine = Nothing
    Set DXSHLine = Nothing
    Set DXSMouse = Nothing
    Set DXSSmMap = Nothing
    Set DXSSmScreen = Nothing
    Set DXSBack = Nothing
    Set DXSFront = Nothing
    
    Set dd = Nothing
    'Set ds = Nothing 'if direct sound
    'Set dp = Nothing 'if direct play
    'If procOld <> 0 Then
    'Call SetWindowLong(hWnd, GWL_WNDPROC, procOld)
    'End If
    'If EventHandle <> 0 Then dx.DestroyEvent EventHandle
    Set dx = Nothing
    ShowCursor 1
    'lTemp = ReleaseDC(DeskTopHWND, DeskTopHDC)

End Sub

'***************************************************
'GAME ACTION HERE !!!!!
'*************************************************
Private Sub TMR_Timer()
 'wait until all the obects are created from the paint event
 If Not DemoStarted Then Exit Sub
 Dim yes As Boolean
 Dim i As Integer, i2 As Integer
 Dim divtempX, divtempY, oldx, oldy
 Dim dxdc As Long, divrestY As Integer, divrestX As Integer
 Dim ScrollMapHorizontal2 As Boolean
 TMR.Enabled = False
 DemoRunning = True
 
 
 While DemoRunning
    '********************************************
    'SCROLL MAP CHECK FOR EVENT
    '********************************************
    If ScrollMapHorizontal And Not SelectSquare Then ViewPortX = GetNewViewPortPosition(DirectionX, ViewPortX, 640, MAPWidth)
    If ScrollMapVertical And Not SelectSquare Then ViewPortY = GetNewViewPortPosition(DirectionY, ViewPortY, 480, MAPHeight)
    '********************************************
  
    '***************************************************
    'SCROLLING FOR THE SMALL MAP
    '***************************************************
    If SmallMapKliked Then
      If Mousex > 10 And Mousex < MAPWidth / 100 + 2 And Mousey > 390 And Mousey < MAPHeight / 100 + 385 Then
        ViewPortX = Int(((Mousex - 10) * 100) / SCROLLSPEED) * SCROLLSPEED
        ViewPortY = Int(((Mousey - 390) * 100) / SCROLLSPEED) * SCROLLSPEED
      Else
        
        SmallMapKliked = False
      End If
    End If
    '***************************************************
   If ScrollMapVertical Or ScrollMapHorizontal Then
    '***************************************************
    'PAINT MAP TILES FROM NEW POSITION
    '***************************************************
    divrestX = ViewPortX Mod 130
    divrestY = ViewPortY Mod 130
    RS.Top = divrestY
    RS.Left = divrestX
    RS.Right = 130
    RS.Bottom = 130
    divtempX = 0
    divtempY = 0
    For i = 0 To 4
       For i2 = 0 To 5
         DXSBackTiled.BltFast (i2 * 130 - divtempX), (i * 130 - divtempY), DXSMapTiles(TableMap(Int(ViewPortX / 130) + i2, Int(ViewPortY / 130) + i)), RS, DDBLTFAST_SRCCOLORKEY
         RS.Left = 0
         divtempX = divrestX
      
       Next
       RS.Left = divrestX
       RS.Top = 0
       divtempY = divrestY
       divtempX = 0
    Next
   End If
    '********************************************
    'FINAL PAINTING BEFORE FLIPING
    '********************************************
    
    RS.Top = 0
    RS.Left = 0
    RS.Right = 640
    RS.Bottom = 480
    DXSBack.BltFast 0, 0, DXSBackTiled, RS, DDBLTFAST_SRCCOLORKEY
    '********************************************
    
    '********************************************
    'PAINTING BASE 01
    '*******************************************************************
     RS.Top = 0
     RS.Left = 0
     RS.Bottom = 110
     RS.Right = 130
     DXSBack.BltFast -ViewPortX + 200, -ViewPortY + 200, DXSBase01, RS, DDBLTFAST_SRCCOLORKEY
    '********************************************
   
    '********************************************
    'IF MOUSE CLICKED PAINT MESSAGE 1
    '********************************************
 
    If MouseClick Then
        Set Message1.MessageSurface = SetDisplayMessage(dd, Message1.MessageText)
        Message1.Position.Right = 15 * Len(Message1.MessageText)
        DXSBack.BltFast 250, 400, Message1.MessageSurface, Message1.Position, DDBLTFAST_SRCCOLORKEY
        Set Message1.MessageSurface = Nothing
        If Message1.DisplayedTime = Message1.DisplayTime Then MouseClick = False
        Message1.DisplayedTime = Message1.DisplayedTime + 1
    End If
    '********************************************
    
    '********************************************
    'BLACK BAR AT THE BOTTOM
    '*******************************************************************
     RS.Bottom = 115
     RS.Right = 640
     DXSBack.BltFast 0, 365, DXSControlBar, RS, DDBLTFAST_SRCCOLORKEY
    '********************************************
  
    '********************************************
    'PAINT SMALL MAP IN THE CORNER
    '********************************************
     RS.Right = 78
     RS.Bottom = 78
     DXSBack.BltFast 10, 390, DXSSmMap, RS, DDBLTFAST_SRCCOLORKEY
    '********************************************
  
    '********************************************
    'PAINT SMALL SCREEN ON THE SMALL MAP
    '********************************************
     RS.Right = 6
     RS.Bottom = 4
     DXSBack.BltFast SMScreenX + Int(ViewPortX / 100), SMScreenY + Int(ViewPortY / 100), DXSSmScreen, RS, DDBLTFAST_SRCCOLORKEY
    '********************************************
     ' DoEvents
    
    DoEvents
    '********************************************
    'IF SELECT SQUARE THEN PAINT THE SQUARE
    '********************************************
    If SelectSquare And Not SmallMapKliked Then
       RS.Bottom = 1
       RS.Right = Abs(Mousex - ClickOrigineX)
       DXSBack.BltFast IIf(Mousex < ClickOrigineX, Mousex, ClickOrigineX), ClickOrigineY, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
       DXSBack.BltFast IIf(Mousex < ClickOrigineX, Mousex, ClickOrigineX), Mousey, DXSVLine, RS, DDBLTFAST_SRCCOLORKEY
       RS.Right = 1
       RS.Bottom = Abs(Mousey - ClickOrigineY) + 1
       DXSBack.BltFast ClickOrigineX, IIf(Mousey < ClickOrigineY, Mousey, ClickOrigineY), DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
       DXSBack.BltFast Mousex, IIf(Mousey < ClickOrigineY, Mousey, ClickOrigineY), DXSHLine, RS, DDBLTFAST_SRCCOLORKEY
     End If
    '********************************************
    
    '********************************************
    'MOUSE PAINTING
    '*******************************************************************
    RS.Left = 0
    RS.Top = 0
    RS.Bottom = 29
    RS.Right = 23
    If Mousey > 450 Then RS.Bottom = 480 - Mousey
    If Mousex > 616 Then RS.Right = 640 - Mousex
    DXSBack.BltFast Mousex, Mousey, DXSMouse, RS, DDBLTFAST_SRCCOLORKEY
    '********************************************
       
    '********************************************
    'Fliping buffers chain
    '********************************************
     On Error Resume Next
     DXSFront.Flip DXSBack, 0
     If Err.Number = DDERR_SURFACELOST Then DXSFront.restore
     'lTemp = BitBlt(DeskTopHDC, Mousex, Mousey, 23, 29, mouseHDC, 0, 0, vbSrcCopy)
     
    '********************************************
        
Wend

Unload Me
End Sub
'********************************************
'Function to create and return a text message
'with the format of a surface with a special font
'**********************************************
Private Function SetDisplayMessage(DXObject As DirectDraw7, ByVal mess As String) As DirectDrawSurface7
Dim DXSMessTemp As DirectDrawSurface7
Dim TempDXD As DDSURFACEDESC2         ' Surface description
Dim ddck As DDCOLORKEY
Dim RStemp As RECT
Dim i
ddck.low = 0 'TRANSPARENT VALUE
ddck.high = 0
'set the surface values
With TempDXD
        '.dwSize = Len(TempDXD)
        .lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_SYSTEMMEMORY
        .lWidth = 15 * Len(mess)
        .lHeight = 15
End With
    ' Create DX surface
    Set DXSMessTemp = DXObject.CreateSurface(TempDXD)
    RStemp.Top = 0
    RStemp.Bottom = 15
    'Convert message to bitmap onto surface
    For i = 1 To Len(mess)
      RStemp.Left = (Asc(Mid(mess, i, 1)) - 97) * 15
      RStemp.Right = RStemp.Left + 15
      DXSMessTemp.BltFast (i - 1) * 15, 0, DXSFont, RStemp, DDBLTFAST_SRCCOLORKEY
    Next
    DXSMessTemp.SetColorKey DDCKEY_SRCBLT, ddck
    Set SetDisplayMessage = DXSMessTemp
End Function

