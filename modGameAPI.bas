Attribute VB_Name = "modGameAPI"
'This module contains API calls helpful for making games, such as graphics,
'sound, and keyboard manipulation.

'***************************************************************************
'BitBlt (Bit Block Transfer) copies a graphics area to another graphics area
'***************************************************************************
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
    ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long

'***************************************************************************
'StretchBlt (Stretch Block Transfer) stretches or shrinks one graphics area
'onto another graphics area
'***************************************************************************
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
       ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
       ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
       ByVal ySrc As Long, ByVal nSrcWidth As Long, _
       ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
       
'***************************************************************************
'GetPixel gets the color value of a pixel from a graphics area
'***************************************************************************
Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long

'***************************************************************************
'SetPixel changes the color of a pixel on a graphics area
'***************************************************************************
Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

'***************************************************************************
'DC (Device Context) API functions
'Basically, DC's are a link between an application, such as this
'one, and the graphics output devices (graphics card, monitor, etc) in memory.
'DC's are referenced by a handle (pointer in Microsoft terminology).  If an
'object is able to participate in a device context, it's hasDC property
'will be true, and it's hDC property is the handle to it's DC.  DC's are
'especially useful for graphics, since they can store bitmap pictures in memory
'and be used to display those bitmaps on graphics output devices.  DC's make
'up the backbone for the graphics subsystem of this game.
'***************************************************************************
'CreateCompatibleDC creates the memory DC to hold graphics
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Any) As Long

'CreateCompatibleBitmap makes a temporary bitmap, to test the DC
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

'DeleteDC frees the memory used by a memory DC
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'SelectObject inserts a bitmap into the specified DC
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'DeleteObject deletes a bitmap from memory
Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

'***************************************************************************
'sndPlaySound plays a sound from the file specified by lpszSoundName
'***************************************************************************
Public Const SND_ASYNC = &H1    'flags
Public Const SND_NODEFAULT = &H2

Public Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'***************************************************************************
'GetKeyState determines if a key with code nVirtKey is pressed or not
'***************************************************************************
Public Const KEY_TOGGLED As Integer = &H1
Public Const KEY_PRESSED As Integer = &H1000

Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
