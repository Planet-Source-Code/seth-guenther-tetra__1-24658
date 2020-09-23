VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BitMapBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////
'BitmapBuffer class
'last mod date: 6/28/01
'
'This class is used to create device contexts (DC) in memory to allow for
'manipulation of bitmap graphics.  An object declared as a bitmap buffer
'can be used in a manner similar to a pre-defined VB object (such as a form or
'picturebox.)  To use the class, include it in your project and define
'a variable as type BitmapBuffer (e.g. Dim bitmap1 as New BitmapBuffer)
'
'Variables used by this class
'   bufferDC        - internal pointer to the DC of the BitmapBuffer
'   bmpOld          - holds temporary bitmap used in DC creation
'   bufferFile      - holds path and filename of bitmap file
'   bmpWidth        - width of the BitmapBuffer
'   bmpHeight       - height of the BitmapBuffer
'Class properties
'   Let Bitmapfile  - used to assign the path and filename of a bitmap file
'                     to an instance of the BitmapBuffer class
'   Get Bitmapfile  - used to obtain the bitmap file path for a particular
'                     BitmapBuffer instance
'   Get hDC         - returns a handle (pointer) to the DC of a BitmapBuffer
'   Get Height      - returns the height (pixels) of a BitmapBuffer
'   Get Width       - returns the width (pixels) of a BitmapBuffer
'Class Procedures/Functions
'   Create          - used to create the memory DC for a BitmapBuffer.  If no
'                     bitmap file is specified when the BitmapBuffer is created,
'                     the height and width of the DC must be specified.
'            IMPORTANT:  Before a BitmapBuffer object can be created using the
'            .Create method, the form containing the object must be made active,
'            otherwise an error will occur.  The simplest way to make a form active
'            is to call its .Show method.
'
'   Destroy         - used to destroy the DC
'   Class_Terminate - destroys the BitmapBuffer when the program terminates or
'                     there are no more instances of the BitmapBuffer class
'   NewDC           - function to create the actual memory DC
'///////////////////////////////////////////////////////////////////////////////

Option Explicit

'Type declarations for reading information from a bitmap file
Private Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

'Variables
Private bufferDC As Long
Private bmpOld As Long
Private bmpWidth As Long
Private bmpHeight As Long
Private bufferFile As String
                               
Property Let bitmapFile(ByVal bmpFile As String)
'This property is used to assign a bitmap file to a BitmapBuffer object.
'The header information from the bitmap file is read to determine the
'size (width x height) that the DC will need to be.

Dim fileNum As Integer            'File handle for the bitmap
Dim fileHeader As BITMAPFILEHEADER  'Bitmap file header
Dim infoHeader As BITMAPINFOHEADER  'Bitmap information header

'First, check to make sure the file specified by bmpFile exists
If Dir(bmpFile) <> "" Then
    fileNum = FreeFile      'Get the next free file handle
    
    'Open the bitmap file and read the information
    Open bmpFile For Binary Access Read As #fileNum
    Get #fileNum, , fileHeader  'must get past the file header first
    Get #fileNum, , infoHeader  'the info header is what we need
    Close #fileNum
    
    'Make sure the height and width are greater than zero
    If infoHeader.biHeight <= 0 Or infoHeader.biWidth <= 0 Then
        Err.Raise Number:=vbObjectError + 32002, Description:="Invalid bitmap dimensions."
        bufferFile = ""
    Else
    'Store the path and filename and set the height and width of the BitmapBuffer
        bufferFile = bmpFile
        bmpHeight = infoHeader.biHeight
        bmpWidth = infoHeader.biWidth
    End If
Else
'If the file doesn't exist, raise an error
    Err.Raise Number:=vbObjectError + 32001, Description:=bmpFile & " does not exist."
    bufferFile = ""
End If
End Property

Property Get bitmapFile() As String
'This procedure returns the path and name of the bitmap file.
    bitmapFile = bufferFile
End Property

Property Get hdc() As Long
'This property returns the handle (pointer) to the BitmapBuffer DC.
    hdc = bufferDC
End Property

Property Get height()
'This property returns the height of the BitmapBuffer.
    height = bmpHeight
End Property

Property Get width()
'This property returns the width of the BitmapBuffer.
    width = bmpWidth
End Property

Sub Create(Optional bWidth As Variant, Optional bHeight As Variant)
'This procedure creates the memory DC for the BitmapBuffer.  It must
'be called before the buffer can be used.  If a bitmap file is
'specified, then the dimensions of the bitmap are used to create the
'DC.  Otherwise, the dimensions (height & width) must be specified
'before the DC can be created.

'Check to see if a bitmap file is specified
If bufferFile = "" Then
    'If it isn't, check if the parameters bWidth and bHeight exist
    If IsMissing(bWidth) Or IsMissing(bHeight) Then Exit Sub
        'If so, make sure both are greater than 0
        If (bWidth <= 0) And (bHeight <= 0) Then
            'If not, set them both to 0, raise an error and exit the procedure
            bmpWidth = 0
            bmpHeight = 0
            Exit Sub
        Else
            'If they are, then set the width and height of the BitmapBuffer
            bmpWidth = bWidth
            bmpHeight = bHeight
        End If
End If

'Create a memory device context for the BitmapBuffer
bufferDC = NewDC(Screen.ActiveForm.hdc, bmpWidth, bmpHeight)
'If bitmapFile is set, load the bitmap into the buffer
If bufferFile <> "" Then bmpOld = SelectObject(bufferDC, LoadPicture(bufferFile))
End Sub

Sub Destroy()
'This procedure destroys both the memory DC and temporary bitmap to
'free system resources.

'Put the temporary bitmap back into the DC
SelectObject bufferDC, bmpOld

'Delete the memory DC as well as the temorary bitmap
'to free the resources used
DeleteDC bufferDC
DeleteObject bmpOld
End Sub

Private Sub Class_Terminate()
'This procedure is invoked when no more instances of the
'BitmapBuffer class are referenced, such as when you
'end the program. This procedure is the same as the Destroy method.
    
'Put the temporary bitmap back into the DC
If bmpOld > 0 Then SelectObject bufferDC, bmpOld

'Delete the memory DC as well as the temporary bitmap
'to free the resources used
If bufferDC > 0 Then DeleteDC bufferDC
DeleteObject bmpOld
End Sub

Private Function NewDC(formDC As Long, DCWidth As Long, DCHeight As Long) As Long
'This function creates the actual memory DC.  It creates a DC that is compatible
'with formDC, then tests the DC for compatibility by inserting a bitmap into
'it.  If successful, the DC is returned.
Dim compatibleDC, compatibleBMP As Long

compatibleDC = CreateCompatibleDC(formDC)   'Create the DC
compatibleBMP = CreateCompatibleBitmap(formDC, DCWidth, DCHeight)   'Create a temporary bitmap

'Now try to insert the temporary bitmap into the new DC.
If SelectObject(compatibleDC, compatibleBMP) = vbNull Then
    NewDC = vbNull     'If unsuccessful failed, return null
Else
    NewDC = compatibleDC      'otherwise, return the new DC
End If
End Function
