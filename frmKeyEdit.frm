VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmKeyEdit 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movement Sensitivity"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   Icon            =   "frmKeyEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   603
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame lblTaken 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1080
      Left            =   5175
      TabIndex        =   28
      Top             =   4785
      Visible         =   0   'False
      Width           =   3855
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Key Is Taken"
         BeginProperty Font 
            Name            =   "ZeroHour"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   585
         TabIndex        =   29
         Top             =   390
         Width           =   2595
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   5
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   930
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   30
         Width           =   3630
      End
   End
   Begin VB.Frame lblMsg 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   5175
      TabIndex        =   26
      Top             =   3555
      Visible         =   0   'False
      Width           =   3795
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Press new key"
         BeginProperty Font 
            Name            =   "ZeroHour"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   300
         TabIndex        =   27
         Top             =   375
         Width           =   3180
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   5
         FillColor       =   &H00A56E38&
         FillStyle       =   0  'Solid
         Height          =   930
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   30
         Width           =   3630
      End
   End
   Begin VB.Frame frmVDrop 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Drop Speed"
      Height          =   1470
      Left            =   5130
      TabIndex        =   20
      Top             =   1740
      Width           =   3840
      Begin ComctlLib.Slider sldVDrop 
         Height          =   630
         Left            =   570
         TabIndex        =   21
         ToolTipText     =   "Use this to change how fast the blocks fall when you press the down key."
         Top             =   615
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   1111
         _Version        =   327682
         BorderStyle     =   1
         Min             =   1
         Max             =   99
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Fast"
         Height          =   330
         Left            =   210
         TabIndex        =   25
         Top             =   705
         Width           =   330
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Slow"
         Height          =   255
         Left            =   3345
         TabIndex        =   24
         Top             =   705
         Width           =   390
      End
   End
   Begin VB.Frame frmHMove 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Movement Sensitivity"
      Height          =   1470
      Left            =   5130
      TabIndex        =   18
      Top             =   135
      Width           =   3840
      Begin ComctlLib.Slider sldHMove 
         Height          =   630
         Left            =   585
         TabIndex        =   19
         ToolTipText     =   "Use this to change how fast blocks move from side to side."
         Top             =   570
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   1111
         _Version        =   327682
         BorderStyle     =   1
         Min             =   50
         Max             =   150
         SelStart        =   100
         TickFrequency   =   10
         Value           =   100
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "High"
         Height          =   375
         Left            =   210
         TabIndex        =   23
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Low"
         Height          =   255
         Left            =   3345
         TabIndex        =   22
         Top             =   630
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdDefault 
      Caption         =   "default"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3570
      TabIndex        =   2
      ToolTipText     =   "Restores default values."
      Top             =   5625
      Width           =   1440
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1815
      TabIndex        =   1
      ToolTipText     =   "Cancels all changes."
      Top             =   5625
      Width           =   1470
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      TabIndex        =   0
      ToolTipText     =   "Click to apply all changes."
      Top             =   5625
      Width           =   1470
   End
   Begin VB.Frame lblSettings 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Key Settings"
      Height          =   3795
      Left            =   75
      TabIndex        =   3
      Top             =   1740
      Width           =   4950
      Begin VB.PictureBox keys 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   120
         ScaleHeight     =   360
         ScaleWidth      =   1125
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Click to change this key."
         Top             =   3210
         Width           =   1185
      End
      Begin VB.PictureBox keys 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   120
         ScaleHeight     =   360
         ScaleWidth      =   1125
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Click to change this key."
         Top             =   2490
         Width           =   1185
      End
      Begin VB.PictureBox keys 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   120
         ScaleHeight     =   360
         ScaleWidth      =   1125
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Click to change this key."
         Top             =   1770
         Width           =   1185
      End
      Begin VB.PictureBox keys 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   120
         ScaleHeight     =   360
         ScaleWidth      =   1125
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Click to change this key."
         Top             =   1050
         Width           =   1185
      End
      Begin VB.PictureBox keys 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000009&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   120
         ScaleHeight     =   360
         ScaleWidth      =   1125
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Click to change this key."
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "rotate"
         BeginProperty Font 
            Name            =   "ZeroHour"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1425
         TabIndex        =   14
         Top             =   3240
         Width           =   1185
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "counter-clockwIse"
         BeginProperty Font 
            Name            =   "ZeroHour"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1425
         TabIndex        =   13
         Top             =   3465
         Width           =   3030
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "rotate clockwIse"
         BeginProperty Font 
            Name            =   "ZeroHour"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1410
         TabIndex        =   12
         Top             =   2595
         Width           =   2880
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Drop"
         BeginProperty Font 
            Name            =   "ZeroHour"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1410
         TabIndex        =   11
         Top             =   1860
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Move rIGht"
         BeginProperty Font 
            Name            =   "ZeroHour"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1410
         TabIndex        =   10
         Top             =   1170
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Move Left"
         BeginProperty Font 
            Name            =   "ZeroHour"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1410
         TabIndex        =   9
         Top             =   435
         Width           =   1965
      End
   End
   Begin VB.Frame lblInstruct 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Instructions"
      Height          =   1470
      Left            =   90
      TabIndex        =   15
      Top             =   135
      Width           =   4935
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "To restore original key settings, click Default and then OK."
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   4530
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "To change key settings, click on key you want to change and press new key.  Click OK when finished, or Cancel to quit."
         Height          =   465
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   4740
      End
   End
End
Attribute VB_Name = "frmKeyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////
'frmKeyEdit
'created: 5/25/01
'last mod date: 6/28/01
'
'This form allows the player to manipulate the keyboard controls to suit
'his or her liking.  Only certain keys are allowed (see determineKey routine)
'and no one key can be used for more than one function.
'
'The following controls are used in this form:
'   cmdCancel   - command button; closes form and returns to game without changin
'                 key settings
'   cmdDefault  - command button; restores key settings to default values
'   cmdOK       - command button; changes key settings to player's specifications
'   keys        - array of picture boxes; displays player's selections
'   sldHMove    - slider; changes how fast blocks move horizontally
'   sldVDrop    - slider; changes how fast blocks drop due to player
'Other labels are also used.
'//////////////////////////////////////////////////////////////////////////////

'---------
'Variables
'---------
Dim curKey As Integer   'value from 1 - 5 representing which key the player is changing
Dim keyAssign(5) As Integer   'array of key assignments

Private Sub Form_Load()
Dim i As Integer    'loop counter

Me.Icon = LoadResPicture("TICON", vbResIcon)    'load game icon

'Set up initial key assignments (will be default values when the program
'is first run)
keyAssign(0) = frmTetra.keyLeft
keyAssign(1) = frmTetra.keyRight
keyAssign(2) = frmTetra.keyDown
keyAssign(3) = frmTetra.keyCRotate
keyAssign(4) = frmTetra.keyCCRotate

'Display key assignments
For i = 0 To 4
    keys(i).Print determineKey(keyAssign(i))
Next i

sldHMove.Value = frmTetra.moveTimer.Interval
sldVDrop.Value = frmTetra.fallTimer.Interval
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'The KeyPreview property of frmKeyEdit is True, therefore when the player
'strikes a key on the keyboard, this procedure will be called.  If the player
'is wanting to change a key, the key struck will be determined and checked
'for validity, and assigned if valid.
Dim newKey As String    'holds value of struck key

'curKey will be greater than zero when the player has first clicked on one
'of the keys picture boxes to change a key.  Otherwise no change will occur.
'Additionally, if the player changes his or her mind and presses escape,
'no change will occur.
If curKey >= 0 And KeyCode <> vbKeyEscape Then
    newKey = determineKey(KeyCode)  'determine the key struck
    If newKey <> "" And notTaken(KeyCode, curKey) Then  'check if key is already assigned
        keyAssign(curKey) = KeyCode     'if not, assign the key
        keys(curKey).Cls                'display change on screen
        keys(curKey).Print newKey
    End If
End If

curKey = -1     'no changes until player clicks another key to change
lblMsg.Visible = False
End Sub

Private Sub keys_Click(Index As Integer)
'When the player clicks on a key with the mouse to change it
'display message for player to change key
lblTaken.Visible = False
lblMsg.Visible = True
curKey = Index  'note which key the user is changing
End Sub

Private Sub cmdCancel_Click()
'When the user clicks the cancel button, unload the form and return to the
'game without making any changes
Unload Me
End Sub

Private Sub cmdDefault_Click()
'When the player clicks the default button, restore the key settings
'to the default values.
Dim i As Integer    'loop counter

keyAssign(0) = vbKeyLeft    'Left arrow key
keyAssign(1) = vbKeyRight   'Right arrow key
keyAssign(2) = vbKeyDown    'Down arrow key
keyAssign(3) = vbKeySpace   'Space bar
keyAssign(4) = vbKeyControl 'Control key

'Display on the screen
For i = 0 To 4
    keys(i).Cls
    keys(i).Print determineKey(keyAssign(i))
Next i

sldHMove.Value = 100
sldVDrop.Value = 50
End Sub

Private Sub cmdOK_Click()
'After the player has made key selections and clicks OK, assign the new
'values to the global key* variables.
frmTetra.keyLeft = keyAssign(0)
frmTetra.keyRight = keyAssign(1)
frmTetra.keyDown = keyAssign(2)
frmTetra.keyCRotate = keyAssign(3)
frmTetra.keyCCRotate = keyAssign(4)
frmTetra.moveTimer.Interval = sldHMove.Value
frmTetra.fallTimer.Interval = sldVDrop.Value
Unload Me   'return to game
End Sub

Private Function notTaken(ByVal key As Long, ByVal Index As Integer) As Boolean
'This fucntion determines if key is already used in the keyAssign array.
Dim i As Integer    'loop counter

notTaken = True     'initial assumption

For i = 0 To 4  'search for a match between key and every element in keyAssign
                'except for the one at index, which is the key the user wants to change
    If keyAssign(i) = key And i <> Index Then
        notTaken = False    'if match is found, key is taken,
        lblTaken.Visible = True   'display message
        Exit For
    End If
Next i

End Function

Private Function determineKey(ByVal KeyCode As Integer) As String
'This function returns a string based on the value of KeyCode, for
'display within the keys picture boxes.  This allows the player to
'see changes as they are made.
Select Case KeyCode
    Case vbKeyBack
        determineKey = "Backspace"
    Case vbKeyReturn
        determineKey = "Enter"
    Case vbKeyShift
        determineKey = "Shift"
    Case vbKeyControl
        determineKey = "Control"
    Case vbKeySpace
        determineKey = "Space"
    Case vbKeyPageUp
        determineKey = "Page Up"
    Case vbKeyPageDown
        determineKey = "Page Down"
    Case vbKeyEnd
        determineKey = "End"
    Case vbKeyHome
        determineKey = "Home"
    Case vbKeyLeft
        determineKey = "Left"
    Case vbKeyUp
        determineKey = "Up"
    Case vbKeyRight
        determineKey = "Right"
    Case vbKeyDown
        determineKey = "Down"
    Case vbKeyInsert
        determineKey = "Insert"
    Case vbKeyDelete
        determineKey = "Delete"
    Case vbKeyMultiply
        determineKey = "*"
    Case vbKeyAdd
        determineKey = "+"
    Case vbKeySubtract
        determineKey = "--"
    Case vbKeyDecimal
        determineKey = "."
    Case vbKeyDivide
        determineKey = "/"
    Case Else
        If KeyCode >= 65 And KeyCode <= 90 Then   'keys A-Z
            determineKey = Chr$(KeyCode)
        ElseIf KeyCode >= 48 And KeyCode <= 57 Then  'keys 0-9 below function keys
            determineKey = "" & (KeyCode - 48)
        ElseIf KeyCode >= 96 And KeyCode <= 105 Then    'keys 0-9 on 10-key pad
            determineKey = "Numpad " & (KeyCode - 96)
        Else
            determineKey = ""   'return empty string if something else
        End If
End Select
End Function
