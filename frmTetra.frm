VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTetra 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tetra"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7545
   Icon            =   "frmTetra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   504
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   503
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer blockTimer 
      Enabled         =   0   'False
      Left            =   6675
      Top             =   6390
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5100
      Picture         =   "frmTetra.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7005
      Width           =   1125
   End
   Begin VB.Timer moveTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4245
      Top             =   6420
   End
   Begin VB.Timer fallTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   7080
      Top             =   6390
   End
   Begin VB.PictureBox picNext 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1560
      Left            =   5070
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   4
      Top             =   4380
      Width           =   1560
   End
   Begin VB.Timer eraseTimer 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   3840
      Top             =   6420
   End
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6300
      MaskColor       =   &H00808080&
      Picture         =   "frmTetra.frx":0F36
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7005
      Width           =   1155
   End
   Begin VB.PictureBox gameScreen 
      BackColor       =   &H00000000&
      ForeColor       =   &H8000000E&
      Height          =   7560
      Left            =   0
      ScaleHeight     =   500
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   0
      Top             =   0
      Width           =   3810
      Begin VB.Label lblPause 
         BackStyle       =   0  'Transparent
         Caption         =   "Paused"
         BeginProperty Font 
            Name            =   "ZeroHour"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   735
         TabIndex        =   1
         Top             =   3360
         Visible         =   0   'False
         Width           =   2115
      End
   End
   Begin MCI.MMControl bgMusic 
      Height          =   330
      Left            =   5460
      TabIndex        =   3
      Top             =   6510
      Visible         =   0   'False
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      _Version        =   393216
      PrevVisible     =   0   'False
      NextVisible     =   0   'False
      PauseVisible    =   0   'False
      BackVisible     =   0   'False
      StepVisible     =   0   'False
      StopVisible     =   0   'False
      RecordVisible   =   0   'False
      EjectVisible    =   0   'False
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label numDouble 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   6285
      TabIndex        =   20
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label numTriple 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   6285
      TabIndex        =   19
      Top             =   2580
      Width           =   495
   End
   Begin VB.Label numTetra 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   6285
      TabIndex        =   18
      Top             =   2985
      Width           =   495
   End
   Begin VB.Label numSingle 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   6285
      TabIndex        =   17
      Top             =   1755
      Width           =   495
   End
   Begin VB.Label numLines 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   480
      Left            =   4155
      TabIndex        =   9
      Top             =   2775
      Width           =   1590
   End
   Begin VB.Label level 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   4155
      TabIndex        =   8
      Top             =   1965
      Width           =   1440
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   390
      Left            =   4155
      TabIndex        =   6
      Top             =   1635
      Width           =   1755
   End
   Begin VB.Label lblLines 
      BackStyle       =   0  'Transparent
      Caption         =   "Lines"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   330
      Left            =   4155
      TabIndex        =   7
      Top             =   2460
      Width           =   1755
   End
   Begin VB.Label lblDouble 
      BackStyle       =   0  'Transparent
      Caption         =   "Double"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   6000
      TabIndex        =   16
      Top             =   1995
      Width           =   990
   End
   Begin VB.Label lblTriple 
      BackStyle       =   0  'Transparent
      Caption         =   "TrIple"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   6000
      TabIndex        =   15
      Top             =   2430
      Width           =   990
   End
   Begin VB.Label lblTetra 
      BackStyle       =   0  'Transparent
      Caption         =   "Tetra"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   6000
      TabIndex        =   14
      Top             =   2835
      Width           =   990
   End
   Begin VB.Label lblSingle 
      BackStyle       =   0  'Transparent
      Caption         =   "SIngle"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Left            =   6000
      TabIndex        =   13
      Top             =   1590
      Width           =   990
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H008000FF&
      BorderWidth     =   5
      FillColor       =   &H00808000&
      FillStyle       =   4  'Upward Diagonal
      Height          =   1785
      Left            =   5910
      Shape           =   4  'Rounded Rectangle
      Top             =   1515
      Width           =   1320
   End
   Begin VB.Shape shapeNext 
      BorderColor     =   &H00A56E38&
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   3330
      Left            =   4410
      Shape           =   3  'Circle
      Top             =   3480
      Width           =   2565
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  'Opaque
      Height          =   75
      Left            =   3780
      Top             =   6840
      Width           =   3900
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   450
      Left            =   4710
      TabIndex        =   12
      Top             =   345
      Width           =   1920
   End
   Begin VB.Label score 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   480
      Left            =   4470
      TabIndex        =   11
      Top             =   735
      Width           =   2385
   End
   Begin MSForms.ToggleButton cmdPause 
      Height          =   450
      Left            =   3900
      TabIndex        =   10
      Top             =   7005
      Width           =   1125
      BackColor       =   12632256
      ForeColor       =   -2147483630
      DisplayStyle    =   6
      Size            =   "1984;794"
      Value           =   "0"
      Picture         =   "frmTetra.frx":1CC0
      FontName        =   "ZeroHour"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H008000FF&
      BorderWidth     =   5
      FillColor       =   &H00808000&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1785
      Left            =   4035
      Shape           =   4  'Rounded Rectangle
      Top             =   1515
      Width           =   1695
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H008000FF&
      BorderWidth     =   6
      FillColor       =   &H00808000&
      FillStyle       =   6  'Cross
      Height          =   1065
      Left            =   4035
      Shape           =   4  'Rounded Rectangle
      Top             =   225
      Width           =   3240
   End
   Begin VB.Menu menuGame 
      Caption         =   "&Game"
      NegotiatePosition=   1  'Left
      Begin VB.Menu menuNewGm 
         Caption         =   "New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu menuPause 
         Caption         =   "Pause"
         Shortcut        =   {F3}
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu menuHScore 
         Caption         =   "View High Scores"
      End
      Begin VB.Menu dash0 
         Caption         =   "-"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuOptions 
      Caption         =   "&Options"
      Begin VB.Menu menuMusicList 
         Caption         =   "Music"
         Index           =   0
         Begin VB.Menu menuMusic 
            Caption         =   "Classic Tetris A"
            Index           =   0
         End
         Begin VB.Menu menuMusic 
            Caption         =   "Classic Tetris B"
            Index           =   1
         End
         Begin VB.Menu menuMusic 
            Caption         =   "Urban"
            Index           =   2
         End
         Begin VB.Menu menuMusic 
            Caption         =   "Jazzy"
            Checked         =   -1  'True
            Index           =   3
         End
         Begin VB.Menu menuMusic 
            Caption         =   "Techno"
            Index           =   4
         End
         Begin VB.Menu menuMusic 
            Caption         =   "Rock"
            Index           =   5
         End
         Begin VB.Menu menuNoMus 
            Caption         =   "Off"
         End
      End
      Begin VB.Menu menuSndFX 
         Caption         =   "Sound Effects"
         Begin VB.Menu menuFX 
            Caption         =   "On"
            Checked         =   -1  'True
            Index           =   0
            Shortcut        =   {F5}
         End
         Begin VB.Menu menuFX 
            Caption         =   "Off"
            Index           =   1
            Shortcut        =   {F6}
         End
      End
      Begin VB.Menu menuBlkTp 
         Caption         =   "Block Type"
         Begin VB.Menu menuBlk 
            Caption         =   "Classic"
            Index           =   0
         End
         Begin VB.Menu menuBlk 
            Caption         =   "Vector "
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu menuBlk 
            Caption         =   "Gradient"
            Index           =   2
         End
         Begin VB.Menu menuBlk 
            Caption         =   "Diamond"
            Index           =   3
         End
         Begin VB.Menu menuBlk 
            Caption         =   "Sphere"
            Index           =   4
         End
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu opShowNext 
         Caption         =   "Show Next"
         Checked         =   -1  'True
      End
      Begin VB.Menu dash3 
         Caption         =   "-"
      End
      Begin VB.Menu menuEditCtrl 
         Caption         =   "Edit Controls"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu menuHow2Play 
         Caption         =   "How to Play"
         Shortcut        =   {F1}
      End
      Begin VB.Menu dash4 
         Caption         =   "-"
      End
      Begin VB.Menu menuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmTetra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////
'Tetra
'version 1.0
'Author:  Seth Guenther
'Created:  5/25/01
'Last mod date: 7/2/01
'
'Tetra is a Tetris clone for Windows.  I wrote this program because I am a
'huge Tetris fan who was seriously disappointed at the existing Tetris
'programs.  None of them seemed to adhere to the original Tetris parameters
'from the NES and Gameboy versions of Tetris, which I still consider to be
'the best.  As such, much time was devoted to the research :) of the old
'console versions of Tetris, and I did my best to mimic their behavior.
'Of course, I had to add a little flair of my own, but I think you will
'find the gameplay similar.  Enjoy!
'hyperlink
'There are 4 forms to this project:
'   frmTetra    - this is the main form where all game play occurs and game
'                 engine code is located
'   frmKeyEdit  - this form allows the user to change the keyboard controls
'                 to his or her liking
'   frmHScore   - the high scores
'   frmAbout    - small splash screen & copyright info
'
'The following controls are used in this form (in alphabetical order):
'   bgMusic     - multimedia control, handles the background MIDI music
'   blockTimer  - timer control; controls the automatic downward movement of blocks
'   cmdPause    - toggle button; pauses the game
'   cmdQuit     - command button; quits the game
'   cmdReset    - command button; resets the game
'   eraseTimer  - timer control; erases horizontal rows when they become full.  The
'                 reason a timer is used for this event instead of a loop is so that
'                 erasure of lines will occur at the same speed regardless of what
'                 machine it runs on, since it depends on time.  The same philosophy
'                 holds for moveTimer
'   fallTimer   - timer control; moves blocks down faster than normal when the down
'                 button is pressed
'   gameScreen  - picture box; this is where the game's graphics are displayed
'   level       - label; the level the player is on
'   moveTimer   - timer control; controls horizontal movement of blocks
'   numDouble   - the number of times the player has cleared 2 lines simultaneously
'   numLines    - label; number of lines player has cleared
'   numSingle   - the number of singles the player has cleared
'   numTetra    - the number of Tetra's (4 lines) the player has cleared
'   numTriple   - the number of triples the player has cleared
'   picNext     - picture box; shows the next block
'   score       - label; player's score
'
'This project makes use of a resource file, tetra.res.  It contains all the
'bitmaps for the various blocks that can be used in the game, as well as the
'game icon and a text file to construct the tile array (see variable declarations.)
'This keeps all the game resources in one place and makes it easier to manage them.
'
'Various other label and shape controls are used, but they are merely cosmetic and not
'critical to the game's function.  In addition, a fully functional menu system is
'employed.  Explanations of each menu item's function can be found under it's _Click()
'event procedure.
'
'Constant, variable, and type definition explanations are found below in the
'declarations section.
'//////////////////////////////////////////////////////////////////////////

Option Explicit

'----------------------------------------------------------------------------------
'Constants
Const ARRAYROWS As Integer = 22     'number of rows in tile array (see tileType)
Const ARRAYCOLS As Integer = 12     'number of columns in tile array
Const TILEWIDTH As Integer = 25     'the width of one tile (pixels)
Const TILEHEIGHT As Integer = 25    'height of one tile
Const SINGLEVALUE As Integer = 40   'base score for clearing a single horizontal row
Const DOUBLEVALUE As Integer = 10   'base score for a double
Const TRIPLEVALUE As Integer = 300  'base score for a triple
Const TETRAVALUE As Integer = 1200  'base score for a tetris

'----------------------------------------------------------------------------------
'Type declarations
'
'The underlying functionality of this game is based on an array of tiles.  Each
'element of the array represents one 25x25 (pixels) tile on the game screen.  The
'tile array is used for collision detection, erasing complete horizontal rows,
'ending the game when the Tetra pit becomes full, etc.
Private Type tileType
    xPos As Integer     'horizontal position (pixels) of array element
    yPos As Integer     'vertical position (pixels) of array element
    occupied As Integer 'binary value indicating occupation:
                        '   0 - unoccupied
                        '   1 - occupied or border
End Type

'Each Tetra block has both an x and y position, as well as a memory DC which
'holds its bitmap picture.
Private Type blockType
    xPos As Integer     'horizontal position (pixels) of block
    yPos As Integer     'vertical position (pixels) of block
    block As New BitMapBuffer   'memory DC for bitmap picture (see BitmapBuffer class
                                'and modGameAPI for details on DC's)
End Type

'High scores are read out of a text file as scoreType records.
Private Type scoreType
    hName As String     'player's name
    hScore As Double    'player's high score
    hLevel As Integer   'level player reached
    hLines As Integer   'number of lines player cleared
End Type

'----------------------------------------------------------------------------------
'Variables
'   tileArray - this is the actual tile array.  It is 22 rows high by 12 rows
'               wide.  Each element of the array is of type tileType, and has
'               an x and y position that corresponds to the game screen.
'   Piece     - this array holds the Tetra blocks that the player controls.
'               It is a one-dimensional, 4 element array, and thus can be used
'               for all 7 Tetra pieces.
'   stage     - this is a memory DC that serves as an off-screen staging area;
'               all graphics are blitted here (see modGameAPI for info on blitting)
'               before they are blitted to the game screen.  This allows for
'               flicker-free animation.
'   backStage - a second memory DC used for erasing rows and moving existing rows
'               down.
'   eraseAnim - a memory DC that holds the animation for erasing full rows
'   animStep  - integer; current frame in line-erase animation
'   animDir   - integer; direction of animation; 1 = forward, -1 = reverse
'   fallTime  - integer; time interval for blockTimer based on level
'   fallScore - integer; score given based on how high block was dropped from
'   moveDir   - integer; direction of horizontal movement; 1 = right, -1 = left
'   rot       - integer; current degree of rotation; 0, 90, 180, and 360 degrees
'   curPiece  - integer; current piece being played (1 - 7)
'   nextPiece - integer; next piece to be played (1 - 7)
'   blkFlag   - string; value is set when player selects block type and is used
'               to get appropriate block bitmap from resource file
'   letUp     - boolean; used to force player to release down key when current
'               piece is dropped before next piece can be dropped
'   gameOver  - boolean; true when game ends when blocks reach top of pit
'   quit      - boolean; true when player quits game
'   fullRows  - this array holds the rows that have become completely full and
'               need to be erased
'   colors    - this array holds the text names of 7 colors
'   key*      - these global variables hold the key values for moving the blocks
'               left, right, down, and rotating them clockwise and counter-clockwise
'   hScores   - array of five high scores
Dim tileArray(ARRAYROWS, ARRAYCOLS) As tileType
Dim Piece(4) As blockType
Dim stage As New BitMapBuffer
Dim backStage As New BitMapBuffer
Dim eraseAnim As New BitMapBuffer
Dim blkFlag As String * 1
Dim moveDir, rot, curPiece, nextPiece, fallTime, animDir, animStep, fallScore As Integer
Dim letUp, gameOver, quit As Boolean
Dim fullRows(1 To ARRAYROWS - 2) As Boolean
Dim colors(7) As String
Public keyDown, keyRight, keyLeft, keyCRotate, keyCCRotate As Long
Dim hScores(5) As scoreType

'********************************************************************************
'The procedures in this section load and unload frmTetra, as well as
'initialize game variables by reading and updating the tetra.cfg file.
'********************************************************************************

Private Sub Form_Load()
'All game variables and other game information is initialized in
'this procedure, and the main game engine is called to begin game play.
Dim i, j As Integer     'loop counters
Dim txtString As String     'text buffer

Me.Show     'activate form
Me.Icon = LoadResPicture("TICON", vbResIcon)    'set the icon

bgMusic.DeviceType = "sequencer"    'MIDI sequencer
bgMusic.TimeFormat = 0      '0=milliseconds (ms)

'Set up the array of color names
colors(1) = "Blue"
colors(2) = "Purple"
colors(3) = "Orange"
colors(4) = "Red"
colors(5) = "Green"
colors(6) = "Yellow"
colors(7) = "Grey"

'Create the off-screen staging areas
stage.Create gameScreen.ScaleWidth, gameScreen.ScaleHeight
backStage.Create gameScreen.ScaleWidth, gameScreen.ScaleHeight

'Save the animation picture from the resource file onto disk, then
'load the bitmap from disk into the animation BitmapBuffer.
SavePicture LoadResPicture("PICNOVA", vbResBitmap), App.Path & "\nova.bmp"
eraseAnim.bitmapFile = App.Path & "\nova.bmp"
eraseAnim.Create

'Initialize game variables (block type, music type, etc) from the
'tetra.cfg file in the application folder.
InitFromConfigFile

'The tile array is filled with values from a text file.  This text file is
'stored in the tetra.res resource file.  The file is a simple matrix of 1's
'and 0's resembling the game screen.  I decided to use a text file instead
'of creating the tile array dynamically to allow for extensibility; there
'could be text files corresponding to arrays already filled with random blocks,
'similar to the B-Type games on Tetris.  I did not implement this, but it can
'be done.
Dim byteFile() As Byte
byteFile = LoadResData(101, "CUSTOM")   'get the data from the resource file

'Save the data to a text file on disk
Open App.Path & "\array.dat" For Binary Access Write As #1
    Put #1, , byteFile
Close #1

'Open the file and read it into the tile array.  If you would like to see
'the file, pause the game and look in the application directory for the
'array.dat file.  Open it with a simple text editor like notepad.
Open App.Path & "\array.dat" For Input As #1
    For i = 0 To ARRAYROWS - 1
        Input #1, txtString     'read in an entire row from the file
        For j = 0 To ARRAYCOLS - 1
            With tileArray(i, j)
                .xPos = j * TILEWIDTH
                .yPos = i * TILEHEIGHT
                'Parse the row and fill the .occupied values for each
                'element in the row of the tile array
                .occupied = Mid$(txtString, j + 1, 1)
            End With
        Next j
    Next i
Close #1

'Call the main game engine
Main
End Sub

Private Sub Form_Unload(Cancel As Integer)
'All the DC's are destroyed and files used during the game are erased
'when frmTetra is unloaded.
Dim i As Integer    'loop counter

'Update the config file to reflect player changes
UpdateConfigFile

blockTimer.Enabled = False  'shut down all timers
moveTimer.Enabled = False
fallTimer.Enabled = False
eraseTimer.Enabled = False
bgMusic.Command = "close"   'stop music

stage.Destroy               'destroy all memory DC's
backStage.Destroy
eraseAnim.Destroy

For i = 1 To 4
    Piece(i).block.Destroy
Next i

Kill App.Path & "\array.dat"    'delete temporary files
Kill App.Path & "\block.bmp"
Kill App.Path & "\nova.bmp"

End
End Sub

Private Sub InitFromConfigFile()
'This procedure reads from the tetra.cfg file in the application folder
'to initialize all the game variables, such as music type, block type,
'key settings, etc.
Dim i As Integer     'loop counter
Dim txtBuffer As String     'holds input from file

'First check to make sure the file exists
If Dir(App.Path & "\tetra.cfg") = "" Then
    'If not, then a backup copy is contained in the resource file
    Dim bytFile() As Byte   'will hold file
    bytFile = LoadResData(102, "CUSTOM")    'load the file into memory
    'Save the file to disk.  This is a stock config file with default values.
    Open "tetra.cfg" For Binary Access Write As #1
        Put #1, , bytFile
    Close #1
End If

Open "tetra.cfg" For Input As #1
    skipTo 1, "[Settings]"  'go to the general settings section
    
    'The first value to read is the music type.
    Input #1, txtBuffer
    If Not (txtBuffer = "off") Then     'play music if not off
        menuMusic_Click (Asc(txtBuffer) - 97)
    Else
        menuNoMus_Click     'else turn music off
    End If

    'Next is whether or not the sound effects are on or off
    Input #1, txtBuffer
    menuFX_Click (Val(txtBuffer))
    
    'Next is the type of block
    Input #1, txtBuffer
    menuBlk_Click (Val(txtBuffer))
    
    'Next is whether or not the next piece is shown
    Input #1, txtBuffer
    If txtBuffer = "hide" Then opShowNext_Click
    
    'Next is the time intervals for moveTimer and fallTimer
    Input #1, txtBuffer
    moveTimer.Interval = Val(txtBuffer)
    Input #1, txtBuffer
    fallTimer.Interval = Val(txtBuffer)
    
    skipTo 1, "[Keys]"  'goto the key settings section
    Input #1, txtBuffer     'read in the key values
    keyLeft = Val(txtBuffer)
    Input #1, txtBuffer
    keyRight = Val(txtBuffer)
    Input #1, txtBuffer
    keyDown = Val(txtBuffer)
    Input #1, txtBuffer
    keyCRotate = Val(txtBuffer)
    Input #1, txtBuffer
    keyCCRotate = Val(txtBuffer)
            
    skipTo 1, "[High Scores]"   'goto the high scores section
    'Read in the 5 high scores and other info
    For i = 1 To 5
        Input #1, hScores(i).hName      'player name
        Input #1, txtBuffer
        hScores(i).hScore = Val(txtBuffer)  'player high score
        Input #1, hScores(i).hLines     'player's lines
        Input #1, hScores(i).hLevel     'player's level
    Next i
        
Close #1
End Sub

Private Sub UpdateConfigFile()
'This procedure creates a new tetra.cfg file.  This is done when the
'game ends and the player gets a new high score, or when the player
'quits the game, so that the player's settings will be remembered
'next time the game is played.
Dim i As Integer    'loop counter

'Delete the old config file
If Dir(App.Path & "\tetra.cfg") <> "" Then Kill (App.Path & "\tetra.cfg")

Open "tetra.cfg" For Output As #1
    Print #1, "[Settings]"   'placeholder
    
    'First put the music type, or if the music is off
    If menuNoMus.Checked Then
        Print #1, "off"
    Else
        'Find the type that is checked
        For i = 0 To 5
            If menuMusic(i).Checked Then
                Print #1, Chr$(i + 97)     'lowercase letter
                Exit For
            End If
        Next i
    End If
    
    'Next, put whether sound effects are on or off
    If menuFX(0).Checked Then
        Print #1, 0
    Else
        Print #1, 1
    End If
    
    'Next, put the block type
    Print #1, blkFlag
    
    'Next, put whether the next piece is shown or not
    If opShowNext.Checked Then
        Print #1, "show"
    Else
        Print #1, "hide"
    End If
    
    'Next, put the time intervals for the move and fall timers
    Print #1, moveTimer.Interval
    Print #1, fallTimer.Interval
    
    'Next is the keys section
    Print #1, "[Keys]"
    Print #1, keyLeft
    Print #1, keyRight
    Print #1, keyDown
    Print #1, keyCRotate
    Print #1, keyCCRotate
    
    'The high scores are next
    Print #1, "[High Scores]"
    For i = 1 To 5
        Print #1, hScores(i).hName
        Print #1, hScores(i).hScore
        Print #1, hScores(i).hLines
        Print #1, hScores(i).hLevel
    Next i
Close #1
End Sub

'********************************************************************************
'The procedures in this section handle input from the player and
'are responsible for displaying graphics to the game screen.  The main
'game engine procedure is in this section.
'********************************************************************************

Private Sub generateNext()
'This procedure generates the next piece to be played and displays it
'in the picNext picture box.
Dim i As Integer    'loop counter
Dim block As New BitMapBuffer   'holds block bitmap
picNext.Cls

'Save the appropriate block (use blkFlag) to disk, then load the bitmap into
'memory.
SavePicture LoadResPicture("bl" & colors(nextPiece) & blkFlag, vbResBitmap), App.Path & "\block.bmp"
block.bitmapFile = App.Path & "\block.bmp"
block.Create

'Blit the piece to the picNext picture box
Select Case nextPiece
    Case 1      'straight Piece
        For i = 1 To 4
            BitBlt picNext.hdc, (i - 1) * TILEWIDTH, TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
        Next i
    Case 2      'square block
        For i = 1 To 2
            BitBlt picNext.hdc, i * TILEWIDTH, TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
            BitBlt picNext.hdc, i * TILEWIDTH, 2 * TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
        Next i
    Case 3      'L block
        BitBlt picNext.hdc, 2 * TILEWIDTH, TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
        For i = 1 To 3
            BitBlt picNext.hdc, (i - 1) * TILEWIDTH, 2 * TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
        Next i
    Case 4      'J block
        BitBlt picNext.hdc, 0, TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
        For i = 1 To 3
            BitBlt picNext.hdc, (i - 1) * TILEWIDTH, 2 * TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
        Next i
    Case 5      'Z block
        For i = 1 To 2
            BitBlt picNext.hdc, (i - 1) * TILEWIDTH, TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
            BitBlt picNext.hdc, i * TILEWIDTH, 2 * TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
        Next i
    Case 6      'S block
        For i = 1 To 2
            BitBlt picNext.hdc, i * TILEWIDTH, TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
            BitBlt picNext.hdc, (i - 1) * TILEWIDTH, 2 * TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
        Next i
    Case 7      'T block
        BitBlt picNext.hdc, TILEWIDTH, TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
        For i = 1 To 3
            BitBlt picNext.hdc, (i - 1) * TILEWIDTH, 2 * TILEHEIGHT, TILEWIDTH, TILEHEIGHT, block.hdc, 0, 0, vbSrcCopy
        Next i
End Select
block.Destroy   'this DC is no longer needed
End Sub

Private Sub generatePiece(p As Integer)
'This procedure generates the piece that the player will use.  The piece
'to be generated is determined by p, an integer from 1 to 7.
Dim i, num As Integer   'loop counters
Randomize Timer

'Randomly generate the next piece
nextPiece = Int(Rnd * 7) + 1
generateNext

'Save the appropriate block bitmap to disk (use blkFlag).  Then load the
'bitmap into the memory DC
SavePicture LoadResPicture("bl" & colors(p) & blkFlag, vbResBitmap), App.Path & "\block.bmp"
For i = 1 To 4
    Piece(i).block.Destroy  'destroying the DC before it gets reused ensures
                            'memory is properly deallocated
    Piece(i).block.bitmapFile = App.Path & "\block.bmp"
    Piece(i).block.Create
Next i

rot = 0     'initial rotation angle of the current piece
curPiece = p    'the current piece is of type p

'Set up the initial x and y positions for each of the blocks in Piece
Select Case p
    Case 1      'straight Piece
        For i = 1 To 4
            Piece(i).yPos = 0
            Piece(i).xPos = (i + 2) * TILEWIDTH
        Next i
    Case 2      'square block
        For i = 1 To 2
            Piece(i).yPos = 0
            Piece(i).xPos = (i + 3) * TILEWIDTH
            Piece(i + 2).yPos = TILEHEIGHT
            Piece(i + 2).xPos = (i + 3) * TILEWIDTH
        Next i
    Case 3      'L block
        Piece(1).yPos = 0
        Piece(1).xPos = 5 * TILEWIDTH
        For i = 2 To 4
            Piece(i).yPos = TILEHEIGHT
            Piece(i).xPos = (i + 1) * TILEWIDTH
        Next i
    Case 4      'J block
        Piece(1).yPos = 0
        Piece(1).xPos = 3 * TILEWIDTH
        For i = 2 To 4
            Piece(i).yPos = TILEHEIGHT
            Piece(i).xPos = (i + 1) * TILEWIDTH
        Next i
    Case 5      'Z block
        For i = 1 To 2
            Piece(i).yPos = 0
            Piece(i).xPos = (i + 2) * TILEWIDTH
            Piece(i + 2).yPos = TILEHEIGHT
            Piece(i + 2).xPos = (i + 3) * TILEWIDTH
        Next i
    Case 6      'S block
        For i = 1 To 2
            Piece(i).yPos = 0
            Piece(i).xPos = (i + 3) * TILEWIDTH
            Piece(i + 2).yPos = TILEHEIGHT
            Piece(i + 2).xPos = (i + 2) * TILEWIDTH
        Next i
    Case 7      'T block
        Piece(1).xPos = 4 * TILEWIDTH
        Piece(1).yPos = 0
        For i = 2 To 4
            Piece(i).xPos = (i + 1) * TILEWIDTH
            Piece(i).yPos = TILEHEIGHT
        Next i
End Select

blockTimer.Interval = fallTime  'initialize and set the automatic timer
blockTimer.Enabled = True

'If any tile that the current piece now occupies is already occupied, then
'the game is over.  This check is performed here because the tile array does
'not yet contain the current piece; the array will contain the piece when
'blockTimer is first executed.
For i = 1 To 4
    With Piece(i)
        gameOver = tileArray((.yPos / TILEHEIGHT) + 1, (.xPos / TILEWIDTH) + 1).occupied = 1
    End With
Next i
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'This procedure is used for rotating blocks.  The GetKeyState API is not
'used in this case because holding down the key will over-rotate the piece.
'The keyboard delay used by this procedure will circumvent that and make
'for concise rotation.
If KeyCode = keyCRotate Then    'clockwise
        RotatePiece (1)
ElseIf KeyCode = keyCCRotate Then   'counter-clockwise
        RotatePiece (-1)
End If
End Sub

Private Sub Main()
'This procedure is the main game engine for Tetra.  All input is handled
'here, with the exception of the keys that rotate pieces, which are handled
'in the Form_Keydown procedure.  Also, the engine makes the calls to redraw
'game graphics.
Dim i As Integer        'loop counter
Randomize Timer

fallTime = 850  'time interval for level 0
score = 0       'initialize game data values
numLines = 0
numSingle = 0
numDouble = 0
numTriple = 0
numTetra = 0
level = 0
fallScore = 0
quit = False
letUp = False
gameOver = False
menuOptions.Enabled = True  'these two get disabled at game over
cmdPause.Enabled = True

generatePiece (Int(Rnd * 7) + 1)    'generate the first piece

'Here is the main game engine loop
Do
    If (GetKeyState(keyLeft) And KEY_PRESSED) And Not (moveTimer.Enabled) Then
    'Move left
        moveDir = -1
        moveTimer.Enabled = True
    ElseIf (GetKeyState(keyRight) And KEY_PRESSED) And Not (moveTimer.Enabled) Then
    'Move right
        moveDir = 1
        moveTimer.Enabled = True
    End If

    If (GetKeyState(keyDown) And KEY_PRESSED) And fallTimer.Interval < blockTimer.Interval Then
    'If the user presses the down key and the interval for fallTimer (50ms) is
    'less than the current blockTimer interval, then the blocks will move down
    'faster.  This is known as dropping the blocks.  When the player drops a block
    'into place (it can't move down any farther,) the player must let go of the
    'down key before the next piece can be dropped.  This prevents the next piece
    'from shooting down unexpectedly when the player drops a block into place.
        If Not (letUp) Then
            blockTimer.Enabled = False
            fallTimer.Enabled = True
        End If
    Else
    'When the player lets go of the down key, reset the height-based score
    'and restart the automatic block timer.
        letUp = False
        fallScore = -1
        fallTimer.Enabled = False
        blockTimer.Enabled = True
    End If
        
    redrawBG    'redraw the graphics
    quit = (GetKeyState(vbKeyEscape) And KEY_PRESSED)  'escape key quits
    DoEvents    'This command allows the OS to take care of other processes
                'waiting to be executed and pending requests for service.
                'This line is necessary inside a non-deterministic loop such
                'as this one.  Without it, the OS would give strict attention
                'to this loop and neglect to execute other processes, such as
                'the timers that make up this game and sub-routines for
                'memory management.
Loop Until quit Or gameOver

If gameOver Then doGameOver

'Loop until the player either quits or resets the game
Do
    DoEvents    'take care of pending requests
Loop Until quit

Unload Me   'If the player quits, end the program
End Sub

Private Sub redrawBG()
'This procedure redraws all graphics to the screen.
Dim i As Integer    'loop counter

'First blit the backStage, which contains all the Tetra pieces the
'player has already positioned, onto the stage
BitBlt stage.hdc, 0, 0, gameScreen.ScaleWidth, gameScreen.ScaleHeight, backStage.hdc, 0, 0, vbSrcCopy

'Next, blit the current piece onto the stage
For i = 1 To 4
    BitBlt stage.hdc, Piece(i).xPos, Piece(i).yPos, TILEWIDTH, TILEHEIGHT, Piece(i).block.hdc, 0, 0, vbSrcCopy
Next i

'Finally, blit the stage onto the game screen
BitBlt gameScreen.hdc, 0, 0, gameScreen.ScaleWidth, gameScreen.ScaleHeight, stage.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub doGameOver()
'This procedure fills the screen with blocks to signify the game is over.
Dim i, j As Integer    'loop counters

blockTimer.Enabled = False  'stop all timers
fallTimer.Enabled = False
moveTimer.Enabled = False
eraseTimer.Enabled = False
cmdPause.Enabled = False    'disable the pause button
picNext.Cls
menuOptions.Enabled = False    'disable the options menu
bgMusic.Command = "close"    'stop the background music
'Play game over sound
SndPlaySound App.Path & "\audio\died.wav", SND_ASYNC Or SND_NODEFAULT

'Fill the screen with blocks of the same color as the last piece played
For i = ARRAYROWS - 2 To 0 Step -1
    For j = 0 To ARRAYROWS - 2
        BitBlt gameScreen.hdc, j * TILEWIDTH, i * TILEHEIGHT, TILEWIDTH, TILEHEIGHT, Piece(1).block.hdc, 0, 0, vbSrcCopy
    Next j
Next i

updateHScores      'check for a new high score

End Sub

Private Sub updateHScores()
'This procedure checks the player's score to see if a new high score
'has been reached.  If so, it adds the player's score to the array of
'high scores, dropping off the lowest score.  The high score form is then
'displayed.
Dim i, j As Integer     'loop counters

'If player's score is greater than lowest score, then player has high score
If Val(score) > hScores(5).hScore Then
    Dim usr As String       'get player's name
    usr = InputBox("Enter your name: ", "High Score!")
    
    'Find the highest high score that the player beat
    For i = 1 To 5
        If Val(score) > hScores(i).hScore Then
        'Move all lower scores down one, dropping off the lowest score
            For j = 5 To (i + 1) Step -1
                hScores(j).hName = hScores(j - 1).hName
                hScores(j).hScore = hScores(j - 1).hScore
                hScores(j).hLevel = hScores(j - 1).hLevel
                hScores(j).hLines = hScores(j - 1).hLines
            Next j
            'Insert the player's high score into the array
            hScores(i).hName = usr
            hScores(i).hScore = Val(score)
            hScores(i).hLevel = level
            hScores(i).hLines = numLines
            Exit For
        End If
    Next i
    
    UpdateConfigFile    'Update the tetra.cfg file to reflect new high score
    frmHScore.Show      'show high scores
End If
End Sub

'********************************************************************************
'This section contains the _Timer() event procedures for all four timers.
'********************************************************************************

Private Sub blockTimer_Timer()
'The blockTimer automatically moves blocks down, and is a building block of
'the Tetra game.  A timer control is used so that the blocks will fall at the
'same speed on all computers; since it works on time, the speed of the host
'computer will make no difference.  This philosophy applies to all the timer
'controls.
Dim i, x, y As Integer  'loop counters

If noObstacleDown Then  'check if the tiles below the current piece are occupied
    'If not, then move each block in the current piece down one tile
    For i = 1 To 4
        Piece(i).yPos = Piece(i).yPos + TILEHEIGHT
    Next i
Else
'If the tiles below are occupied, stop the timer
    blockTimer.Enabled = False
    If fallTimer.Enabled Then
        fallTimer.Enabled = False   'if the player dropped the piece manually,
        score = score + fallScore   'increase the score based on height
        letUp = True    'the player must now let up the down key if he/she wishes
                        'to manually drop the next piece
    End If
    
    'Play dropping sound
    If menuFX(0).Checked Then SndPlaySound App.Path & "\audio\drop.wav", SND_ASYNC Or SND_NODEFAULT
    
    'Insert each block of the new piece into the tile array
    For i = 1 To 4
        With Piece(i)
            x = (Piece(i).xPos / TILEWIDTH) + 1     'x position in the tile array
            y = (Piece(i).yPos / TILEHEIGHT) + 1    'y """"
                        
            tileArray(y, x).occupied = 1    'now occupied
            'Rather than blit the blocks to the stage, which changes frequently,
            'the backStage will be used, and the blocks will remain there
            'until they are erased by completing a row.
            BitBlt backStage.hdc, .xPos, .yPos, TILEWIDTH, TILEHEIGHT, .block.hdc, 0, 0, vbSrcCopy
        End With
    Next i
                
    EraseLines  'erase any full rows
    generatePiece (nextPiece)   'the next piece will now be the current piece
End If
End Sub

Private Sub fallTimer_Timer()
'This event happens when the user is holding the down key to drop the blocks
'into place faster.
fallScore = fallScore + 1   'score is based on number of executions of this
                            'procedure, or how high the block is when the
                            'player begins to drop it
'The code for moving blocks down is already contained in the blockTimer
'procedure, so just use it.
blockTimer_Timer
End Sub

Private Sub moveTimer_Timer()
'This procedure moves the current piece horizontally left or right.
Dim i As Integer    'loop counter

If moveDir = -1 And noObstacleLeft Then     'left
    For i = 1 To 4      'check for obstacles to the left
        Piece(i).xPos = Piece(i).xPos - TILEWIDTH
    Next i
ElseIf moveDir = 1 And noObstacleRight Then     'right
    For i = 1 To 4      'check for obstacles to the right
        Piece(i).xPos = Piece(i).xPos + TILEWIDTH
    Next i
End If

moveTimer.Enabled = False   'disable the timer after every move to
                            'prevent over-sensitivity to input
End Sub

Private Sub eraseTimer_Timer()
'This event procedure plays the erase animation to erase complete
'horizontal rows.  It continues to execute until the animation is
'complete, whereupon the timer will be disabled.  For singles, doubles,
'and triples, the animation is played only once.  For Tetra's, the
'animation is played twice, once forward, once backward, for a neat
'effect just for getting a Tetra.
Dim i, j As Integer     'loop counters

'Loop through the fullRows array, checking for rows that are full
For i = 1 To ARRAYROWS - 2
    If fullRows(i) Then
        'When a full row is found, blit the frame specified by animStep
        'onto the row.  This loop is executed each time this procedure is
        'called, producing an animation effect.
        BitBlt gameScreen.hdc, 0, (i - 1) * TILEHEIGHT, gameScreen.ScaleWidth, TILEHEIGHT, eraseAnim.hdc, 0, animStep * TILEHEIGHT, vbSrcCopy
    End If
Next i

animStep = animStep + (animDir * 1) 'go to next frame (either forwards or reverse)
'Stop when the animation is complete (when the frame is either the
'first or last)
If animStep = 12 Or animStep = -1 Then eraseTimer.Enabled = False
End Sub

'********************************************************************************
'This section contains various procedures important to game play.
'********************************************************************************

Private Sub EraseLines()
'This procedure determines if any rows are full, and if so, starts the timer
'control to erase those lines.  This procedure is called every time blockTimer
'is executed.
Dim i, j, x, y As Integer   'loop counters
Dim row As String           'text buffer
Dim lineCount As Integer    'count of how many rows were full

lineCount = 0   'init

'Loop through the tile array row by row.  For each row, check if any
'tiles in that row have an occupied value of 0.  If not, then that
'row is completely full and needs to be erased.
For i = 1 To ARRAYROWS - 2
    'Create the row by concatenating the occupied values of each element
    'of the tile array in a row.
    row = ""
    For j = 1 To ARRAYCOLS - 2
        row = row & tileArray(i, j).occupied
    Next j
        
    If Not (strContains1(row, "0")) Then    'check if the row contains 0's
    'If there are no zero's, then that row is full and will be erased
        fullRows(i) = True      'mark the row as full
        lineCount = lineCount + 1
        numLines = numLines + 1
        'Move the occupied values of all rows above the full row down.
        For y = i - 1 To 0 Step -1
            For x = 1 To 10
                tileArray(y + 1, x).occupied = tileArray(y, x).occupied
            Next x
        Next y
        'The top row will always be cleared, so that blocks that stick above
        'the ceiling will be truncated.  In the original Tetris, this was not
        'a problem since new blocks always started one tile below the ceiling,
        'but I chose to go this route to add that extra little height.
        For x = 1 To 10
            tileArray(1, x).occupied = 0
        Next x
    End If
Next i

If lineCount > 0 Then       'If at least one row was full,
    blockTimer.Enabled = False  'disable all timers
    fallTimer.Enabled = False
    moveTimer.Enabled = False
    animStep = 0       'start animation at first frame and
    animDir = 1        'proceed forwards
    eraseTimer.Enabled = True   'start timer to erase
    'Play erasing sound
    If menuFX(0).Checked Then SndPlaySound App.Path & "\audio\erase.wav", SND_ASYNC Or SND_NODEFAULT
    Do
        DoEvents    'take care of pending requests
    Loop Until eraseTimer.Enabled = False   'idle until animation completes
End If

Select Case lineCount   'increase the line counters
    Case 1
        numSingle = numSingle + 1
    Case 2
        numDouble = numDouble + 1
    Case 3
        numTriple = numTriple + 1
    Case 4
    '4 lines is a Tetra, and it receives a special animation
        numTetra = numTetra + 1
    
        animStep = 11   'start the animtion at the last frame
        animDir = -1    'and proceed backwards
        eraseTimer.Enabled = True   'start timer to erase
        'Play second erasing sound
        If menuFX(0).Checked Then SndPlaySound App.Path & "\audio\tetrasnd.wav", SND_ASYNC Or SND_NODEFAULT
        Do
            DoEvents    'take care of pending requests
        Loop Until eraseTimer.Enabled = False   'Idle until animation complete
End Select

'Now the actual blocks are moved down to take up the space of the erased rows.
For i = 1 To ARRAYROWS - 2
    If fullRows(i) Then
        fullRows(i) = False     'rows are no longer full
        'For each row that was full, move all blocks above it down by one tile
        BitBlt backStage.hdc, 0, TILEHEIGHT, gameScreen.ScaleWidth, (i - 1) * TILEHEIGHT, stage.hdc, 0, 0, vbSrcCopy
        BitBlt backStage.hdc, 0, 0, gameScreen.ScaleWidth, TILEHEIGHT, stage.hdc, 0, 0, vbBlackness     'erase very top row
        BitBlt stage.hdc, 0, 0, gameScreen.ScaleWidth, gameScreen.ScaleHeight, backStage.hdc, 0, 0, vbSrcCopy
    End If
Next i

increaseScore lineCount     'increase the score
checkLevel                  'increase level if necessary
End Sub

Private Sub checkLevel()
'This procedure is responsible for increasing the player's level.

'Divide the current number of lines by 10.  If the integer result is greater
'than the current level, increase the level.
If (numLines \ 10) > level Then
    level = level + 1
    'Play leveling up sound
    If menuFX(0).Checked Then SndPlaySound App.Path & "\audio\levelup.wav", SND_ASYNC Or SND_NODEFAULT
    
    'Decrease the time interval for blockTimer to make the blocks fall faster
    'as the level increases.
    If level < 10 Then      'levels 1-9 decrease 75 ms per level
        fallTime = fallTime - 75
    ElseIf level < 15 Then  'levels 10-15 decrease 15 ms per level
        fallTime = fallTime - 15
    Else                    'levels 15 and above - 10 ms per level
        fallTime = fallTime - 10
    End If
    
    '1 ms is the lowest possible time interval
    If fallTime < 1 Then fallTime = 1
    blockTimer.Interval = fallTime  'set the time interval
End If

End Sub

Private Sub increaseScore(n As Integer)
'This procedure increases the player's score based on the number of
'lines cleared (n) and the player's level.
Select Case n
    Case 1
        score = score + (level + 1) * SINGLEVALUE
    Case 2
        score = score + (level + 1) * DOUBLEVALUE
    Case 3
        score = score + (level + 1) * TRIPLEVALUE
    Case 4
        score = score + (level + 1) * TETRAVALUE
End Select
End Sub

'********************************************************************************
'This section contains the procedures used for collision detection.
'********************************************************************************

Private Function noObstacleDown() As Boolean
'This procedure checks to see if any tiles below the current
'Tetra piece are occupied.
Dim i, x, y As Integer  'loop counters

noObstacleDown = True   'initial assumption

For i = 1 To 4
    x = (Piece(i).xPos / TILEWIDTH) + 1     'x position in tile array
    y = (Piece(i).yPos / TILEWIDTH) + 1     'y position in tile array
    
    If tileArray(y + 1, x).occupied = 1 Then    'if at least one tile is
        noObstacleDown = False                  'occupied, then an obstacle exists.
        Exit For                                'Exit the loop.
    End If
Next i
End Function

Private Function noObstacleRight() As Boolean
'This procedure checks to see if any tiles to the right of the current
'Tetra piece are occupied.
Dim i, x, y As Integer  'loop counters

noObstacleRight = True  'initial assumption

For i = 1 To 4
    x = (Piece(i).xPos / TILEWIDTH) + 1     'x position in tile array
    y = (Piece(i).yPos / TILEWIDTH) + 1     'y position in tile array
    
    If tileArray(y, x + 1).occupied = 1 Then    'if at least one tile is occupied,
        noObstacleRight = False                 'then an obstacle exists.
        Exit For                                'Exit the loop.
    End If
Next i
End Function

Private Function noObstacleLeft() As Boolean
'This procedure checks to see if any tiles to the left of the current
'Tetra piece are occupied.
Dim i, x, y As Integer  'loop counters

noObstacleLeft = True   'initial assumption

For i = 1 To 4
    x = (Piece(i).xPos / TILEWIDTH) + 1     'x position in tile array
    y = (Piece(i).yPos / TILEWIDTH) + 1     'y position in tile array
    
    If tileArray(y, x - 1).occupied = 1 Then     'if at least one tile is occupied,
        noObstacleLeft = False                   'then an obstacle exists.
        Exit For                                 'Exit the loop.
    End If
Next i
End Function

'********************************************************************************
'This section contains the two procedures used to rotate pieces.
'********************************************************************************

Private Sub RotatePiece(clock As Integer)
'This procedure rotates the current Tetra piece.  The clock paramter determines
'the direction of rotation; a value of 1 indicates clockwise rotation, a value
'of -1 means counter-clockwise.  Rather than rotate the piece first and then
'check it, a temporary piece is created that will be rotated.  The temporary
'piece is then checked, and if valid, its values will be assigned to the current
'piece and rotation will be complete.
'Rather than document each and every rotation, I have included a .jpeg file
'which illustrates the various positions in the \help directory in the application
'folder.
Dim posPiece(4) As blockType    'temporary piece
Dim tempRot As Integer          'temporary rotation angle
Dim i As Integer                'loop counter

'Assign the x and y positions of the current piece to the
'temporary piece
For i = 1 To 4
    posPiece(i).xPos = Piece(i).xPos
    posPiece(i).yPos = Piece(i).yPos
Next i

Select Case curPiece
    Case 1          'straight piece
        If rot = 0 Then
            tempRot = 90
                posPiece(1).xPos = Piece(1).xPos + TILEWIDTH
                posPiece(1).yPos = Piece(1).yPos - TILEHEIGHT
                posPiece(3).xPos = Piece(3).xPos - TILEWIDTH
                posPiece(3).yPos = Piece(3).yPos + TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - 2 * TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + 2 * TILEHEIGHT
        ElseIf rot = 90 Then
            tempRot = 0
                posPiece(1).xPos = Piece(1).xPos - TILEWIDTH
                posPiece(1).yPos = Piece(1).yPos + TILEHEIGHT
                posPiece(3).xPos = Piece(3).xPos + TILEWIDTH
                posPiece(3).yPos = Piece(3).yPos - TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + 2 * TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - 2 * TILEHEIGHT
        End If
    'case 2 is the square block - it doesn't rotate
    Case 3          'L block
        If rot = 0 Then
            If clock = 1 Then
                tempRot = 90
                posPiece(1).yPos = Piece(1).yPos + 2 * TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
            Else
                tempRot = 270
                posPiece(1).xPos = Piece(1).xPos - 2 * TILEWIDTH
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
            End If
        ElseIf rot = 90 Then
            If clock = 1 Then
                tempRot = 180
                posPiece(1).xPos = Piece(1).xPos - 2 * TILEWIDTH
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
            Else
                tempRot = 0
                posPiece(1).yPos = Piece(1).yPos - 2 * TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
            End If
        ElseIf rot = 180 Then
            If clock = 1 Then
                tempRot = 270
                posPiece(1).yPos = Piece(1).yPos - 2 * TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
            Else
                tempRot = 90
                posPiece(1).xPos = Piece(1).xPos + 2 * TILEWIDTH
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
            End If
        Else
            If clock = 1 Then
                tempRot = 0
                posPiece(1).xPos = Piece(1).xPos + 2 * TILEWIDTH
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
            Else
                tempRot = 180
                posPiece(1).yPos = Piece(1).yPos + 2 * TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
            End If
        End If
    Case 4      'J block
        If rot = 0 Then
            If clock = 1 Then
                tempRot = 90
                posPiece(1).xPos = Piece(1).xPos + 2 * TILEWIDTH
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
            Else
                tempRot = 270
                posPiece(1).yPos = Piece(1).yPos + 2 * TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
            End If
        ElseIf rot = 90 Then
            If clock = 1 Then
                tempRot = 180
                posPiece(1).yPos = Piece(1).yPos + 2 * TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
            Else
                tempRot = 0
                posPiece(1).xPos = Piece(1).xPos - 2 * TILEWIDTH
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
            End If
        ElseIf rot = 180 Then
            If clock = 1 Then
                tempRot = 270
                posPiece(1).xPos = Piece(1).xPos - 2 * TILEWIDTH
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
            Else
                tempRot = 90
                posPiece(1).yPos = Piece(1).yPos - 2 * TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
            End If
        Else
            If clock = 1 Then
                tempRot = 0
                posPiece(1).yPos = Piece(1).yPos - 2 * TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
            Else
                tempRot = 180
                posPiece(1).xPos = Piece(1).xPos + 2 * TILEWIDTH
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
            End If
        End If
    Case 5      'Z block
        If rot = 0 Then
            tempRot = 90
            posPiece(1).xPos = Piece(1).xPos + TILEWIDTH
            posPiece(1).yPos = Piece(1).yPos - TILEHEIGHT
            posPiece(3).xPos = Piece(3).xPos - TILEWIDTH
            posPiece(3).yPos = Piece(3).yPos - TILEHEIGHT
            posPiece(4).xPos = Piece(4).xPos - 2 * TILEWIDTH
        ElseIf rot = 90 Then
            tempRot = 0
            posPiece(1).xPos = Piece(1).xPos - TILEWIDTH
            posPiece(1).yPos = Piece(1).yPos + TILEHEIGHT
            posPiece(3).xPos = Piece(3).xPos + TILEWIDTH
            posPiece(3).yPos = Piece(3).yPos + TILEHEIGHT
            posPiece(4).xPos = Piece(4).xPos + 2 * TILEWIDTH
        End If
    Case 6      'S block
        If rot = 0 Then
            tempRot = 90
            posPiece(2).xPos = Piece(2).xPos - 2 * TILEWIDTH
            posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
            posPiece(3).yPos = Piece(3).yPos - TILEHEIGHT
        ElseIf rot = 90 Then
            tempRot = 0
            posPiece(2).xPos = Piece(2).xPos + 2 * TILEWIDTH
            posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
            posPiece(3).yPos = Piece(3).yPos + TILEHEIGHT
        End If
    Case 7      'T block
        If rot = 0 Then
            If clock = 1 Then
                tempRot = 90
                posPiece(1).xPos = Piece(1).xPos + TILEWIDTH
                posPiece(1).yPos = Piece(1).yPos + TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
            Else
                tempRot = 270
                posPiece(1).xPos = Piece(1).xPos - TILEWIDTH
                posPiece(1).yPos = Piece(1).yPos + TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
            End If
        ElseIf rot = 90 Then
            If clock = 1 Then
                tempRot = 180
                posPiece(1).xPos = Piece(1).xPos - TILEWIDTH
                posPiece(1).yPos = Piece(1).yPos + TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
            Else
                tempRot = 0
                posPiece(1).xPos = Piece(1).xPos - TILEWIDTH
                posPiece(1).yPos = Piece(1).yPos - TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
            End If
        ElseIf rot = 180 Then
            If clock = 1 Then
                tempRot = 270
                posPiece(1).xPos = Piece(1).xPos - TILEWIDTH
                posPiece(1).yPos = Piece(1).yPos - TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos + TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos - TILEHEIGHT
            Else
                tempRot = 90
                posPiece(1).xPos = Piece(1).xPos + TILEWIDTH
                posPiece(1).yPos = Piece(1).yPos - TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
            End If
        ElseIf rot = 270 Then
            If clock = 1 Then
                tempRot = 0
                posPiece(1).xPos = Piece(1).xPos + TILEWIDTH
                posPiece(1).yPos = Piece(1).yPos - TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos - TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos + TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
            Else
                tempRot = 180
                posPiece(1).xPos = Piece(1).xPos + TILEWIDTH
                posPiece(1).yPos = Piece(1).yPos + TILEHEIGHT
                posPiece(2).xPos = Piece(2).xPos + TILEWIDTH
                posPiece(2).yPos = Piece(2).yPos - TILEHEIGHT
                posPiece(4).xPos = Piece(4).xPos - TILEWIDTH
                posPiece(4).yPos = Piece(4).yPos + TILEHEIGHT
            End If
        End If
End Select

'Check the temporary piece to see if it's valid (i.e. not stuck inside another
'block or going into the walls.)  If it is valid, assign the values from
'the possible piece to the current piece.
If CheckPiece(posPiece) Then
    'Play rotating sound
    If menuFX(0).Checked Then SndPlaySound App.Path & "\audio\rotate.wav", SND_ASYNC Or SND_NODEFAULT
    rot = tempRot       'assign new rotation angle
    For i = 1 To 4      'assign x and y values
        Piece(i).xPos = posPiece(i).xPos
        Piece(i).yPos = posPiece(i).yPos
    Next i
End If
End Sub

Private Function CheckPiece(p() As blockType) As Boolean
'This function checks the Tetra piece p to see if it is valid.  To be valid,
'all four blocks must be in positions in the tile array that are not
'occupied, and they must not extend beyond the walls.
Dim i, x, y As Integer      'loop counters
    
CheckPiece = True   'initial assumption

For i = 1 To 4
    x = (p(i).xPos / TILEWIDTH) + 1     'x position in tile array
    y = (p(i).yPos / TILEHEIGHT) + 1    'y position in tile array
    
    If (x < 0 Or x > 10) Or (y < 0 Or y > 20) Then
    'Make sure x and y positions are within the walls
        CheckPiece = False
        Exit For
    ElseIf (tileArray(y, x).occupied = 1) Then
    'If they are, make sure the tile is not occupied
        CheckPiece = False
        Exit For
    End If
Next i
End Function

'********************************************************************************
'This section contains the procedures for the command and toggle buttons.
'********************************************************************************

Private Sub cmdPause_Click()
'This procedure pauses or unpauses the game when the player clicks
'the Pause toggle button.
If cmdPause.Value = True Then   'If the button is pushed (pause)
    blockTimer.Enabled = False  'disable all game timers
    fallTimer.Enabled = False
    moveTimer.Enabled = False
    lblPause.Visible = True     'show the pause message
    'Loop indefinately.  The only way to exit the loop is to quit the
    'game or click the pause button once again
    Do
        DoEvents    'take care of pending requests
    Loop Until cmdPause.Value = False Or quit
End If
'When unpaused, hide the pause message and restart game timers
lblPause.Visible = False
gameScreen.SetFocus
blockTimer.Enabled = True
End Sub

Private Sub cmdReset_Click()
'This procedure resets the game when the player clicks the
'Reset command button.
Dim i As Integer    'loop counter
gameScreen.SetFocus

'Update any changes made by the player
UpdateConfigFile

blockTimer.Enabled = False  'disable all timers
moveTimer.Enabled = False
fallTimer.Enabled = False
eraseTimer.Enabled = False
cmdPause.Value = False      'unpause the game if paused
bgMusic.Command = "close"   'turn off the music

stage.Destroy               'destroy all memory DC's
backStage.Destroy
eraseAnim.Destroy

For i = 1 To 4
    Piece(i).block.Destroy
Next i

Form_Load    'reload the form, which will restart the game
End Sub

Private Sub cmdQuit_Click()
'Quit when the player clicks the Quit command button.
quit = True
End Sub

'********************************************************************************
'This section contains all the procedures for the menu system.
'********************************************************************************

Private Sub menuNewGm_Click()
'Reset the game when the player selects New Game from the Game menu.
cmdReset_Click
End Sub

Private Sub menuPause_Click()
'Pause or unpause the game when the player selects pause from the Game menu
'or presses the F3 key.
cmdPause.Value = Not (cmdPause.Value)
End Sub

Private Sub menuHScore_Click()
'High Scores option under the File menu.
frmHScore.Show
End Sub

Private Sub menuExit_Click()
'Quit the game when the player selects Exit from the Game menu.
quit = True
End Sub

Private Sub menuMusic_Click(Index As Integer)
'This procedure is called when the player selects a music type
'from the Music option under the Options menu
Dim i As Integer    'loop counter

menuMusic(Index).Checked = True      'check the type the player clicked
menuNoMus.Checked = False            'uncheck No Music option
For i = 0 To Index - 1               'uncheck all types before this one...
    menuMusic(i).Checked = False
Next i
For i = Index + 1 To 5               '...and after it
    menuMusic(i).Checked = False
Next i

'Load and play the new music
bgMusic.Command = "close"
bgMusic.FileName = App.Path & "\audio\tetris" & Chr(Index + 97) & ".mid"
bgMusic.Command = "open"
bgMusic.Command = "play"
End Sub

Private Sub menuNoMus_Click()
'This procedure turns off all music.
Dim i As Integer    'loop counter
bgMusic.Command = "close"   'turn music off

menuNoMus.Checked = True    'check this option
For i = 1 To 5      'uncheck all music types
    menuMusic(i).Checked = False
Next i
End Sub

Private Sub bgMusic_Done(NotifyCode As Integer)
'This procedure restarts the background music when it has finished playing.
'This method is used so that a separate timer isn't needed to periodically
'check if the music is still playing; this procedure is automatically called
'when that happens.

'A NotifyCode of 1 indicates the multimedia control has successfully
'completed the command it was given (play).  Only restart music if
'the No Music option is not checked.
If NotifyCode = 1 And Not (menuNoMus.Checked) Then
    bgMusic.From = 0    'begin playing music at 0 ms (beginning of song)
    bgMusic.Command = "play"
End If
End Sub

Private Sub menuFX_Click(Index As Integer)
'This procedure turns the sound effects on or off.  This option is
'found in the Options menu.
If Index = 0 Then   'Sound effects on
    menuFX(0).Checked = True
    menuFX(1).Checked = False
Else                'Sound effects off
    menuFX(0).Checked = False
    menuFX(1).Checked = True
End If
End Sub

Private Sub menuBlk_Click(Index As Integer)
'This procedure changes the block type when the player selects a type
'of block from the Options menu.
Dim i As Integer    'loop counter

menuBlk(Index).Checked = True   'check the block the player selected
For i = 0 To Index - 1          'uncheck all the types before...
    menuBlk(i).Checked = False
Next i
For i = Index + 1 To 4          '...and after
    menuBlk(i).Checked = False
Next i

blkFlag = Index      'set the flag based on the player's selection
End Sub

Private Sub opShowNext_Click()
'This procedure shows or hides the Next Piece display when the user
'clicks Show Next under the Options menu.

opShowNext.Checked = Not (opShowNext.Checked)  'invert the check mark
shapeNext.Visible = opShowNext.Checked      'show if checked, hide if unchecked
picNext.Visible = opShowNext.Checked
End Sub

Private Sub menuEditCtrl_Click()
'This procedure loads the key editing form when the
'player clicks the Edit Controls option from the Options menu.
frmKeyEdit.Top = Me.Top
frmKeyEdit.Left = Me.Left
frmKeyEdit.Show
End Sub

Private Sub menuHow2Play_Click()
'This procedure opens the html help file the application folder when
'the user clicks on How to Play under the Help menu.  The Shell command
'simply opens the html help file with IE.
If Dir(App.Path & "\help\tetrahelp.html") <> "" Then
    Shell ("explorer " & App.Path & "\help\tetrahelp.html")
Else
    MsgBox "Help file not found."
End If
End Sub

Private Sub menuAbout_Click()
'Show the About form when the player selects About from the Help menu.
frmAbout.Show
End Sub
