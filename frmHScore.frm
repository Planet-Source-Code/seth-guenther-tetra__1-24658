VERSION 5.00
Begin VB.Form frmHScore 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "frmHScore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   17
      Left            =   2865
      TabIndex        =   21
      Top             =   3495
      Width           =   1260
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      Index           =   3
      X1              =   14
      X2              =   445
      Y1              =   221
      Y2              =   221
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      Index           =   2
      X1              =   14
      X2              =   445
      Y1              =   177
      Y2              =   177
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      Index           =   1
      X1              =   14
      X2              =   445
      Y1              =   130
      Y2              =   130
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      Index           =   0
      X1              =   14
      X2              =   445
      Y1              =   85
      Y2              =   85
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   19
      Left            =   5850
      TabIndex        =   23
      Top             =   3495
      Width           =   510
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   18
      Left            =   4380
      TabIndex        =   22
      Top             =   3495
      Width           =   750
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   16
      Left            =   240
      TabIndex        =   20
      Top             =   3570
      Width           =   2385
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   15
      Left            =   5865
      TabIndex        =   19
      Top             =   2820
      Width           =   510
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   14
      Left            =   4380
      TabIndex        =   18
      Top             =   2820
      Width           =   750
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   13
      Left            =   2865
      TabIndex        =   17
      Top             =   2820
      Width           =   1260
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   12
      Left            =   240
      TabIndex        =   16
      Top             =   2910
      Width           =   2385
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   5865
      TabIndex        =   15
      Top             =   2145
      Width           =   510
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Index           =   10
      Left            =   4380
      TabIndex        =   14
      Top             =   2145
      Width           =   750
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   9
      Left            =   2865
      TabIndex        =   13
      Top             =   2145
      Width           =   1260
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   240
      TabIndex        =   12
      Top             =   2220
      Width           =   2385
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   7
      Left            =   5865
      TabIndex        =   11
      Top             =   1500
      Width           =   510
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   4380
      TabIndex        =   10
      Top             =   1500
      Width           =   750
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   5
      Left            =   2865
      TabIndex        =   9
      Top             =   1500
      Width           =   1260
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   8
      Top             =   1575
      Width           =   2385
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5865
      TabIndex        =   7
      Top             =   855
      Width           =   495
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4380
      TabIndex        =   6
      Top             =   855
      Width           =   750
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   2865
      TabIndex        =   5
      Top             =   855
      Width           =   1260
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      FillColor       =   &H0000FFFF&
      Height          =   3360
      Left            =   165
      Shape           =   4  'Rounded Rectangle
      Top             =   585
      Width           =   6585
   End
   Begin VB.Label labels 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   915
      Width           =   2385
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      FillColor       =   &H0000FFFF&
      Height          =   3255
      Left            =   5370
      Shape           =   4  'Rounded Rectangle
      Top             =   630
      Width           =   30
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      FillColor       =   &H0000FFFF&
      Height          =   3270
      Left            =   4185
      Shape           =   4  'Rounded Rectangle
      Top             =   630
      Width           =   30
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      FillColor       =   &H0000FFFF&
      Height          =   3285
      Left            =   2670
      Shape           =   4  'Rounded Rectangle
      Top             =   615
      Width           =   30
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   5520
      TabIndex        =   3
      Top             =   225
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Lines"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   4275
      TabIndex        =   2
      Top             =   225
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2850
      TabIndex        =   1
      Top             =   225
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "ZeroHour"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   930
      TabIndex        =   0
      Top             =   225
      Width           =   1155
   End
End
Attribute VB_Name = "frmHScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////////////
'frmHScore
'created: 5/25/01
'last mod date: 6/29/01
'
'This form simply displays the high scores.  It reads in the scores from
'the 'tetra.cfg' file in the application directory and displays them on
'the screen.
'
'The following controls are used in this form:
'   labels  - array of labels; used to display high scores
'Other shape and line controls are used.
'//////////////////////////////////////////////////////////////////////////////

Private Sub Form_Load()
Dim i As Integer    'loop counter
Dim field As String    'string buffer

Me.Icon = LoadResPicture("TICON", vbResIcon)    'load game icon

Open "tetra.cfg" For Input As #1
'Input from the file to the buffer and display on the screen.
'The file is structured in such a way that the sequential reading in and
'placement of text will cause the output to appear properly on the screen.
skipTo 1, "[High Scores]"       'skip to high scores section
For i = 0 To 19
    Input #1, field
    labels(i) = field
Next i

Close #1

End Sub
