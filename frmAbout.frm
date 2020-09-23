VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Tetra"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   3000
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Original design, concept and program by Alexey Pazhitnov"
      Height          =   480
      Left            =   735
      TabIndex        =   3
      Top             =   2580
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      Height          =   255
      Left            =   2610
      TabIndex        =   2
      Top             =   1920
      Width           =   885
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "crosslight@aol.com"
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   90
      TabIndex        =   1
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "(c) 2001 Seth Guenther"
      Height          =   270
      Left            =   90
      TabIndex        =   0
      Top             =   1920
      Width           =   1845
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////////////
'frmAbout
'created: 5/25/01
'last mod date: 6/29/01
'
'Just a simple About form.
'/////////////////////////////////////////////////////////////////////////

Private Sub Form_Load()
Me.Icon = LoadResPicture("TICON", vbResIcon)    'load game icon
End Sub
