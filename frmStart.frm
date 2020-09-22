VERSION 5.00
Begin VB.Form frmStart 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      BackColor       =   &H80000009&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H80000009&
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H80000009&
      Caption         =   "Play Slots!"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblCreator 
      BackStyle       =   0  'Transparent
      Caption         =   "Made by Chan Park"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Slots!"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   4695
   End
   Begin VB.Shape shpBorder2 
      BorderColor     =   &H00000080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   5
      Height          =   2895
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3150
   End
   Begin VB.Shape shpBorder1 
      BorderColor     =   &H00000040&
      BorderWidth     =   5
      Height          =   2775
      Left            =   120
      Top             =   240
      Width           =   3600
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'**************************************
'Initiates border dimensions/positions*
'**************************************
Private Sub Form_Load()
With shpBorder1
    .Top = 22
    .Left = 20
    .Width = Me.Width - 65
    .Height = Me.Height - 65
End With

With shpBorder2
    .Top = 25
    .Left = 20
    .Width = Me.Width - 75
    .Height = Me.Height - 75
End With
End Sub
'**************************************
'**************************************


'**************************************
'Hide me, show game screen*************
'**************************************
Private Sub cmdPlay_Click()
Me.Hide
frmMain.Show
End Sub
'**************************************
'**************************************


'**************************************
'Show info screen**********************
'**************************************
Private Sub cmdHelp_Click()
bHelpFrom = True    'was accessed from start menu
frmHelp.Show
End Sub
'**************************************
'**************************************


'**************************************
'Show message upon exiting*************
'**************************************
Private Sub cmdQuit_Click()
MsgBox "Thanks for playing 'Slots!'", , "Goodbye"
End
End Sub
'**************************************
'**************************************
