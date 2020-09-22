VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label lblCombo10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Sevens : 10x Original Bet"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1335
      TabIndex        =   13
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label lblCombo9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Dollars : 9x Original Bet"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1335
      TabIndex        =   12
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblCombo8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Cherries : 8x Original Bet"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1275
      TabIndex        =   11
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label lblCombo7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Lemons : 7x Original Bet"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1380
      TabIndex        =   10
      Top             =   2880
      Width           =   2685
   End
   Begin VB.Label lblCombo5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Characters ($'s and 7's) : 5x Original Bet"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   9
      Top             =   2400
      Width           =   5415
   End
   Begin VB.Label lblCombo4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Fruits : 4x Original Bet"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   8
      Top             =   2160
      Width           =   5385
   End
   Begin VB.Label lblCombo6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All Bars : 6x Original Bet"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   5445
   End
   Begin VB.Label lblCombo3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Two Sevens : 3x Original Bet"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   5445
   End
   Begin VB.Label lblCombo2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Two Dollars: 2x Original Bet"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label lblCombo1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Two Bars : 1x Original Bet"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label lblCombosLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Combos"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   5460
   End
   Begin VB.Label lblObjective 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To walk in with $5.00 and walk out with $50.00"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   5415
   End
   Begin VB.Label lblObjectiveLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Objective"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin VB.Shape shpBorder2 
      BorderColor     =   &H00000080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   5
      Height          =   2055
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3015
   End
   Begin VB.Shape shpBorder1 
      BorderColor     =   &H00000040&
      BorderWidth     =   5
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "frmHelp"
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
    .Width = Me.Width - 40
    .Height = Me.Height - 50
End With

With shpBorder2
    .Top = 25
    .Left = 20
    .Width = Me.Width - 50
    .Height = Me.Height - 60
End With
End Sub
'**************************************
'**************************************


'**************************************
'Depending on where the form is opened*
'from, will hide or show the OK button*
'**************************************
Private Sub Form_Activate()
If bHelpFrom = True Then 'if opened from start menu
    cmdOK.Visible = True    'show the button
Else                     'if opened from game screen
    cmdOK.Visible = False   'hide the button
End If
End Sub
'**************************************
'**************************************


'**************************************
'Hide me and show start menu***********
'**************************************
Private Sub cmdOK_Click()
Me.Hide
frmStart.Show
End Sub
'**************************************
'**************************************


