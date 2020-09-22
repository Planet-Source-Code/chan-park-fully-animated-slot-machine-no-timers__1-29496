VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Slots!"
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   4  'Icon
   ScaleHeight     =   6045
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H8000000E&
      Caption         =   "Exit Game"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdBetUp 
      Caption         =   "/\"
      Height          =   510
      Left            =   5475
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4140
      Width           =   465
   End
   Begin VB.CommandButton cmdBetDown 
      Caption         =   "\/"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7095
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4140
      Width           =   465
   End
   Begin VB.PictureBox picHdl 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Enabled         =   0   'False
      Height          =   3450
      Left            =   8370
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   3450
      ScaleWidth      =   1965
      TabIndex        =   4
      Top             =   1095
      Width           =   1965
   End
   Begin VB.Timer tmrLoad 
      Interval        =   10
      Left            =   480
      Top             =   5400
   End
   Begin VB.PictureBox picSlot 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   1920
      Index           =   2
      Left            =   6510
      ScaleHeight     =   1920
      ScaleWidth      =   1350
      TabIndex        =   3
      Top             =   1160
      Width           =   1350
   End
   Begin VB.PictureBox picSlot 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   1920
      Index           =   1
      Left            =   4530
      ScaleHeight     =   1920
      ScaleWidth      =   1350
      TabIndex        =   2
      Top             =   1160
      Width           =   1350
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   525
      Left            =   8895
      MousePointer    =   2  'Cross
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1125
      Width           =   570
   End
   Begin VB.PictureBox picSlot 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   1920
      Index           =   0
      Left            =   2460
      ScaleHeight     =   1920
      ScaleWidth      =   1350
      TabIndex        =   0
      Top             =   1160
      Width           =   1350
   End
   Begin VB.Label lblSpins 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   5895
      TabIndex        =   16
      Top             =   5055
      Width           =   1215
   End
   Begin VB.Label lblGoal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "$50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   3240
      TabIndex        =   15
      Top             =   5055
      Width           =   1215
   End
   Begin VB.Label lblSpinsLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Spins"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5550
      TabIndex        =   14
      Top             =   4695
      Width           =   1950
   End
   Begin VB.Label lblGoalLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Goal"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3120
      TabIndex        =   13
      Top             =   4695
      Width           =   1440
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Welcome!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1965
      TabIndex        =   12
      Top             =   3345
      Width           =   6390
   End
   Begin VB.Label lblBetLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Your Bet:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5715
      TabIndex        =   9
      Top             =   3825
      Width           =   1590
   End
   Begin VB.Label lblMoneyLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Cash:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3105
      TabIndex        =   8
      Top             =   3840
      Width           =   1590
   End
   Begin VB.Label lblBet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0.25"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   5895
      TabIndex        =   7
      Top             =   4185
      Width           =   1215
   End
   Begin VB.Label lblMoney 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4.75"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   420
      Left            =   3225
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Slots!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   855
      Left            =   3885
      TabIndex        =   5
      Top             =   120
      Width           =   2775
   End
   Begin VB.Shape shpBorder1 
      BorderColor     =   &H00000040&
      BorderWidth     =   5
      Height          =   6015
      Left            =   0
      Top             =   0
      Width           =   10440
   End
   Begin VB.Shape shpBorder2 
      BorderColor     =   &H00000080&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   5
      Height          =   6015
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   9990
   End
   Begin VB.Shape shpSlotBorder 
      BackColor       =   &H80000008&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   5
      Height          =   2055
      Index           =   2
      Left            =   6450
      Top             =   1095
      Width           =   1470
   End
   Begin VB.Shape shpSlotBorder 
      BackColor       =   &H80000008&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   5
      Height          =   2055
      Index           =   1
      Left            =   4470
      Top             =   1095
      Width           =   1470
   End
   Begin VB.Shape shpSlotBorder 
      BackColor       =   &H80000008&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0FF&
      BorderWidth     =   5
      Height          =   2055
      Index           =   0
      Left            =   2400
      Top             =   1095
      Width           =   1470
   End
   Begin VB.Shape shpSlotBack 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   7410
      Left            =   1950
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6420
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'*   Program Name: SLOTS!                                         *
'*   Made by: Chan Park                                           *
'*       Note: - Repeated runnings may deplete memory and         *
'*               cause computer to crash, and/or cause            *
'*               program to malfunction                           *
'******************************************************************

Option Explicit

Dim bMouseDwn As Boolean
Dim sMouseXPos As Single
Dim sMouseYPos As Single

Dim NoOfLoops As Long

Const GOAL As Integer = 50          'Sets the Target amount of money
Const MNYSTRT As Currency = 5       'Sets the starting amount of money
Const BETITVL As Currency = 0.25    'Sets interval of bets

Dim iSpins As Integer               'Keeps track of spins
Dim cMoney As Currency              'Keeps track of money
Dim cBet As Currency                'Keeps track of bet

'-------------------------------------
'Stores filenames
Const SVNF As String = "\1SVN.bmp"
Const DLRF As String = "\2DLR.bmp"
Const CHRYF As String = "\3CHRY.bmp"
Const LMNF As String = "\4LMN.bmp"
Const BARF As String = "\5BAR.bmp"
'-------------------------------------

'-------------------------------------
'Sets values to represent the different
'pictures
Const SVN As Integer = 0
Const DLR As Integer = 1
Const CHRY As Integer = 2
Const LMN As Integer = 3
Const BAR As Integer = 4
'-------------------------------------

Dim sPicName(4) As String           'Stores names of pictures

'-------------------------------------
'Arrays to store the variables needed
'for animation and image loading
Dim lPicDC(4) As Long               'Stores device contexts of loaded pictures
Dim lPicH(4) As Long                'Stores handles of loaded pictures
Dim lPicF(4) As String              'Stores filename constants
'-------------------------------------

'-------------------------------------
'Stores dimensions of the slots in
'pixels
Dim iSlotPixW As Integer
Dim iSlotPixH As Integer
'-------------------------------------

Const ANIF As String = "\Animation\HdlAnim"     'path of handle animation

Dim lAniDCH As Long         'device context of blue handle highlight image
Dim lAniHH As Long          'handle of blue handle highlight image

Dim lAniDC(22) As Long      'device contexts of images used in handle animation
Dim lAniH(22) As Long       'handles of images used in handle animation
Dim iHdlPixW As Integer     'dimensions of the picture box containing
Dim iHdlPixH As Integer         'the handle in pixels

Dim bRollin As Boolean      'has the rolling of the slots been initiated?

Dim iRslt(2) As Integer     'stores the 3-number combo resulting the roll
Dim iPRslt(2) As Integer    'stores the previous round's 3-number combo


'*********************************************************
' Various initialization values***************************
' Loads images to be used from file***********************
'*********************************************************
Private Sub Form_Load()
Dim i As Integer                    'miscellaneous variable for use as counter etc.

bErrorCheck = False
NoOfLoops = LoopsPerSecond(True, 10000, 20)

'-------------------------------------
'Initial positions and dimenstions
'of the borders
With shpBorder1
    .Top = .Top + 22
    .Left = .Left + 20
    .Width = Me.Width - 65
    .Height = Me.Height - 65
End With

With shpBorder2
    .Top = .Top + 25
    .Left = .Left + 20
    .Width = Me.Width - 75
    .Height = Me.Height - 75
End With
'-------------------------------------

'-------------------------------------
'Puts file names of the pictures into
'an array
lPicF(0) = SVNF
lPicF(1) = DLRF
lPicF(2) = CHRYF
lPicF(3) = LMNF
lPicF(4) = BARF
'-------------------------------------

'-------------------------------------
'Puts names of pictures into array
sPicName(0) = "Seven"
sPicName(1) = "Dollar"
sPicName(2) = "Cherrie"
sPicName(3) = "Lemon"
sPicName(4) = "Bar"
'-------------------------------------

'-------------------------------------
'Loads the picture files for animation
For i = 0 To 4
    lPicDC(i) = GenerateDC(App.Path & lPicF(i), lPicH(i))
    If lPicDC(i) = 0 Then
        bErrorCheck = 1
    End If
Next
'-------------------------------------

'-------------------------------------
'Loads images for handle animation and
'stores their DC's in an array
For i = 0 To 22         '23 images in the animation
    'puts generated DC's into lAniDC(), passing the function the images' file
    'names and a place for the function to put their handles in
    lAniDC(i) = GenerateDC(App.Path & ANIF & i + 1 & ".bmp", lAniH(i))
    If lAniDC(i) = 0 Then 'there is an error
        bErrorCheck = 1
    End If
Next
'-------------------------------------

'assigns a DC for the blue highlight image
lAniDCH = GenerateDC(App.Path & ANIF & "H" & ".bmp", lAniHH)

Call ErrorCheck 'has there been an error?

'-------------------------------------
'Converts twips value of dimensions of
'the Slots to pixels and stores it
iSlotPixW = picSlot(0).Width        'stores twips values
iSlotPixH = picSlot(0).Height
Call TwipsToPix(iSlotPixW, iSlotPixH)   'converts twips to pixels
'-------------------------------------

'-------------------------------------
'Converts twips value of dimensions of
'the Handle to pixels and stores it
iHdlPixW = picHdl.Width
iHdlPixH = picHdl.Height
Call TwipsToPix(iHdlPixW, iHdlPixH)
'-------------------------------------

'Initiates a timer to create a delay before the
'initial images (three sevens) are loaded
tmrLoad.Interval = 10
tmrLoad.Enabled = True

'-------------------------------------
'Initial values
For i = 0 To 2
    iRslt(i) = Empty
Next

cMoney = MNYSTRT
lblMoney = "$" & MNYSTRT
cBet = BETITVL * 1
lblBet = "$" & cBet

lblStatus = "Welcome"

iSpins = 0
lblSpins = iSpins

cmdBetUp.Enabled = True
cmdBetDown.Enabled = False
'-------------------------------------
End Sub
'***************************************************
'***************************************************


'***************************************************
'Sub routine that converts passed X and Y twips*****
'values into pixels*********************************
'***************************************************
Sub TwipsToPix(ByRef intTwipX As Integer, ByRef intTwipY As Integer)
intTwipX = intTwipX \ Screen.TwipsPerPixelX
intTwipY = intTwipY \ Screen.TwipsPerPixelY
End Sub
'***************************************************
'***************************************************


'***************************************************
'Used to create a delay when loading initial slot***
'pictures.  Otherwise, the image will load before***
'the form/picture box does, resulting in blank slots
'***************************************************
Private Sub tmrLoad_Timer()
Dim i As Integer
For i = 0 To 2
    picSlot(i).Picture = LoadPicture(App.Path & SVNF)
Next
tmrLoad.Enabled = False
End Sub
'***************************************************
'***************************************************


'***************************************************
'These buttons raise and lower the betting amount***
'accordingly in intervals of BETITVL and with a max*
'amount of $2.5*************************************
'***************************************************
Private Sub cmdBetUp_Click()    'raises bet
cBet = cBet + BETITVL
cmdBetDown.Enabled = True 'because when bet is raised, it can always be lowered

If cBet = 2.5 Then              'sets max bet value
    cmdBetUp.Enabled = False
End If

If cBet = cMoney Then           'makes sure that you can't bet more than you have
    cmdBetUp.Enabled = False
End If

lblBet = "$" & cBet             'refreshes value of bet on to label
End Sub

Private Sub cmdBetDown_Click()  'lowers bet
cBet = cBet - BETITVL
cmdBetUp.Enabled = True 'because when bet is lowered, it can always be raised

If cBet = 0.25 Then             'sets min bet value
    cmdBetDown.Enabled = False
End If

lblBet = "$" & cBet
End Sub
'***************************************************
'***************************************************

'***************************************************
'Initiates the Slot Machine*************************
'***************************************************
Private Sub cmdStart_click()
Dim i As Integer
Dim n As Long
Dim z As Long

cmdStart.Visible = False    'so that button can't be clicked until all ends
lblStatus = "Spinning..."
cMoney = cMoney - cBet
lblMoney = "$" & cMoney
bRollin = True
DoEvents                    'creates a pause so that above changes can take effect
                                'before anything else

'-------------------------------------
'Animates the Handle
For i = 1 To 22
    Call AniHdl(i)          'calls function to redraw each frame of animation
    For z = 1 To NoOfLoops / 40  'creates a delay so that
        DoEvents                    'animation can
    Next                            'look natural
Next
'-------------------------------------

Call Roll                   'Roll the slots and make them stop
Call CalcWin                'Did the user win? If so, how much?

bRollin = False
cmdStart.Visible = True

'-------------------------------------
'If user bet an amount that exceeds
'the current amount after winnings are
'calculated, bet amount is lowered
If cBet > cMoney Then
    cBet = cMoney
    lblBet = "$" & cBet
    cmdBetUp.Enabled = False 'because bet is already at maximum value
    If cBet = 0.25 Then
        cmdBetDown.Enabled = False 'so bet can't also be lowered if cMoney = 0.25
    End If
End If
'-------------------------------------
End Sub
'***************************************************
'***************************************************

'***************************************************
'Sub routine that draws an image of the handle******
'animation into picHdl by accepting which of the****
'images to draw as a value**************************
'***************************************************
Sub AniHdl(ByVal n As Integer)
Dim lReturn As Long
lReturn = BitBlt(picHdl.hdc, 0, 0, iHdlPixW, iHdlPixH, _
                lAniDC(n), 1, 0, SRCCOPY)

If lReturn = 0 Then 'an error has occurred
    bErrorCheck = True
End If

Call ErrorCheck
End Sub
'***************************************************
'***************************************************

'***************************************************
'Handles the actual animation of the pictures on the
'slot, and makes them stop randomly*****************
'***************************************************
Sub Roll()
Dim z As Long
Dim c As Integer
Dim iRevs As Integer            'Number of revolutions before stopping
Dim iPic As Integer             'Which picture to draw
Dim iOffset As Integer          'How far from the top to draw the picture
Dim iSlot As Integer            'Which slot to draw the picture on

Dim bStopSlot(2) As Boolean     'Should a slot be stopped?
Dim iSlotStop As Integer        'Which slot is being stopped

'-------------------------------------
'Generates the 3-number combo that
'represents which pictures to show
For c = 0 To 2
    Randomize
    iPRslt(c) = iRslt(c)    'iPRslt = value of iRslt before a new value is created
    iRslt(c) = Int(5 * Rnd) 'Min = 0, Max = 4 (because it rounds down)
Next
'-------------------------------------


'-------------------------------------
'Animates the slots before they are
'stopped
iRevs = 2
For c = 1 To iRevs      'Every revolution
    For iPic = SVN To BAR   'goes through each picture in order,
        For iOffset = 0 To 127  'draws them from the top of the slot to the bottom,
            For iSlot = 0 To 2      'in each slot.
                'In each slot, a maximum of 2 pictures can appear at a time
                'therefore, the first picture slowly goes out of view
                For z = 1 To NoOfLoops / 3000
                    DoEvents
                Next
                Call DrawSlot(iSlot, (iPic + iPRslt(iSlot)) Mod 5, iOffset)
                'while the picture after it, iPic + 1, slowly comes into view
                Call DrawSlot(iSlot, ((iPic + 1) + iPRslt(iSlot)) Mod 5, iOffset - 127)
            Next
        Next
    Next
Next
'-------------------------------------

'-------------------------------------
'Stops the animation in each slot
'sequentially so that there is a delay
'after every slot stops
For iSlotStop = 0 To 2  'The slots will stop sequentially,
    For iPic = SVN To BAR   'showing all pictures sequentially,
        For iOffset = 0 To 127  'moving from top of slot to bottom,
            For iSlot = 0 To 2      'in all slots
                'If the current picture of the slot thats supposed to stop
                'corresponds to what its supposed to stop at
                If (iPic + iPRslt(iSlotStop)) Mod 5 = iRslt(iSlotStop) Then
                    'always stop this slot from moving
                    bStopSlot(iSlotStop) = True
                End If
                If bStopSlot(iSlot) = False Then    'all other slots
                For z = 1 To NoOfLoops / 3200
                    DoEvents
                Next
                    'continue animating
                    Call DrawSlot(iSlot, (iPic + iPRslt(iSlot)) Mod 5, iOffset)
                    Call DrawSlot(iSlot, ((iPic + 1) + iPRslt(iSlot)) Mod 5, iOffset - 127)
                End If
            Next
        Next
    Next
Next
'-------------------------------------

'-------------------------------------
'Makes the current pictures in each
'slot become static(not disappear when
', for example, another window moves
'on top of the game window)
For c = 0 To 2
    picSlot(c) = LoadPicture(App.Path & lPicF(iRslt(c)))
Next
'-------------------------------------
End Sub
'***************************************************
'***************************************************


'***************************************************
'This one, like sub AniHdl, handles the animation of
'the slots, receiving which slot to draw, which*****
'picture in, at how much of a y offset as values****
'***************************************************
Sub DrawSlot(ByVal iSlot As Integer, ByVal iPic As Integer, ByVal iYOffset As Integer)
Dim lReturn As Long
lReturn = BitBlt(picSlot(iSlot).hdc, 0, iYOffset, iSlotPixW, iSlotPixH, _
                lPicDC(iPic), 0, 0, SRCCOPY)

If lReturn = 0 Then 'an error has occurred
    bErrorCheck = True
End If

Call ErrorCheck
End Sub
'***************************************************
'***************************************************


'***************************************************
'Did the user win? If so, by how much?**************
'Did the user beat the game? or did the user lose?**
'***************************************************
Sub CalcWin()
Dim iCheck As Integer   'Did the user win? If so what combo did the user get?
Dim lPlay As Long       'The user won or lost. Will he/she play again?

iCheck = CheckWin(iRslt())  'Did the user win? If so what combo did the user get?

'-------------------------------------
'Telling the user what happened
If iCheck = 0 Then
    lblStatus = "You lose"
ElseIf iCheck = 1 Then
    lblStatus = "You got two Bars: you get your bet of $" & cBet & " back"
ElseIf iCheck = 2 Then
    lblStatus = "You got two Dollars: you get 2 times your bet, $" & 2 * cBet
ElseIf iCheck = 3 Then
    lblStatus = "You got two Sevens: you get 3 times your bet, $" & 3 * cBet
ElseIf iCheck = 4 Then
    lblStatus = "You got all fruits: you get 4 times your bet, $" & 4 * cBet
ElseIf iCheck = 5 Then
    lblStatus = "You got all characters: you get 5 times your bet, $" & 5 * cBet
Else
    lblStatus = "You got all " & sPicName(10 - iCheck) & "s: you get " & _
                iCheck & " times your bet, $" & iCheck * cBet
End If
'-------------------------------------

cMoney = cMoney + (cBet * iCheck)
lblMoney = "$" & cMoney

iSpins = iSpins + 1
lblSpins = iSpins

'--------------------------------------
'Will the user play again after losing?
If cMoney = 0 Then
    lblStatus = "You are broke!"
    lPlay = MsgBox("Play Again?", vbYesNo, "Game Over")
    If lPlay = vbYes Then   'If the user will play, then
        Call Restart            'reset the game
    Else                    'Otherwise, if the user won't play again, then
        Call Restart            'reset game in case user comes back,
        Me.Hide                 'hide this form,
        frmStart.Show           'and show the Starting form
    End If
End If
'-------------------------------------

'---------------------------------------
'Will the user play again after winning?
If cMoney >= 50 Then
    lblStatus = "You reached your goal!"
    lPlay = MsgBox("You got $" & GOAL & " in " & iSpins & _
                    " spins. Play again?", vbYesNo, "You Won!")
    If lPlay = vbYes Then
        Call Restart
    Else
        Call Restart
        Me.Hide
        frmStart.Show
    End If
End If
'-------------------------------------
End Sub
'***************************************************
'***************************************************


'***************************************************
'Resets all required settings for a game to restart*
'***************************************************
Sub Restart()
Dim i As Integer

tmrLoad.Interval = 10
tmrLoad.Enabled = True

For i = 0 To 2
    iRslt(i) = Empty
Next

cMoney = MNYSTRT
lblMoney = "$" & MNYSTRT
cBet = BETITVL * 1
lblBet = "$" & cBet

lblStatus = "Welcome"

iSpins = 0
lblSpins = iSpins

cmdBetUp.Enabled = True
cmdBetDown.Enabled = False
End Sub
'***************************************************
'***************************************************


'***************************************************
'Did the user win or lose? If so, what combo did****
'the user get?**************************************
'By passing the iRslt(), this function determines***
'which combo was obtained, if there was one at all**
'***************************************************
Function CheckWin(ByRef iRslt() As Integer) As Integer
Dim j As Integer
Dim i As Integer
Dim c As Integer
'this variable represents the value of the combo obtained. combos are represented
'by numbers reflecting their values. i.e.:when there are all sevens,
'then icheck = 10, the highest value
Dim iCheck As Integer

'-------------------------------------
'Did the user get the same of
'everything?
j = iRslt(0)
For i = 1 To 2
    If j = iRslt(i) Then
        iCheck = 10 - j 'if so, then his bet is multiplied by 10 minus the number
                            'that he got(represented by constants SVN, DLR etc.)
    Else
        iCheck = 0
        Exit For
    End If
Next
'-------------------------------------

'-------------------------------------
'Did the user get all fruits?
If iCheck = 0 Then  'if the user still didn't get a rewardable combo then
    For i = 0 To 2      'for every of the 3 numbers
        If iRslt(i) = CHRY Or iRslt(i) = LMN Then   'check if they are either a
                                                    'cherry or lemon
            iCheck = 4
        Else
            iCheck = 0  'if even one of them aren't a fruit,
            Exit For        'stop checking
        End If
    Next
End If
'-------------------------------------

'-------------------------------------
'Did the user get all characters?
If iCheck = 0 Then
    For i = 0 To 2
        If iRslt(i) = SVN Or iRslt(i) = DLR Then
            iCheck = 5
        Else
            iCheck = 0
            Exit For
        End If
    Next
End If
'-------------------------------------

'-------------------------------------
'Did the user get 2 sevens, dollars,
'or bars?
If iCheck = 0 Then  'if the combo is not the same of everything then
    For c = 4 To 6      'For each of the three pictures being checked,
                            'sevens, dollars, and bars,
        For i = 0 To 2          'check whether:
            'if iRslt(0) and (1), or (1) and (2), or (2) and (1), are equal
            If iRslt(i) = c Mod 5 And iRslt((i + 1) Mod 3) = c Mod 5 Then
                'if they are and
                If c = 4 Then 'the user got 2 bars
                    iCheck = 1
                ElseIf c = 5 Then 'the user got 2 sevens
                    iCheck = 3
                Else
                    iCheck = 2 'the user got 2 dollars
                End If
            End If
        Next
    Next
End If
'-------------------------------------

CheckWin = iCheck       'returns the value of the obtained combo
End Function
'***************************************************
'***************************************************


'***************************************************
'Unload all loaded images**************************
'***************************************************
Private Sub Form_Terminate()
Dim i As Integer

'--------------------------------------
'Unload the DCs of all Handle animation
'images
For i = 0 To 22
    DeleteDC lAniDC(i)
Next
'-------------------------------------

'-------------------------------------
'Unload the DCs of all slot pictures
For i = 0 To 4
    DeleteDC lPicDC(i)
Next
'-------------------------------------

DeleteDC lAniDCH    'unload the dc of the blue highlight handle
End Sub
'***************************************************
'***************************************************


'***************************************************
'Is the user sure that he/she wants to exit the*****
'current game?**************************************
'***************************************************
Private Sub cmdExit_Click()
If MsgBox("Are you sure you want to exit the current game?", _
            vbYesNo, "Exit Game") = vbYes Then
    Call Form_Terminate
    Unload Me
    frmStart.Show
End If
End Sub
'***************************************************
'***************************************************


'***************************************************
'When mouse is over the help button, shows frmHelp**
'***************************************************
Private Sub cmdHelp_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
bHelpFrom = False   'the form was accessed from frmMain
frmHelp.Show        'show the form
End Sub
'***************************************************
'***************************************************


'***************************************************
'When mouse is over the knob of handle, it is*******
'highlighted in blue********************************
'***************************************************
Private Sub cmdStart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
picHdl.Picture = LoadPicture(App.Path & ANIF & "H.bmp", , , 2)
End Sub
'***************************************************
'***************************************************


'***************************************************
'When mouse is anywhere on the form, hide frmHelp,
'and revert from the highlighted handle knob
'as long as the slot machine hasn't been initiated
'***************************************************
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If bRollin = False Then
    picHdl.Picture = LoadPicture(App.Path & ANIF & "1.bmp")
End If

frmHelp.Hide

If bMouseDwn = True Then
    Me.Left = Me.Left - (sMouseXPos - x)
    Me.Top = Me.Top - (sMouseYPos - y)
End If
End Sub
'***************************************************
'***************************************************

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDwn = True
sMouseXPos = x
sMouseYPos = y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
bMouseDwn = False
End Sub

