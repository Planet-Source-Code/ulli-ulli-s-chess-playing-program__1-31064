VERSION 5.00
Object = "{5431BC50-A509-11D2-921B-9337960C1B8F}#11.0#0"; "MSGHOOK.OCX"
Begin VB.Form fChess 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Ulli's Chess Program"
   ClientHeight    =   9645
   ClientLeft      =   1395
   ClientTop       =   1770
   ClientWidth     =   11460
   ClipControls    =   0   'False
   FillStyle       =   0  'Ausgefüllt
   ForeColor       =   &H00C0C0C0&
   Icon            =   "chess.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "chess.frx":0152
   Moveable        =   0   'False
   ScaleHeight     =   643
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   764
   Begin VB.CheckBox ckWarn 
      BackColor       =   &H0000FFFF&
      DownPicture     =   "chess.frx":045C
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10485
      MaskColor       =   &H00FFFF00&
      MousePointer    =   1  'Pfeil
      Picture         =   "chess.frx":09AE
      Style           =   1  'Grafisch
      TabIndex        =   40
      ToolTipText     =   "Attacked Square Warning"
      Top             =   3075
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin MsgFilter.MsgHook mhFocus 
      Left            =   2385
      Top             =   1275
      _ExtentX        =   1138
      _ExtentY        =   1931
   End
   Begin VB.ListBox lstPromo 
      Appearance      =   0  '2D
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   930
      ItemData        =   "chess.frx":0F00
      Left            =   1350
      List            =   "chess.frx":0F10
      MousePointer    =   1  'Pfeil
      Style           =   1  'Kontrollkästchen
      TabIndex        =   37
      Top             =   1275
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Timer tmrElapsed 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3075
      Top             =   1275
   End
   Begin VB.Frame fr 
      Caption         =   "Time to think"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1725
      Index           =   2
      Left            =   9135
      TabIndex        =   28
      Top             =   3765
      Width           =   1875
      Begin VB.CommandButton btBreak 
         BackColor       =   &H008080FF&
         Caption         =   "Stop thinking"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   120
         MousePointer    =   1  'Pfeil
         Style           =   1  'Grafisch
         TabIndex        =   38
         Top             =   1065
         Width           =   1620
      End
      Begin VB.HScrollBar scrTimeToThink 
         Height          =   240
         LargeChange     =   5
         Left            =   120
         Max             =   51
         Min             =   1
         TabIndex        =   29
         Top             =   300
         Value           =   5
         Width           =   1620
      End
      Begin VB.Label lbTime 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fest Einfach
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "Approximate time to think"
         Top             =   660
         Width           =   1620
      End
   End
   Begin VB.CommandButton btEdit 
      Caption         =   "Edit Board"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   10065
      TabIndex        =   1
      ToolTipText     =   "Set up individual piece positions"
      Top             =   810
      Width           =   930
   End
   Begin VB.CheckBox ckView 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Show planned moves"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   9120
      TabIndex        =   8
      ToolTipText     =   "Show the planned moves"
      Top             =   5685
      Width           =   1875
   End
   Begin VB.PictureBox picEinst 
      BackColor       =   &H00E0E0E0&
      Height          =   2790
      Left            =   9135
      Picture         =   "chess.frx":0F31
      ScaleHeight     =   2730
      ScaleWidth      =   1800
      TabIndex        =   18
      ToolTipText     =   "Don't disturb me - I'm thinking"
      Top             =   5970
      Visible         =   0   'False
      Width           =   1860
      Begin VB.Label lb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "thinking..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   2
         Left            =   345
         TabIndex        =   19
         Top             =   2415
         Width           =   990
      End
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3525
      Top             =   1275
   End
   Begin VB.CommandButton btNewGame 
      Caption         =   "Standard Board"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   9135
      TabIndex        =   0
      ToolTipText     =   "Set up standard piece positions"
      Top             =   810
      Width           =   930
   End
   Begin VB.CommandButton btGo 
      BackColor       =   &H0080FF80&
      Caption         =   "Start Game"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   9150
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Start a game"
      Top             =   8940
      Width           =   1845
   End
   Begin VB.Frame fr 
      Caption         =   "First Move"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   930
      Index           =   1
      Left            =   9135
      TabIndex        =   5
      Top             =   2745
      Width           =   1860
      Begin VB.OptionButton opComp 
         Caption         =   "Computer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   195
         TabIndex        =   6
         ToolTipText     =   "Computer makes first move"
         Top             =   330
         Width           =   990
      End
      Begin VB.OptionButton opHuman 
         Caption         =   "Human"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   195
         TabIndex        =   7
         ToolTipText     =   "Human makes first move"
         Top             =   585
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         X1              =   1245
         X2              =   1245
         Y1              =   120
         Y2              =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   1230
         X2              =   1230
         Y1              =   105
         Y2              =   900
      End
   End
   Begin VB.Frame fr 
      Caption         =   "Mode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1200
      Index           =   0
      Left            =   9135
      TabIndex        =   2
      Top             =   1455
      Width           =   1860
      Begin VB.CheckBox ckReverse 
         Caption         =   "Flip Board"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   210
         TabIndex        =   33
         ToolTipText     =   "Board Display Mode  "
         Top             =   840
         Width           =   1020
      End
      Begin VB.OptionButton opPlayAlt 
         Caption         =   "Alternate Players"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   195
         TabIndex        =   4
         ToolTipText     =   "Alternate players make the moves"
         Top             =   585
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.OptionButton opPlaySelf 
         Caption         =   "Same Player"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   195
         TabIndex        =   3
         ToolTipText     =   "Same Player makes the moves for both sides"
         Top             =   330
         Width           =   1200
      End
   End
   Begin VB.ListBox lsPV 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   9150
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5970
      Width           =   1815
   End
   Begin VB.PictureBox pcSquare 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00002040&
      BorderStyle     =   0  'Kein
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1650
      Index           =   0
      Left            =   960
      Picture         =   "chess.frx":EAD3
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   31
      Top             =   6990
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox pcSquare 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H00204570&
      BorderStyle     =   0  'Kein
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1650
      Index           =   1
      Left            =   2610
      Picture         =   "chess.frx":179BD
      ScaleHeight     =   1650
      ScaleWidth      =   1650
      TabIndex        =   32
      Top             =   6990
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Image imgUMGEDV 
      BorderStyle     =   1  'Fest Einfach
      Height          =   765
      Left            =   15
      Picture         =   "chess.frx":208A7
      Top             =   8865
      Width           =   825
   End
   Begin VB.Label lbMate 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mate "
      BeginProperty Font 
         Name            =   "Amaze"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   5265
      TabIndex        =   39
      Top             =   4125
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lbCheck 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "Amaze"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   5265
      TabIndex        =   36
      Top             =   4575
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblClock 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   3420
      TabIndex        =   35
      ToolTipText     =   "Black elapsed time"
      Top             =   375
      Width           =   945
   End
   Begin VB.Label lblClock 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3420
      TabIndex        =   34
      ToolTipText     =   "White elapsed Time"
      Top             =   105
      Width           =   945
   End
   Begin VB.Image imWhiteQueen 
      Appearance      =   0  '2D
      Height          =   720
      Index           =   0
      Left            =   5760
      Picture         =   "chess.frx":2258D
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imBlackPawn 
      Appearance      =   0  '2D
      Height          =   720
      Index           =   0
      Left            =   5040
      Picture         =   "chess.frx":24257
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imBlackBishop 
      Appearance      =   0  '2D
      Height          =   720
      Index           =   0
      Left            =   2880
      Picture         =   "chess.frx":25F21
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imBlackKnight 
      Appearance      =   0  '2D
      Height          =   720
      Index           =   0
      Left            =   3600
      Picture         =   "chess.frx":27BEB
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imBlackRook 
      Appearance      =   0  '2D
      Height          =   720
      Index           =   0
      Left            =   4320
      Picture         =   "chess.frx":298B5
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imBlackQueen 
      Appearance      =   0  '2D
      Height          =   720
      Index           =   0
      Left            =   1425
      Picture         =   "chess.frx":2B57F
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imBlackKing 
      Appearance      =   0  '2D
      Height          =   720
      Left            =   2160
      Picture         =   "chess.frx":2D249
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imWhitePawn 
      Appearance      =   0  '2D
      Height          =   720
      Index           =   0
      Left            =   9360
      Picture         =   "chess.frx":2EF13
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imWhiteBishop 
      Appearance      =   0  '2D
      Height          =   720
      Index           =   0
      Left            =   7200
      Picture         =   "chess.frx":30BDD
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imWhiteKnight 
      Appearance      =   0  '2D
      Height          =   720
      Index           =   0
      Left            =   7920
      Picture         =   "chess.frx":328A7
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imWhiteRook 
      Appearance      =   0  '2D
      Height          =   720
      Index           =   0
      Left            =   8640
      Picture         =   "chess.frx":34571
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image imWhiteKing 
      Appearance      =   0  '2D
      Height          =   720
      Left            =   6480
      Picture         =   "chess.frx":3623B
      Top             =   9750
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Shape sh 
      BackColor       =   &H00000000&
      BorderColor     =   &H0000FFFF&
      Height          =   7710
      Index           =   3
      Left            =   945
      Top             =   945
      Width           =   7710
   End
   Begin VB.Label lblVers 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   165
      Left            =   705
      TabIndex        =   27
      Top             =   540
      Width           =   1155
   End
   Begin VB.Label lbPly 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5235
      TabIndex        =   26
      ToolTipText     =   "Maximum ply analysed"
      Top             =   8940
      Width           =   945
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Max Ply"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   4515
      TabIndex        =   25
      Top             =   8970
      Width           =   675
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cutoffs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   4575
      TabIndex        =   24
      Top             =   9285
      Width           =   615
   End
   Begin VB.Label lbCutoff 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5235
      TabIndex        =   23
      ToolTipText     =   "Number of Cutoffs"
      Top             =   9255
      Width           =   945
   End
   Begin VB.Label lbPosns 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7860
      TabIndex        =   22
      ToolTipText     =   "Number of analysed positions"
      Top             =   9255
      Width           =   945
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Posns visited"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   6675
      TabIndex        =   21
      Top             =   9285
      Width           =   1140
   End
   Begin VB.Label lbMsg 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Left            =   4665
      TabIndex        =   13
      ToolTipText     =   "Various Messages"
      Top             =   210
      Width           =   6330
   End
   Begin VB.Label lb 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   7335
      TabIndex        =   17
      Top             =   8970
      Width           =   480
   End
   Begin VB.Label lbScore 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7860
      TabIndex        =   16
      ToolTipText     =   "The best future value"
      Top             =   8940
      Width           =   945
   End
   Begin VB.Label lbYMM 
      Alignment       =   1  'Rechts
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   825
      TabIndex        =   15
      Top             =   9150
      Width           =   1290
   End
   Begin VB.Label lbMoves 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "New Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   360
      Left            =   2190
      TabIndex        =   14
      ToolTipText     =   "The Current Move"
      Top             =   9060
      Width           =   1650
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "to move"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   2550
      TabIndex        =   12
      Top             =   255
      Width           =   690
   End
   Begin VB.Label lbSide 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fest Einfach
      Height          =   300
      Left            =   2175
      TabIndex        =   11
      ToolTipText     =   "Who's turn is it?"
      Top             =   210
      Width           =   300
   End
   Begin VB.Label lb 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Chess"
      BeginProperty Font 
         Name            =   "Amaze"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   525
      Index           =   0
      Left            =   300
      TabIndex        =   10
      ToolTipText     =   "Ulli's Chess Program   © 2001 UMGEDV"
      Top             =   0
      Width           =   1185
   End
   Begin VB.Shape sh 
      BackColor       =   &H00000000&
      BorderColor     =   &H00404080&
      BorderWidth     =   9
      Height          =   7860
      Index           =   0
      Left            =   870
      Top             =   870
      Width           =   7860
   End
End
Attribute VB_Name = "fChess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ToDo##
'Quiescence search##
'Forward Pruning in Quiescence Search ##
'Interrupt Search on timeout if less than 90% done ##
'Transposition Table TT##
'Opening Book##
'AI still needs a lot of cooking until done##
'I think Stalemate recognition is wrong##
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jan21,2002 UMG
'
'Fixed Edit Board: Pawn on 1st or 8th rank is not possible
'                  General tidying up of Edit function, the function now
'                  has a certain amount of intelligence
'Fixed Captured King during Move Generation (MaxMaterial)
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Jan15,2002 UMG
'
'Modfied Castling Permission update
'Added Move Notation - XLatMove
'Added Edit Board
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dec20,2001 UMG
'
'Hurray it plays
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'In Human vs Human mode the rules are enforced but no "mate"
'is announced (there is no search and no board evaluation)
'Doubleclick on an empty square to see all legal destination squares.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' The board is laid out as follows:
'
'       G  G  G  G  G  G  G  G  G  G
'       G  G  G  G  G  G  G  G  G  G
'       G  A8 B8 C8 D8 E8 F8 G8 H8 G
'       G  A7 B7 C7 D7 E7 F7 g7 H7 G
'       G  A6 B6 C6 D6 E6 F6 G6 H6 G
'       G  A5 B5 C5 D5 E5 F5 G5 H5 G
'       G  A4 B4 C4 D4 E4 F4 G4 H4 G
'       G  A3 B3 C3 D3 E3 F3 G3 H3 G
'       G  A2 B2 C2 D2 E2 F2 G2 H2 G
'       G  A1 B1 C1 D1 E1 F1 G1 H1 G
'       G  G  G  G  G  G  G  G  G  G
'       G  G  G  G  G  G  G  G  G  G
'
' G are guard squares to prevent the pieces from moving / jumping off the board
'
' The board is then streched to one dimension (the lower left corner having index 0)
' because a lone index is faster than two indexes
'
' Thus the legal square indexes are as follows:

Private Enum SquareIndexes

    A8 = 91
    B8 = 92
    C8 = 93
    D8 = 94
    E8 = 95
    F8 = 96
    G8 = 97
    H8 = 98

    A7 = 81
    B7 = 82
    C7 = 83
    D7 = 84
    E7 = 85
    F7 = 86
    G7 = 87
    H7 = 88

    A6 = 71
    B6 = 72
    C6 = 73
    D6 = 74
    E6 = 75
    F6 = 76
    G6 = 77
    H6 = 78

    A5 = 61
    B5 = 62
    C5 = 63
    D5 = 64
    E5 = 65
    F5 = 66
    G5 = 67
    H5 = 68

    A4 = 51
    B4 = 52
    C4 = 53
    D4 = 54
    E4 = 55
    F4 = 56
    G4 = 57
    H4 = 58

    A3 = 41
    B3 = 42
    C3 = 43
    D3 = 44
    E3 = 45
    F3 = 46
    G3 = 47
    H3 = 48

    A2 = 31
    B2 = 32
    C2 = 33
    D2 = 34
    E2 = 35
    F2 = 36
    G2 = 37
    H2 = 38

    A1 = 21
    B1 = 22
    C1 = 23
    D1 = 24
    E1 = 25
    F1 = 26
    G1 = 27
    H1 = 28

End Enum

' The squares contain the following byte values:
'
'         msb                                lsb
' Bit-->   7    6    5    4    3    2    1    0
'         -------------  ---  ------------------
'         ------------------
'               |        -----------------------
'               |         |           |
'               |         |            ----> Piece List Index
'               |         |
'               |          ------> Piece Color 0 White
'               |                              1 Black
'               |
'                -------> Piece Type  0 Free
'                                     1 King
'                                     2 Queen
'                                     3 Rook
'                                     4 Bishop
'                                     5 Knight
'                                     6 Pawn
'                                     7 Off Board (Guard)
'
' Bits 0 to 4 are combined to yield the index for the the piece list, so
' the white pieces are listed in 0..15 and the black pieces are listed in 16..31
'
' Bits 4 to 7 are also combined to yield the piece type, so &H2x is the white king and
' &H3x is the black king, &H4x is a white queen and &H5x is a black queen and so on.
'
' The entries in the Piece List have the following byte structure:
'
'         msb                                lsb
' Bit-->   7    6    5    4    3    2    1    0
'         ---  ---------------------------------
'          |                   |
'          |                   |
'          |                    ----> Position on Board
'          |
'           ----> Life  1 living
'                       0 dead (captured)

' Thus there is a cross reference both ways.
'
' The Kings are always listed in Positions 0 (white) and 16 (black) of the Piece list
'
' There is a third list having the piece material values; the index positions
' correspond the piece list: ---->  Value   100 Pawn
'                                           325 Knight
'                                           350 Bishop
'                                           510 Rook
'                                           900 Queen
'                                             0 King (value is irrelevant since it is priceless)
'
Private Const WM_SETFOCUS           As Long = 7

Private Const PieceTypeMask         As Byte = 32 Or 64 Or 128
Private Const PieceImageMask        As Byte = 1 Or 2 Or 4 Or 8
Private Const PieceColorMask        As Byte = 16
Private Const PieceTypeColorMask    As Byte = PieceTypeMask Or PieceColorMask
Private Const PieceSquareMask       As Byte = 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64
Private Const PieceListIndexMask    As Byte = 1 Or 2 Or 4 Or 8 Or PieceColorMask
Private Const PieceLivingMask       As Byte = 128
Private Const WhiteCastleShortMask  As Byte = 1 'White can castle kingside
Private Const WhiteCastleLongMask   As Byte = 2 'White can castle queenside
Private Const BlackCastleShortMask  As Byte = 4 'Black can castle kingside
Private Const BlackCastleLongMask   As Byte = 8 'Black can castle queenside
Private Const CastleMask            As Byte = WhiteCastleShortMask Or WhiteCastleLongMask Or BlackCastleShortMask Or BlackCastleLongMask
Private Const WhiteEndgameMask      As Byte = 16 'White is in endgame
Private Const BlackEndgameMask      As Byte = 32 'Black is in endgame
Private Const PawnMovedMask         As Byte = 64 'This bit governs the pawn structure evaluation, on after pawn or king move
Private Const Free                  As Byte = 0 'used with PieceTypeMask
Private Const King                  As Byte = 1 * 32
Private Const Queen                 As Byte = 2 * 32
Private Const Rook                  As Byte = 3 * 32
Private Const Bishop                As Byte = 4 * 32
Private Const Knight                As Byte = 5 * 32
Private Const Pawn                  As Byte = 6 * 32
Private Const Forbidden             As Byte = 7 * 32

Private Const White                 As Byte = 0
Private Const Black                 As Byte = PieceColorMask

Private Const WhiteKing             As Byte = King Or White 'used with PieceTypeColorMask
Private Const WhiteQueen            As Byte = Queen Or White
Private Const WhiteRook             As Byte = Rook Or White
Private Const WhiteBishop           As Byte = Bishop Or White
Private Const WhiteKnight           As Byte = Knight Or White
Private Const WhitePawn             As Byte = Pawn Or White
Private Const BlackKing             As Byte = King Or Black
Private Const BlackQueen            As Byte = Queen Or Black
Private Const BlackRook             As Byte = Rook Or Black
Private Const BlackBishop           As Byte = Bishop Or Black
Private Const BlackKnight           As Byte = Knight Or Black
Private Const BlackPawn             As Byte = Pawn Or Black
Private Const QueenValue            As Integer = 900
Private Const RookValue             As Integer = 510
Private Const BishopValue           As Integer = 350
Private Const KnightValue           As Integer = 325
Private Const PawnValue             As Integer = 100
Private Const EnoughAdvantage       As Integer = PawnValue + 1
Private Const MaxMaterial           As Long = 15 * QueenValue
Private Const TTSize                As Long = 2 ^ 21 - 1

Private Const WhitesTurnColor       As Long = vbWhite
Private Const BlacksTurnColor       As Long = vbBlack
Private Const NeutralColor          As Long = &HFFC0FF
Private Const Draw                  As String = "Stalemate, the game is a draw"

Private Const Infinity              As Integer = 30000
Private Const PlyLimit              As Long = 60

Private Enum Compass
    North = 10
    West = -1
    South = -10
    East = 1
    NorthNorth = North + North
    EastEast = East + East
    SouthSouth = South + South
    WestWest = West + West
    WestWestWest = WestWest + West
    NorthEast = North + East
    NorthWest = North + West
    SouthEast = South + East
    SouthWest = South + West
End Enum

Private Enum Reasons 'for game end
    InProgress = 0
    ByLoss = 1
    ByResign = 2
    ByDraw = 3
    ByWin = 4
    ByUser = 5
End Enum

'move and jump distances
Private QueenDistances              As Variant 'queen (rook and bishop)
Private KnightDistances             As Variant 'knight

'for black board evaluation
Private Mirror                      As Variant

'translate square to rank and file
Private Rank                        As Variant
Private File                        As Variant

'knowledge arrays
'''''''''''''''''
'square mobility values
Private Mobility                    As Variant

'piece on square values
Private NormalPawnOnSquare          As Variant
Private EndgamePawnOnSquare         As Variant
Private KnightOnSquare              As Variant
Private BishopOnSquare              As Variant
Private RookOnSquare                As Variant
Private NormalKingOnSquare          As Variant
Private EndgameKingOnSquare         As Variant

Private CastleBits                  As Variant

Private Type Board
    Squares(0 To 119)               As Byte 'squares and guard squares
    Pieces(0 To 31)                 As Byte 'white king in 0; black king in 16
    SideToMove                      As Byte 'corresponds to PieceColor Mask
    QuietMoveListFrom               As Integer 'indexes for generate moves and get moves
    QuietMoveListTo                 As Integer
    CaptureMoveListFrom             As Integer
    CaptureMoveListTo               As Integer
    BestCapture                     As Integer 'index into capturemovelist: bestcapture
    MoveCategory                    As Byte 'type of move to analyse next (for move ordering)
    CurrMoveIndex                   As Integer 'index for the current move in the move lists
    EnPassant                       As Byte 'has last pawn dest square if pawn moved two squares, and zero else
    MiscBits                        As Byte 'has all four castling rights; the 2 endgame states and the pawn moved bit - used to trigger pawn eval
    TTIdx                           As Long 'Index into Transposition Table
    TTCnf                           As Long 'Transposition Table Entry Confirmation
    Material(0 To 1)                As Integer '0 white; 1 black
    PawnStrucValue                  As Integer 'latest computed pawn structure value
End Type
'since the board is copied into tempboard very frequently it is kept
'as short as possible 182 bytes

Private Board As Board

Private X, Y, i, j, k, l
Private EndSplash                   As Date
Private LastX                       As Single
Private LastY                       As Single
Private QuietMoveListIx
Private CaptureMoveListIx
Private WhiteKingCount
Private BlackKingCount
Private WhiteQueensUsed
Private WhiteRooksUsed
Private WhiteBishopsUsed
Private WhiteKnightsUsed
Private WhitePawnsUsed
Private BlackQueensUsed
Private BlackRooksUsed
Private BlackBishopsUsed
Private BlackKnightsUsed
Private BlackPawnsUsed
Private InitialAlpha
Private InitialBeta
Private Iteration
Private ClickButton                 As Integer
Private MaxPlySearched
Private CutOffs
Private PosnsVisited
Private Score
Private Pivot
Private MoveCount
Private TimeLimit                   As Single
Private GameStart                   As Date
Private WhiteTime                   As Date
Private BlackTime                   As Date
Private Result
Private QuResult
Private Alpha
Private Beta
Private GameEnds                    As Reasons
Private Editing                     As Boolean
Private PawnStructureValue 'result of the pawn structure evaluation
Private Seed 'used for random board painting
Private InPrinVar                   As Boolean 'true while search is in principal variation
Private Stalemate                   As Boolean
Private Recur                       As Boolean 'prevent recursion on castling moves
Private BreakRequested              As Boolean 'Stop search
Private SaverActive                 As Long 'the previous screensaver state
Private Mode                        As String 'game mode
Private Hilited                     As String 'the hilited squares
Private FormClicked                 As Boolean
Private ClickEnabled                As Boolean
Private KeyDown                     As Boolean 'used in individual board setup
Private HumanMove                   As String 'human move in decimal notation
Private QuietMoves(0 To 60 * PlyLimit) As String 'max 60 avg QuietMoves per ply
Private CaptureMoves(0 To 20 * PlyLimit) As String 'max 20 avg CaptureMoves per ply
Private PieceValues(0 To 31)        As Integer 'white king in 0; black king in 16
Private CurrPV(0 To PlyLimit, 0 To PlyLimit) As String
Private PrevPV(-2 To PlyLimit)      As String
Private Killer1(0 To PlyLimit)      As String
Private Killer2(0 To PlyLimit)      As String

'piece images with permission from RJSoft of West Tennessee
Private imPieces(0 To 31)           As Image

Private Type IndConf
    Index                           As Long
    Confirm                         As Long
End Type

Private HashValues(0 To Forbidden * 8 Or H8) As IndConf 'the index is ((piece shiftleft 3) or square)

Private Type TTEntry
    Confirm                         As Long
    Confidence                      As Integer
    Value                           As Integer
End Type

Private TT(0 To TTSize)             As TTEntry 'transposition table

'wait a little
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'painting the board
Private Declare Sub BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long)
Private Const SRCCOPY As Long = &HCC0020

'pretend to be a screensaver
Private Declare Sub SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long)
Private Const SPI_GETSCREENSAVEACTIVE As Long = 16
Private Const SPI_SETSCREENSAVEACTIVE As Long = 17

Private Sub btBreak_Click()

    BreakRequested = True

End Sub

Private Sub btEdit_Click()

  'individual board setup

  Dim WhitePawnsConsumed    As Long
  Dim BlackPawnsConsumed    As Long
  Dim SelNum                As Long

    Editing = True
    GameEnds = ByUser 'user interupted game
    ResetClocks
    ClearLabels
    lbMsg = ""
    lbYMM = ""
    btGo.Enabled = False
    btEdit.Enabled = False
    btNewGame.Enabled = False
    fr(0).Enabled = False
    fr(1).Enabled = False
    lbCheck.Visible = False
    lbMate.Visible = False
    For i = 1 To Len(Hilited)
        UnHiliteSquare Asc(Mid$(Hilited, i, 1))
    Next i
    Hilited = ""
    lbCutoff = ""
    lbMoves = ""
    lbPly = ""
    lbScore = ""
    lbSide.BackColor = BackColor
    PrevPV(1) = "" 'invalidate PV
    With Board
        For i = 0 To 19
            .Squares(i) = Forbidden
        Next i
        For i = 100 To 119
            .Squares(i) = Forbidden
        Next i
        For i = 20 To 90 Step 10
            For j = 1 To 8
                .Squares(i + j) = Free
            Next j
            .Squares(i) = Forbidden
            .Squares(i + 9) = Forbidden
        Next i
        For i = 0 To 31
            .Pieces(i) = Free
            Set imPieces(i) = Nothing
        Next i
        imWhiteKing.Visible = False
        imBlackKing.Visible = False
        For i = 0 To 9
            imWhiteQueen(i).Visible = False
            imWhiteRook(i).Visible = False
            imWhiteBishop(i).Visible = False
            imWhiteKnight(i).Visible = False
            imWhitePawn(i).Visible = False
            imBlackQueen(i).Visible = False
            imBlackRook(i).Visible = False
            imBlackBishop(i).Visible = False
            imBlackKnight(i).Visible = False
            imBlackPawn(i).Visible = False
        Next i
        WhiteKingCount = 0
        BlackKingCount = 0
        WhiteQueensUsed = 0
        WhiteRooksUsed = 0
        WhiteBishopsUsed = 0
        WhiteKnightsUsed = 0
        WhitePawnsUsed = 0
        BlackQueensUsed = 0
        BlackRooksUsed = 0
        BlackBishopsUsed = 0
        BlackKnightsUsed = 0
        BlackPawnsUsed = 0
        ClickEnabled = True
        lbMsg = "White - leftclick, Black - rightclick; spacebar when done..."
        HumanMove = ""
        Do
            Do
                DoEvents
            Loop Until FormClicked Or KeyDown
            If FormClicked Then
                ClickEnabled = False
                FormClicked = False
                i = Asc(HumanMove + Chr$(0))
                HumanMove = ""
                If i Then
                    X = TLX(i) + 8
                    Y = TLY(i) + 8
                    If ClickButton = vbRightButton Then 'Black
                        Select Case .Squares(i) And PieceTypeColorMask
                          Case WhiteKing
                            WhiteKingCount = 0
                            imWhiteKing.Visible = False
                            .Squares(i) = Free
                          Case WhiteQueen
                            WhiteQueensUsed = WhiteQueensUsed - 1
                            imWhiteQueen(WhiteQueensUsed).Visible = False
                            .Squares(i) = Free
                          Case WhiteRook
                            WhiteRooksUsed = WhiteRooksUsed + 1
                            imWhiteRook(WhiteRooksUsed).Visible = False
                            .Squares(i) = Free
                          Case WhiteBishop
                            WhiteBishopsUsed = WhiteBishopsUsed - 1
                            imWhiteBishop(WhiteBishopsUsed).Visible = False
                            .Squares(i) = Free
                          Case WhiteKnight
                            WhiteKnightsUsed = WhiteKnightsUsed - 1
                            imWhiteKnight(WhiteKnightsUsed).Visible = False
                            .Squares(i) = Free
                          Case WhitePawn
                            WhitePawnsUsed = WhitePawnsUsed - 1
                            imWhitePawn(WhitePawnsUsed).Visible = False
                            .Squares(i) = Free
                        End Select

                        Select Case .Squares(i) And PieceTypeColorMask
                          Case Free
                            SelNum = 1
                            GoSub WhichBlackPiece
                          Case BlackKing
                            imBlackKing.Visible = False
                            BlackKingCount = 0
                            SelNum = 2
                            GoSub WhichBlackPiece
                          Case BlackPawn
                            BlackPawnsUsed = BlackPawnsUsed - 1
                            imBlackPawn(.Squares(i) And PieceImageMask).Visible = False
                            SelNum = 3
                            GoSub WhichBlackPiece
                          Case BlackKnight
                            BlackKnightsUsed = BlackKnightsUsed - 1
                            imBlackKnight(.Squares(i) And PieceImageMask).Visible = False
                            SelNum = 4
                            GoSub WhichBlackPiece
                          Case BlackBishop
                            BlackBishopsUsed = BlackBishopsUsed - 1
                            imBlackBishop(.Squares(i) And PieceImageMask).Visible = False
                            SelNum = 5
                            GoSub WhichBlackPiece
                          Case BlackRook
                            BlackRooksUsed = BlackRooksUsed - 1
                            imBlackRook(.Squares(i) And PieceImageMask).Visible = False
                            SelNum = 6
                            GoSub WhichBlackPiece
                          Case BlackQueen
                            BlackQueensUsed = BlackQueensUsed - 1
                            imBlackQueen(.Squares(i) And PieceImageMask).Visible = False
                            .Squares(i) = Free
                        End Select

                      Else 'NOT CLICKBUTTON...
                        Select Case .Squares(i) And PieceTypeColorMask
                          Case BlackKing
                            BlackKingCount = 0
                            imBlackKing.Visible = False
                            .Squares(i) = Free
                          Case BlackQueen
                            BlackQueensUsed = BlackQueensUsed - 1
                            imBlackQueen(.Squares(i) And PieceImageMask).Visible = False
                            .Squares(i) = Free
                          Case BlackRook
                            BlackRooksUsed = BlackRooksUsed - 1
                            imBlackRook(.Squares(i) And PieceImageMask).Visible = False
                            .Squares(i) = Free
                          Case BlackBishop
                            BlackBishopsUsed = BlackBishopsUsed - 1
                            imBlackBishop(.Squares(i) And PieceImageMask).Visible = False
                            .Squares(i) = Free
                          Case BlackKnight
                            BlackKnightsUsed = BlackKnightsUsed - 1
                            imBlackKnight(.Squares(i) And PieceImageMask).Visible = False
                            .Squares(i) = Free
                          Case BlackPawn
                            BlackPawnsUsed = BlackPawnsUsed - 1
                            imBlackPawn(.Squares(i) And PieceImageMask).Visible = False
                            .Squares(i) = Free
                        End Select
    
                        Select Case .Squares(i) And PieceTypeColorMask
                          Case Free
                            SelNum = 1
                            GoSub WhichWhitePiece
                          Case WhiteKing
                            imWhiteKing.Visible = False
                            WhiteKingCount = 0
                            SelNum = 2
                            GoSub WhichWhitePiece
                          Case WhitePawn
                            WhitePawnsUsed = WhitePawnsUsed - 1
                            imWhitePawn(.Squares(i) And PieceImageMask).Visible = False
                            SelNum = 3
                            GoSub WhichWhitePiece
                          Case WhiteKnight
                            WhiteKnightsUsed = WhiteKnightsUsed - 1
                            imWhiteKnight(.Squares(i) And PieceImageMask).Visible = False
                            SelNum = 4
                            GoSub WhichWhitePiece
                          Case WhiteBishop
                            WhiteBishopsUsed = WhiteBishopsUsed - 1
                            imWhiteBishop(.Squares(i) And PieceImageMask).Visible = False
                            SelNum = 5
                            GoSub WhichWhitePiece
                          Case WhiteRook
                            WhiteRooksUsed = WhiteRooksUsed - 1
                            imWhiteRook(.Squares(i) And PieceImageMask).Visible = False
                            SelNum = 6
                            GoSub WhichWhitePiece
                          Case WhiteQueen
                            WhiteQueensUsed = WhiteQueensUsed - 1
                            imWhiteQueen(.Squares(i) And PieceImageMask).Visible = False
                            .Squares(i) = Free
                        End Select

                    End If
                End If 'HumanMove valid
                DoEvents
                Sleep 111
                lbMoves = ""
                ClickEnabled = True
            End If 'FormClicked
        Loop Until KeyDown
        lbMsg = ""
        .Material(0) = 0
        .Material(1) = 0
        j = 1
        k = 1
        For i = A1 To H8 'all squares
            Select Case .Squares(i) And PieceTypeColorMask
              Case BlackKing
                .Pieces(Black) = i Or PieceLivingMask
                .Squares(i) = .Squares(i) And PieceTypeColorMask
                Set imPieces(Black) = imBlackKing
              Case BlackQueen
                Set imPieces(Black + k) = imBlackQueen(.Squares(i) And PieceImageMask)
                .Material(1) = .Material(1) + QueenValue
                PieceValues(Black + k) = QueenValue
              Case BlackRook
                Set imPieces(Black + k) = imBlackRook(.Squares(i) And PieceImageMask)
                .Material(1) = .Material(1) + RookValue
                PieceValues(Black + k) = RookValue
              Case BlackBishop
                Set imPieces(Black + k) = imBlackBishop(.Squares(i) And PieceImageMask)
                .Material(1) = .Material(1) + BishopValue
                PieceValues(Black + k) = BishopValue
              Case BlackKnight
                Set imPieces(Black + k) = imBlackKnight(.Squares(i) And PieceImageMask)
                .Material(1) = .Material(1) + KnightValue
                PieceValues(Black + k) = KnightValue
              Case BlackPawn
                Set imPieces(Black + k) = imBlackPawn(.Squares(i) And PieceImageMask)
                .Material(1) = .Material(1) + PawnValue
                PieceValues(Black + k) = PawnValue
              Case WhiteKing
                .Pieces(White) = i Or PieceLivingMask
                .Squares(i) = .Squares(i) And PieceTypeColorMask
                Set imPieces(White) = imWhiteKing
              Case WhiteQueen
                Set imPieces(White + j) = imWhiteQueen(.Squares(i) And PieceImageMask)
                .Material(0) = .Material(0) + QueenValue
                PieceValues(White + k) = QueenValue
              Case WhiteRook
                Set imPieces(White + j) = imWhiteRook(.Squares(i) And PieceImageMask)
                .Material(0) = .Material(0) + RookValue
                PieceValues(White + k) = RookValue
              Case WhiteBishop
                Set imPieces(White + j) = imWhiteBishop(.Squares(i) And PieceImageMask)
                .Material(0) = .Material(0) + BishopValue
                PieceValues(White + k) = BishopValue
              Case WhiteKnight
                Set imPieces(White + j) = imWhiteKnight(.Squares(i) And PieceImageMask)
                .Material(0) = .Material(0) + KnightValue
                PieceValues(White + k) = KnightValue
              Case WhitePawn
                Set imPieces(White + j) = imWhitePawn(.Squares(i) And PieceImageMask)
                .Material(0) = .Material(0) + PawnValue
                PieceValues(White + k) = PawnValue
            End Select
            If (.Squares(i) And PieceTypeMask) <> King Then
                Select Case .Squares(i) And PieceColorMask
                  Case Black
                    .Pieces(Black + k) = i Or PieceLivingMask
                    .Squares(i) = (.Squares(i) And PieceTypeColorMask) Or k
                    k = k + 1
                  Case White
                    If .Squares(i) <> Free And .Squares(i) <> Forbidden Then 'forbidden squares look like they have a white piece
                        .Pieces(White + j) = i Or PieceLivingMask
                        .Squares(i) = (.Squares(i) And PieceTypeColorMask) Or j
                        j = j + 1
                    End If
                End Select
            End If
        Next i
        Board.SideToMove = White
        j = 0
        With fRights
            .txEPB.Enabled = (Board.SideToMove = Black)
            .txEPW.Enabled = (Board.SideToMove = White)
            .txEPB = ""
            .txEPW = ""
            .ckCastleKB = vbUnchecked
            .ckCastleKW = vbUnchecked
            .ckCastleQB = vbUnchecked
            .ckCastleQW = vbUnchecked
            If Board.SideToMove = White Then
                .opWhite = True
              Else 'NOT BOARD.SIDETOMOVE...
                .opBlack = True
            End If
            If (Board.Squares(E1) And PieceTypeColorMask) = WhiteKing Then
                .ckCastleQW.Enabled = ((Board.Squares(A1) And PieceTypeColorMask) = WhiteRook)
                .ckCastleKW.Enabled = ((Board.Squares(H1) And PieceTypeColorMask) = WhiteRook)
              Else 'NOT (BOARD.SQUARES(E1)...
                .ckCastleQW.Enabled = False
                .ckCastleKW.Enabled = False
            End If
            If (Board.Squares(E8) And PieceTypeColorMask) = BlackKing Then
                .ckCastleQB.Enabled = ((Board.Squares(A8) And PieceTypeColorMask) = BlackRook)
                .ckCastleKB.Enabled = ((Board.Squares(H8) And PieceTypeColorMask) = BlackRook)
              Else 'NOT (BOARD.SQUARES(E8)...
                .ckCastleQB.Enabled = False
                .ckCastleKB.Enabled = False
            End If
            Do
                .Show vbModal, Me
                i = ((.ckCastleKW = vbChecked) And WhiteCastleShortMask) Or ((.ckCastleKB = vbChecked) And BlackCastleShortMask) Or ((.ckCastleQW = vbChecked) And WhiteCastleLongMask) Or ((.ckCastleQB = vbChecked) And BlackCastleLongMask)
                If .Square Then
                    If Board.SideToMove = White Then
                        If (Board.Squares(.Square + South) And PieceTypeColorMask) = BlackPawn Then
                            Board.EnPassant = .Square + South
                            j = 1
                          Else 'NOT (BOARD.SQUARES(.SQUARE...
                            .SetFocusOnText = True
                        End If
                      Else 'NOT BOARD.SIDETOMOVE...
                        If (Board.Squares(.Square + North) And PieceTypeColorMask) = WhitePawn Then
                            Board.EnPassant = .Square + North
                            j = 1
                          Else 'NOT (BOARD.SQUARES(.SQUARE...
                            .SetFocusOnText = True
                        End If
                    End If
                  Else '.SQUARE = FALSE
                    j = 1
                End If
            Loop Until j
            Board.SideToMove = IIf(.opBlack, Black, White)
            lbSide.BackColor = IIf(Board.SideToMove = White, WhitesTurnColor, BlacksTurnColor)
        End With 'FRIGHTS
        Unload fRights
        opComp_Click
        .MiscBits = i Or PawnMovedMask 'to trigger pawn evaluation also
        If IsAttacked(Board, .Pieces(.SideToMove Xor Black) And PieceSquareMask, .SideToMove) Then
            lbMsg = "Illegal Position, " & IIf(.SideToMove = White, "White", "Black") & " to move and the " & IIf(.SideToMove = White, "black", "white") & " king is attacked."
          Else 'NOT ISATTACKED(BOARD,...
            btGo.Enabled = True
        End If
    End With 'BOARD
    MousePointer = vbHourglass
    CreateHash
    MousePointer = vbNormal
    btEdit.Enabled = True
    btNewGame.Enabled = True
    fr(0).Enabled = True
    fr(1).Enabled = True
    ClickEnabled = False
    KeyDown = False
    lbMoves.ForeColor = NeutralColor
    lbMoves = "New Game"
    Editing = False

Exit Sub

WhichBlackPiece:

    If (BlackKnightsUsed + BlackBishopsUsed + BlackRooksUsed + BlackQueensUsed + BlackPawnsUsed) < 16 Then
        BlackPawnsConsumed = BlackPawnsUsed
        If BlackKnightsUsed > 2 Then
            BlackPawnsConsumed = BlackPawnsConsumed + BlackKnightsUsed - 2
        End If
        If BlackBishopsUsed > 2 Then
            BlackPawnsConsumed = BlackPawnsConsumed + BlackBishopsUsed - 2
        End If
        If BlackRooksUsed > 2 Then
            BlackPawnsConsumed = BlackPawnsConsumed + BlackRooksUsed - 2
        End If
        If BlackQueensUsed > 1 Then
            BlackPawnsConsumed = BlackPawnsConsumed + BlackQueensUsed - 1
        End If

        With Board
            'on entry i has the square
            .Squares(i) = Free
            Select Case SelNum
              Case 1 'free square
                If BlackKingCount Then
                    SelNum = 2
                    GoSub WhichBlackPiece
                  Else 'BLACKKINGCOUNT = FALSE
                    .Squares(i) = BlackKing
                    imBlackKing.Move X, Y
                    imBlackKing.Visible = True
                    BlackKingCount = 1
                End If
              Case 2 'there was a king on this square or we have a king aready
                If i < A2 Or i > H7 Or BlackPawnsConsumed = 8 Then
                    SelNum = 3
                    GoSub WhichBlackPiece
                  ElseIf BlackPawnsConsumed < 8 Then 'NOT I...
                    .Squares(i) = BlackPawn Or BlackPawnsUsed
                    For j = 0 To 7
                        If imBlackPawn(j).Visible = False Then
                            Exit For '>---> Next
                        End If
                    Next j
                    imBlackPawn(j).Move X, Y
                    imBlackPawn(j).Visible = True
                    BlackPawnsUsed = BlackPawnsUsed + 1
                End If
              Case 3 'there was a pawn on this square or we have 8 pawns or the square cannot have a pawn
                If BlackKnightsUsed < 2 Or BlackPawnsConsumed < 8 Then
                    .Squares(i) = BlackKnight Or BlackKnightsUsed
                    For j = 0 To 9
                        If imBlackKnight(j).Visible = False Then
                            Exit For '>---> Next
                        End If
                    Next j
                    imBlackKnight(j).Move X, Y
                    imBlackKnight(j).Visible = True
                    BlackKnightsUsed = BlackKnightsUsed + 1
                  Else 'NOT BLACKKNIGHTSUSED...
                    SelNum = 4
                    GoSub WhichBlackPiece
                End If
              Case 4
                If BlackBishopsUsed < 2 Or BlackPawnsConsumed < 8 Then
                    .Squares(i) = BlackBishop Or BlackBishopsUsed
                    For j = 0 To 9
                        If imBlackBishop(j).Visible = False Then
                            Exit For '>---> Next
                        End If
                    Next j
                    imBlackBishop(j).Move X, Y
                    imBlackBishop(j).Visible = True
                    BlackBishopsUsed = BlackBishopsUsed + 1
                  Else 'NOT BLACKBISHOPSUSED...
                    SelNum = 5
                    GoSub WhichBlackPiece
                End If
              Case 5
                If BlackRooksUsed < 2 Or BlackPawnsConsumed < 8 Then
                    .Squares(i) = BlackRook Or BlackRooksUsed
                    For j = 0 To 9
                        If imBlackRook(j).Visible = False Then
                            Exit For '>---> Next
                        End If
                    Next j
                    imBlackRook(j).Move X, Y
                    imBlackRook(j).Visible = True
                    BlackRooksUsed = BlackRooksUsed + 1
                  Else 'NOT BLACKROOKSUSED...
                    SelNum = 6
                    GoSub WhichBlackPiece
                End If
              Case 6
                If BlackQueensUsed < 1 Or BlackPawnsConsumed < 8 Then
                    .Squares(i) = BlackQueen Or BlackQueensUsed
                    For j = 0 To 8
                        If imBlackQueen(j).Visible = False Then
                            Exit For '>---> Next
                        End If
                    Next j
                    imBlackQueen(j).Move X, Y
                    imBlackQueen(j).Visible = True
                    BlackQueensUsed = BlackQueensUsed + 1
                End If
            End Select
        End With 'BOARD
    End If
    Return

WhichWhitePiece:

    If (WhiteKnightsUsed + WhiteBishopsUsed + WhiteRooksUsed + WhiteQueensUsed + WhitePawnsUsed) < 16 Then
        WhitePawnsConsumed = WhitePawnsUsed
        If WhiteKnightsUsed > 2 Then
            WhitePawnsConsumed = WhitePawnsConsumed + WhiteKnightsUsed - 2
        End If
        If WhiteBishopsUsed > 2 Then
            WhitePawnsConsumed = WhitePawnsConsumed + WhiteBishopsUsed - 2
        End If
        If WhiteRooksUsed > 2 Then
            WhitePawnsConsumed = WhitePawnsConsumed + WhiteRooksUsed - 2
        End If
        If WhiteQueensUsed > 1 Then
            WhitePawnsConsumed = WhitePawnsConsumed + WhiteQueensUsed - 1
        End If

        With Board
            'on entry i has the square
            .Squares(i) = Free
            Select Case SelNum
              Case 1 'free square
                If WhiteKingCount Then
                    SelNum = 2
                    GoSub WhichWhitePiece
                  Else 'WHITEKINGCOUNT = FALSE
                    .Squares(i) = WhiteKing
                    imWhiteKing.Move X, Y
                    imWhiteKing.Visible = True
                    WhiteKingCount = 1
                End If
              Case 2 'there was a king on this square or we have a king aready
                If i < A2 Or i > H7 Or WhitePawnsConsumed = 8 Then
                    SelNum = 3
                    GoSub WhichWhitePiece
                  ElseIf WhitePawnsConsumed < 8 Then 'NOT I...
                    .Squares(i) = WhitePawn Or WhitePawnsUsed
                    For j = 0 To 7
                        If imWhitePawn(j).Visible = False Then
                            Exit For '>---> Next
                        End If
                    Next j
                    imWhitePawn(j).Move X, Y
                    imWhitePawn(j).Visible = True
                    WhitePawnsUsed = WhitePawnsUsed + 1
                End If
              Case 3 'there was a pawn in this square or we have 8 pawns or the square cannot have a pawn
                If WhiteKnightsUsed < 2 Or WhitePawnsConsumed < 8 Then
                    .Squares(i) = WhiteKnight Or WhiteKnightsUsed
                    For j = 0 To 9
                        If imWhiteKnight(j).Visible = False Then
                            Exit For '>---> Next
                        End If
                    Next j
                    imWhiteKnight(j).Move X, Y
                    imWhiteKnight(j).Visible = True
                    WhiteKnightsUsed = WhiteKnightsUsed + 1
                  Else 'NOT WHITEKNIGHTSUSED...
                    SelNum = 4
                    GoSub WhichWhitePiece
                End If
              Case 4
                If WhiteBishopsUsed < 2 Or WhitePawnsConsumed < 8 Then
                    .Squares(i) = WhiteBishop Or WhiteBishopsUsed
                    For j = 0 To 9
                        If imWhiteBishop(j).Visible = False Then
                            Exit For '>---> Next
                        End If
                    Next j
                    imWhiteBishop(j).Move X, Y
                    imWhiteBishop(j).Visible = True
                    WhiteBishopsUsed = WhiteBishopsUsed + 1
                  Else 'NOT WHITEBISHOPSUSED...
                    SelNum = 5
                    GoSub WhichWhitePiece
                End If
              Case 5
                If WhiteRooksUsed < 2 Or WhitePawnsConsumed < 8 Then
                    .Squares(i) = WhiteRook Or WhiteRooksUsed
                    For j = 0 To 9
                        If imWhiteRook(j).Visible = False Then
                            Exit For '>---> Next
                        End If
                    Next j
                    imWhiteRook(j).Move X, Y
                    imWhiteRook(j).Visible = True
                    WhiteRooksUsed = WhiteRooksUsed + 1
                  Else 'NOT WHITEROOKSUSED...
                    SelNum = 6
                    GoSub WhichWhitePiece
                End If
              Case 6
                If WhiteQueensUsed < 1 Or WhitePawnsConsumed < 8 Then
                    .Squares(i) = WhiteQueen Or WhiteQueensUsed
                    For j = 0 To 8
                        If imWhiteQueen(j).Visible = False Then
                            Exit For '>---> Next
                        End If
                    Next j
                    imWhiteQueen(j).Move X, Y
                    imWhiteQueen(j).Visible = True
                    WhiteQueensUsed = WhiteQueensUsed + 1
                End If
            End Select
        End With 'BOARD
    End If
    Return

End Sub

Private Sub btGo_Click()

  'start a game

    fr(0).Enabled = False
    fr(1).Enabled = False
    btGo.Enabled = False
    btEdit.Enabled = False
    'kill PV
    For i = 0 To PlyLimit
        CurrPV(0, i) = ""
    Next i
    ResetClocks
    GameStart = Now
    lbMoves = ""
    tmrElapsed.Enabled = True

    PlayGame

    tmrElapsed.Enabled = False
    btNewGame.Enabled = True
    btEdit.Enabled = True
    picEinst.Visible = False
    MousePointer = vbNormal
    fr(0).Enabled = True
    fr(1).Enabled = True

End Sub

Private Sub btNewGame_Click()

  'set up standard board

    GameEnds = ByUser 'user interrupted game
    For i = 1 To Len(Hilited)
        UnHiliteSquare Asc(Mid$(Hilited, i, 1))
    Next i
    Hilited = ""
    DoEvents
    lbMsg = ""
    For i = 1 To 9
        Unload imWhiteQueen(i)
        Unload imBlackQueen(i)
        Unload imWhiteBishop(i)
        Unload imBlackBishop(i)
        Unload imWhiteKnight(i)
        Unload imBlackKnight(i)
        Unload imWhiteRook(i)
        Unload imBlackRook(i)
        Unload imWhitePawn(i)
        Unload imBlackPawn(i)
    Next i
    btGo.Enabled = True
    btEdit.Enabled = True
    lbCheck.Visible = False
    lbMate.Visible = False
    PrevPV(1) = "" 'invalidate PV
    Form_Load
    ClearLabels
    fr(0).Enabled = True
    fr(1).Enabled = True
    lbMoves.ForeColor = NeutralColor
    lbMoves = "New Game"

End Sub

Private Function CastlingIsLegal(Board As Board, ByVal Which As Long) As Boolean

  'check legal castling

    If Board.SideToMove = White Then
        If Not IsAttacked(Board, E1, Black) Then
            If Which = 5 Then 'queenside castling
                CastlingIsLegal = Not IsAttacked(Board, D1, Black)
              Else 'NOT WHICH...
                CastlingIsLegal = Not IsAttacked(Board, F1, Black)
            End If
        End If
      Else 'NOT BOARD.SIDETOMOVE...
        If Not IsAttacked(Board, E8, White) Then
            If Which = 5 Then 'queenside castling
                CastlingIsLegal = Not IsAttacked(Board, D8, White)
              Else 'NOT WHICH...
                CastlingIsLegal = Not IsAttacked(Board, F8, White)
            End If
        End If
    End If
    'the square which the king will occupy is not checked here; this is deferred
    'until the next search level (which will capture any illegally placed king)

End Function

Private Sub ckReverse_Click()

  'reverse board

    For i = 0 To 31
        If Not imPieces(i) Is Nothing Then
            imPieces(i).Visible = False
            MovePieceImage i, Board.Pieces(i) And PieceSquareMask
            imPieces(i).Visible = Board.Pieces(i) And PieceLivingMask
        End If
    Next i

End Sub

Private Sub ckView_Click()

  'show planned moves (the pricipal variation)

    lsPV.Clear
    If ckView Then
        lsPV.ToolTipText = "Here are the planned Moves"
      Else 'CKVIEW = FALSE
        lsPV.ToolTipText = "No Moves showing"
    End If

End Sub

Private Sub ClearLabels()

    lbMoves = ""
    lbPly = ""
    lbCutoff = ""
    lbScore = ""
    lbPosns = ""

End Sub

Private Sub ComputerMoves()

  'find a move for the computer

    With Board
        picEinst.Visible = True
        btNewGame.Enabled = False
        btEdit.Enabled = False
        Iteration = 0
        InPrinVar = False
        PosnsVisited = 0
        MaxPlySearched = 0
        CutOffs = 0
        lbYMM = "Thinking about"
        lbMoves.ForeColor = IIf(.SideToMove = White, WhitesTurnColor, BlacksTurnColor)
        lbMoves = "??-??"
        MousePointer = vbHourglass
        InitialAlpha = Evaluate(Board, False) - PawnValue
        InitialBeta = InitialAlpha + PawnValue + PawnValue
        BreakRequested = False
        For i = 0 To PlyLimit
            Killer1(i) = ""
            Killer2(i) = ""
            'recycle the principal variation
            PrevPV(i - 2) = CurrPV(0, i)
        Next i
        InPrinVar = (PrevPV(0) <> "" And PrevPV(-1) <> "")  'previous PV was not invalidated; the opponent made the expected move
        TimeLimit = Timer + IIf(scrTimeToThink > 50, 10000000000#, scrTimeToThink * 2)
        Do 'iterative search deepening
            Stalemate = False
            .QuietMoveListFrom = 1
            .QuietMoveListTo = 0
            .CaptureMoveListFrom = 1
            .CaptureMoveListTo = 0
            Iteration = Iteration + 1
            Alpha = InitialAlpha
            Beta = InitialBeta
            Score = Search(Board, 0, Iteration, Alpha, Beta)
            If Score <= InitialAlpha Or Score >= InitialBeta Then
                're-search with open alpha beta window
                Score = Search(Board, 0, Iteration, -Infinity, Infinity)
            End If
            InitialAlpha = Score - PawnValue
            InitialBeta = Score + PawnValue
            'save completed PV
            For i = 0 To PlyLimit
                PrevPV(i) = CurrPV(0, i)
                If PrevPV(i) = "" Then
                    Exit For '>---> Next
                End If
            Next i
            If PrevPV(0) <> "" Then
                lbMoves.ForeColor = IIf(.SideToMove = White, WhitesTurnColor, BlacksTurnColor)
                lbMoves = XlatMove(PrevPV(0), True)
                InPrinVar = True 'second time round there is a PV move
                btBreak.Enabled = True
            End If
        Loop While Timer < TimeLimit And Iteration < PlyLimit And Abs(Score) <= MaxMaterial And Not BreakRequested And Not Stalemate
        btBreak.Enabled = False
        MousePointer = vbNormal
        picEinst.Visible = (Mode = "CC")
        lbPly = MaxPlySearched
        lbScore = Score
        lbCutoff = CutOffs
        lbPosns = PosnsVisited
        Select Case Score
          Case -Infinity
            lbCheck.Visible = False
            DoEvents
            GameEnds = ByLoss
            lbMsg = "I have lost."
            lbMate.Move lbCheck.Left, lbCheck.Top
            lbMate.Visible = True
            lbScore = "-Infinity"
          Case Is < -MaxMaterial
            GameEnds = ByResign
            lbMsg = IIf(Mode = "CC", IIf(.SideToMove = White, "White", "Black") & " resigns.", "I resign.")
            lbScore.ForeColor = IIf(.SideToMove = Black, BlacksTurnColor, WhitesTurnColor)
          Case Else
            If Score > MaxMaterial Then
                i = Infinity - Score
                lbMsg = IIf(Mode = "CC", IIf(.SideToMove = White, "White wins", "Black wins"), "I win") & " in " & i & " move" & IIf(i > 1, "s.", ".")
                DoEvents
                Sleep 555
            End If
            If Stalemate Then
                GameEnds = ByDraw
                lbMsg = Draw
              Else 'STALEMATE = FALSE
                lbYMM = "My move"
                lbMoves.ForeColor = IIf(.SideToMove = White, WhitesTurnColor, BlacksTurnColor)
                lbMoves = XlatMove(CurrPV(0, 0), True)
                If ckView = vbChecked Then
                    lsPV.Clear
                    lsPV.AddItem "PLANNED MOVES"
                    i = 1
                    Do While Len(CurrPV(0, i))
                        If Len(CurrPV(0, i + 1)) = 0 Then
                            lsPV.AddItem "    " & XlatMove(CurrPV(0, i), False)
                            Exit Do '>---> Loop
                          Else 'NOT LEN(CURRPV(0,...
                            lsPV.AddItem "    " & XlatMove(CurrPV(0, i), False) & vbTab & XlatMove(CurrPV(0, i + 1), False)
                        End If
                        i = i + 2
                    Loop
                End If
                lbCheck.Visible = False
                HiliteSquare Asc(Left$(CurrPV(0, 0), 1))
                HiliteSquare Asc(Mid$(CurrPV(0, 0), 2))
                DoEvents
                Sleep 2000
                MakeMove Board, CurrPV(0, 0), True
                MoveCount = MoveCount + 1
                UnHiliteSquare Asc(Left$(CurrPV(0, 0), 1))
                UnHiliteSquare Asc(Mid$(CurrPV(0, 0), 2))
                If Score = Infinity - 1 Then
                    lbScore.ForeColor = IIf(.SideToMove = White, WhitesTurnColor, BlacksTurnColor)
                    lbScore = "+Infinity"
                    DoEvents
                    GameEnds = ByWin
                    lbCheck.Visible = False
                    lbMsg = IIf(Mode = "CC", IIf(.SideToMove = White, "Black", "White") & " has", "You have") & " lost."
                    lbMate.Move TLX(.Pieces(.SideToMove Xor Black) And PieceSquareMask), TLY(.Pieces(.SideToMove Xor Black) And PieceSquareMask) + 15
                    lbMate.Visible = True
                End If
            End If
        End Select
    End With 'BOARD

End Sub

Private Sub CreateHash()

  'create the initial hash for the board

    With Board
        For i = A1 To H8
            j = (.Squares(i) And PieceTypeColorMask) * 8 'free lower 7 bits
            If j Then
                .TTIdx = .TTIdx Xor HashValues(i Or j).Index
                .TTCnf = .TTCnf Xor HashValues(i Or j).Confirm
            End If
        Next i
    End With 'BOARD

End Sub

Private Function DistanceToKing(Board As Board, Square As Long, KingColor As Byte) As Long

  'compute distance square to king of side not to move

    DistanceToKing = Sqr((Rank(Square) - Rank(Board.Pieces(KingColor) And PieceSquareMask)) ^ 2 + (File(Square) - File(Board.Pieces(KingColor) And PieceSquareMask)) ^ 2)

End Function

Private Sub DrawBoard()

  'bitblt the squares

    Rnd -Seed
    For X = 64 To 575 Step 64
        For Y = 64 To 575 Step 64
            BitBlt Me.hDC, X, Y, 64, 64, pcSquare(IIf((X And 64) Xor (Y And 64), 0, 1)).hDC, Rnd * 40, Rnd * 40, SRCCOPY
    Next Y, X

End Sub

Private Function Evaluate(Board As Board, DoPosEval As Boolean) As Long

  'material and positional evaluation

  'todo##
  'pawn structure evaluation##
  'king safety evaluation##

  'bonusses and penalties

  Const EARLY_QUEEN_PENALTY         As Long = 28
  Const DOUBLED_PAWN_PENALTY        As Long = 10
  Const TRIPLED_PAWN_PENALTY        As Long = 40
  Const ISOLATED_PAWN_PENALTY       As Long = 20
  Const BACKWARDS_PAWN_PENALTY      As Long = 8
  Const LOST_CASTLE_RIGHT_PENALTY   As Long = 50
  Const PASSED_PAWN_BONUS           As Long = 20
  Const ROOK_SEMI_OPEN_FILE_BONUS   As Long = 10
  Const ROOK_OPEN_FILE_BONUS        As Long = 15
  Const ROOK_ON_7_BONUS             As Long = 20
  Const DOUBLED_ROOK_BONUS          As Long = 25

  Dim Total                         As Long
  Dim Square                        As Long
  Dim SqFile                        As Long
  Dim SqRank                        As Long
  Dim RookFile                      As Long

    With Board
        Total = .Material(0) - .Material(1)
        If Abs(Total) < EnoughAdvantage Then 'posnl eval if less than enough (currenty enough is more than a pawn)
            If DoPosEval Then
                'White Pieces(bonusses are added and penalties are subtracted)
                RookFile = 0
                For i = 0 To 15
                    If .Pieces(i) And PieceLivingMask Then
                        Square = .Pieces(i) And PieceSquareMask
                        SqFile = File(Square)
                        SqRank = Rank(Square)
                        Select Case .Squares(Square) And PieceTypeMask
                          Case Pawn
                            If .MiscBits And WhiteEndgameMask Then
                                Total = Total + EndgamePawnOnSquare(Square)
                              Else 'NOT .MISCBITS...
                                Total = Total + NormalPawnOnSquare(Square)
                            End If
                          Case Knight
                            Total = Total + KnightOnSquare(Square)
                            If .MiscBits And WhiteEndgameMask Then
                                Total = Total - DistanceToKing(Board, Square, Black) * 1.2
                            End If
                          Case Bishop
                            Total = Total + BishopOnSquare(Square)
                            If .MiscBits And WhiteEndgameMask Then
                                Total = Total - DistanceToKing(Board, Square, Black)
                            End If
                          Case Rook
                            If SqFile = RookFile Then
                                Total = Total + DOUBLED_ROOK_BONUS
                              Else 'NOT SQFILE...
                                RookFile = SqFile
                            End If
                            If SqRank = 7 Then
                                Total = Total + ROOK_ON_7_BONUS
                            End If
                          Case Queen
                            If SqRank > 2 Then
                                If .MiscBits And (WhiteCastleLongMask Or WhiteCastleShortMask) Then
                                    Total = Total - EARLY_QUEEN_PENALTY * (SqRank - 2)
                                End If
                            End If
                          Case King
                            If .MiscBits And WhiteEndgameMask Then
                                Total = Total + EndgameKingOnSquare(Square)
                              Else 'NOT .MISCBITS...
                                Total = Total + NormalKingOnSquare(Square)
                                If Square = A5 Then
                                    If (.MiscBits And WhiteCastleLongMask) = 0 Then
                                        Total = Total - LOST_CASTLE_RIGHT_PENALTY + 2
                                    End If
                                    If (.MiscBits And WhiteCastleShortMask) = 0 Then
                                        Total = Total - LOST_CASTLE_RIGHT_PENALTY - 2
                                    End If
                                End If
                            End If
                        End Select
                    End If
                Next i
                'Black Pieces (bonusses are subtracted and penalties are added)
                RookFile = 0
                For i = 16 To 31
                    If .Pieces(i) And PieceLivingMask Then
                        Square = Mirror(.Pieces(i) And PieceSquareMask)
                        SqFile = File(Square)
                        SqRank = Rank(Square)
                        Select Case .Squares(Square) And PieceTypeMask
                          Case Pawn
                            If .MiscBits And BlackEndgameMask Then
                                Total = Total - EndgamePawnOnSquare(Square)
                              Else 'NOT .MISCBITS...
                                Total = Total - NormalPawnOnSquare(Square)
                            End If
                          Case Knight
                            Total = Total - KnightOnSquare(Square)
                            If .MiscBits And BlackEndgameMask Then
                                Total = Total + DistanceToKing(Board, Square, White) * 1.2
                            End If
                          Case Bishop
                            Total = Total - BishopOnSquare(Square)
                            If .MiscBits And BlackEndgameMask Then
                                Total = Total + DistanceToKing(Board, Square, White)
                            End If
                          Case Rook
                            If SqFile = RookFile Then
                                Total = Total - DOUBLED_ROOK_BONUS
                              Else 'NOT SQFILE...
                                RookFile = SqFile
                            End If
                            If SqRank = 7 Then
                                Total = Total - ROOK_ON_7_BONUS
                            End If
                          Case Queen
                            If SqRank > 2 Then
                                If .MiscBits And (WhiteCastleLongMask Or WhiteCastleShortMask) Then
                                    Total = Total + EARLY_QUEEN_PENALTY * (SqRank - 2)
                                End If
                            End If
                          Case King
                            If .MiscBits And BlackEndgameMask Then
                                Total = Total - EndgameKingOnSquare(Square)
                              Else 'NOT .MISCBITS...
                                Total = Total - NormalKingOnSquare(Square)
                                If Square = A5 Then
                                    If (.MiscBits And WhiteCastleLongMask) = 0 Then
                                        Total = Total + LOST_CASTLE_RIGHT_PENALTY - 2
                                    End If
                                    If (.MiscBits And WhiteCastleShortMask) = 0 Then
                                        Total = Total + LOST_CASTLE_RIGHT_PENALTY + 2
                                    End If
                                End If
                            End If
                        End Select
                    End If
                Next i
                If .MiscBits And PawnMovedMask Then
                    .MiscBits = .MiscBits Xor PawnMovedMask
                    'todo##
                    'pawn structure evaluation##
                    'king safety evaluation##
                End If
            End If
        End If
        If .SideToMove = Black Then
            Evaluate = -Total
          Else 'NOT .SIDETOMOVE...
            Evaluate = Total
        End If
    End With 'BOARD

End Function

Private Sub ckWarn_Click()

    ckWarn.BackColor = IIf(ckWarn = vbChecked, vbCyan, vbYellow)

End Sub

Private Sub Form_Initialize()

  'initialization of program variables

    Screen.MousePointer = vbHourglass
    EndSplash = Now + 4 / 24 / 60 / 60 '4 seconds min splash time
    'Pretend to be a Screensaver
    SystemParametersInfo SPI_GETSCREENSAVEACTIVE, 0&, SaverActive, 0&
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, 0&, ByVal 0&, 0&
    Randomize
    Seed = Rnd * 100
    'used to update castling rights
    CastleBits = Array( _
                 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, _
                 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, _
                 255, 255 And Not WhiteCastleLongMask, 255, 255, 255, 255 And Not WhiteCastleLongMask And Not WhiteCastleShortMask, 255, 255, 255 And Not WhiteCastleShortMask, 255, _
                 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, _
                 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, _
                 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, _
                 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, _
                 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, _
                 255, 255, 255, 255, 255, 255, 255, 255, 255, 255, _
                 255, 255 And Not BlackCastleLongMask, 255, 255, 255, 255 And Not BlackCastleLongMask And Not BlackCastleShortMask, 255, 255, 255 And Not BlackCastleShortMask)

    'setup distance arrays
    QueenDistances = Array( _
                     North, _
                     West, _
                     South, _
                     East, _
                     NorthEast, _
                     NorthWest, _
                     SouthEast, _
                     SouthWest, _
                     South, _
                     East, _
                     North, _
                     West, _
                     SouthWest, _
                     SouthEast, _
                     NorthWest, _
                     NorthEast)

    KnightDistances = Array( _
                      NorthNorth + West, _
                      NorthNorth + East, _
                      WestWest + North, _
                      EastEast + North, _
                      EastEast + South, _
                      WestWest + South, _
                      SouthSouth + East, _
                      SouthSouth + West, _
                      SouthSouth + East, _
                      SouthSouth + West, _
                      EastEast + South, _
                      WestWest + South, _
                      WestWest + North, _
                      EastEast + North, _
                      NorthNorth + West, _
                      NorthNorth + East)

    Mirror = Array( _
             110, 111, 112, 113, 114, 115, 116, 117, 118, 119, _
             100, 101, 102, 102, 104, 105, 106, 107, 108, 109, _
             90, 91, 92, 93, 94, 95, 96, 97, 98, 99, _
             80, 81, 82, 83, 84, 85, 86, 87, 88, 89, _
             70, 71, 72, 73, 74, 75, 76, 77, 78, 79, _
             60, 61, 62, 63, 64, 65, 66, 67, 68, 69, _
             50, 51, 52, 53, 54, 55, 56, 57, 58, 59, _
             40, 41, 42, 43, 44, 45, 46, 47, 48, 49, _
             30, 31, 32, 33, 34, 35, 36, 37, 38, 39, _
             20, 21, 22, 23, 24, 25, 26, 27, 28)

    File = Array( _
           0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
           0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
           0, 1, 2, 3, 4, 5, 6, 7, 8, 0, _
           0, 1, 2, 3, 4, 5, 6, 7, 8, 0, _
           0, 1, 2, 3, 4, 5, 6, 7, 8, 0, _
           0, 1, 2, 3, 4, 5, 6, 7, 8, 0, _
           0, 1, 2, 3, 4, 5, 6, 7, 8, 0, _
           0, 1, 2, 3, 4, 5, 6, 7, 8, 0, _
           0, 1, 2, 3, 4, 5, 6, 7, 8, 0, _
           0, 1, 2, 3, 4, 5, 6, 7, 8)

    Rank = Array( _
           0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
           0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
           1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
           2, 2, 2, 2, 2, 2, 2, 2, 2, 2, _
           3, 3, 3, 3, 3, 3, 3, 3, 3, 3, _
           4, 4, 4, 4, 4, 4, 4, 4, 4, 4, _
           5, 5, 5, 5, 5, 5, 5, 5, 5, 5, _
           6, 6, 6, 6, 6, 6, 6, 6, 6, 6, _
           7, 7, 7, 7, 7, 7, 7, 7, 7, 7, _
           8, 8, 8, 8, 8, 8, 8, 8, 8)

    Mobility = Array( _
               0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
               0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
               0, 2, 3, 3, 3, 3, 3, 3, 2, 0, _
               0, 3, 3, 4, 4, 4, 4, 3, 3, 0, _
               0, 3, 4, 5, 6, 6, 5, 4, 3, 0, _
               0, 3, 4, 6, 9, 9, 6, 4, 3, 0, _
               0, 3, 4, 6, 9, 9, 6, 4, 3, 0, _
               0, 3, 4, 5, 6, 6, 5, 4, 3, 0, _
               0, 3, 3, 4, 4, 4, 4, 3, 3, 0, _
               0, 2, 3, 3, 3, 3, 3, 3, 2)

    NormalPawnOnSquare = Array( _
                         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                         0, 0, 0, 0, -15, -35, -3, -4, 0, 0, _
                         0, 0, 1, 2, 4, -10, 0, -2, -3, 0, _
                         0, 1, 2, 4, 8, 8, 4, 2, 1, 0, _
                         0, 3, 5, 9, 12, 12, 9, 6, 3, 0, _
                         0, 4, 8, 12, 16, 16, 12, 8, 4, 0, _
                         0, 5, 10, 15, 20, 20, 15, 10, 5)

    EndgamePawnOnSquare = Array( _
                          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                          0, 2, 2, 2, 2, 2, 2, 2, 2, 0, _
                          0, 4, 4, 4, 4, 4, 4, 4, 4, 0, _
                          0, 8, 8, 8, 8, 8, 8, 8, 8, 0, _
                          0, 12, 12, 12, 12, 12, 12, 12, 12, 0, _
                          0, 16, 16, 16, 16, 16, 16, 16, 16, 0, _
                          0, 20, 20, 20, 20, 20, 20, 20, 20)

    KnightOnSquare = Array( _
                     0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                     0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                     0, -10, -28, -10, -10, -10, -10, -28, -10, 0, _
                     0, -10, 0, 0, 0, 0, 0, 0, -10, 0, _
                     0, -10, 0, 5, 5, 5, 5, 0, -10, 0, _
                     0, -10, 0, 5, 10, 10, 5, 0, -10, 0, _
                     0, -10, 0, 5, 10, 10, 5, 0, -10, 0, _
                     0, -10, 0, 5, 7, 5, 7, 0, -10, 0, _
                     0, -10, 0, 0, 0, 7, 0, 0, -10, 0, _
                     0, -10, -10, -10, -10, -10, -10, -10, -10)

    BishopOnSquare = Array( _
                     0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                     0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                     0, -15, -10, -20, -10, -10, -20, -10, -15, 0, _
                     0, -10, 0, 0, 0, 0, 0, 0, -10, 0, _
                     0, -10, 0, 5, 5, 5, 5, 0, -10, 0, _
                     0, -10, 0, 5, 10, 10, 5, 0, -10, 0, _
                     0, -10, 0, 5, 10, 10, 5, 0, -10, 0, _
                     0, -10, 0, 5, 5, 5, 5, 0, -10, 0, _
                     0, -10, 0, 0, 0, 0, 0, 0, -10, 0, _
                     0, -15, -10, -10, -10, -10, -10, -10, -15)

    NormalKingOnSquare = Array( _
                         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                         0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                         0, 10, 30, 35, -10, 0, -10, 40, 30, 0, _
                         0, -20, -20, -20, -20, -20, -20, -20, -20, 0, _
                         0, -40, -40, -40, -40, -40, -40, -40, -40, 0, _
                         0, -40, -40, -40, -40, -40, -40, -40, -40, 0, _
                         0, -40, -40, -40, -40, -40, -40, -40, -40, 0, _
                         0, -40, -40, -40, -40, -40, -40, -40, -40, 0, _
                         0, -40, -40, -40, -40, -40, -40, -40, -40, 0, _
                         0, -40, -40, -40, -40, -40, -40, -40, -40)

    EndgameKingOnSquare = Array( _
                          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
                          0, 0, 10, 20, 30, 30, 20, 10, 0, 0, _
                          0, 10, 20, 30, 40, 40, 30, 20, 10, 0, _
                          0, 20, 30, 40, 50, 50, 40, 30, 20, 0, _
                          0, 30, 40, 50, 60, 60, 50, 40, 30, 0, _
                          0, 30, 40, 50, 60, 60, 50, 40, 30, 0, _
                          0, 20, 30, 40, 50, 50, 40, 30, 20, 0, _
                          0, 10, 20, 30, 40, 40, 30, 20, 10, 0, _
                          0, 0, 10, 20, 30, 30, 20, 10, 0)

    'setup piece values
    PieceValues(0) = 0
    PieceValues(1) = QueenValue
    PieceValues(2) = BishopValue
    PieceValues(3) = KnightValue
    PieceValues(4) = RookValue
    PieceValues(5) = BishopValue
    PieceValues(6) = KnightValue
    PieceValues(7) = RookValue
    For i = 8 To 15
        PieceValues(i) = PawnValue
    Next i
    PieceValues(16) = 0
    PieceValues(17) = QueenValue
    PieceValues(18) = BishopValue
    PieceValues(19) = KnightValue
    PieceValues(20) = RookValue
    PieceValues(21) = BishopValue
    PieceValues(22) = KnightValue
    PieceValues(23) = RookValue
    For i = 24 To 31
        PieceValues(i) = PawnValue
    Next i

    'initialize hash square-piece values
    For i = LBound(HashValues) To UBound(HashValues)
        HashValues(i).Index = Int(Rnd * (TTSize + 1))
        HashValues(i).Confirm = Int(Rnd * (TTSize + 1))
    Next i
    For i = LBound(TT) To UBound(TT)
        With TT(i)
            .Confirm = 0
            .Value = 0
            .Confidence = 0
        End With 'TT(I)
    Next i

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyDown = (WhiteKingCount = 1 And BlackKingCount = 1) 'no exit out of edit unless one king each color
    KeyCode = 0 'Consume

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyDown = False
    KeyCode = 0 'Consume

End Sub

Private Sub Form_Load()

  'create a standard board

    fr(0).BackColor = BackColor
    fr(1).BackColor = BackColor
    ckView.BackColor = BackColor
    opPlaySelf.BackColor = BackColor
    opPlayAlt.BackColor = BackColor
    opComp.BackColor = BackColor
    opHuman.BackColor = BackColor
    lblVers = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'prepare the piece images (images stolen from and by permission of RJSoft)
    For i = 1 To 9
        Load imWhiteQueen(i)
        Load imBlackQueen(i)
        Load imWhiteBishop(i)
        Load imBlackBishop(i)
        Load imWhiteKnight(i)
        Load imBlackKnight(i)
        Load imWhiteRook(i)
        Load imBlackRook(i)
        Load imWhitePawn(i)
        Load imBlackPawn(i)
    Next i
    opComp_Click
    ckView = vbUnchecked
    ckView_Click
    ckReverse = vbUnchecked
    StandardBoard
    Board.SideToMove = White
    lbSide.BackColor = WhitesTurnColor
    ClickEnabled = False
    scrTimeToThink = 30 '30/10 Minutes = 3 Minutes
    If EndSplash Then
        With mhFocus
            'add message hooks to absorb WM_SETFOCUS
            .Add btBreak.hWnd, WM_SETFOCUS
            .Add btEdit.hWnd, WM_SETFOCUS
            .Add btGo.hWnd, WM_SETFOCUS
            .Add btNewGame.hWnd, WM_SETFOCUS
            .Add ckReverse.hWnd, WM_SETFOCUS
            .Add ckView.hWnd, WM_SETFOCUS
            .Add ckWarn.hWnd, WM_SETFOCUS
            .Add opComp.hWnd, WM_SETFOCUS
            .Add opHuman.hWnd, WM_SETFOCUS
            .Add opPlayAlt.hWnd, WM_SETFOCUS
            .Add opPlaySelf.hWnd, WM_SETFOCUS
            .Add scrTimeToThink.hWnd, WM_SETFOCUS
        End With 'MHFOCUS
        Do Until Now > EndSplash
        Loop
        Unload fSplash
        Set fSplash = Nothing
        EndSplash = 0
        Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    End If
    Show
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim DispColor As Long

    If X > 63 And X < 576 And Y > 64 And Y < 576 Then
        If ClickEnabled Then
            lbCheck.Visible = False
            X = X \ 64
            Y = Y \ 64
            DispColor = BlacksTurnColor
            If Button = (vbLeftButton And Editing) Or (Board.SideToMove = White And Not Editing) Then
                DispColor = WhitesTurnColor
            End If
            SquareToLabel IIf(ckReverse = vbChecked, (Y + 1) * 10 + 9 - X, (10 - Y) * 10 + X), DispColor
        End If
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If ClickEnabled Then
        If X > 63 And X < 576 And Y > 63 And Y < 576 Then
            MousePointer = vbCustom
          Else 'NOT X...
            MousePointer = vbNormal
        End If
      Else 'CLICKENABLED = FALSE
        MousePointer = IIf(btNewGame.Enabled, vbNormal, vbHourglass)
    End If
    LastX = X
    LastY = Y

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ClickButton = Button
    FormClicked = ClickEnabled And X > 63 And X < 576 And Y > 63 And Y < 576

End Sub

Private Sub Form_Paint()

    Enabled = False
    DrawBoard
    If Len(Hilited) > 1 Then
        For i = 1 To Len(Hilited)
            HiliteSquare Asc(Mid$(Hilited, i, 1))
        Next i
      Else 'NOT LEN(HILITED)...
        Hilited = ""
    End If
    Enabled = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    BreakRequested = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    With mhFocus
        'remove the messagehooks
        .Remove btBreak.hWnd, WM_SETFOCUS
        .Remove btEdit.hWnd, WM_SETFOCUS
        .Remove btGo.hWnd, WM_SETFOCUS
        .Remove btNewGame.hWnd, WM_SETFOCUS
        .Remove ckReverse.hWnd, WM_SETFOCUS
        .Remove ckView.hWnd, WM_SETFOCUS
        .Remove ckWarn.hWnd, WM_SETFOCUS
        .Remove opComp.hWnd, WM_SETFOCUS
        .Remove opHuman.hWnd, WM_SETFOCUS
        .Remove opPlayAlt.hWnd, WM_SETFOCUS
        .Remove opPlaySelf.hWnd, WM_SETFOCUS
        .Remove scrTimeToThink.hWnd, WM_SETFOCUS
    End With 'MHFOCUS
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, SaverActive, ByVal 0&, 0&
    End

End Sub

Private Function GenerateMoves(Board As Board, ForHuman As Boolean) As Long

  'Returns the move list index of the best capture if any
  'or (-1) if it could capture the opposing king or zero else

  Dim f
  Dim SquareFrom
  Dim SquareTo
  Dim Direction
  Dim CapturedPiece As Byte
  Dim PromoteTo     As Byte
  Dim CapturedColor As Byte
  Dim PromotionDelta
  Dim MaterialGain
  Dim BestGain
  Dim BestMove
  Dim StopSlide     As Boolean
  Dim MobilityValue

    BestGain = 0
    BestMove = 0
    With Board
        QuietMoveListIx = .QuietMoveListFrom - 1
        CaptureMoveListIx = .CaptureMoveListFrom - 1
        If .SideToMove = White Then
            f = 0
          Else 'NOT .SIDETOMOVE...
            f = 8
        End If
        For i = .SideToMove + 15 To .SideToMove Step -1 'start with the pawn moves
            If .Pieces(i) And PieceLivingMask Then
                SquareFrom = .Pieces(i) And PieceSquareMask
                Select Case .Squares(SquareFrom) And PieceTypeMask
                  Case Pawn
                    If .SideToMove = White Then
                        GoSub WhitePawnMoves
                      Else 'NOT .SIDETOMOVE...
                        GoSub BlackPawnMoves
                    End If
                  Case Knight
                    For j = f To f + 7
                        SquareTo = SquareFrom + KnightDistances(j)
                        GoSub CheckToSquare
                    Next j
                  Case Bishop
                    GoSub BishopMoves
                  Case Rook
                    GoSub RookMoves
                  Case Queen
                    GoSub BishopMoves
                    GoSub RookMoves
                  Case King
                    For j = f To f + 7
                        SquareTo = SquareFrom + QueenDistances(j)
                        GoSub CheckToSquare
                    Next j
                    If Board.MiscBits And CastleMask Then
                        If Board.SideToMove = White Then
                            If Board.MiscBits And WhiteCastleShortMask Then
                                GoSub CastleShort
                            End If
                            If Board.MiscBits And WhiteCastleLongMask Then
                                GoSub CastleLong
                            End If
                          Else 'NOT BOARD.SIDETOMOVE...
                            If Board.MiscBits And BlackCastleShortMask Then
                                GoSub CastleShort
                            End If
                            If Board.MiscBits And BlackCastleLongMask Then
                                GoSub CastleLong
                            End If
                        End If
                    End If
                End Select
            End If
            If BestMove = -1 Then 'could capture the opposing king
                Exit For '>---> Next
            End If
        Next i
        .QuietMoveListTo = QuietMoveListIx
        .CaptureMoveListTo = CaptureMoveListIx
    End With 'BOARD
    GenerateMoves = BestMove

Exit Function

    '-------------------------------------------------------------------------------
    'Local Subs

CastleShort:
    If (Board.Squares(SquareFrom + East) And PieceTypeMask) = Free And (Board.Squares(SquareFrom + EastEast) And PieceTypeMask) = Free Then
        SquareTo = SquareFrom + EastEast
        If ForHuman Then
            If Not (IsAttacked(Board, SquareTo, Board.SideToMove Xor Black) Or IsAttacked(Board, SquareTo + West, Board.SideToMove Xor Black) Or IsAttacked(Board, SquareTo + WestWest, Board.SideToMove Xor Black) Or IsAttacked(Board, SquareFrom, Board.SideToMove Xor Black)) Then
                'put opponent's castling move into list only if it is legal
                QuietMoveListIx = QuietMoveListIx + 1
                QuietMoves(QuietMoveListIx) = Chr$(SquareFrom) & Chr$(SquareTo) & "00"
            End If
          Else 'FORHUMAN = FALSE
            'for the computer the legal check is deferred until we actually make the move
            'we might not need it because of cutoffs and the check for Attacked is time-expensive
            QuietMoveListIx = QuietMoveListIx + 1
            QuietMoves(QuietMoveListIx) = Chr$(SquareFrom) & Chr$(SquareTo) & "00"
        End If
    End If
    Return

CastleLong:
    If (Board.Squares(SquareFrom + West) And PieceTypeMask) = Free And (Board.Squares(SquareFrom + WestWest) And PieceTypeMask) = Free And (Board.Squares(SquareFrom + WestWestWest) And PieceTypeMask) = Free Then
        SquareTo = SquareFrom + WestWest
        If ForHuman Then
            If Not (IsAttacked(Board, SquareTo, Board.SideToMove Xor Black) Or IsAttacked(Board, SquareTo + East, Board.SideToMove Xor Black) Or IsAttacked(Board, SquareTo + EastEast, Board.SideToMove Xor Black) Or IsAttacked(Board, SquareFrom, Board.SideToMove Xor Black)) Then
                'put opponent's castling move into list only if it is legal
                QuietMoveListIx = QuietMoveListIx + 1
                QuietMoves(QuietMoveListIx) = Chr$(SquareFrom) & Chr$(SquareTo) & "000"
            End If
          Else 'FORHUMAN = FALSE
            'for the computer the legal check is deferred until we actually make the move
            QuietMoveListIx = QuietMoveListIx + 1
            QuietMoves(QuietMoveListIx) = Chr$(SquareFrom) & Chr$(SquareTo) & "000"
        End If
    End If
    Return

BishopMoves:
    For j = f + 4 To f + 7
        Direction = QueenDistances(j)
        GoSub SlidePiece
    Next j
    Return

RookMoves:
    For j = f To f + 3
        Direction = QueenDistances(j)
        GoSub SlidePiece
    Next j
    Return

SlidePiece:
    StopSlide = False
    SquareTo = SquareFrom
    Do
        SquareTo = SquareTo + Direction
        GoSub CheckToSquare
    Loop Until StopSlide
    Return

CheckToSquare:
    CapturedPiece = Board.Squares(SquareTo) And PieceTypeMask
    CapturedColor = Board.Squares(SquareTo) And PieceColorMask
    If CapturedPiece = Forbidden Then
        StopSlide = True
      ElseIf CapturedPiece = Free Then 'NOT CAPTUREDPIECE...
        GoSub RecordQuietMove
      ElseIf CapturedColor = Board.SideToMove Then 'NOT CAPTUREDPIECE...
        StopSlide = True
      Else 'NOT CAPTUREDCOLOR...
        GoSub RecordCapture
        StopSlide = True
    End If
    Return

WhitePawnMoves:
    SquareTo = SquareFrom + North
    If (Board.Squares(SquareTo) And PieceTypeMask) = Free Then
        If SquareTo > H7 Then
            GoSub WhitePromotion
          Else 'NOT SQUARETO...
            GoSub RecordQuietMove
        End If
        If SquareFrom < A3 Then
            'double advance
            SquareTo = SquareFrom + NorthNorth
            If (Board.Squares(SquareTo) And PieceTypeMask) = Free Then
                GoSub RecordQuietMove
            End If
        End If
    End If
    'en passant
    If SquareFrom + West = Board.EnPassant Or SquareFrom + East = Board.EnPassant Then
        SquareTo = Board.EnPassant + North
        GoSub RecordCapture
    End If
    'capture
    SquareTo = SquareFrom + NorthEast
    GoSub CheckWhitePawnCapture
    SquareTo = SquareFrom + NorthWest
    'dropthru

CheckWhitePawnCapture:
    CapturedColor = Board.Squares(SquareTo) And PieceColorMask
    If CapturedColor = Black Then
        CapturedPiece = Board.Squares(SquareTo) And PieceTypeMask
        If SquareTo > H7 Then
            GoSub WhitePromotion
          Else 'NOT SQUARETO...
            GoSub RecordCapture
        End If
    End If
    Return

WhitePromotion:
    If ForHuman Then
        PromoteTo = 1 'dont know what he will do
        GoSub RecordCapture
      Else 'FORHUMAN = FALSE
        PromoteTo = WhiteQueen
        PromotionDelta = QueenValue - PawnValue
        GoSub RecordCapture
        PromoteTo = WhiteRook
        PromotionDelta = RookValue - PawnValue
        GoSub RecordCapture
        PromoteTo = WhiteBishop
        PromotionDelta = BishopValue - PawnValue
        GoSub RecordCapture
        PromoteTo = WhiteKnight
        PromotionDelta = KnightValue - PawnValue
        GoSub RecordCapture
    End If
    PromoteTo = 0
    PromotionDelta = 0
    Return

BlackPawnMoves:
    SquareTo = SquareFrom + South
    If (Board.Squares(SquareTo) And PieceTypeMask) = Free Then
        If SquareTo < A2 Then
            GoSub BlackPromotion
          Else 'NOT SQUARETO...
            GoSub RecordQuietMove
        End If
        If SquareFrom > H6 Then
            'double advance
            SquareTo = SquareFrom + SouthSouth
            If (Board.Squares(SquareTo) And PieceTypeMask) = Free Then
                GoSub RecordQuietMove
            End If
        End If
    End If
    'en passant
    If SquareFrom + West = Board.EnPassant Or SquareFrom + East = Board.EnPassant Then
        SquareTo = Board.EnPassant + South
        GoSub RecordCapture
    End If
    'capture
    SquareTo = SquareFrom + SouthWest
    GoSub CheckBlackPawnCapture
    SquareTo = SquareFrom + SouthEast
    'dropthru

CheckBlackPawnCapture:
    CapturedPiece = Board.Squares(SquareTo) And PieceTypeMask
    CapturedColor = Board.Squares(SquareTo) And PieceColorMask
    If CapturedPiece <> Free Then
        If CapturedColor = White Then
            If CapturedPiece <> Forbidden Then
                If SquareTo < A2 Then
                    GoSub BlackPromotion
                  Else 'NOT SQUARETO...
                    GoSub RecordCapture
                End If
            End If
        End If
    End If
    Return

BlackPromotion:
    If ForHuman Then
        PromoteTo = 1 'dont know what he will do
        GoSub RecordCapture
      Else 'FORHUMAN = FALSE
        PromoteTo = BlackQueen
        PromotionDelta = QueenValue - PawnValue
        GoSub RecordCapture
        PromoteTo = BlackRook
        PromotionDelta = RookValue - PawnValue
        GoSub RecordCapture
        PromoteTo = BlackBishop
        PromotionDelta = BishopValue - PawnValue
        GoSub RecordCapture
        PromoteTo = BlackKnight
        PromotionDelta = KnightValue - PawnValue
        GoSub RecordCapture
    End If
    PromoteTo = 0
    PromotionDelta = 0
    Return

RecordQuietMove:
    QuietMoveListIx = QuietMoveListIx + 1
    QuietMoves(QuietMoveListIx) = Chr$(SquareFrom) & Chr$(SquareTo)
    MobilityValue = MobilityValue + Mobility(SquareTo)
    Return

RecordCapture:
    If ForHuman Then
        'all human moves go into the quiet list; makes it easier to check them
        QuietMoveListIx = QuietMoveListIx + 1
        QuietMoves(QuietMoveListIx) = Chr$(SquareFrom) & Chr$(SquareTo) & Chr$(PromoteTo)
      Else 'FORHUMAN = FALSE
        If CapturedPiece = King Then
            BestMove = -1
            BestGain = MaxMaterial
          Else 'NOT CAPTUREDPIECE...
            CaptureMoveListIx = CaptureMoveListIx + 1
            CaptureMoves(CaptureMoveListIx) = Chr$(SquareFrom) & Chr$(SquareTo) & Chr$(PromoteTo)
            'most valuable victim / least valuable aggressor may be better##
            MaterialGain = (PieceValues(Board.Squares(SquareTo) And PieceListIndexMask)) + PromotionDelta
            If MaterialGain > BestGain Then
                BestMove = CaptureMoveListIx
                BestGain = MaterialGain
            End If
        End If
        MobilityValue = MobilityValue + Mobility(SquareTo)
    End If
    Return

End Function

Private Function GetPromotionPiece(Square As Long) As Long

  'show promotion list

    lstPromo.Move TLX(Square) + 1, TLY(Square) + 1
    With lstPromo
        For j = 0 To 3
            .Selected(j) = False
        Next j
        .ListIndex = -1
        lstPromo.Visible = True
        .SetFocus
        Do While lstPromo.Visible
            DoEvents
        Loop
        Select Case True
          Case .Selected(0)
            GetPromotionPiece = Queen
          Case .Selected(1)
            GetPromotionPiece = Rook
          Case .Selected(2)
            GetPromotionPiece = Bishop
          Case .Selected(3)
            GetPromotionPiece = Knight
        End Select
    End With 'LSTPROMO

End Function

Private Sub HiliteSquare(Square As Long)

  'draw frame around a square

  Dim HiliteColor As OLE_COLOR

    If IsAttacked(Board, Square, Board.SideToMove Xor Black) Then
        HiliteColor = ckWarn.BackColor
      Else 'NOT ISATTACKED(BOARD,...
        HiliteColor = vbYellow
    End If
    DrawStyle = vbDot
    Line (TLX(Square) + 1, TLY(Square) + 1)-Step(61, 0), HiliteColor
    Line -Step(0, 61), HiliteColor
    Line -Step(-61, 0), HiliteColor
    Line -Step(0, -61), HiliteColor

End Sub

Private Sub HumanMoves()

  'accept and check a human move

  Dim TempBoard     As Board

    With Board
        .QuietMoveListFrom = 1
        InPrinVar = False
        If GenerateMoves(Board, True) Or .QuietMoveListFrom <= .QuietMoveListTo Then   'opponent can move
            lbYMM = "Your move"
            ClickEnabled = True
            Form_MouseMove 0, 0, LastX, LastY
            lbMsg = "Enter your move:"
            HumanMove = ""
            lbMoves = ""
            Do
                FormClicked = False
                btNewGame.Enabled = True
                btEdit.Enabled = True
                Do
                    DoEvents
                    If GameEnds Then
                        Exit Sub '>---> Bottom
                    End If
                Loop While Not FormClicked
                lbCutoff = ""
                lbPly = ""
                lbPosns = ""
                lbScore = ""
                If FormClicked Then 'search opponent's move list
                    For j = 1 To Len(Hilited)
                        UnHiliteSquare Asc(Mid$(Hilited, j, 1))
                    Next j
                    Hilited = Left$(HumanMove, 1)
                    j = 0
                    For i = 1 To QuietMoveListIx
                        If HumanMove = Left$(QuietMoves(i), Len(HumanMove)) Then
                            j = j + 1
                            k = i
                            Hilited = Hilited & Mid$(QuietMoves(i), 2, 1)
                        End If
                    Next i
                    Select Case j
                      Case 0 'not in opponent's movelist
                        lbMsg = "You tried to enter an illegal move."
                        HumanMove = ""
                        lbMoves = ""
                      Case 1 'move uniquely defined - make it (the touched piece rule)
                        lbMsg = ""
                        TempBoard = Board
                        'try move on tempboard (without paint)
                        MakeMove TempBoard, Left$(QuietMoves(k), 2), False
                        For j = 1 To Len(Hilited)
                            UnHiliteSquare Asc(Mid$(Hilited, j, 1))
                        Next j
                        Hilited = ""
                        If IsAttacked(TempBoard, TempBoard.Pieces(.SideToMove) And PieceSquareMask, .SideToMove Xor Black) Then 'king in check
                            lbMsg = "That move would put / leave you in Check."
                            HumanMove = ""
                            lbMoves = ""
                          Else 'move is ok -> make it on real board with paint'NOT ISATTACKED(TEMPBOARD,...
                            If Len(QuietMoves(k)) = 3 Then
                                If Asc(Right$(QuietMoves(k), 1)) = 1 Then
                                    Mid$(QuietMoves(k), 3, 1) = Chr$(GetPromotionPiece(Asc(Mid$(QuietMoves(k), 2, 1))) Or .SideToMove)
                                End If
                            End If
                            lbMoves.ForeColor = IIf(.SideToMove = White, WhitesTurnColor, BlacksTurnColor)
                            lbMoves = XlatMove(QuietMoves(k), True)
                            If QuietMoves(k) <> PrevPV(1) Then
                                'opponent made a different move - invalidate PV
                                PrevPV(2) = ""
                            End If
                            MakeMove Board, QuietMoves(k), True
                            If Mode = "HH" Then
                                lbScore = Evaluate(Board, True)
                            End If
                            DoEvents
                            Sleep 500
                            Exit Do 'and exit '>---> Loop
                        End If
                      Case Else 'not unique yet
                        For j = 1 To Len(Hilited)
                            HiliteSquare Asc(Mid$(Hilited, j, 1))
                        Next j
                    End Select
                End If
            Loop
            ClickEnabled = False
            If Rnd < 0.1 Then
                lbMsg = "Do you mind me smoking?"
            End If
          Else 'opponent cannot move'NOT GENERATEMOVES(BOARD,...
            GameEnds = ByDraw
            lbMsg = Draw
            lbScore.ForeColor = NeutralColor
            lbScore = "0"
        End If
    End With 'BOARD

End Sub

Private Sub imBlackBishop_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imBlackBishop(Index).Left, imBlackBishop(Index).Top

End Sub

Private Sub imBlackBishop_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imBlackBishop(Index).Left, imBlackBishop(Index).Top

End Sub

Private Sub imBlackKing_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imBlackKing.Left, imBlackKing.Top

End Sub

Private Sub imBlackKing_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imBlackKing.Left, imBlackKing.Top

End Sub

Private Sub imBlackKnight_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imBlackKnight(Index).Left, imBlackKnight(Index).Top

End Sub

Private Sub imBlackKnight_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imBlackKnight(Index).Left, imBlackKnight(Index).Top

End Sub

Private Sub imBlackPawn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imBlackPawn(Index).Left, imBlackPawn(Index).Top

End Sub

Private Sub imBlackPawn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imBlackPawn(Index).Left, imBlackPawn(Index).Top

End Sub

Private Sub imBlackQueen_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imBlackQueen(Index).Left, imBlackQueen(Index).Top

End Sub

Private Sub imBlackQueen_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imBlackQueen(Index).Left, imBlackQueen(Index).Top

End Sub

Private Sub imBlackRook_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imBlackRook(Index).Left, imBlackRook(Index).Top

End Sub

Private Sub imBlackRook_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imBlackRook(Index).Left, imBlackRook(Index).Top

End Sub

Private Sub imWhiteBishop_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imWhiteBishop(Index).Left, imWhiteBishop(Index).Top

End Sub

Private Sub imWhiteBishop_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imWhiteBishop(Index).Left, imWhiteBishop(Index).Top

End Sub

Private Sub imWhiteKing_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imWhiteKing.Left, imWhiteKing.Top

End Sub

Private Sub imWhiteKing_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imWhiteKing.Left, imWhiteKing.Top

End Sub

Private Sub imWhiteKnight_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imWhiteKnight(Index).Left, imWhiteKnight(Index).Top

End Sub

Private Sub imWhiteKnight_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imWhiteKnight(Index).Left, imWhiteKnight(Index).Top

End Sub

Private Sub imWhitePawn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imWhitePawn(Index).Left, imWhitePawn(Index).Top

End Sub

Private Sub imWhitePawn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imWhitePawn(Index).Left, imWhitePawn(Index).Top

End Sub

Private Sub imWhiteQueen_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imWhiteQueen(Index).Left, imWhiteQueen(Index).Top

End Sub

Private Sub imWhiteQueen_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imWhiteQueen(Index).Left, imWhiteQueen(Index).Top

End Sub

Private Sub imWhiteRook_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form

    Form_MouseDown Button, Shift, imWhiteRook(Index).Left, imWhiteRook(Index).Top

End Sub

Private Sub imWhiteRook_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Form_MouseUp Button, Shift, imWhiteRook(Index).Left, imWhiteRook(Index).Top

End Sub

Private Function IsAttacked(Board As Board, Square As Long, ByColor As Byte) As Boolean

  'ceck if square is attacked by color

  Dim SquareToCheck
  Dim AttackingPiece As Byte
  Dim Found          As Boolean

    With Board
        If (.Squares(Square) And PieceTypeMask) <> Forbidden Then
            If ByColor = White Then 'white attack ?
                If (.Squares(Square + SouthWest) And PieceTypeColorMask) = WhitePawn Or (.Squares(Square + SouthEast) And PieceTypeColorMask) = WhitePawn Then
                    Found = True
                  Else 'NOT (.SQUARES(SQUARE...
                    For l = 0 To 7
                        If (.Squares(Square + QueenDistances(l)) And PieceTypeColorMask) = WhiteKing Then
                            Found = True
                          ElseIf (.Squares(Square + KnightDistances(l)) And PieceTypeColorMask) = WhiteKnight Then 'NOT (.SQUARES(SQUARE...
                            Found = True
                          Else 'NOT (.SQUARES(SQUARE...
                            SquareToCheck = Square
                            Do
                                SquareToCheck = SquareToCheck + QueenDistances(l)
                                AttackingPiece = (.Squares(SquareToCheck) And PieceTypeColorMask)
                                If AttackingPiece = WhiteQueen Then
                                    Found = True
                                    Exit Do '>---> Loop
                                End If
                                If AttackingPiece = WhiteRook Then
                                    If QueenDistances(l) = North Or QueenDistances(l) = West Or QueenDistances(l) = South Or QueenDistances(l) = East Then
                                        Found = True
                                        Exit Do '>---> Loop
                                    End If
                                  ElseIf AttackingPiece = WhiteBishop Then 'NOT ATTACKINGPIECE...
                                    If QueenDistances(l) = NorthWest Or QueenDistances(l) = NorthEast Or QueenDistances(l) = SouthWest Or QueenDistances(l) = SouthEast Then
                                        Found = True
                                        Exit Do '>---> Loop
                                    End If
                                End If
                            Loop While AttackingPiece = Free
                        End If
                        If Found Then
                            Exit For '>---> Next
                        End If
                    Next l
                End If
              Else 'black attack ?'NOT BYCOLOR...
                If (.Squares(Square + NorthWest) And PieceTypeColorMask) = BlackPawn Or (.Squares(Square + NorthEast) And PieceTypeColorMask) = BlackPawn Then
                    Found = True
                  Else 'NOT (.SQUARES(SQUARE...
                    For l = 8 To 15
                        If (.Squares(Square + QueenDistances(l)) And PieceTypeColorMask) = BlackKing Then
                            Found = True
                          ElseIf (.Squares(Square + KnightDistances(l)) And PieceTypeColorMask) = BlackKnight Then 'NOT (.SQUARES(SQUARE...
                            Found = True
                          Else 'NOT (.SQUARES(SQUARE...
                            SquareToCheck = Square
                            Do
                                SquareToCheck = SquareToCheck + QueenDistances(l)
                                AttackingPiece = (.Squares(SquareToCheck) And PieceTypeColorMask)
                                If AttackingPiece = BlackQueen Then
                                    Found = True
                                    Exit Do '>---> Loop
                                End If
                                If AttackingPiece = BlackRook Then
                                    If QueenDistances(l) = North Or QueenDistances(l) = West Or QueenDistances(l) = South Or QueenDistances(l) = East Then
                                        Found = True
                                        Exit Do '>---> Loop
                                    End If
                                  ElseIf AttackingPiece = BlackBishop Then 'NOT ATTACKINGPIECE...
                                    If QueenDistances(l) = NorthWest Or QueenDistances(l) = NorthEast Or QueenDistances(l) = SouthWest Or QueenDistances(l) = SouthEast Then
                                        Found = True
                                        Exit Do '>---> Loop
                                    End If
                                End If
                            Loop While AttackingPiece = Free
                        End If
                        If Found Then
                            Exit For '>---> Next
                        End If
                    Next l
                End If
            End If
        End If
    End With 'BOARD
    IsAttacked = Found

End Function

Private Sub lbCheck_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'pass click to form
  'lbCheck disappears on mouse_down and will not see mouse_up and the king underneath
  'will not see it either so we have to fake a mouse_up for the form

    Form_MouseDown vbLeftButton, 0, lbCheck.Left + lbCheck.Width / 2, lbCheck.Top
    Form_MouseUp vbLeftButton, 0, lbCheck.Left + lbCheck.Width / 2, lbCheck.Top

End Sub

Private Sub lbMsg_Change()

  'a message was output - so restart the sponge timer

    tmr.Enabled = False
    tmr.Enabled = (GameEnds = 0) 'restart timer unless game ends

End Sub

Private Sub lstPromo_ItemCheck(Item As Integer)

  'hide promotion list once the user has selected the promotion piece

    Sleep 300
    lstPromo.Visible = False

End Sub

Private Sub MakeMove(Board As Board, Move As String, Paint As Boolean)

  'Alters the Board, PieceList and Material Values, and updates the UI and the TT

  Dim MovingPiece   As Byte
  Dim CapturedPiece As Byte
  Dim FromSquare    As Long
  Dim ToSquare      As Long
  Dim NewPiece      As Byte
  Dim MoveDist      As Long
  Dim LostPiece
  Dim PieceNum
  Dim SideNum

    With Board
        FromSquare = Asc(Left$(Move, 1))
        ToSquare = Asc(Mid$(Move, 2))
        MovingPiece = .Squares(FromSquare) And PieceTypeMask
        CapturedPiece = .Squares(ToSquare)
        .EnPassant = 0 'reset enpassant square, may be set anew
        Select Case True
          Case MovingPiece = Pawn
            MoveDist = Abs(FromSquare - ToSquare)
            If MoveDist = NorthEast Or MoveDist = NorthWest Then
                If CapturedPiece = Free Then 'enpassant
                    MoveDist = Sgn(ToSquare - FromSquare) * North
                    ToSquare = ToSquare - MoveDist
                    'capture enpassant pawn
                    MakeMove Board, Chr$(FromSquare) & Chr$(ToSquare), Paint
                    'move to final destination
                    FromSquare = ToSquare
                    ToSquare = FromSquare + MoveDist
                End If
            End If
            If MoveDist = NorthNorth Then 'move dist is abs
                .EnPassant = ToSquare
            End If
          Case MovingPiece = King
            If .SideToMove = White Then
                .MiscBits = .MiscBits And Not WhiteCastleLongMask And Not WhiteCastleShortMask
              Else 'NOT .SIDETOMOVE...
                .MiscBits = .MiscBits And Not BlackCastleLongMask And Not BlackCastleShortMask
            End If
            If Abs(FromSquare - ToSquare) = 2 Then 'castling
                If Not Recur Then
                    Recur = True 'prevent recursion
                    'make king move
                    MakeMove Board, Move, Paint
                    'prepare rookmove
                    If ToSquare < FromSquare Then 'castling queenside
                        FromSquare = ToSquare + WestWest
                        ToSquare = ToSquare + East
                      Else 'castling kingside'NOT TOSQUARE...
                        FromSquare = ToSquare + East
                        ToSquare = ToSquare + West
                    End If
                    MovingPiece = .Squares(FromSquare) And PieceTypeMask
                    Recur = False
                End If
            End If
        End Select
        'update hashvalues, castling rights, and pawn moved bit
        .MiscBits = .MiscBits And (CastleBits(FromSquare) And CastleBits(ToSquare)) Or (PawnMovedMask And ((MovingPiece = Pawn Or MovingPiece = King) Or (CapturedPiece = Pawn)))
        LostPiece = .Squares(ToSquare) And PieceListIndexMask
        .Squares(ToSquare) = .Squares(FromSquare)
        l = (.Squares(FromSquare) And PieceTypeColorMask) * 8
        j = l Or FromSquare
        l = l Or ToSquare
        'xor FromSquare out-of and ToSquare into hash
        .TTIdx = .TTIdx Xor HashValues(j).Index Xor HashValues(l).Index
        .TTCnf = .TTCnf Xor HashValues(j).Confirm Xor HashValues(l).Confirm
        .Squares(FromSquare) = Free
        PieceNum = .Squares(ToSquare) And PieceListIndexMask
        .Pieces(PieceNum) = ToSquare Or PieceLivingMask
        SideNum = IIf(.SideToMove = White, 0, 1)
        If CapturedPiece <> Free Then  'capture
            .Material(1 - SideNum) = .Material(1 - SideNum) - PieceValues(LostPiece)
            .Pieces(CapturedPiece And PieceListIndexMask) = .Pieces(CapturedPiece And PieceListIndexMask) And Not PieceLivingMask
            If Paint Then
                'visibly remove the captured piece
                imPieces(CapturedPiece And PieceListIndexMask).Visible = False
                DoEvents
                Sleep 200
            End If
            CapturedPiece = CapturedPiece \ 16
            j = CapturedPiece * 128 Or ToSquare
            'xor captured-piece out of hash
            .TTIdx = .TTIdx Xor HashValues(j).Index
            .TTCnf = .TTCnf Xor HashValues(j).Confirm
        End If
        If Len(Move) = 3 Then
            NewPiece = Asc(Right$(Move, 1))
            If NewPiece Then 'promotion
                .Squares(ToSquare) = .Squares(ToSquare) And Not PieceTypeMask Or NewPiece
                If Paint Then 'create the new piece image
                    imPieces(.Squares(ToSquare) And PieceListIndexMask).Visible = False
                    Select Case NewPiece
                      Case WhiteQueen
                        Set imPieces(.Squares(ToSquare) And PieceListIndexMask) = imWhiteQueen(WhiteQueensUsed)
                        WhiteQueensUsed = WhiteQueensUsed + 1
                      Case WhiteRook
                        Set imPieces(.Squares(ToSquare) And PieceListIndexMask) = imWhiteRook(WhiteRooksUsed)
                        WhiteRooksUsed = WhiteRooksUsed + 1
                      Case WhiteBishop
                        Set imPieces(.Squares(ToSquare) And PieceListIndexMask) = imWhiteBishop(WhiteBishopsUsed)
                        WhiteBishopsUsed = WhiteBishopsUsed + 1
                      Case WhiteKnight
                        Set imPieces(.Squares(ToSquare) And PieceListIndexMask) = imWhiteKnight(WhiteKnightsUsed)
                        WhiteKnightsUsed = WhiteKnightsUsed + 1
                      Case BlackQueen
                        Set imPieces(.Squares(ToSquare) And PieceListIndexMask) = imBlackQueen(BlackQueensUsed)
                        BlackQueensUsed = BlackQueensUsed + 1
                      Case BlackRook
                        Set imPieces(.Squares(ToSquare) And PieceListIndexMask) = imBlackRook(BlackRooksUsed)
                        BlackRooksUsed = BlackRooksUsed + 1
                      Case BlackBishop
                        Set imPieces(.Squares(ToSquare) And PieceListIndexMask) = imBlackBishop(BlackBishopsUsed)
                        BlackBishopsUsed = BlackBishopsUsed + 1
                      Case BlackKnight
                        Set imPieces(.Squares(ToSquare) And PieceListIndexMask) = imBlackKnight(BlackKnightsUsed)
                        BlackKnightsUsed = BlackKnightsUsed + 1
                    End Select
                    imPieces(.Squares(ToSquare) And PieceListIndexMask).Visible = True
                End If
                j = NewPiece * 8 Or ToSquare
                l = ((Pawn Or .SideToMove) * 8) Or ToSquare
                'xor pawn out-of and promoted piece into hash
                .TTIdx = .TTIdx Xor HashValues(j).Index Xor HashValues(l).Index
                .TTCnf = .TTCnf Xor HashValues(j).Confirm Xor HashValues(l).Confirm
                NewPiece = NewPiece And PieceTypeMask
                PieceValues(.Squares(ToSquare) And PieceListIndexMask) = Switch(NewPiece = Queen, QueenValue, NewPiece = Rook, RookValue, NewPiece = Bishop, BishopValue, NewPiece = Knight, KnightValue)
                .Material(SideNum) = .Material(SideNum) + PieceValues(.Squares(ToSquare) And PieceListIndexMask) - PawnValue
            End If
        End If
        If Paint Then
            MovePieceImage .Squares(ToSquare) And PieceListIndexMask, ToSquare
            i = .Pieces(.SideToMove Xor Black) And PieceSquareMask
            If Mode <> "CC" Then
                If IsAttacked(Board, i, .SideToMove) Then
                    lbCheck.Move TLX(i), TLY(i) + 15
                    lbCheck.Visible = True
                End If
            End If
        End If
    End With 'BOARD

End Sub

Private Sub mhFocus_MsgReceived(ByVal hWnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long, lResult As Long)

  'Consume WM_SETFOCUS message

    lResult = 0

End Sub

Private Sub MovePieceImage(PieceNum As Long, Square As Long)

    imPieces(PieceNum).Move TLX(Square) + 8, TLY(Square) + 8

End Sub

Private Sub opComp_Click()

  'set up options

    fr(2).Enabled = True
    lbTime.ForeColor = vbBlack
    If opComp Then
        If opPlaySelf Then
            Mode = "CC"
          Else 'OPPLAYSELF = FALSE
            If Board.SideToMove = White Then
                Mode = "CH"
                ckReverse = vbChecked
              Else 'NOT BOARD.SIDETOMOVE...
                Mode = "HC"
                ckReverse = vbUnchecked
            End If
        End If
      Else 'OPCOMP = FALSE
        If opPlaySelf Then
            Mode = "HH"
            fr(2).Enabled = False
            lbTime.ForeColor = lbTime.BackColor
          Else 'OPPLAYSELF = FALSE
            If Board.SideToMove = White Then
                Mode = "HC"
                ckReverse = vbUnchecked
              Else 'NOT BOARD.SIDETOMOVE...
                Mode = "CH"
                ckReverse = vbChecked
            End If
        End If
    End If

End Sub

Private Sub opHuman_Click()

    opComp_Click

End Sub

Private Sub opPlayAlt_Click()

    opComp_Click

End Sub

Private Sub opPlaySelf_Click()

    opComp_Click

End Sub

Private Sub PlayGame()

  'the main game loop

    With Board
        GameEnds = False
        Do Until GameEnds
            lbSide.BackColor = IIf(.SideToMove = White, WhitesTurnColor, BlacksTurnColor)
            lbScore.ForeColor = IIf(.SideToMove = Black, WhitesTurnColor, BlacksTurnColor)
            Select Case Mode
              Case "CC"
                CurrPV(0, 1) = "" 'invalidate PV in CC mode
                ComputerMoves
              Case "HH"
                HumanMoves
              Case "CH"
                If .SideToMove = White Then
                    ComputerMoves
                  Else 'NOT .SIDETOMOVE...
                    HumanMoves
                End If
              Case "HC"
                If .SideToMove = White Then
                    HumanMoves
                  Else 'NOT .SIDETOMOVE...
                    ComputerMoves
                End If
            End Select
            If GameEnds = 0 Then 'switch color only if game in progress
                .SideToMove = .SideToMove Xor Black
            End If
        Loop
    End With 'BOARD
    'game finished
    lbYMM = ""
    If GameEnds <> ByUser Then 'game ended normally
        lbMoves.ForeColor = NeutralColor
        lbMoves = "Game ends"
        lbSide.BackColor = BackColor
    End If
    btNewGame.Enabled = True
    btEdit.Enabled = True
    picEinst.Visible = False
    MousePointer = vbNormal

End Sub

Private Sub ResetClocks()

    tmrElapsed.Enabled = False
    WhiteTime = 0
    BlackTime = 0
    lblClock(0) = "00:00:00"
    lblClock(1) = "00:00:00"

End Sub

Private Sub scrTimeToThink_Change()

    lbTime.BackColor = &HE0E0E0
    lbTime.ForeColor = vbBlack
    Select Case scrTimeToThink
      Case Is < 10
        lbTime = scrTimeToThink * 6 & " Seconds"
      Case 10
        lbTime = "1 Minute"
      Case 51
        lbTime = "Unlimited"
        lbTime.BackColor = vbYellow
        lbTime.ForeColor = vbRed
      Case Else
        lbTime = scrTimeToThink / 10 & " Minutes"
    End Select

End Sub

Private Sub scrTimeToThink_Scroll()

    scrTimeToThink_Change

End Sub

Private Function Search(Board As Board, ByVal Ply, ByVal MinPly, Alpha, Beta) As Long

  'the search function (find the best move in the current situation)

  Dim Piece
  Dim BestResult
  Dim TempBoard     As Board
  Dim t             As Boolean
  Dim CurrMove      As String

    If Ply > MaxPlySearched Then
        MaxPlySearched = Ply
    End If
    If (PosnsVisited And 1023&) = 0 Then
        DoEvents
    End If
    PosnsVisited = PosnsVisited + 1
    With Board
        .MoveCategory = 2
        If InPrinVar Then
            If Len(PrevPV(Ply)) Then
                .MoveCategory = 1
              Else 'LEN(PREVPV(PLY)) = FALSE
                InPrinVar = False
            End If
        End If
        .BestCapture = GenerateMoves(Board, False)
        CurrPV(Ply, Ply) = ""
        If .BestCapture Or .QuietMoveListFrom <= .QuietMoveListTo Then 'Moves generated
            If .BestCapture = -1 Then 'could capture king
                BestResult = Infinity - Ply + 1
              ElseIf .BestCapture = 0 And Ply = MinPly Then 'no captures and at horizon'NOT .BESTCAPTURE...
                BestResult = Evaluate(Board, True)
              Else 'NOT .BESTCAPTURE...
                If Ply > MinPly Then 'beyond horizon - quiescence search
                    BestResult = Evaluate(Board, True)
                  Else 'NOT PLY...
                    BestResult = -Infinity
                End If
                If Ply < MinPly + 4 Then 'temporary fwd pruning ##
                    Do
                        'move ordering: get next move to analyze
                        On .MoveCategory GoSub GetPrinvar, GetKiller1, GetKiller2, GetBestCapture, GetCaptures, GetQuiets
                        If Len(CurrMove) > 3 Then  'castling move
                            t = (CastlingIsLegal(Board, Len(CurrMove)))
                          Else 'NOT LEN(CURRMOVE)...
                            t = Len(CurrMove)
                        End If
                        If t Then 'that's a legal move worth analyzing
                            TempBoard = Board
                            MakeMove TempBoard, CurrMove, False
                            With TempBoard
                                .CaptureMoveListFrom = .CaptureMoveListTo + 1
                                .QuietMoveListFrom = .QuietMoveListTo + 1
                                .SideToMove = .SideToMove Xor Black
                                Result = -Search(TempBoard, Ply + 1, MinPly, -Beta, -Alpha) 'recursion
                            End With 'TEMPBOARD
                            If BreakRequested Then
                                Exit Function '>---> Bottom
                            End If
                            If Result > BestResult Then 'found better move
                                BestResult = Result 'update bestresult and pv
                                For i = Ply + 1 To PlyLimit - 1
                                    CurrPV(Ply, i) = CurrPV(Ply + 1, i)
                                    If CurrPV(Ply, i) = "" Then
                                        Exit For '>---> Next
                                    End If
                                Next i
                                CurrPV(Ply, Ply) = CurrMove 'save this move in pv
                                If BestResult >= Beta Then 'cutoff
                                    CutOffs = CutOffs + 1
                                    If .MoveCategory > 4 Then 'save this move - it has produced a cutoff
                                        If CurrMove <> Killer1(Ply) Then
                                            If CurrMove <> Killer2(Ply) Then
                                                Killer2(Ply) = Killer1(Ply)
                                                Killer1(Ply) = CurrMove
                                            End If
                                        End If
                                    End If
                                    Exit Do 'no need to evaluate more moves '>---> Loop
                                  ElseIf BestResult > Alpha Then 'NOT BESTRESULT...
                                    Alpha = BestResult
                                End If 'cutoff
                            End If 'found better move
                        End If 'legal move analyzing
                    Loop While Len(CurrMove)
                    'all moves done with
                    If BestResult = Ply - Infinity Then
                        If Not IsAttacked(Board, .Pieces(.SideToMove) And PieceSquareMask, .SideToMove Xor Black) Then 'king is not attacked
                            BestResult = 0 'stalemate
                            Stalemate = (Ply = 0)
                        End If
                    End If
                End If
            End If 'could capture king
          Else 'no moves generated'NOT .BESTCAPTURE...
            BestResult = 0
            Stalemate = (Ply = 0)
        End If 'moves generated
        Search = BestResult
    End With 'BOARD

Exit Function

    '-------------------------------------------------------------------------------------
    'Local Subs for move ordering

GetPrinvar:
    'MoveCat 1
    Board.MoveCategory = 2
    CurrMove = PrevPV(Ply)
    Return

GetKiller1:
    'MoveCat 2
    Board.MoveCategory = 3
    CurrMove = Killer1(Ply)
    If Len(CurrMove) Then
        For i = Board.CaptureMoveListFrom To Board.CaptureMoveListTo
            If CaptureMoves(i) = CurrMove Then 'found move
                Exit For '>---> Next
            End If
        Next i
        If i <= Board.CaptureMoveListTo Then 'move is in list
            Return
        End If
        For i = Board.QuietMoveListFrom To Board.QuietMoveListTo
            If QuietMoves(i) = CurrMove Then 'found move
                Exit For '>---> Next
            End If
        Next i
        If i <= Board.QuietMoveListTo Then 'move is in list
            Return
        End If
    End If

GetKiller2:
    'MoveCat 3
    Board.MoveCategory = 4
    CurrMove = Killer2(Ply)
    If Len(CurrMove) Then
        For i = Board.CaptureMoveListFrom To Board.CaptureMoveListTo
            If CaptureMoves(i) = CurrMove Then 'found move
                Exit For '>---> Next
            End If
        Next i
        If i <= Board.CaptureMoveListTo Then 'move is in list
            Return
        End If
        For i = Board.QuietMoveListFrom To Board.QuietMoveListTo
            If QuietMoves(i) = CurrMove Then 'found move
                Exit For '>---> Next
            End If
        Next i
        If i <= Board.QuietMoveListTo Then 'move is in list
            Return
        End If
    End If

GetBestCapture:
    'MoveCat 4
    Board.MoveCategory = 5
    Board.CurrMoveIndex = Board.CaptureMoveListFrom - 1
    If Board.BestCapture Then
        CurrMove = CaptureMoves(Board.BestCapture)
        If CurrMove <> Killer1(Ply) And CurrMove <> Killer2(Ply) Then
            Return
        End If
    End If

GetCaptures:
    'MoveCat 5
    Board.CurrMoveIndex = Board.CurrMoveIndex + 1
    If Board.CurrMoveIndex = Board.BestCapture Then
        'skip best capture move
        Board.CurrMoveIndex = Board.CurrMoveIndex + 1
    End If
    If Board.CurrMoveIndex > Board.CaptureMoveListTo Then
        Board.MoveCategory = 6
        Board.CurrMoveIndex = Board.QuietMoveListFrom - 1
      Else 'NOT BOARD.CURRMOVEINDEX...
        CurrMove = CaptureMoves(Board.CurrMoveIndex)
        If CurrMove = Killer1(Ply) Or CurrMove = Killer2(Ply) Then
            GoSub GetCaptures
        End If
        Return
    End If

GetQuiets:
    'MoveCat 6
    Board.CurrMoveIndex = Board.CurrMoveIndex + 1
    If Board.CurrMoveIndex > Board.QuietMoveListTo Or Ply > MinPly Then 'out of moves or quiescence search
        CurrMove = ""
      Else 'NOT BOARD.CURRMOVEINDEX...
        CurrMove = QuietMoves(Board.CurrMoveIndex)
        If CurrMove = Killer1(Ply) Or CurrMove = Killer2(Ply) Then
            GoSub GetQuiets
        End If
    End If
    Return

End Function

Private Sub SquareToLabel(Square As Byte, Color As Long)

  'translate a square index to user friendly notation

    lbMoves.ForeColor = Color
    lbMoves = XN(Square)
    HumanMove = HumanMove & Chr$(Square)

End Sub

Private Sub StandardBoard()

  'set up standard board

    With Board
        For i = 0 To 19
            .Squares(i) = Forbidden
        Next i
        For i = 100 To 119
            .Squares(i) = Forbidden
        Next i
        For i = 20 To 90 Step 10
            For j = 1 To 8
                .Squares(i + j) = Free
            Next j
            .Squares(i) = Forbidden
            .Squares(i + 9) = Forbidden
        Next i
        .Squares(E1) = WhiteKing
        .Squares(D1) = WhiteQueen Or 1
        .Squares(F1) = WhiteBishop Or 2
        .Squares(G1) = WhiteKnight Or 3
        .Squares(H1) = WhiteRook Or 4
        .Squares(C1) = WhiteBishop Or 5
        .Squares(B1) = WhiteKnight Or 6
        .Squares(A1) = WhiteRook Or 7
        .Squares(A2) = WhitePawn Or 8
        .Squares(B2) = WhitePawn Or 9
        .Squares(C2) = WhitePawn Or 10
        .Squares(D2) = WhitePawn Or 11
        .Squares(E2) = WhitePawn Or 12
        .Squares(F2) = WhitePawn Or 13
        .Squares(G2) = WhitePawn Or 14
        .Squares(H2) = WhitePawn Or 15
        .Pieces(0) = E1 Or PieceLivingMask
        .Pieces(1) = D1 Or PieceLivingMask
        .Pieces(2) = F1 Or PieceLivingMask
        .Pieces(3) = G1 Or PieceLivingMask
        .Pieces(4) = H1 Or PieceLivingMask
        .Pieces(5) = C1 Or PieceLivingMask
        .Pieces(6) = B1 Or PieceLivingMask
        .Pieces(7) = A1 Or PieceLivingMask
        .Pieces(8) = A2 Or PieceLivingMask
        .Pieces(9) = B2 Or PieceLivingMask
        .Pieces(10) = C2 Or PieceLivingMask
        .Pieces(11) = D2 Or PieceLivingMask
        .Pieces(12) = E2 Or PieceLivingMask
        .Pieces(13) = F2 Or PieceLivingMask
        .Pieces(14) = G2 Or PieceLivingMask
        .Pieces(15) = H2 Or PieceLivingMask
        .Squares(E8) = BlackKing
        .Squares(D8) = BlackQueen Or 1
        .Squares(F8) = BlackBishop Or 2
        .Squares(G8) = BlackKnight Or 3
        .Squares(H8) = BlackRook Or 4
        .Squares(C8) = BlackBishop Or 5
        .Squares(B8) = BlackKnight Or 6
        .Squares(A8) = BlackRook Or 7
        .Squares(A7) = BlackPawn Or 8
        .Squares(B7) = BlackPawn Or 9
        .Squares(C7) = BlackPawn Or 10
        .Squares(D7) = BlackPawn Or 11
        .Squares(E7) = BlackPawn Or 12
        .Squares(F7) = BlackPawn Or 13
        .Squares(G7) = BlackPawn Or 14
        .Squares(H7) = BlackPawn Or 15
        .Pieces(16) = E8 Or PieceLivingMask
        .Pieces(17) = D8 Or PieceLivingMask
        .Pieces(18) = F8 Or PieceLivingMask
        .Pieces(19) = G8 Or PieceLivingMask
        .Pieces(20) = H8 Or PieceLivingMask
        .Pieces(21) = C8 Or PieceLivingMask
        .Pieces(22) = B8 Or PieceLivingMask
        .Pieces(23) = A8 Or PieceLivingMask
        .Pieces(24) = A7 Or PieceLivingMask
        .Pieces(25) = B7 Or PieceLivingMask
        .Pieces(26) = C7 Or PieceLivingMask
        .Pieces(27) = D7 Or PieceLivingMask
        .Pieces(28) = E7 Or PieceLivingMask
        .Pieces(29) = F7 Or PieceLivingMask
        .Pieces(30) = G7 Or PieceLivingMask
        .Pieces(31) = H7 Or PieceLivingMask
        Set imPieces(0) = imWhiteKing
        Set imPieces(1) = imWhiteQueen(0)
        Set imPieces(2) = imWhiteBishop(0)
        Set imPieces(3) = imWhiteKnight(0)
        Set imPieces(4) = imWhiteRook(0)
        Set imPieces(5) = imWhiteBishop(1)
        Set imPieces(6) = imWhiteKnight(1)
        Set imPieces(7) = imWhiteRook(1)
        For i = 0 To 7
            Set imPieces(i + 8) = imWhitePawn(i)
        Next i
        Set imPieces(16) = imBlackKing
        Set imPieces(17) = imBlackQueen(0)
        Set imPieces(18) = imBlackBishop(0)
        Set imPieces(19) = imBlackKnight(0)
        Set imPieces(20) = imBlackRook(0)
        Set imPieces(21) = imBlackBishop(1)
        Set imPieces(22) = imBlackKnight(1)
        Set imPieces(23) = imBlackRook(1)
        For i = 0 To 7
            Set imPieces(i + 24) = imBlackPawn(i)
        Next i
        WhiteQueensUsed = 1
        BlackQueensUsed = 1
        WhiteRooksUsed = 2
        WhiteBishopsUsed = 2
        WhiteKnightsUsed = 2
        WhitePawnsUsed = 8
        BlackRooksUsed = 2
        BlackBishopsUsed = 2
        BlackKnightsUsed = 2
        BlackPawnsUsed = 8
        .EnPassant = 0
        .MiscBits = WhiteCastleShortMask Or WhiteCastleLongMask Or BlackCastleShortMask Or BlackCastleLongMask
        .QuietMoveListFrom = 0
        .QuietMoveListTo = 0
        .CaptureMoveListFrom = 0
        .CaptureMoveListTo = 0
        MoveCount = 0
        'calculate material
        .Material(0) = 0
        For i = 0 To 15
            .Material(0) = .Material(0) + (PieceValues(i))
        Next i
        .Material(1) = 0
        For i = 16 To 31
            .Material(1) = .Material(1) + (PieceValues(i))
        Next i
        opHuman = True
        ckReverse_Click
        ResetClocks
        lbYMM = ""
    End With 'BOARD
    CreateHash

End Sub

Private Function TLX(Square As Long) As Long

  'returns the top left x coordinate of square

    If ckReverse = vbChecked Then
        TLX = (9 - (Square Mod 10)) * 64
      Else 'NOT CKREVERSE...
        TLX = (Square Mod 10) * 64
    End If

End Function

Private Function TLY(Square As Long) As Long

  'returns the top left y coordinate of square

    If ckReverse = vbChecked Then
        TLY = ((Square \ 10) - 1) * 64
      Else 'NOT CKREVERSE...
        TLY = (10 - (Square \ 10)) * 64
    End If

End Function

Private Sub tmr_Timer()

  'sponge timer - reset some displays

    lbMsg = ""
    lbCheck.Visible = False
    tmr.Enabled = False

End Sub

Private Sub tmrElapsed_Timer()

  'feed the clocks

    If Board.SideToMove = White Then
        WhiteTime = Now - GameStart - BlackTime
        lblClock(0) = Format$(WhiteTime, "hh:mm:ss")
      Else 'NOT BOARD.SIDETOMOVE...
        BlackTime = Now - GameStart - WhiteTime
        lblClock(1) = Format$(BlackTime, "hh:mm:ss")
    End If

End Sub

Private Sub UnHiliteSquare(ByVal Square As Long)

  'remove frame around square

  Dim Color As Long

    X = TLX(Square)
    Y = TLY(Square)
    Color = Point(X + 60, Y + 60)
    DrawStyle = vbSolid
    Line (X + 1, Y + 1)-Step(61, 0), Color
    Line -Step(0, 61), Color
    Line -Step(-61, 0), Color
    Line -Step(0, -61), Color

End Sub

Private Function XlatMove(Move As String, Full As Boolean) As String

  'translate an internal move to user friendly notation

  Dim Conn        As String
  Dim Suffix      As String
  Dim PromTo      As Byte

    Select Case Len(Move)
      Case 0
        XlatMove = "None"
      Case 5
        XlatMove = "0-0-0"
      Case 4
        XlatMove = "0-0"
      Case Else
        If (Len(Move) = 3) And Full Then
            Conn = "x"
            If Right$(Move, 1) = Chr$(0) Then
                If (Board.Squares(Asc(Mid$(Move, 2))) And PieceTypeMask) = Free Then
                    Suffix = "ep"
                End If
              Else 'NOT RIGHT$(MOVE,...
                If (Board.Squares(Asc(Mid$(Move, 2))) And PieceTypeMask) = Free Then
                    Conn = "-"
                End If
                PromTo = Asc(Right$(Move, 1)) And PieceTypeMask
                Suffix = Switch(PromTo = Queen, "Q", PromTo = Rook, "R", PromTo = Bishop, "B", PromTo = Knight, "N")
            End If
          Else 'NOT (LEN(MOVE)...
            Conn = "-"
        End If
        XlatMove = XN(Asc(Left$(Move, 1))) & Conn & XN(Asc(Mid$(Move, 2))) & Suffix
    End Select

End Function

Private Function XN(Square As Byte) As String

  'translate square index

    XN = Chr$(File(Square) + 64) & Rank(Square)

End Function

':) Ulli's VB Code Formatter V2.9.4 (22.01.2002 08:17:27) 410 + 2958 = 3368 Lines
