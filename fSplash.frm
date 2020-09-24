VERSION 5.00
Object = "{A034B639-50EC-11D4-B07A-FBBD7E43DB02}#9.0#0"; "GRADIENT.OCX"
Begin VB.Form fSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3825
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Image imWhiteKing 
      Appearance      =   0  '2D
      Height          =   720
      Left            =   1665
      Picture         =   "fSplash.frx":0000
      Top             =   930
      Width           =   720
   End
   Begin VB.Image imBlackKing 
      Appearance      =   0  '2D
      Height          =   720
      Left            =   2220
      Picture         =   "fSplash.frx":1CCA
      Top             =   405
      Width           =   720
   End
   Begin VB.Image imgUMGEDV 
      BorderStyle     =   1  'Fest Einfach
      Height          =   765
      Left            =   210
      Picture         =   "fSplash.frx":3994
      Top             =   615
      Width           =   825
   End
   Begin VB.Label lblVers 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00C000C0&
      Height          =   165
      Left            =   1005
      TabIndex        =   1
      Top             =   1680
      Width           =   1800
   End
   Begin GradientOCX.Gradient gra 
      Left            =   75
      Top             =   2205
      _ExtentX        =   529
      _ExtentY        =   503
      FromColor       =   11579647
      ToColor         =   16756991
      ColorSequence   =   1
   End
   Begin VB.Label lbl 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ulli's Chess Program"
      BeginProperty Font 
         Name            =   "Amaze"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H006040C0&
      Height          =   525
      Left            =   180
      TabIndex        =   0
      Top             =   -15
      Width           =   3405
   End
End
Attribute VB_Name = "fSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    lblVers = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

Private Sub Form_Paint()

    gra.Paint

End Sub

':) Ulli's VB Code Formatter V2.9.4 (22.01.2002 08:16:53) 1 + 14 = 15 Lines
