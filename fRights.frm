VERSION 5.00
Begin VB.Form fRights 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Enter additional data"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   ControlBox      =   0   'False
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton opBlack 
      BackColor       =   &H00808080&
      Height          =   195
      Left            =   345
      TabIndex        =   12
      ToolTipText     =   "...has first move"
      Top             =   1635
      Width           =   195
   End
   Begin VB.OptionButton opWhite 
      BackColor       =   &H00808080&
      Height          =   195
      Left            =   345
      TabIndex        =   11
      ToolTipText     =   "...has first move"
      Top             =   180
      Width           =   195
   End
   Begin VB.Frame fr 
      BackColor       =   &H00808080&
      Caption         =   "    Black..."
      BeginProperty Font 
         Name            =   "Amaze"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1320
      Index           =   1
      Left            =   210
      TabIndex        =   10
      Top             =   1590
      Width           =   3060
      Begin VB.CheckBox ckCastleQB 
         Alignment       =   1  'Rechts ausgerichtet
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   315
         Width           =   2640
      End
      Begin VB.CheckBox ckCastleKB 
         Alignment       =   1  'Rechts ausgerichtet
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   630
         Width           =   2640
      End
      Begin VB.TextBox txEPB 
         Height          =   285
         Left            =   2595
         MaxLength       =   2
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   870
         Width           =   315
      End
      Begin VB.Label lb 
         BackColor       =   &H00808080&
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   930
         Width           =   2160
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00808080&
      Caption         =   "    White..."
      BeginProperty Font 
         Name            =   "Amaze"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1320
      Index           =   0
      Left            =   210
      TabIndex        =   9
      Top             =   135
      Width           =   3060
      Begin VB.TextBox txEPW 
         Height          =   285
         Left            =   2595
         MaxLength       =   2
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   870
         Width           =   315
      End
      Begin VB.CheckBox ckCastleKW 
         Alignment       =   1  'Rechts ausgerichtet
         BackColor       =   &H00808080&
         Caption         =   "...has right to castle &King Side"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   630
         Width           =   2640
      End
      Begin VB.CheckBox ckCastleQW 
         Alignment       =   1  'Rechts ausgerichtet
         BackColor       =   &H00808080&
         Caption         =   "...has right to castle &Queen Side"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   315
         Width           =   2640
      End
      Begin VB.Label lb 
         BackColor       =   &H00808080&
         Caption         =   "...could capture &en passant on"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   930
         Width           =   2160
      End
   End
   Begin VB.CommandButton btOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   450
      Left            =   1208
      TabIndex        =   8
      Top             =   3090
      Width           =   1065
   End
End
Attribute VB_Name = "fRights"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Square           As Long
Public SetFocusOnText   As Boolean

Private Sub btOK_Click()

    Hide
    
End Sub

Private Function ConvertSquare(Box As TextBox, Rank As String) As Long

    If Len(Box) = 2 Then
        If Left$(Box, 1) Like "[A-H]" And Right$(Box, 1) = Rank Then
            ConvertSquare = Val(Rank) * 10 + 10 + InStr("ABCDEFGH", Left$(Box, 1))
          Else 'NOT LEFT$(BOX,...
            ConvertSquare = -1
        End If
    End If
    Box.SelStart = 0
    Box.SelLength = 2

End Function

Private Sub Form_Load()
    
    Move fChess.Left + (fChess.Width - Width) / 2, fChess.Top + (fChess.Height - Height) / 2
    SetFocusOnText = False
    lb(0) = lb(1)
    ckCastleQB.Caption = ckCastleQW.Caption
    ckCastleKB.Caption = ckCastleKW.Caption

End Sub

Private Sub Form_Paint()

    Square = 0
    If SetFocusOnText Then
        If txEPB.Enabled Then
            txEPB.SelStart = 0
            txEPB.SelLength = 2
            txEPB.SetFocus
          Else 'TXEPB.ENABLED = FALSE
            txEPW.SelStart = 0
            txEPW.SelLength = 2
            txEPW.SetFocus
        End If
    End If

End Sub

Private Sub opBlack_Click()
    
    txEPB.Enabled = opBlack
    txEPW.Enabled = Not opBlack
    txEPW = ""
    
End Sub

Private Sub opWhite_Click()

    txEPW.Enabled = opWhite
    txEPB.Enabled = Not opWhite
    txEPB = ""
    
End Sub

Private Sub txEPB_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub txEPB_Validate(Cancel As Boolean)

    Square = ConvertSquare(txEPB, "3")
    Cancel = (Square < 0)

End Sub

Private Sub txEPW_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub txEPW_Validate(Cancel As Boolean)

    Square = ConvertSquare(txEPW, "6")
    Cancel = (Square < 0)
    
End Sub

':) Ulli's VB Code Formatter V2.9.4 (22.01.2002 08:16:52) 4 + 91 = 95 Lines
