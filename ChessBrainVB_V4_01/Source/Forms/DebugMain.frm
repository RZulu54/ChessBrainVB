VERSION 5.00
Begin VB.Form frmDebugMain 
   Caption         =   "ChessBrainVB debug console"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14715
   Icon            =   "DebugMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   14715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRunUCI 
      Caption         =   "Calc UCI-Pos"
      Height          =   330
      Left            =   4200
      TabIndex        =   10
      Top             =   600
      Width           =   1305
   End
   Begin VB.TextBox txtUciPosition 
      Height          =   285
      Left            =   5640
      TabIndex        =   9
      Text            =   $"DebugMain.frx":0442
      Top             =   600
      Width           =   8535
   End
   Begin VB.CommandButton cmdThink 
      Caption         =   "Think"
      Height          =   330
      Left            =   1800
      TabIndex        =   8
      Top             =   600
      Width           =   1425
   End
   Begin VB.CommandButton cmdNewgame 
      Caption         =   "New game"
      Height          =   330
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1425
   End
   Begin VB.CommandButton cmdTx 
      Caption         =   "f1xf4"
      Height          =   330
      Left            =   13200
      TabIndex        =   6
      Top             =   240
      Width           =   945
   End
   Begin VB.CommandButton cmdT2 
      Caption         =   "g4xh6"
      Height          =   330
      Left            =   11880
      TabIndex        =   5
      Top             =   240
      Width           =   945
   End
   Begin VB.CommandButton cmdTest1 
      Caption         =   "e3e8+ M8"
      Height          =   330
      Left            =   10440
      TabIndex        =   4
      Top             =   240
      Width           =   1065
   End
   Begin VB.CommandButton cmdFakeInput 
      Caption         =   "Send"
      Height          =   330
      Left            =   9000
      TabIndex        =   1
      Top             =   240
      Width           =   1065
   End
   Begin VB.ComboBox cboFakeInput 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   240
      Width           =   8796
   End
   Begin VB.TextBox txtIO 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8892
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   14310
   End
   Begin VB.Label lblDescr 
      BackStyle       =   0  'Transparent
      Caption         =   "Input"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmDebugMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'= frmDebugMain:
'= debug form
'==================================================
Option Explicit


Private Sub cboFakeInput_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then cmdFakeInput_Click
End Sub

Private Sub cmdFakeInput_Click()
  FakeInput = cboFakeInput.Text & vbLf
  FakeInputState = True
  cboFakeInput.SelStart = 0
  cboFakeInput.SelLength = Len(cboFakeInput.Text)
  cboFakeInput.SetFocus
  
  UCIMode = False
 ' pbIsOfficeMode = True ' TEst
End Sub

Private Sub cmdWb_Click()
  bPostMode = True
  ParseCommand "setboard r1b2rk1/pp1n2pp/2p1p3/2Pp4/1q1Pp3/4P1PN/PP2QPBP/2R2RK1 w - - 0 15" & vbLf
  ParseCommand "sd 10" & vbLf
  ParseCommand "go" & vbLf
  
End Sub

Private Sub cmdNewgame_Click()
 UCIMode = True
 cboFakeInput.Text = "ucinewgame"
 cmdFakeInput_Click
End Sub

Private Sub cmdRunUCI_Click()
UCIPositionSetup "position fen r1bqk2r/pp1nbppp/2p1p3/3n4/4N3/3P1NP1/PPP1QPBP/R1B1K2R w KQkq - 0 1 moves e1g1 e8g8 c2c4 d5f6 e4c3 e6e5 f3e5 d7e5 e2e5 f8e8 d3d4 e7b4 e5f4 b4c3 b2c3 c8e6 f1e1 e6c4 e1e8 d8e8 c1d2 e8d7 a2a4 f6d5 f4e4 a8e8 e4c2 d5f6 c2b2 g7g6 a4a5 f6g4 g2f3 g4f6 b2b4 d7f5 f3g2 f5d3 d2e1 d3e2 h2h3 a7a6 b4b1 e8e7 b1d1 e2e6 d1d2 g8g7 d2f4 c4d5 f2f3 e6c8 g3g4 d5b3 f4d6 f6d5 d6a3 b3c4 a3c5 e7e2 g2f1 c8e8 f1e2 e8e2 c5d6 e2f3 g4g5 f3f1 g1h2 h7h6 g5h6 g7h7 d6e5 f7f6 e5e4 d5f4 e4e7 h7h6 e7f8 h6h7 f8e7 h7g8 e7d8 g8f7 d8d7 f7f8 d7d8 f8g7 d8e7 c4f7 e7f6 g7g8 f6d8 g8g7 d8f6 g7g8 f6d8"
FixedDepth = 15: MovesToTC = 0: TimeLeft = 20: TimeIncrement = 10: bPostMode = True
'--- start computing --------------
StartEngine

End Sub

Private Sub cmdT2_Click()
  cboFakeInput.Text = "bench 21"
  TestStart = 8
  TestEnd = 8
End Sub

Private Sub cmdTest1_Click()
  cboFakeInput.Text = "bench 23"
  TestStart = 1
  TestEnd = 1
End Sub


Private Sub cmdThink_Click()
 cboFakeInput.Text = "go"
 cmdFakeInput_Click
End Sub

Private Sub cmdTx_Click()
  cboFakeInput.Text = "bench 21"
  TestStart = 2
  TestEnd = 2
End Sub

Private Sub Form_Load()
  'txtIO = "* STDIN HANDLE: " & hStdIn & vbTab & "STDOUT HANDLE: " & hStdOut & " *" & vbCrLf
  txtIO = ""
  cboFakeInput = "bench 14"
  cboFakeInput.AddItem "analyze"
  cboFakeInput.AddItem "eval"  ' input in Immediate window and Tracexxx.txt
  cboFakeInput.AddItem "bench 6"
  cboFakeInput.AddItem "writeepd"
  cboFakeInput.AddItem "display"
  cboFakeInput.AddItem "list"
  cboFakeInput.AddItem "new"
  cboFakeInput.AddItem "setboard 1b5k/7P/p1p2np1/2P2p2/PP3P2/4RQ1R/q2r3P/6K1 w - - 0 1"
  cboFakeInput.AddItem "setboard r1b2rk1/pp1nq1p1/2p1p2p/3p1p2/2PPn3/2NBPN2/PPQ2PPP/2R2RK1 b - -"
  cboFakeInput.AddItem "setboard 2br2k1/ppp2p1p/4p1p1/4P2q/2P1Bn2/2Q5/PP3P1P/4R1RK b - -"
  cboFakeInput.AddItem "setboard 8/8/R3k3/1R6/8/8/8/2K5 b - -"
  cboFakeInput.AddItem "setboard 2k4r/1pr1n3/p1p1q2p/5pp1/3P1P2/P1P1P3/1R2Q1PP/1RB3K1 w KQkq -"
  cboFakeInput.AddItem "setboard 6k1/1b1nqpbp/pp4p1/5P2/1PN5/4Q3/P5PP/1B2B1K1 b - -"
  cboFakeInput.AddItem "perft 3"
  cboFakeInput.AddItem "xboard" & vbLf & "new" & vbLf & "random" & vbLf & "level 40 5 0" & vbLf & "post"
  cboFakeInput.AddItem "xboard" & vbLf & "new" & vbLf & "random" & vbLf & "sd 4" & vbLf & "post"
  cboFakeInput.AddItem "time 30000" & vbLf & "otim 30000" & vbLf & "e2e4"
  cboFakeInput.AddItem "force" & vbLf & "quit"
  cboFakeInput.AddItem "setboard rnbqkbnr/ppp2ppp/4p3/3pP3/3P4/8/PPP2PPP/RNBQKBNR b KQkq -"
  cboFakeInput.AddItem "setboard 8/p1b1k1p1/Pp4p1/1Pp2pPp/2P2P1P/3B1K2/8/8 w - -"
  cboFakeInput.AddItem "setboard 8/2R5/1r3kp1/2p4p/2P2P2/p3K1P1/P6P/8 w - -"
  cboFakeInput.AddItem "setboard 7k/p7/6K1/5Q2/8/8/8/8 w - -"
  cboFakeInput.AddItem "debug1"
  DebugMode = True
  cmdTest1_Click
  UCIMode = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ExitProgram
End Sub

Private Sub Form_Resize()
  On Local Error Resume Next

'  With txtIO
'    .Move .Left, .Top, Me.ScaleWidth - (.Left * 2), Me.ScaleHeight - 800
'  End With
'
'  cboFakeInput.Width = txtIO.Width - cmdFakeInput.Width - 100
'  cmdFakeInput.Left = cboFakeInput.Left + cboFakeInput.Width + 100
'  On Local Error GoTo 0
End Sub

