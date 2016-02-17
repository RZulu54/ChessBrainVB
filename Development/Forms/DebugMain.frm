VERSION 5.00
Begin VB.Form frmDebugMain 
   Caption         =   "ChessBrainVB debug console"
   ClientHeight    =   9960
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   14712
   Icon            =   "DebugMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9960
   ScaleWidth      =   14712
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdFakeInput 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   330
      Left            =   9000
      TabIndex        =   1
      Top             =   288
      Width           =   1065
   End
   Begin VB.ComboBox cboFakeInput 
      Height          =   288
      Left            =   105
      TabIndex        =   0
      Top             =   315
      Width           =   8796
   End
   Begin VB.TextBox txtIO 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   8892
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   735
      Width           =   14310
   End
   Begin VB.Label lblDescr 
      BackStyle       =   0  'Transparent
      Caption         =   "Input"
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   3
      Top             =   105
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

Private Sub cmdFakeInput_Click()

  FakeInput = cboFakeInput.Text & vbLf
  FakeInputState = True
  cboFakeInput.SelStart = 0
  cboFakeInput.SelLength = Len(cboFakeInput.Text)
  cboFakeInput.SetFocus

End Sub
Private Sub Form_Load()

  'txtIO = "* STDIN HANDLE: " & hStdIn & vbTab & "STDOUT HANDLE: " & hStdOut & " *" & vbCrLf
  txtIO = ""
  cboFakeInput = "bench 8"
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  ExitProgram

End Sub

Private Sub Form_Resize()

  On Local Error Resume Next

  With txtIO
    .Move .Left, .Top, Me.ScaleWidth - (.Left * 2), Me.ScaleHeight - 800
  End With

  cboFakeInput.Width = txtIO.Width - cmdFakeInput.Width - 100
  cmdFakeInput.Left = cboFakeInput.Left + cboFakeInput.Width + 100
  On Local Error GoTo 0

End Sub

