VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ChessBrain VB"
   ClientHeight    =   3435
   ClientLeft      =   2580
   ClientTop       =   1920
   ClientWidth     =   5250
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "In option General->Commandline parameters please add   -xboard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   2
      Left            =   1800
      TabIndex        =   7
      ToolTipText     =   "GNU General Public License"
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   2850
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDescr 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "based on engines: LarsenVB (by Luca Dormio) and Faile (by Adrien M. Regimbald) / Stockfish"
      Height          =   390
      Index           =   4
      Left            =   795
      TabIndex        =   6
      ToolTipText     =   "GNU General Public License"
      Top             =   615
      UseMnemonic     =   0   'False
      Width           =   3525
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblDescr 
      BackStyle       =   0  'Transparent
      Caption         =   "ChessBrainVB 3.70"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   840
      TabIndex        =   5
      Top             =   210
      UseMnemonic     =   0   'False
      Width           =   3405
   End
   Begin VB.Label lblDescr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please use a winboard chess GUI (i.e. ARENA)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   1692
      TabIndex        =   4
      ToolTipText     =   "GNU General Public License"
      Top             =   1332
      UseMnemonic     =   0   'False
      Width           =   2256
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCmd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   252
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2700
      Width           =   432
   End
   Begin VB.Label lblCmd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book editor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   216
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1752
      Width           =   1212
   End
   Begin VB.Label lblCmd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Play game  "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   216
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1332
      Width           =   1224
   End
   Begin VB.Label lblDescr 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright: GNU GENERAL PUBLIC LICENSE V3"
      Height          =   330
      Index           =   1
      Left            =   1440
      TabIndex        =   0
      ToolTipText     =   "GNU General Public License"
      Top             =   3120
      UseMnemonic     =   0   'False
      Width           =   3600
   End
   Begin VB.Image imgIco 
      Height          =   480
      Left            =   105
      Top             =   105
      Width           =   480
   End
   Begin VB.Image imgPointer 
      Height          =   480
      Left            =   735
      Picture         =   "Main.frx":0442
      Top             =   105
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==================================================
'= frmMain:
'= Main form ( not shown under winboard)
'==================================================
Option Explicit
Private sWBPath As String   'path winboard.exe

Private Function BrowseForFolders() As String
  BrowseForFolders = InputBox("Enter path of Winboard.exe (or edit INI file):")
  ' WinAPI removed to avoid problems with missing reference
  '###WIN32 sTitle = StrConv("Select location of winboard.exe (or use ARENA GUI) :", vbFromUnicode)
  '###WIN32 BInfo.hwndOwner = Me.hWnd
  '###WIN32 BInfo.lpszTitle = StrPtr(sTitle)
  '###WIN32 BInfo.ulFlags = BIF_RETURNONLYFSDIRS
  '###WIN32 lpIdList = SHBrowseForFolder(BInfo)
  '###WIN32 If lpIdList Then
  '###WIN32     sFolderName = String$(260, 0)
  '###WIN32     SHGetPathFromIDList lpIdList, sFolderName
  '###WIN32     sFolderName = Left$(sFolderName, InStr(sFolderName, Chr(0)) - 1)
  '###WIN32     CoTaskMemFree lpIdList
  '###WIN32 End If
  '###WIN32 BrowseForFolders = sFolderName
End Function

'---------------------------------------------------------------------------
'GetCmdLine() - pass command line to ChessBrainVB
'
'---------------------------------------------------------------------------
Private Function GetCmdLine() As String
 ' GetCmdLine = " -cp -fcp ""ChessBrainVB -xboard"" -fd """ & psEnginePath & """  -scp ""ChessBrainVB -xboard"" -sd """ & psEnginePath & """"
End Function

Private Sub SetWBPath()
  sWBPath = ReadINISetting("WINBOARD", "")
  If sWBPath = "" Then
    sWBPath = BrowseForFolders
  Else
    On Local Error Resume Next
    If Dir$(sWBPath & "\winboard.exe") = "" Then
      sWBPath = BrowseForFolders
    End If
    On Local Error GoTo 0
  End If
End Sub

Private Sub Form_Load()
  Dim i As Long
  imgIco.Picture = Me.Icon
  Set Me.Icon = Nothing

  With App
    Me.Caption = Me.Caption & "   ver. " & .Major & "." & Format(.Minor, "00") & "." & Format(.Revision, "0000")
    'lblDescr(1) = .LegalCopyright
  End With

  For i = 0 To lblCmd.UBound
    lblCmd(i).MouseIcon = imgPointer.Picture
  Next

  'lblDescr(0) = LoadResString(resMainTitle)
  'lblDescr(2) = LoadResString(resMainOption)
  'lblCmd(0) = LoadResString(resMainPlay)
  'lblCmd(1) = LoadResString(resMainBookEd)
  'lblCmd(2) = LoadResString(resMainQuit)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Long

  For i = 0 To lblCmd.UBound
    lblCmd(i).Font.Underline = False
    lblCmd(i).Font.Bold = False
  Next

End Sub

Private Sub lblCmd_Click(Index As Integer)
  Dim sBookName As String

  Select Case Index
    Case 0      'Winboard
      SetWBPath
      On Local Error GoTo CmdError:
      Shell sWBPath & "\winboard.exe" & GetCmdLine, vbNormalFocus
      WriteINISetting "WINBOARD", sWBPath
      End
    Case 1      'BookEdit
      Screen.MousePointer = vbHourglass
      sBookName = ReadINISetting(USE_BOOK_KEY, "")
      On Local Error GoTo CmdError:
      If sBookName = "" Then
        Shell psEnginePath & "\BookEdit.exe", vbNormalFocus
      Else
        Shell psEnginePath & "\BookEdit.exe " & psEnginePath & "\" & sBookName, vbNormalFocus
      End If
      Screen.MousePointer = vbDefault
      End
    Case 2      'Quit
      End
  End Select

  On Local Error GoTo 0
  Exit Sub
CmdError:

  Select Case Err.Number
    Case 53

      Select Case Index
        Case 0
          MsgBox "Cannot find Winboard", vbCritical
        Case 1
          MsgBox "Cannot find  BookEdit", vbCritical
      End Select
  End Select

  Screen.MousePointer = vbDefault
End Sub

Private Sub lblCmd_MouseMove(Index As Integer, _
                             Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)
  Dim i            As Long
  Static LastIndex As Long
  If Index <> LastIndex Then

    For i = 0 To lblCmd.UBound
      lblCmd(i).Font.Underline = False
      lblCmd(i).Font.Bold = False
    Next

    LastIndex = Index
  End If
  lblCmd(Index).Font.Underline = True
  lblCmd(Index).Font.Bold = True
End Sub
