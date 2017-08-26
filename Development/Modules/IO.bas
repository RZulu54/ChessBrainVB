Attribute VB_Name = "IObas"
'==================================================
'= IOBas:
'= Winboard communication / output of think results
'==================================================
Option Explicit


'--- Win32 API functions
Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long

Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Declare Function PeekNamedPipe _
        Lib "kernel32" (ByVal hNamedPipe As Long, _
                        lpBuffer As Any, _
                        ByVal nBufferSize As Long, _
                        lpBytesRead As Long, _
                        lpTotalBytesAvail As Long, _
                        lpBytesLeftThisMessage As Long) As Long
                        
Declare Function ReadFile _
        Lib "kernel32" (ByVal hFile As Long, _
                        lpBuffer As Any, _
                        ByVal nNumberOfBytesToRead As Long, _
                        lpNumberOfBytesRead As Long, _
                        lpOverlapped As Any) As Long
                        
Declare Function WriteFile _
        Lib "kernel32" (ByVal hFile As Long, _
                        ByVal lpBuffer As String, _
                        ByVal nNumberOfBytesToWrite As Long, _
                        lpNumberOfBytesWritten As Long, _
                        lpOverlapped As Any) As Long
                        
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Declare Function GetPrivateProfileString _
        Lib "kernel32" _
        Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                          ByVal lpKeyName As Any, _
                                          ByVal lpDefault As String, _
                                          ByVal lpReturnedString As String, _
                                          ByVal nSize As Long, _
                                          ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString _
        Lib "kernel32" _
        Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                            ByVal lpKeyName As Any, _
                                            ByVal lpString As Any, _
                                            ByVal lpFileName As String) As Long

Public Declare Sub ZeroMemory2 Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

            
Public hStdIn  As Long   ' Handle Standard Input
Public hStdOut As Long   ' Handle Standard Output
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&

Public psEnginePath    As String   ' path of engine directory (init different VB6 / Office)
Public psDocumentPath  As String   ' path of office document
Public pbIsOfficeMode  As Boolean
Public plLastPostNodes As Long ' to avoid duplicate outputs

Public TableBasesRootEnabled As Boolean
Public TableBasesSearchEnabled As Boolean
Public LastTbProbeTime As Date
Public Const TB_MAX_PIECES = 5
Private oProxy As Object ' for tablebases
Private bTbBaseTrace As Boolean

'---------------------------------
' Log file
'---------------------------------
Public bLogPV          As Boolean  ' log PV in post mode
Public bLogMode        As Boolean
Public LogFile         As Long

Public LastFullPV As String

Private LanguageENArr(200) As String
Private LanguageArr(200) As String
Public LangCnt As Long


'---------------------------------------------------------------------------

Public Sub OpenCommHandles()
  ' Open IO channels to Winboard
  hStdIn = GetStdHandle(STD_INPUT_HANDLE)
  hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)
End Sub

Public Sub CloseCommChannels()
  ' Close IO channels to Winboard
  CloseHandle hStdIn
  CloseHandle hStdOut
End Sub

'---------------------------------------------------------------------------
'PollCommand() - check standard input
'
' returns TRUE if data found
'---------------------------------------------------------------------------
Function PollCommand() As Boolean

If ThreadNum <= 0 Then
  #If DEBUG_MODE <> 0 Then
    ' from Debug form
    PollCommand = FakeInputState
  #Else
    ' winboard input
    Dim sBuff       As String
    Dim lBytesRead  As Long
    Dim lTotalBytes As Long
    Dim lAvailBytes As Long
    Dim rc          As Long
    
    sBuff = String(2048, Chr$(0))
    rc = PeekNamedPipe(hStdIn, ByVal sBuff, 2048, lBytesRead, lTotalBytes, lAvailBytes)
    
    PollCommand = CBool(rc And lBytesRead > 0)
  #End If
Else
  '--- Multi-thread mode: helper threads get commands from main thread
  MainThreadStatus = ReadMainThreadStatus()
  'If bThreadTrace Then WriteTrace "PollCommand: ThreadStatusCheck:" & MainThreadStatus & " " & LastThreadStatus & " / " & Now()
  Select Case MainThreadStatus
  Case 1
    If LastThreadStatus <> MainThreadStatus Then
      ThreadCommand = "go" & vbLf: PollCommand = True
      WriteTrace "PollCommand: MainThreadStatus = 1" & " / " & Now()
    End If
  Case 0
   If LastThreadStatus <> MainThreadStatus Then
     ThreadCommand = "exit" & vbLf: PollCommand = True: bTimeExit = True
     WriteTrace "PollCommand: MainThreadStatus = 0" & " / " & Now()
   End If
  End Select
  LastThreadStatus = MainThreadStatus
End If
End Function

'---------------------------------------------------------------------------
'ReadCommand()
'---------------------------------------------------------------------------
Function ReadCommand() As String
  If ThreadNum > 0 Then
     If bThreadTrace Then WriteTrace "ReadCommand: ThreadCommand = " & ThreadCommand & " / " & Now()
     ReadCommand = ThreadCommand
     ThreadCommand = ""
     Exit Function
  End If
  
  #If DEBUG_MODE <> 0 Then
    ReadCommand = FakeInput ' from Debug form
    FakeInputState = False
    FakeInput = ""
  #Else
    Dim sBuff      As String
    Dim lBytesRead As Long
    Dim rc         As Long
    
    sBuff = String(2048, Chr$(0))
    rc = ReadFile(hStdIn, ByVal sBuff, 2048, lBytesRead, ByVal 0&)
    ReadCommand = Left$(sBuff, lBytesRead)
  #End If

End Function

'---------------------------------------------------------------------------
'SendCommand()
'
'---------------------------------------------------------------------------
Function SendCommand(ByVal sCommand As String) As String

  #If VBA_MODE = 1 Then
    ' OFFICE VBA
    With frmChessX
      If .txtIO.Visible Then
        If Len(.txtIO) > 32000 Then .txtIO = ""
        .txtIO = .txtIO & vbCrLf & sCommand
        .txtIO.SetFocus
        .txtIO.SelStart = Len(.txtIO)
        .txtIO.SelLength = 0
        DoEvents
      End If
    End With
  #End If

  #If DEBUG_MODE <> 0 Then
    ' VB DEBUG FORM
    With frmDebugMain
      If Len(.txtIO) > 32000 Then .txtIO = ""
      .txtIO = .txtIO & vbCrLf & sCommand
      .txtIO.SelStart = Len(.txtIO)
      .txtIO.SelLength = 0
      .Refresh
    End With
  #End If

  #If DEBUG_MODE = 0 And VBA_MODE = 0 Then
    ' WINBOARD STDOUT channel
    Dim lBytesWritten As Long
    Dim lBytes        As Long
    Dim rc            As Long
    
    sCommand = vbLf & sCommand & vbLf
    
    lBytes = Len(sCommand)
    
    rc = WriteFile(hStdOut, ByVal sCommand, lBytes, lBytesWritten, ByVal 0&)
  #End If

  SendCommand = sCommand

End Function

Public Sub WriteGame(sFile As String)
  '--- Write PGN file for game
  '
  ' Format:
  '[Event "F/S Return Match"]
  '[Site "Belgrade, Serbia Yugoslavia|JUG"]
  '[Date "1992.11.04"]
  '[Round "29"]
  '[White "Fischer, Robert J."]
  '[Black "Spassky, Boris V."]
  '[Result "1/2-1/2"]
  ' 1. e4 d5 2. d4 dxe4 3. Nf3

  Dim i As Long, h As Long, s As String, MoveCnt As Long, Cnt As Long
  Cnt = GameMovesCnt

  If Cnt = 0 Then Exit Sub
  s = "": MoveCnt = 0

  For i = 1 To Cnt Step 2
    MoveCnt = MoveCnt + 1
    s = s & CStr(MoveCnt) & ". " & CompToCoord(arGameMoves(i))

    If i + 1 <= Cnt Then s = s & " " & CompToCoord(arGameMoves(i + 1)) & " "
  Next i
 
  If s <> "" Then
    h = FreeFile()
    Open sFile For Append Lock Write As #h
    Print #h, "[Date " & Chr$(34) & Format(Now(), "YYYY.MM.DD HH:NN") & Chr$(34) & "]"
    Print #h, "[White " & Chr$(34) & "?" & Chr$(34) & "]"
    Print #h, "[Black " & Chr$(34) & "?" & Chr$(34) & "]"
    Print #h, "[Result " & Chr$(34) & "?" & Chr$(34) & "]"
  
    Print #h, s
    Close #h
  End If

End Sub

Public Sub ReadGame(sFile As String)
  ' Read PGN File
  Dim h            As Long, s As String, m As Long, sInp As String, m1 As String, m2 As String
  Dim asMoveList() As String
 
  InitGame
  bForceMode = True
 
  h = 10 'FreeFile()
  Open sFile For Input As #h

  Do Until EOF(h)
    Line Input #h, sInp
    sInp = Trim(sInp) & "  "

    If Left(sInp, 1) <> "[" Then '--- Ignore Header Tags
      asMoveList = Split(sInp, ".") ' split at move number dot

      For m = 0 To UBound(asMoveList)
        s = asMoveList(m)
        s = Replace(s, "-", "")
        s = Replace(s, "x", "")
        s = Replace(s, "+", "")
        s = Left(s, 10)
        
        If Left(s, 1) = " " Then ' behind move number
          s = Trim(s)
          'Debug.Print s
          m1 = Trim(Left(s, 4))

          If Len(m1) = 4 Then
            'Debug.Print m1, asMoveList(m)
            ParseCommand m1 & vbLf
          End If

          If Len(s) > 8 Then
            m2 = Trim(Mid(s, 6, 4))
            If Len(m2) >= 4 Then
            'Debug.Print m2, asMoveList(m)
              ParseCommand m2 & vbLf
            End If
          End If
        End If

      Next

    End If

  Loop

  Close #h
End Sub

Public Sub SendThinkInfo(Elapsed As Single, ActDepth As Long, CurrentScore As Long)
  Static FinalMoveForHint As TMOVE
  Dim sPost               As String, j As Long, sPostPV As String
 
  If pbIsOfficeMode Then
    '--- MS OFFICE
    sPost = " " & Translate("Depth") & ":" & ActDepth & "/" & MaxPly & " " & Translate("Score") & ":" & FormatScore(EvalSFTo100(CurrentScore)) & " " & Translate("Nodes") & ":" & Format("0.000", Nodes) & " " & Translate("Sec") & ":" & Format(Elapsed, "0.00")

    If plLastPostNodes <> Nodes Then
      SendCommand sPost
      plLastPostNodes = Nodes
                    
      sPostPV = "      >" & Translate("Line") & ": "

      For j = 1 To PVLength(1) - 1
        sPostPV = sPostPV & " " & MoveText(PV(1, j))

        ' Save Hint move
        If j = 1 And Not MovesEqual(FinalMoveForHint, PV(1, 1)) Then HintMove = EmptyMove ' for case that 1. ply as hash move only
        If j = 2 Then
          If PV(1, j).From > 0 Then HintMove = PV(1, j): FinalMoveForHint = PV(1, 1)
        End If

      Next

      SendCommand sPostPV
      
      ShowMoveInfo MoveText(FinalMove), ActDepth, MaxPly, EvalSFTo100(CurrentScore), Elapsed
    End If

  Else
 
    '--- VB6
    sPost = ActDepth & " " & EvalSFTo100(CurrentScore) & " " & (Int(Elapsed) * 100) & " " & Nodes
    sPostPV = ""

    For j = 1 To PVLength(1)
      If PV(1, j).From <> 0 Then sPostPV = sPostPV & " " & MoveText(PV(1, j))
    Next
    If Len(Trim(sPostPV)) > 8 Then
        LastFullPV = sPostPV
    Else
      If Left(Trim(sPostPV), 5) = Left(Trim(LastFullPV), 5) Then
        If Len(Trim(sPostPV)) < Len(Trim(LastFullPV)) Then
          sPostPV = LastFullPV
        End If
      End If
    End If
    
    sPost = sPost & sPostPV & " (" & MaxPly & "/" & HashUsagePerc & ")"
    If Not GotExitCommand() Then
      SendCommand sPost
    End If
  End If
 
End Sub

Public Sub SendRootInfo(Elapsed As Single, ActDepth As Long, CurrentScore As Long)
  Dim sPost As String, j As Long, sPV As String
 
  If pbIsOfficeMode Then
    '--- MS OFFICE
    sPost = " " & Translate("Depth") & ":" & ActDepth & "/" & MaxPly & " " & Translate("Score") & ":" & FormatScore(EvalSFTo100(CurrentScore)) & " " & Translate("Nodes") & ":" & Format("0.000", Nodes) & " " & Translate("Sec") & ":" & Format(Elapsed, "0.00")
    
    If plLastPostNodes <> Nodes Or Nodes = 0 Then
      SendCommand sPost
      plLastPostNodes = Nodes
      
      sPost = "      >Line: "
      For j = 1 To PVLength(1) - 1
        sPost = sPost & " " & MoveText(PV(1, j))
      Next
      SendCommand sPost
      
      ShowMoveInfo MoveText(FinalMove), ActDepth, MaxPly, EvalSFTo100(CurrentScore), Elapsed
    End If
  Else
    ' VB6
    sPost = ActDepth & " " & EvalSFTo100(CurrentScore) & " " & (Int(Elapsed) * 100) & " " & Nodes
    sPV = ""
    For j = 1 To PVLength(1) - 1
       If PV(1, j).From <> 0 Then sPV = sPV & " " & MoveText(PV(1, j))
    Next
    
    If Len(Trim(sPV)) > 8 Then
        LastFullPV = sPV
    Else
      If Trim(Left(sPV, 5)) = Trim(Left(LastFullPV, 5)) Then
        sPV = LastFullPV
      End If
    End If
    sPost = sPost & sPV
    If Not GotExitCommand() Then
      SendCommand sPost
    End If
  End If

  If bWinboardTrace Then If bLogPV Then LogWrite Space(6) & sPost
End Sub

Public Function GotExitCommand() As Boolean
 Dim sInput As String
  GotExitCommand = False
  If PollCommand Then
    sInput = ReadCommand
    If Left$(sInput, 1) = "." Then
      SendAnalyzeInfo
    Else
      If sInput <> "" Then
        ParseCommand sInput
        GotExitCommand = bExitReceived
      End If
    End If
  End If
End Function

Public Function FormatScore(ByVal lScore As Long) As String
  If lScore < -MATE_IN_MAX_PLY And lScore >= -MATE0 Then
    FormatScore = "-M" & CStr((Abs(MATE0) - Abs(lScore)) \ 2)
  ElseIf lScore > MATE_IN_MAX_PLY And lScore <= MATE0 Then
    FormatScore = "+M" & (MATE0 - lScore) \ 2
  ElseIf lScore = UNKNOWN_SCORE Then
    FormatScore = "?"
  Else
    FormatScore = Format$(lScore / 100#, "+0.00;-0.00")
  End If
End Function
            
Public Sub SendAnalyzeInfo()
  Dim sPost As String, Elapsed As Single
  Elapsed = TimerDiff(StartThinkingTime, Timer)
  sPost = "stat01: " & Int(Elapsed) & " " & Nodes & " " & IterativeDepth & " " & "1 1"
  If Not GotExitCommand() Then
    SendCommand sPost
  End If
End Sub

Public Sub WriteTrace(s As String)
  Dim h As Long
  On Error Resume Next
  'Debug.Print s
  If s <> "" Then
    h = FreeFile()
    If ThreadNum <= 0 Then
      Open psEnginePath & "\Trace_" & Format(Date, "YYMMDD") & ".txt" For Append Lock Write As #h
    Else
      Open psEnginePath & "\Trace_" & Format(Date, "YYMMDD") & "_T" & Trim(CStr(GetMax(0, ThreadNum))) & ".txt" For Append Lock Write As #h
    End If
          
    Print #h, s
    Close #h
  End If
 
  If pbIsOfficeMode Then SendCommand s
 
End Sub


'---------------------------------------------------------------------------
'ReadINISetting: Read values form INI file
'---------------------------------------------------------------------------
Function ReadINISetting(ByVal sSetting As String, ByVal sDefault As String) As String

  Dim sBuffer    As String
  Dim lBufferLen As Long

  sBuffer = Space(260)

  lBufferLen = GetPrivateProfileString("Engine", sSetting, sDefault, sBuffer, 260, psEnginePath & "\" & INI_FILE)
 
  If lBufferLen > 0 Then
    ReadINISetting = Left$(sBuffer, lBufferLen)
  Else
    'LogWrite "Error retrieving setting: " & sSetting, True, True
  End If

End Function

'---------------------------------------------------------------------------
' WriteINISetting: write values to INI file
'---------------------------------------------------------------------------
Function WriteINISetting(ByVal sSetting As String, ByVal sValue As String) As Boolean
  Dim lBufferLen As Long

  lBufferLen = WritePrivateProfileString("Engine", sSetting, sValue, psEnginePath & "\" & INI_FILE)
 
  If lBufferLen > 0 Then
    WriteINISetting = True
  Else
    LogWrite "Error writing setting: " & sSetting & "=" & sValue, True
    WriteINISetting = False
  End If

End Function

'---------------------------------------------------------------------------
'LogWrite: Write log file
'bTime adds the time
'---------------------------------------------------------------------------
Public Sub LogWrite(sLogString As String, _
                    Optional ByVal bTime As Boolean)

  Dim sStr As String
  LogFile = FreeFile
  sStr = sLogString

  If bTime Then sStr = Now & " - " & sStr

  Open psEnginePath & "\" & LCase(psAppName) & ".log" For Append Lock Write As #LogFile
  Print #LogFile, sStr
  'Debug.Print sStr
  Close #LogFile

End Sub

Public Sub ShowMoveInfo(ByVal sMove As String, _
                        ByVal lDepth As Long, _
                        ByVal lMaxPly As Long, _
                        ByVal lScore As Long, _
                        ByVal lTime As Single)
  #If VBA_MODE Then
    With frmChessX
      If InStr(sMove, "x") = 0 Then
        .lblMove = Translate("Move") & ": " & UCase(Left$(sMove, 2)) & "-" & UCase$(Mid$(sMove, 3))
      Else
        .lblMove = Translate("Move") & ": " & UCase(Left$(sMove, 2)) & "x" & UCase$(Mid$(sMove, 4))
      End If
      .lblDepth = Translate("Depth") & ": " & CStr(lDepth) & "/" & CStr(lMaxPly) & ":" & CStr(RootMoveCnt)
      .lblScore = Translate("Score") & " : " & FormatScore(lScore)
      .lblTime = Translate("Time") & ": " & Format(lTime, "0.00") & "s"
      DoEvents
    End With
  #End If
End Sub

Public Function FieldNumToCoord(ByVal ilFieldNum As Long) As String
  FieldNumToCoord = Chr$(Asc("a") + ((ilFieldNum - 1) Mod 8)) & Chr$(Asc("1") + ((ilFieldNum - 1) \ 8))
End Function

'
'--- Translate functions ---
'

Public Sub ReadLangFile(ByVal isLanguage As String)
  '--- sample: isLanguage = "DE"
  
  Dim sLine As String
  Dim i As Long
  Dim sFile As String
  Dim f As Long
  Dim c As String
  
  Dim sTextEN As String
  Dim sText As String
  sFile = psEnginePath & "\ChessBrainVB_Language_" & isLanguage & ".txt"
  LangCnt = 0

  If Dir(sFile) <> "" Then
    f = FreeFile()
    Open sFile For Input As #f
    
    Do While Not EOF(f)
      Line Input #f, sLine
      sLine = Trim$(sLine) 'Input
      If Not sLine = "" Then
        'Debug.Print sLine
        c = Left$(LTrim$(sLine), 1)
        If c <> ";" Then
          If StringSplit(sLine, sTextEN, sText) Then
            LangCnt = LangCnt + 1
            LanguageENArr(LangCnt) = sTextEN
            LanguageArr(LangCnt) = sText
          End If
        End If
      End If
    Loop
    Close #f
  End If ' File Exists
  
End Sub

Public Sub InitTranslate()
  Dim sLang As String
  sLang = ReadINISetting("LANGUAGE", "EN")
  If sLang = "EN" Then
    LangCnt = 0
  Else
    ReadLangFile sLang
  End If
End Sub

Public Function Translate(ByVal isTextEN As String) As String
  Dim i As Long
  If pbIsOfficeMode Then
    For i = 1 To LangCnt
      If LanguageENArr(i) = isTextEN Then Translate = LanguageArr(i): Exit Function
    Next
  End If
  Translate = isTextEN
End Function


Private Function StringSplit(sInput As String, _
                            ByRef sTextEN As String, _
                            ByRef sText As String) As Boolean
  'Split String from Format "english#languageX#"

  Dim v As Variant
  v = Split(sInput, "#", -1, vbBinaryCompare)
  If Not UBound(v) = 2 Then
    StringSplit = False
    Exit Function
  End If
  sTextEN = v(0): sText = v(1): StringSplit = True
End Function

  
Public Function InitTableBases() As Boolean
  ' Documentation: http://www.lokasoft.nl/tbapi.aspx
  Dim sURL As String
  On Error GoTo lblErr
  
  ' Tracing
  bTbBaseTrace = CBool(ReadINISetting("TBBASE_TRACE", "0") <> "0")
  
  sURL = ReadINISetting("TB_URL", "http://www.lokasoft.nl/tbweb/tbapi.wsdl")
  If bTbBaseTrace Then WriteTrace "Init endgame tablebase for: " & sURL & " / " & Now()
  
  Set oProxy = GetObject("soap:wsdl=" & sURL)
  InitTableBases = True
  If bTbBaseTrace Then WriteTrace "Init endgame tablebase OK! "
lblExit:
  Exit Function
  
lblErr:
  If bTbBaseTrace Then WriteTrace "Init endgame tablebase:ERROR! "
  InitTableBases = False
  TableBasesRootEnabled = False
  TableBasesSearchEnabled = False
  Resume lblExit
End Function
  
 Public Function IsTimeForTbBaseProbe() As Boolean
   '  max 20 sec for initial TB call needed, expect refresh after 30 min pause
   IsTimeForTbBaseProbe = CBool(TimeLeft > 20 Or FixedDepth <> NO_FIXED_DEPTH Or (DateDiff("n", LastTbProbeTime, Now()) < 30 And TimeLeft > 2))
   If bTbBaseTrace And Not IsTimeForTbBaseProbe Then WriteTrace "No time for endgame tablebase access: " & TimeLeft
 End Function
 
  
Public Function IsTbBasePosition(ByVal ActPly As Long) As Boolean
 Dim i As Long, ActPieceCnt As Long
 ActPieceCnt = PieceCntRoot
 For i = 1 To ActPly - 1
   If MovesList(i).Captured <> NO_PIECE Then ActPieceCnt = ActPieceCnt - 1
 Next
 IsTbBasePosition = CBool(ActPieceCnt <= TB_MAX_PIECES)
End Function

Public Sub TestTableBase()
  Dim sFEN As String, GameResultScore As Long, BestMove As String, BestMovesList As String
  Dim i As Long
  
  For i = 1 To 3
    If i Mod 2 = BCOL Then
     sFEN = "6k1/6p1/8/8/8/8/4P2P/6K1 b - -"
    Else
      sFEN = "7k/4P3/6K1/8/8/8/8/8 w - -"
     'sFEN = "R7/P4k2/8/8/8/8/r7/6K1 w - -"
    End If
    If ProbeTablebases(sFEN, GameResultScore, True, BestMove, BestMovesList) Then
      Debug.Print sFEN & " / Score: " & GameResultScore & "  > " & BestMove & " / " & Left(BestMovesList, 80)
      DoEvents
    Else
      Debug.Print "Error"
    End If
 Next
End Sub

  
Public Function ProbeTablebases(ByVal sFEN As String, ByRef GameResultScore As Long, ByVal bShowBestMoves As Boolean, ByRef BestMove As String, ByRef BestMovesList As String) As Boolean
  ' Online Web Access needed !
 ' Documentation: http://www.lokasoft.nl/tbapi.aspx
 ' Comsvcs.dll needed
 ' function returns false if no result
 Static bInitDone As Boolean
 Static bInitOK As Boolean
 Dim sResult As String
 
  GameResultScore = UNKNOWN_SCORE: BestMove = "": BestMovesList = "": ProbeTablebases = False
  If Not bInitDone Then
    bInitOK = InitTableBases()
    bInitDone = True
  End If
 
  If Not bInitOK Then ProbeTablebases = False: Exit Function
 
  On Error GoTo lblErr
 
  ' The score is given as distance to mat, or 0 when the position is a draw.
  ' An error response is returned when position is invalid or not in database. '
  ' e.g.  M5 = color to move gives mate in 5 , -M3 = color to move gets mated in 5 moves.
  sResult = Trim(oProxy.ProbePosition(sFEN))
  If sResult = "0" Then
    GameResultScore = 0
  ElseIf Left(sResult, 1) = "M" Then
    GameResultScore = MATE0 - 2 * Val("0" & Mid$(sResult, 2))
  ElseIf Left(sResult, 2) = "-M" Then
    GameResultScore = -MATE0 + 2 * Val("0" & Mid$(sResult, 3))
  End If
  
  ' Shows list of best move with score separated by Char=10 (vbLF)
  ' Moves = oProxy.GetBestMoves("6k1/6p1/8/8/8/8/4P2P/6K1 w - -")  => "Ra8-h8 M21, Ra8-b8 0, Ra8-c8 0, Kg1-f1 0, Kg1-h1 0, Ra8-d8 0, Ra8-g8 -M15, Ra8-e8 -M15, Ra8-f8 -M15,"
  If GameResultScore <> UNKNOWN_SCORE Then
    ProbeTablebases = True
    If bShowBestMoves Then
      BestMovesList = Replace(oProxy.GetBestMoves(sFEN), vbLf, ", ")
      ' Extract first move in internal format e2e4
      BestMove = ExtractFirstTbMove(BestMovesList)
    End If
  End If

  If bTbBaseTrace Then WriteTrace "endgame tablebase move: " & BestMove & " / Score: " & GameResultScore & " " & Now() & vbCrLf & PrintPos()

lblExit:
  Exit Function
  
lblErr:
  bInitDone = False
  ProbeTablebases = False
  Resume lblExit
End Function

Public Function ExtractFirstTbMove(ByVal sMoveList As String) As String
  Dim sMove As String, p As Long, c As String
  For p = 1 To Len(sMoveList)
    c = Mid$(sMoveList, p, 1)
    If (c >= "a" And c <= "h") Or (c >= "0" And c <= "9") Then
      If Len(sMove) <= 4 Then sMove = sMove & c
    ElseIf InStr("QRNB", c) > 0 Then
      ' Promote piece
      If Len(sMove) = 4 Then sMove = sMove & c
    ElseIf c = " " Or c = Chr$(10) Then
      Exit For
    End If
  Next
  If Len(sMove) = 4 Or Len(sMove) = 5 Then
    ExtractFirstTbMove = sMove
  Else
    ExtractFirstTbMove = ""
  End If
End Function






