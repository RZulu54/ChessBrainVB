Attribute VB_Name = "basIO"
'==================================================
'= IOBas:
'= Winboard / UCI communication / output of think results
'==================================================
Option Explicit
'--- Win API functions

#If VBA7 And Win64 Then
'Note: Win64 = Office64 bit (not Windows 64 bit)
  Declare PtrSafe Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
  Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
  Declare PtrSafe Function PeekNamedPipe _
          Lib "kernel32" (ByVal hNamedPipe As Long, _
                          lpBuffer As Any, _
                          ByVal nBufferSize As Long, _
                          lpBytesRead As Long, _
                          lpTotalBytesAvail As Long, _
                          lpBytesLeftThisMessage As Long) As Long
  Declare PtrSafe Function ReadFile _
          Lib "kernel32" (ByVal hFile As Long, _
                          lpBuffer As Any, _
                          ByVal nNumberOfBytesToRead As Long, _
                          lpNumberOfBytesRead As Long, _
                          lpOverlapped As Any) As Long
  Declare PtrSafe Function WriteFile _
          Lib "kernel32" (ByVal hFile As Long, _
                          ByVal lpBuffer As String, _
                          ByVal nNumberOfBytesToWrite As Long, _
                          lpNumberOfBytesWritten As Long, _
                          lpOverlapped As Any) As Long
  Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
  Declare PtrSafe Function GetPrivateProfileString _
          Lib "kernel32" _
          Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                            ByVal lpKeyName As Any, _
                                            ByVal lpDefault As String, _
                                            ByVal lpReturnedString As String, _
                                            ByVal nSize As Long, _
                                            ByVal lpFileName As String) As Long
  Declare PtrSafe Function WritePrivateProfileString _
          Lib "kernel32" _
          Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                              ByVal lpKeyName As Any, _
                                              ByVal lpString As Any, _
                                              ByVal lpFileName As String) As Long
  Public Declare PtrSafe Sub ZeroMemory2 _
                 Lib "kernel32.dll" _
                 Alias "RtlZeroMemory" (Destination As Any, _
                                        ByVal Length As Long)
#Else
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
  Public Declare Sub ZeroMemory2 _
                 Lib "kernel32.dll" _
                 Alias "RtlZeroMemory" (Destination As Any, _
                                        ByVal Length As Long)

#End If
                                      
'-------------
Public hStdIn  As Long   ' Handle Standard Input
Public hStdOut As Long   ' Handle Standard Output
Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public psEnginePath            As String   ' path of engine directory (init different VB6 / Office)
Public psDocumentPath          As String   ' path of office document
Public pbIsOfficeMode          As Boolean
Public plLastPostNodes         As Long ' to avoid duplicate outputs
Public EGTBasesEnabled         As Boolean
Public EGTBasesMaxPieces       As Long  ' 3,4,5,6 piece set
Public EGTBasesMaxPly          As Long ' max ply using EGTB in search
Public EGTBasesPath            As String  ' SYZYGY EGTB files path
Private oProxy                 As Object ' for online tablebases
Public bEGTbBaseTrace          As Boolean
Public EGTBasesHitsCnt         As Long ' count for GUI output
Public EGTBRootProbeDone       As Boolean
Public EGTBRootResultScore     As Long
Public EGTBBestMoveStr         As String, EGTBBestMoveListStr As String
Public EGTBMoveListCnt(MAX_PV) As Long, EGTBMoveList(MAX_PV, 199) As String
Public UCISyzygyPath           As String
Public UCISyzygyMaxPieceSet    As Long
Public UCISyzygyMaxPly         As Long
'---------------------------------
' Log file
'---------------------------------
Public bLogPV                  As Boolean  ' log PV in post mode
Public bLogMode                As Boolean
Public LogFile                 As Long
Public LastFullPV              As String
Public LanguageENArr(200)     As String
Public LanguageArr(200)       As String
Public LangCnt                 As Long
Public psLanguage              As String

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
  If EGTBasesEnabled And Not DebugMode Then
    ' wait to avoid windows error when programs exits in AREAN after tablesbase access  in Win7 ( ok for Win10)
    Dim i As Long

    For i = 1 To 15
      Sleep 500
      DoEvents
    Next

  End If
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
      
      sBuff = String(4096, Chr$(0))
      rc = PeekNamedPipe(hStdIn, ByVal sBuff, 4096, lBytesRead, lTotalBytes, lAvailBytes)
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
          If bThreadTrace Then WriteTrace "PollCommand: MainThreadStatus = 1" & " / " & Now()
      End If
      Case 0
        If LastThreadStatus <> MainThreadStatus Then
          ThreadCommand = "exit" & vbLf: PollCommand = True: bTimeExit = True
          If bThreadTrace Then WriteTrace "PollCommand: MainThreadStatus = 0" & " / " & Now()
        Else
          Sleep 25
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
    sBuff = String$(4096, Chr$(0))
    rc = ReadFile(hStdIn, ByVal sBuff, 4096, lBytesRead, ByVal 0&)
    ReadCommand = Left$(sBuff, lBytesRead)
  #End If
End Function

'---------------------------------------------------------------------------
'SendCommand()
'
'---------------------------------------------------------------------------
Function SendCommand(ByVal sCommand As String) As String
  Dim p As Long, s As String, sOut As String
  #If VBA_MODE = 1 Then

    ' OFFICE VBA
    With frmChessX
      If .txtIO.Visible Then
        If Len(.txtIO) > 64000 Then .txtIO = ""
        If .txtIO <> "" Then .txtIO = .txtIO & vbCrLf
        If Len(sCommand) > 120 Then
          s = sCommand: sOut = ""
          Do While Len(s) > 120
            p = InStrRev(Left$(s, 120), " ")
            sOut = sOut & Left$(s, p)
            s = Trim$(Mid$(s, p + 1))
            If s <> "" Then sOut = sOut & vbCrLf & Space(14)
          Loop
          sOut = sOut & s
          .txtIO = .txtIO & sOut
        Else
          .txtIO = .txtIO & sCommand
        End If
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
      If Len(.txtIO) > 32000 Then .txtIO = Left$(.txtIO, 8000)
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
  '--- Write file for game un UCI format
  '
  ' Format:
  '[Event "F/S Return Match"]
  '[Site "Belgrade, Serbia Yugoslavia|JUG"]
  '[Date "1992.11.04"]
  '[Round "29"]
  '[White "Fischer, Robert J."]
  '[Black "Spassky, Boris V."]
  '[Result "1/2-1/2"]
  ' 1. e2e4 e7e5 2. c2c4 f8e7 3. d2d4 e5d4 4. b1c3 d4c3
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
  BookMovePossible = False
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
            m2 = Trim$(Mid(s, 6, 4))
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

Public Sub SendThinkInfo(Elapsed As Single, ActDepth As Long, CurrentScore As Long, Alpha As Long, Beta As Long)
  Static FinalMoveForHint As TMOVE
  Static sLastInfo As String
  Dim sPost               As String, j As Long, sPostPV As String
  'pbIsOfficeMode = False 'Test
  If pbIsOfficeMode Then
    '--- MS OFFICE
    sPost = " " & Translate("Depth") & ":" & ActDepth & "/" & MaxPly & " " & Translate("Score") & ":" & FormatScore(EvalSFTo100(CurrentScore)) & " " & Translate("Nodes") & ":" & Format("0.000", CalcNodes()) & " " & Translate("Sec") & ":" & Format(Elapsed, "0.00")
    If plLastPostNodes <> CalcNodes() Then
      SendCommand sPost
      plLastPostNodes = CalcNodes()
      sPostPV = "      >" & Translate("Line") & ": "

      For j = 1 To PVLength(1) - 1
        sPostPV = sPostPV & " " & GUIMoveText(PV(1, j))
        ' Save Hint move
        If j = 1 And Not MovesEqual(FinalMoveForHint, PV(1, 1)) Then HintMove = EmptyMove ' for case that 1. ply as hash move only
        If j = 2 Then
          If PV(1, j).From > 0 Then HintMove = PV(1, j): FinalMoveForHint = PV(1, 1)
        End If
      Next

      If sPost <> sLastInfo Then
        SendCommand sPostPV
        sLastInfo = sPost
        ShowMoveInfo MoveText(FinalMove), ActDepth, MaxPly, EvalSFTo100(CurrentScore), Elapsed ' VBA mode only
      End If
    End If
  Else
    '--- VB6
    If UCIMode Then
      ' format: info depth 1 seldepth 1 multipv 1 score cp 417 nodes 51 nps 25500 tbhits 0 time 2 pv e8g8
      sPost = "info depth " & ActDepth & " seldepth " & MaxPly & " multipv 1 score " & UciGUIScore(CurrentScore, Alpha, Beta)
      If Nodes > 1000 Then sPost = sPost & " hashfull " & HashUsageUCI()
      sPost = sPost & " nodes " & CalcNodes() & " nps " & CalcNPS(Elapsed) & " tbhits " & EGTBasesHitsCnt & " time " & Int(Elapsed * 1000#) & " pv"
    Else
      sPost = ActDepth & " " & EvalSFTo100(CurrentScore) & " " & (Int(Elapsed) * 100) & " " & CalcNodes()
    End If
    
    sPostPV = ""
    For j = 1 To GetMax(1, PVLength(1) - 1)
      If PV(1, j).From <> 0 Then sPostPV = sPostPV & " " & GUIMoveText(PV(1, j))
    Next
    
    Dim bLastFullPVUsed As Boolean
    bLastFullPVUsed = False
    If Len(Trim(sPostPV)) > 12 Then ' more than 2 moves
      If Len(Trim(sPostPV)) < Len(Trim(LastFullPV)) Then
        If Left(Trim(LastFullPV), Len(Trim(sPostPV))) = Trim(sPostPV) Then
         sPostPV = LastFullPV
         bLastFullPVUsed = True
        End If
      End If
      If Not bLastFullPVUsed Then
        LastFullPV = sPostPV
        LastFullPVLen = PVLength(1)
        For j = 1 To PVLength(1): SetMove LastFullPVArr(j), PV(1, j): Next
      End If
    Else
      If Left(Trim(sPostPV), 5) = Left(Trim(LastFullPV), 5) Then
        If Len(Trim(sPostPV)) < Len(Trim(LastFullPV)) Then
          sPostPV = LastFullPV
        End If
      End If
    End If
    sPost = sPost & sPostPV
    If Not UCIMode And Not bWbPvInUciFormat Then sPost = sPost & "(" & MaxPly & "/" & HashUsagePerc & ")"
    If Not GotExitCommand() Then
      If sPost <> sLastInfo Then
       SendCommand sPost
       sLastInfo = sPost
      End If
    End If
  End If
End Sub

'Public Sub SendRootInfo(Elapsed As Single, ActDepth As Long, CurrentScore As Long, Alpha As Long, Beta As Long)
'  Dim sPost As String, j As Long, sPV As String
'  'CurrentScore = ScaleScoreByEGTB(CurrentScore)
'  If pbIsOfficeMode Then
'    '--- MS OFFICE
'    sPost = " " & Translate("Depth") & ":" & ActDepth & "/" & MaxPly & " " & Translate("Score") & ":" & FormatScore(EvalSFTo100(CurrentScore)) & " " & Translate("Nodes") & ":" & Format("0.000", CalcNodes()) & " " & Translate("Sec") & ":" & Format(Elapsed, "0.00")
'    If plLastPostNodes <> Nodes Or Nodes = 0 Then
'      SendCommand sPost
'      plLastPostNodes = Nodes
'      sPost = "      >Line: "
'
'      For j = 1 To PVLength(1) - 1
'        sPost = sPost & " " & MoveText(PV(1, j))
'      Next
'
'      SendCommand sPost
'      ShowMoveInfo MoveText(FinalMove), ActDepth, MaxPly, EvalSFTo100(CurrentScore), Elapsed
'    End If
'  Else
'    ' VB6
'    If UCIMode Then
'      ' format: info depth 1 seldepth 1 multipv 1 score cp 417 nodes 51 nps 25500 tbhits 0 time 2 pv e8g8
'      sPost = "info depth " & ActDepth & " seldepth " & MaxPly & " multipv 1 score " & UciGUIScore(CurrentScore, Alpha, Beta) & " nodes " & CalcNodes() & " nps " & CalcNPS(Elapsed) & " tbhits " & EGTBasesHitsCnt & " time " & Int(Elapsed * 1000#) & " pv"
'    Else
'      sPost = ActDepth & " " & EvalSFTo100(CurrentScore) & " " & (Int(Elapsed) * 100) & " " & CalcNodes()
'    End If
'    sPV = ""
'
'    For j = 1 To PVLength(1) - 1
'      If PV(1, j).From <> 0 Then sPV = sPV & " " & GUIMoveText(PV(1, j))
'    Next
'
'    If Len(Trim(sPV)) > 8 Then
'      LastFullPV = sPV
'    Else
'      If Trim(Left(sPV, 5)) = Trim(Left(LastFullPV, 5)) Then
'        sPV = LastFullPV
'      End If
'    End If
'    sPost = sPost & sPV
'    If Not GotExitCommand() Then
'      SendCommand sPost
'    End If
'  End If
'  If bWinboardTrace Then If bLogPV Then LogWrite Space(6) & sPost
'End Sub

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
  ElseIf lScore = VALUE_NONE Then
    FormatScore = "?"
  Else
    FormatScore = Format$(lScore / 100#, "+0.00;-0.00")
  End If
End Function
            
Public Sub SendAnalyzeInfo()
  Dim sPost As String, Elapsed As Single
  Elapsed = TimeElapsed
  sPost = "stat01: " & Int(Elapsed) & " " & CalcNodes() & " " & RootDepth & " " & "1 1"
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
'ReadINISetting: Read values fromm INI file
'---------------------------------------------------------------------------
Function ReadINISetting(ByVal sSetting As String, ByVal sDefault As String) As String
  
  Dim sBuffer    As String
  Dim lBufferLen As Long
  If pbMSExcelRunning Then
    #If VBA_MODE = 1 Then
      ReadINISetting = ReadINISettingExcel(sSetting, sDefault)
    #End If
  Else
    sBuffer = Space(260)
    lBufferLen = GetPrivateProfileString("Engine", sSetting, sDefault, sBuffer, 260, psEnginePath & "\" & INI_FILE)
    If lBufferLen > 0 Then
      ReadINISetting = Left$(sBuffer, lBufferLen)
    Else
      'LogWrite "Error retrieving setting: " & sSetting, True, True
    End If
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
Public Sub LogWrite(sLogString As String, Optional ByVal BTime As Boolean)
  Dim sStr As String
  LogFile = FreeFile
  sStr = sLogString
  If BTime Then sStr = Now & " - " & sStr
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
  Dim sLine   As String
  Dim i       As Long
  Dim sFile   As String
  Dim f       As Long
  Dim c       As String
  Dim sTextEN As String
  Dim sText   As String
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
  
  If pbMSExcelRunning Then
    psLanguage = "EN"
    #If VBA_MODE = 1 Then
    InitTranslateExcel
    #End If
  Else ' other VBA Office
    psLanguage = "EN"
    ReadLangFile "DE"
  End If
End Sub

Public Function Translate(ByVal isTextEN As String) As String
  Dim i As Long
  If pbIsOfficeMode And psLanguage = "DE" Then
    For i = 1 To LangCnt
      If LanguageENArr(i) = isTextEN Then Translate = LanguageArr(i): Exit Function
    Next
  End If
  Translate = isTextEN
End Function

Public Function StringSplit(sInput As String, _
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
  On Error GoTo lblErr
  EGTBasesEnabled = CBool(Trim(ReadINISetting("EGTB_ENABLED", "0")) = "1") Or TableBasesRootEnabled
  If Not EGTBasesEnabled Then InitTableBases = False: Exit Function
  'pbIsOfficeMode = True 'TEST
  If pbIsOfficeMode Then ' for VBA-GUI only
    ' Online endgame tablebases
    ' see: https://github.com/lichess-org/lila-tablebase
    EGTBasesMaxPieces = 7
    EGTBasesMaxPly = 1
    InitTableBases = True
  Else
    ' winboard / UCI mode: using SYZYGY endgame tablebases
    EGTBasesPath = Trim(ReadINISetting("TB_SYZYGY_PATH", psEnginePath))
    If UCIMode And Trim$(UCISyzygyPath) <> "" Then
      EGTBasesPath = UCISyzygyPath
    End If
    EGTBasesMaxPieces = Val("0" & ReadINISetting("TB_SYZYGY_MAX_PIECES", "0"))
    If UCIMode And UCISyzygyMaxPieceSet > 0 Then
      EGTBasesMaxPieces = UCISyzygyMaxPieceSet
    End If
    ' probe for first x plies only
    EGTBasesMaxPly = Val("0" & ReadINISetting("TB_SYZYGY_MAX_PLY", "1"))  ' ply 1=root
    InitTableBases = (EGTBasesMaxPieces > 2 And EGTBasesPath <> "")
    If UCIMode And UCISyzygyMaxPly > 0 Then
      EGTBasesMaxPly = UCISyzygyMaxPly
    End If
    If Trim$(EGTBasesPath) = "" Then EGTBasesEnabled = False: Exit Function
    '
    EGTBasesHitsCnt = 0
    If InitTableBases Then
      Dim ResultScore As Long, BestMove As String, MoveListStr As String, MoveCnt As Long
      InitTableBases = ProbeEGTB("8/8/8/3k4/5P2/5K2/8/8 b - - 0 1", ResultScore, True, BestMove, MoveListStr)
      If UCIMode Then
        If InitTableBases Then
          SendCommand "info string tablebases found"
        Else
          SendCommand "info string tablebases not found at:" & EGTBasesPath
        End If
      End If
    End If
    If bEGTbBaseTrace Then WriteTrace "InitTableBases: Path:" & EGTBasesPath & " PieceSet:" & EGTBasesMaxPieces & " > " & InitTableBases
  End If
  If bEGTbBaseTrace Then WriteTrace "Init endgame tablebase OK! "
lblExit:
  Exit Function
lblErr:
  If bEGTbBaseTrace Then WriteTrace "Init endgame tablebase:ERROR! "
  InitTableBases = False
  EGTBasesEnabled = False
  Resume lblExit
End Function
  
Public Function IsTimeForEGTbBaseProbe() As Boolean
  If Not pbIsOfficeMode Then
    IsTimeForEGTbBaseProbe = False
    If FixedDepth <> NO_FIXED_DEPTH Then IsTimeForEGTbBaseProbe = True: Exit Function
    ' If Ply < GetMax(3, RootDepth \ 3) Then
    If CBool(TimeLeft > 1.5) Then
      IsTimeForEGTbBaseProbe = True
    End If
    ' End If
  Else
    '  max 20 sec for initial online TB call needed, expect refresh after 30 min pause
    IsTimeForEGTbBaseProbe = CBool(TimeLeft >= 5 Or FixedDepth <> NO_FIXED_DEPTH)
  End If
  If bEGTbBaseTrace And Not IsTimeForEGTbBaseProbe Then WriteTrace "No time for endgame tablebase access: " & TimeLeft
End Function
 
Public Function IsEGTbBasePosition() As Boolean
  Dim ActPieceCnt As Long
  ActPieceCnt = 2 + WNonPawnPieces + PieceCnt(WPAWN) + BNonPawnPieces + PieceCnt(BPAWN)
  IsEGTbBasePosition = CBool(ActPieceCnt <= EGTBasesMaxPieces)
End Function

Public Sub TestTableBase()
  Dim sFEN As String, GameResultScore As Long, BestMove As String, BestMovesList As String
  Dim i    As Long
pbIsOfficeMode = True
TableBasesRootEnabled = True
InitTableBases

  For i = 1 To 1
    If i Mod 2 = BCOL Then
      sFEN = "6k1/6p1/8/8/8/8/4P2P/6K1 b - -"
    Else
      sFEN = "7k/4P3/6K1/8/8/8/8/8 w - -"
      'sFEN = "R7/P4k2/8/8/8/8/r7/6K1 w - -"
    End If
    sFEN = "8/6k1/6p1/8/7r/3P1KP1/8/8 w - - 0 1"
    sFEN = "2Q1k3/8/4K3/8/8/3P4/8/8 b - - 0 11"
    sFEN = "8/k7/2K5/8/3P4/1Q6/8/8 b - - 0 11"
    
    If ProbeTablebases(sFEN, GameResultScore, True, BestMove, BestMovesList) Then
      Debug.Print sFEN & " / Score: " & GameResultScore & "  > " & BestMove & " / " & Left(BestMovesList, 80)
      DoEvents
    Else
      Debug.Print "Error"
    End If
  Next

End Sub

Public Function ProbeTablebases(ByVal sFEN As String, _
                                ByRef GameResultScore As Long, _
                                ByVal bShowBestMoves As Boolean, _
                                ByRef BestMove As String, _
                                ByRef BestMovesList As String) As Boolean
  If pbIsOfficeMode Then
    ProbeTablebases = ProbeOnlineEGTB(sFEN, GameResultScore, BestMove, BestMovesList)
  Else
    ProbeTablebases = ProbeEGTB(sFEN, GameResultScore, bShowBestMoves, BestMove, BestMovesList)
  End If
End Function
  
'Public Function ProbeOnlineEGTB(ByVal sFEN As String, _
'                                ByRef GameResultScore As Long, _
'                                ByVal bShowBestMoves As Boolean, _
'                                ByRef BestMove As String, _
'                                ByRef BestMovesList As String) As Boolean
'  ' Online Web Access needed !
'  ' Documentation: http://www.lokasoft.nl/tbapi.aspx
'  ' Comsvcs.dll needed
'  ' function returns false if no result
'  Static bInitDone As Boolean
'  Static bInitOk   As Boolean
'  Dim sResult      As String, sCommand As String
'  GameResultScore = VALUE_NONE: BestMove = "": BestMovesList = "": ProbeOnlineEGTB = False
'  If Not bInitDone Then
'    bInitOk = InitTableBases()
'    bInitDone = True
'  End If
'  If Not bInitOk Then ProbeOnlineEGTB = False: Exit Function
'  On Error GoTo lblErr
'  ' The score is given as distance to mat, or 0 when the position is a draw.
'  ' An error response is returned when position is invalid or not in database. '
'  ' e.g.  M5 = color to move gives mate in 5 , -M3 = color to move gets mated in 5 moves.
'  sCommand = "curl http://tablebase.lichess.ovh/standard/mainline?fen=4k3/8/8/8/8/4K3/8/8_w_-_"
'  sResult = GetCommandOutput(sCommand)
'
'  If Trim$(sResult) = "" Then Exit Function  ' TODO Trim$(oProxy.ProbePosition(sFEN))
'  If sResult = "0" Then
'    GameResultScore = 0
'  ElseIf Left$(sResult, 1) = "M" Then
'    GameResultScore = MATE0 - 2 * Val("0" & Mid$(sResult, 2))
'  ElseIf Left$(sResult, 2) = "-M" Then
'    GameResultScore = -MATE0 + 2 * Val("0" & Mid$(sResult, 3))
'  End If
'  ' Shows list of best move
'  If GameResultScore <> VALUE_NONE Then
'    ProbeOnlineEGTB = True
'    If bShowBestMoves Then
'      BestMovesList = "" ' TODO
'    End If
'  End If
'  If bEGTbBaseTrace Then WriteTrace "endgame tablebase move: " & BestMove & " / Score: " & GameResultScore & " " & Now() & vbCrLf & PrintPos()
'lblExit:
'  Exit Function
'lblErr:
'  bInitDone = False
'  ProbeOnlineEGTB = False
'  Resume lblExit
'End Function

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

Public Function ProbeEGTB(ByVal sFEN As String, _
                          ByRef GameResultScore As Long, _
                          ByVal bShowBestMoves As Boolean, _
                          ByRef BestMove As String, _
                          ByRef BestMovesListStr As String) As Boolean
  '
  '--- Use Fathom.exe to access Syzygy Endgame Tablebases / no longer used since V4.00 > too slow
  '--- Output string is parsed for result and bestmove
  '
  Dim sCommand As String, sRet As String, p As Long, p2 As Long, i As Long, sResult As String, sSearch As String, sOut As String, MoveList() As String, TmpMove As TMOVE, MoveCnt As Long, DTZ As Long
  GameResultScore = VALUE_NONE: BestMove = "": BestMovesListStr = "": ProbeEGTB = False:  EGTBMoveListCnt(Ply) = 0: DTZ = 0
  On Error GoTo lblErr
  '
  '--- Call Fathom.exe and return output
  '
  sCommand = psEnginePath & "\Fathom.exe --path=" & Chr$(34) & EGTBasesPath & Chr$(34) & " " & Chr$(34) & sFEN & Chr$(34)
  sOut = GetCommandOutput(sCommand)
  If Trim$(sOut) = "" Then Exit Function
  sOut = Replace(sOut, Chr$(34), "") ' Remove "
  ' search for DTZ (distance to zero for fifty counter): [DTZ 11]
  sRet = Trim$(sOut)
  p = InStr(sRet, "[DTZ")
  If p > 0 Then
    sRet = Mid$(sRet, p + Len("[DTZ") + 1)
    p = InStr(sRet, "]"): If p = 0 Then Exit Function
    sRet = Trim$(Left$(sRet, GetMax(p - 1, 0)))
    DTZ = Val("0" & Trim$(sRet))
  End If
  sRet = Trim$(sOut)
  'Debug.Print sOut
  ' search for result: [WDL "Win"]
  p = InStr(sRet, "[WDL "): If p = 0 Then Exit Function
  sRet = Mid$(sRet, p + 5)
  p = InStr(sRet, "]"): If p = 0 Then Exit Function
  sResult = Left$(sRet, p - 1)

  Select Case sResult
    Case "Win"
      sSearch = "[WinningMoves"
      GameResultScore = ScorePawn.EG * 20# - 3 * (Ply + DTZ): ProbeEGTB = True
    Case "Draw", "CursedWin", "BlessedLoss" 'CursedWin/BlessedLoss: 50 move draw avoids loss/win
      sSearch = "[DrawingMoves"
      GameResultScore = 0: ProbeEGTB = True
    Case "Loss"
      sSearch = "[LosingMoves"
      GameResultScore = -(ScorePawn.EG * 20# - 3 * (Ply + DTZ)): ProbeEGTB = True
    Case Else
      sSearch = "????"
      Exit Function
  End Select

  EGTBasesHitsCnt = EGTBasesHitsCnt + 1
  ' search for moves: [WinningMoves "Rexd1, Re6, Rdxd1, Rc3"]
  p = InStr(sRet, sSearch): If p = 0 Then Exit Function
  sRet = Mid$(sRet, p + Len(sSearch) + 1)
  p = InStr(sRet, "]"): If p = 0 Then Exit Function
  sRet = Trim$(Left$(sRet, GetMax(p - 1, 0)))
  Dim s As String, CaptureVal As Long, BestCaptureVal As Long, tmp As String
  If sRet <> "" Then
    ' Convert best move to internal move (Rexd1  => e1d1), generate moves and find matching move
    MoveList = Split(sRet, " ")
    CaptureVal = -99999

    For i = 0 To UBound(MoveList())
      s = Trim$(MoveList(i))
      If s <> "" And InStr(s, ".") = 0 Then ' ignore move cnt '1. '
        If InStr(s, "-") = 0 Then ' ignore result '1-0'
          EGTBMoveListCnt(Ply) = EGTBMoveListCnt(Ply) + 1
          EGTBMoveList(Ply, EGTBMoveListCnt(Ply)) = CompToCoord(GetMoveFromSAN(s))
          If EGTBMoveListCnt(Ply) = 1 Then
            BestMove = EGTBMoveList(Ply, 1)
            'Debug.Print MoveText(BestMove)
          End If
          tmp = EGTBMoveList(Ply, EGTBMoveListCnt(Ply))
          TmpMove = TextToMove(tmp)
          If InStr(s, "x") > 0 Or Len(tmp) = 5 Then ' prefer captures/promotions
            If Len(tmp) = 5 Then
              CaptureVal = PieceAbsValue(TmpMove.Promoted) - PieceAbsValue(TmpMove.Piece) ' promotion
            Else
              CaptureVal = GetSEE(TmpMove)  ' try best capture
            End If
          Else
            CaptureVal = (PsqVal(1, TmpMove.Piece, TmpMove.Target) - PsqVal(1, TmpMove.Piece, TmpMove.From))
          End If
          If CaptureVal > BestCaptureVal Then
            BestCaptureVal = CaptureVal
            BestMove = EGTBMoveList(Ply, EGTBMoveListCnt(Ply))
          End If
          'Debug.Print MoveCnt & ">:" & s
        End If
      End If
    Next

    ' If sResult = "Loss" Then ' do not return move filter
    '   EGTBMoveListCnt = 0
    ' End If
  End If
  ' Find first move of best line " 1. d8=Q Kg4 2. Ke6 Kf4
  If bShowBestMoves Then
    BestMovesListStr = Mid$(sOut, InStrRev(sOut, "]") + 5)  ' find last ] from  [LosingMoves..]
  End If
  sRet = Trim$(Replace(BestMovesListStr, "...", ".")) & " " ' black to move : "1..."
  MoveCnt = 0
  MoveList = Split(sRet, " ")

  For i = 0 To UBound(MoveList())
    s = Trim$(MoveList(i))
    If s <> "" And InStr(s, ".") = 0 Then ' ignore move cnt '1. '
      If InStr(s, "-") = 0 Then ' ignore result '1-0'
        MoveCnt = MoveCnt + 1
        ' If MoveCnt = 1 Then
        '   BestMove = CompToCoord(GetMoveFromSAN(s))
        'Debug.Print MoveText(BestMove)
        ' End If
        'Debug.Print MoveCnt & ">:" & s
      End If
    End If
  Next

  'If MoveCnt > 0 Then
  '  Select Case sResult
  '  Case "Win"
  '    If BestCaptureVal > 150 Then MoveCnt = MoveCnt \ 2
  '    GameResultScore = ScorePawn.EG * 20# - 3 * MoveCnt
  '  Case "Loss"
  '    If BestCaptureVal > 150 Then MoveCnt = MoveCnt + 200 ' prefer good captures
  '    GameResultScore = -(ScorePawn.EG * 20# - 6 * MoveCnt)
  '  Case Else
  '    ' keep 0
  '  End Select
  'End If
lblExit:
  Exit Function
lblErr:
  ProbeEGTB = False
  Resume lblExit
End Function

Public Function CalcNodes() As Long
  Dim TotalNodes As Double
  If NoOfThreads > 1 Then TotalNodes = CDbl(NoOfThreads) * CDbl(Nodes) Else TotalNodes = Nodes
  If TotalNodes > 2147483647# Then CalcNodes = 9999999 Else CalcNodes = TotalNodes
End Function

Public Function CalcNPS(ByVal ElapsedTime As Single) As Long
  Dim TotalNodes As Double
  If NoOfThreads > 1 Then TotalNodes = CDbl(NoOfThreads) * CDbl(Nodes) Else TotalNodes = Nodes
  CalcNPS = CDbl(TotalNodes) / GetMaxSingle(0.01, ElapsedTime)
End Function

Public Function ScaleScoreByEGTB(Score As Long) As Long
  If EGTBRootResultScore = VALUE_NONE Or Abs(Score) > MATE_IN_MAX_PLY Or Ply > 1 Then
    ScaleScoreByEGTB = Score
  ElseIf EGTBRootResultScore > 0 Then
    ScaleScoreByEGTB = ScorePawn.EG * 20 + Score
  ElseIf EGTBRootResultScore < 0 Then
    ScaleScoreByEGTB = -ScorePawn.EG * 20 + Abs(Score)
  ElseIf EGTBRootResultScore = 0 Then
    ScaleScoreByEGTB = Score \ 10
  End If
End Function

Public Function UciGUIScore(ByVal UciScore As Long, ByVal Alpha As Long, ByVal Beta As Long) As String
     If UciScore <= -MATE_IN_MAX_PLY Then
        UciGUIScore = "mate -" & CStr((MATE0 - Abs(UciScore)) \ 2)
      ElseIf UciScore >= MATE_IN_MAX_PLY Then
        UciGUIScore = "mate " & CStr((MATE0 - UciScore) \ 2)
      Else
        UciGUIScore = "cp " & EvalSFTo100(UciScore)
        If UciScore <= Alpha Then
          UciGUIScore = UciGUIScore & " upperbound"
        ElseIf UciScore >= Beta Then
          UciGUIScore = UciGUIScore & " lowerbound"
        End If
      End If
End Function

Public Function TestEGTB() As String

' see: https://github.com/lichess-org/lila-tablebase

'Public Function TestEGTB(GameResult As enumEndOfGame) As String
' --- curl http://tablebase.lichess.ovh/standard/mainline?fen=4k3/6KP/8/8/6r1/8/7p/8_w_-_-
'curl http://tablebase.lichess.ovh/standard/mainline?fen=4k3/8/8/8/8/4K3/8/8_w_-_-
'{"mainline":[],"winner":null,"dtz":0,"precise_dtz":0}
Dim sInp As String, i As Long, sWinner As String, sCommand As String
Dim sTBMoves As String
Dim EGTBArr() As String
Dim GameResult As enumEndOfGame


TestEGTB = ""

sCommand = "curl http://tablebase.lichess.ovh/standard/mainline?fen=4k3/P7/8/8/8/4K3/p7/8_w_-_-"
sInp = GetCommandOutput(sCommand)


If sInp = "too many pieces" Then GameResult = NO_MATE: Exit Function

'sInp = "{""mainline"":[{""uci"":""g7h8"",""san"":""Kh8"",""dtz"":3,""precise_dtz"":3},{""uci"":""g4a4"",""san"":""Ra4"",""dtz"":-2,""precise_dtz"":-2},{""uci"":""h8g7"",""san"":""Kg7"",""dtz"":1,""precise_dtz"":1},{""uci"":""h2h1q"",""san"":""h1=Q"",""dtz"":-2,""precise_dtz"":-2},{""uci"":""g7f6"",""san"":""Kf6"",""dtz"":1,""precise_dtz"":1},{""uci"":""h1h7"",""san"":""Qxh7"",""dtz"":-5,""precise_dtz"":-5},{""uci"":""f6e5"",""san"":""Ke5"",""dtz"":4,""precise_dtz"":4},{""uci"":""h7g6"",""san"":""Qg6"",""dtz"":-3,""precise_dtz"":-3},{""uci"":""e5d5"",""san"":""Kd5"",""dtz"":2,""precise_dtz"":2},{""uci"":""g6d6"",""san"":""Qd6+"",""dtz"":-1,""precise_dtz"":-1},{""uci"":""d5d6"",""san"":""Kxd6"",""dtz"":21,""precise_dtz"":21},{""uci"":""a4a5"",""san"":""Ra5"",""dtz"":-20,""precise_dtz"":-20},{""uci"":""d6e6"",""san"":""Ke6"",""dtz"":19,""precise_dtz"":19},{""uci"":""a5h5"",""san"":""Rh5"",""dtz"":-18,""precise_dtz"":-18},{""uci"":""e6d6"",""san"":""Kd6"" "
'sInp = sInp & " ""winner"":""b"",""dtz"":-4,""precise_dtz"":-4}"
''sInp = """mainline"":[],""winner"":null,""dtz"":0,""precise_dtz"":0}"


EGTBArr() = Split(sInp, "uci"":""")
sTBMoves = ""
For i = 1 To UBound(EGTBArr)
  sTBMoves = sTBMoves & Left$(EGTBArr(i), 4) & " "
Next i
sTBMoves = Trim(sTBMoves)

i = InStr(sInp, "winner")
sWinner = ""
If i > 0 Then sWinner = Mid$(sInp, i + 9, 1) ' w =white, b =black, u =Null(Draw)
Select Case sWinner
Case "w"
  GameResult = WHITE_WON
Case "b"
  GameResult = BLACK_WON
Case "u"
  GameResult = DRAW_RESULT
Case Else
  GameResult = NO_MATE
End Select

' Public Enum enumEndOfGame ' Game result
'  NO_MATE = 0
'  WHITE_WON = 1
'  BLACK_WON = 2
'  DRAW_RESULT = 3
'  DRAW3REP_RESULT = 4
' End Enum

Debug.Print sTBMoves, sWinner
End Function

Public Function ProbeOnlineEGTB(ByVal sFEN As String, _
                                ByRef GameResultScore As Long, _
                                ByRef BestMove As String, _
                                ByRef BestMovesList As String) As Boolean
  ' Online Web Access needed ! Uses Windows program curl.exe (comes with Windows)
  ' Documentation: see: https://github.com/lichess-org/lila-tablebase
  ' sample call : curl.exe http://tablebase.lichess.ovh/standard/mainline?fen=4k3/6KP/8/8/6r1/8/7p/8_w_-_-
  ' function returns false if no result
  ' sampel string returned:
  ' 'sResult = "{""mainline"":[{""uci"":""g7h8"",""san"":""Kh8"",""dtz"":3,""precise_dtz"":3},{""uci"":""g4a4"",""san"":""Ra4"",""dtz"":-2,""precise_dtz"":-2},{""uci"":""h8g7"",""san"":""Kg7"",""dtz"":1,""precise_dtz"":1},{""uci"":""h2h1q"",""san"":""h1=Q"",""dtz"":-2,""precise_dtz"":-2},{""uci"":""g7f6"",""san"":""Kf6"",""dtz"":1,""precise_dtz"":1},{""uci"":""h1h7"",""san"":""Qxh7"",""dtz"":-5,""precise_dtz"":-5},{""uci"":""f6e5"",""san"":""Ke5"",""dtz"":4,""precise_dtz"":4},{""uci"":""h7g6"",""san"":""Qg6"",""dtz"":-3,""precise_dtz"":-3},{""uci"":""e5d5"",""san"":""Kd5"",""dtz"":2,""precise_dtz"":2},{""uci"":""g6d6"",""san"":""Qd6+"",""dtz"":-1,""precise_dtz"":-1},{""uci"":""d5d6"",""san"":""Kxd6"",""dtz"":21,""precise_dtz"":21},{""uci"":""a4a5"",""san"":""Ra5"",""dtz"":-20,""precise_dtz"":-20},{""uci"":""d6e6"",""san"":""Ke6"",""dtz"":19,""precise_dtz"":19},{""uci"":""a5h5"",""san"":""Rh5"",""dtz"":-18,""precise_dtz"":-18},{""uci"":""e6d6"",""san"":""Kd6"" "
  ' 'sResult = sResult & " ""winner"":""b"",""dtz"":-4,""precise_dtz"":-4}"
  '  ''sInp = """mainline"":[],""winner"":null,""dtz"":0,""precise_dtz"":0}"
  ' Test FEN/EPD: 8/8/1P6/5pr1/8/4R3/7k/2K5 w - -

  Static bInitDone As Boolean
  Static bInitOk   As Boolean
  Dim sResult      As String, sCommand As String
  Dim i As Long, sWinner As String
  Dim lDTM As Long ' Distance to mate
  Dim sTBMove As String, bMate As Boolean
  Dim GameResult As enumEndOfGame


  GameResultScore = VALUE_NONE: BestMove = "": BestMovesList = "": GameResult = NO_MATE: ProbeOnlineEGTB = False
  If Not bInitDone Then
    bInitOk = InitTableBases()
    bInitDone = True
  End If
  If Not bInitOk Then ProbeOnlineEGTB = False: Exit Function
  On Error GoTo lblErr
  ' The score is given as distance to mat, or 0 when the position is a draw.
  ' An error response is returned when position is invalid or not in database. '
  ' e.g.  M5 = color to move gives mate in 5 , -M3 = color to move gets mated in 5 moves.
 
  sCommand = "curl http://tablebase.lichess.ovh/standard?fen=" & Replace(sFEN, " ", "_")
  sResult = Trim(GetCommandOutput(sCommand))
  
'  sResult = """checkmate"":false,""stalemate"":false,""variant_win"":false,""variant_loss"":false,""insufficient_material"":false,""dtz"":-4,""precise_dtz"":-4,""dtm"":-10,""category"":""loss"",""moves"":[{""uci"":""g7h8"",""san"":""Kh8"",""zeroing"":false,""checkmate"":false,""stalemate"":false,""variant_win"":false,""variant_loss"":false,""insufficient_material"":false,""dtz"":3,""precise_dtz"":3,""dtm"":9,""category"":""win""}"
  
  'sCommand = "curl http://tablebase.lichess.ovh/standard/mainline?fen=" & Replace(sFEN, " ", "_")
  'sResult = "{""mainline"":[{""uci"":""g7h8"",""san"":""Kh8"",""dtz"":3,""precise_dtz"":3},{""uci"":""g4a4"",""san"":""Ra4"",""dtz"":-2,""precise_dtz"":-2},{""uci"":""h8g7"",""san"":""Kg7"",""dtz"":1,""precise_dtz"":1},{""uci"":""h2h1q"",""san"":""h1=Q"",""dtz"":-2,""precise_dtz"":-2},{""uci"":""g7f6"",""san"":""Kf6"",""dtz"":1,""precise_dtz"":1},{""uci"":""h1h7"",""san"":""Qxh7"",""dtz"":-5,""precise_dtz"":-5},{""uci"":""f6e5"",""san"":""Ke5"",""dtz"":4,""precise_dtz"":4},{""uci"":""h7g6"",""san"":""Qg6"",""dtz"":-3,""precise_dtz"":-3},{""uci"":""e5d5"",""san"":""Kd5"",""dtz"":2,""precise_dtz"":2},{""uci"":""g6d6"",""san"":""Qd6+"",""dtz"":-1,""precise_dtz"":-1},{""uci"":""d5d6"",""san"":""Kxd6"",""dtz"":21,""precise_dtz"":21},{""uci"":""a4a5"",""san"":""Ra5"",""dtz"":-20,""precise_dtz"":-20},{""uci"":""d6e6"",""san"":""Ke6"",""dtz"":19,""precise_dtz"":19},{""uci"":""a5h5"",""san"":""Rh5"",""dtz"":-18,""precise_dtz"":-18},{""uci"":""e6d6"",""san"":""Kd6""  ""winner"":""b"",""dtz"":-4,""precise_dtz"":-4}"
      

  ' more than 7 pieces?
  If sResult = "" Or sResult = "too many pieces" Then ProbeOnlineEGTB = False: Exit Function

  '--- search for UCI moves in result string
  sResult = Replace(sResult, """", " ")
  
  bMate = False
  If InStr(Left$(sResult, 90), "checkmate :true") > 0 Then
    lDTM = 0
    bMate = True
  Else
    i = InStr(sResult, "uci :")
    sTBMove = Trim$(Mid$(sResult, i + 6, 5)) ' 5th character for promotion: qrbn
    sResult = Mid$(sResult, i)
    i = InStr(sResult, "dtm :")
    If Mid$(sResult, i + 5, 4) = "null" Then
      lDTM = -1 ' NO_MATE
      ' search for DTZ: moves to next capture or Pawn move
      i = InStr(sResult, "dtz :")
      If Mid$(sResult, i + 5, 4) <> "null" Then
        lDTM = Val(Trim$(Mid$(sResult, i + 5, 5)))
        If lDTM > 0 Then lDTM = lDTM + 1000 Else lDTM = lDTM - 1000
      End If
    Else
     If InStr(sResult, "checkmate :true") > 0 Then
       lDTM = 0
       bMate = True
     Else
       lDTM = Val(Trim$(Mid$(sResult, i + 5, 5)))
     End If
    End If
  End If
  
  If lDTM < 0 Then
    If bWhiteToMove Then
      GameResult = WHITE_WON: GameResultScore = MATE0 - lDTM
    Else
      GameResult = BLACK_WON: GameResultScore = -MATE0 + lDTM
    End If
  ElseIf lDTM > 0 Then
     If Not bWhiteToMove Then
      GameResult = WHITE_WON: GameResultScore = MATE0 - Abs(lDTM)
    Else
      GameResult = BLACK_WON: GameResultScore = -MATE0 + Abs(lDTM)
    End If
  ElseIf lDTM = 0 Then
    If bMate Then ' Mate
      If bWhiteToMove Then
        GameResult = WHITE_WON: GameResultScore = MATE0
      Else
        GameResult = BLACK_WON: GameResultScore = -MATE0
      End If
    Else ' draw
      GameResult = DRAW_RESULT: GameResultScore = 0
    End If
  Else
    GameResult = NO_MATE
    GameResultScore = VALUE_NONE
  End If
  
  ' Shows list of best move
  If GameResult <> NO_MATE Then
    ProbeOnlineEGTB = True
    BestMove = Trim$(Left$(sTBMove, 5))
    BestMovesList = BestMove
  End If
  
  If bEGTbBaseTrace Then WriteTrace "endgame tablebase move: " & BestMove & " / Score: " & GameResultScore & " " & Now() & vbCrLf & PrintPos()
lblExit:
  Exit Function
lblErr:
  bInitDone = False
  ProbeOnlineEGTB = False
  Resume lblExit
End Function



