Attribute VB_Name = "SearchBas"
Option Explicit

'=======================================================
'= SearchBas:
'=
'= Search functions: Think->SearchRoot->Search->QSearch
'=
'= Think.....: Init search and call "SearchRoot" with increasing iterative depth 1,2,3... until time is over
'= SearchRoot: create root moves at ply 1 and call "Search" starting with ply 2
'= Search....: search for best move by recursive calls to itself down to iterative depth or time is over
'=             when iterative depth reached, call QSearch
'= QSearch...: quiescence search calculates all captures and check (first QS-ply only) by recursive calls to itself
'=             when all captures are done, the final position evaluation is returned
'=======================================================

Public Result                     As enumEndOfGame
Public BestScore                  As Long
Private CurrentScore              As Long
Public IterativeDepth             As Long
Public Nodes                      As Long
Public QNodes                     As Long
Public QNodesPerc                 As Double
Public EvalCnt                    As Long
Public bEndgame                   As Boolean
Public PlyScore(MAX_DEPTH)        As Long
Public MaxPly                     As Long

Public PV(MAX_PV, MAX_PV)         As TMove '--- principal variation(PV): best path of moves in current search tree
Public PVLength(MAX_PV)           As Long
Private bSearchingPV              As Boolean '--- often used for special handling (more exact search)
Public HintMove                   As TMove ' user hint move for GUI

Public MovesList(MAX_PV)          As TMove '--- currently searched move path
Public CntRootMoves               As Long
Public PliesFromNull(MAX_PV)      As Long '--- number of moves since last null move : for 3x draw detection

Public TempMove                   As TMove
Public FinalMove                  As TMove, FinalScore As Long '--- Final move selected
Public BadRootMove                As Boolean
Public PieceCntRoot As Long

Private bOnlyMove                 As Boolean  ' direct response if only one move
Private RootStartScore            As Long ' Eval score at root from view of side to move
Public PrevGameMoveScore          As Long ' Eval score at root from view of side to move
Private RootMatScore              As Long ' Material score at root from view of side to move
Public RootMoveCnt                As Long ' current root move for GUI

'--- Search performance: move ordering, cuts of search tree ---
Public History(2, MAX_BOARD, MAX_BOARD)    As Long     ' move history From square -> To square for color
Public HistoryVal(MAX_PV) As Long
Public CounterMove(15, MAX_BOARD) As TMove ' Good move against previous move
Public CounterMoves(15 * MAX_BOARD, 15 * MAX_BOARD) As Integer  ' Integer for less memory
Public CmhPtr(MAX_PV) As Long ' Pointer to first move of Countermoves

Public Type TKiller
 Killer1            As TMove 'killer moves: good moves for better move ordering
 Killer2            As TMove
 Killer3            As TMove
End Type

Public Killer(MAX_PV) As TKiller
Public Killer0 As TKiller
Public Killer2 As TKiller
Public EmptyKiller As TKiller

Public bSkipEarlyPruning          As Boolean  '--- no more cuts in search when null move tried

Public FutilityMoveCounts(1, MAX_PV)                     As Long '  [worse][depth]
Public Reductions(1, 1, 63, 63)   As Long ' [pv][worse][depth][moveNumber]
Public BestMovePly(MAX_PV)        As TMove
Public EmptyMove                  As TMove
Public RazorMargin(4) As Long

'--- piece bit constants for attack arrays
Public Const PLAttackBit = 1  ' Pawn left (from white view)
Public Const PRAttackBit = 2  ' Pawn right (from white view)
Public Const NAttackBit = 4
Public Const BAttackBit = 8
Public Const RAttackBit = 16
Public Const QAttackBit = 32
Public Const KAttackBit = 64
Public Const BXrayAttackBit = 128 ' Xray attack through own bishop/queen
Public Const RXrayAttackBit = 256 ' Xray attack through own rook/queen
Public Const QXrayAttackBit = 512 ' Xray attack through own bishop/rook/queen

'--- combined attack bits
Public Const PAttackBit = PLAttackBit Or PRAttackBit
Public Const PBNAttackBit = PAttackBit Or NAttackBit Or BAttackBit
Public Const QBAttackBit = QAttackBit Or BAttackBit
Public Const QRAttackBit = QAttackBit Or RAttackBit
Public Const QRBAttackBit = QAttackBit Or RAttackBit Or BAttackBit  ' slider attacks, detect pinned pieces
Public Const PNBRAttackBit = PAttackBit Or NAttackBit Or BAttackBit Or RAttackBit
'----

Public AttackBitCnt(1024) As Long   ' Returns number of attack bits set

Public EasyMove As TMove
Public EasyMovePV(3) As TMove
Public EasyMoveStableCnt As Long
Public bEasyMovePlayed As Boolean
Public QSDepth As Long

Private TmpMove         As TMove
Public bFirstRootMove   As Boolean
Public bFailedLowAtRoot As Boolean
Public bEvalBench       As Boolean
Public LegalRootMovesOutOfCheck As Long
Public IsTBScore           As Boolean


'--- end if declarations -----------------------------------------------------------------------------------------------------------

'---------------------------------------------------------------------------
'StartEngine: starts the chess engine to return a move
'---------------------------------------------------------------------------
Public Sub StartEngine()

  Dim CompMove      As TMove
  Dim sCoordMove    As String
  Dim bOldEvalTrace As Boolean

  '--- in winboard FORCE mode return, also check side to move
  'Debug.Print bCompIsWhite, bWhiteToMove, bForceMode, Result

  If bCompIsWhite <> bWhiteToMove Or bForceMode Or Result <> NO_MATE Then Exit Sub

  ' Init Search data
  QNodes = 0
  Nodes = 0
  Ply = 0
  Result = NO_MATE
  TimeStart = Timer
  bOldEvalTrace = bEvalTrace

  '
  '--- Start search ---
  '
  CompMove = Think()  '--- Calculate engine move

  If bAnalyzeMode Or bOldEvalTrace Then
    bAnalyzeMode = False
    bCompIsWhite = Not bCompIsWhite
    Exit Sub
  End If

  '--- Set time
  SearchTime = TimerDiff(TimeStart, Timer)
  TimeLeft = (TimeLeft - SearchTime) + TimeIncrement

  '--- Check  search result
  sCoordMove = CompToCoord(CompMove)
  Select Case Result
    Case NO_MATE
      PlayMove CompMove
      GameMovesAdd CompMove
      SendCommand Translate("move") & " " & sCoordMove
    
    Case BLACK_WON
      ' Mate?
      If CompMove.From <> 0 Then
        PlayMove CompMove
        GameMovesAdd CompMove
        SendCommand Translate("move") & " " & sCoordMove
      End If
      SendCommand "0-1 {" & Translate("Black Mates") & "}"
    Case WHITE_WON
      ' Mate?
      If CompMove.From <> 0 Then
        PlayMove CompMove
        GameMovesAdd CompMove
        SendCommand Translate("move") & " " & sCoordMove
      End If
      SendCommand "1-0 {" & Translate("White Mates") & "}"
    Case DRAW3REP_RESULT
      ' Draw?
      PlayMove CompMove
      GameMovesAdd CompMove
      SendCommand Translate("move") & " " & sCoordMove
      SendCommand "1/2-1/2 {" & Translate("Draw by repetition") & "}"
    Case Else
      ' Send move
      If CompMove.From <> 0 Then
        PlayMove CompMove
        GameMovesAdd CompMove
        SendCommand Translate("move") & " " & sCoordMove
      End If
      '--- Draw?
      If Fifty >= 100 Then
        SendCommand "1/2-1/2 {" & Translate("50 Move Rule") & "}"
      Else '--- no move
        SendCommand "1/2-1/2 {" & Translate("Draw") & "}"
      End If
  End Select

  'WriteTrace "move: " & CompMove & vbCrLf ' & "(t:" & Format(SearchTime, "###0.00") & " s:" & BestScore ' & " n:" & Nodes & " qn:" & QNodes & " q%:" & Format(QNodesPerc, "###0.00") & ")"

End Sub

'------------------------------------------------------------------------------
' Think: Start of Search with iterative deepening
'        aspiration windows used in 3 steps (slow without hash implementation)
'        called by: STARTENGINE, calls: SEARCH
'------------------------------------------------------------------------------
Public Function Think() As TMove

  Dim TimeUsed            As Single, Elapsed As Single
  Dim CompMove            As TMove, LastMove As TMove
  Dim IMax                As Long, i As Long
  Dim BoardTmp(MAX_BOARD) As Long
  Dim bOutOfBook          As Boolean
  Dim GoodMoves           As Long
  Dim RootAlpha           As Long
  Dim RootBeta            As Long
  Dim TimeFactor          As Single
  Dim OldScore            As Long, Delta As Long
  Dim bOldEvalTrace       As Boolean
  
  CompMove = EmptyMove
  ResetMaterial
  MaxPly = 0
  CurrentScore = -MATE0
  bSkipEarlyPruning = False
  bAddExtraTime = False
  LastNodesCnt = 0: RootMoveCnt = 0
  plLastPostNodes = 0: IsTBScore = False
  UnstablePvFactor = 1
  NextHashGeneration ' set next generation for hash entries
  LastFullPV = ""
  
  ' init easy move
  EasyMove = GetEasyMove() ' get easy move from previous Think call
  If bTimeTrace Then WriteTrace "Think: Easymove: " & MoveText(EasyMove)
  ClearEasyMove
  bEasyMovePlayed = False
  BestMoveChanges = 0
 
  ' Tracing
  bTimeTrace = CBool(ReadINISetting("TIMETRACE", "0") <> "0")
  If bTimeTrace Then
    WriteTrace " "
    WriteTrace "----- Start thinking, GAME MOVE >>>: " & GameMovesCnt \ 2 & " <<<"
  ElseIf bLogPV Then
    If bWinboardTrace Then LogWrite Space(6) & "----- Start thinking, GAME MOVE >>>: " & GameMovesCnt \ 2 & " <<<"
  End If

  For i = 0 To 99: PlyScore(i) = 0: MovesList(i).From = 0: MovesList(i).Target = 0: Next i
  For i = 0 To 20: TestCnt(i) = 0:  Next

    bTimeExit = False '--- Used for stop search, currently searched line result is not valid!!

    ' Opening book
    If BookPly < BOOK_MAX_PLY Then
      CompMove = ChooseBookMove
      If CompMove.From <> 0 Then
        CurrentScore = 0
        BestScore = CurrentScore
        SendCommand "0 0 0 0 (Book Move)"
        Think = CompMove
        Exit Function
      Else
        BookPly = BOOK_MAX_PLY + 1
        bOutOfBook = True
      End If
    End If

    ' init counters
    Nodes = 0
    QNodes = 0
    EvalCnt = 0
    InitEval

    bEvalTrace = bEvalTrace Or CBool(ReadINISetting("EVALTRACE", "0") <> "0") ' after InitEval
    bOldEvalTrace = bEvalTrace

    ' Scores
    BestScore = -MATE0
    RootStartScore = Eval()   ' Output for EvalTrace, sets EvalTrace=false
    If bOldEvalTrace Then Think = EmptyMove: Exit Function ' Exit if we only want an EVAl trace

    'LogWrite "Start Think "

    '--- Timer ---
    StartThinkingTime = Timer
    TotalTimeGiven = AllocateTime(PrevGameMoveScore)
    TimeForIteration = TotalTimeGiven
    'LogWrite " Given start: " & TotalTimeGiven
    If bAnalyzeMode Then TimeForIteration = 99999: TotalTimeGiven = 99999

    '  InitHash ' Init Hash: PieceValues may be changed or endgame phase => different eval

    HashBoard
    InHashCnt = 0

    IMax = MAX_DEPTH

    ' copy current board before start of search
    CopyIntArr Board, BoardTmp
    

    '--- Init search data
    Erase History()
    Erase CounterMove()
    Erase CounterMoves()
    Erase PV()
    CntRootMoves = 0

    Erase Killer()

    Erase MovesList()

    bExtraTime = False
    BadRootMove = False
    LastChangeMove = ""

    FinalScore = UNKNOWN_SCORE
    Result = NO_MATE

    '----------------------------
    '--- Iterative deepening ----
    '----------------------------
    For IterativeDepth = 1 To IMax
      Elapsed = TimerDiff(StartThinkingTime, Timer)
      TimeUsed = TimeUsed + (Elapsed - TimeUsed)
      TimeForIteration = TotalTimeGiven - TimeUsed
      bResearching = False
      BestMoveChanges = BestMoveChanges * 0.505 '  Age out PV variability metric
    
      If Not FixedDepthMode And FixedTime = 0 And Not bAnalyzeMode Then
        If MovesToTC <= 1 Then TimeFactor = 0.9 Else TimeFactor = 0.66 ' enough time for next interation?
        If TimeUsed > (TimeFactor * TotalTimeGiven) And IterativeDepth > LIGHTNING_DEPTH And Not bExtraTime Then
          If bTimeTrace Then WriteTrace "Exit SearchRoot: Used: " & Format$(TimeUsed, "0.00") & ", Given:" & Format$(TotalTimeGiven, "0.00") & ", Given*Factor=" & Format$(TotalTimeGiven * TimeFactor, "0.00")
          Exit For
        End If
        If TimeUsed > TotalTimeGiven And IterativeDepth > 1 Then
          If bTimeTrace Then WriteTrace "Exit SearchRoot2: Used: " & Format$(TimeUsed, "0.00") & ", Given:" & Format$(TotalTimeGiven, "0.00")
          Exit For
        End If
      Else
        If IterativeDepth > FixedDepth Then Exit For ' Fixed depth reached -> Exit
      End If
      bSearchingPV = True
      GoodMoves = 0
      PlyScore(IterativeDepth) = 0
      '--- Aspiration Window
    
      ' Delta = Eval100ToSF(25) ' aspiration window size
      RootAlpha = -MATE0: RootBeta = MATE0: Delta = -MATE0
      CurrentScore = RootAlpha
      
      OldScore = PlyScore(IterativeDepth - 1)
      If IterativeDepth >= 4 Then
          ' Delta = 45 '55 ' 30 ' 18 ' aspiration window size
          Delta = GetMin(25, 45 - IterativeDepth)
          
          RootAlpha = GetMax(OldScore - Delta, -MATE0)
          RootBeta = GetMin(OldScore + Delta, MATE0)
          If OldScore > MATE_IN_MAX_PLY Then
            RootBeta = MATE0
          ElseIf OldScore < -MATE_IN_MAX_PLY Then
            RootAlpha = -MATE0
          End If
      End If
      
      bFailedLowAtRoot = False
     
     Do While (True)
      '
      '--------- SEARCH ROOT ----------------
      '
      
      LastMove = SearchRoot(RootAlpha, RootBeta, IterativeDepth, GoodMoves)
     
      If bTimeExit Then Exit Do
        '
        '--- Research:no move found in Alpha-Beta window
        '
        bSearchingPV = True: GoodMoves = 0
          '-- SF6 logic
          If CurrentScore <= RootAlpha Then
            RootBeta = (RootAlpha + RootBeta) \ 2
            RootAlpha = GetMax(CurrentScore - Delta, -MATE0)
            bResearching = True
          ElseIf CurrentScore >= RootBeta Then
            RootAlpha = (RootAlpha + RootBeta) \ 2
            RootBeta = GetMin(CurrentScore + Delta, MATE0)
            bResearching = True
          Else
            Exit Do
          End If
          
          If CurrentScore > 2 * ScoreQueen.EG And CurrentScore <> MATE0 Then
            RootBeta = MATE0
          ElseIf CurrentScore < -2 * ScoreQueen.EG And CurrentScore <> -MATE0 Then
            RootAlpha = -MATE0
          End If
        
        #If DEBUG_MODE Then
         If IterativeDepth > 5 Then
          SendCommand "Research D:" & IterativeDepth & "/" & " SC:" & CurrentScore & " A:" & RootAlpha & ", B:" & RootBeta & " Last:" & OldScore & " Delta:" & Delta
         End If
        #End If
        
        Delta = Delta + (Delta \ 4 + 5)

        DoEvents
        bFailedLowAtRoot = CBool(CurrentScore < RootAlpha)
      Loop
    
      '--- Search result for current iteration ---
      If FinalScore <> UNKNOWN_SCORE Then
        CompMove = FinalMove
        BestScore = FinalScore
        PlyScore(IterativeDepth) = BestScore
        If (IterativeDepth > 1 Or IsTBScore) And bPostMode And PVLength(1) >= 1 Then
          Elapsed = TimerDiff(StartThinkingTime, Timer)
          If Not bExitReceived Then SendThinkInfo Elapsed, FinalScore ' Output to GUI
        End If
      End If
      

      CopyIntArr BoardTmp, Board  ' copy new position to main board

      If bOnlyMove Or IsTBScore Then
        bOnlyMove = False: Exit For
      End If
      ' LogWrite "THINK move End: IDepth:" & IterativeDepth & " : " & MovesPlyList
  
      If IterativeDepth > 2 And CurrentScore > MATE0 - IterativeDepth Then
        Exit For
      End If
         
      If IterativeDepth > 4 Then PVInstability ' Update with BestMoveCHanges
      
      If IterativeDepth >= 7 - 3 * Abs(pbIsOfficeMode) And EasyMove.From > 0 And Not FixedDepthMode And Not FixedTime > 0 Then
        If bTimeTrace Then WriteTrace "Easy check PV (IT:" & IterativeDepth & "): EM:" & MoveText(EasyMove) & ": PV1:" & MoveText(PV(1, 1))
        If MovesEqual(PV(1, 1), EasyMove) Then
          If bTimeTrace Then WriteTrace "Easy check2 bestmove: " & Format(BestMoveChanges, "0.000")
          If BestMoveChanges < 0.03 Then
            Elapsed = TimerDiff(StartThinkingTime, Timer)
            If bTimeTrace Then WriteTrace "Easy check3 Elapsed: " & Format$(Elapsed, "0.00") & Format$(TotalTimeGiven * 5# / 42#, "0.00")

            If Elapsed > TotalTimeGiven * 5# / 42# Then
              bEasyMovePlayed = True
              bTimeExit = True
              If bTimeTrace Then
                WriteTrace "Easy move played: " & MoveText(EasyMove) & " Elaspsed:" & Format$(Elapsed, "0.00") & ", Given:" & Format$(TotalTimeGiven, "0.00")
              End If
            End If
          End If
          
        End If
      End If
      
      If bTimeExit Then Exit For
      
      If PV(1, 3).From > 0 Then
        UpdateEasyMove
      Else
        If EasyMovePV(3).From > 0 Then ClearEasyMove
      End If
      
      
    Next ' Iteration <<<<<<<<

    If Nodes > 0 Then QNodesPerc = (QNodes / Nodes) * 100

    '--- Time management
    Elapsed = TimerDiff(StartThinkingTime, Timer)
    If EasyMoveStableCnt < 6 Or bEasyMovePlayed Then ClearEasyMove
    

    If bOutOfBook Then
      'LogWrite "out of book"
      'LogWrite Space(6) & "line: " & OpeningHistory
      'LogWrite Space(6) & "score: " & BestScore
    End If

    'LogWrite "End Think " & MoveText(CompMove) & " Result:" & Result
    If FinalScore <> UNKNOWN_SCORE Then PrevGameMoveScore = FinalScore Else PrevGameMoveScore = 0

    Think = CompMove '--- Return move

  End Function

'---------------------------------------------------------------------------
' SearchRoot: Search root moves
'             called by THINK,  calls SEARCH
'---------------------------------------------------------------------------
Private Function SearchRoot(ByVal Alpha As Long, _
                            ByVal Beta As Long, _
                            ByVal Depth As Long, _
                            GoodMoves As Long) As TMove

  Dim RootScore As Long, CurrMove As Long
  Dim bLegalMove     As Boolean, LegalMoveCnt As Long, bCheckBest As Boolean, QuietMoves As Long
  Dim Elapsed        As Single, lExtension As Long
  Dim BestMove       As TMove, CurrentMove As TMove
  Dim PrevMove As TMove
  Dim CutNode As Boolean, r As Long, bDoFullDepthSearch As Long
  Dim NewDepth       As Long, Depth1 As Long, bCaptureOrPromotion As Boolean
  Dim TimeUsed       As Single, bMoveCountPruning As Boolean

  '---------------------------------------------
  
  Ply = 1  ' start with ply 1
  
  CutNode = False: QSDepth = 0
  bOnlyMove = False
  GoodMoves = 0: RootMoveCnt = 0
  PrevMove = EmptyMove
  PliesFromNull(0) = Fifty: PliesFromNull(1) = Fifty
  If GameMovesCnt > 0 Then PrevMove = arGameMoves(GameMovesCnt)
  
  ' init history values
  HistoryVal(Ply + 1) = 0
  CmhPtr(Ply) = 0
  With Killer(Ply + 2)
    .Killer1 = EmptyMove: .Killer2 = EmptyMove: .Killer3 = EmptyMove
  End With
  
 ' Debug.Print "-------------"

  If bEvalBench Then
    'Benchmark evalutaion
    Dim start As Single, ElapsedT As Single, lCnt As Long
    start = Timer
    For lCnt = 1 To 1500000
      RootStartScore = Eval()
    Next
    ElapsedT = TimerDiff(start, Timer)
    MsgBox Format$(ElapsedT, "0.000")
    End
  End If

  LegalMoveCnt = 0
  QuietMoves = 0
  bFirstRootMove = True
  bFailedLowAtRoot = False

  PVLength(Ply) = Ply
  SearchStart = Timer
  

  ' Root check extent
  If InCheck Then
    Depth = Depth + 1
  End If

  InitPieceSquares

  RootStartScore = Eval()
  PieceCntRoot = 2 + PieceCnt(WPAWN) + PieceCnt(WKNIGHT) + PieceCnt(WBISHOP) + PieceCnt(WROOK) + PieceCnt(WQUEEN) + PieceCnt(BPAWN) + PieceCnt(BKNIGHT) + PieceCnt(BBISHOP) + PieceCnt(BROOK) + PieceCnt(BQUEEN) ' For TableBases
  StaticEvalArr(0) = RootStartScore

  ' PlyMatScore (1) = WMaterial - BMaterial
  RootMatScore = WMaterial - BMaterial: If Not bWhiteToMove Then RootMatScore = -RootMatScore

  '
  '---  Root moves loop --------------------
  '
  If IterativeDepth = 1 Then
      GenerateMoves 1, False, CntRootMoves
      OrderMoves 1, CntRootMoves, PrevMove, EmptyMove, EmptyMove, False, LegalRootMovesOutOfCheck
      SortMovesStable 1, 0, CntRootMoves - 1   ' Sort by OrderVal
  Else
      SortMovesStable 1, 0, CntRootMoves - 1  ' Sort by last iteration scores
      '  For CurrMove = 0 To CntRootMoves - 1: Debug.Print IterativeDepth, CurrMove, MoveText(Moves(1, CurrMove)), Moves(1, CurrMove).OrderValue: Next
      For CurrMove = 1 To CntRootMoves - 1: Moves(1, CurrMove).OrderValue = -100000000: Next
  End If
  SearchRoot = EmptyMove: IsTBScore = False
  
  '--- Endgame Tablebase check for root position
  If TableBasesRootEnabled Then
    If IsTbBasePosition(1) And IsTimeForTbBaseProbe Then ' min 20 sec for initial TB call needed
      Dim sTbFEN As String, TBGameResultScore As Long, sTBBestMove As String, sTBBestMovesList As String
      sTbFEN = WriteEPD()
      If ProbeTablebases(sTbFEN, TBGameResultScore, True, sTBBestMove, sTBBestMovesList) Then
        sTBBestMove = LCase(sTBBestMove) ' lower promoted piece
        For CurrMove = 0 To CntRootMoves - 1
          If CompToCoord(Moves(1, CurrMove)) = sTBBestMove Then
            SearchRoot = Moves(1, CurrMove)
            FinalMove = SearchRoot: FinalScore = TBGameResultScore: CurrentScore = FinalScore: PV(1, 1) = SearchRoot: PVLength(1) = 2
            If Fifty > 100 Then
              Result = DRAW_RESULT
            ElseIf FinalScore = 0 Then
              ' go on an try to win if opponent makes bas move
            ElseIf Abs(FinalScore) >= MATE0 - 2 Then
              If bWhiteToMove Then
                If FinalScore > 0 Then Result = WHITE_WON Else Result = BLACK_WON
              Else
                If FinalScore > 0 Then Result = BLACK_WON Else Result = WHITE_WON
              End If
            End If
            Elapsed = TimerDiff(TimeStart, Timer)
            Nodes = 1
            If Not bExitReceived Then SendRootInfo Elapsed, FinalScore   ' Output to GUI
            IsTBScore = True
            'MsgBox ">TB hit: " & MoveText(FinalMove)
            Exit Function
          End If
        Next
      End If
   End If
  End If
  
  
  For CurrMove = 0 To CntRootMoves - 1
    CurrentMove = Moves(1, CurrMove)
    MovePickerDat(Ply).CurrMoveNum = CurrMove
    
    RootScore = UNKNOWN_SCORE
    CmhPtr(Ply) = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target

    RemoveEpPiece
    MakeMove CurrentMove
    Ply = Ply + 1

    bLegalMove = False
    bCheckBest = False
  
    If CheckLegal(CurrentMove) Then
      Nodes = Nodes + 1
      bLegalMove = True: LegalMoveCnt = LegalMoveCnt + 1: RootMoveCnt = LegalMoveCnt
      bCaptureOrPromotion = CurrentMove.Captured <> NO_PIECE Or CurrentMove.Promoted <> 0

      bMoveCountPruning = (Depth < 16 And LegalMoveCnt >= FutilityMoveCounts(1, Depth))
    '  bMoveCountPruning = Depth < 15 And LegalMoveCnt >= (GetMax(0, (MovePickerDat(Ply).EndMoves - 15)) \ 5) + FutilityMoveCounts(1, Depth) + Abs(RootScore = DrawContempt) * 10
       
      If pbIsOfficeMode And IterativeDepth > 3 Then ' Show move cnt
        ShowMoveInfo MoveText(FinalMove), IterativeDepth, MaxPly, EvalSFTo100(FinalScore), Elapsed
      End If
        
      bFirstRootMove = CBool(LegalMoveCnt = 1)
        
      bSkipEarlyPruning = False
      MovesList(Ply - 1) = CurrentMove
      StaticEvalArr(Ply - 1) = RootStartScore
      
      RootMove = CurrentMove
      ' Debug.Print "Root:" & IterativeDepth & ": " & MoveText(CurrentMove), FinalScore
      r = 0: bDoFullDepthSearch = True
      lExtension = 0
      If (CurrentMove.IsChecking) Then
        If SEEGreaterOrEqual(CurrentMove, 0) Then
          lExtension = 1
        End If
      End If
      NewDepth = GetMax(0, Depth + lExtension - 1)
         
      'If IterativeDepth <= 4 Then GoTo lblNoMoreReductions
          
      '
      '--- Step 15. Reduced depth search (LMR). If the move fails high it will be re-searched at full depth.
      '
      If Depth >= 3 And LegalMoveCnt > 1 And (Not bCaptureOrPromotion Or bMoveCountPruning) Then
      
        r = Reduction(PV_NODE, 1, Depth, LegalMoveCnt)
           
        If bCaptureOrPromotion Then
          If r > 0 Then r = r - 1
        Else
          If CutNode Then
            r = r + 2
          Else
            '--- Decrease reduction for moves that escape a capture
            If CurrentMove.Castle = NO_CASTLE Then
              TmpMove.From = CurrentMove.Target: TmpMove.Target = CurrentMove.From: TmpMove.Piece = CurrentMove.Piece: TmpMove.Captured = NO_PIECE: TmpMove.SeeValue = UNKNOWN_SCORE
              ' Move back to old square, were we in danger there?
              If Not SEEGreaterOrEqual(TmpMove, -MAX_SEE_DIFF) Then r = r - 2  ' old square was dangerous
            End If
          End If
           
          HistoryVal(Ply) = History(PieceColor(CurrentMove.Piece), CurrentMove.From, CurrentMove.Target) - 4000
          
          '--- Decrease/increase reduction for moves with a good/bad history
          r = GetMax(0, r - HistoryVal(Ply) \ 20000)
          
        End If
        
        Depth1 = GetMax(NewDepth - r, 1)

        '--- Reduced SEARCH ---------
        RootScore = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, Depth1, CurrentMove, EmptyMove, True)
           
        bDoFullDepthSearch = (RootScore > Alpha And Depth1 <> NewDepth)
        r = 0
        
      Else
        bDoFullDepthSearch = (LegalMoveCnt > 1)
      End If

lblNoMoreReductions:
    '---  Step 16. Full depth search when LMR is skipped or fails high
    If bDoFullDepthSearch Then
         
        '------------------------------------------------
        '--->>>>  S E A R C H <<<<-----------------------
        '------------------------------------------------
        If (NewDepth <= 0) Then
          RootScore = -QSearch(NON_PV_NODE, -(Alpha + 1), -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          RootScore = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, NewDepth, CurrentMove, EmptyMove, False)
        End If
      End If
            
      ' For PV nodes only, do a full PV search on the first move or after a fail
      ' high (in the latter case search only if value < beta), otherwise let the
      ' parent node fail low with value <= alpha and to try another move.
      If (LegalMoveCnt = 1 Or RootScore > Alpha) Or RootScore = UNKNOWN_SCORE Then
        If NewDepth < 1 Then
          RootScore = -QSearch(PV_NODE, -Beta, -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          RootScore = -Search(PV_NODE, -Beta, -Alpha, NewDepth, CurrentMove, EmptyMove, False)
        End If
      End If
      
      If (Not bTimeExit) Or (LegalMoveCnt <= 1) Then bCheckBest = True
      
    End If
    '--- Unmake move
    RemoveEpPiece
    Ply = Ply - 1
    UnmakeMove CurrentMove
    ResetEpPiece
    
    ' check for best legal move
    If RootScore > Alpha And bLegalMove And bCheckBest Then
        
      BestMove = CurrentMove
      Alpha = RootScore
        
      CurrentScore = Alpha
        
      If LegalMoveCnt > 1 Then BestMoveChanges = BestMoveChanges + 1#
        
      If Not bTimeExit Then
        GoodMoves = GoodMoves + 1
      End If
      '--- Save final move
      If Not bTimeExit Or FinalScore = UNKNOWN_SCORE Then
        FinalMove = BestMove: FinalScore = CurrentScore
      
        ' Set root move order value for next iteration <<<<<<<<<<<<<<<<<
        Moves(1, CurrMove).OrderValue = CurrentScore
      End If
        
      If IterativeDepth > 3 Then
        If BestScore < PlyScore(IterativeDepth - 1) - 30 Then BadRootMove = True Else BadRootMove = False
      End If
        
      ' Store PV
      UpdatePV Ply, BestMove

      'Extra Time ?
      If Not FixedDepthMode() And FixedTime = 0 And Not bExtraTime And IterativeDepth > 3 And TimeLeftCorr > 3 And (MovesToTC > 1 Or MovesToTC = 0) Then
        Elapsed = TimerDiff(StartThinkingTime, Timer)
        TimeUsed = TimeUsed + (Elapsed - TimeUsed)
        If TimeUsed > TimeForIteration / 4# Then
          bAddExtraTime = False
          If LastChangeMove <> "" And IterativeDepth > 4 And LastChangeDepth >= IterativeDepth - 1 And LastChangeMove <> MoveText(PV(1, 1)) And Abs(FinalScore - PrevGameMoveScore) > 40 + Abs(PrevGameMoveScore) \ 10 Then
            bAddExtraTime = True
            If bTimeTrace Then WriteTrace "ExtraTime  LastChangeDepth: " & LastChangeDepth
          ElseIf IterativeDepth > 5 And Abs(FinalScore - PrevGameMoveScore) > 80 + Abs(PrevGameMoveScore) \ 10 Then
            bAddExtraTime = True ' drastic score change
            If bTimeTrace Then WriteTrace "ExtraTime  DiffScore: " & Abs(FinalScore - PrevGameMoveScore) & "," & PrevGameMoveScore
          ElseIf bResearching Then
            bAddExtraTime = True
            If bTimeTrace Then WriteTrace "ExtraTime  Researching: "
          End If
            
          If bAddExtraTime Then
            AllocateExtraTime '-- bAddExtraTime
          End If
        End If
      End If '-- Extra time
        
      LastChangeDepth = IterativeDepth
      LastChangeMove = MoveText(PV(1, 1))
        
      If (IterativeDepth >= 3 Or Abs(BestScore) >= MATE_IN_MAX_PLY) And bPostMode And (Not bTimeExit) Then
        Elapsed = TimerDiff(TimeStart, Timer)
        If Not bExitReceived Then SendRootInfo Elapsed, CurrentScore   ' Output to GUI
      End If
    End If

    If Not FixedDepthMode And Not bTimeExit And GoodMoves > 0 And Not bAnalyzeMode Then
      If FixedTime > 0 Then
        If TimerDiff(StartThinkingTime, Timer) >= FixedTime - 0.1 Then
          bTimeExit = True
        End If
      ElseIf (IterativeDepth > LIGHTNING_DEPTH) Then ' Time for next move?
        SearchTime = TimerDiff(TimeStart, Timer)
        If SearchTime > TotalTimeGiven * 0.75 Then
          If bTimeTrace Then WriteTrace "Exit SearchRoot3: Used:" & Format$(SearchTime, "0.00") & " TotalTimeGiven:" & Format$(TotalTimeGiven, "0.00")
          bTimeExit = True
        End If
      End If
    End If

    If (bTimeExit And LegalMoveCnt > 0) Or RootScore = MATE0 - 1 Then Exit For
    If pbIsOfficeMode Then
      If bTimeExit Then
        SearchTime = TimerDiff(TimeStart, Timer)
        'Debug.Print Nodes, SearchTime
      End If
      #If VBA_MODE = 1 Then
       '-- Office sometimes lost focus for Powepoint
       If Application.Name = "Microsoft PowerPoint" Then
         If IterativeDepth > 4 Then frmChessX.cmdStop.SetFocus
       End If
      #End If
      If IterativeDepth > 2 Then DoEvents
    Else
      If IterativeDepth > 6 Then DoEvents
    End If
    '--- Add Quiet move, used for pruning and history update
    If CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 And QuietMoves < 64 Then
      If Not MovesEqual(BestMove, CurrentMove) Then QuietMoves = QuietMoves + 1: QuietsSearched(Ply, QuietMoves) = CurrentMove
    End If
    
    If LegalMoveCnt > 0 And RootScore >= Beta Then Exit For
    If bTimeExit Then Exit For
    
  Next CurrMove
  '---<<< End of root moves loop -------------

  If LegalMoveCnt = 0 Then
    If InCheck Then
      If bWhiteToMove Then
        Result = BLACK_WON
      Else
        Result = WHITE_WON
      End If
    Else
      Result = DRAW_RESULT
    End If
    GoodMoves = -1
  Else
    If LegalMoveCnt = 1 And Not bTimeExit Then bOnlyMove = True
    
    If RootScore = MATE0 - 2 Then

      If bWhiteToMove Then
        Result = WHITE_WON
      Else
        Result = BLACK_WON
      End If
    Else
      If Fifty > 100 Then
        Result = DRAW_RESULT
      End If
    End If
  End If
  
  If FinalMove.From > 0 And Not bTimeExit And FinalMove.Captured = NO_PIECE And FinalMove.Promoted = 0 Then
     UpdateStats Ply, FinalMove, QuietMoves, EmptyMove, StatBonus(IterativeDepth)
  End If
  

  SearchRoot = FinalMove
  'WriteDebug "Root: " & IterativeDepth & " Best:" & MoveText(SearchRoot) & " Sc:" & CurrentScore & " M:" & GoodMoves

End Function

'---------------------------------------------------------------------------
' Search: Search moves from ply=2 to x, finally calls QSearch
'         called by SEARCHROOT, calls SEARCH recursively , then QSEARCH
'---------------------------------------------------------------------------
Private Function Search(ByVal PVNode As Boolean, _
                        ByVal Alpha As Long, _
                        ByVal Beta As Long, _
                        ByVal Depth As Long, _
                        InPrevMove As TMove, _
                        ExcludedMove As TMove, _
                        ByVal CutNode As Boolean) As Long

  Dim CurrentMove            As TMove, Score As Long, bNoMoves As Boolean, bLegalMove As Boolean
  Dim NullScore              As Long, PrevMove As TMove, QuietMoves As Long, rBeta As Long, rDepth As Long
  Dim StaticEval             As Long, GoodMoves As Long, NewDepth As Long, LegalMoveCnt As Long, MoveCnt As Long
  Dim bExtraTimeDone         As Boolean, lExtension As Long, lPlyExtension As Long
  Dim bMoveCountPruning As Boolean, bKillerMove As Boolean

  Dim r            As Long, Improv As Long, bCaptureOrPromotion As Boolean, LmrDepth As Long, bDoFullDepthSearch As Boolean, Depth1 As Long
  Dim BestValue              As Long, bIsNullMove As Boolean, ThreatMove As TMove, TryBestMove As TMove
  Dim bHashFound             As Boolean, ttHit As Boolean, HashEvalType As Long, HashScore As Long, HashStaticEval As Long, HashDepth As Long
  Dim EvalScore              As Long, HashKey As THashKey, HashMove As TMove, ttMove As TMove, ttValue As Long
  Dim BestMove               As TMove, sInput As String
  Dim Cmh                    As Long, Fmh As Long, FMh2 As Long, HistVal As Long, CurrPtr As Long, Cm_Ok As Boolean, Fm_Ok As Boolean, F2_Ok As Boolean
  Dim IsTbPos                As Boolean, bSingularExtensionNode As Boolean, Penalty As Long, bSkipQuiets As Boolean
  '----
  Debug.Assert Not (PVNode And CutNode)
  Debug.Assert (PVNode Or (Alpha = Beta - 1))
  Debug.Assert (-MATE0 <= Alpha And Alpha < Beta And Beta <= MATE0)
  
  '
  '--- Step 1. Initialize node for search -------------------------------------upd
  '
  PrevMove = InPrevMove '--- bug fix: make copy to avoid changes in parameter use
  BestValue = UNKNOWN_SCORE: BestMove = EmptyMove: BestMovePly(Ply) = EmptyMove
  EvalScore = UNKNOWN_SCORE: StaticEval = UNKNOWN_SCORE: StaticEvalArr(Ply) = UNKNOWN_SCORE
  ThreatMove = EmptyMove
  bIsNullMove = (PrevMove.From < SQ_A1)
  If bSearchingPV Then PVNode = True: CutNode = False
  'If Nodes = 1127 Then Stop
  If Ply > MaxPly Then MaxPly = Ply '--- Max depth reached in normal search
  If Depth < 0 Then Depth = 0
  HistoryVal(Ply + 1) = 0
  
  CmhPtr(Ply) = 0
  With Killer(Ply + 2)
    .Killer1 = EmptyMove: .Killer2 = EmptyMove: .Killer3 = EmptyMove
  End With

  
  '
  '--- Step 2. Check for aborted search and immediate draw
  '
  HashKey = HashBoard() ' Save current position hash keys for insert later
  If Not bIsNullMove Then
    '--- Step 2. Check immediate draw
    '--- Draw ?
    If Is3xDraw(HashKey, GameMovesCnt, Ply) Then
      If CompToMove() Then Search = DrawContempt Else Search = -DrawContempt
      PVLength(Ply) = 0
      Exit Function
    End If
    GamePosHash(GameMovesCnt + Ply - 1) = HashKey
  Else
    GamePosHash(GameMovesCnt + Ply - 1) = EmptyHash
  End If
   
  ' Endgame tablebase position?
  IsTbPos = False
  If TableBasesSearchEnabled And Ply = 2 Then ' For first computer ply only because web access is very slow
    If IsTbBasePosition(Ply) And IsTimeForTbBaseProbe Then IsTbPos = True
  End If
  '
  '--- Step 4. Transposition table lookup
  '
  bHashFound = False: ttHit = False: HashMove = EmptyMove
  ttHit = False: ttMove = EmptyMove: ttValue = UNKNOWN_SCORE:  NullScore = UNKNOWN_SCORE
    
  If Depth >= 0 Then
    ttHit = IsInHashTable(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
    If ttHit Then ttMove = HashMove: ttValue = HashScore
        
    If (Not PVNode Or HashDepth = TT_TB_BASE_DEPTH) And HashDepth >= Depth And ttHit And ttValue <> UNKNOWN_SCORE And HashMove.From > 0 Then
      If ttValue >= Beta Then
        bHashFound = (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT)
      Else
        bHashFound = (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT)
      End If
            
      If bHashFound And ExcludedMove.From = 0 Then
        If IsTbPos And HashDepth <> TT_TB_BASE_DEPTH Then
           ' Ignore Hash and continue with TableBase query
        Else
          If ttMove.From >= SQ_A1 Then
            '--- Save PV ---
            If ttValue > Alpha And ttValue < Beta Then UpdatePV Ply, HashMove
                
            If ttValue >= Beta Then
              If ttMove.Captured = NO_PIECE And ttMove.Promoted = 0 Then
                '--- Update statistics
                UpdateStats Ply, ttMove, 0, PrevMove, StatBonus(Depth)
              End If
          
              ' Extra penalty for a quiet TT move in previous ply when it gets refuted
              If PrevMove.Captured = NO_PIECE Then
                If PrevMove.From > 0 And Ply > 2 Then
                  If MovePickerDat(Ply - 1).CurrMoveNum = 0 Then
                    UpdateCmStats Ply - 1, PrevMove.Piece, PrevMove.Target, -StatBonus(Depth + 1)
                  End If
                End If
              End If
            ElseIf PrevMove.Captured = NO_PIECE And PrevMove.Promoted = 0 Then
              Penalty = -StatBonus(Depth + 1)
              UpdHistory ttMove.Piece, ttMove.From, ttMove.Target, Penalty
              If MovePickerDat(Ply - 1).CurrMoveNum = 0 Then
                UpdateCmStats Ply, ttMove.Piece, ttMove.Target, Penalty
              End If
            End If ' ttValue >= Beta
          End If ' ttMove.From >= SQ_A1
          
          BestMove = ttMove: Search = ttValue
          Exit Function  ' <<<< exit with TT move
        End If
      End If
    End If
  End If  '--- End Hash
   
  ' ---- Q S E A R C H -----
  If Depth <= 0 Or Ply >= MAX_DEPTH Then
    Search = QSearch(PVNode, Alpha, Beta, MAX_DEPTH, PrevMove, QS_CHECKS)
    Exit Function  '<<<<<<< R E T U R N >>>>>>>>
  End If
    
  StaticEval = UNKNOWN_SCORE
  StaticEvalArr(Ply + 1) = UNKNOWN_SCORE
  bNoMoves = True
  BestMovePly(Ply + 1) = EmptyMove
    
  '--- Check Time ---
  If Not FixedDepthMode Then
    '-- Fix:Nodes Mod 1000 > not working because nodes are incremented in QSearch too
    If Nodes > LastNodesCnt + GUICheckIntervalNodes And (IterativeDepth > LIGHTNING_DEPTH) Then
      'If pbIsOfficeMode Then DoEvents
      ' --- Check new commands from GUI (i.e. analyze stop)
      If PollCommand Then
        sInput = ReadCommand
        If Left$(sInput, 1) = "." Then
          SendAnalyzeInfo
        Else
          If sInput <> "" Then
            ParseCommand sInput
          End If
        End If
      End If
           
      LastNodesCnt = Nodes
      If bTimeExit Then Search = 0: Exit Function
      If FixedTime > 0 Then
        If Not bAnalyzeMode And TimerDiff(TimeStart, Timer) >= FixedTime - 0.1 Then bTimeExit = True: Exit Function
      ElseIf TimeForIteration - (TimerDiff(SearchStart, Timer)) <= 0 And Not bAnalyzeMode Then
        If BadRootMove And Not bExtraTime And TimeLeftCorr > 5 * TimeForIteration Then
          bExtraTimeDone = AllocateExtraTime()
        Else
          bExtraTimeDone = False
        End If
        If Not bExtraTimeDone Then
          If bTimeTrace Then WriteTrace "Exit Search: TimeUsed: " & Format$(TimerDiff(SearchStart, Timer), "0.00") & ", Given:" & Format$(TimeForIteration, "0.00")
          bTimeExit = True: Search = 0: Exit Function
        End If
      End If
    End If
  End If
    
  '
  '--- Step 3.:  Mate distance pruning
  '
  Alpha = GetMax(-MATE0 + Ply, Alpha)
  Beta = GetMin(MATE0 - Ply, Beta)
  If Alpha >= Beta Then Search = Alpha: Exit Function

  '--- / Step 4a. Tablebase (endgame) : TODO
  ' Tablebase access (switch to 5 men only for web online access)
  If IsTbPos And HashDepth <> TT_TB_BASE_DEPTH Then ' Postion already done and saved in hash?
    Dim sTbFEN As String, TBGameResultScore As Long, sTBBestMove As String, sTBBestMovesList As String
    sTbFEN = WriteEPD()
    If ProbeTablebases(sTbFEN, TBGameResultScore, True, sTBBestMove, sTBBestMovesList) Then
      BestMove = TextToMove(sTBBestMove)
      Search = TBGameResultScore
      InsertIntoHashTable HashKey, TT_TB_BASE_DEPTH, BestMove, TT_EXACT, TBGameResultScore, TBGameResultScore
      If BestMove.From > 0 Then
        If PVNode Then UpdatePV Ply, BestMove '--- Save PV ---
        BestMovePly(Ply) = BestMove
        If TBGameResultScore < -MATE_IN_MAX_PLY Then IsTBScore = True  ' if opp mated don't search for better moves , i.e draws
      End If
      Exit Function
    End If
  End If
        
  '--- / Step 5. Evaluate the position statically
  If PrevMove.IsChecking Then
    StaticEval = UNKNOWN_SCORE: EvalScore = StaticEval
    GoTo lblIID
  ElseIf ttHit Then
    If HashStaticEval = UNKNOWN_SCORE Then StaticEval = Eval() Else StaticEval = HashStaticEval
    EvalScore = StaticEval
        
    If ttValue <> UNKNOWN_SCORE Then
      If ttValue > EvalScore Then
        If (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT) Then EvalScore = ttValue
      Else
        If (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT) Then EvalScore = ttValue
      End If
    End If
  Else
    If StaticEval = UNKNOWN_SCORE Then
      'If bIsNullMove Then  '--- Removed because of asymmetric eval
      '   StaticEval = -StaticEvalArr(Ply - 1) + 2 * TEMPO_BONUS ' Tempo bonus = 20
      'Else
         StaticEval = Eval()
      'End If
    End If
    InsertIntoHashTable HashKey, DEPTH_NONE, EmptyMove, TT_NO_BOUND, UNKNOWN_SCORE, StaticEval
    EvalScore = StaticEval
  End If
    
  StaticEvalArr(Ply) = StaticEval
    
  '--- Check for dangerous moves => do not cut here
  If bSkipEarlyPruning Then GoTo lblMovesLoop
  If IterativeDepth <= 4 Then GoTo lblMovesLoop 'lblNoRazor

  '
  '--- Step 6. Razoring (skipped when in check)
  '
  '    If Not PVNode And Depth < 4 And ttMove.From = 0 Then
  If Not PVNode And Depth < 4 Then
    If EvalScore + RazorMargin(Depth) <= Alpha And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN Then
      'If Not PawnOnRank7() Then
      If Depth <= 1 Then
        Search = QSearch(NON_PV_NODE, Alpha, Alpha + 1, MAX_DEPTH, PrevMove, QS_CHECKS)
        Exit Function
      End If
        
      Dim rAlpha As Long
      rAlpha = Alpha - RazorMargin(Depth)
      Score = QSearch(NON_PV_NODE, rAlpha, rAlpha + 1, MAX_DEPTH, PrevMove, QS_CHECKS)
      If Score <= rAlpha Then
        Search = Score
        Exit Function
      End If
      'End If
    End If
  End If
    
  '
  '--- Step 7. Futility pruning: child node (skipped when in check)
  '
  If Depth < 7 Then
    If (bWhiteToMove And CBool(WNonPawnMaterial > 0)) Or (Not bWhiteToMove And CBool(BNonPawnMaterial > 0)) Then
      If EvalScore < VALUE_KNOWN_WIN And EvalScore - FutilityMargin(Depth, PVNode) >= Beta Then
        Search = EvalScore - FutilityMargin(Depth, PVNode)
        Exit Function
      End If
    End If
  End If
   
lblNoRazor:
    
  '
  '--- Step 8. NULL MOVE ------------
  '
  If Not PVNode And Depth >= 2 And EvalScore >= Beta Then
    If Fifty < 80 And Abs(Beta) < MATE_IN_MAX_PLY And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN Then
     If (StaticEval >= Beta - 35 * (Depth - 6)) Or Depth >= 13 Then
      If (bWhiteToMove And WNonPawnPieces > 0) Or (Not bWhiteToMove And BNonPawnPieces > 0) Then
        '--- Do NULLMOVE ---
        Dim bOldToMove As Boolean
       
        bOldToMove = bWhiteToMove
        bWhiteToMove = Not bWhiteToMove 'MakeNullMove
        bSkipEarlyPruning = True: BestMovePly(Ply + 1) = EmptyMove
        CmhPtr(Ply) = 0
        RemoveEpPiece
        MovesList(Ply) = EmptyMove
        Ply = Ply + 1
        EpPosArr(Ply) = 0: PliesFromNull(Ply) = 0: Fifty = Fifty + 1
        CurrentMove = EmptyMove
        Debug.Assert EvalScore - Beta >= 0
        
        '--- Stockfish
        r = (823 + 67 * Depth) \ 256 + GetMin((EvalScore - Beta) \ ScorePawn.MG, 3) '3 + Depth \ 4 + GetMin((StaticEval - Beta) \ ValueP,3) ' SF6 (problems: WAC 288,200)'
        If Depth - r <= 0 Then
          NullScore = -QSearch(NON_PV_NODE, -Beta, -Beta + 1, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          NullScore = -Search(NON_PV_NODE, -Beta, -Beta + 1, Depth - r, CurrentMove, EmptyMove, Not CutNode)
        End If
        RemoveEpPiece
        Ply = Ply - 1
        ResetEpPiece
        Fifty = Fifty - 1
        
        bSkipEarlyPruning = False
        ' UnMake NullMove
        bWhiteToMove = bOldToMove
        
        If bTimeExit Then Search = 0: Exit Function
            
        If NullScore < -MATE_IN_MAX_PLY Then ' Mate threat : not SF logic
          ThreatMove = BestMovePly(Ply + 1)
          lPlyExtension = 1: GoTo lblMovesLoop
        End If
            
        If NullScore >= Beta Then
          If NullScore >= MATE_IN_MAX_PLY Then NullScore = Beta '  Do not return unproven mate scores
                
          If (Depth < 12 And Abs(Beta) < VALUE_KNOWN_WIN) Then
            Search = NullScore
            Exit Function '--- Return Null Score
          End If
                  
          ' Do verification search at high depths
          bSkipEarlyPruning = True
          If Depth - r <= 0 Then
            Score = QSearch(NON_PV_NODE, Beta - 1, Beta, MAX_DEPTH, CurrentMove, QS_CHECKS)
          Else
            Score = Search(NON_PV_NODE, Beta - 1, Beta, Depth - r, CurrentMove, EmptyMove, False)
          End If
          bSkipEarlyPruning = False
          If Score >= Beta Then
            Search = NullScore
            Exit Function '--- Return Null Score
          End If
                
        End If
           
        '--- Capture Threat?  ( not SF logic )
        If (BestMovePly(Ply + 1).Captured <> NO_PIECE Or NullScore < -MATE_IN_MAX_PLY) Then
          ThreatMove = BestMovePly(Ply + 1)
        End If
      End If
     End If
   End If
  End If

lblNoNullMove:
    
  '--- Step 9. ProbCut (skipped when in check)
  ' If we have a very good capture (i.e. SEE > seeValues[captured_piece_type])
  ' and a reduced search returns a value much above beta, we can (almost) safely prune the previous move.
  If Not PVNode And Depth >= 5 Then
    If Abs(Beta) < MATE_IN_MAX_PLY And Abs(StaticEval) < 2 * VALUE_KNOWN_WIN Then
      rBeta = GetMin(Beta + 200, MATE0)
      rDepth = Depth - 4
      Debug.Assert rDepth >= 1
      Debug.Assert PrevMove.Target > 0
      
      MovePickerInit Ply, EmptyMove, PrevMove, ThreatMove, True, False, GENERATE_ALL_MOVES
      Do While MovePicker(Ply, CurrentMove, LegalMovesOutOfCheck)
        If CurrentMove.Captured <> NO_PIECE Or CurrentMove.Promoted > 0 Then
          If SEEGreaterOrEqual(CurrentMove, PieceAbsValue(CurrentMove.Captured)) Then
          'If SEEGreaterOrEqual(CurrentMove, rBeta - StaticEval) Then
            '--- Make move            -
            RemoveEpPiece
            MakeMove CurrentMove
            Ply = Ply + 1
            bLegalMove = False
            If CheckLegal(CurrentMove) Then
              bLegalMove = True: MovesList(Ply - 1) = CurrentMove
              Score = -Search(NON_PV_NODE, -rBeta, -rBeta + 1, rDepth, CurrentMove, EmptyMove, Not CutNode)
            End If
            '--- Undo move ------------
            RemoveEpPiece
            Ply = Ply - 1
            UnmakeMove CurrentMove
            ResetEpPiece
            If Score >= rBeta And bLegalMove Then
              Search = Score
              Exit Function '---<<< Return
            End If
          End If
        End If
      Loop
    End If
  End If
    
  '--- Step 10. Internal iterative deepening (skipped when in check)
  ' Original depths in SF6: PVNode 5, NonPV: 8. But lower depth are better because of bad move ordering
lblIID:
  If (ttMove.From = 0) And ((PVNode And Depth >= 4) Or (Not PVNode And Depth >= 6)) Then
    If StaticEval = UNKNOWN_SCORE Then StaticEval = Eval()
    If (PVNode Or (CutNode And (StaticEval + 256 >= Beta))) Then
      Depth1 = 3 * Depth \ 4 - 2: If Depth1 <= 0 Then Depth1 = 1
      'Depth1 = Depth - 2: If Not PVNode Then Depth1 = Depth1 - Depth \ 4
      bSkipEarlyPruning = True
      '--- Set BestMovePly(Ply)
      Score = Search(PVNode, Alpha, Beta, Depth1, PrevMove, EmptyMove, CutNode)
      bSkipEarlyPruning = False
        
      ttMove = EmptyMove
      ttHit = IsInHashTable(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)
      If ttHit And HashMove.Target > 0 Then
        ttMove = HashMove
      End If
    End If
  End If
    
  '--- Prepare values for move loop
  If StaticEval = UNKNOWN_SCORE Or StaticEvalArr(Ply - 2) = UNKNOWN_SCORE Or bIsNullMove Or PrevMove.IsChecking Then
    Improv = 1
  Else
    If StaticEval >= StaticEvalArr(Ply - 2) Then Improv = 1 Else Improv = 0
  End If



  '-- SF6: Depth>= 8
  bSingularExtensionNode = (Depth >= 8 And ttMove.From > 0 And ttValue <> UNKNOWN_SCORE And (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT) And ExcludedMove.From = 0 And HashDepth >= Depth - 3)

  bSkipQuiets = False
 
  '----------------------------------------------------
  '---- Step 11. Loop through moves        ------------
  '----------------------------------------------------
lblMovesLoop:
    
  bSkipEarlyPruning = False
  PVLength(Ply) = Ply
  LegalMoveCnt = 0: QuietMoves = 0: MoveCnt = 0
    
  If ttMove.From > 0 Then TryBestMove = ttMove Else TryBestMove = EmptyMove

  ' Init MovePicker
  MovePickerInit Ply, TryBestMove, PrevMove, ThreatMove, False, False, GENERATE_ALL_MOVES
  Score = BestValue
  
  ' Set move history values
  Cmh = CmhPtr(Ply - 1): Cm_Ok = (MovesList(Ply - 1).From > 0)
  Fmh = 0: Fm_Ok = False: FMh2 = 0: F2_Ok = False
  If Ply > 2 Then
    Fmh = CmhPtr(Ply - 2): If MovesList(Ply - 2).From > 0 Then Fm_Ok = True
    If Ply > 4 Then FMh2 = CmhPtr(Ply - 4): If MovesList(Ply - 4).From > 0 Then F2_Ok = True
  End If
  
  '--- Loop over moves
  Do While MovePicker(Ply, CurrentMove, LegalMovesOutOfCheck)
      
    'If ExcludedMove.From > 0 Then If MovesEqual(CurrentMove, ExcludedMove) Then GoTo lblNextMove
    If PrevMove.IsChecking And Not CurrentMove.IsLegal Then GoTo lblNextMove '--- Legal already tested in Ordermoves
    bLegalMove = False: MoveCnt = MoveCnt + 1
    
    'bMoveCountPruning = (Depth < 16 And MoveCnt >= FutilityMoveCounts(Improv, Depth))
    bMoveCountPruning = Depth < 15 And LegalMoveCnt >= (GetMax(0, (MovePickerDat(Ply).EndMoves - 15)) \ 5) + FutilityMoveCounts(Abs(Improv > 0 Or PVNode), GetMax(0, Depth - 1)) + Abs(BestValue = DrawContempt) * 10
    
    bCaptureOrPromotion = (CurrentMove.Captured <> NO_PIECE Or CurrentMove.Promoted <> 0)
    CurrPtr = CurrentMove.Piece * MAX_BOARD + CurrentMove.Target
    CmhPtr(Ply) = CurrPtr
    
    HistVal = UNKNOWN_SCORE
    If Not bCaptureOrPromotion And bMoveCountPruning Then
      If bSkipQuiets And LegalMoveCnt > 0 And Not PrevMove.IsChecking Then
          HistVal = History(PieceColor(CurrentMove.Piece), CurrentMove.From, CurrentMove.Target)
          If Cmh > 0 Then HistVal = HistVal + CounterMoves(Cmh, CurrPtr)
          If Fmh > 0 Then HistVal = HistVal + CounterMoves(Fmh, CurrPtr)
          If FMh2 > 0 Then HistVal = HistVal + CounterMoves(FMh2, CurrPtr)
          If HistVal < 0 Then
            TestCnt(8) = TestCnt(8) + 1
            GoTo lblNextMove
          End If
      End If
    End If
    
    bDoFullDepthSearch = True
            
    lExtension = lPlyExtension
    bKillerMove = IsKiller1Move(Ply, CurrentMove)
    

    '
    '--- Step 12. CHECK EXTENSION ---
    '
    If (CurrentMove.IsChecking) And lExtension = 0 Then
      If SEEGreaterOrEqual(CurrentMove, 0) Then
        lExtension = 1
      End If
    End If
    
    '- check single move escape extent
  '  If (PrevMove.ischecking) And lExtension = 0 Then
  '    If LegalMovesOutOfCheck <= 1 Then
  '      lExtension = 1
  '    End If
  '  End If
  
    '----  Singular extension search.
    '--- SF logic
   ' If bSingularExtensionNode Then
   '   If lExtension = 0 And CurrentMove.From = ttMove.From And CurrentMove.Target = ttMove.Target And CurrentMove.Promoted = ttMove.Promoted Then
   '    If MovePossible(CurrentMove) Then
   '     rBeta = ttValue - 2 * Depth
   '     bSkipEarlyPruning = True
   '     '--- Current move excluded
   '     Score = Search(NON_PV_NODE, rBeta - 1, rBeta, Depth \ 2, False, PrevMove, CurrentMove, 0, CutNode)
   '     bSkipEarlyPruning = False
   '     If Score < rBeta Then
   '       If CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0 And Not bIsNullMove Then
   '         CounterMove(PrevMove.piece, PrevMove.Target) = CurrentMove
   '       End If
   '       lExtension = 1
   '     End If
   '    End If
   '   End If
   ' End If
        
    ' own cheaper logic but less cases
    If bSingularExtensionNode And lExtension = 0 Then
     If bCaptureOrPromotion Then ' Capture or promotion
      If lExtension = 0 And CurrentMove.From = ttMove.From And CurrentMove.Target = ttMove.Target And CurrentMove.Promoted = ttMove.Promoted Then
       If MovePossible(CurrentMove) Then
         If PrevMove.IsChecking And LegalMovesOutOfCheck <= 2 Then
           lExtension = 1: Improv = 1: bSingularExtensionNode = False ' extend this check evasion  move
         ElseIf SEEGreaterOrEqual(CurrentMove, ScorePawn.MG) Then
           lExtension = 1: Improv = 1: bSingularExtensionNode = False ' extend this good move
            'TestCnt(12) = TestCnt(12) + 1 ' output in debug mode
         End If
       End If
      End If
     End If
    End If
        
    NewDepth = GetMax(0, Depth - 1 + lExtension)
         
    '
    '--- Reductions ---------
    '
    '--- Step 13. Pruning at shallow depth
    If BestValue > -MATE_IN_MAX_PLY Then
     If Not bCaptureOrPromotion And Not CurrentMove.IsChecking And Not AdvancedPawnPush(CurrentMove.Piece, CurrentMove.Target) Then
        '--- LMP --- move count based, different formular to SF includes total number of moves and Improv
        If Not bKillerMove And bMoveCountPruning Then
          If BestMovePly(Ply + 1).From > 0 Then ThreatMove = BestMovePly(Ply + 1) ' Threat move not implemented in SF
          If ThreatMove.Captured <> NO_PIECE Or ThreatMove.Promoted <> 0 Or ThreatMove.IsChecking Then
            ' don't skip threat esacpe
            If Not (CurrentMove.From = ThreatMove.Target Or CurrentMove.From = ThreatMove.From) Or IsBlockingMove(ThreatMove, CurrentMove) Then
              'TestCnt(10) = TestCnt(10) + 1
              bSkipQuiets = True
              GoTo lblNextMove  ' not a threat defeat
            End If
          Else
            'TestCnt(10) = TestCnt(10) + 1
            bSkipQuiets = True
            GoTo lblNextMove ' not a threat
          End If
        End If
        
        LmrDepth = GetMax(NewDepth - Reduction(PVNode, Improv, Depth, MoveCnt), 0)
    
        '--- Countermoves based pruning
        If LmrDepth < 3 Then
          If (CounterMoves(Cmh, CurrPtr) < 0 Or Not Cm_Ok) And _
             (CounterMoves(Fmh, CurrPtr) < 0 Or Not Fm_Ok) And _
             ((CounterMoves(FMh2, CurrPtr) < 0 Or (Not F2_Ok) Or (Cm_Ok And Fm_Ok))) Then
             ' TestCnt(11) = TestCnt(11) + 1
           GoTo lblNextMove
          End If
        End If
                       
        '--- Futility pruning: parent node
        If LmrDepth < 7 And Not PrevMove.IsChecking Then
          If StaticEval + 256 + 200 * LmrDepth <= Alpha Then
              'TestCnt(12) = TestCnt(12) + 1
            GoTo lblNextMove
          End If
        End If
        
        '--- SEE based LMP
        If LmrDepth < 8 Then
          If Not SEEGreaterOrEqual(CurrentMove, -35 * LmrDepth * LmrDepth) Then
            'TestCnt(13) = TestCnt(13) + 1
            GoTo lblNextMove
          End If
        End If
        
      Else
        If Depth < 7 And Not CurrentMove.IsChecking And lExtension = 0 Then ' IsChecking better for me, not for SF
          If Not SEEGreaterOrEqual(CurrentMove, -ScorePawn.EG * Depth) Then
            'TestCnt(14) = TestCnt(14) + 1
            GoTo lblNextMove
          End If
        End If
      End If
     
    End If
    
lblMakeMove:
    '--------------------------
    '--- Step 14. Make move   -
    '--------------------------
    RemoveEpPiece
    MakeMove CurrentMove
    Ply = Ply + 1
        
    If CheckLegal(CurrentMove) Then
      Nodes = Nodes + 1: LegalMoveCnt = LegalMoveCnt + 1
      bNoMoves = False: bLegalMove = True
      MovesList(Ply - 1) = CurrentMove
      
      '
      '--- Step 15. Reduced depth search (LMR). If the move fails high it will be re-searched at full depth.
      '
      If Depth >= 3 And LegalMoveCnt > 1 And (Not bCaptureOrPromotion Or bMoveCountPruning) Then
               
        r = Reduction(PVNode, Improv, Depth, MoveCnt)
                
        If bCaptureOrPromotion Then
          If r > 0 Then r = r - 1
        Else
 
          If CutNode Then
            r = r + 2
          ElseIf CurrentMove.Castle = NO_CASTLE Then
            '--- Decrease reduction for moves that escape a capture
            TmpMove.From = CurrentMove.Target: TmpMove.Target = CurrentMove.From: TmpMove.Piece = CurrentMove.Piece: TmpMove.Captured = NO_PIECE: TmpMove.SeeValue = UNKNOWN_SCORE
            ' Move back to old square, were we in danger there?
            If Not SEEGreaterOrEqual(TmpMove, -MAX_SEE_DIFF) Then r = r - 2 ' old square was dangerous
          End If
          
          If HistVal = UNKNOWN_SCORE Then
            HistVal = History(PieceColor(CurrentMove.Piece), CurrentMove.From, CurrentMove.Target)
            If Cmh > 0 Then HistVal = HistVal + CounterMoves(Cmh, CurrPtr)
            If Fmh > 0 Then HistVal = HistVal + CounterMoves(Fmh, CurrPtr)
            If FMh2 > 0 Then HistVal = HistVal + CounterMoves(FMh2, CurrPtr)
          End If
          HistoryVal(Ply) = HistVal - 4000
          
          '--- Decrease/increase reduction by comparing opponent's stat score
          If HistoryVal(Ply) > 0 And HistoryVal(Ply - 1) < 0 Then
            r = r - 1
          ElseIf HistoryVal(Ply) < 0 And HistoryVal(Ply - 1) > 0 Then
            r = r + 1
          End If
          r = r - HistoryVal(Ply) \ 20000
          
          '--- Decrease/increase reduction for moves with a good/bad history
          r = GetMax(0, r)
        End If ' bCaptureOrPromotion
        
        If NewDepth <= 3 And Not bCaptureOrPromotion And CurrentMove.Castle = NO_CASTLE Then
          If PieceType(CurrentMove.Piece) = PT_KING And Not PrevMove.IsChecking Then
            If BNonPawnPieces > 0 Then r = r + 1
          End If
        End If

        Depth1 = GetMax(NewDepth - r, 1)
        
        '--- Reduced SEARCH ---------
        Score = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, Depth1, CurrentMove, EmptyMove, True)
        bDoFullDepthSearch = (Score > Alpha And Depth1 <> NewDepth)
        r = 0
      Else
        bDoFullDepthSearch = (Not PVNode Or LegalMoveCnt > 1)
      End If '  Depth >= 3 ...
            
lblNoMoreReductions:
         
      '------------------------------------------------
      '--->>>>  S E A R C H <<<<-----------------------
      '------------------------------------------------

      If (Alpha > MATE_IN_MAX_PLY And GoodMoves > 0) Or (Ply + Depth + lExtension > MAX_DEPTH) Then lExtension = 0
                  
      NewDepth = GetMax(0, Depth - 1 + lExtension)
            
      '------------------------------------------------
      '--->>>>  R E C U R S I V E  S E A R C H <<<<----
      '------------------------------------------------
      '
      'Step 16. Full depth search when LMR is skipped or fails high
      '
      If bDoFullDepthSearch Then
        If (NewDepth <= 0) Or (Ply >= MAX_DEPTH) Then
          Score = -QSearch(NON_PV_NODE, -(Alpha + 1), -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          Score = -Search(NON_PV_NODE, -(Alpha + 1), -Alpha, NewDepth, CurrentMove, EmptyMove, Not CutNode)
        End If
      End If
            
      ' For PV nodes only, do a full PV search on the first move or after a fail
      ' high (in the latter case search only if value < beta), otherwise let the
      ' parent node fail low with value <= alpha and to try another move.
      If (PVNode And (LegalMoveCnt = 1 Or (Score > Alpha And Score < Beta))) Or Score = UNKNOWN_SCORE Then
        If NewDepth <= 0 Or (Ply >= MAX_DEPTH) Then
          Score = -QSearch(PV_NODE, -Beta, -Alpha, MAX_DEPTH, CurrentMove, QS_CHECKS)
        Else
          Score = -Search(PV_NODE, -Beta, -Alpha, NewDepth, CurrentMove, EmptyMove, False)
        End If
      End If
               
lblSkipMove:
    End If '--- CheckLegal
        
    '--------------------------
    '---  Step 17. Undo move --
    '--------------------------
    RemoveEpPiece
    Ply = Ply - 1
    UnmakeMove CurrentMove
    ResetEpPiece
        
    If bTimeExit Then Search = 0: Exit Function
                
    '-
    '--- Step 18. Check for a new best move
    '-
    If Score > BestValue And bLegalMove Then
      BestValue = Score
        
      If (Score > Alpha) Then
        GoodMoves = GoodMoves + 1
        BestMove = CurrentMove

        If PVNode Then UpdatePV Ply, CurrentMove '--- Save PV ---
            
        If PVNode And Score < Beta Then
          Alpha = Score
        Else
          '--- Fail High  ---
          Exit Do
        End If
            
      End If
    End If
        
        
    If bLegalMove Then
          
      '--- Add Quiet move, used for pruning and history update
      If Not bCaptureOrPromotion And QuietMoves < 64 Then
        If Not MovesEqual(BestMove, CurrentMove) Then QuietMoves = QuietMoves + 1: QuietsSearched(Ply, QuietMoves) = CurrentMove
      End If
    Else
      MoveCnt = MoveCnt - 1
    End If
        
lblNextMove:
  Loop '--- next Move ---

  '
  '--- Step 20. Check for mate and stalemate ---
  '
  If bNoMoves Then
    'If ExcludedMove.From > 0 Then
    '  BestValue = Alpha
    If InCheck() Then '-- do check again to be sure
      BestValue = -MATE0 + Ply ' mate in N plies
    Else
      If CompToMove() Then BestValue = DrawContempt Else BestValue = -DrawContempt
    End If
  
  ElseIf BestMove.From > 0 Then
    ' New best move
    BestMovePly(Ply) = BestMove
    
    If BestMove.Captured = NO_PIECE And BestMove.Promoted = 0 Then
      UpdateStats Ply, BestMove, QuietMoves, PrevMove, StatBonus(Depth)
    End If
    
    ' Extra penalty for a quiet TT move in previous ply when it gets refuted
    If PrevMove.Captured = NO_PIECE Then
      If PrevMove.From > 0 And Ply > 2 And Cmh > 0 Then
        If MovePickerDat(Ply - 1).CurrMoveNum = 0 Then
          UpdateCmStats Ply - 1, PrevMove.Piece, PrevMove.Target, -StatBonus(Depth + 1)
        End If
      End If
    End If
    
  Else
    '--- failed low - no best move
    BestMovePly(Ply) = EmptyMove
    
    ' Bonus for prior countermove that caused the fail low
    If Depth >= 3 Then
      If PrevMove.Captured = NO_PIECE Then
        If Cm_Ok And Ply > 2 Then
          UpdateCmStats Ply - 1, PrevMove.Piece, PrevMove.Target, StatBonus(Depth)
        End If
      End If
    End If
  End If
  
  If Fifty > 100 Then ' Draw ?
    If CompToMove() Then BestValue = DrawContempt Else BestValue = -DrawContempt
  End If
    
  '--- Save Hash values ---
  If BestValue >= Beta Then
    HashEvalType = TT_LOWER_BOUND
  ElseIf PVNode And BestMove.From >= SQ_A1 Then
    HashEvalType = TT_EXACT
  Else
    HashEvalType = TT_UPPER_BOUND
  End If
  InsertIntoHashTable HashKey, Depth, BestMove, HashEvalType, BestValue, StaticEval

  Search = BestValue
End Function
'------------------------------------------------------------------------------------------------------
'- end of SEARCH
'------------------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------------------
' QSearch SF:  search for quiet position until no more capture possible, finally calls position evalution
'           called by SEARCH, calls QSEARCH recursively , then EVAL
'------------------------------------------------------------------------------------------------------
Private Function QSearch(ByVal PVNode As Boolean, _
                         ByVal Alpha As Long, _
                         ByVal Beta As Long, _
                         ByVal Depth As Long, _
                         InPrevMove As TMove, _
                         ByVal GenerateQSChecks As Boolean) As Long

  Dim CurrentMove As TMove, bNoMoves As Boolean, Score As Long, BestMove As TMove
  Dim bLegalMove  As Boolean, PrevMove As TMove, FutilBase As Long, FutilScore As Long, StaticEval As Long
  Dim bPrunable   As Boolean, BestValue As Long, ttDepth As Long
  Dim bHashFound  As Boolean, ttHit As Boolean, HashEvalType As Long, HashScore As Long, HashStaticEval As Long, HashDepth As Long
  Dim HashKey     As THashKey, HashMove As TMove, bCapturesOnly As Boolean, bHashBoardDone As Boolean, OldAlpha As Long
  
  BestMovePly(Ply) = EmptyMove: BestMove = EmptyMove: PrevMove = InPrevMove
  BestValue = UNKNOWN_SCORE: StaticEval = UNKNOWN_SCORE: HashScore = UNKNOWN_SCORE
  QSDepth = QSDepth + 1

  OldAlpha = Alpha
  bHashFound = False: ttHit = False: HashMove = EmptyMove: bHashBoardDone = False
  If Fifty > 3 Then
    HashKey = HashBoard(): bHashBoardDone = True ' Save current keys for insert later
    If Is3xDraw(HashKey, GameMovesCnt, Ply) Then
      If CompToMove() Then QSearch = DrawContempt Else QSearch = -DrawContempt
      QSDepth = QSDepth - 1
      Exit Function ' -- Exit
    End If
  End If
  If Not PrevMove.From = 0 Then GamePosHash(GameMovesCnt + Ply - 1) = HashKey Else GamePosHash(GameMovesCnt + Ply - 1) = EmptyHash

  If (Depth <= 0 Or Ply >= MAX_DEPTH) Then
    QSearch = Eval(): QSDepth = QSDepth - 1
    Exit Function  '-- Exit
  Else
    '--- Check Hash ---------------
    If Not bHashBoardDone Then HashKey = HashBoard() ' Save current keys for insert later
    GamePosHash(GameMovesCnt + Ply - 1) = HashKey

    If PrevMove.IsChecking Or GenerateQSChecks Then
      ttDepth = DEPTH_QS_CHECKS   ' = 0
    Else
      ttDepth = DEPTH_QS_NO_CHECKS ' = -1
    End If
  
    ttHit = IsInHashTable(HashKey, HashDepth, HashMove, HashEvalType, HashScore, HashStaticEval)

    If Not PVNode And ttHit And HashScore <> UNKNOWN_SCORE And HashDepth >= ttDepth Then
      If HashScore >= Beta Then
        bHashFound = (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT)
      Else
        bHashFound = (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT)
      End If

      If bHashFound Then
        QSearch = HashScore: QSDepth = QSDepth - 1
        Exit Function ' -- Exit
      End If
    End If
    If ttHit And HashMove.From > 0 Then BestMovePly(Ply) = HashMove
  
    '-----------------------
  
    If PrevMove.IsChecking Then
      FutilBase = UNKNOWN_SCORE  ':StaticEvalArr(Ply) = UNKNOWN_SCORE:
      bCapturesOnly = False ' search all moves to prove mate
    Else
  
      '--- SEARCH CAPTURES ONLY ----
      If ttHit Then
        If HashStaticEval = UNKNOWN_SCORE Then
          StaticEval = Eval()
        Else
          StaticEval = HashStaticEval
        End If
        BestValue = StaticEval ': StaticEvalArr(Ply) = StaticEval
        
        If HashScore <> UNKNOWN_SCORE Then
          If HashScore > BestValue Then
            If (HashEvalType = TT_LOWER_BOUND Or HashEvalType = TT_EXACT) Then BestValue = HashScore
          Else
            If (HashEvalType = TT_UPPER_BOUND Or HashEvalType = TT_EXACT) Then BestValue = HashScore
          End If
        End If
      Else
        '--- Removed because of asymmetric eval
        'If PrevMove.From = 0 Then ' Nullmove? Can happen at first cal form normal search only
        '  StaticEval = -StaticEvalArr(Ply - 1) '+ 2 * TEMPO_BONUS ' Tempo bonus for nullmove
        'Else
          StaticEval = Eval()
        'End If
        BestValue = StaticEval  ': StaticEvalArr(Ply) = StaticEval
      End If
    
      '--- Stand pat. Return immediately if static value is at least beta
      If BestValue >= Beta Then
        If Not ttHit Then InsertIntoHashTable HashKey, DEPTH_NONE, EmptyMove, TT_LOWER_BOUND, BestValue, StaticEval
        QSearch = BestValue
        QSDepth = QSDepth - 1
        Exit Function '-- exit
      End If
      If PVNode And BestValue > Alpha Then Alpha = BestValue
      FutilBase = BestValue + 128: bCapturesOnly = True ' Captures only
    End If
   
    PVLength(Ply) = Ply: bNoMoves = True
        
    '
    '---- QSearch moves loop ---------------
    '
    If HashMove.From > 0 Then ' Hash move is capture or check ?
      If bCapturesOnly And HashMove.Captured = NO_PIECE Then
        HashMove = EmptyMove
      ElseIf Not GenerateQSChecks And HashMove.IsChecking Then
        HashMove = EmptyMove
      End If
    End If
    
    MovePickerInit Ply, HashMove, PrevMove, EmptyMove, bCapturesOnly, False, GenerateQSChecks
    Do While MovePicker(Ply, CurrentMove, LegalMovesOutOfCheck)
      ' Debug.Print "QS:" & Ply, MoveText(CurrentMove)
      If PrevMove.IsChecking And LegalMovesOutOfCheck = 0 Then
        '--- Mate
        QSearch = -MATE0 + Ply: QSDepth = QSDepth - 1
        Exit Function
      End If
      If PrevMove.IsChecking And Not CurrentMove.IsLegal Then GoTo lblNext
      
    '  If Not PrevMove.IsChecking And QSDepth > 5 And bCapturesOnly Then
    '     ' Only recaptures for QsDepth > 5
    '    If CurrentMove.Target <> PrevMove.Target Then TestCnt(19) = TestCnt(19) + 1: GoTo lblNext
    '  End If
      
      Score = UNKNOWN_SCORE
       
      '--- Futil Pruning
      If Not PrevMove.IsChecking And Not CurrentMove.IsChecking And FutilBase > -VALUE_KNOWN_WIN Then
        If Not AdvancedPawnPush(CurrentMove.Piece, CurrentMove.Target) Then
          FutilScore = FutilBase
          If CurrentMove.Captured <> NO_PIECE Then FutilScore = FutilScore + PieceAbsValue(CurrentMove.Captured)
          
          If FutilScore <= Alpha Then
            BestValue = GetMax(BestValue, FutilScore)
            GoTo lblNext
          End If
            
          If FutilBase <= Alpha Then
            If Not SEEGreaterOrEqual(CurrentMove, 1) Then
              BestValue = GetMax(BestValue, FutilBase)
              GoTo lblNext
            End If
          End If
        End If
      End If
      
      bPrunable = (PrevMove.IsChecking And CurrentMove.Captured = NO_PIECE And BestValue > -MATE_IN_MAX_PLY And QSDepth > 1)
      If (Not PrevMove.IsChecking Or bPrunable) And CurrentMove.Promoted = 0 Then
        If Not SEEGreaterOrEqual(CurrentMove, 0) Then GoTo lblNext
      End If
                  
      '--- Make move -----------------
      RemoveEpPiece
      MakeMove CurrentMove
      Ply = Ply + 1: bLegalMove = False
      If CheckLegal(CurrentMove) Then
        Nodes = Nodes + 1: QNodes = QNodes + 1
        bLegalMove = True: bNoMoves = False
        MovesList(Ply - 1) = CurrentMove
          
        '-------------------------------------
        '--- QSearch recursive  --------------
        '-------------------------------------
        Score = -QSearch(PVNode, -Beta, -Alpha, Depth - 1, CurrentMove, QS_NO_CHECKS)
      End If
      RemoveEpPiece
      Ply = Ply - 1
      UnmakeMove CurrentMove
      ResetEpPiece
        
      If (Score > BestValue) And bLegalMove Then
        BestValue = Score
        If Score > Alpha Then
          If bSearchingPV And PVNode Then UpdatePV Ply, CurrentMove
                
          If PVNode And Score < Beta Then
            Alpha = BestValue
            BestMove = CurrentMove
          Else
            InsertIntoHashTable HashKey, ttDepth, CurrentMove, TT_LOWER_BOUND, Score, StaticEval
            BestMovePly(Ply) = CurrentMove

            '--- Fail high: >= Beta
            QSearch = Score: QSDepth = QSDepth - 1
            Exit Function
          End If
        End If
      End If
lblNext:
    Loop
  End If
  
  '--- Mate?
  If PrevMove.IsChecking And bNoMoves Then
    If InCheck() Then
      QSearch = -MATE0 + Ply ' mate in N plies, check again to be sure
      QSDepth = QSDepth - 1
      Exit Function
    End If
  End If

  QSearch = BestValue
  BestMovePly(Ply) = BestMove

  '--- Save Hash values ---
  If PVNode And BestValue > OldAlpha Then HashEvalType = TT_EXACT Else HashEvalType = TT_UPPER_BOUND
  InsertIntoHashTable HashKey, ttDepth, BestMove, HashEvalType, QSearch, StaticEval
  QSDepth = QSDepth - 1
  
End Function

'---------------------------------------------------------------------------
'- OrderMoves()
'- Assign an order value to the generated move list
'---------------------------------------------------------------------------
Private Sub OrderMoves(ByVal Ply As Long, _
                       ByVal NumMoves As Long, _
                       PrevMove As TMove, _
                       BestMove As TMove, _
                       ThreatMove As TMove, _
                       ByVal bCapturesOnly As Boolean, _
                       LegalMovesOutOfCheck As Long)
  Dim i               As Long, From As Long, Target As Long, Promoted As Long, Captured As Long, lValue As Long, Piece As Long
  Dim bSearchingPVNew As Boolean, BestValue As Long, BestIndex As Long, WhiteMoves As Boolean, Cmh As Long
  Dim bLegalsOnly     As Boolean, TmpVal As Long, PieceVal As Long, CounterMoveTmp As TMove, KingLoc As Long, v As Long
  Dim Fmh As Long, Fm2 As Long, CurrPtr As Long, bIsChecking As Boolean
  
  LegalMovesOutOfCheck = 0
  If NumMoves = 0 Then Exit Sub
  bSearchingPVNew = False
  BestValue = -9999999: BestIndex = -1 '--- save highest score
  WhiteMoves = CBool(Board(Moves(Ply, 0).From) Mod 2 = 1) ' to be sure to have correct side ...
  Killer0 = Killer(Ply)
  If Ply > 2 Then Killer2 = Killer(Ply - 2) Else Killer2 = EmptyKiller
  
  bLegalsOnly = PrevMove.IsChecking And Not bCapturesOnly ' Count legal moves in normal search (not in QSearch)
  If bWhiteToMove Then KingLoc = WKingLoc Else KingLoc = BKingLoc
  Cmh = PrevMove.Piece * MAX_BOARD + PrevMove.Target
  If Ply > 2 Then Fmh = CmhPtr(Ply - 2) Else Fmh = 0
  If Ply > 4 Then Fm2 = CmhPtr(Ply - 4) Else Fm2 = 0
  
  CounterMoveTmp = CounterMove(PrevMove.Piece, PrevMove.Target)
    
  For i = 0 To NumMoves - 1
    With Moves(Ply, i)
      From = .From: Target = .Target: Promoted = .Promoted: Captured = .Captured: Piece = .Piece
      .IsLegal = False: .IsChecking = False: .SeeValue = UNKNOWN_SCORE
    End With
    lValue = 0
     
    ' Count legal moves if in check
    If bLegalsOnly Then
      If Moves(Ply, i).Castle = NO_CASTLE Then ' castling not allowed in check
        ' Avoid costly legal proof for moves with cannot be a check evasion
        If From <> KingLoc And PieceType(Captured) <> PT_KNIGHT And Not SameXRay(From, KingLoc) And Not SameXRay(Target, KingLoc) Then
          ' ignore
          'TestCnt(18) = TestCnt(18) + 1
        Else
          ' Make move
          RemoveEpPiece
          MakeMove Moves(Ply, i)
          'TestCnt(19) = TestCnt(19) + 1
          
          If CheckEvasionLegal() Then Moves(Ply, i).IsLegal = True: LegalMovesOutOfCheck = LegalMovesOutOfCheck + 1
          
          ' UnMake
          UnmakeMove Moves(Ply, i)
          ResetEpPiece
        End If
      End If
      
      If Moves(Ply, i).IsLegal Then
        lValue = lValue + 3 * MATE0 '- Out of check moves
      Else
        lValue = -999999
        GoTo lblIgnoreMove
      End If
    End If
    
    PieceVal = PieceAbsValue(Piece)
    
    '--- Is Move checking ?
    bIsChecking = Moves(Ply, i).IsChecking
    If Not bIsChecking Then bIsChecking = IsCheckingMove(Piece, From, Target, Promoted)
    If bIsChecking Then
      If Not bCapturesOnly Then
        If Captured = NO_PIECE Then lValue = lValue + 9000 ' 11000
      Else
        lValue = lValue + 800 '  in QSearch search captures first??
      End If
      lValue = lValue + PieceVal \ 6
      If Ply > 2 Then
        If MovesList(Ply - 2).IsChecking Then lValue = lValue + 500 ' Repeated check
      End If
      Moves(Ply, i).IsChecking = True
    End If
        
        
    'bonus  pv:
    If bSearchingPV Then
      If From = PV(1, Ply).From And Target = PV(1, Ply).Target And Promoted = PV(1, Ply).Promoted Then
        bSearchingPVNew = True: lValue = lValue + 2 * MATE0 ' Highest score
        GoTo lblNextMove
      End If
    End If
    
    If ThreatMove.From <> 0 Then
      If Target = ThreatMove.From Then
        lValue = lValue + 600  ' Try capture
      End If
      If From = ThreatMove.Target Then ' Try escape
       If PieceVal > PieceAbsValue(Board(ThreatMove.From)) + 80 Then
         lValue = lValue + 4000 + (PieceVal - PieceAbsValue(Board(ThreatMove.From))) \ 2
       Else
         lValue = lValue + 2000 + PieceVal \ 4
       End If
      End If
    End If
        
    '--- Capture bonus
    If Captured <> NO_PIECE Then
      '-- Captures
      If Not bEndgame Then
        If bWhiteToMove Then lValue = lValue - 100 * Rank(Target) Else lValue = lValue - 100 * (9 - Rank(Target))
      End If
      
      If Piece = WKING Or Piece = BKING Then
        TmpVal = PieceAbsValue(Captured) ' cannot be defended because legal move
      Else
        TmpVal = (PieceAbsValue(Captured) - PieceVal) \ 2
      End If
      If TmpVal > MAX_SEE_DIFF Then
        '--- Winning capture
        lValue = lValue + TmpVal * 10 + 6000
      ElseIf TmpVal > -MAX_SEE_DIFF Then
        '--- Equal capture
        lValue = lValue + PieceAbsValue(Captured) - PieceVal \ 2 + 800
      Else
        '--- Loosing capture? Check with SEE later in MovePicker
        lValue = lValue + PieceAbsValue(Captured) \ 2 - PieceVal
      End If
      
      If Target = PrevMove.Target Then lValue = lValue + 250 ' Recapture
  
      '-- King attack?
      If WhiteMoves Then
        If Piece <> WPAWN Then If MaxDistance(Target, BKingLoc) = 1 Then lValue = lValue + PieceVal \ 2 + 150
      Else
        If Piece <> BPAWN Then If MaxDistance(Target, WKingLoc) = 1 Then lValue = lValue + PieceVal \ 2 + 150
      End If
    Else
      '
      '--- Not a Capture, substract 10000 to select captures first
      '
      If Not bCapturesOnly Then lValue = lValue + MOVE_ORDER_QUIETS
                  
      If PrevMove.Target <> 0 Then
        If CounterMoveTmp.Target = Target Then
          lValue = lValue + 250 ' Bonus for Countermove
          If CounterMoveTmp.Piece = Piece Then lValue = lValue + 250 - PieceVal \ 20
        End If
      End If
      
      'bonus per killer move:
      If From = Killer0.Killer1.From And Target = Killer0.Killer1.Target Then
        lValue = lValue + 3000: GoTo lblKillerDone
      ElseIf From = Killer0.Killer2.From And Target = Killer0.Killer2.Target Then
        lValue = lValue + 2500: GoTo lblKillerDone
      ElseIf From = Killer0.Killer3.From And Target = Killer0.Killer3.Target Then
        lValue = lValue + 2200: GoTo lblKillerDone
    End If
                 
      If Ply > 2 Then '--- killer bonus for previous move of same color
        If From = Killer2.Killer1.From And Target = Killer2.Killer1.Target Then
          lValue = lValue + 2700 ' !!! better!?! 300
        ElseIf From = Killer2.Killer2.From And Target = Killer2.Killer2.Target Then
          lValue = lValue + 200
        End If ' Killer3 not better
      End If
lblKillerDone:
    End If
                 
    '--- value for piece square table  difference of move
    lValue = lValue + PieceAbsValue(Promoted) \ 2 + (PsqVal(Abs(bEndgame), Piece, Target) - PsqVal(Abs(bEndgame), Piece, From)) * 2 ' * (PieceVal \ 100))
              
    '--- Attacked by pawn or pawn push?
    If WhiteMoves Then
      If Piece = WPAWN Then
        If Rank(Target) >= 6 Then If AdvancedPawnPush(Piece, Target) Then lValue = lValue + 250
      Else
        If ((Board(Target + 9) = BPAWN Or Board(Target + 11)) = BPAWN) Then lValue = lValue - PieceVal \ 4    '--- Attacked by Pawn
        If ((Board(Target - 9) = WPAWN Or Board(Target - 11)) = WPAWN) Then lValue = lValue + 50 + PieceVal \ 8    '--- Defended by Pawn
        TmpVal = MaxDistance(Target, BKingLoc): lValue = lValue - TmpVal * TmpVal ' closer to opp king
      End If
    Else
      If Piece = BPAWN Then
        If Rank(Target) <= 3 Then If AdvancedPawnPush(Piece, Target) Then lValue = lValue + 250
      Else
        If ((Board(Target - 9) = WPAWN Or Board(Target - 11) = WPAWN)) Then lValue = lValue - PieceVal \ 4  '--- Attacked by Pawn
        If ((Board(Target + 9) = BPAWN Or Board(Target + 11)) = BPAWN) Then lValue = lValue + 50 + PieceVal \ 8    '--- Defended by Pawn
        TmpVal = MaxDistance(Target, WKingLoc): lValue = lValue - TmpVal * TmpVal ' closer to opp king
      End If
    End If
               
    If PrevMove.IsChecking Then
      If Piece = WKING Or Piece = BKING Then lValue = lValue + 200  ' King check escape move?
      If Target = PrevMove.Target Then lValue = lValue + 200 ' Capture checking piece?
    End If
      
    
    ' CounterMovesHist
    If PrevMove.Target > 0 And Captured = NO_PIECE And Promoted = 0 And Not PrevMove.IsChecking Then
      CurrPtr = Piece * MAX_BOARD + Target
      v = History(PieceColor(Piece), From, Target) + _
          CounterMoves(Cmh, CurrPtr) + CounterMoves(Fmh, CurrPtr) + CounterMoves(Fm2, CurrPtr)
      'If v > TestCnt(2) Then TestCnt(2) = v
      lValue = lValue + v \ 6 ' bonus per history heuristic: Caution: big effects!
    End If
                
lblNextMove:
    '--- Hashmove
    If BestMove.From = From And BestMove.Target = Target Then
      lValue = lValue + MATE0 \ 2
    ElseIf BestMovePly(Ply).From = From And BestMovePly(Ply).Target = Target Then
      '--- Move from Internal Iterative Depening
      lValue = lValue + MATE0 \ 2
    End If
        
    If lValue > BestValue Then BestValue = lValue: BestIndex = i '- save best for first move
lblIgnoreMove:
    Moves(Ply, i).OrderValue = lValue
  Next '---- Move
    
  bSearchingPV = bSearchingPVNew
    
  'Debug:  for i=0 to nummoves-1: Debug.Print i,Moves(ply,i).ordervalue, MoveText(Moves(ply,i)):next
  If BestIndex > 0 Then
    ' Swap best move to top
    TempMove = Moves(Ply, 0): Moves(Ply, 0) = Moves(Ply, BestIndex): Moves(Ply, BestIndex) = TempMove
  End If

End Sub

'------------------------------------------------------------------------------------
' BestMoveAtFirst: get best move from generated move list, scored by OrderMoves.
'                  Faster than SortMoves if alpha/beta cut in the first moves
'------------------------------------------------------------------------------------
Public Sub BestMoveAtFirst(ByVal Ply As Long, _
                           ByVal StartIndex As Long, _
                           ByVal NumMoves As Long)
  Dim TempMove As TMove, i As Long, MaxScore As Long, MaxPtr As Long
  If StartIndex = NumMoves Then Exit Sub
  MaxScore = -9999999
  MaxPtr = StartIndex
  For i = StartIndex To NumMoves
    If Moves(Ply, i).OrderValue > MaxScore Then MaxScore = Moves(Ply, i).OrderValue: MaxPtr = i
  Next i
  If MaxPtr > StartIndex Then
    TempMove = Moves(Ply, StartIndex): Moves(Ply, StartIndex) = Moves(Ply, MaxPtr): Moves(Ply, MaxPtr) = TempMove
  End If
  
 ' For i = StartIndex To NumMoves
 '   If Moves(Ply, StartIndex - 1).OrderValue < Moves(Ply, i - 1).OrderValue Then Stop
 ' Next
End Sub

' Stable sort
Private Sub SortMovesStable(ByVal Ply As Long, _
                       ByVal iStart As Long, _
                       ByVal iEnd As Long)
                       
Dim i As Long, j As Long, iMin As Long, IMax As Long, TempMove As TMove

iMin = iStart + 1: IMax = iEnd
i = iMin: j = i + 1
Do While i <= IMax
    If Moves(Ply, i).OrderValue > Moves(Ply, i - 1).OrderValue Then
      TempMove = Moves(Ply, i): Moves(Ply, i) = Moves(Ply, i - 1): Moves(Ply, i - 1) = TempMove ' Swap
      If i > iMin Then i = i - 1
    Else
      i = j: j = j + 1
    End If
Loop
'For i = iStart To iEnd - 1 ' Check sort order
' If Moves(Ply, i).OrderValue < Moves(Ply, i + 1).OrderValue Then Stop
'Next

End Sub

'
'--- init move list
'
Public Function MovePickerInit(ByVal ActPly As Long, _
                               BestMove As TMove, _
                               PrevMove As TMove, _
                               ThreatMove As TMove, _
                               ByVal bCapturesOnly As Boolean, _
                               ByVal bMovesGenerated As Boolean, _
                               ByVal bGenerateQSChecks As Boolean)
  With MovePickerDat(ActPly)
    .CurrMoveNum = 0
    .EndMoves = 0
    .BestMove = BestMove
    .bBestMoveChecked = False
    .bBestMoveDone = False
    .PrevMove = PrevMove
    .ThreatMove = ThreatMove
    .bCapturesOnly = bCapturesOnly
    .bMovesGenerated = bMovesGenerated
    .LegalMovesOutOfCheck = -1
    If bGenerateQSChecks Then .GenerateQSChecksCnt = 1 Else .GenerateQSChecksCnt = 0
  End With
End Function

Public Function MovePicker(ByVal ActPly As Long, _
                           Move As TMove, _
                           LegalMovesOutOfCheck As Long) As Boolean
  '
  '-- Returns next move in "Move"  or function returns false if no more moves
  '
  Dim SeeVal As Long, NumMovesPly As Long, BestMove As TMove, bBestMoveDone As Boolean
  
  MovePicker = False: LegalMovesOutOfCheck = 0
  With MovePickerDat(ActPly)
    ' First: try BestMove. If Cutoff then no move generation needed.
    BestMove = .BestMove: bBestMoveDone = .bBestMoveDone
    If Not .bBestMoveChecked Then
      .bBestMoveChecked = True
      If .BestMove.From > 0 And Not .PrevMove.IsChecking Then ' Check: Generate all out of check move, LegalMovesOutOfCheck needed
        If MovePossible(BestMove) Then
          Move = BestMove: .bBestMoveDone = True: MovePicker = True: Move.OrderValue = 5 * MATE0
          If bSearchingPV And Move.From = PV(1, ActPly).From And Move.Target = PV(1, ActPly).Target And Move.Promoted = PV(1, ActPly).Promoted Then
            ' keep SearchingPV
          Else
            bSearchingPV = False
          End If
          Exit Function '--- return best move before move generation
        End If
      End If
    End If
    
    If Not .bMovesGenerated Then
      ' Generate all moves
      GenerateMoves ActPly, .bCapturesOnly, .EndMoves
      ' Order moves
      OrderMoves ActPly, .EndMoves, .PrevMove, .BestMove, .ThreatMove, .bCapturesOnly, .LegalMovesOutOfCheck
      .bMovesGenerated = True: .GenerateQSChecksCnt = 0: .CurrMoveNum = 0
    End If
    LegalMovesOutOfCheck = .LegalMovesOutOfCheck
  
    .CurrMoveNum = .CurrMoveNum + 1  '  array index starts at 0 = nummoves-1
    
    ' ignore Hash move
    If bBestMoveDone And MovesEqual(BestMove, Moves(ActPly, .CurrMoveNum - 1)) Then .CurrMoveNum = .CurrMoveNum + 1
    
    NumMovesPly = .EndMoves
    If NumMovesPly <= 0 Or .CurrMoveNum > NumMovesPly Then Move = EmptyMove: Exit Function
    
    If .CurrMoveNum = 1 Then
      Move = Moves(ActPly, 0):  MovePicker = True: Exit Function  ' First move is already sorted to top in OrderMoves
    Else
      BestMoveAtFirst ActPly, .CurrMoveNum - 1, NumMovesPly - 1
    End If
  
  'If .CurrMoveNum = 2 Then
  '  SortMovesQS Ply, 1, NumMovesPly - 1 ' Sort rest of moves
  'End If

  Do
    Move = Moves(ActPly, .CurrMoveNum - 1)
    If .CurrMoveNum >= NumMovesPly Or (Not Move.IsChecking And Move.Captured = NO_PIECE) Or Move.OrderValue < MOVE_ORDER_BAD_CAPTURES + 5000 Or Move.OrderValue > 1000 Then
      MovePicker = True: Exit Function ' Last move
    End If
    If PieceAbsValue(Move.Captured) - PieceAbsValue(Move.Piece) < -MAX_SEE_DIFF Then
       '-- Bad capture?
       SeeVal = GetSEE(Move): Move.SeeValue = SeeVal ' Slow! Delay the costly SEE until this move is needed - may be not needed if cutoffs earlier
       Moves(ActPly, .CurrMoveNum - 1).SeeValue = SeeVal  ' Save for later use
       If SeeVal >= -MAX_SEE_DIFF Then
         MovePicker = True: Exit Function
       Else
         Move.OrderValue = MOVE_ORDER_BAD_CAPTURES + SeeVal * 5 ' negative See!  - Set to fit condition above < -15000
         '- to avoid new list sort: append this bad move to the end of the move list (add new record), skip current list entry
         'Moves(ActPly, .CurrMoveNum  - 1).From = 0 ' Delete move in list,not needed ??
         NumMovesPly = NumMovesPly + 1: MovePickerDat(ActPly).EndMoves = NumMovesPly: Moves(ActPly, NumMovesPly - 1) = Move
       End If
    Else
      MovePicker = True: Exit Function  ' good captures
    End If
    .CurrMoveNum = .CurrMoveNum + 1
  Loop
End With
 
End Function

Public Function CompToMove() As Boolean
  If bCompIsWhite Then CompToMove = bWhiteToMove Else CompToMove = Not bWhiteToMove
End Function

Private Function FixedDepthMode() As Boolean
  '--- if no time limit use depth limit
  FixedDepthMode = CBool(FixedDepth <> NO_FIXED_DEPTH)
End Function

Public Function IsAnyLegalMove(ByVal NumMoves As Long) As Boolean
  ' Count legal moves
  Dim i As Long
  IsAnyLegalMove = False
  For i = 0 To NumMoves - 1
    RemoveEpPiece
    MakeMove Moves(Ply, i)
    If CheckLegal(Moves(Ply, i)) Then IsAnyLegalMove = True
    UnmakeMove Moves(Ply, i)
    ResetEpPiece
    If IsAnyLegalMove = True Then Exit Function
  Next i
End Function

'
'--- Check 3xRepetion Draw in current moves (only equal from-target combinations)
'
Public Function Is3xDraw(HashKey As THashKey, _
                         ByVal GameMoves As Long, _
                         ByVal SearchPly As Long) As Boolean
  Dim i As Long, Repeats As Long, EndPos As Long, StartPos As Long
  Is3xDraw = False
  If Fifty < 4 Or PliesFromNull(Ply) < 4 Then Exit Function
  If SearchPly > 1 Then SearchPly = SearchPly - 1
  Repeats = 0
  StartPos = GetMax(0, GameMoves + SearchPly - 1)
  If CompToMove Then
    EndPos = GetMax(0, GameMoves + SearchPly - GetMin(Fifty, PliesFromNull(Ply)))
  Else
    EndPos = GetMax(0, GameMoves + SearchPly - GetMin(Fifty - 1, PliesFromNull(Ply) - 1))
  End If
  If StartPos - EndPos < 2 Then Exit Function
  For i = StartPos To EndPos Step -1
    If HashKey.HashKey1 = GamePosHash(i).HashKey1 Then
      If HashKey.HashKey2 = GamePosHash(i).HashKey2 And HashKey.HashKey1 <> 0 Then
        ' 1 repeated position in search=>Draw; or 1 in game plus 1 in search(except root) = 2 => draw
        If i > GameMoves Or SearchPly > 1 Then Is3xDraw = True: Exit Function
        Repeats = Repeats + 1
        If Repeats >= 2 Then Is3xDraw = True: Exit Function
      End If
    End If
  Next i
End Function

Public Function InitRazorMargin()
  ' margin based on depth
  RazorMargin(1) = 483
  RazorMargin(2) = 570
  RazorMargin(3) = 603
  RazorMargin(4) = 554
End Function

'Private Function RazorMargin(ByVal iDepth As Long) As Long
'  RazorMargin = 512& + 32& * CLng(iDepth)
'End Function

Private Function IsKillerMove(ByVal ActPly As Long, Move As TMove) As Boolean
  
  IsKillerMove = True
  With Killer(ActPly)
    If Move.From = .Killer1.From And Move.Target = .Killer1.Target Then Exit Function
    If Move.From = .Killer2.From And Move.Target = .Killer2.Target Then Exit Function
    If Move.From = .Killer3.From And Move.Target = .Killer3.Target Then Exit Function
  End With
  IsKillerMove = False
End Function

Private Function IsKiller1Move(ByVal ActPly As Long, Move As TMove) As Boolean
  
  IsKiller1Move = True
  With Killer(ActPly)
    If Move.From = .Killer1.From And Move.Target = .Killer1.Target Then Exit Function
  End With
  IsKiller1Move = False
End Function

Public Sub InitFutilityMoveCounts()
  Dim d As Single
  For d = 0 To 15
   ' FutilityMoveCounts(0, d) = Int(2.4 + 0.773 * ((CDbl(d) + 0#) ^ 1.8))
    FutilityMoveCounts(0, d) = Int(2.4 + 0.74 * ((CDbl(d) + 0#) ^ 1.78))
    FutilityMoveCounts(1, d) = 5 + d * (d - 1)
    'Debug.Print d, FutilityMoveCounts(0, d), FutilityMoveCounts(1, d)
  Next d
  
End Sub

Public Function FutilityMargin(ByVal iDepth As Long, ByVal PVNode As Boolean) As Long
  FutilityMargin = 150& * CLng(iDepth)
End Function

Public Sub InitReductionArray()
  '  Init reductions array
  Dim d As Long, mc As Long, Improv As Long, r As Double

  For Improv = 0 To 1
    For d = 1 To 63
      For mc = 1 To 63
      
        r = Log(CDbl(d)) * Log(CDbl(mc)) / 1.95
      
        Reductions(0, Improv, d, mc) = Round(r) ' 0=NonPV
        Reductions(1, Improv, d, mc) = GetMaxDbl(Reductions(0, Improv, d, mc) - 1, 0) ' 1=PV
          
        ' Increase reduction when eval is not improving
        If Improv = 0 And Reductions(0, Improv, d, mc) >= 2 Then
          Reductions(0, Improv, d, mc) = Reductions(0, Improv, d, mc) + 1
        End If
      Next mc
    Next d
  Next Improv
End Sub


Private Function Reduction(PVNode As Boolean, _
                           Improving As Long, _
                           Depth As Long, _
                           MoveNumber As Long) As Long
  Dim lPV As Long
  If PVNode Then lPV = 1 Else lPV = 0
  Reduction = Reductions(lPV, Improving, GetMin(Depth, 63), GetMin(MoveNumber, 63))
End Function

Private Function UpdateStats(ByVal ActPly As Long, CurrentMove As TMove, _
                                  ByVal QuietMoveCounter As Long, _
                                  PrevMove As TMove, _
                                  ByVal Bonus As Long)
  '
  '--- Update Killer moves and History-Score
  '
  Dim j As Long
 
  Debug.Assert (CurrentMove.Captured = NO_PIECE And CurrentMove.Promoted = 0)
 
  '--- Killers
  With Killer(Ply)
    If .Killer1.From <> CurrentMove.From Or .Killer1.Target <> CurrentMove.Target Or .Killer1.Piece <> CurrentMove.Piece Then
      .Killer3 = .Killer2: .Killer2 = .Killer1: .Killer1 = CurrentMove
    End If
  End With
                                
  UpdHistory CurrentMove.Piece, CurrentMove.From, CurrentMove.Target, Bonus
  UpdateCmStats ActPly, CurrentMove.Piece, CurrentMove.Target, Bonus
      
  If PrevMove.From >= SQ_A1 And PrevMove.Captured = NO_PIECE Then
    '--- CounterMove:
    CounterMove(PrevMove.Piece, PrevMove.Target) = CurrentMove
  End If
  
  '--- Decrease History for previous tried quiet moves that did not cut off
  For j = 1 To QuietMoveCounter
    With QuietsSearched(Ply, j)
     If .From = CurrentMove.From And .Target = CurrentMove.Target And .Piece = CurrentMove.Piece Then
      ' ignore
     Else
      UpdHistory .Piece, .From, .Target, -Bonus
      If PrevMove.Target > 0 Then UpdateCmStats ActPly, .Piece, .Target, -Bonus
     End If
    End With
  Next j
  
End Function

Public Sub UpdHistory(ByVal Piece As Long, ByVal From As Long, ByVal Target As Long, ByVal ScoreVal As Long)
  If Abs(ScoreVal) >= 324 Then Exit Sub
  History(PieceColor(Piece), From, Target) = History(PieceColor(Piece), From, Target) - (History(PieceColor(Piece), From, Target) * Abs(ScoreVal) \ 324) + (ScoreVal * 32)
End Sub

Public Sub UpdateCmStats(ByVal ActPly As Long, ByVal Piece As Long, ByVal Square As Long, ByVal Bonus As Long)
 
 If ActPly > 1 Then
   If MovesList(ActPly - 1).From > 0 Then
     UpdateCmVal MovesList(ActPly - 1).Piece, MovesList(ActPly - 1).Target, Piece, Square, Bonus
   End If
   If ActPly > 2 Then
     If MovesList(ActPly - 2).From > 0 Then
       UpdateCmVal MovesList(ActPly - 2).Piece, MovesList(ActPly - 2).Target, Piece, Square, Bonus
     End If
   End If
   If ActPly > 4 Then
     If MovesList(ActPly - 4).From > 0 Then
       UpdateCmVal MovesList(ActPly - 4).Piece, MovesList(ActPly - 4).Target, Piece, Square, Bonus
     End If
   End If
 End If
End Sub

Public Sub UpdateCmVal(ByVal PrevPiece As Long, ByVal PrevSquare As Long, ByVal Piece As Long, ByVal Square As Long, ByVal ScoreVal As Long)
  If Abs(ScoreVal) >= 324 Then Exit Sub
  Dim PrevPtr As Long, CurrPtr As Long
  PrevPtr = PrevPiece * MAX_BOARD + PrevSquare: CurrPtr = Piece * MAX_BOARD + Square
  CounterMoves(PrevPtr, CurrPtr) = CounterMoves(PrevPtr, CurrPtr) - CounterMoves(PrevPtr, CurrPtr) * (Abs(ScoreVal)) \ 936 + ScoreVal * 32
End Sub

Public Sub UpdatePV(ByVal ActPly As Long, Move As TMove)
  Dim j As Long
 
  PV(ActPly, ActPly) = Move
  If PVLength(ActPly + 1) > 0 Then
    For j = ActPly + 1 To PVLength(ActPly + 1) - 1
      PV(ActPly, j) = PV(ActPly + 1, j)
    Next
    PVLength(ActPly) = PVLength(ActPly + 1)
  End If
End Sub


Public Function MovePossible(Move As TMove) As Boolean
  ' for test of HashMove before move generation
  Dim Offset As Long, sq As Long, Diff As Long, AbsDiff As Long, OldPiece As Long
  MovePossible = False

  OldPiece = Move.Piece: If Move.Promoted > 0 Then OldPiece = Board(Move.From)
  If Move.From < SQ_A1 Or Move.From > SQ_H8 Or OldPiece < 1 Or Move.From = Move.Target Or OldPiece = NO_PIECE Then Exit Function
  If Board(Move.Target) = FRAME Then Exit Function
  If Board(Move.From) <> OldPiece Then Exit Function
  If Move.Captured <> NO_PIECE Then If Board(Move.Target) <> Move.Captured Then Exit Function
  
  If bWhiteToMove Then
    If OldPiece Mod 2 <> 1 Then Exit Function
  Else

    If OldPiece Mod 2 <> 0 Then Exit Function
  End If
 
  If Board(Move.Target) <> NO_PIECE Then
    If Board(Move.Target) Mod 2 = OldPiece Mod 2 Then Exit Function  ' same color
  End If

  Diff = Move.Target - Move.From: AbsDiff = Abs(Diff)
 
  If PieceType(OldPiece) = PT_PAWN Then
    If (AbsDiff = 9 Or AbsDiff = 11) And Board(Move.Target) = NO_PIECE Then Exit Function
    If AbsDiff = 10 And Board(Move.Target) <> NO_PIECE Then Exit Function
    If AbsDiff = 20 Then
      If Board(Move.From + 10 * Sgn(Diff)) <> NO_PIECE Then Exit Function
      If Board(Move.Target) <> NO_PIECE Then Exit Function
    End If

    MovePossible = True
    Exit Function
  ElseIf OldPiece = WKNIGHT Or OldPiece = BKNIGHT Then

    ' Knight
    Select Case AbsDiff
    Case 8, 12, 19, 21
        MovePossible = True ' OK
    End Select

    Exit Function
  ElseIf OldPiece = WKING Then

    ' WKing: Castling
    If AbsDiff = 2 Then
      If Move.From <> WKING_START Or Moved(WKING_START) > 0 Then Exit Function
      If Diff = 2 Then
        If Board(Move.From + 1) <> NO_PIECE Or Board(Move.From + 2) <> NO_PIECE Or Board(Move.From + 3) <> WROOK Then Exit Function
      ElseIf Diff = -2 Then
        If Board(Move.From - 1) <> NO_PIECE Or Board(Move.From - 2) <> NO_PIECE Or Board(Move.From - 3) <> NO_PIECE Or Board(Move.From - 4) <> WROOK Then Exit Function
      End If
    End If

    MovePossible = True
    Exit Function
  ElseIf OldPiece = BKING Then

    ' BKing: Castling
    If AbsDiff = 2 Then
      If Move.From <> BKING_START Or Moved(BKING_START) > 0 Then Exit Function
      If Diff = 2 Then
        If Board(Move.From + 1) <> NO_PIECE Or Board(Move.From + 2) <> NO_PIECE Or Board(Move.From + 3) <> BROOK Then Exit Function
      ElseIf Diff = -2 Then
        If Board(Move.From - 1) <> NO_PIECE Or Board(Move.From - 2) <> NO_PIECE Or Board(Move.From - 3) <> NO_PIECE Or Board(Move.From - 4) <> BROOK Then Exit Function
      End If
    End If

    MovePossible = True
    Exit Function
  End If

  '--- Sliding piece blocked?
  If MaxDistance(Move.From, Move.Target) > 1 Then
    If AbsDiff Mod 9 = 0 Then
      Offset = Sgn(Diff) * 9
    ElseIf AbsDiff Mod 11 = 0 Then
      Offset = Sgn(Diff) * 11
    ElseIf AbsDiff Mod 10 = 0 Then
      Offset = Sgn(Diff) * 10
    Else
      Offset = Sgn(Diff) * 1
    End If
 
    Select Case OldPiece
      Case WROOK, BROOK:
        If Abs(Offset) <> 1 And Abs(Offset) <> 10 Then Exit Function
      Case WBISHOP, BBISHOP:
        If Abs(Offset) <> 9 And Abs(Offset) <> 11 Then Exit Function
      Case WQUEEN, BQUEEN:
        If Abs(Offset) <> 1 And Abs(Offset) <> 10 And Abs(Offset) <> 9 And Abs(Offset) <> 11 Then Exit Function
    End Select
  
    For sq = Move.From + Offset To Move.Target - Offset Step Offset
      If Board(sq) < NO_PIECE Then Exit Function
    Next
  End If
  MovePossible = True
End Function

Public Function PawnOnRank7() As Boolean
  ' check if side to move has a pawn on relative rank 7
  Dim i As Long
  
  If bWhiteToMove Then
    For i = SQ_A7 To SQ_H7
      If Board(i) = WPAWN Then PawnOnRank7 = True: Exit Function
    Next
  Else
    For i = SQ_A2 To SQ_H2
      If Board(i) = BPAWN Then PawnOnRank7 = True: Exit Function
    Next
  End If

  PawnOnRank7 = False
End Function

Public Sub ClearEasyMove()
 If bTimeTrace Then WriteTrace "Clear EasyMovePV"
 EasyMovePV(1) = EmptyMove: EasyMovePV(2) = EmptyMove: EasyMovePV(3) = EmptyMove
 EasyMoveStableCnt = 0
End Sub

Public Sub UpdateEasyMove()
   Dim i As Long, bDoUpdate As Boolean
   If MovesEqual(PV(1, 3), EasyMovePV(3)) Then
     EasyMoveStableCnt = EasyMoveStableCnt + 1
   Else
     EasyMoveStableCnt = 0
   End If
   
   bDoUpdate = False
   For i = 1 To 3
     If PV(1, i).From > 0 Then If Not MovesEqual(EasyMovePV(i), PV(1, i)) Then bDoUpdate = True
   Next
   If bDoUpdate Then
     For i = 1 To 3: EasyMovePV(i) = PV(1, i): Next
     If bTimeTrace Then WriteTrace "UpdateEasyMove: " & MoveText(PV(1, 1)) & " " & MoveText(PV(1, 2)) & " " & MoveText(PV(1, 3))
   End If
End Sub

Public Function GetEasyMove() As TMove
  ' Return Easy move if previous moves are as expected
  GetEasyMove = EmptyMove
  If GameMovesCnt >= 2 And EasyMovePV(3).From > 0 Then
    If bTimeTrace Then WriteTrace "GetEasyMove: EM3" & MoveText(EasyMovePV(3)) & " ( EM1:" & MoveText(EasyMovePV(1)) & " = GM1:" & MoveText(arGameMoves(GameMovesCnt - 1)) & "  / EM2:" & MoveText(EasyMovePV(1)) & " = GM2:" & MoveText(arGameMoves(GameMovesCnt))
    If MovesEqual(EasyMovePV(1), arGameMoves(GameMovesCnt - 1)) And _
       MovesEqual(EasyMovePV(2), arGameMoves(GameMovesCnt)) Then
      GetEasyMove = EasyMovePV(3)
    End If
  End If
End Function

Public Sub InitAttackBitCnt()
  Dim i As Long, Cnt As Long
  For i = 1 To 1024
    Cnt = 0
    If i And PLAttackBit Then Cnt = Cnt + 1
    If i And PRAttackBit Then Cnt = Cnt + 1
    If i And NAttackBit Then Cnt = Cnt + 1
    If i And BAttackBit Then Cnt = Cnt + 1
    If i And RAttackBit Then Cnt = Cnt + 1
    If i And QAttackBit Then Cnt = Cnt + 1
    If i And KAttackBit Then Cnt = Cnt + 1
    If i And BXrayAttackBit Then Cnt = Cnt + 1
    If i And RXrayAttackBit Then Cnt = Cnt + 1
    If i And QXrayAttackBit Then Cnt = Cnt + 1
    AttackBitCnt(i) = Cnt
  Next
End Sub

Public Function StatBonus(ByVal Depth As Long) As Long
  StatBonus = Depth * Depth + 2 * Depth - 2
End Function





