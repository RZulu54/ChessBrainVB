Attribute VB_Name = "EvalBas"
'========================================
'=  EVAL : Evaluation of board position =
'========================================
Option Explicit
' Game phase
Const PHASE_MIDGAME               As Long = 128
Const PHASE_ENDGAME               As Long = 0
Public Const MAX_SEE_DIFF         As Long = 80  ' greater than value bishop minus value Knight
Public Const TEMPO_BONUS          As Long = 23   ' 20 bonus for side to move
Public Const SPACE_THRESHOLD      As Long = 12222 ' compute space eval for opening phase only

'-- Endgame eval scale factors
Const SCALE_FACTOR_DRAW = 0
Const SCALE_FACTOR_ONEPAWN = 48
Const SCALE_FACTOR_NORMAL = 64
Const SCALE_FACTOR_MAX = 128
Const SCALE_FACTOR_NONE = 255

'SF7: Penalties for enemy's safe checks
Const QueenCheck                  As Long = 780
Const RookCheck                   As Long = 880
Const BishopCheck                 As Long = 435
Const KnightCheck                 As Long = 790
'---
Public IsolatedPenalty(1)         As TScore
Public BackwardPenalty(1)         As TScore
Public DoubledPenalty             As TScore
Public ConnectedBonus(1, 1, 2, 8) As TScore
Public LeverBonus(8)              As TScore
Public ShelterWeakness(4, 8)      As Long
Public StormDanger(4, 4, 8)       As Long
Public ThreatenedByHangingPawn    As TScore
Public ThreatByRank               As TScore
Public WeakUnopposedPawn          As TScore
Public Hanging                    As TScore
Public Overload                   As TScore
Public SafeCheck                  As TScore
Public OtherCheck                 As TScore
Public PawnlessFlank              As TScore
Public ScorePawn                  As TScore
Public ScoreKnight                As TScore
Public ScoreBishop                As TScore
Public ScoreRook                  As TScore
Public ScoreQueen                 As TScore
Public PieceScore(17)             As Long
Public PieceAbsValue(17)          As Long
Public PieceTypeValue(6)          As Long
Public WMaterial                  As Long
Public WNonPawnMaterial           As Long
Public BMaterial                  As Long
Public BNonPawnMaterial           As Long
Public Material                   As Long
Public NonPawnMaterial            As Long
Public DrawContempt               As Long
Dim WAttack(MAX_BOARD)            As Integer   '- Fields around king: count attacks ' public+Erase is 2x faster than local in Eval function !
Dim BAttack(MAX_BOARD)            As Integer   '- Fields around king: count attacks
Dim WThreat                       As TScore, BThreat As TScore
Public PiecePosScaleFactor        As Long ' set in INI file
Public CompKingDefScaleFactor     As Long ' set in INI file
Public OppKingAttScaleFactor      As Long ' set in INI file
Public PawnStructScaleFactor      As Long ' set in INI file
Public PassedPawnsScaleFactor     As Long ' set in INI file
Public MobilityScaleFactor        As Long ' set in INI file
Public ThreatsScaleFactor         As Long ' set in INI file
Public WKingScaleFactor           As Long, BKingScaleFactor As Long
Public PawnsWMax(9)               As Long  '--- Pawn max rank (2-7) for file A-H
Public PawnsWMin(9)               As Long  '--- Pawn min rank (2-7) for file A-H
Public WPawns(9)                  As Long  '--- number of pawns for file A-H
Public PawnsBMax(9)               As Long
Public PawnsBMin(9)               As Long
Public BPawns(9)                  As Long
Public RootMove                   As TMOVE
Public LastNodesCnt               As Long
Public LastThreadCheckNodesCnt    As Long
Public StaticEvalArr(MAX_PV)      As Long ' Eval history
Public TestCnt(20)                As Long '--- Counter for special debug cases
Public MidGameLimit               As Long
Public EndgameLimit               As Long
'------------------------
'--- Piece square tables
'------------------------
Public PsqtWP(MAX_BOARD)          As TScore
Public PsqtBP(MAX_BOARD)          As TScore
Public PsqtWB(MAX_BOARD)          As TScore
Public PsqtBB(MAX_BOARD)          As TScore
Public PsqtWN(MAX_BOARD)          As TScore
Public PsqtBN(MAX_BOARD)          As TScore
Public PsqtWQ(MAX_BOARD)          As TScore
Public PsqtBQ(MAX_BOARD)          As TScore
Public PsqtWR(MAX_BOARD)          As TScore
Public PsqtBR(MAX_BOARD)          As TScore
Public PsqtWK(MAX_BOARD)          As TScore
Public PsqtBK(MAX_BOARD)          As TScore
Public PsqVal(1, 16, MAX_BOARD)   As Long ' piece square score for piece: (endgame,piece,square)
'--- Mobility values for pieces
Public MobilityN(9)               As TScore
Public MobilityB(15)              As TScore
Public MobilityR(15)              As TScore
Public MobilityQ(29)              As TScore
Public ZeroScore                  As TScore
Public ThreatBySafePawn(5)        As TScore
Public OutpostBonusKnight(1)      As TScore
Public OutpostBonusBishop(1)      As TScore
Public ReachableOutpostKnight(1)  As TScore
Public ReachableOutpostBishop(1)  As TScore
Public KingAttackWeights(6)       As Long
Public QueenMinorsImbalance(12)   As Long
Public WBestPawnVal               As Long, BBestPawnVal As Long, WBestPawn As Long, BBestPawn As Long
Public GamePhase                  As Long
Public WKingAttackersWeight       As Long, WKingAttackersCount As Long, BKingAttackersWeight As Long, BKingAttackersCount As Long
Public bEvalTrace                 As Boolean
Public bTimeTrace                 As Boolean
Public bHashTrace                 As Boolean
Public bWinboardTrace             As Boolean
Public bWbPvInUciFormat           As Boolean
Public bThreadTrace               As Boolean
Dim PassedPawns(16)               As Long ' List of passed pawns (Square)
Dim PassedPawnsCnt                As Long
Dim WPassedPawnAttack             As Long, BPassedPawnAttack As Long
Public PushClose(8)               As Long
Public PushAway(8)                As Long
Public PushToEdges(MAX_BOARD)     As Long
Public WOutpostSq(MAX_BOARD)      As Boolean
Public BOutpostSq(MAX_BOARD)      As Boolean
' endgame
Public KRPPKRP_SFactor(8)         As Long

'--- Threat list
Dim ThreatCnt                     As Long

Public Type TThreatList
  HangCol         As enumColor
  HangPieceType   As Long
  AttackerPieceType    As Long
  AttackerSquare As Long
  AttackedSquare  As Long
End Type

Dim ThreatList(32)                 As TThreatList
' Pawn Eval
Dim Passed                         As Boolean, Opposed As Boolean, Backward As Boolean
Dim Neighbours                     As Boolean, Doubled As Boolean, Lever As Long, Supported As Long, Phalanx As Long, LeverPush As Long
Public PassedPawnFileBonus(8)      As TScore
Public PassedPawnRankBonus(8)      As TScore
Public PassedDanger(8)             As Long
Private OwnAttCnt                  As Long, OppAttCnt As Long
' Threats
Public ThreatByMinor(6)            As TScore ' Attacker is defended minor (B/N)
Public ThreatByRook(6)             As TScore
Public ThreatByAttackOnQueen       As TScore
Public KingOnOneBonus              As TScore
Public KingOnManyBonus             As TScore
' King protection
Public KingProtector(5)            As TScore
' Material imbalance (SF6)
Public QuadraticOurs(5, 5)         As Long
Public QuadraticTheirs(5, 5)       As Long
Public PawnSet(8)                  As Long
Public ImbPieceCount(COL_WHITE, 5) As Long
Private bWIsland                   As Boolean, bBIsland As Boolean
Private PieceSqList(15, 10)        As Integer ' <Piece type> <list number> Square List of pieces for multiple runs thorugh piece list
Private PieceSqListCnt(15)         As Integer ' counter for  PieceLoc
' temp
Private bIniReadDone               As Boolean


'---------------------------------------------------------------------------
'InitEval(ThreatMove)  Set piece values and piece square tables
'---------------------------------------------------------------------------
Public Sub InitEval()
  Dim Score As Long, bSaveEvalTrace As Boolean
  ZeroScore.MG = 0: ZeroScore.EG = 0
  '--- Limit  high eval values ( VERY important for playing style!)
  If Not bIniReadDone Then
    bIniReadDone = True
    '--- Default used if INI file is missing
    PiecePosScaleFactor = Val(ReadINISetting("POSITION_FACTOR", "100"))
    MobilityScaleFactor = Val(ReadINISetting("MOBILITY_FACTOR", "100"))
    PawnStructScaleFactor = Val(ReadINISetting("PAWNSTRUCT_FACTOR", "100"))
    PassedPawnsScaleFactor = Val(ReadINISetting("PASSEDPAWNS_FACTOR", "130"))
    ThreatsScaleFactor = Val(ReadINISetting("THREATS_FACTOR", "150"))
    OppKingAttScaleFactor = Val(ReadINISetting("OPPKINGATT_FACTOR", "100"))
    CompKingDefScaleFactor = Val(ReadINISetting("COMPKINGDEF_FACTOR", "100"))
    '
    '--- Piece values  MG=midgame / EG=endgame58
    '--- SF6 values  ( scale to centipawns: \256 )
    '
    ScorePawn.MG = Val(ReadINISetting("PAWN_VAL_MG", "142"))
    ScorePawn.EG = Val(ReadINISetting("PAWN_VAL_EG", "207"))
    ScoreKnight.MG = Val(ReadINISetting("KNIGHT_VAL_MG", "784"))
    ScoreKnight.EG = Val(ReadINISetting("KNIGHT_VAL_EG", "868"))
    ScoreBishop.MG = Val(ReadINISetting("BISHOP_VAL_MG", "828"))
    ScoreBishop.EG = Val(ReadINISetting("BISHOP_VAL_EG", "916"))
    ScoreRook.MG = Val(ReadINISetting("ROOK_VAL_MG", "1286"))
    ScoreRook.EG = Val(ReadINISetting("ROOK_VAL_EG", "1378"))
    ScoreQueen.MG = Val(ReadINISetting("QUEEN_VAL_MG", "2528"))
    ScoreQueen.EG = Val(ReadINISetting("QUEEN_VAL_EG", "2698"))
    MidGameLimit = Val(ReadINISetting("MIDGAME_LIMIT", "15258")) ' for game phase
    EndgameLimit = Val(ReadINISetting("ENDGAME_LIMIT", "3915"))  ' for game phase
    ' Draw contempt in centipawns > scale to SF (needs ScorePawn.EG set)
    DrawContempt = Val(ReadINISetting(CONTEMPT_KEY, "1"))
    DrawContempt = Eval100ToSF(DrawContempt) ' in centipawns
  End If
  '--- Detect endgame stage ---
  bSaveEvalTrace = bEvalTrace: bEvalTrace = False ' Save trace setting, trace not needed here before init done
  Score = Eval() ' Set material,NonPawnMaterial for GamePhase calculation
  bEvalTrace = bSaveEvalTrace
  SetGamePhase NonPawnMaterial ' Set GamePhase, PieceValues, bEndGame
  InitPieceValue
  '--- Pawn values needed, so init here
  'InitRecaptureMargins ' no longer used
  InitFutilityMoveCounts
  InitReductionArray
  InitRazorMargin
  InitConnectedPawns
  InitOutpostSq
End Sub

Public Sub InitPieceValue()
  '--- Piece values, always absolut, positive value
  PieceAbsValue(FRAME) = 0
  PieceAbsValue(WPAWN) = ScorePawn.MG: PieceAbsValue(BPAWN) = ScorePawn.MG
  PieceAbsValue(WKNIGHT) = ScoreKnight.MG: PieceAbsValue(BKNIGHT) = ScoreKnight.MG
  PieceAbsValue(WBISHOP) = ScoreBishop.MG: PieceAbsValue(BBISHOP) = ScoreBishop.MG
  PieceAbsValue(WROOK) = ScoreRook.MG: PieceAbsValue(BROOK) = ScoreRook.MG
  PieceAbsValue(WQUEEN) = ScoreQueen.MG: PieceAbsValue(BQUEEN) = ScoreQueen.MG
  PieceAbsValue(WKING) = 5000: PieceAbsValue(BKING) = 5000
  PieceAbsValue(13) = 0: PieceAbsValue(14) = 0
  PieceAbsValue(WEP_PIECE) = ScorePawn.MG: PieceAbsValue(BEP_PIECE) = ScorePawn.MG
  '--- Piece SCore: positive for White, negative for Black
  PieceScore(FRAME) = 0
  PieceScore(WPAWN) = ScorePawn.MG: PieceScore(BPAWN) = -ScorePawn.MG
  PieceScore(WKNIGHT) = ScoreKnight.MG: PieceScore(BKNIGHT) = -ScoreKnight.MG
  PieceScore(WBISHOP) = ScoreBishop.MG: PieceScore(BBISHOP) = -ScoreBishop.MG
  PieceScore(WROOK) = ScoreRook.MG: PieceScore(BROOK) = -ScoreRook.MG
  PieceScore(WQUEEN) = ScoreQueen.MG: PieceScore(BQUEEN) = -ScoreQueen.MG
  PieceScore(WKING) = 5000: PieceScore(BKING) = -PieceScore(WKING)
  PieceScore(13) = 0: PieceScore(14) = 0
  PieceScore(WEP_PIECE) = ScorePawn.MG: PieceScore(BEP_PIECE) = -ScorePawn.MG
  PieceTypeValue(PT_PAWN) = ScorePawn.MG
  PieceTypeValue(PT_KNIGHT) = ScoreKnight.MG
  PieceTypeValue(PT_BISHOP) = ScoreBishop.MG
  PieceTypeValue(PT_ROOK) = ScoreRook.MG
  PieceTypeValue(PT_QUEEN) = ScoreQueen.MG
  PieceTypeValue(PT_KING) = PieceScore(WKING)
End Sub

Public Function SetGamePhase(ByVal NonPawnMaterial As Long) As Long
  Debug.Assert NonPawnMaterial >= 0
  NonPawnMaterial = GetMax(EndgameLimit, GetMin(NonPawnMaterial, MidGameLimit))
  GamePhase = (((NonPawnMaterial - EndgameLimit) * PHASE_MIDGAME) / (MidGameLimit - EndgameLimit))
  bEndgame = (GamePhase <= PHASE_ENDGAME)
End Function

'---------------------------------------------------------------------------------------------------
'---  Eval() - Evaluation of position
'---           Returns value from view of side to move (positive if black to move and black is better)
'---           Value scaled to stockfish pawn endgame value (258 = 1 pawn)
'---
'---  Steps:
'---         Init: inits attacks arrays, pawn arrays, material values for pieces
'---         Check material draw or special endgame positions
'---         STEPS:
'---         1. Loop over all pieces to fill pawn structure array and pawn threats
'---         2. Loop over all pieces types: evaluate each piece except kings.
'---            do a move generation to calculate mobility, attackers, defenders. fill attack array with piece bitcode
'---         3. Pass for pawn push (locate here because full attack info needed)
'---         4. Calculate king safety ( shelter, pawn storm, check attacks ), king distance to best pawn
'---         5. Calculate threats
'---         6. Calculate trapped bishops, passed pawns, center control, pawn islands
'---         7. Calculate total material values and endgame scale factors
'---         8. Calculate weights and total eval
'---             Add all evalution terms weighted by variables set in INI file:
'---             Material + Position(general) + PawnStructure + PassedPawns + Mobility +
'---             KingSafetyComputer + KingSafetyOpponent + Threats
'---         9. Invert score for black to move
'---        10. Add tempo value for side to move
'---------------------------------------------------------------------------------------------------
Public Function Eval() As Long
  Dim a                       As Long, i As Long, Square As Long, Target As Long, Offset As Long, MobCnt As Long, r As Long, rr As Long, AttackBit As Long, k As Long, ForkCnt As Long, SC As TScore
  Dim WPos                    As TScore, BPos As TScore, WPassed As TScore, BPassed As TScore, WMobility As TScore, BMobility As TScore
  Dim WPawnStruct             As TScore, BPawnStruct As TScore, Piece As Long, WPawnCnt As Long, BPawnCnt As Long
  Dim WKSafety                As TScore, BKSafety As TScore, bDoWKSafety As Boolean, bDoBKSafety As Boolean
  Dim WKingAdjacentZoneAttCnt As Long, BKingAdjacentZoneAttCnt As Long, WKingAttPieces As Long, BKingAttPieces As Long
  Dim KingDanger              As Long, Undefended As Long, RankNum As Long, RelRank As Long, QueenWeak As Boolean
  Dim FileNum                 As Long, MinWKingPawnDistance As Long, MinBKingPawnDistance As Long, KingSidePawns As Long, QueenSidePawns As Long
  Dim DefByPawn               As Long, AttByPawn As Long, bAllDefended As Boolean, BlockSqDefended As Boolean, WPinnedCnt As Long, BPinnedCnt As Long
  Dim RankPath                As Long, sq As Long ', WSemiOpenFiles As Long, BSemiOpenFiles As Long
  Dim BlockSq                 As Long, MBonus As Long, EBonus As Long, UnsafeCnt As Long, PieceAttackBit As Long
  Dim OwnCol                  As Long, OppCol As Long, MoveUp As Long, OwnKingLoc As Long, OppKingLoc As Long, BlockSqUnsafe As Boolean
  Dim WBishopsOnBlackSq       As Long, WBishopsOnWhiteSq As Long, BBishopsOnBlackSq As Long, BBishopsOnWhiteSq As Long, WCenterPawnsBlocked As Long, BCenterPawnsBlocked As Long
  Dim WPawnCntOnWhiteSq       As Long, BPawnCntOnWhiteSq As Long, WWeakUnopposedCnt As Long, BWeakUnopposedCnt As Long
  Dim WKingFile               As Long, BKingFile As Long, WFrontMostPassedPawnRank As Long, BFrontMostPassedPawnRank As Long, ScaleFactor As Long
  Dim WChecksCounted          As Long, BChecksCounted As Long, WUnsafeChecks As Long, BUnsafeChecks As Long, KingLevers As Long
  '
  '------ Init Eval
  '
  If bEvalTrace Then WriteTrace "------- Start Eval ------"
  EvalCnt = EvalCnt + 1
  Eval = 0
  WPawnCnt = PieceCnt(WPAWN): BPawnCnt = PieceCnt(BPAWN)
  WKingFile = File(WKingLoc): BKingFile = File(BKingLoc)
  WNonPawnMaterial = PieceCnt(WQUEEN) * ScoreQueen.MG + PieceCnt(WROOK) * ScoreRook.MG + PieceCnt(WBISHOP) * ScoreBishop.MG + PieceCnt(WKNIGHT) * ScoreKnight.MG
  WMaterial = WNonPawnMaterial + WPawnCnt * ScorePawn.MG
  BNonPawnMaterial = PieceCnt(BQUEEN) * ScoreQueen.MG + PieceCnt(BROOK) * ScoreRook.MG + PieceCnt(BBISHOP) * ScoreBishop.MG + PieceCnt(BKNIGHT) * ScoreKnight.MG
  BMaterial = BNonPawnMaterial + BPawnCnt * ScorePawn.MG
  NonPawnMaterial = WNonPawnMaterial + BNonPawnMaterial
  Material = WMaterial - BMaterial
  SetGamePhase NonPawnMaterial

  'Debug.Assert PieceSqListCnt(WPAWN) = PieceCnt(WPAWN)
  'Debug.Assert PieceSqListCnt(BPAWN) = PieceCnt(BPAWN)
  '
  '--- Endgame function available?
  '
  Select Case WPawnCnt + BPawnCnt
    Case 0 ' no pawns
      ' KQKR
      If (WMaterial = ScoreQueen.MG And BMaterial = ScoreRook.MG) Or (BMaterial = ScoreQueen.MG And WMaterial = ScoreRook.MG) Then
        Eval = Eval_KQKR(): GoTo lblEndEval
      End If
      '--- Insufficent material draw?
      If IsMaterialDraw() Then
        Eval = 0: Exit Function '- Endgame draw: not sufficent material for mate
      End If
    Case 1 ' one pawn
      If (WMaterial = ScoreRook.MG And BMaterial = ScorePawn.MG) Or (BMaterial = ScoreRook.MG And WMaterial = ScorePawn.MG) Then
        Eval = Eval_KRKP(): GoTo lblEndEval ' KRKP
      ElseIf (WMaterial = ScoreQueen.MG And BMaterial = ScorePawn.MG) Or (BMaterial = ScoreQueen.MG And WMaterial = ScorePawn.MG) Then
        Eval = Eval_KQKP(): GoTo lblEndEval ' KQKP
      End If
  End Select

  '----- Init Eval ---------------------
  WBestPawnVal = UNKNOWN_SCORE: WBestPawn = 0
  BBestPawnVal = UNKNOWN_SCORE: BBestPawn = 0
  WPassedPawnAttack = 0: BPassedPawnAttack = 0
  ThreatCnt = 0: WThreat = ZeroScore: BThreat = ZeroScore

  '--- Fill Pawn Arrays: number of pawns in file
  Erase WPawns: Erase BPawns: Erase PawnsWMax: Erase PawnsBMax
  For a = 0 To 9
    PawnsWMin(a) = 9: PawnsBMin(a) = 9
  Next

  WPawns(0) = -1: BPawns(0) = -1
  WPawns(9) = -1: BPawns(9) = -1
  PassedPawnsCnt = 0
  Erase WAttack(): Erase BAttack() 'Init attack arrays  (fast)
  Erase PieceSqListCnt()
  MinWKingPawnDistance = 9: MinBKingPawnDistance = 9

  '--- Step 1. loop over pieces: count pieces for material totals and game phase calculation. add piece square table score.
  '----                          calc pawn min/max rank positions per file; pawn attacks(for mobility used later)
  
  For a = 1 To NumPieces
    Square = Pieces(a): If Square = 0 Or Board(Square) >= NO_PIECE Then GoTo lblNextPieceCnt
    r = Board(Square):  PieceSqListCnt(r) = PieceSqListCnt(r) + 1: PieceSqList(r, PieceSqListCnt(r)) = Square ' fill piece list

    Select Case r
      Case WPAWN
        WAttack(Square + SQ_UP_LEFT) = WAttack(Square + SQ_UP_LEFT) Or PLAttackBit: WAttack(Square + SQ_UP_RIGHT) = WAttack(Square + SQ_UP_RIGHT) Or PRAttackBit  ' Set pawn attack here for use in pieces eval
        FileNum = File(Square): RankNum = Rank(Square): WPawns(FileNum) = WPawns(FileNum) + 1
        If RankNum < PawnsWMin(FileNum) Then PawnsWMin(FileNum) = RankNum
        If RankNum > PawnsWMax(FileNum) Then PawnsWMax(FileNum) = RankNum
        If MaxDistance(WKingLoc, Square) < MinWKingPawnDistance Then MinWKingPawnDistance = MaxDistance(WKingLoc, Square)
        If ColorSq(Square) = COL_WHITE Then WPawnCntOnWhiteSq = WPawnCntOnWhiteSq + 1  ' for Bishop eval
       ' If FileNum < FILE_E Then QueenSidePawns = QueenSidePawns + 1 Else KingSidePawns = KingSidePawns + 1
      Case BPAWN
        BAttack(Square + SQ_DOWN_LEFT) = BAttack(Square + SQ_DOWN_LEFT) Or PLAttackBit: BAttack(Square + SQ_DOWN_RIGHT) = BAttack(Square + SQ_DOWN_RIGHT) Or PRAttackBit
        FileNum = File(Square): RankNum = Rank(Square): BPawns(FileNum) = BPawns(FileNum) + 1
        If RankNum < PawnsBMin(FileNum) Then PawnsBMin(FileNum) = RankNum
        If RankNum > PawnsBMax(FileNum) Then PawnsBMax(FileNum) = RankNum
        If MaxDistance(BKingLoc, Square) < MinBKingPawnDistance Then MinBKingPawnDistance = MaxDistance(BKingLoc, Square)
        If ColorSq(Square) = COL_WHITE Then BPawnCntOnWhiteSq = BPawnCntOnWhiteSq + 1 ' for Bishop eval
       ' If FileNum < FILE_E Then QueenSidePawns = QueenSidePawns + 1 Else KingSidePawns = KingSidePawns + 1
    End Select

lblNextPieceCnt:
  Next

  '--- KPK endgame: Eval if promoted pawn cannot be captured
  If NonPawnMaterial = 0 And (WPawnCnt + BPawnCnt = 1) Then
    If WPawnCnt = 1 Then
      sq = PieceSqList(WPAWN, 1)
      If File(sq) = FILE_A Or File(sq) = FILE_H Then
        If File(BKingLoc) = File(sq) And Rank(BKingLoc) > Rank(sq) Then Eval = 0:  GoTo lblEndEval
      End If

      If bWhiteToMove Then
        If Rank(sq) = 7 Then
          If sq + SQ_UP <> WKingLoc Then ' own king not at promote square
            If MaxDistance(BKingLoc, sq + SQ_UP) > 1 Or MaxDistance(WKingLoc, sq + SQ_UP) = 1 Then
              Eval = VALUE_KNOWN_WIN: GoTo lblEndEval
            End If
          End If
        End If
        '--- Draw if opp king 2 rows in front of pawn (not at rank 8) and own king behind
        If Rank(BKingLoc) <> 8 Then
          If BKingLoc >= sq + SQ_UP + SQ_UP_LEFT And BKingLoc <= sq + SQ_UP + SQ_UP_RIGHT Then
            If WKingLoc >= sq + SQ_DOWN_LEFT And WKingLoc <= sq + SQ_DOWN_RIGHT Then Eval = 0:  GoTo lblEndEval
          End If
        End If
        '
      End If
    Else
      sq = PieceSqList(BPAWN, 1)
      If File(sq) = FILE_A Or File(sq) = FILE_H Then
        If File(WKingLoc) = File(sq) And Rank(WKingLoc) < Rank(sq) Then Eval = 0:  GoTo lblEndEval
      End If
      
      If Not bWhiteToMove Then
        If Rank(sq) = 2 Then
          If sq + SQ_DOWN <> BKingLoc Then ' own king not at promote square
            If MaxDistance(WKingLoc, sq + SQ_DOWN) > 1 Or MaxDistance(BKingLoc, sq + SQ_DOWN) = 1 Then
              Eval = -VALUE_KNOWN_WIN: GoTo lblEndEval
            End If
          End If
        End If
        '--- Draw if opp king in front of pawn (not at rank 1) and own king behind
        If Rank(WKingLoc) <> 1 Then
          If WKingLoc >= sq + SQ_DOWN + SQ_DOWN_LEFT And WKingLoc <= sq + SQ_DOWN + SQ_DOWN_RIGHT Then
            If BKingLoc >= sq + SQ_UP_LEFT And BKingLoc <= sq + SQ_UP_RIGHT Then Eval = 0: GoTo lblEndEval
          End If
        End If
      End If
    End If
  End If
  '
  '--- King safety needed?
  '
  bDoWKSafety = CBool(BNonPawnMaterial >= ScoreQueen.MG)
  bDoBKSafety = CBool(WNonPawnMaterial >= ScoreQueen.MG)
  WKingAttackersCount = 0: WKingAttackersWeight = 0: BKingAttackersCount = 0: BKingAttackersWeight = 0
  '--- King Position
  WKSafety = ZeroScore: BKSafety = ZeroScore
  If WNonPawnMaterial > 0 And BMaterial = 0 Then
    WPos.EG = WPos.EG + (7 - MaxDistance(BKingLoc, WKingLoc)) * 12 ' follow opp king to edge for mate (KRK, KQK)
    BPos.EG = BPos.EG + PsqtBK(BKingLoc).EG
  ElseIf BNonPawnMaterial > 0 And WMaterial = 0 Then
    BPos.EG = BPos.EG + (7 - MaxDistance(WKingLoc, BKingLoc)) * 12
    WPos.EG = WPos.EG + PsqtWK(WKingLoc).EG
  Else
    AddScore WPos, PsqtWK(WKingLoc)
    AddScore BPos, PsqtBK(BKingLoc)
  End If

  '--------------------------------------------------------------------
  '--- Step 2: EVAL Loop over pieces ------------------------------------------
  '--------------------------------------------------------------------
  '
  '--------------------------------------------------------------------
  '---- WHITE PAWNs ------------------------------------
  '--------------------------------------------------------------------
  For a = 1 To PieceSqListCnt(WPAWN)
    Square = PieceSqList(WPAWN, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = RankNum: SC.MG = 0: SC.EG = 0
    WPos.MG = WPos.MG + PsqtWP(Square).MG: WPos.EG = WPos.EG + PsqtWP(Square).EG
    DefByPawn = AttackBitCnt(WAttack(Square) And PAttackBit) ' counts 1 or 2 pawns
    AttByPawn = AttackBitCnt(BAttack(Square) And PAttackBit) ' counts 1 or 2 pawns

    If bEndgame And RankNum > 4 Then If MaxDistance(Square, BKingLoc) = 1 Then SC.EG = SC.EG + 10 ' advanced pawn supported by king
    'If BPawns(FileNum) = 0 Then WSemiOpenFiles = WSemiOpenFiles + 12 \ WPawns(FileNum) ' only count once per file, so 12 \ WPawns(FileNum) works for 1,2,3,4 pawns
    Opposed = (BPawns(FileNum) > 0) And RankNum < PawnsBMax(FileNum)
    Lever = AttByPawn
    Supported = DefByPawn
    LeverPush = AttackBitCnt(BAttack(Square + SQ_UP) And PAttackBit)
    Doubled = (Board(Square + SQ_DOWN) = WPAWN) ' not SQ_UP!
    Neighbours = (WPawns(FileNum + 1) > 0 Or WPawns(FileNum - 1) > 0)
    Phalanx = AttackBitCnt(WAttack(Square + SQ_UP) And PAttackBit)
    '
    If Not Neighbours Or Lever Or RelRank >= 5 Then
      Backward = False
    Else
      r = GetMin(PawnsWMin(FileNum - 1), PawnsWMin(FileNum + 1))
      If r <= RankNum Then
        Backward = False
      Else
        Backward = True
        If r = RankNum + 1 Then ' can safely advance to not backward rank?
          If LeverPush = 0 Then If Board(Square + SQ_UP) <> BPAWN Then Backward = False
        End If
      End If
    End If

    ' Blocked pawn on center files? Needed for bishop eval
    If FileNum >= FILE_C Then If FileNum <= FILE_F Then If Board(Square + SQ_UP) < NO_PIECE Then WCenterPawnsBlocked = WCenterPawnsBlocked + 1

    '
    '-----  Passed pawn?
    '
    Passed = False
    If Doubled Then GoTo lblEndWPassed
    
    ' Stopper two or more ranks in front?
    For k = -1 To 1
      If PawnsBMax(FileNum + k) > RankNum + 1 Then GoTo lblEndWPassed
    Next k
    If Board(Square + SQ_UP) = BPAWN Then
      ' phalanx neighbour can capture block opp pawn and became a passer
      If Phalanx > LeverPush Then If Supported >= Lever And RankNum >= 5 And bWhiteToMove Then Passed = True: GoTo lblEndWPassed
    Else
      If AttByPawn = 0 Then
        Passed = True:  GoTo lblEndWPassed
      ElseIf Phalanx >= LeverPush Then
        ' debug.print printpos, LocCoord(square)
        If Supported >= Lever Then Passed = True: GoTo lblEndWPassed
      End If
    End If
    '
    If Not Passed And Supported > 0 And RankNum >= 5 Then ' sacrify supporter pawn to create passer?
      If PawnsBMax(FileNum) = RankNum + 1 Then ' blocker pawn
        If PawnsBMax(FileNum - 1) < RankNum Then ' no other stopper left side
          If CBool(WAttack(Square) And PRAttackBit) Then ' left side supporter pawn (attacks to right)
            If Board(Square + SQ_LEFT) >= NO_PIECE Then  ' can move forward to attack stopper
              If Not CBool(BAttack(Square + SQ_LEFT) And PRAttackBit) Then ' no second left to right attacker from file-2
                Passed = True:  GoTo lblEndWPassed
              End If
            End If
          End If
        End If
        If Not Passed Then
          If PawnsBMax(FileNum + 1) < RankNum Then
            If CBool(WAttack(Square) And PLAttackBit) Then ' right side supporter pawn (attacks from left)
              If Board(Square + SQ_RIGHT) >= NO_PIECE Then  ' can move forward to attack stopper
                If Not CBool(BAttack(Square + SQ_RIGHT) And PLAttackBit) Then ' no second right to left attacker
                  Passed = True:  GoTo lblEndWPassed
                End If
              End If
            End If
          End If
        End If
      End If
    End If
lblEndWPassed:

    '--- pawn score
    If Lever Then AddScore SC, LeverBonus(RelRank)
    If Supported Or Phalanx Then ' Connected
      AddScore SC, ConnectedBonus(Abs(Opposed), Abs(Phalanx <> 0), DefByPawn, RelRank)
    ElseIf Not Neighbours Then
      MinusScore SC, IsolatedPenalty(Abs(Opposed))
      If Not Opposed Then WWeakUnopposedCnt = WWeakUnopposedCnt + 1
    ElseIf Backward Then
      MinusScore SC, BackwardPenalty(Abs(Opposed))
    End If
    If Doubled And Supported = 0 Then MinusScore SC, DoubledPenalty
    '---------------------
    If bEndgame Then
      If FileNum = 1 Or FileNum = 8 Then AddScore SC, PsqtWP(Square)
      If WPawnCnt = 1 Then SC.EG = SC.EG + 2 * RelRank * RelRank
      If SC.EG + PsqtWP(Square).EG > WBestPawnVal Then
        WBestPawnVal = SC.EG + PsqtWP(Square).EG: WBestPawn = Square
      ElseIf SC.EG = WBestPawnVal Then
        If WBestPawn = 0 Or MaxDistance(Square, WKingLoc) < MaxDistance(WBestPawn, WKingLoc) Then
          WBestPawnVal = SC.EG: WBestPawn = Square
        End If
      End If
    End If
    ' Passed : eval later when full attack is available
    If Passed Then
      PassedPawnsCnt = PassedPawnsCnt + 1: PassedPawns(PassedPawnsCnt) = Square
      If RankNum > 4 Then If Abs(FileNum - BKingFile) <= 2 Then WPassedPawnAttack = WPassedPawnAttack + 1
    End If
    '
    AddScore WPawnStruct, SC
    If bEvalTrace Then WriteTrace "WPawn: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG
  Next a

  '--------------------------------------------------------------------
  '---- BLACK PAWNs ------------------------------------
  '--------------------------------------------------------------------
  For a = 1 To PieceSqListCnt(BPAWN)
    Square = PieceSqList(BPAWN, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = (9 - RankNum): SC.MG = 0: SC.EG = 0
    'Debug.Assert Board(Square) = BPAWN
    BPos.MG = BPos.MG + PsqtBP(Square).MG: BPos.EG = BPos.EG + PsqtBP(Square).EG
    DefByPawn = AttackBitCnt(BAttack(Square) And PAttackBit) ' counts 1 or 2 pawns
    AttByPawn = AttackBitCnt(WAttack(Square) And PAttackBit) ' counts 1 or 2 pawns

    If bEndgame And RelRank > 4 Then If MaxDistance(Square, WKingLoc) = 1 Then SC.EG = SC.EG + 10  ' advanced pawn supported by king
    'If WPawns(FileNum) = 0 Then BSemiOpenFiles = BSemiOpenFiles + 12 \ BPawns(FileNum)
    Opposed = RankNum > PawnsWMin(FileNum)  ' PawnsWMin=9 if no pawn
    Lever = AttByPawn
    Supported = DefByPawn
    LeverPush = AttackBitCnt(WAttack(Square + SQ_DOWN) And PAttackBit)
    Doubled = Abs(Board(Square + SQ_UP) = BPAWN)
    Neighbours = (BPawns(FileNum + 1) > 0 Or BPawns(FileNum - 1) > 0)
    Phalanx = AttackBitCnt(BAttack(Square + SQ_DOWN) And PAttackBit)
    
    If Not Neighbours Or Lever Or RelRank >= 5 Then
      Backward = False
    Else
      r = GetMax(PawnsBMax(FileNum - 1), PawnsBMax(FileNum + 1))
      If r >= RankNum Then
        Backward = False
      Else
        Backward = True
        If r = RankNum - 1 Then ' can safely advance to not backward rank?
          If LeverPush = 0 Then If Board(Square + SQ_DOWN) <> WPAWN Then Backward = False
        End If
      End If
    End If
    
    ' Blocked pawn on center files? Needed for bishop eval
    If FileNum >= FILE_C Then If FileNum <= FILE_F Then If Board(Square + SQ_DOWN) < NO_PIECE Then BCenterPawnsBlocked = BCenterPawnsBlocked + 1
    
    '
    '-----  Passed pawn?
    '
    Passed = False
    If Doubled Then GoTo lblEndBPassed
    
    ' Stopper two or more ranks in front
    For k = -1 To 1
      If PawnsWMin(FileNum + k) < RankNum - 1 Then GoTo lblEndBPassed
    Next k
    If Board(Square - SQ_UP) = WPAWN Then
       If Phalanx > LeverPush Then If Supported >= Lever And RankNum <= 4 And Not bWhiteToMove Then Passed = True: GoTo lblEndBPassed
    Else
      If AttByPawn = 0 Then
        Passed = True: GoTo lblEndBPassed
      ElseIf Phalanx >= LeverPush Then
        If Supported >= Lever Then Passed = True: GoTo lblEndBPassed
      End If
    End If
    
    If Not Passed And Supported And RankNum <= 4 Then ' sacrify supporter pawn to create passer?
      If PawnsWMin(FileNum) = RankNum - 1 Then
        If PawnsWMin(FileNum - 1) > RankNum Then ' no other stopper left side (PawnsWMin=9 if no pawn)
          If CBool(BAttack(Square) And PRAttackBit) Then ' left side supporter pawn
            If Board(Square + SQ_LEFT) >= NO_PIECE Then  ' can move forward to attack stopper
              If Not CBool(WAttack(Square + SQ_LEFT) And PRAttackBit) Then ' no second left to right attacker from file-2
                Passed = True:  GoTo lblEndBPassed
              End If
            End If
          End If
        End If
        If Not Passed Then
          If PawnsWMin(FileNum + 1) > RankNum Then
            If CBool(BAttack(Square) And PLAttackBit) Then ' right side supporter pawn
              If Board(Square + SQ_RIGHT) >= NO_PIECE Then  ' can move forward to attack stopper
                If Not CBool(WAttack(Square + SQ_RIGHT) And PLAttackBit) Then ' no second right to left attacker
                  Passed = True: GoTo lblEndBPassed
                End If
              End If
            End If
          End If
        End If
      End If
    End If
lblEndBPassed:

    '--- pawn score
    If Lever Then AddScore SC, LeverBonus(RelRank)
    If Supported Or Phalanx Then ' Connected
      AddScore SC, ConnectedBonus(Abs(Opposed), Abs(Phalanx <> 0), DefByPawn, RelRank)
    ElseIf Not Neighbours Then
      MinusScore SC, IsolatedPenalty(Abs(Opposed))
      If Not Opposed Then BWeakUnopposedCnt = BWeakUnopposedCnt + 1
    ElseIf Backward Then
      MinusScore SC, BackwardPenalty(Abs(Opposed))
    End If
    If Doubled And Supported = 0 Then MinusScore SC, DoubledPenalty
    '-------------------------
    If bEndgame Then
      If FileNum = 1 Or FileNum = 8 Then AddScore SC, PsqtBP(Square)
      If BPawnCnt = 1 Then SC.EG = SC.EG + 2 * RelRank * RelRank
      If SC.EG + PsqtBP(Square).EG > BBestPawnVal Then
        BBestPawnVal = SC.EG + PsqtBP(Square).EG: BBestPawn = Square
      ElseIf SC.EG = BBestPawnVal Then
        If BBestPawn = 0 Or MaxDistance(Square, BKingLoc) < MaxDistance(BBestPawn, BKingLoc) Then
          BBestPawnVal = SC.EG: BBestPawn = Square
        End If
      End If
    End If
    ' Passed : eval later when full attack is available
    If Passed And Not Doubled Then
      PassedPawnsCnt = PassedPawnsCnt + 1: PassedPawns(PassedPawnsCnt) = Square
      If RelRank > 4 Then If Abs(FileNum - WKingFile) <= 2 Then BPassedPawnAttack = BPassedPawnAttack + 1
    End If
    '
    AddScore BPawnStruct, SC
    If bEvalTrace Then WriteTrace "BPawn: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG
  Next a

  '--------------------------------------------------------------------
  '---- WHITE KNIGHTs ------------------------------------
  '--------------------------------------------------------------------  '
  For a = 1 To PieceSqListCnt(WKNIGHT)
    Square = PieceSqList(WKNIGHT, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = RankNum: SC.MG = 0: SC.EG = 0
    WPos.MG = WPos.MG + PsqtWN(Square).MG: WPos.EG = WPos.EG + PsqtWN(Square).EG: r = 0
    ' Outpost bonus
    If WOutpostSq(Square) Then
      If Not CBool(BAttack(Square) And PAttackBit) Then ' not attacked by pawn
        ' Defended by pawn?
        AddScore SC, OutpostBonusKnight(Abs(CBool(WAttack(Square) And PAttackBit))): r = 3 ' ignore ReachableOutpost
        If bEvalTrace Then WriteTrace "WKight: " & LocCoord(Square) & "> Outpost:" & OutpostBonusKnight(Abs(CBool(WAttack(Square) And PAttackBit))).MG
      End If
    End If
    '--- Mobility
    If Moved(Square) = 0 Then SC.MG = SC.MG - 45
    ForkCnt = 0: MobCnt = 0
    If a = 1 Then PieceAttackBit = N1AttackBit Else PieceAttackBit = N2AttackBit
    
    For i = 0 To 7
      Offset = KnightOffsets(i): Target = Square + Offset
      If Board(Target) <> FRAME Then
        WAttack(Target) = WAttack(Target) Or PieceAttackBit

        Select Case Board(Target)
          Case NO_PIECE:
            If (Not CBool(BAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1: SC.MG = SC.MG + 3
          Case WPAWN: SC.MG = SC.MG + 2: SC.EG = SC.EG + 2
            If RankNum > 3 Then If Board(Target + SQ_UP) >= NO_PIECE Then If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
          Case BPAWN: SC.MG = SC.MG + 7: SC.EG = SC.EG + 7: If Rank(Target) >= 6 Then SC.MG = SC.MG + 4
            If (Not CBool(BAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1:  AddThreat COL_BLACK, PT_PAWN, PT_KNIGHT, Square, Target
          Case BKNIGHT, BBISHOP: If (Not CBool(BAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
            AddThreat COL_BLACK, PieceType(Board(Target)), PT_KNIGHT, Square, Target '-- no Score for WKnight : total is zero
          Case BROOK, BQUEEN: If (Not CBool(BAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
            AddThreat COL_BLACK, PieceType(Board(Target)), PT_KNIGHT, Square, Target: ForkCnt = ForkCnt + 1
          Case WKING, WQUEEN: ' ignore
          Case BKING: MobCnt = MobCnt + 1: ForkCnt = ForkCnt + 1
          Case WEP_PIECE, BEP_PIECE:
            If (Not CBool(BAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1: SC.MG = SC.MG + 3
          Case Else: If (Not CBool(BAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
        End Select

        If r < 2 Then ' choose best square only
          If WOutpostSq(Target) Then ' Empty or opp piece: square can be occupied.
            ' not attacked by opp pawn? Else if not blocked by own piece
            If Not CBool(BAttack(Target) And PAttackBit) Then
              r = 2: rr = 1 + Abs(CBool(WAttack(Target) And PAttackBit)) ' supported by own pawn? Factor 2
            Else
              If r = 0 Then If PieceColor(Board(Target)) <> COL_WHITE Then r = 1: rr = 1 + Abs(CBool(WAttack(Target) And PAttackBit)) ' supported by own pawn? Factor 2
            End If
          End If
        End If
      End If
    Next

    If ForkCnt > 1 Then AddScoreVal SC, 7 * ForkCnt * ForkCnt, 5 * ForkCnt * ForkCnt: If bWhiteToMove Then AddScoreVal SC, 35, 35
    AddScore WMobility, MobilityN(MobCnt)
    ' Minor behind pawn bonus
    If RelRank < 5 Then
      If PieceType(Board(Square + SQ_UP)) = PT_PAWN Then SC.MG = SC.MG + 16: If bEvalTrace Then WriteTrace "WKnight: " & LocCoord(Square) & "> Behind pawn 16"
    End If
    If r > 0 And r < 3 Then AddScoreWithFactor SC, ReachableOutpostKnight(r - 1), rr
    If CBool(BAttack(Square) And PAttackBit) Then AddPawnThreat BThreat, COL_WHITE, PieceType(Board(Square)), Square
    AddScoreWithFactor SC, KingProtector(PT_KNIGHT), MaxDistance(Square, WKingLoc) ' defends king?
    AddScore WPos, SC
    If bEvalTrace Then WriteTrace "WKnight: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & WPos.MG & ", " & WPos.EG
  Next a

  '--------------------------------------------------------------------
  '---- BLACK KNIGHTs ------------------------------------
  '--------------------------------------------------------------------  '
  For a = 1 To PieceSqListCnt(BKNIGHT)
    Square = PieceSqList(BKNIGHT, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = (9 - RankNum): SC.MG = 0: SC.EG = 0
    BPos.MG = BPos.MG + PsqtBN(Square).MG: BPos.EG = BPos.EG + PsqtBN(Square).EG: r = 0
    ' Outpost bonus
    If BOutpostSq(Square) Then
      If Not CBool(WAttack(Square) And PAttackBit) Then ' not attacked by pawn
        ' Defended by pawn?
        AddScore SC, OutpostBonusKnight(Abs(CBool(BAttack(Square) And PAttackBit))): r = 3 ' ignore ReachableOutpost
        If bEvalTrace Then WriteTrace "BKight: " & LocCoord(Square) & "> Outpost:" & OutpostBonusKnight(Abs(CBool(BAttack(Square) And PAttackBit))).MG
      End If
    End If
    If Moved(Square) = 0 Then SC.MG = SC.MG - 45
    '--- Mobility
    ForkCnt = 0: MobCnt = 0
    If a = 1 Then PieceAttackBit = N1AttackBit Else PieceAttackBit = N2AttackBit
    
    For i = 0 To 7
      Offset = KnightOffsets(i)
      Target = Square + Offset
      If Board(Target) <> FRAME Then
        BAttack(Target) = BAttack(Target) Or PieceAttackBit

        Select Case Board(Target)
          Case NO_PIECE:
            If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1: SC.MG = SC.MG + 3
          Case BPAWN: SC.MG = SC.MG + 2: SC.EG = SC.EG + 2
            If RankNum < 6 Then If Board(Target + SQ_DOWN) >= NO_PIECE Then If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
          Case WPAWN: SC.MG = SC.MG + 7: SC.EG = SC.EG + 7: If Rank(Target) <= 3 Then SC.MG = SC.MG + 4
            If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
            AddThreat COL_WHITE, PT_PAWN, PT_KNIGHT, Square, Target
          Case WKNIGHT, WBISHOP: If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
            AddThreat COL_WHITE, PieceType(Board(Target)), PT_KNIGHT, Square, Target  '-- no Score for WKnight : total is zero
          Case WROOK, WQUEEN: If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
            AddThreat COL_WHITE, PieceType(Board(Target)), PT_KNIGHT, Square, Target: ForkCnt = ForkCnt + 1
          Case BKING, BQUEEN: ' Ignore
          Case WKING: If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
            If (Not CBool(WAttack(Target) And PAttackBit)) Then ForkCnt = ForkCnt + 1
          Case WEP_PIECE, BEP_PIECE:
            If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1: SC.MG = SC.MG + 3
          Case Else: If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
        End Select

        If r < 2 Then
          If BOutpostSq(Target) Then ' Empty or opp piece: square can be occupied
            ' not attacked by opp pawn? Else if not blocked by own piece
            If Not CBool(WAttack(Target) And PAttackBit) Then
              r = 2: rr = 1 + Abs(CBool(BAttack(Target) And PAttackBit)) ' supported by own pawn? Factor 2
            Else
              If r = 0 Then If PieceColor(Board(Target)) <> COL_BLACK Then r = 1: rr = 1 + Abs(CBool(BAttack(Target) And PAttackBit)) ' supported by own pawn? Factor 2
            End If
          End If
        End If
      End If
    Next

    If ForkCnt > 1 Then AddScoreVal SC, 7 * ForkCnt * ForkCnt, 5 * ForkCnt * ForkCnt: If Not bWhiteToMove Then AddScoreVal SC, 35, 35
    AddScore BMobility, MobilityN(MobCnt)
    ' Minor behind pawn bonus
    If RelRank < 5 Then
      If PieceType(Board(Square + SQ_DOWN)) = PT_PAWN Then SC.MG = SC.MG + 16: If bEvalTrace Then WriteTrace "BKnight: " & LocCoord(Square) & "> Behind pawn 16"
    End If
    If r > 0 And r < 3 Then AddScoreWithFactor SC, ReachableOutpostKnight(r - 1), rr
    If CBool(WAttack(Square) And PAttackBit) Then AddPawnThreat WThreat, COL_BLACK, PieceType(Board(Square)), Square
    AddScoreWithFactor SC, KingProtector(PT_KNIGHT), MaxDistance(Square, BKingLoc)  ' defends king?
    AddScore BPos, SC
    If bEvalTrace Then WriteTrace "BKnight: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & BPos.MG & ", " & BPos.EG
  Next a

  '--------------------------------------------------------------------
  '---- WHITE BISHOPs ------------------------------------
  '--------------------------------------------------------------------  '
  For a = 1 To PieceSqListCnt(WBISHOP)
    Square = PieceSqList(WBISHOP, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = RankNum: SC.MG = 0: SC.EG = 0
    If ColorSq(Square) = COL_WHITE Then WBishopsOnWhiteSq = WBishopsOnWhiteSq + 1 Else WBishopsOnBlackSq = WBishopsOnBlackSq + 1
    WPos.MG = WPos.MG + PsqtWB(Square).MG: WPos.EG = WPos.EG + PsqtWB(Square).EG: r = 0
    ' Outpost bonus
    If WOutpostSq(Square) Then
      If Not CBool(BAttack(Square) And PAttackBit) Then ' not attacked by pawn
        ' Defended by pawn?
        AddScore SC, OutpostBonusBishop(Abs(CBool(WAttack(Square) And PAttackBit))): r = 3 ' ignore ReachableOutpost
        If bEvalTrace Then WriteTrace "WBishop: " & LocCoord(Square) & "> Outpost:" & OutpostBonusBishop(Abs(CBool(WAttack(Square) And PAttackBit))).MG
      End If
    End If
    '--- Mobility
    MobCnt = 0
    If a = 1 Then PieceAttackBit = B1AttackBit Else PieceAttackBit = B2AttackBit

    For i = 4 To 7
      Offset = QueenOffsets(i): Target = Square + Offset: AttackBit = PieceAttackBit

      Do While Board(Target) <> FRAME
        WAttack(Target) = WAttack(Target) Or AttackBit

        Select Case Board(Target)
          Case NO_PIECE:
            If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: If Offset > 0 Then SC.MG = SC.MG + 2
          Case WPAWN: SC.MG = SC.MG + 2: SC.EG = SC.EG + 3:
            If RankNum > 3 Then If Board(Target + SQ_UP) >= NO_PIECE Then If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            If Offset > 0 Then WAttack(Target + Offset) = WAttack(Target + Offset) Or BXrayAttackBit
            Exit Do
          Case BPAWN: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = PieceAttackBit Then AddThreat COL_BLACK, PT_PAWN, PT_BISHOP, Square, Target: SC.MG = SC.MG + 7: SC.EG = SC.EG + 7
            Exit Do
          Case BKNIGHT, BBISHOP, BROOK, BQUEEN: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = PieceAttackBit Then AddThreat COL_BLACK, PieceType(Board(Target)), PT_BISHOP, Square, Target ' Reattack: no SC because x-x=0
            Exit Do
          Case WKING: Exit Do ' ignore
          Case BKING: MobCnt = MobCnt + 1
            Exit Do
          Case WQUEEN: AttackBit = BXrayAttackBit  '--- Continue xray
          Case WEP_PIECE, BEP_PIECE:
            If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: If Offset > 0 Then SC.MG = SC.MG + 2
          Case Else: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            Exit Do ' own bishop or knight
        End Select

        If r < 2 Then
          If WOutpostSq(Target) Then ' Empty or opp piece: square can be occupied
            ' not attacked by opp pawn? Else if not blocked by own piece
            If Not CBool(BAttack(Target) And PAttackBit) Then
              r = 2: rr = 1 + Abs(CBool(WAttack(Target) And PAttackBit)) ' supported by own pawn? Factor 2
            Else
              If r = 0 Then If PieceColor(Board(Target)) <> COL_WHITE Then r = 1: rr = 1 + Abs(CBool(WAttack(Target) And PAttackBit)) ' supported by own pawn? Factor 2
            End If
          End If
        End If
        Target = Target + Offset
      Loop

    Next

    AddScore WMobility, MobilityB(MobCnt)
    If bEvalTrace Then WriteTrace "WBishop: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & WPos.MG & ", " & WPos.EG
    ' Minor behind pawn bonus
    If RelRank < 5 Then
      If PieceType(Board(Square + SQ_UP)) = PT_PAWN Then SC.MG = SC.MG + 16: If bEvalTrace Then WriteTrace "WBishop: " & LocCoord(Square) & "> Behind pawn 16"
    End If
    If r > 0 And r < 3 Then AddScoreWithFactor SC, ReachableOutpostBishop(r - 1), rr
    If CBool(BAttack(Square) And PAttackBit) Then AddPawnThreat BThreat, COL_WHITE, PieceType(Board(Square)), Square
    AddScoreWithFactor SC, KingProtector(PT_BISHOP), MaxDistance(Square, WKingLoc) ' defends king?
    AddScore WPos, SC
  Next a

  '--------------------------------------------------------------------
  '---- BLACK BISHOPs ------------------------------------
  '--------------------------------------------------------------------  '
  For a = 1 To PieceSqListCnt(BBISHOP)
    Square = PieceSqList(BBISHOP, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = (9 - RankNum): SC.MG = 0: SC.EG = 0
    If ColorSq(Square) = COL_WHITE Then BBishopsOnWhiteSq = BBishopsOnWhiteSq + 1 Else BBishopsOnBlackSq = BBishopsOnBlackSq + 1
    BPos.MG = BPos.MG + PsqtBB(Square).MG: BPos.EG = BPos.EG + PsqtBB(Square).EG: r = 0
    ' Outpost bonus
    If BOutpostSq(Square) Then
      If Not CBool(WAttack(Square) And PAttackBit) Then ' not attacked by pawn
        ' Defended by pawn?
        AddScore SC, OutpostBonusBishop(Abs(CBool(BAttack(Square) And PAttackBit))): r = 3 ' ignore ReachableOutpost
        If bEvalTrace Then WriteTrace "BBishop: " & LocCoord(Square) & "> Outpost:" & OutpostBonusBishop(Abs(CBool(BAttack(Square) And PAttackBit))).MG
      End If
    End If
    '--- Mobility
    MobCnt = 0
    If a = 1 Then PieceAttackBit = B1AttackBit Else PieceAttackBit = B2AttackBit

    For i = 4 To 7
      Offset = QueenOffsets(i): Target = Square + Offset:  AttackBit = PieceAttackBit

      Do While Board(Target) <> FRAME
        BAttack(Target) = BAttack(Target) Or AttackBit

        Select Case Board(Target)
          Case NO_PIECE:
            If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1:  If Offset < 0 Then SC.MG = SC.MG + 2
          Case BPAWN: SC.MG = SC.MG + 2: SC.EG = SC.EG + 3
            If RankNum < 6 Then If Board(Target + SQ_DOWN) >= NO_PIECE Then If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            If Offset < 0 Then BAttack(Target + Offset) = BAttack(Target + Offset) Or BXrayAttackBit
            Exit Do
          Case WPAWN: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = PieceAttackBit Then AddThreat COL_WHITE, PT_PAWN, PT_BISHOP, Square, Target: SC.MG = SC.MG + 7: SC.EG = SC.EG + 7
            Exit Do
          Case WKNIGHT, WBISHOP, WROOK, WQUEEN: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = PieceAttackBit Then AddThreat COL_WHITE, PieceType(Board(Target)), PT_BISHOP, Square, Target
            Exit Do
          Case BKING: Exit Do ' Ignore
          Case WKING: MobCnt = MobCnt + 1
            Exit Do
          Case BQUEEN: AttackBit = BXrayAttackBit '--- Continue xray
          Case WEP_PIECE, BEP_PIECE:
            If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1:  If Offset < 0 Then SC.MG = SC.MG + 2
          Case Else: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            Exit Do ' own bishop or knight
        End Select

        If r < 2 Then
          If BOutpostSq(Target) Then ' Empty or opp piece: square can be occupied
            ' not attacked by opp pawn? Else if not blocked by own piece
            If Not CBool(WAttack(Target) And PAttackBit) Then
              r = 2: rr = 1 + Abs(CBool(BAttack(Target) And PAttackBit)) ' supported by own pawn? Factor 2
            Else
              If r = 0 Then If PieceColor(Board(Target)) <> COL_BLACK Then r = 1: rr = 1 + Abs(CBool(BAttack(Target) And PAttackBit)) ' supported by own pawn? Factor 2
            End If
          End If
        End If
        Target = Target + Offset
      Loop

    Next

    AddScore BMobility, MobilityB(MobCnt)
    If bEvalTrace Then WriteTrace "BBishop: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & BPos.MG & ", " & BPos.EG
    ' Minor behind pawn bonus
    If RelRank < 5 Then
      If PieceType(Board(Square + SQ_DOWN)) = PT_PAWN Then SC.MG = SC.MG + 16: If bEvalTrace Then WriteTrace "BBishop: " & LocCoord(Square) & "> Behind pawn 16"
    End If
    If r > 0 And r < 3 Then AddScoreWithFactor SC, ReachableOutpostBishop(r - 1), rr
    If CBool(WAttack(Square) And PAttackBit) Then AddPawnThreat WThreat, COL_BLACK, PieceType(Board(Square)), Square
    AddScoreWithFactor SC, KingProtector(PT_BISHOP), MaxDistance(Square, BKingLoc) ' defends king?
    AddScore BPos, SC
  Next a

  '--------------------------------------------------------------------
  '---- WHITE ROOKs ------------------------------------
  '--------------------------------------------------------------------  '
  For a = 1 To PieceSqListCnt(WROOK)
    Square = PieceSqList(WROOK, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = RankNum: SC.MG = 0: SC.EG = 0
    WPos.MG = WPos.MG + PsqtWR(Square).MG: WPos.EG = WPos.EG + PsqtWR(Square).EG
    If WPawns(FileNum) = 0 Then
      If BPawns(FileNum) = 0 Then
        SC.MG = SC.MG + 45: SC.EG = SC.EG + 20
      Else
        SC.MG = SC.MG + 20: SC.EG = SC.EG + 7
      End If
    End If
    '--- Mobility
    MobCnt = 0
    If a = 1 Then PieceAttackBit = R1AttackBit Else PieceAttackBit = R2AttackBit

    For i = 0 To 3
      Offset = QueenOffsets(i): Target = Square + Offset: AttackBit = PieceAttackBit

      Do While Board(Target) <> FRAME
        WAttack(Target) = WAttack(Target) Or AttackBit

        Select Case Board(Target)
          Case NO_PIECE:
            If Not CBool(BAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1: If Abs(Offset) = 10 Then SC.MG = SC.MG + 7
          Case WPAWN: SC.MG = SC.MG + 2: SC.EG = SC.EG + 5:
            If RankNum > 3 Then If Board(Target + SQ_UP) >= NO_PIECE Then If Not CBool(BAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1
            Exit Do
          Case BPAWN:
            SC.MG = SC.MG + 7: SC.EG = SC.EG + 10 '--- no reattack possible
            If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = PieceAttackBit Then AddThreat COL_BLACK, PT_PAWN, PT_ROOK, Square, Target
            If RankNum >= 5 Then SC.MG = SC.MG + 8: SC.EG = SC.EG + 25  ' aligned pawns
            Exit Do
          Case BKNIGHT, BBISHOP:
            If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = PieceAttackBit Then AddThreat COL_BLACK, PieceType(Board(Target)), PT_ROOK, Square, Target  '--- no reattack possible
            Exit Do
          Case BROOK:  If AttackBit = PieceAttackBit Then AddThreat COL_BLACK, PT_ROOK, PT_ROOK, Square, Target
            MobCnt = MobCnt + 1
            Exit Do ' equal exchange, ok for mobility
          Case WKING: Exit Do ' ignore
          Case BKING: MobCnt = MobCnt + 1
            Exit Do
          Case BQUEEN: MobCnt = MobCnt + 1:  If AttackBit = PieceAttackBit Then AddThreat COL_BLACK, PT_QUEEN, PT_ROOK, Square, Target
            Exit Do
          Case WROOK, WQUEEN:
            If Offset = 10 Then
              If WPawns(FileNum) = 0 Then SC.MG = SC.MG + 12: If BPawns(FileNum) = 0 Then SC.MG = SC.MG + 15
            End If
            If Board(Target) = WROOK Then If Not CBool(BAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1
            If a = 1 Then AttackBit = R1XrayAttackBit Else AttackBit = R2XrayAttackBit '--- double lines , continue xray
          Case WEP_PIECE, BEP_PIECE:
            If Not CBool(BAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1: If Abs(Offset) = 10 Then SC.MG = SC.MG + 7
          Case Else: If Not CBool(BAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1
            Exit Do ' own bishop or knight
        End Select

        Target = Target + Offset
      Loop

    Next

    AddScore WMobility, MobilityR(MobCnt)
    ' Trapped rook by king : worse when cannot castle
    If Not bEndgame Then
      If MobCnt <= 3 Then
        If WPawns(FileNum) > 0 Then
          If RankNum = Rank(WKingLoc) Or Rank(WKingLoc) = 1 Then
            r = 0
            If WKingFile < FILE_E Then
              If FileNum < WKingFile Then r = -1
            Else
              If FileNum > WKingFile Then r = 1
            End If
            If r <> 0 Then
  
              For k = WKingFile + r To FileNum - r Step r ' own blocking pawns on files between king an rook
                If WPawns(k) = 0 Then
                  r = 0: Exit For
                ElseIf PawnsWMin(k) > RankNum + 2 Then
                  r = 0: Exit For
                End If
              Next
  
              If r <> 0 Then SC.MG = SC.MG - (92 - MobCnt * 22) * (1 + Abs(Rank(WKingLoc) = 1 And (Moved(WKING_START) > 0 Or Moved(Square) > 0)))
            End If
          End If
        End If
      End If
    Else
      If WPawns(FileNum) > 0 And BPawns(FileNum) = 0 And PawnsWMin(FileNum) >= 5 Then
        SC.MG = SC.MG + (PawnsWMin(FileNum)): SC.EG = SC.EG + 5 * PawnsWMin(FileNum)
      End If
    End If
    If CBool(BAttack(Square) And PAttackBit) Then AddPawnThreat BThreat, COL_WHITE, PieceType(Board(Square)), Square
    AddScoreWithFactor SC, KingProtector(PT_ROOK), MaxDistance(Square, WKingLoc) ' defends king?
    AddScore WPos, SC
    If bEvalTrace Then WriteTrace "WRook: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & WPos.MG & ", " & WPos.EG
  Next a

  '--------------------------------------------------------------------
  '---- BLACK ROOKs ------------------------------------
  '--------------------------------------------------------------------  '
  For a = 1 To PieceSqListCnt(BROOK)
    Square = PieceSqList(BROOK, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = (9 - RankNum): SC.MG = 0: SC.EG = 0
    BPos.MG = BPos.MG + PsqtBR(Square).MG: BPos.EG = BPos.EG + PsqtBR(Square).EG
    If BPawns(FileNum) = 0 Then
      If WPawns(FileNum) = 0 Then
        SC.MG = SC.MG + 45: SC.EG = SC.EG + 20
      Else
        SC.MG = SC.MG + 20: SC.EG = SC.EG + 7
      End If
    End If
    '--- Mobility
    MobCnt = 0
    If a = 1 Then PieceAttackBit = R1AttackBit Else PieceAttackBit = R2AttackBit

    For i = 0 To 3
      Offset = QueenOffsets(i): Target = Square + Offset: AttackBit = PieceAttackBit

      Do While Board(Target) <> FRAME
        BAttack(Target) = BAttack(Target) Or AttackBit
        
        Select Case Board(Target)
          Case NO_PIECE:
            If Not CBool(WAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1: If Abs(Offset) = 10 Then SC.MG = SC.MG + 7
          Case BPAWN: SC.MG = SC.MG + 2: SC.EG = SC.EG + 5
            If RankNum < 6 Then If Board(Target + SQ_DOWN) >= NO_PIECE Then If Not CBool(WAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1
            Exit Do
          Case WPAWN:
            SC.MG = SC.MG + 7: SC.EG = SC.EG + 10  '--- no reattack possible
            If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = PieceAttackBit Then AddThreat COL_WHITE, PT_PAWN, PT_ROOK, Square, Target
            If RankNum <= 4 Then SC.MG = SC.MG + 8: SC.EG = SC.EG + 25  ' aligned pawns
            Exit Do
          Case WKNIGHT, WBISHOP:
            If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = PieceAttackBit Then AddThreat COL_WHITE, PieceType(Board(Target)), PT_ROOK, Square, Target
            Exit Do   '--- no reattack possible
          Case WROOK:  If AttackBit = PieceAttackBit Then AddThreat COL_WHITE, PT_ROOK, PT_ROOK, Square, Target
            MobCnt = MobCnt + 1
            Exit Do  ' equal exchange ok for mobility
          Case BKING: Exit Do ' Ignore
          Case WKING: MobCnt = MobCnt + 1
            Exit Do
          Case WQUEEN: MobCnt = MobCnt + 1:  If AttackBit = PieceAttackBit Then AddThreat COL_WHITE, PT_QUEEN, PT_ROOK, Square, Target
            Exit Do
          Case BROOK, BQUEEN:
            If Offset = -10 Then
              If BPawns(FileNum) = 0 Then SC.MG = SC.MG + 12: If WPawns(FileNum) = 0 Then SC.MG = SC.MG + 15
            End If
            If Board(Target) = BROOK Then If Not CBool(WAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1
            If a = 1 Then AttackBit = R1XrayAttackBit Else AttackBit = R2XrayAttackBit '--- double lines , continue xray
          Case WEP_PIECE, BEP_PIECE:
            If Not CBool(WAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1: If Abs(Offset) = 10 Then SC.MG = SC.MG + 7
          Case Else: If Not CBool(WAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1
            Exit Do ' own bishop or knight
        End Select

        Target = Target + Offset
      Loop

    Next

    AddScore BMobility, MobilityR(MobCnt)
    ' Trapped rook by king : worse when cannot castle
    If Not bEndgame Then
      If MobCnt <= 3 Then
        If BPawns(FileNum) > 0 Then
          If RankNum = Rank(BKingLoc) Or Rank(BKingLoc) = 1 Then
            r = 0
            If BKingFile < FILE_E Then
              If FileNum < BKingFile Then r = -1
            Else
              If FileNum > BKingFile Then r = 1
            End If
            If r <> 0 Then
  
              For k = BKingFile + r To FileNum - r Step r ' own blocking pawns on files between king an rook
                If BPawns(k) = 0 Then
                  r = 0: Exit For
                ElseIf PawnsBMax(k) < RankNum - 2 Then
                  r = 0: Exit For
                End If
              Next
  
              If r <> 0 Then SC.MG = SC.MG - (92 - MobCnt * 22) * (1 + Abs(Rank(BKingLoc) = 8 And (Moved(BKING_START) > 0 Or Moved(Square) > 0)))
            End If
          End If
        End If
      End If
    Else
      If BPawns(FileNum) > 0 And WPawns(FileNum) = 0 And PawnsBMax(FileNum) <= 4 Then
        SC.MG = SC.MG + (9 - PawnsBMin(FileNum)): SC.EG = SC.EG + 5 * (9 - PawnsBMin(FileNum))
      End If
    End If
    If CBool(WAttack(Square) And PAttackBit) Then AddPawnThreat WThreat, COL_BLACK, PieceType(Board(Square)), Square
    AddScoreWithFactor SC, KingProtector(PT_ROOK), MaxDistance(Square, BKingLoc) ' defends king?
    AddScore BPos, SC
    If bEvalTrace Then WriteTrace "BRook: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & BPos.MG & ", " & BPos.EG
  Next a

  '--------------------------------------------------------------------
  '---- WHITE QUEENs ( last - full attack info needed for mobility )  -
  '--------------------------------------------------------------------
  For a = 1 To PieceSqListCnt(WQUEEN)
    Square = PieceSqList(WQUEEN, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = RankNum: SC.MG = 0: SC.EG = 0: QueenWeak = False
    WPos.MG = WPos.MG + PsqtWQ(Square).MG: WPos.EG = WPos.EG + PsqtWQ(Square).EG
    '--- Mobility
    MobCnt = 0

    For i = 0 To 7
      Offset = QueenOffsets(i): Target = Square + Offset: AttackBit = QAttackBit

      Do While Board(Target) <> FRAME
        WAttack(Target) = WAttack(Target) Or AttackBit

        Select Case Board(Target)
          Case NO_PIECE: If Not CBool(BAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
          Case WPAWN: SC.MG = SC.MG + 2: SC.EG = SC.EG + 2
            If RankNum > 3 Then If Board(Target + SQ_UP) >= NO_PIECE Then If Not CBool(BAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
            If Offset = SQ_UP_LEFT Or Offset = SQ_UP_RIGHT Then WAttack(Target + Offset) = WAttack(Target + Offset) Or QXrayAttackBit
            If CBool(BAttack(Target) And RBAttackBit) Then CheckWQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            Exit Do   'Defends pawn
          Case BPAWN:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1: If AttackBit = QAttackBit Then AddThreat COL_BLACK, PT_PAWN, PT_QUEEN, Square, Target
            Else
              If CBool(BAttack(Target) And RBAttackBit) Then CheckWQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            SC.MG = SC.MG + 7: SC.EG = SC.EG + 7
            Exit Do   'Attack pawn
          Case BKNIGHT:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1: If AttackBit = QAttackBit Then AddThreat COL_BLACK, PT_KNIGHT, PT_QUEEN, Square, Target
            Else
              If CBool(BAttack(Target) And RBAttackBit) Then CheckWQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            If AttackBit = QAttackBit Then AddThreat COL_BLACK, PieceType(Board(Target)), PT_QUEEN, Square, Target
            Exit Do
          Case BBISHOP:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1: If AttackBit = QAttackBit Then AddThreat COL_BLACK, PT_BISHOP, PT_QUEEN, Square, Target
            Else
              If CBool(BAttack(Target) And RBAttackBit) Then CheckWQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            If AttackBit = QAttackBit Then AddThreat COL_BLACK, PieceType(Board(Target)), PT_QUEEN, Square, Target
            If CBool(BAttack(Target) And RBAttackBit) Then CheckWQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            Exit Do
          Case BROOK:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1: If AttackBit = QAttackBit Then AddThreat COL_BLACK, PT_ROOK, PT_QUEEN, Square, Target
            Else
              If CBool(BAttack(Target) And RBAttackBit) Then CheckWQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            If AttackBit = QAttackBit Then AddThreat COL_BLACK, PieceType(Board(Target)), PT_QUEEN, Square, Target
            If CBool(BAttack(Target) And RBAttackBit) Then CheckWQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            Exit Do
          Case WKING: Exit Do ' ignore
          Case BKING: MobCnt = MobCnt + 1
            Exit Do
          Case BQUEEN: If AttackBit = QAttackBit Then AddThreat COL_BLACK, PT_QUEEN, PT_QUEEN, Square, Target: MobCnt = MobCnt + 1
            Exit Do
          Case WBISHOP:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1: SC.MG = SC.MG + 4: SC.EG = SC.EG + 2
            Else
              If CBool(BAttack(Target) And RBAttackBit) Then CheckWQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            If i > 3 Then AttackBit = QXrayAttackBit Else Exit Do
          Case WKNIGHT:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1
            Else
              If CBool(BAttack(Target) And RBAttackBit) Then CheckWQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            Exit Do
          Case WROOK:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then
              If Offset = 10 Then
                If WPawns(FileNum) = 0 Then
                  SC.MG = SC.MG + 10: SC.EG = SC.EG + 5
                ElseIf BPawns(FileNum) = 0 Then
                  SC.MG = SC.MG + 15: SC.EG = SC.EG + 5
                End If
              End If
              MobCnt = MobCnt + 1 '--- double lines
            Else
              If CBool(BAttack(Target) And RBAttackBit) Then CheckWQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            If i < 4 Then AttackBit = QXrayAttackBit Else Exit Do
          Case WEP_PIECE, BEP_PIECE: If Not CBool(BAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
          Case Else:
            Exit Do
        End Select

        Target = Target + Offset
      Loop

    Next

    AddScore WMobility, MobilityQ(MobCnt)
    If CBool(BAttack(Square) And PAttackBit) Then AddPawnThreat BThreat, COL_WHITE, PieceType(Board(Square)), Square
    AddScoreWithFactor SC, KingProtector(PT_QUEEN), MaxDistance(Square, WKingLoc) ' defends king?
    If QueenWeak Then SC.MG = SC.MG - 50: SC.EG = SC.EG - 10
    AddScore WPos, SC
    If bEvalTrace Then WriteTrace "WQueen: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & WPos.MG & ", " & WPos.EG
  Next a

  '--------------------------------------------------------------------
  '---- BLACK QUEENs ( last - full attack info needed for mobility ) --
  '--------------------------------------------------------------------
  For a = 1 To PieceSqListCnt(BQUEEN)
    Square = PieceSqList(BQUEEN, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = (9 - RankNum): SC.MG = 0: SC.EG = 0: QueenWeak = False
    BPos.MG = BPos.MG + PsqtBQ(Square).MG: BPos.EG = BPos.EG + PsqtBQ(Square).EG
    '--- Mobility
    MobCnt = 0

    For i = 0 To 7
      Offset = QueenOffsets(i): Target = Square + Offset: AttackBit = QAttackBit

      Do While Board(Target) <> FRAME
        BAttack(Target) = BAttack(Target) Or AttackBit

        Select Case Board(Target)
          Case NO_PIECE: If Not CBool(WAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
          Case BPAWN: SC.MG = SC.MG + 2: SC.EG = SC.EG + 2
            If RankNum < 6 Then If Board(Target + SQ_DOWN) >= NO_PIECE Then If Not CBool(WAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
            If Offset = SQ_DOWN_LEFT Or Offset = SQ_DOWN_RIGHT Then BAttack(Target + Offset) = BAttack(Target + Offset) Or QXrayAttackBit
            If CBool(WAttack(Target) And RBAttackBit) Then CheckBQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            Exit Do   'Defends pawn
          Case WPAWN:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1: If AttackBit = QAttackBit Then AddThreat COL_WHITE, PT_PAWN, PT_QUEEN, Square, Target
            Else
              If CBool(WAttack(Target) And RBAttackBit) Then CheckBQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            SC.MG = SC.MG + 7: SC.EG = SC.EG + 7
            Exit Do   'Attack pawn
          Case WKNIGHT:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1
            Else
              If CBool(WAttack(Target) And RBAttackBit) Then CheckBQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            If AttackBit = QAttackBit Then AddThreat COL_WHITE, PieceType(Board(Target)), PT_QUEEN, Square, Target
            Exit Do
          Case WBISHOP:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1
            Else
              If CBool(WAttack(Target) And RBAttackBit) Then CheckBQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            If AttackBit = QAttackBit Then AddThreat COL_WHITE, PieceType(Board(Target)), PT_QUEEN, Square, Target
            Exit Do
          Case WROOK:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1
            Else
              If CBool(WAttack(Target) And RBAttackBit) Then CheckBQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            If AttackBit = QAttackBit Then AddThreat COL_WHITE, PieceType(Board(Target)), PT_QUEEN, Square, Target
            Exit Do
          Case BKING: Exit Do ' Ignore
          Case WKING: MobCnt = MobCnt + 1
            Exit Do
          Case WQUEEN:  If AttackBit = QAttackBit Then AddThreat COL_WHITE, PT_QUEEN, PT_QUEEN, Square, Target: MobCnt = MobCnt + 1
            Exit Do
          Case BBISHOP:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1: SC.MG = SC.MG + 4: SC.EG = SC.EG + 2
            Else
              If CBool(WAttack(Target) And RBAttackBit) Then CheckBQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            If i > 3 Then AttackBit = QXrayAttackBit Else Exit Do
          Case BKNIGHT:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then
              MobCnt = MobCnt + 1
            Else
              If CBool(WAttack(Target) And RBAttackBit) Then CheckBQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            Exit Do
          Case BROOK:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then
              If Offset = -10 Then
                If BPawns(FileNum) = 0 Then
                  SC.MG = SC.MG + 10: SC.EG = SC.EG + 5
                ElseIf WPawns(FileNum) = 0 Then
                  SC.MG = SC.MG + 15: SC.EG = SC.EG + 5
                End If
              End If
              MobCnt = MobCnt + 1
            Else
              If CBool(WAttack(Target) And RBAttackBit) Then CheckBQueenWeek Target, Offset, i, QueenWeak ' pin oder discovered attack?
            End If
            If i < 4 Then AttackBit = QXrayAttackBit Else Exit Do
          Case WEP_PIECE, BEP_PIECE: If Not CBool(WAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
          Case Else:
            Exit Do
        End Select
        Target = Target + Offset
      Loop
    Next

    AddScore BMobility, MobilityQ(MobCnt)
    If CBool(WAttack(Square) And PAttackBit) Then AddPawnThreat WThreat, COL_BLACK, PieceType(Board(Square)), Square
    AddScoreWithFactor SC, KingProtector(PT_QUEEN), MaxDistance(Square, BKingLoc) ' defends king?
    If QueenWeak Then SC.MG = SC.MG - 50: SC.EG = SC.EG - 10
    AddScore BPos, SC
    If bEvalTrace Then WriteTrace "BQueen: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & BPos.MG & ", " & BPos.EG
  Next a

  '--------------------------------------------------------------------
  '---- Step 3.: Pass for pawn push ( full attack info needed for mobility )
  '--------------------------------------------------------------------
  SC = ZeroScore

  For a = 1 To PieceSqListCnt(WPAWN)
    Square = PieceSqList(WPAWN, a): RelRank = Rank(Square)

    ' bonus if safe pawn push attacks an enemy piece
    For rr = 1 To 1 + Abs(RelRank = 2)
      Target = Square + SQ_UP * rr
      If Board(Target) >= NO_PIECE Then ' empty or ep-dummy piece
        SC.MG = SC.MG + 8: SC.EG = SC.EG + 8 ' pawn mobility
        ' Safe pawn push: push field not attacked by opp pawn AND defend by own piece or not attacked by opp
        If BAttack(Target) = 0 Or WAttack(Target) > 0 Then
          If Not (rr = 2 And CBool(BAttack(Square + SQ_UP) And PAttackBit)) Then ' check EnPassant capture

            For i = 9 To 11 Step 2
              r = Board(Target + i)
              If PieceColor(r) = COL_BLACK And r <> BPAWN Then
                If Not CBool(WAttack(Target + i) And PAttackBit) Then ' already attacked by own pawn?
                  SC.MG = SC.MG + 38: SC.EG = SC.EG + 22 ' pawn threats non pawn enemy
                End If
              End If
            Next i

          End If
        End If
      Else
        Exit For
      End If
    Next
  Next a

  If SC.MG > 0 Then AddScore WPos, SC
  SC = ZeroScore

  For a = 1 To PieceSqListCnt(BPAWN)
    Square = PieceSqList(BPAWN, a): RelRank = (9 - Rank(Square))

    ' bonus if safe pawn push attacks an enemy piece
    For rr = 1 To 1 + Abs(RelRank = 2)
      Target = Square + SQ_DOWN * rr
      If Board(Target) >= NO_PIECE Then
        SC.MG = SC.MG + 8: SC.EG = SC.EG + 8 ' pawn mobility
        ' Safe pawn push: push field not attacked by opp pawn AND defend by own piece and not attacked by opp
        If WAttack(Target) = 0 Or BAttack(Target) > 0 Then
          If Not (rr = 2 And CBool(WAttack(Square + SQ_DOWN) And PAttackBit)) Then ' check EnPassant capture

            For i = 9 To 11 Step 2
              r = Board(Target - i)
              If PieceColor(r) = COL_WHITE And r <> WPAWN Then
                If Not CBool(BAttack(Target - i) And PAttackBit) Then ' already attacked by own pawn?
                  SC.MG = SC.MG + 38: SC.EG = SC.EG + 22 ' pawn threats non pawn enemy
                End If
              End If
            Next i

          End If
        End If
      Else
        Exit For
      End If
    Next rr
  Next a

  If SC.MG > 0 Then AddScore BPos, SC
  '--- End pass for pawn push <<<<
   
   
  '----------------------------------------------
  '--- Step 4:  King Safety   -------------------
  '----------------------------------------------
  If bEndgame Then
    WKSafety = ZeroScore: BKSafety = ZeroScore
  Else
    Dim Bonus            As Long
    Dim KingOnlyDefended As Long, bSafe As Boolean, Tropism As Long
    '----------------------------------------------
    '--- White King Safety Eval -------------------
    '----------------------------------------------
    RankNum = Rank(WKingLoc): FileNum = WKingFile: Bonus = 0
    If (PieceCnt(BQUEEN) * 2 + PieceCnt(BROOK)) > 1 Then
      KingDanger = 0
      If WPawnCnt = 0 Then MinWKingPawnDistance = 0 Else MinWKingPawnDistance = MinWKingPawnDistance - 1
      If RankNum > 4 Then
        WKSafety.EG = WKSafety.EG - 16 * MinWKingPawnDistance
      Else
        Bonus = WKingShelterStorm(WKingLoc)
        If WhiteCastled = NO_CASTLE Then
          If WKingLoc = SQ_E1 Then
            If WPawns(7) > 0 And PawnsWMin(7) < 4 Then
              If WCanCastleOO() Then
                Bonus = GetMax(Bonus, WKingShelterStorm(SQ_G1))
              End If
            End If
            If (WPawns(3) > 0 And PawnsWMin(3) < 4) Or (WPawns(2) > 0 And PawnsWMin(2) < 4) Then
              If WCanCastleOOO() Then
                Bonus = GetMax(Bonus, WKingShelterStorm(SQ_C1))
              End If
            End If
          End If
        End If
        AddScoreVal WKSafety, Bonus, -16 * MinWKingPawnDistance
      End If
      If bDoWKSafety Then
      
          ' King tropism: firstly, find squares that opponent attacks in our king flank
          ' Secondly, add the squares which are attacked twice in that flank
          GetKingFlankFiles WKingLoc, r, rr: Tropism = 0
          For k = SQ_A1 - 1 To SQ_A1 - 1 + 40 Step 10 ' start square - 1 of rank 1-5 (camp)
            For Square = k + r To k + rr     ' files king flank
              If BAttack(Square) <> 0 Then
                Tropism = Tropism + 1: If AttackBitCnt(BAttack(Square)) > 1 Then Tropism = Tropism + 1  ' Attacked twice?
              End If
            Next
          Next
          
          ' Pawnless king flank penalty
          k = 0
          For i = r To rr
            If WPawns(i) + BPawns(i) > 0 Then k = 1: Exit For
          Next
          If k = 0 Then MinusScore WKSafety, PawnlessFlank
                          
                  
          '--- Check threats at king ring
          Undefended = 0: KingOnlyDefended = 0: WKingAttPieces = 0: KingLevers = 0
          '  add the 2 or 3 squares in front of king ring: king G1 => F3+G3+H3
          If RankNum = 1 Then
            For Target = WKingLoc + 19 To WKingLoc + 21
              If Board(Target) <> FRAME Then
                If BAttack(Target) <> 0 Then
                  If WAttack(Target) = 0 Or WAttack(Target) = QAttackBit Then Undefended = Undefended + 1
                  ' exclude double pawn defended squares
                  If AttackBitCnt(WAttack(Target) And PAttackBit) < 2 Then WKingAttPieces = WKingAttPieces Or BAttack(Target)
                  If Board(Target) = WPAWN Then
                    If CBool(BAttack(Target) And PAttackBit) Then KingLevers = KingLevers + 1
                  End If
                End If
              End If
            Next
          End If

          For i = 0 To 7 ' for all directions from king square
            Offset = QueenOffsets(i): Target = WKingLoc + Offset
            If Board(Target) <> FRAME Then
              If BAttack(Target) <> 0 Then
                ' King attacks are added later in attack array, so distance=1 and WAttack=0 is equal to king attack only
                If WAttack(Target) = 0 Then KingOnlyDefended = KingOnlyDefended + 1
                WKingAdjacentZoneAttCnt = WKingAdjacentZoneAttCnt + AttackBitCnt(BAttack(Target) And Not PAttackBit)
                ' exclude double pawn defended squares
                If AttackBitCnt(WAttack(Target) And PAttackBit) < 2 Then WKingAttPieces = WKingAttPieces Or BAttack(Target)
                If Board(Target) = WPAWN Then
                  If CBool(BAttack(Target) And PAttackBit) Then KingLevers = KingLevers + 1
                End If
              End If
              rr = 1 ' rr=Distance to King

              Do  ' loop for a direction
                r = BAttack(Target)
                If CBool(r And QRBAttackBit) Then
                  bSafe = False ' Safe attack square?
                  If PieceColor(Board(Target)) <> BCOL Then
                    If WAttack(Target) = 0 Then
                      If rr = 1 Then
                        If AttackBitCnt(BAttack(Target)) > 1 Then bSafe = True
                      Else
                        bSafe = True
                      End If
                    End If
                  End If
                  ' Queen safe checks
                  If bSafe Then
                    If CBool(r And QAttackBit) Then
                      If Not CBool(WChecksCounted And QAttackBit) Then
                        KingDanger = KingDanger + QueenCheck
                        WChecksCounted = (WChecksCounted Or QAttackBit)
                      End If
                    End If
                  End If
                  If CBool(r And RBOrXrayAttackBit) Then
                    If Not bSafe And rr > 1 Then ' not defended by king
                      ' For minors and rooks, also consider the square as safe if attacked twice,
                      ' and only defended by our queen.
                      If CBool(WAttack(Target) = QAttackBit) Then
                        If AttackBitCnt(BAttack(Target)) > 1 Then
                          If Not (AttackBitCnt(WAttack(Target)) > 1 Or PieceColor(Board(Target)) = BCOL) Then
                            bSafe = True
                          End If
                        End If
                      End If
                    End If
                    '(i=0-3: orthogonal offset, 4-7:diagonal)
                    ' Rook checks
                    If i < 4 Then
                      If CBool(r And ROrXrayAttackBit) Then ' R1Attackbit or R2Attackbit set, if 2 rooks only one is counted per square
                        If bSafe Then
                          ' look for both rooks, different to SF
                          If CBool(r And R1OrXrayAttackBit) Then
                           If Not CBool(WChecksCounted And R1AttackBit) Then ' count only once per square!
                             If CBool(r And R1XrayAttackBit) Then
                               If SqBetween(Target, PieceSqList(BROOK, 1), WKingLoc) Then ' xray attack only if in direct line to opp king
                                 KingDanger = KingDanger + RookCheck \ 3: WChecksCounted = (WChecksCounted Or R1AttackBit)
                               Else
                                 KingDanger = KingDanger + 20 ' may be an attack plan
                               End If
                             Else
                               KingDanger = KingDanger + RookCheck: WChecksCounted = (WChecksCounted Or R1AttackBit)
                             End If
                           End If
                          End If
                          If CBool(r And R2OrXrayAttackBit) Then
                           If Not CBool(WChecksCounted And R2AttackBit) Then ' count only once per square!
                             If CBool(r And R2XrayAttackBit) Then
                               If SqBetween(Target, PieceSqList(BROOK, 2), WKingLoc) Then ' xray attack only if in direct line to opp king
                                 KingDanger = KingDanger + RookCheck \ 3: WChecksCounted = (WChecksCounted Or R2AttackBit)
                               Else
                                 KingDanger = KingDanger + 20 ' may be an attack plan
                               End If
                             Else
                               KingDanger = KingDanger + RookCheck: WChecksCounted = (WChecksCounted Or R2AttackBit)
                             End If
                           End If
                          End If
                        Else
                          WUnsafeChecks = WUnsafeChecks + 1
                        End If
                      End If
                    Else ' i >= 4
                      ' Bishop checks
                      If CBool(r And BXrayAttackBit) Then ' B1Attackbit or B2Attackbit set, if 2 rooks only one is counted
                        If Not CBool(WChecksCounted And B1AttackBit) Then ' count only once! only one bishop has same color as king
                          If bSafe Then
                             If CBool(r And BXrayAttackBit) Then
                               If SqBetween(Target, PieceSqList(BBISHOP, 1), WKingLoc) Or _
                                  SqBetween(Target, PieceSqList(BBISHOP, 2), WKingLoc) Then  ' xray attack only if in direct line to opp king
                                  ' do not count xray if through a blocked pawn
                                  If Board(Target + Offset) <> BPAWN Or Board(Target + Offset + SQ_DOWN) >= NO_PIECE Then KingDanger = KingDanger + BishopCheck \ 3: WChecksCounted = (WChecksCounted Or B1AttackBit)
                               Else
                                 KingDanger = KingDanger + 10 ' may be an attack plan
                               End If
                             Else
                               KingDanger = KingDanger + BishopCheck: WChecksCounted = (WChecksCounted Or B1AttackBit)
                             End If
                          Else
                            WUnsafeChecks = WUnsafeChecks + 1
                          End If
                        End If
                      End If
                    End If
                  End If
                End If ' r and QRBAttackbit
                If Board(Target) < NO_PIECE Then ' Piece found
                  '
                  '--- Check for pinned pieces
                  '
                  If (Board(Target) Mod 2 = WCOL) Then ' own piece, look for pinned
                    If i < 4 Then ' orthogonal
                      If CBool(BAttack(Target) And QRAttackBit) Then  ' rook or queen, direction not clear
                        For k = 1 To 7
                          sq = Target + Offset * k: Piece = Board(sq)
                          If Piece = FRAME Then Exit For
                          If Piece < NO_PIECE Then
                            If Piece = BQUEEN Or Piece = BROOK Then
                              If (PieceType(Piece) <> PieceType(Board(Target))) Then
                                If Piece = BROOK And Board(Target) = WQUEEN Then
                                  If bWhiteToMove Then
                                    AddScoreVal BThreat, 30, 50
                                    If BAttack(sq) <> 0 And WAttack(sq) = QAttackBit Then
                                      AddScoreVal BThreat, 75, 100 ' attacker defended? less because may be blocker move?
                                      If MaxDistance(Target, sq) = 1 Then AddScoreVal BThreat, 400, 500 ' no blocker option
                                    End If
                                  Else
                                    AddScoreVal BThreat, 1200, 1400
                                  End If
                                End If
                                WPinnedCnt = WPinnedCnt + 1
                                ' if pinned pawn then add bonus for attacker
                                If Board(Target) = WPAWN Then AddScoreVal BPos, ThreatByRank.MG \ 2 * Rank(Target), ThreatByRank.EG \ 2 * Rank(Target)
                              End If
                            End If
                            Exit For
                          Else
                            If Not (CBool(BAttack(sq) And QRAttackBit)) Then Exit For
                          End If
                        Next k
                      End If
                    Else ' i>4 diagonal
                      If CBool(BAttack(Target) And QBAttackBit) Then  ' bishop or queen, direction not clear

                        For k = 1 To 7
                          sq = Target + Offset * k: Piece = Board(sq)
                          If Piece = FRAME Then Exit For
                          If Piece < NO_PIECE Then
                            If Piece = BQUEEN Or Piece = BBISHOP Then
                              If (PieceType(Piece) <> PieceType(Board(Target))) Then
                                WPinnedCnt = WPinnedCnt + 1
                                If Piece = BBISHOP And Board(Target) = WQUEEN Then
                                  If bWhiteToMove Then
                                    AddScoreVal BThreat, 50, 70
                                    If BAttack(sq) <> 0 And WAttack(sq) = QAttackBit Then
                                      AddScoreVal BThreat, 100, 130  ' attacker defended? less because may be blocker move?
                                      If MaxDistance(Target, sq) = 1 Then AddScoreVal BThreat, 400, 500 ' no blocker option
                                    End If
                                  Else
                                    AddScoreVal BThreat, 1300, 1500
                                  End If
                                End If
                                ' if pinned pawn then add bonus for attacker (if pawn cannot capture attacker = distance>1)
                                If Board(Target) = WPAWN Then If MaxDistance(Target, sq) > 1 Or Offset < 0 Then AddScoreVal BPos, ThreatByRank.MG \ 2 * Rank(Target), ThreatByRank.EG \ 2 * Rank(Target)
                              End If
                            End If
                            Exit For
                          Else
                            If Not (CBool(BAttack(sq) And QBAttackBit)) Then Exit For
                          End If
                        Next k

                      End If
                    End If
                  End If
                  ' --- Piece found - exit direction loop
                  If Board(Target) <> WQUEEN Then Exit Do ' threat Q+K
                End If
                Target = Target + Offset: rr = rr + 1
              Loop While Board(Target) <> FRAME

            End If ' <<< Board(Target) <> FRAME
            
            ' Knight Check
            If PieceCnt(BKNIGHT) > 0 Then
              Target = WKingLoc + KnightOffsets(i)
              If Board(Target) <> FRAME Then
                If CBool(BAttack(Target) And NAttackBit) Then
                  bSafe = False ' Safe attack square?
                  If PieceColor(Board(Target)) <> BCOL Then If WAttack(Target) = 0 Then bSafe = True
                  If Not bSafe Then
                    If CBool(WAttack(Target) = QAttackBit) Then
                      If AttackBitCnt(BAttack(Target)) > 1 Then
                        If Not (AttackBitCnt(WAttack(Target)) > 1 Or PieceColor(Board(Target)) = BCOL) Then bSafe = True
                      End If
                    End If
                  End If
                  If Not CBool(WChecksCounted And N1AttackBit) Then ' count only once per square!
                    If bSafe Then
                      KingDanger = KingDanger + KnightCheck: WChecksCounted = (WChecksCounted Or N1AttackBit) ' only one knight check expected, two are very rare
                    Else
                      WUnsafeChecks = WUnsafeChecks + 1
                    End If
                  End If
                  ' Knight check fork?
                  If WAttack(Target) = 0 Or (WAttack(Target) = QAttackBit And (BAttack(Target) <> NAttackBit)) Then ' no attack

                   If PieceCnt(WQUEEN) + PieceCnt(WROOK) > 0 Then
                    For k = 0 To 7

                      Select Case Board(Target + KnightOffsets(k))
                        Case WQUEEN: AddScoreVal BThreat, 25, 35
                        Case WROOK: AddScoreVal BThreat, 15, 20
                      End Select

                    Next
                   End If
                  End If
                End If '<<< CBool(BAttack(Target) And NAttackBit)
              End If '<<<  Board(Target) <> FRAME
            End If '<<< PieceCnt(BKNIGHT) > 0
          Next i '<<< direction

          If WKingAttPieces <> 0 Then AddWKingAttackers WKingAttPieces

          If WKingAttackersCount > 1 - PieceCnt(BQUEEN) Then
                      
            ' total KingDanger
            KingDanger = KingDanger + WKingAttackersCount * WKingAttackersWeight + 65 * WKingAdjacentZoneAttCnt + Abs(KingLevers > 0) * 64 _
                         + 191 * (KingOnlyDefended + Undefended) + 152 * (WPinnedCnt + WUnsafeChecks) - 885 * Abs(PieceCnt(BQUEEN) = 0) - 6 * Bonus \ 8 - 30 _
                         + Tropism * Tropism \ 4 + (BMobility.MG - WMobility.MG)
            
            ' Penalty for king on open or semi-open file
            If NonPawnMaterial > 9000 And WPawns(FileNum) = 0 And WKingLoc <> WKING_START Then
              If BPawns(FileNum) = 0 Then KingDanger = KingDanger + 18 Else KingDanger = KingDanger + 9
            End If
            r = KingDanger + BPassedPawnAttack * 8 ' passed pawn attacking king?
            If r > 0 Then
              WKSafety.MG = WKSafety.MG - (r * r) \ 4096
              WKSafety.EG = WKSafety.EG - r \ 16
            End If
          End If
        
        End If

        
        ' Bonus for a dangerous pawn in the center near the opponent king, for instance pawn e5 against king g8.
        If FileNum >= 4 Then If Board(SQ_E4) = BPAWN Then WKSafety.MG = WKSafety.MG - 5
        If FileNum <= 5 Then If Board(SQ_D4) = BPAWN Then WKSafety.MG = WKSafety.MG - 5 ' both possible if king centered
        

        ' King tropism bonus, to anticipate slow motion attacks on our king
        WKSafety.MG = WKSafety.MG - 7 * Tropism ' closeEnemies
        
    End If
    
    
    '----------------------------------------------
    '--- Black King Safety Eval -------------------
    '----------------------------------------------
    RankNum = Rank(BKingLoc): RelRank = (9 - RankNum): FileNum = BKingFile: Bonus = 0: KingLevers = 0
    If (PieceCnt(WQUEEN) * 2 + PieceCnt(WROOK)) > 1 Then
      KingDanger = 0
      If BPawnCnt = 0 Then MinBKingPawnDistance = 0 Else MinBKingPawnDistance = MinBKingPawnDistance - 1
      If RelRank > 4 Then
        BKSafety.EG = BKSafety.EG - 16 * MinBKingPawnDistance
      Else
        Bonus = BKingShelterStorm(BKingLoc)
        If BlackCastled = NO_CASTLE Then
          If BKingLoc = SQ_E8 Then
            If BPawns(7) > 0 And PawnsBMax(7) > 5 Then
              If BCanCastleOO() Then
                Bonus = GetMax(Bonus, BKingShelterStorm(SQ_G8))
              End If
            End If
            If (BPawns(3) > 0 And PawnsBMax(3) > 5) Or (BPawns(2) > 0 And PawnsBMax(2) > 5) Then
              If BCanCastleOOO() Then
                Bonus = GetMax(Bonus, BKingShelterStorm(SQ_C8))
              End If
            End If
          End If
        End If
        AddScoreVal BKSafety, Bonus, -16 * MinBKingPawnDistance
      End If
      If bDoBKSafety Then
      
          ' King tropism: firstly, find squares that opponent attacks in our king flank
          ' Secondly, add the squares which are attacked twice in that flank
          GetKingFlankFiles BKingLoc, r, rr: Tropism = 0
          For k = SQ_A1 - 1 + 30 To SQ_A1 - 1 + 70 Step 10 ' start square - 1 of rank 5-8
            For Square = k + r To k + rr     ' files king flank
              If WAttack(Square) <> 0 Then
                Tropism = Tropism + 1: If AttackBitCnt(WAttack(Square)) > 1 Then Tropism = Tropism + 1  ' Attacked twice?
              End If
            Next
          Next
        
          ' Pawnless king flank penalty
          k = 0
          For i = r To rr
            If WPawns(i) + BPawns(i) > 0 Then k = 1: Exit For
          Next
          If k = 0 Then MinusScore BKSafety, PawnlessFlank
           
          '--- Check threats at king ring
          Undefended = 0: KingOnlyDefended = 0: BKingAttPieces = 0
          '  add the 2 or 3 squares in front of king ring: king G8 => F6+G6+H6
          If RankNum = 8 Then

            For Target = BKingLoc - 21 To BKingLoc - 19
              If Board(Target) <> FRAME Then
                If WAttack(Target) <> 0 Then
                  If BAttack(Target) = 0 Or BAttack(Target) = QAttackBit Then Undefended = Undefended + 1
                  ' exclude double pawn defended squares
                  If AttackBitCnt(BAttack(Target) And PAttackBit) < 2 Then BKingAttPieces = BKingAttPieces Or WAttack(Target)
                  If Board(Target) = BPAWN Then
                    If CBool(WAttack(Target) And PAttackBit) Then KingLevers = KingLevers + 1
                  End If
                End If
              End If
            Next

          End If

          For i = 0 To 7
            Offset = QueenOffsets(i): Target = BKingLoc + Offset
            If Board(Target) <> FRAME Then
              If WAttack(Target) <> 0 Then
                ' King attacks are added later in attack array, so distance=1 and BAttack=0 is equal to king attack only
                If BAttack(Target) = 0 Then KingOnlyDefended = KingOnlyDefended + 1
                BKingAdjacentZoneAttCnt = BKingAdjacentZoneAttCnt + AttackBitCnt(WAttack(Target) And Not PAttackBit)
                ' exclude double pawn defended squares
                If AttackBitCnt(BAttack(Target) And PAttackBit) < 2 Then BKingAttPieces = BKingAttPieces Or WAttack(Target)
                If Board(Target) = BPAWN Then
                  If CBool(WAttack(Target) And PAttackBit) Then KingLevers = KingLevers + 1
                End If
              End If
              rr = 1 ' rr=Distance to King

              Do ' loop for a direction
                r = WAttack(Target)
                If CBool(r And QRBAttackBit) Then
                  bSafe = False ' Safe attack square?
                  If PieceColor(Board(Target)) <> WCOL Then
                    If BAttack(Target) = 0 Then
                      If rr = 1 Then
                        If AttackBitCnt(WAttack(Target)) > 1 Then bSafe = True
                      Else
                        bSafe = True
                      End If
                    End If
                  End If
                  ' Queen safe checks
                  If bSafe Then
                    If CBool(r And QAttackBit) Then
                      If Not CBool(BChecksCounted And QAttackBit) Then
                        KingDanger = KingDanger + QueenCheck
                        BChecksCounted = (BChecksCounted Or QAttackBit)
                      End If
                    End If
                  End If
                  If CBool(r And RBOrXrayAttackBit) Then
                    If Not bSafe And rr > 1 Then ' not defended by king
                      ' For minors and rooks, also consider the square as safe if attacked twice,
                      ' and only defended by our queen.
                      If CBool(BAttack(Target) = QAttackBit) Then
                        If AttackBitCnt(WAttack(Target)) > 1 Then
                          If Not (AttackBitCnt(BAttack(Target)) > 1 Or PieceColor(Board(Target)) = WCOL) Then
                            bSafe = True
                          End If
                        End If
                      End If
                    End If
                    '(i=0-3: orthogonal offset, 4-7:diagonal)
                    ' Rook checks
                    If i < 4 Then
                      If CBool(r And ROrXrayAttackBit) Then ' R1Attackbit or R2Attackbit set, if 2 rooks only one is counted per square
                        If bSafe Then
                          ' look for both rooks, different to SF
                          If CBool(r And R1OrXrayAttackBit) Then
                           If Not CBool(BChecksCounted And R1AttackBit) Then ' count only once per square!
                             If CBool(r And R1XrayAttackBit) Then
                               If SqBetween(Target, PieceSqList(WROOK, 1), BKingLoc) Then ' xray attack only if in direct line to opp king
                                 KingDanger = KingDanger + RookCheck \ 3: BChecksCounted = (BChecksCounted Or R1AttackBit)
                               Else
                                 KingDanger = KingDanger + 20 ' may be an attack plan
                               End If
                             Else
                               KingDanger = KingDanger + RookCheck: BChecksCounted = (BChecksCounted Or R1AttackBit)
                             End If
                           End If
                          End If
                          '
                          If CBool(r And R2OrXrayAttackBit) Then
                           If Not CBool(BChecksCounted And R2AttackBit) Then ' count only once per square!
                             If CBool(r And R2XrayAttackBit) Then
                               If SqBetween(Target, PieceSqList(WROOK, 2), BKingLoc) Then ' xray attack only if in direct line to opp king
                                 KingDanger = KingDanger + RookCheck \ 3: BChecksCounted = (BChecksCounted Or R2AttackBit)
                               Else
                                 KingDanger = KingDanger + 20 ' may be an attack plan
                               End If
                             Else
                               KingDanger = KingDanger + RookCheck: BChecksCounted = (BChecksCounted Or R2AttackBit)
                             End If
                           End If
                          End If
                        Else
                          BUnsafeChecks = BUnsafeChecks + 1
                        End If
                      End If
                    Else ' i >= 4
                      ' Bishop checks
                      If CBool(r And BXrayAttackBit) Then ' B1Attackbit or B2Attackbit set, if 2 rooks only one is counted
                        If Not CBool(BChecksCounted And B1AttackBit) Then ' count only once! only one bishop has same color as king
                          If bSafe Then
                             If CBool(r And BXrayAttackBit) Then
                               If SqBetween(Target, PieceSqList(WBISHOP, 1), BKingLoc) Or _
                                  SqBetween(Target, PieceSqList(WBISHOP, 2), BKingLoc) Then  ' xray attack only if in direct line to opp king
                                 ' do not count xray if through a blocked pawn
                                 If Board(Target + Offset) <> WPAWN Or Board(Target + Offset + SQ_UP) >= NO_PIECE Then KingDanger = KingDanger + BishopCheck \ 3: BChecksCounted = (BChecksCounted Or B1AttackBit)
                               Else
                                 KingDanger = KingDanger + 10 ' may be an attack plan
                               End If
                             Else
                               KingDanger = KingDanger + BishopCheck: BChecksCounted = (BChecksCounted Or B1AttackBit)
                             End If
                          Else
                            BUnsafeChecks = BUnsafeChecks + 1
                          End If
                        End If
                      End If
                    End If
                  End If
                End If ' r and QRBAttackbit
                If Board(Target) < NO_PIECE Then ' Piece found
                  '
                  '--- Check for pinned pieces
                  '
                  If (Board(Target) Mod 2 = BCOL) Then ' own piece
                    If i < 4 Then ' orthogonal
                      If CBool(WAttack(Target) And QRAttackBit) Then  ' rook or queen, direction not clear
                        For k = 1 To 7
                          sq = Target + Offset * k: Piece = Board(sq)
                          If Piece = FRAME Then Exit For
                          If Piece < NO_PIECE Then
                            If Piece = WQUEEN Or Piece = WROOK Then
                              If (PieceType(Piece) <> PieceType(Board(Target))) Then
                                If Piece = WROOK And Board(Target) = BQUEEN Then
                                  If Not bWhiteToMove Then
                                    AddScoreVal WThreat, 30, 50
                                    If WAttack(sq) <> 0 And BAttack(sq) = QAttackBit Then
                                      AddScoreVal WThreat, 75, 100 ' attacker defended? less because may be blocker move?
                                      If MaxDistance(Target, sq) = 1 Then AddScoreVal WThreat, 400, 500 ' no blocker option
                                    End If
                                  Else
                                    AddScoreVal WThreat, 1200, 1400
                                  End If
                                End If
                               BPinnedCnt = BPinnedCnt + 1
                                ' if pinned pawn then add bonus for attacker
                                If Board(Target) = BPAWN Then AddScoreVal WPos, ThreatByRank.MG \ 2 * (9 - Rank(Target)), ThreatByRank.EG \ 2 * (9 - Rank(Target))
                              End If
                            End If
                            Exit For
                          Else
                            If Not CBool(WAttack(sq) And QRAttackBit) Then Exit For
                          End If
                        Next k
                      End If
                    Else ' i>4 diagonal
                      If CBool(WAttack(Target) And QBAttackBit) Then  ' bishop or queen, direction not clear

                        For k = 1 To 7
                          sq = Target + Offset * k: Piece = Board(sq)
                          If Piece = FRAME Then Exit For
                          If Piece < NO_PIECE Then
                            If Piece = WQUEEN Or Piece = WBISHOP Then
                              If (PieceType(Piece) <> PieceType(Board(Target))) Then
                                BPinnedCnt = BPinnedCnt + 1
                                If Piece = WBISHOP And Board(Target) = BQUEEN Then
                                  If Not bWhiteToMove Then
                                    AddScoreVal WThreat, 50, 70
                                    If WAttack(sq) <> 0 And BAttack(sq) = QAttackBit Then
                                      AddScoreVal WThreat, 100, 130 ' attacker defended? less because may be blocker move?
                                      If MaxDistance(Target, sq) = 1 Then AddScoreVal WThreat, 400, 500 ' no blocker option
                                    End If
                                  Else
                                    AddScoreVal WThreat, 1300, 1500
                                  End If
                                End If
                                ' if pinned pawn then add bonus for attacker (if pawn cannot capture attacker = distance>1)
                                If Board(Target) = BPAWN Then If MaxDistance(Target, sq) > 1 Or Offset > 0 Then AddScoreVal WPos, ThreatByRank.MG \ 2 * (9 - Rank(Target)), ThreatByRank.EG \ 2 * (9 - Rank(Target))
                              End If
                            End If
                            Exit For
                          Else
                            If Not CBool(WAttack(sq) And QBAttackBit) Then Exit For
                          End If
                        Next k

                      End If
                    End If
                  End If
                  ' --- Piece found - exit direction loop
                  If Board(Target) <> BQUEEN Then Exit Do
                End If
                Target = Target + Offset: rr = rr + 1
              Loop While Board(Target) <> FRAME

            End If ' <<< Board(Target) <> FRAME
            ' Knight Check
            If PieceCnt(WKNIGHT) > 0 Then
              Target = BKingLoc + KnightOffsets(i)
              If Board(Target) <> FRAME Then
                If CBool(WAttack(Target) And NAttackBit) Then
                  bSafe = False ' Safe attack square?
                  If PieceColor(Board(Target)) <> WCOL Then If BAttack(Target) = 0 Then bSafe = True
                  If Not bSafe Then
                    If CBool(BAttack(Target) = QAttackBit) Then
                      If AttackBitCnt(WAttack(Target)) > 1 Then
                        If Not (AttackBitCnt(BAttack(Target)) > 1 Or PieceColor(Board(Target)) = WCOL) Then
                          bSafe = True
                        End If
                      End If
                    End If
                  End If
                  If Not CBool(BChecksCounted And N1AttackBit) Then ' count only once per square!
                    If bSafe Then
                      KingDanger = KingDanger + KnightCheck: BChecksCounted = (BChecksCounted Or N1AttackBit) ' only one knight check expected, two are very rare
                    Else
                      BUnsafeChecks = BUnsafeChecks + 1
                    End If
                  End If
                  ' Knight check fork?
                  If BAttack(Target) = 0 Or (BAttack(Target) = QAttackBit And (WAttack(Target) <> NAttackBit)) Then ' field not defended or by queen only but other attacker

                   If PieceCnt(BQUEEN) + PieceCnt(BROOK) > 0 Then
                    For k = 0 To 7

                      Select Case Board(Target + KnightOffsets(k))
                        Case BQUEEN: AddScoreVal WThreat, 25, 35
                        Case BROOK: AddScoreVal WThreat, 15, 20
                      End Select

                    Next
                   End If

                  End If
                End If  '<<< CBool(WAttack(Target) And NAttackBit)
              End If ' <<< Board(Target) <> FRAME
            End If '<<< PieceCnt(WKNIGHT) > 0
          Next i '<<< direction
          
          If BKingAttPieces <> 0 Then AddBKingAttackers BKingAttPieces
          
          If BKingAttackersCount > 1 - PieceCnt(WQUEEN) Then
  

            ' total KingDanger
            KingDanger = KingDanger + BKingAttackersCount * BKingAttackersWeight + 65 * BKingAdjacentZoneAttCnt + Abs(KingLevers > 0) * 64 _
                         + 191 * (KingOnlyDefended + Undefended) + 152 * (BPinnedCnt + BUnsafeChecks) - 885 * Abs(PieceCnt(WQUEEN) = 0) - 6 * Bonus \ 8 - 30 _
                         + Tropism * Tropism \ 4 + (WMobility.MG - BMobility.MG)
                         
            ' Penalty for king on open or semi-open file
            If NonPawnMaterial > 9000 And BPawns(FileNum) = 0 And BKingLoc <> BKING_START Then
              If WPawns(FileNum) = 0 Then KingDanger = KingDanger + 18 Else KingDanger = KingDanger + 9
            End If
            r = KingDanger + WPassedPawnAttack * 8 ' passed pawn attacking king?
            If r > 0 Then
              BKSafety.MG = BKSafety.MG - (r * r) \ 4096
              BKSafety.EG = BKSafety.EG - r \ 16
            End If
          End If
        
        End If
        
        ' Bonus for a dangerous pawn in the center near the opponent king, for instance pawn e5 against king g8.
        If FileNum >= 4 Then If Board(SQ_E5) = WPAWN Then BKSafety.MG = BKSafety.MG - 5
        If FileNum <= 5 Then If Board(SQ_D5) = WPAWN Then BKSafety.MG = BKSafety.MG - 5 ' both possible if king centered
        
       
        ' King tropism bonus, to anticipate slow motion attacks on our king
        BKSafety.MG = BKSafety.MG - 7 * Tropism ' closeEnemies
        
      End If
  End If ' Endgame
  
  
  '--- Endgame King distance to best pawn. Not in PawnHash because "Fifty" may be different
  If bEndgame Or (WPawnCnt + BPawnCnt <= 8) Then
    If WBestPawn > 0 Then
      i = MaxDistance(WBestPawn, WKingLoc)
      AddScoreVal WPos, 0, (7 - i) * (7 - i) * 6
      If Rank(WBestPawn) >= 5 Then AddScoreVal WPos, 0, ((Rank(WBestPawn - 4) * Rank(WBestPawn - 4)) * (Fifty + 1)) \ 3 * 2 '--- bonus for move pawn in endgame
    ElseIf BBestPawn > 0 Then
      i = MaxDistance(BBestPawn, WKingLoc)     '--- Close to Opp Pawn
      AddScoreVal WPos, 0, (7 - i) * (7 - i)
    End If
    If BBestPawn > 0 Then
      i = MaxDistance(BBestPawn, BKingLoc)
      If i > 2 Then AddScoreVal BPos, 0, (7 - i) * (7 - i) * 6
      If RankB(BBestPawn) >= 5 Then AddScoreVal BPos, 0, ((RankB(BBestPawn - 4) * RankB(BBestPawn - 4)) * (Fifty + 1)) \ 3 * 2 '--- bonus for move pawn in endgame
    ElseIf WBestPawn > 0 Then
      i = MaxDistance(WBestPawn, BKingLoc)     '--- Close to Opp Pawn
      AddScoreVal BPos, 0, (7 - i) * (7 - i)
    End If
  Else
    '--- Midgame
  End If
  ' add kings to attack array
  r = 0: rr = 0

  For i = 0 To 7
    Offset = QueenOffsets(i)
    Target = WKingLoc + Offset
    If Board(Target) <> FRAME Then
      WAttack(Target) = WAttack(Target) Or KAttackBit
      If PieceColor(Board(Target)) = COL_BLACK Then
        If AttackBitCnt(WAttack(Target)) > AttackBitCnt(BAttack(Target)) Then r = r + 1  ' King attacks unprotected piece
      End If
    End If
    Target = BKingLoc + Offset
    If Board(Target) <> FRAME Then
      BAttack(Target) = BAttack(Target) Or KAttackBit
      If PieceColor(Board(Target)) = COL_WHITE Then
        If AttackBitCnt(BAttack(Target)) > AttackBitCnt(WAttack(Target)) Then rr = rr + 1  ' King attacks unprotected piece
      End If
    End If
  Next

  If r > 0 Then
    If r = 1 Then AddScore WThreat, KingOnOneBonus Else AddScore WThreat, KingOnManyBonus
  End If
  If rr > 0 Then
    If rr = 1 Then AddScore BThreat, KingOnOneBonus Else AddScore BThreat, KingOnManyBonus
  End If
  
  '--------------------------------------------------
  '--- Step 6: Eval threats -------------------------
  '--------------------------------------------------
  CalcThreats  ' in WThreat and BThreat
  
  If WWeakUnopposedCnt > 0 Then
    If PieceCnt(BQUEEN) + PieceCnt(BROOK) > 0 Then AddScoreWithFactor BThreat, WeakUnopposedPawn, WWeakUnopposedCnt
  End If
  If BWeakUnopposedCnt > 0 Then
    If PieceCnt(WQUEEN) + PieceCnt(WROOK) > 0 Then AddScoreWithFactor WThreat, WeakUnopposedPawn, BWeakUnopposedCnt
  End If

  ' Trapped bishops at a7/h7, a2/h2
  If PieceCnt(WBISHOP) > 0 Then
    ' white bishop not defended trapped at A7 by black pawn B6 (or if pawn can move to B6)
    If Board(SQ_A7) = WBISHOP Then
      If BAttack(SQ_B6) > 0 And WAttack(SQ_A7) = 0 Then
        If Board(SQ_B6) = BPAWN Or (Not bWhiteToMove And Board(SQ_B6) >= NO_PIECE And Board(SQ_B7) = BPAWN) Then
          AddScoreVal BThreat, ScoreBishop.MG \ 3, ScoreBishop.MG \ 4
        End If
      End If
    End If
    If Board(SQ_H7) = WBISHOP Then
      If BAttack(SQ_G6) > 0 And WAttack(SQ_H7) = 0 Then
        If Board(SQ_G6) = BPAWN Or (Not bWhiteToMove And Board(SQ_G6) >= NO_PIECE And Board(SQ_G7) = BPAWN) Then
          AddScoreVal BThreat, ScoreBishop.MG \ 3, ScoreBishop.MG \ 4
        End If
      End If
    End If
  End If
  If PieceCnt(BBISHOP) > 0 Then
    If Board(SQ_A2) = BBISHOP Then
      If WAttack(SQ_B3) > 0 And BAttack(SQ_A2) = 0 Then
        If Board(SQ_B3) = WPAWN Or (bWhiteToMove And Board(SQ_B3) >= NO_PIECE And Board(SQ_B2) = WPAWN) Then
          AddScoreVal WThreat, ScoreBishop.MG \ 3, ScoreBishop.MG \ 4
        End If
      End If
    End If
    If Board(SQ_H2) = BBISHOP Then
      If WAttack(SQ_G3) > 0 And BAttack(SQ_H2) = 0 Then
        If Board(SQ_G3) = WPAWN Or (bWhiteToMove And Board(SQ_G3) >= NO_PIECE And Board(SQ_G2) = WPAWN) Then
          AddScoreVal WThreat, ScoreBishop.MG \ 3, ScoreBishop.MG \ 4
        End If
      End If
    End If
  End If
  '
  '--- Passed pawns (white and black). done here because full attack info is needed
  '
  WFrontMostPassedPawnRank = 0: BFrontMostPassedPawnRank = 0

  For a = 1 To PassedPawnsCnt
    Dim AttackedFromBehind As Long, DefendedFromBehind As Long
    Square = PassedPawns(a): FileNum = File(Square): RankNum = Rank(Square)
    MBonus = 0: EBonus = 0: UnsafeCnt = 0
    If PieceColor(Board(Square)) = COL_WHITE Then
      ' White piece
      OwnCol = COL_WHITE: OppCol = COL_BLACK: MoveUp = SQ_UP
      RelRank = RankNum: OwnKingLoc = WKingLoc: OppKingLoc = BKingLoc
      ' Attack Opp King?
      If RelRank > WFrontMostPassedPawnRank Then WFrontMostPassedPawnRank = RankNum
      If PieceCnt(WBISHOP) > 0 Then    ' Bishop with same color as promote square? ( not SF logic )
        sq = SQ_A1 + FileNum - 1 + 7 * MoveUp
        If ColorSq(sq) = COL_WHITE Then
          r = Sgn(Sgn(WBishopsOnWhiteSq) - Sgn(BBishopsOnWhiteSq)) ' 0 if both sides have same bishop color, else +1 or -1
          If r <> 0 Then MBonus = MBonus + r * (10 + (RelRank - 1) * 3): EBonus = EBonus + r * (20 + (RelRank - 1) * (RelRank - 1))
        Else
          r = Sgn(Sgn(WBishopsOnBlackSq) - Sgn(BBishopsOnBlackSq)) ' 0 if both sides have same bishop color, else +1 or -1
          If r <> 0 Then MBonus = MBonus + r * (10 + (RelRank - 1) * 3): EBonus = EBonus + r * (20 + (RelRank - 1) * (RelRank - 1))
        End If
      End If
    Else
      OwnCol = COL_BLACK: OppCol = COL_WHITE: MoveUp = SQ_DOWN
      ' Black piece
      RelRank = (9 - RankNum):  OwnKingLoc = BKingLoc: OppKingLoc = WKingLoc
      If RelRank > BFrontMostPassedPawnRank Then BFrontMostPassedPawnRank = RelRank
      If PieceCnt(BBISHOP) > 0 Then  ' Bishop with same color as promote square?
        sq = SQ_A1 + FileNum - 1
        If ColorSq(sq) = COL_WHITE Then
          r = Sgn(Sgn(BBishopsOnWhiteSq) - Sgn(WBishopsOnWhiteSq)) ' 0 if both sides have same bishop color, else +1 or -1
          If r <> 0 Then MBonus = MBonus + r * (10 + (RelRank - 1) * 3): EBonus = EBonus + r * (20 + (RelRank - 1) * (RelRank - 1))
        Else
          r = Sgn(Sgn(BBishopsOnBlackSq) - Sgn(WBishopsOnBlackSq)) ' 0 if both sides have same bishop color, else +1 or -1
          If r <> 0 Then MBonus = MBonus + r * (10 + (RelRank - 1) * 3): EBonus = EBonus + r * (20 + (RelRank - 1) * (RelRank - 1))
        End If
      End If
    End If
    '
    '--- Path to promote square blocked? => penalty
    '
    r = RelRank
    rr = PassedDanger(RelRank)
    MBonus = MBonus + PassedPawnRankBonus(r).MG: EBonus = EBonus + PassedPawnRankBonus(r).EG
    ' Bonus based on rank ' SF9
    If rr <> 0 Then
      BlockSq = Square + MoveUp
      If Board(BlockSq) <> FRAME Then
        '  Adjust bonus based on the king's proximity
        AttackedFromBehind = 0: DefendedFromBehind = 0
        EBonus = EBonus + (GetMin(5, MaxDistance(BlockSq, OppKingLoc)) * 5 - GetMin(5, MaxDistance(BlockSq, OwnKingLoc)) * 2) * rr
        'If blockSq is not the queening square then consider also a second push
        If RelRank <> 7 Then EBonus = EBonus - MaxDistance(BlockSq + MoveUp, OwnKingLoc) * rr
        'If the pawn is free to advance, then increase the bonus
        If Board(BlockSq) >= NO_PIECE Then
          k = 0: bAllDefended = True: BlockSqDefended = True: BlockSqUnsafe = False
          ' Rook or Queen attacking/defending from behind?
          If CBool(BAttack(Square) And QRAttackBit) Or CBool(WAttack(Square) And QRAttackBit) Then
            sq = Square
            For RankPath = RelRank - 1 To 1 Step -1
              sq = sq - MoveUp ' move down to rank 1
              Select Case Board(sq)
                Case NO_PIECE:
                Case BROOK, BQUEEN:
                  If OwnCol = COL_WHITE Then
                    BlockSqUnsafe = True: AttackedFromBehind = 1
                  Else
                    DefendedFromBehind = 1:
                  End If
                  Exit For
                Case WROOK, WQUEEN:
                  If OwnCol = COL_BLACK Then
                    BlockSqUnsafe = True: AttackedFromBehind = 1
                  Else
                    DefendedFromBehind = 1
                  End If
                  Exit For
                Case Else:
                  Exit For
              End Select
            Next
          End If

          sq = Square
          For RankPath = RelRank + 1 To 8
            sq = sq + MoveUp
            OwnAttCnt = AttackBitCnt(AttackByCol(OwnCol, sq)) + DefendedFromBehind
            If OwnAttCnt = 0 And sq <> OwnKingLoc Then
              bAllDefended = False: If sq = BlockSq Then BlockSqDefended = False
            End If
            If PieceColor(Board(sq)) = OppCol Then
              If sq = BlockSq Then BlockSqUnsafe = True Else UnsafeCnt = UnsafeCnt + 1
            ElseIf AttackBitCnt(AttackByCol(OppCol, sq)) + AttackedFromBehind > 0 Then
              If CBool(AttackByCol(OwnCol, sq) And PAttackBit) And Not CBool(AttackByCol(OppCol, sq) And PAttackBit) Then
                ' Own pawn support but no enemy pawn attack: square is safe ( NOT SF LOGIC )
              Else
                If sq = BlockSq Then BlockSqUnsafe = True Else UnsafeCnt = UnsafeCnt + 1
              End If
            End If
          Next RankPath

          If BlockSqUnsafe Then UnsafeCnt = UnsafeCnt + 1
          If UnsafeCnt = 0 Then
            k = 20
          ElseIf Not BlockSqUnsafe Then
            k = 9 '- UnsafeCnt
          Else
            k = 0
          End If
          If bAllDefended Then
            k = k + 6 '- UnsafeCnt \ 2
          ElseIf BlockSqDefended Then
            k = k + 4 '- UnsafeCnt \ 2
          End If
          ' If protected by more than one rook or queen, assign extra bonus
          If k > 0 Then
            If OwnCol = COL_WHITE Then
              If AttackBitCnt((WAttack(Square) And QOrXrayROrXrayAttackBit)) > 1 Then k = k + 2
            Else
              If AttackBitCnt((BAttack(Square) And QOrXrayROrXrayAttackBit)) > 1 Then k = k + 2
            End If
          End If
          '-- add bonus
          If k <> 0 Then MBonus = MBonus + k * rr: EBonus = EBonus + k * rr
        Else
          If PieceColor(Board(BlockSq)) = OwnCol Then MBonus = MBonus + rr + (r - 1) * 2: EBonus = EBonus + rr + (r - 1) * 2 ' r-1 because rank r is 0 based in C
        End If
      End If
    End If ' rr>0
    '
    If UnsafeCnt > 0 Then MBonus = MBonus - UnsafeCnt * 8: EBonus = EBonus - UnsafeCnt ' hindered passed pawn
    '
    If OwnCol = COL_WHITE Then
      If WPawnCnt > BPawnCnt Then EBonus = EBonus + EBonus \ 4
      MBonus = MBonus + PassedPawnFileBonus(FileNum).MG: EBonus = EBonus + PassedPawnFileBonus(FileNum).EG
      If BNonPawnMaterial = 0 Then EBonus = EBonus + 20
      If Board(Square + SQ_UP) = BPAWN Then MBonus = MBonus \ 2: EBonus = EBonus \ 2 ' supporter sacrify needed
      If bWhiteToMove Then MBonus = (MBonus * 105) \ 100:   EBonus = (EBonus * 105) \ 100
      AddScoreVal WPassed, MBonus, EBonus
      If 1000 + EBonus > WBestPawnVal Then WBestPawn = Square: WBestPawnVal = 1000 + EBonus ' new best pawn
      If bEvalTrace Then WriteTrace "WPassed: " & LocCoord(Square) & ">" & MBonus & ", " & EBonus
    ElseIf OwnCol = COL_BLACK Then
      If BPawnCnt > WPawnCnt Then EBonus = EBonus + EBonus \ 4
      MBonus = MBonus + PassedPawnFileBonus(FileNum).MG: EBonus = EBonus + PassedPawnFileBonus(FileNum).EG
      If WNonPawnMaterial = 0 Then EBonus = EBonus + 20
      If Board(Square + SQ_DOWN) = WPAWN Then MBonus = MBonus \ 2: EBonus = EBonus \ 2 ' supporter sacrify needed
      If Not bWhiteToMove Then MBonus = (MBonus * 105) \ 100:   EBonus = (EBonus * 105) \ 100
      AddScoreVal BPassed, MBonus, EBonus
      If 1000 + EBonus > BBestPawnVal Then BBestPawn = Square: BBestPawnVal = 1000 + EBonus ' new best pawn
      If bEvalTrace Then WriteTrace "BPassed: " & LocCoord(Square) & ">" & MBonus & ", " & EBonus
    End If
  Next a
  '
  '---<<< end  Passed pawn
  '
  '--- If both sides have only pawns, score for potential unstoppable pawns
  If WNonPawnMaterial + BNonPawnMaterial = 0 Then
    If WFrontMostPassedPawnRank > 0 Then AddScoreVal WPassed, 0, WFrontMostPassedPawnRank * 20
    If BFrontMostPassedPawnRank > 0 Then AddScoreVal BPassed, 0, BFrontMostPassedPawnRank * 20 ' RelRank is used, so >0 is correct
  End If
  '
  '---  Penalty for pawns on same color square of bishop
  '
  If PieceCnt(WBISHOP) > 0 Then
    r = WPawnCntOnWhiteSq * WBishopsOnWhiteSq + (WPawnCnt - WPawnCntOnWhiteSq) * WBishopsOnBlackSq
    If r <> 0 Then
      r = r * (1 + WCenterPawnsBlocked)
      AddScoreVal WPos, -3 * r, -5 * r
    End If
    ' Bonus for bishop on a long diagonal if it can "see" both center squares and no pawns
    If WBishopsOnWhiteSq > 0 And Not bEndgame Then
      If CBool(WAttack(SQ_E4) And BAttackBit) Then
        If PieceType(Board(SQ_E4)) <> PT_PAWN Then
          If CBool(WAttack(SQ_D5) And BAttackBit) Then If PieceType(Board(SQ_D5)) <> PT_PAWN Then WPos.MG = WPos.MG + 22
        End If
      End If
    End If
    If WBishopsOnBlackSq > 0 Then
      If CBool(WAttack(SQ_D4) And BAttackBit) Then
        If PieceType(Board(SQ_D4)) <> PT_PAWN Then
          If CBool(WAttack(SQ_E5) And BAttackBit) Then If PieceType(Board(SQ_E5)) <> PT_PAWN Then WPos.MG = WPos.MG + 22
        End If
      End If
    End If
  End If
  If PieceCnt(BBISHOP) > 0 Then
    r = BPawnCntOnWhiteSq * BBishopsOnWhiteSq + (BPawnCnt - BPawnCntOnWhiteSq) * BBishopsOnBlackSq
    If r <> 0 Then
      r = r * (1 + BCenterPawnsBlocked)
      AddScoreVal BPos, -3 * r, -5 * r
    End If
    ' Bonus for bishop on a long diagonal if it can "see" both center squares and no pawns
    If BBishopsOnWhiteSq > 0 And Not bEndgame Then
      If CBool(BAttack(SQ_D5) And BAttackBit) Then
        If PieceType(Board(SQ_D5)) <> PT_PAWN Then
          If CBool(BAttack(SQ_E4) And BAttackBit) Then If PieceType(Board(SQ_E4)) <> PT_PAWN Then BPos.MG = BPos.MG + 22
        End If
      End If
    End If
    If BBishopsOnBlackSq > 0 Then
      If CBool(BAttack(SQ_E5) And BAttackBit) Then
        If PieceType(Board(SQ_E5)) <> PT_PAWN Then
          If CBool(BAttack(SQ_D4) And BAttackBit) Then If PieceType(Board(SQ_D4)) <> PT_PAWN Then BPos.MG = BPos.MG + 22
        End If
      End If
    End If
  End If
  '
  '--->>> Pawn Islands (groups of pawns) ---
  '
  r = 0: bWIsland = False  ' r : white islands
  rr = 0: bBIsland = False ' rr: black islands

  For FileNum = 1 To 9
    If WPawns(FileNum) <= 0 Then  ' File WPawns(9) = -1
      bWIsland = True
    ElseIf bWIsland Then
      r = r + 1: bWIsland = False ' empty file and pawn onleft side > island
    End If
    If BPawns(FileNum) <= 0 Then  ' File BPawns(9) = -1
      bBIsland = True
    ElseIf bBIsland Then
      rr = rr + 1: bBIsland = False ' empty file and pawn onleft side > island
    End If
  Next

  If r > 0 Then AddScoreVal WPawnStruct, -15 * r, -25 * r ' Penalty for each island
  If rr > 0 Then AddScoreVal BPawnStruct, -15 * rr, -25 * rr
  '---<<< end Pawn Islands ---
  '
  '-----------------------------------------------------------------------------
  '--- Step 7: Calculate total material values and endgame scale factors     ---
  '-----------------------------------------------------------------------------
  '
  ' Piece values were set in SetGamePhase
  Dim AllTotal As TScore, MatEval As Long
  AllTotal.MG = Material ' Based on MG, no need to recalc ' (PieceCnt(WQUEEN) - PieceCnt(BQUEEN)) * ScoreQueen.MG + (PieceCnt(WROOK) - PieceCnt(BROOK)) * ScoreRook.MG + (PieceCnt(WBISHOP) - PieceCnt(BBISHOP)) * ScoreBishop.MG + (PieceCnt(WKNIGHT) - PieceCnt(BKNIGHT)) * ScoreKnight.MG + (WPawnCnt - BPawnCnt) * ScorePawn.MG
  AllTotal.EG = (PieceCnt(WQUEEN) - PieceCnt(BQUEEN)) * ScoreQueen.EG + (PieceCnt(WROOK) - PieceCnt(BROOK)) * ScoreRook.EG + (PieceCnt(WBISHOP) - PieceCnt(BBISHOP)) * ScoreBishop.EG + (PieceCnt(WKNIGHT) - PieceCnt(BKNIGHT)) * ScoreKnight.EG + (WPawnCnt - BPawnCnt) * ScorePawn.EG
  MatEval = ScaleScore(AllTotal)
  If bEvalTrace Then
    Debug.Print "Material: " & EvalSFTo100(AllTotal.MG) & "," & EvalSFTo100(AllTotal.EG)
  End If

  '
  '--- Calculate SPACE in opening phase for safe squares in center (files 3-6, ranks 2-4)
  '
  If NonPawnMaterial > SPACE_THRESHOLD Then
    r = 0: rr = 0
    For k = 3 To 6  ' files 3-6
      '--- White space
      Offset = PawnsWMin(k): Target = k + 20
      For RankNum = 2 To 4 ' WHITE
        Target = Target + 10
        If Board(Target) <> WPAWN Then
          If Not CBool(BAttack(Target) And PAttackBit) Then
            r = r + 1: If RankNum < Offset Then If RankNum >= Offset - 3 Then r = r + 1 ' extra bonus if at most three squares behind some friendly pawn
          End If
        End If
      Next
      
      '--- Black space
      Offset = PawnsBMin(k): Target = k + 50
      For RankNum = 5 To 7 '
        Target = Target + 10
        If Board(Target) <> BPAWN Then
          If Not CBool(WAttack(Target) And PAttackBit) Then
            rr = rr + 1: If RankNum <= Offset + 3 And RankNum > Offset Then rr = rr + 1 ' extra bonus if at most three squares behind some friendly pawn
          End If
        End If
      Next
    Next

    If r + rr <> 0 Then
      ' weight for space
      k = 0
      For i = 1 To 8 ' count open files
        If WPawns(i) = 0 Then If BPawns(i) = 0 Then k = k + 1
      Next
      If r > 0 Then
        a = WNonPawnPieces + 1 + WPawnCnt - 2 * k
        WPos.MG = WPos.MG + r * a * a \ 16
      End If
      If rr > 0 Then
        a = BNonPawnPieces + 1 + BPawnCnt - 2 * k
        BPos.MG = BPos.MG + rr * a * a \ 16
      End If
    End If
  End If ' <<< Space


  '-----------------------------------------------
  '--- Step 8: Calculate weights and total eval -
  '-----------------------------------------------
  '
  '--- evaluate_initiative() /Complexity computes the initiative correction value for the
  '--- position, i.e., second order bonus/malus based on the known attacking/defending status of the players.
  '--- Semiopenfiles \12 because tricky counting to avoid count duplicate pawns per file
  '------ REMOVED: slightly better result without this logic. More complex better because no EGTB?
'  k = 12 * (Abs(WKingFile - BKingFile) - Abs(Rank(WKingLoc) - Rank(BKingLoc))) _
'      + 8 * (Abs(WSemiOpenFiles + BSemiOpenFiles) \ 12 + PassedPawnsCnt) _
'      + 12 * (WPawnCnt + BPawnCnt) _
'      + 16 * Abs(KingSidePawns > 0 And QueenSidePawns > 0) _
'      + 48 * Abs(NonPawnMaterial = 0) _
'      - 136
'
'  rr = MatEval + (WPos.EG - BPos.EG) + (WPassed.EG - BPassed.EG) ' strong side?
'  If rr > 0 Then
'    WPos.EG = WPos.EG + GetMax(k, -Abs(rr))
'  ElseIf rr < 0 Then
'    BPos.EG = BPos.EG + GetMax(k, -Abs(rr))
'  End If

  '
  '--- Material Imbalance / Score trades
  '
  Dim TradeEval       As Long
  If MatEval = 0 Then
    TradeEval = 0
  Else
    TradeEval = Imbalance() ' SF6
    AllTotal.MG = AllTotal.MG + TradeEval: AllTotal.EG = AllTotal.EG + TradeEval
  End If
  AllTotal.MG = AllTotal.MG + ((WPos.MG - BPos.MG) * PiecePosScaleFactor) \ 100&
  AllTotal.EG = AllTotal.EG + ((WPos.EG - BPos.EG) * PiecePosScaleFactor) \ 100&
  AllTotal.MG = AllTotal.MG + ((WPawnStruct.MG - BPawnStruct.MG) * PawnStructScaleFactor) \ 100&
  AllTotal.EG = AllTotal.EG + ((WPawnStruct.EG - BPawnStruct.EG) * PawnStructScaleFactor) \ 100&
  AllTotal.MG = AllTotal.MG + ((WPassed.MG - BPassed.MG) * PassedPawnsScaleFactor) \ 100&
  AllTotal.EG = AllTotal.EG + ((WPassed.EG - BPassed.EG) * PassedPawnsScaleFactor) \ 100&
  AllTotal.MG = AllTotal.MG + ((WMobility.MG - BMobility.MG) * MobilityScaleFactor) \ 100&
  AllTotal.EG = AllTotal.EG + ((WMobility.EG - BMobility.EG) * MobilityScaleFactor) \ 100&
  '
  ' different weights for defending computer king / attacking opp king
  If bCompIsWhite Then
    WKingScaleFactor = CompKingDefScaleFactor: BKingScaleFactor = OppKingAttScaleFactor
  Else
    BKingScaleFactor = CompKingDefScaleFactor: WKingScaleFactor = OppKingAttScaleFactor
  End If
  If bWhiteToMove Then
    WKingScaleFactor = WKingScaleFactor + 5
  Else
    BKingScaleFactor = BKingScaleFactor + 5
  End If
  
  AllTotal.MG = AllTotal.MG + (WKSafety.MG * WKingScaleFactor) \ 100&
  AllTotal.EG = AllTotal.EG + (WKSafety.EG * WKingScaleFactor) \ 100&
  AllTotal.MG = AllTotal.MG - (BKSafety.MG * BKingScaleFactor) \ 100&
  AllTotal.EG = AllTotal.EG - (BKSafety.EG * BKingScaleFactor) \ 100&
  AllTotal.MG = AllTotal.MG + ((WThreat.MG - BThreat.MG) * ThreatsScaleFactor) \ 100&
  AllTotal.EG = AllTotal.EG + ((WThreat.EG - BThreat.EG) * ThreatsScaleFactor) \ 100&
  '---------------------------------------------------------------------------------------------------
  
  '
  '--- Scale Factor ---
  '
  ScaleFactor = SCALE_FACTOR_NORMAL ' Normal ScaleFactor, scales EG value only
  If GamePhase < PHASE_MIDGAME Then
    ' KRPPKRP endgame
    'if the defending king is actively placed and not passed pawn for strong side, the position is drawish
    If WNonPawnMaterial = ScoreRook.MG And BNonPawnMaterial = ScoreRook.MG Then
      If WPawnCnt = 2 And BPawnCnt = 1 Then  ' white is strong side
        If WFrontMostPassedPawnRank = 0 Then ' no passed pawn for strong side
          Square = PieceSqList(WPAWN, 1) ' 1. pawn
          If Rank(BKingLoc) > Rank(Square) Then ' Opp king in front
            If Abs(File(Square) - File(BKingLoc)) <= 1 Then ' and nearby file
              r = Rank(Square): Square = PieceSqList(WPAWN, 2) ' 2. pawn
              If Rank(BKingLoc) > Rank(Square) Then ' Opp king in front
                If Abs(File(Square) - File(BKingLoc)) <= 1 Then ' and nearby file
                  ScaleFactor = KRPPKRP_SFactor(GetMax(r, Rank(Square))): GoTo lblEndScaleFactor
                End If
              End If
            End If
          End If
        End If
      ElseIf BPawnCnt = 2 And WPawnCnt = 1 Then ' black is strong side
        If BFrontMostPassedPawnRank = 0 Then
          Square = PieceSqList(BPAWN, 1) ' 1. pawn
          If RelativeRank(COL_BLACK, WKingLoc) > RelativeRank(COL_BLACK, Square) Then ' Opp king in front
            If Abs(File(Square) - File(WKingLoc)) <= 1 Then ' and nearby file
              r = RelativeRank(COL_BLACK, Square)
              Square = PieceSqList(BPAWN, 2) ' 2. pawn
              If RelativeRank(COL_BLACK, WKingLoc) > RelativeRank(COL_BLACK, Square) Then ' Opp king in front
                If Abs(File(Square) - File(WKingLoc)) <= 1 Then ' and nearby file
                  ScaleFactor = KRPPKRP_SFactor(GetMax(r, RelativeRank(COL_BLACK, Square))): GoTo lblEndScaleFactor
                End If
              End If
            End If
          End If
        End If
      End If
    End If
  
    '- zero or just one pawn makes it difficult to win
    If AllTotal.EG > 0 Then ' white stronger
      If NonPawnMat = 0 Then If WPawnCnt > BPawnCnt Then ScaleFactor = 96 ' A small advantage is typically decisive in a pure pawn endings
      Select Case WPawnCnt
      Case 0:
        If WNonPawnMaterial - BNonPawnMaterial <= ScoreBishop.MG Then
          If WNonPawnMaterial < ScoreRook.MG Then
            ScaleFactor = SCALE_FACTOR_DRAW
          Else
            If BNonPawnMaterial <= ScoreBishop.MG Then ScaleFactor = 4 Else ScaleFactor = 44
          End If
        End If
      Case 1: If WNonPawnMaterial - BNonPawnMaterial <= ScoreBishop.MG Then ScaleFactor = SCALE_FACTOR_ONEPAWN
      End Select
    ElseIf AllTotal.EG < 0 Then
      If NonPawnMat = 0 Then If BPawnCnt > WPawnCnt Then ScaleFactor = 96 ' A small advantage is typically decisive in a pure pawn endings
      Select Case BPawnCnt
      Case 0:
        If BNonPawnMaterial - WNonPawnMaterial <= ScoreBishop.MG Then
          If BNonPawnMaterial < ScoreRook.MG Then
            ScaleFactor = SCALE_FACTOR_DRAW
          Else
            If WNonPawnMaterial <= ScoreBishop.MG Then ScaleFactor = 4 Else ScaleFactor = 44
          End If
        End If
      Case 1: If BNonPawnMaterial - WNonPawnMaterial <= ScoreBishop.MG Then ScaleFactor = SCALE_FACTOR_ONEPAWN
      End Select
    End If
    
    '
    '- Endgame with opposite-colored bishops and no other pieces (ignoring pawns)
    '- is almost a draw, in case of KBP vs KB, it is even more a draw.
    If PieceCnt(WBISHOP) = 1 And PieceCnt(BBISHOP) = 1 And WBishopsOnWhiteSq = BBishopsOnBlackSq Then ' opposite-colored bishops
      If (WNonPawnMaterial = ScoreBishop.MG Or WNonPawnMaterial = ScoreBishop.MG + ScoreQueen.MG) And BNonPawnMaterial = WNonPawnMaterial Then
        If WPawnCnt + BPawnCnt > 1 Then ScaleFactor = 31 Else ScaleFactor = 9
      Else
        ' Endgame with opposite-colored bishops, but also other pieces. Still
        ' a bit drawish, but not as drawish as with only the two bishops.
        If PieceCnt(WQUEEN) + PieceCnt(BQUEEN) = 0 Then ScaleFactor = 46
      End If
    Else
      If WNonPawnMaterial + BNonPawnMaterial = 0 Then
        ' KPsK: K and two or more pawns vs K. There is just a single rule here: If all pawns
        ' are on the same rook file and are blocked by the defending king, it's a draw.
        If WPawnCnt >= 2 And BPawnCnt = 0 Then
          If File(PieceSqList(WPAWN, 1)) = 1 Or File(PieceSqList(WPAWN, 1)) = 8 Then
            r = 0

            For a = 1 To PieceSqListCnt(WPAWN)
              If File(PieceSqList(WPAWN, a)) <> File(PieceSqList(WPAWN, 1)) Then r = 1: Exit For
              If Abs(File(PieceSqList(WPAWN, a)) - File(BKingLoc)) > 1 Then r = 1: Exit For
              If Rank(PieceSqList(WPAWN, a)) >= Rank(BKingLoc) Then r = 1: Exit For
            Next

            If r = 0 Then ScaleFactor = 0 ' Draw
          End If
        ElseIf BPawnCnt >= 2 And WPawnCnt = 0 Then
          If File(PieceSqList(BPAWN, 1)) = 1 Or File(PieceSqList(BPAWN, 1)) = 8 Then
            r = 0

            For a = 1 To PieceSqListCnt(BPAWN)
              If File(PieceSqList(BPAWN, a)) <> File(PieceSqList(BPAWN, 1)) Then r = 1: Exit For
              If Abs(File(PieceSqList(BPAWN, a)) - File(WKingLoc)) > 1 Then r = 1: Exit For
              If Rank(PieceSqList(BPAWN, a)) <= Rank(WKingLoc) Then r = 1: Exit For
            Next

            If r = 0 Then ScaleFactor = 0 ' Draw
          End If
        End If
      End If
      If ScaleFactor = SCALE_FACTOR_NORMAL Or ScaleFactor = SCALE_FACTOR_ONEPAWN Then
        ' Endings where weaker side can stop one of the enemy's pawn are drawish.
        If AllTotal.EG > 0 Then ' White is strong side
              ScaleFactor = GetMin(40 + 7 * WPawnCnt, ScaleFactor)
        ElseIf AllTotal.EG < 0 Then ' Black is strong side
              ScaleFactor = GetMin(40 + 7 * BPawnCnt, ScaleFactor)
        End If
      End If
    End If
  End If
  
lblEndScaleFactor:
  
  
  '
  '--- Added all to eval score (SF based scaling:  Eval*100/SFPawnEndGameValue= 100 centipawns =1 pawn)
  '--- Example: Eval=240 => 1.00 pawn
  Eval = AllTotal.MG * GamePhase + AllTotal.EG * CLng(PHASE_MIDGAME - GamePhase) * ScaleFactor \ SCALE_FACTOR_NORMAL
  Eval = Eval \ PHASE_MIDGAME
  ' Initiative: Keep more pawns when attacking
  Bonus = (50 * (14 - (WPawnCnt + BPawnCnt))) \ 14
  If Eval > 0 Then
    Eval = GetMax(Eval - Bonus, Eval \ 2)
  ElseIf Eval < 0 Then
    Eval = GetMin(Eval + Bonus, Eval \ 2)
  End If
lblEndEval:
  If bEvalTrace Then
    bEvalTrace = False
    WriteTrace "---- EVAL TRACE : " & Now()
    WriteTrace PrintPos
    WriteTrace "Material: " & EvalSFTo100(MatEval)
    WriteTrace "Trades  : " & EvalSFTo100(TradeEval)
    WriteTrace "Position: " & ShowScoreDiff100(WPos, BPos) & "  => W" & ShowScore(WPos) & ", B" & ShowScore(BPos)
    WriteTrace "PawnStru: " & ShowScoreDiff100(WPawnStruct, BPawnStruct) & " => W" & ShowScore(WPawnStruct) & ", B" & ShowScore(BPawnStruct)
    WriteTrace "PassedPw: " & ShowScoreDiff100(WPassed, BPassed) & " => W" & ShowScore(WPassed) & ", B" & ShowScore(BPassed)
    WriteTrace "Mobility: " & ShowScoreDiff100(WMobility, BMobility) & " => W(" & ShowScore(WMobility) & ", B" & ShowScore(BMobility)
    WriteTrace "KSafety : " & ShowScoreDiff100(WKSafety, BKSafety) & " => W(" & ShowScore(WKSafety) & ", B" & ShowScore(BKSafety)
    WriteTrace "Threats : " & ShowScoreDiff100(WThreat, BThreat) & " => W(" & ShowScore(WThreat) & ", B" & ShowScore(BThreat)
    WriteTrace "Eval    : " & Eval & "  (" & EvalSFTo100(Eval) & "cp)"
    WriteTrace "-----------------"
    bTimeExit = True
  End If
  
  '------------------------------------------------
  '--- Step 9: Invert score for black to move   ---
  '------------------------------------------------
  If Not bWhiteToMove Then Eval = -Eval '--- Invert for black
  
  '-------------------------------------------------
  '--- Step 10: Add tempo value for side to move ---
  '-------------------------------------------------
  'Eval = Eval + TEMPO_BONUS ' Tempo for side to move
  Eval = Eval + (16 + NonPawnMaterial \ ScoreKnight.MG \ 2) ' use dynamic tempo, more during opening
  '
  If Eval = DrawContempt Then Eval = Eval + 1 ' if not a real draw then make a difference
  '
End Function

'---------------------------------
'-------- END OF EVAL ------------
'---------------------------------

Private Function Eval__EndOfEval_DUMMY()
  ' for faster navigation in source
End Function

Private Function IsMaterialDraw() As Boolean
  '( Protector logic )
  IsMaterialDraw = False
  If PieceCnt(WPAWN) + PieceCnt(BPAWN) = 0 Then ' no pawns
    '---  no heavies Q/R */
    If PieceCnt(WROOK) = 0 And PieceCnt(BROOK) = 0 And PieceCnt(WQUEEN) = 0 And PieceCnt(BQUEEN) = 0 Then
      If PieceCnt(BBISHOP) = 0 And PieceCnt(WBISHOP) = 0 Then
        '---  only knights */
        '---  it pretty safe to say this is a draw */
        If PieceCnt(WKNIGHT) < 3 And PieceCnt(BKNIGHT) < 3 Then IsMaterialDraw = True: Exit Function
      ElseIf PieceCnt(WKNIGHT) = 0 And PieceCnt(BKNIGHT) = 0 Then
        '---  only bishops */
        '---  not a draw if one side two other side zero
        '---  else its always a draw                     */
        If Abs(PieceCnt(WBISHOP) - PieceCnt(BBISHOP)) < 2 Then IsMaterialDraw = True: Exit Function
      ElseIf (PieceCnt(WKNIGHT) < 3 And PieceCnt(WBISHOP) = 0) Or (PieceCnt(WBISHOP) = 1 And PieceCnt(WKNIGHT) = 0) Then
        '---  we cant win, but can black? */
        If (PieceCnt(BKNIGHT) < 3 And PieceCnt(BBISHOP) = 0) Or (PieceCnt(BBISHOP) = 1 And PieceCnt(BKNIGHT) = 0) Then IsMaterialDraw = True: Exit Function '---  guess not */
      End If
    ElseIf PieceCnt(WQUEEN) = 0 And PieceCnt(BQUEEN) = 0 Then
      If PieceCnt(WROOK) = 1 And PieceCnt(BROOK) = 1 Then
        '---  rooks equal */
        '---  one minor difference max: a draw too usually */
        If (PieceCnt(WKNIGHT) + PieceCnt(WBISHOP)) < 2 And (PieceCnt(BKNIGHT) + PieceCnt(BBISHOP)) < 2 Then IsMaterialDraw = True: Exit Function
      ElseIf (PieceCnt(WROOK) = 1 And PieceCnt(BROOK) = 0) Then
        '---  one rook */
        '---  draw if no minors to support AND minors to defend  */
        If (PieceCnt(WKNIGHT) + PieceCnt(WBISHOP) = 0) And ((PieceCnt(BKNIGHT) + PieceCnt(BBISHOP) = 1) Or (PieceCnt(BKNIGHT) + PieceCnt(BBISHOP) = 2)) Then IsMaterialDraw = True: Exit Function
      ElseIf PieceCnt(BROOK) = 1 And PieceCnt(WROOK) = 0 Then
        '---  one rook */
        '---  draw if no minors to support AND minors to defend  */
        If (PieceCnt(BKNIGHT) + PieceCnt(BBISHOP) = 0) And ((PieceCnt(WKNIGHT) + PieceCnt(WBISHOP) = 1) Or (PieceCnt(WKNIGHT) + PieceCnt(WBISHOP) = 2)) Then IsMaterialDraw = True: Exit Function
      End If
    End If
  End If
End Function

Public Function AdvancedPawnPush(ByVal Piece As Long, ByVal Target As Long) As Boolean
  AdvancedPawnPush = False
  If Piece = WPAWN Then

    Select Case Rank(Target)
      Case 7, 8: AdvancedPawnPush = True
      Case 6:
        '--- if no enemy in front and no enemy pawns left or right
        If (Board(Target + SQ_UP) >= NO_PIECE Or Board(Target + SQ_UP) Mod 2 = WCOL) Then If Board(Target + SQ_UP_LEFT) <> BPAWN And Board(Target + SQ_UP_RIGHT) <> BPAWN Then AdvancedPawnPush = True
    End Select

  ElseIf Piece = BPAWN Then

    Select Case Rank(Target)
      Case 1, 2: AdvancedPawnPush = True
      Case 3:
        If (Board(Target + SQ_DOWN) >= NO_PIECE Or Board(Target + SQ_DOWN) Mod 2 = BCOL) Then If Board(Target + SQ_DOWN_LEFT) <> WPAWN And Board(Target + SQ_DOWN_RIGHT) <> WPAWN Then AdvancedPawnPush = True
    End Select

  End If
End Function

Public Function PieceSquareVal(ByVal Piece As Long, ByVal Square As Long) As Long
  '--- Piece value for a square
  PieceSquareVal = 0
  If bEndgame Then

    Select Case Piece
      Case NO_PIECE
      Case WPAWN
        PieceSquareVal = PsqtWP(Square).EG
      Case BPAWN
        PieceSquareVal = PsqtBP(Square).EG
      Case WKNIGHT
        PieceSquareVal = PsqtWN(Square).EG
      Case BKNIGHT
        PieceSquareVal = PsqtBN(Square).EG
      Case WBISHOP
        PieceSquareVal = PsqtWB(Square).EG
      Case BBISHOP
        PieceSquareVal = PsqtBB(Square).EG
      Case WROOK
        PieceSquareVal = PsqtWR(Square).EG
      Case BROOK
        PieceSquareVal = PsqtBR(Square).EG
      Case WQUEEN
        PieceSquareVal = PsqtWQ(Square).EG
      Case BQUEEN
        PieceSquareVal = PsqtBQ(Square).EG
      Case WKING
        PieceSquareVal = PsqtWK(Square).EG
      Case BKING
        PieceSquareVal = PsqtBK(Square).EG
    End Select

  Else

    Select Case Piece
      Case NO_PIECE
      Case WPAWN
        PieceSquareVal = PsqtWP(Square).MG
      Case BPAWN
        PieceSquareVal = PsqtBP(Square).MG
      Case WKNIGHT
        PieceSquareVal = PsqtWN(Square).MG
      Case BKNIGHT
        PieceSquareVal = PsqtBN(Square).MG
      Case WBISHOP
        PieceSquareVal = PsqtWB(Square).MG
      Case BBISHOP
        PieceSquareVal = PsqtBB(Square).MG
      Case WROOK
        PieceSquareVal = PsqtWR(Square).MG
      Case BROOK
        PieceSquareVal = PsqtBR(Square).MG
      Case WQUEEN
        PieceSquareVal = PsqtWQ(Square).MG
      Case BQUEEN
        PieceSquareVal = PsqtBQ(Square).MG
      Case WKING
        PieceSquareVal = PsqtWK(Square).MG
      Case BKING
        PieceSquareVal = PsqtBK(Square).MG
    End Select

  End If
End Function

Public Sub FillPieceSquareVal()
  Dim Piece As Long, Target As Long

  For Piece = 1 To 16
    For Target = SQ_A1 To SQ_H8
      bEndgame = False
      PsqVal(0, Piece, Target) = PieceSquareVal(Piece, Target)
      bEndgame = True
      PsqVal(1, Piece, Target) = PieceSquareVal(Piece, Target)
    Next
  Next

End Sub

Private Function AttackByCol(Col As Long, Square As Long) As Long
  If Col = COL_WHITE Then AttackByCol = WAttack(Square) Else AttackByCol = BAttack(Square)
End Function

Public Sub AddPawnThreat(Score As TScore, _
                         ByVal HangCol As enumColor, _
                         ByVal PieceType As enumPieceType, _
                         ByVal Square As Long)
  'SF6:  const Score ThreatBySafePawn[PIECE_TYPE_NB] = {
  '         S(0, 0), S(0, 0), S(107, 138), S(84, 122), S(114, 203), S(121, 217)
  '      const Score ThreatenedByHangingPawn = S(71, 61);
  '--- attack by black pawn?
  If HangCol = COL_WHITE Then
    If Board(Square + SQ_UP_LEFT) = BPAWN Then
      If Board(Square + SQ_UP_LEFT + SQ_UP_LEFT) = BPAWN Or Board(Square + SQ_UP_LEFT + SQ_UP_RIGHT) = BPAWN Then
        AddScore Score, ThreatBySafePawn(PieceType)
      Else
        AddScore Score, ThreatenedByHangingPawn
      End If
    ElseIf Board(Square + SQ_UP_RIGHT) = BPAWN Then
      If Board(Square + SQ_UP_RIGHT + SQ_UP_LEFT) = BPAWN Or Board(Square + SQ_UP_RIGHT + SQ_UP_RIGHT) = BPAWN Then
        AddScore Score, ThreatBySafePawn(PieceType)
      Else
        AddScore Score, ThreatenedByHangingPawn
      End If
    End If
  Else ' attack by white pawn?
    If Board(Square + SQ_DOWN_LEFT) = WPAWN Then
      If Board(Square + SQ_DOWN_LEFT + SQ_DOWN_LEFT) = WPAWN Or Board(Square + SQ_DOWN_LEFT + SQ_DOWN_RIGHT) = WPAWN Then
        AddScore Score, ThreatBySafePawn(PieceType)
      Else
        AddScore Score, ThreatenedByHangingPawn
      End If
    ElseIf Board(Square + SQ_DOWN_RIGHT) = WPAWN Then
      If Board(Square + SQ_DOWN_RIGHT + SQ_DOWN_LEFT) = WPAWN Or Board(Square + SQ_DOWN_RIGHT + SQ_DOWN_RIGHT) = WPAWN Then
        AddScore Score, ThreatBySafePawn(PieceType)
      Else
        AddScore Score, ThreatenedByHangingPawn
      End If
    End If
  End If
End Sub

Public Sub AddThreat(ByVal HangCol As enumColor, _
                     ByVal HangPieceType As enumPieceType, _
                     ByVal AttackerPieceType As enumPieceType, _
                     ByVal AttackerSquare As Long, _
                     ByVal AttackedSquare As Long)
  ' Add threat to threat list. calculate score later when full attack array data is available
  ThreatCnt = ThreatCnt + 1

  With ThreatList(ThreatCnt)
    .HangCol = HangCol
    .HangPieceType = HangPieceType
    .AttackerPieceType = AttackerPieceType
    .AttackerSquare = AttackerSquare
    .AttackedSquare = AttackedSquare
  End With

End Sub

Public Sub CalcThreats()
  If ThreatCnt = 0 Then Exit Sub
  Dim i As Long, Defended As Boolean, StronglyProtected As Boolean, Weak As Boolean
  Dim UsAttackCnt As Long, ThemAttackCnt As Long, RelRank As Long, PawnProtected As Boolean, Score As TScore

  For i = 1 To ThreatCnt

    With ThreatList(i)
      '
      ' Add a bonus according to the kind of attacking pieces
      '
      Score = ZeroScore
      If .HangCol = COL_WHITE Then ' view from attacker side = us, attacked = them
        UsAttackCnt = AttackBitCnt(BAttack(.AttackedSquare)): ThemAttackCnt = AttackBitCnt(WAttack(.AttackedSquare))
        PawnProtected = CBool(WAttack(.AttackedSquare) And PAttackBit): RelRank = Rank(.AttackedSquare)
      Else ' Black
        UsAttackCnt = AttackBitCnt(WAttack(.AttackedSquare)): ThemAttackCnt = AttackBitCnt(BAttack(.AttackedSquare))
        PawnProtected = CBool(BAttack(.AttackedSquare) And PAttackBit): RelRank = 9 - Rank(.AttackedSquare)
      End If
      '
      ' StronglyProtected: by pawn or by more defenders then attackers
      StronglyProtected = PawnProtected Or (ThemAttackCnt > UsAttackCnt)
      ' Non-pawn enemies strongly defended
      Defended = .HangPieceType <> PT_PAWN And StronglyProtected
      ' Enemies not strongly defended and under our attack
      Weak = Not StronglyProtected
      If Defended Or Weak Then
        If .AttackerPieceType = PT_BISHOP Or .AttackerPieceType = PT_KNIGHT Then
          AddScore Score, ThreatByMinor(.HangPieceType)
          If .HangPieceType <> PT_PAWN Then
            AddScoreVal Score, ThreatByRank.MG * RelRank, ThreatByRank.EG * RelRank
          End If
        End If
        If Weak Then If ThemAttackCnt = 0 Then AddScore Score, Hanging ' hanging
        If .HangPieceType <> PT_PAWN Then ' Overload: attacked and defended only once
          If ThemAttackCnt = 1 Then AddScore Score, Overload
        End If
      End If
      If (.HangPieceType = PT_QUEEN Or Weak) And .AttackerPieceType = PT_ROOK Then
        AddScore Score, ThreatByRook(.HangPieceType)
        If .HangPieceType <> PT_PAWN Then
          AddScoreVal Score, ThreatByRank.MG * RelRank, ThreatByRank.EG * RelRank
        End If
      End If
      If Score.MG <> 0 Or Score.EG <> 0 Then If .HangCol = COL_WHITE Then AddScore BThreat, Score Else AddScore WThreat, Score
    End With
   
lblNext:
  Next

End Sub

Public Sub AddWKingAttackers(ByVal AttackBit As Long)
 If AttackBit And PLAttackBit Then AddWKingAttack PT_PAWN
 If AttackBit And PRAttackBit Then AddWKingAttack PT_PAWN
 If AttackBit And N1AttackBit Then AddWKingAttack PT_KNIGHT
 If AttackBit And N2AttackBit Then AddWKingAttack PT_KNIGHT
 If AttackBit And B1AttackBit Then AddWKingAttack PT_BISHOP
 If AttackBit And B2AttackBit Then AddWKingAttack PT_BISHOP
 If AttackBit And BXrayAttackBit Then If Not (AttackBit And (B1AttackBit Or B2AttackBit)) Then WKingAttackersCount = WKingAttackersCount + 1
 If AttackBit And (R1AttackBit Or R1XrayAttackBit) Then AddWKingAttack PT_ROOK
 If AttackBit And (R2AttackBit Or R2XrayAttackBit) Then AddWKingAttack PT_ROOK
 If AttackBit And (QAttackBit Or QXrayAttackBit) Then AddWKingAttack PT_QUEEN
End Sub

Public Sub AddBKingAttackers(ByVal AttackBit As Long)
 If AttackBit And PLAttackBit Then AddBKingAttack PT_PAWN
 If AttackBit And PRAttackBit Then AddBKingAttack PT_PAWN
 If AttackBit And N1AttackBit Then AddBKingAttack PT_KNIGHT
 If AttackBit And N2AttackBit Then AddBKingAttack PT_KNIGHT
 If AttackBit And B1AttackBit Then AddBKingAttack PT_BISHOP
 If AttackBit And B2AttackBit Then AddBKingAttack PT_BISHOP
 If AttackBit And BXrayAttackBit Then If Not (AttackBit And (B1AttackBit Or B2AttackBit)) Then BKingAttackersCount = BKingAttackersCount + 1
 If AttackBit And (R1AttackBit Or R1XrayAttackBit) Then AddBKingAttack PT_ROOK
 If AttackBit And (R2AttackBit Or R2XrayAttackBit) Then AddBKingAttack PT_ROOK
 If AttackBit And (QAttackBit Or QXrayAttackBit) Then AddBKingAttack PT_QUEEN
End Sub

Public Sub AddWKingAttack(PT As enumPieceType)
  WKingAttackersCount = WKingAttackersCount + 1
  WKingAttackersWeight = WKingAttackersWeight + KingAttackWeights(PT)
End Sub

Public Sub AddBKingAttack(PT As enumPieceType)
  BKingAttackersCount = BKingAttackersCount + 1
  BKingAttackersWeight = BKingAttackersWeight + KingAttackWeights(PT)
End Sub

Public Function InitConnectedPawns()
  ' SF6
  Dim Seed(8) As Long, Opposed As Long, Phalanx As Long, Support As Long, r As Long, v As Long, x As Long
  ReadLngArr Seed(), 0, 0, 13, 24, 18, 76, 100, 175, 330

  For Opposed = 0 To 1
    For Phalanx = 0 To 1
      For Support = 0 To 2
        For r = 2 To 7
          If Phalanx > 0 Then x = (Seed(r + 1) - Seed(r)) / 2 Else x = 0
          v = 17 * Support
          v = v + Seed(r)
          If Phalanx > 0 Then v = v + (Seed(r + 1) - Seed(r)) \ 2
          If Opposed > 0 Then v = v / 2 ' >>  operator for opposed in VB: /2
          ConnectedBonus(Opposed, Phalanx, Support, r).MG = v
          ConnectedBonus(Opposed, Phalanx, Support, r).EG = v * ((r - 1) - 2) \ 4 ' rank r ist zero based in C, so (r-1)
        Next
      Next
    Next
  Next

End Function

Public Sub InitImbalance()  ' SF6
  ' // pair pawn knight bishop rook queen  OUR PIECES
  ReadIntArr2 QuadraticOurs(), 0, 1667 ' Bishop pair
  ReadIntArr2 QuadraticOurs(), PT_PAWN, 40, 0  ' Pawn
  ReadIntArr2 QuadraticOurs(), PT_KNIGHT, 32, 255, -3        ' Knight
  ReadIntArr2 QuadraticOurs(), PT_BISHOP, 0, 104, 4, 0       ' Bishop
  ReadIntArr2 QuadraticOurs(), PT_ROOK, -26, -2, 47, 105, -149          ' Rook
  ReadIntArr2 QuadraticOurs(), PT_QUEEN, -185, 24, 122, 137, -134, 0    ' Queen
  ' // pair pawn knight bishop rook queen  THEIR PIECES
  ReadIntArr2 QuadraticTheirs(), 0, 0 ' Bishop pair
  ReadIntArr2 QuadraticTheirs(), PT_PAWN, 36, 0     ' Pawn
  ReadIntArr2 QuadraticTheirs(), PT_KNIGHT, 9, 63, 0              ' Knight
  ReadIntArr2 QuadraticTheirs(), PT_BISHOP, 59, 65, 42, 0         ' Bishop
  ReadIntArr2 QuadraticTheirs(), PT_ROOK, 46, 39, 24, -24, 0            ' Rook
  ReadIntArr2 QuadraticTheirs(), PT_QUEEN, 101, 100, -37, 141, 268, 0   ' Queen
  ' // PawnSet[pawn count] contains a bonus/malus indexed by number of pawns
  ReadIntArr PawnSet(), 24, -32, 107, -51, 117, -9, -126, -21, 31
End Sub

Public Function Imbalance() As Long ' SF
  Dim v As Long, Key As Long
  Key = CalcMaterialKey()
  Imbalance = ProbeMaterialHash(Key)
  If Imbalance <> UNKNOWN_SCORE Then Exit Function
  ImbPieceCount(COL_WHITE, 0) = Abs(PieceCnt(WBISHOP) > 1)  ' index 0 used for bishop pair
  ImbPieceCount(COL_BLACK, 0) = Abs(PieceCnt(BBISHOP) > 1)  ' index 0 used for bishop pair
  ImbPieceCount(COL_WHITE, PT_PAWN) = PieceCnt(WPAWN)
  ImbPieceCount(COL_BLACK, PT_PAWN) = PieceCnt(BPAWN)
  ImbPieceCount(COL_WHITE, PT_KNIGHT) = PieceCnt(WKNIGHT)
  ImbPieceCount(COL_BLACK, PT_KNIGHT) = PieceCnt(BKNIGHT)
  ImbPieceCount(COL_WHITE, PT_BISHOP) = PieceCnt(WBISHOP)
  ImbPieceCount(COL_BLACK, PT_BISHOP) = PieceCnt(BBISHOP)
  ImbPieceCount(COL_WHITE, PT_ROOK) = PieceCnt(WROOK)
  ImbPieceCount(COL_BLACK, PT_ROOK) = PieceCnt(BROOK)
  ImbPieceCount(COL_WHITE, PT_QUEEN) = PieceCnt(WQUEEN)
  ImbPieceCount(COL_BLACK, PT_QUEEN) = PieceCnt(BQUEEN)
  v = (ColImbalance(COL_WHITE) - ColImbalance(COL_BLACK)) \ 16
  'If Imbalance <> UNKNOWN_SCORE And Imbalance <> v Then MsgBox "Diff"
  Imbalance = v
  SaveMaterialHash Key, Imbalance
End Function

Public Function ColImbalance(ByVal Col As enumColor) As Long
  Dim Bonus As Long, pt1 As Long, pt2 As Long, Us As Long, Them As Long, v As Long
  If Col = COL_WHITE Then
    Us = COL_WHITE: Them = COL_BLACK: Bonus = PawnSet(PieceCnt(WPAWN))
    If PieceCnt(WQUEEN) = 1 Then If PieceCnt(BQUEEN) = 0 Then Bonus = Bonus + QueenMinorsImbalance(PieceCnt(BKNIGHT) + PieceCnt(BBISHOP))
  Else
    Us = COL_BLACK: Them = COL_WHITE: Bonus = PawnSet(PieceCnt(BPAWN))
    If PieceCnt(BQUEEN) = 1 Then If PieceCnt(WQUEEN) = 0 Then Bonus = Bonus + QueenMinorsImbalance(PieceCnt(WKNIGHT) + PieceCnt(WBISHOP))
  End If

  For pt1 = 0 To PT_QUEEN
    If ImbPieceCount(Us, pt1) > 0 Then
      v = 0

      For pt2 = 0 To pt1
        v = v + QuadraticOurs(pt1, pt2) * ImbPieceCount(Us, pt2) + QuadraticTheirs(pt1, pt2) * ImbPieceCount(Them, pt2)
      Next pt2

      Bonus = Bonus + ImbPieceCount(Us, pt1) * v
    End If
  Next pt1

  ColImbalance = Bonus
End Function

Public Sub AddScore(ScoreTotal As TScore, ScoreAdd As TScore)
  ScoreTotal.MG = ScoreTotal.MG + ScoreAdd.MG: ScoreTotal.EG = ScoreTotal.EG + ScoreAdd.EG
End Sub

Public Sub AddScoreWithFactor(ScoreTotal As TScore, ScoreAdd As TScore, Factor As Long)
  ScoreTotal.MG = ScoreTotal.MG + ScoreAdd.MG * Factor: ScoreTotal.EG = ScoreTotal.EG + ScoreAdd.EG * Factor
End Sub

Public Sub AddScore100(ScoreTotal As TScore, ScoreAdd As TScore)
  ' Score 100 centipawns based: scale to SF pawn value
  ScoreTotal.MG = ScoreTotal.MG + (ScoreAdd.MG * ScorePawn.EG) \ 100&: ScoreTotal.EG = ScoreTotal.EG + (ScoreAdd.EG * ScorePawn.EG) \ 100&
End Sub

Public Sub AddScoreVal(ScoreTotal As TScore, ByVal MGScore As Long, ByVal EGSCore As Long)
  ScoreTotal.MG = ScoreTotal.MG + MGScore: ScoreTotal.EG = ScoreTotal.EG + EGSCore
End Sub

Public Sub SetScoreVal(ScoreSet As TScore, ByVal MGScore As Long, ByVal EGSCore As Long)
  ScoreSet.MG = MGScore: ScoreSet.EG = EGSCore
End Sub

Public Function EvalSFTo100(ByVal Eval As Long) As Long
  If Abs(Eval) < MATE_IN_MAX_PLY Then EvalSFTo100 = (Eval * 100&) / CLng(ScorePawn.EG) Else EvalSFTo100 = Eval
End Function

Public Function Eval100ToSF(ByVal Eval As Long) As Long
  Eval100ToSF = (Eval * CLng(ScorePawn.EG)) / 100&
End Function

Public Sub MinusScore(ScoreTotal As TScore, ScoreMinus As TScore)
  ScoreTotal.MG = ScoreTotal.MG - ScoreMinus.MG: ScoreTotal.EG = ScoreTotal.EG - ScoreMinus.EG
End Sub

Public Function ScaleScore(Score As TScore) As Long
  ' Calculate score for game phase
  ScaleScore = Score.MG * GamePhase + Score.EG * CLng(PHASE_MIDGAME - GamePhase) '  * SF6 / SCALE_FACTOR_NORMAL
  ScaleScore = ScaleScore \ PHASE_MIDGAME
End Function

Public Function ScaleScore100(Score As TScore, ByVal ScaleVal As Long) As TScore
  ScaleScore100.MG = (Score.MG * ScaleVal) \ 100&: ScaleScore100.EG = (Score.EG * ScaleVal) \ 100&
End Function

Public Function ShowScore(Score As TScore) As String
  ' show MG, EG Score as text
  ShowScore = "(" & CStr(Score.MG) & "," & CStr(Score.EG) & ")=" & ScaleScore(Score)
End Function

Public Function ShowScoreDiff100(Score1 As TScore, Score2 As TScore) As String
  ' show MG, EG Score as text
  Dim Diff As TScore
  Diff.MG = Score1.MG - Score2.MG: Diff.EG = Score1.EG - Score2.EG
  ShowScoreDiff100 = ShowScore(Diff)
End Function

Public Function PieceSQ(ByVal Side As enumColor, _
                        ByVal SearchPieceType As enumPieceType) As Long
  Dim a As Long, p As Long

  For a = 1 To NumPieces
    p = Board(Pieces(a)): If PieceType(p) = SearchPieceType And PieceColor(p) = Side Then PieceSQ = Pieces(a): Exit Function
  Next

End Function

Public Function Eval_KRKP() As Long
  Dim WKSq          As Long, BKSq As Long, RookSq As Long, PawnSq As Long, StrongSide As enumColor, WeakSide As enumColor
  Dim StrongKingLoc As Long, WeakKingLoc As Long, QueeningSq As Long, Result As Long, SideToMove As enumColor
  If WMaterial > BMaterial Then
    StrongSide = COL_WHITE: WeakSide = COL_BLACK: StrongKingLoc = WKingLoc: WeakKingLoc = BKingLoc
  Else
    StrongSide = COL_BLACK: WeakSide = COL_WHITE: StrongKingLoc = BKingLoc: WeakKingLoc = WKingLoc
  End If
  If bWhiteToMove Then SideToMove = COL_WHITE Else SideToMove = COL_BLACK
  WKSq = RelativeSq(StrongSide, StrongKingLoc)
  BKSq = RelativeSq(StrongSide, WeakKingLoc)
  RookSq = RelativeSq(StrongSide, PieceSQ(StrongSide, PT_ROOK))
  PawnSq = RelativeSq(WeakSide, PieceSQ(WeakSide, PT_PAWN))
  QueeningSq = SQ_A1 + File(PawnSq) - 1 + 7 * SQ_UP
  '-- If the stronger side's king is in front of the pawn, it's a win
  If WKSq < PawnSq And File(WKSq) = File(PawnSq) Then
    Result = ScoreRook.EG - MaxDistance(WKSq, PawnSq)
    '-- If the weaker side's king is too far from the pawn and the rook, it's a win.
  ElseIf MaxDistance(BKSq, PawnSq) >= (3 + Abs(SideToMove = WeakSide)) And MaxDistance(BKSq, RookSq) >= 3 Then
    Result = ScoreRook.EG - MaxDistance(WKSq, PawnSq)
    '-- If the pawn is far advanced and supported by the defending king, the position is drawish
  ElseIf Rank(BKSq) <= 3 And MaxDistance(BKSq, PawnSq) = 1 And Rank(WKSq) >= 4 And MaxDistance(WKSq, PawnSq) > (2 + Abs(SideToMove = StrongSide)) Then
    Result = 80 - 8 * MaxDistance(WKSq, PawnSq)
  Else
    Result = 200 - 8 * (MaxDistance(WKSq, PawnSq + SQ_DOWN) - MaxDistance(BKSq, PawnSq + SQ_DOWN) - MaxDistance(PawnSq, QueeningSq))
  End If
  If StrongSide = SideToMove Then Eval_KRKP = Result Else Eval_KRKP = -Result
  If Not bWhiteToMove Then Eval_KRKP = -Eval_KRKP
End Function

Public Function Eval_KQKP() As Long
  ' KQ vs KP. In general, this is a win for the stronger side, but there are a
  ' few important exceptions. A pawn on 7th rank and on the A,C,F or H files
  ' with a king positioned next to it can be a draw, so in that case, we only
  ' use the distance between the kings.
  Dim WinnerKSq As Long, LoserKSq As Long, PawnSq As Long, StrongSide As enumColor, WeakSide As enumColor
  Dim Result    As Long, SideToMove As enumColor
  If WMaterial > BMaterial Then
    StrongSide = COL_WHITE: WeakSide = COL_BLACK: WinnerKSq = WKingLoc: LoserKSq = BKingLoc
  Else
    StrongSide = COL_BLACK: WeakSide = COL_WHITE: WinnerKSq = BKingLoc: LoserKSq = WKingLoc
  End If
  PawnSq = PieceSQ(WeakSide, PT_PAWN)
  If bWhiteToMove Then SideToMove = COL_WHITE Else SideToMove = COL_BLACK
  Result = PushClose(MaxDistance(WinnerKSq, LoserKSq))
  If RelativeRank(WeakSide, PawnSq) <> 7 Or MaxDistance(LoserKSq, PawnSq) <> 1 Then
    Result = Result + ScoreQueen.EG - ScorePawn.EG
  Else

    Select Case File(PawnSq) ' For File A,C,F,H
      Case 2, 4, 5, 7: Result = Result + ScoreQueen.EG - ScorePawn.EG
    End Select

  End If
  If StrongSide = SideToMove Then Eval_KQKP = Result Else Eval_KQKP = -Result
  If Not bWhiteToMove Then Eval_KQKP = -Eval_KQKP
End Function
 
Public Function Eval_KQKR() As Long
  Dim WinnerKSq As Long, LoserKSq As Long, StrongSide As enumColor, WeakSide As enumColor
  Dim Result    As Long, SideToMove As enumColor
  If WMaterial > BMaterial Then
    StrongSide = COL_WHITE: WeakSide = COL_BLACK: WinnerKSq = WKingLoc: LoserKSq = BKingLoc
  Else
    StrongSide = COL_BLACK: WeakSide = COL_WHITE: WinnerKSq = BKingLoc: LoserKSq = WKingLoc
  End If
  If bWhiteToMove Then SideToMove = COL_WHITE Else SideToMove = COL_BLACK
  Result = ScoreQueen.EG - ScoreRook.EG + PushToEdges(LoserKSq) + PushClose(MaxDistance(WinnerKSq, LoserKSq))
  If StrongSide = SideToMove Then Eval_KQKR = Result Else Eval_KQKR = -Result
  If Not bWhiteToMove Then Eval_KQKR = -Eval_KQKR
End Function
 
Private Function WKingShelterStorm(ShelterKingLoc As Long) As Long
  Dim Center As Long, k As Long, r As Long, RelFile As Long, Safety As Long, RankUs As Long, RankThem As Long, RankNum As Long
  Safety = 258
  ' Opp pawn rank A/H protects king
  If File(WKingLoc) = 1 Or File(WKingLoc) = 8 Then
    If Rank(WKingLoc) <= 2 Then If Board(WKingLoc + SQ_UP) = BPAWN Then Safety = 350
  End If
  
  '--- Pawn shelter
  Center = GetMax(2, GetMin(7, File(ShelterKingLoc))): RankNum = Rank(ShelterKingLoc) ' FIle A=>B, File H=>G

  For k = Center - 1 To Center + 1
    ' Pawn shelter/storm
    RankUs = 1
    If WPawns(k) > 0 Then If PawnsWMin(k) >= RankNum Then RankUs = PawnsWMin(k)
    RankThem = 1
    If BPawns(k) > 0 Then If PawnsBMin(k) >= RankNum Then RankThem = PawnsBMin(k)
    If RankThem = RankNum + 1 And k = File(ShelterKingLoc) Then
      r = 1 ' BlockedByKing
    ElseIf RankUs = 1 Then
      r = 2 ' NoFriendlyPawn
    ElseIf RankThem = RankUs + 1 Then
      r = 3 ' BlockedByPawn
    Else
      r = 4 ' Unblocked
    End If
    RelFile = GetMin(k, 9 - k)
    Safety = Safety - (ShelterWeakness(RelFile, RankUs) + StormDanger(r, RelFile, RankThem))
  Next

  If Center >= 6 Then
    If Board(SQ_H3) = BPAWN Then
      If Board(SQ_H2) = WPAWN Then If Board(SQ_G3) = WPAWN Then If Board(SQ_F2) = WPAWN Then Safety = Safety + 300
    End If
    If Board(SQ_F3) = BPAWN Then
      If Board(SQ_H2) = WPAWN Then If Board(SQ_G3) = WPAWN Then If Board(SQ_F2) = WPAWN Then Safety = Safety + 300
    End If
  ElseIf Center <= 3 Then
    If Board(SQ_A3) = BPAWN Then
      If Board(SQ_A2) = WPAWN Then If Board(SQ_B3) = WPAWN Then If Board(SQ_C2) = WPAWN Then Safety = Safety + 300
    End If
    If Board(SQ_C3) = BPAWN Then
      If Board(SQ_A2) = WPAWN Then If Board(SQ_B3) = WPAWN Then If Board(SQ_C2) = WPAWN Then Safety = Safety + 300
    End If
  End If

  WKingShelterStorm = Safety
End Function

Private Function BKingShelterStorm(ByVal ShelterKingLoc As Long) As Long
  Dim Center As Long, k As Long, r As Long, RelFile As Long, Safety As Long, RankUs As Long, RankThem As Long, RankNum As Long
  Safety = 258
  ' Opp pawn rank A/H protects king
  If File(BKingLoc) = 1 Or File(BKingLoc) = 8 Then
    If Rank(BKingLoc) >= 7 Then If Board(BKingLoc + SQ_DOWN) = WPAWN Then Safety = 350
  End If  '--- Pawn shelter
  Center = GetMax(2, GetMin(7, File(ShelterKingLoc))): RankNum = 9 - Rank(ShelterKingLoc) ' FIle A=>B, File H=>G

  For k = Center - 1 To Center + 1
    ' Pawn shelter/storm
    RankUs = 1
    If BPawns(k) > 0 Then If 9 - PawnsBMax(k) >= RankNum Then RankUs = (9 - PawnsBMax(k))
    RankThem = 1
    If WPawns(k) > 0 Then If 9 - PawnsWMax(k) >= RankNum Then RankThem = (9 - PawnsWMax(k))
    If RankThem = RankNum + 1 And k = File(ShelterKingLoc) Then
      r = 1 ' BlockedByKing
    ElseIf RankUs = 1 Then
      r = 2 ' NoFriendlyPawn
    ElseIf RankThem = RankUs + 1 Then
      r = 3 ' BlockedByPawn
    Else
      r = 4 ' Unblocked
    End If
    RelFile = GetMin(k, 9 - k)
    Safety = Safety - (ShelterWeakness(RelFile, RankUs) + StormDanger(r, RelFile, RankThem))
  Next
  If Center >= 6 Then
    If Board(SQ_H6) = WPAWN Then
      If Board(SQ_H7) = BPAWN Then If Board(SQ_G6) = BPAWN Then If Board(SQ_F7) = BPAWN Then Safety = Safety + 250
    End If
    If Board(SQ_F6) = WPAWN Then
      If Board(SQ_H7) = BPAWN Then If Board(SQ_G6) = BPAWN Then If Board(SQ_F7) = BPAWN Then Safety = Safety + 150
    End If
  ElseIf Center <= 3 Then
    If Board(SQ_A6) = WPAWN Then
      If Board(SQ_A7) = BPAWN Then If Board(SQ_B6) = BPAWN Then If Board(SQ_C7) = BPAWN Then Safety = Safety + 250
    End If
    If Board(SQ_C6) = WPAWN Then
      If Board(SQ_A7) = BPAWN Then If Board(SQ_B6) = BPAWN Then If Board(SQ_C7) = BPAWN Then Safety = Safety + 150
    End If
  End If

  BKingShelterStorm = Safety
End Function

Private Sub GetKingFlankFiles(ByVal KingLoc As Long, FileFrom As Long, FileTo As Long)

  Select Case File(KingLoc)
    Case 1 To 3: FileFrom = FILE_A: FileTo = FILE_D  ' File A-C> A-D
    Case 4 To 5: FileFrom = FILE_C: FileTo = FILE_F  ' File D-E> C-F
    Case 6 To 8: FileFrom = FILE_E: FileTo = FILE_H  ' File F-H> E-H
  End Select

End Sub

Public Function PinnedPieceDir(ByVal PinnedLoc As Long, ByVal MoveTarget As Long, PieceCol As enumColor) As Long
 '  check if a piece is pinned to king and returns direction offset from piece to king, if not pinned =  0
 PinnedPieceDir = 0
 If PieceCol = COL_WHITE Then
  PinnedPieceDir = WPinnedPieceDir(PinnedLoc)
  If PinnedPieceDir <> 0 Then
    If SameXRay(MoveTarget, WKingLoc) Then PinnedPieceDir = 0 ' move in pinned direction Ok
  End If
 ElseIf PieceCol = COL_BLACK Then
   PinnedPieceDir = BPinnedPieceDir(PinnedLoc)
  If PinnedPieceDir <> 0 Then
    If SameXRay(MoveTarget, BKingLoc) Then PinnedPieceDir = 0 ' move in pinned direction Ok
  End If
 End If
'If PinnedPieceDir <> 0 Then Stop
End Function

Public Function WPinnedPieceDir(ByVal PinnedLoc As Long) As Long
  '-- check if a piece is pinned to king and returns direction offset from piece to king, if not pinned =  0
  Dim k As Long, sq As Long, Offset As Long, Piece As Long
  WPinnedPieceDir = 0
  If PinnedLoc = WKingLoc Then Exit Function
  Offset = DirOffset(PinnedLoc, WKingLoc)  ' Find direction to king
  If Offset = 0 Then Exit Function
  
  ' no other piece between piece and own king?
  sq = PinnedLoc
  For k = 1 To 7
    sq = sq + Offset: If sq = WKingLoc Then Exit For ' pinned possible
    Piece = Board(sq) ': If Piece = FRAME Then Exit For ' should not happen
    If Piece < NO_PIECE Then Exit Function ' other piece found > not pinned
  Next k

  ' check other direction for attacker
  sq = PinnedLoc
  For k = 1 To 7
    sq = sq - Offset
    Piece = Board(sq): If Piece = FRAME Then Exit For
    If Piece < NO_PIECE Then
      Select Case Piece
      Case BQUEEN:
        WPinnedPieceDir = Offset: Exit Function ' pinned by queen
      Case BROOK:
        If Abs(Offset) = 10 Or Abs(Offset) = 1 Then WPinnedPieceDir = Offset: Exit Function ' pinned by rook
      Case BBISHOP:
        If Abs(Offset) = 9 Or Abs(Offset) = 11 Then WPinnedPieceDir = Offset: Exit Function ' pinned by bishop
      End Select
      Exit Function ' other piece found
    End If
  Next k
  ' not pinned here
End Function

Public Function BPinnedPieceDir(ByVal PinnedLoc As Long) As Long
  '-- check if a piece is pinned to king and returns direction offset from piece to king, if not pinned =  0
  Dim k As Long, sq As Long, Offset As Long, Piece As Long
  BPinnedPieceDir = 0
  If PinnedLoc = BKingLoc Then Exit Function
  Offset = DirOffset(PinnedLoc, BKingLoc)  ' Find direction to king
  If Offset = 0 Then Exit Function

  ' no other piece between piece and own king?
  sq = PinnedLoc
  For k = 1 To 7
    sq = sq + Offset: If sq = BKingLoc Then Exit For ' pinned possible
    Piece = Board(sq) ': If Piece = FRAME Then Exit For ' should not happen
    If Piece < NO_PIECE Then Exit Function ' other piece found > not pinned
  Next k

  ' check other direction for attacker
  sq = PinnedLoc
  For k = 1 To 7
    sq = sq - Offset
    Piece = Board(sq): If Piece = FRAME Then Exit For
    If Piece < NO_PIECE Then
      Select Case Piece
      Case WQUEEN:
        BPinnedPieceDir = Offset: Exit Function ' pinned by queen
      Case WROOK:
        If Abs(Offset) = 10 Or Abs(Offset) = 1 Then BPinnedPieceDir = Offset: Exit Function ' pinned by rook
      Case WBISHOP:
        If Abs(Offset) = 9 Or Abs(Offset) = 11 Then BPinnedPieceDir = Offset: Exit Function ' pinned by bishop
      End Select
      Exit Function ' other piece found
    End If
  Next k
  ' not pinned here
End Function

'Public Function PinnedPieceW(ByVal PinnedLoc As Long, ByVal Direction As Long) As Boolean
'  ' white pieces it threatend by pinned pieces and slider attack?
'  Dim k As Long, sq As Long, Offset As Long, AttackBit As Long, Piece As Long
'  PinnedPieceW = False
'  If Direction < 4 Then ' Queen or rook orthogonal
'    If Not CBool(BAttack(PinnedLoc) And QRAttackBit) Then Exit Function
'    AttackBit = QRAttackBit
'  Else ' Queen or bishop diagonal
'    If Not CBool(BAttack(PinnedLoc) And QBAttackBit) Then Exit Function
'    AttackBit = QBAttackBit
'  End If
'  Offset = QueenOffsets(Direction)
'
'  For k = 1 To 8
'    sq = PinnedLoc + Offset * k: Piece = Board(sq)
'    If Piece = FRAME Then Exit For
'    If Piece < NO_PIECE Then
'      If Piece = BQUEEN Then PinnedPieceW = True: Exit Function
'      If Piece = BROOK Then If Direction < 4 Then PinnedPieceW = True: Exit Function
'      If Piece = BBISHOP Then If Direction >= 4 Then PinnedPieceW = True: Exit Function
'      Exit For
'    Else
'      If Not (CBool(BAttack(sq) And AttackBit)) Then Exit For
'    End If
'  Next k
'
'End Function
'
'Public Function PinnedPieceB(ByVal PinnedLoc As Long, ByVal Direction As Long) As Boolean
'  ' black pieces it threatend by pinned pieces and slider attack?
'  Dim k As Long, sq As Long, Offset As Long, AttackBit As Long, Piece As Long
'  PinnedPieceB = False
'  If Direction < 4 Then ' Queen or rook orthogonal
'    If Not CBool(WAttack(PinnedLoc) And QRAttackBit) Then Exit Function
'    AttackBit = QRAttackBit
'  Else ' Queen or bishop diagonal
'    If Not CBool(WAttack(PinnedLoc) And QBAttackBit) Then Exit Function
'    AttackBit = QBAttackBit
'  End If
'  Offset = QueenOffsets(Direction)
'
'  For k = 1 To 8
'    sq = PinnedLoc + Offset * k: Piece = Board(sq)
'    If Piece = FRAME Then Exit For
'    If Piece < NO_PIECE Then
'      If Piece = WQUEEN Then PinnedPieceB = True: Exit Function
'      If Piece = WROOK Then If Direction < 4 Then PinnedPieceB = True: Exit Function
'      If Piece = WBISHOP Then If Direction >= 4 Then PinnedPieceB = True: Exit Function
'      Exit For
'    Else
'      If Not (CBool(WAttack(sq) And AttackBit)) Then Exit For
'    End If
'  Next k
'
'End Function

Public Sub InitOutpostSq()
  Dim sq As Long

  For sq = SQ_A1 To SQ_H8
    If Rank(sq) >= 4 And Rank(sq) <= 6 Then WOutpostSq(sq) = True
    If Rank(sq) >= 3 And Rank(sq) <= 5 Then BOutpostSq(sq) = True
  Next sq

End Sub

Public Function NonPawnMatForSide(ByVal UseColOfSideToMove As Boolean) As Long
  If UseColOfSideToMove Then
    If bWhiteToMove Then
      NonPawnMatForSide = PieceCnt(WQUEEN) * ScoreQueen.MG + PieceCnt(WROOK) * ScoreRook.MG + PieceCnt(WBISHOP) * ScoreBishop.MG + PieceCnt(WKNIGHT) * ScoreKnight.MG
    Else
      NonPawnMatForSide = PieceCnt(BQUEEN) * ScoreQueen.MG + PieceCnt(BROOK) * ScoreRook.MG + PieceCnt(BBISHOP) * ScoreBishop.MG + PieceCnt(BKNIGHT) * ScoreKnight.MG
    End If
  Else
    If Not bWhiteToMove Then
      NonPawnMatForSide = PieceCnt(WQUEEN) * ScoreQueen.MG + PieceCnt(WROOK) * ScoreRook.MG + PieceCnt(WBISHOP) * ScoreBishop.MG + PieceCnt(WKNIGHT) * ScoreKnight.MG
    Else
      NonPawnMatForSide = PieceCnt(BQUEEN) * ScoreQueen.MG + PieceCnt(BROOK) * ScoreRook.MG + PieceCnt(BBISHOP) * ScoreBishop.MG + PieceCnt(BKNIGHT) * ScoreKnight.MG
    End If
  End If
End Function

Public Function NonPawnMat() As Long
  NonPawnMat = (PieceCnt(WQUEEN) + PieceCnt(BQUEEN)) * ScoreQueen.MG + (PieceCnt(WROOK) + PieceCnt(BROOK)) * ScoreRook.MG + (PieceCnt(WBISHOP) + PieceCnt(BBISHOP)) * ScoreBishop.MG + (PieceCnt(WKNIGHT) + PieceCnt(BKNIGHT)) * ScoreKnight.MG
End Function

Public Function MaterialTotal() As Long
  ' from view of white
  MaterialTotal = (PieceCnt(WQUEEN) - PieceCnt(BQUEEN)) * ScoreQueen.MG + (PieceCnt(WROOK) - PieceCnt(BROOK)) * ScoreRook.MG + (PieceCnt(WBISHOP) - PieceCnt(BBISHOP)) * ScoreBishop.MG + (PieceCnt(WKNIGHT) - PieceCnt(BKNIGHT)) * ScoreKnight.MG + (PieceCnt(WPAWN) - PieceCnt(BPAWN)) * ScorePawn.MG
End Function

Public Sub CheckWQueenWeek(ByVal sq As Long, ByVal Offset As Long, ByVal Direction As Long, ByRef Result As Boolean)
  ' Queen pinned or discovered threat possible
  If Result Then Exit Sub ' count only once
  Dim r As Long
  Result = False: r = sq + Offset ' next sq in same direction
  Select Case Board(r)
  Case BROOK: If Direction < 4 Then Result = True
  Case BBISHOP: If Direction > 3 Then Result = True
  Case NO_PIECE:
    If Direction < 4 Then ' Rook
      ' 2nd part: compare both attackbits, may be from different rooks: R1Attackbit or R2Attackbit
      If CBool(BAttack(sq) And RAttackBit) Then If (BAttack(r) And RAttackBit) = (BAttack(sq) And RAttackBit) Then Result = True
    Else ' Bishop?
      If CBool(BAttack(sq) And BAttackBit) Then If (BAttack(r) And BAttackBit) = (BAttack(sq) And BAttackBit) Then Result = True
    End If
  End Select
End Sub

Public Sub CheckBQueenWeek(ByVal sq As Long, ByVal Offset As Long, ByVal Direction As Long, ByRef Result As Boolean)
  ' Queen pinned or discovered threat possible
  If Result Then Exit Sub ' count only once
  Dim r As Long
  Result = False: r = sq + Offset ' next sq in same direction
  Select Case Board(r)
  Case WROOK: If Direction < 4 Then Result = True
  Case WBISHOP: If Direction > 3 Then Result = True
  Case NO_PIECE:
    If Direction < 4 Then ' Rook
      ' 2nd part: compare both attackbits, may be from different rooks: R1Attackbit or R2Attackbit
      If CBool(WAttack(sq) And RAttackBit) Then If (WAttack(r) And RAttackBit) = (WAttack(sq) And RAttackBit) Then Result = True
    Else ' Bishop?
      If CBool(WAttack(sq) And BAttackBit) Then If (WAttack(r) And BAttackBit) = (WAttack(sq) And BAttackBit) Then Result = True
    End If
  End Select
End Sub

