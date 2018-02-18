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
Public Hanging                    As TScore
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
Public bThreadTrace               As Boolean
Dim PassedPawns(16)               As Long ' List of passed pawns (Square)
Dim PassedPawnsCnt                As Long
Dim WPassedPawnAttack             As Long, BPassedPawnAttack As Long
Public PushClose(8)               As Long
Public PushAway(8)                As Long
Public PushToEdges(MAX_BOARD)     As Long
Public WOutpostSq(MAX_BOARD)      As Boolean
Public BOutpostSq(MAX_BOARD)      As Boolean
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
Dim Stoppers                       As Boolean, Neighbours As Boolean, Doubled As Boolean, Lever As Boolean, Supported As Boolean, Phalanx As Long, LeverPush As Long
Public PassedPawnFileBonus(8)      As TScore
Public PassedPawnRankBonus(8)      As TScore
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
    ScorePawn.MG = Val(ReadINISetting("PAWN_VAL_MG", "171"))
    ScorePawn.EG = Val(ReadINISetting("PAWN_VAL_EG", "240"))
    ScoreKnight.MG = Val(ReadINISetting("KNIGHT_VAL_MG", "764"))
    ScoreKnight.EG = Val(ReadINISetting("KNIGHT_VAL_EG", "848"))
    ScoreBishop.MG = Val(ReadINISetting("BISHOP_VAL_MG", "826"))
    ScoreBishop.EG = Val(ReadINISetting("BISHOP_VAL_EG", "891"))
    ScoreRook.MG = Val(ReadINISetting("ROOK_VAL_MG", "1282"))
    ScoreRook.EG = Val(ReadINISetting("ROOK_VAL_EG", "1373"))
    ScoreQueen.MG = Val(ReadINISetting("QUEEN_VAL_MG", "2526"))
    ScoreQueen.EG = Val(ReadINISetting("QUEEN_VAL_EG", "2646"))
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
'---           Value scaled to stockfish pawn endgame value (240 = 1 pawn)
'---
'---  Steps:
'---         1. Loop over all pieces to fill pawn structure array, pawn threats,
'---            calculate material + piece square values
'---         2. Check material draw
'---         3. Loop over all pieces: evaluate each piece except kings.
'---            do a move generation to calculate mobility, attackers, defenders
'---         3b. Loop over pawns: (attack info needed) evaluate save pawn pushes
'---         4. calculate king safety ( shelter, pawn storm, check attacks )
'---         5. calculate trapped bishops, passed pawns, king distance to best pawn, center control
'---         6. calculate threats
'---         7. Add all evalution terms weighted by variables set in INI file:
'---             Material + Position(general) + PawnStructure + PassedPawns + Mobility +
'---             KingSafetyComputer + KingSafetyOpponent + Threats
'---         8. Add tempo value for side to move
'---         9. invert score for black to move and return evaluation value
'---------------------------------------------------------------------------------------------------
Public Function Eval() As Long
  Dim a                       As Long, i As Long, Square As Long, Target As Long, Offset As Long, MobCnt As Long, r As Long, rr As Long, AttackBit As Long, ForkCnt As Long, SC As TScore
  Dim WPos                    As TScore, BPos As TScore, WPassed As TScore, BPassed As TScore, WMobility As TScore, BMobility As TScore
  Dim WPawnStruct             As TScore, BPawnStruct As TScore, Piece As Long, WPawnCnt As Long, BPawnCnt As Long
  Dim WKSafety                As TScore, BKSafety As TScore, bDoWKSafety As Boolean, bDoBKSafety As Boolean
  Dim WKingAdjacentZoneAttCnt As Long, BKingAdjacentZoneAttCnt As Long, WKingAttPieces As Long, BKingAttPieces As Long
  Dim KingDanger              As Long, Undefended As Long, RankNum As Long, RelRank As Long
  Dim FileNum                 As Long, MinWKingPawnDistance As Long, MinBKingPawnDistance As Long
  Dim DefByPawn               As Long, AttByPawn As Long, bAllDefended As Boolean, BlockSqDefended As Boolean, WPinnedCnt As Long, BPinnedCnt As Long
  Dim RankPath                As Long, sq As Long, WSemiOpenFiles As Long, BSemiOpenFiles As Long
  Dim BlockSq                 As Long, MBonus As Long, EBonus As Long, k As Long, UnsafeCnt As Long, PieceAttackBit As Long
  Dim OwnCol                  As Long, OppCol As Long, MoveUp As Long, OwnKingLoc As Long, OppKingLoc As Long, BlockSqUnsafe As Boolean
  Dim WBishopsOnBlackSq       As Long, WBishopsOnWhiteSq As Long, BBishopsOnBlackSq As Long, BBishopsOnWhiteSq As Long
  Dim WPawnCntOnWhiteSq       As Long, BPawnCntOnWhiteSq As Long
  Dim WKingFile               As Long, BKingFile As Long, WFrontMostPassedPawnRank As Long, BFrontMostPassedPawnRank As Long, ScaleFactor As Long
  Dim WChecksCounted          As Long, BChecksCounted As Long, WOtherChecksCounted As Long, BOtherChecksCounted As Long, KingLevers As Long
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

  '--- 1. loop over pieces: count pieces for material totals and game phase calculation. add piece square table score.
  '----                     calc pawn min/max rank positions per file; pawn attacks(for mobility used later)
  
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
      Case BPAWN
        BAttack(Square + SQ_DOWN_LEFT) = BAttack(Square + SQ_DOWN_LEFT) Or PLAttackBit: BAttack(Square + SQ_DOWN_RIGHT) = BAttack(Square + SQ_DOWN_RIGHT) Or PRAttackBit
        FileNum = File(Square): RankNum = Rank(Square): BPawns(FileNum) = BPawns(FileNum) + 1
        If RankNum < PawnsBMin(FileNum) Then PawnsBMin(FileNum) = RankNum
        If RankNum > PawnsBMax(FileNum) Then PawnsBMax(FileNum) = RankNum
        If MaxDistance(BKingLoc, Square) < MinBKingPawnDistance Then MinBKingPawnDistance = MaxDistance(BKingLoc, Square)
        If ColorSq(Square) = COL_WHITE Then BPawnCntOnWhiteSq = BPawnCntOnWhiteSq + 1 ' for Bishop eval
    End Select

lblNextPieceCnt:
  Next

  '--- KPK endgame: Eval if promoted pawn cannot be captured
  If NonPawnMaterial = 0 And (WPawnCnt + BPawnCnt = 1) Then
    If WPawnCnt = 1 Then
      If bWhiteToMove Then
        sq = PieceSqList(WPAWN, 1)
        If Rank(sq) = 7 Then
          If sq + SQ_UP <> WKingLoc Then ' own king not at promote square
            If MaxDistance(BKingLoc, sq + SQ_UP) > 1 Or MaxDistance(WKingLoc, sq + SQ_UP) = 1 Then
              Eval = VALUE_KNOWN_WIN: GoTo lblEndEval
            End If
          End If
        End If
      End If
    Else
      If Not bWhiteToMove Then
        sq = PieceSqList(BPAWN, 1)
        If Rank(sq) = 2 Then
          If sq + SQ_DOWN <> BKingLoc Then ' own king not at promote square
            If MaxDistance(WKingLoc, sq + SQ_DOWN) > 1 Or MaxDistance(BKingLoc, sq + SQ_DOWN) = 1 Then
              Eval = -VALUE_KNOWN_WIN: GoTo lblEndEval
            End If
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
  '--- EVAL Loop over pieces ------------------------------------------
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
    If BPawns(FileNum) = 0 Then WSemiOpenFiles = WSemiOpenFiles + 12 \ WPawns(FileNum) ' only count once per file, so 12 \ WPawns(FileNum) works for 1,2,3,4 pawns
    Opposed = (BPawns(FileNum) > 0) And RankNum < PawnsBMax(FileNum)
    Stoppers = (PawnsBMax(FileNum + 1) > RankNum Or PawnsBMax(FileNum) > RankNum Or PawnsBMax(FileNum - 1) > RankNum)
    Lever = (AttByPawn > 0)
    LeverPush = AttackBitCnt(BAttack(Square + SQ_UP) And PAttackBit)
    Doubled = (Board(Square + SQ_UP) = WPAWN)
    Neighbours = (WPawns(FileNum + 1) > 0 Or WPawns(FileNum - 1) > 0)
    Phalanx = Abs(Board(Square + SQ_LEFT) = WPAWN) + Abs(Board(Square + SQ_RIGHT) = WPAWN)
    Supported = (DefByPawn > AttByPawn)
    If Not Neighbours Or Lever Or RelRank >= 5 Then
      Backward = False
    Else
      If Board(Square + SQ_UP) = BPAWN Then
        Backward = True
      Else
        Backward = (PawnsWMin(FileNum + 1) > RankNum And PawnsWMin(FileNum - 1) > RankNum)
      End If
    End If
    Passed = False
    If Not Stoppers And Not Lever And Not LeverPush Then
      If DefByPawn >= AttByPawn Or bWhiteToMove Then
        If Phalanx >= LeverPush Then Passed = True
      End If
    End If
    If Not Passed And Supported And RankNum >= 5 Then ' sacrify supporter pawn to create passer?
      If PawnsBMax(FileNum) = RankNum + 1 Then
        If PawnsBMax(FileNum - 1) < RankNum Then
          If CBool(WAttack(Square) And PRAttackBit) Then ' left side supporter pawn (attacks to right)
            If Board(Square + SQ_LEFT) >= NO_PIECE Then  ' can move forward to attack stopper
              If Not CBool(BAttack(Square + SQ_LEFT) And PRAttackBit) Then ' no second left to right attacker from file-2
                Passed = True
              End If
            End If
          End If
        End If
        If Not Passed Then
          If PawnsBMax(FileNum + 1) < RankNum Then
            If CBool(WAttack(Square) And PLAttackBit) Then ' right side supporter pawn (attacks from left)
              If Board(Square + SQ_RIGHT) >= NO_PIECE Then  ' can move forward to attack stopper
                If Not CBool(BAttack(Square + SQ_RIGHT) And PLAttackBit) Then ' no second right to left attacker
                  Passed = True
                End If
              End If
            End If
          End If
        End If
      End If
    End If
    '--- pawn score
    If Lever Then AddScore SC, LeverBonus(RelRank)
    If Supported Or Phalanx Then ' Connected
      AddScore SC, ConnectedBonus(Abs(Opposed), Abs(Phalanx <> 0), DefByPawn, RelRank)
    Else
      If Not Neighbours Then MinusScore SC, IsolatedPenalty(Abs(Opposed))
      If Backward Then MinusScore SC, BackwardPenalty(Abs(Opposed))
      If Doubled Then MinusScore SC, DoubledPenalty
    End If
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
    If Passed And Not Doubled Then
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
    If WPawns(FileNum) = 0 Then BSemiOpenFiles = BSemiOpenFiles + 12 \ BPawns(FileNum)
    Opposed = (WPawns(FileNum) > 0) And RankNum > PawnsWMin(FileNum)
    Stoppers = PawnsWMin(FileNum + 1) < RankNum Or PawnsWMin(FileNum) < RankNum Or PawnsWMin(FileNum - 1) < RankNum
    Lever = (AttByPawn > 0)
    LeverPush = AttackBitCnt(WAttack(Square + SQ_DOWN) And PAttackBit)
    Doubled = (Board(Square + SQ_DOWN) = BPAWN)
    Neighbours = (BPawns(FileNum + 1) > 0 Or BPawns(FileNum - 1) > 0)
    Phalanx = Abs(Board(Square + SQ_LEFT) = BPAWN) + Abs(Board(Square + SQ_RIGHT) = BPAWN)
    Supported = (DefByPawn > AttByPawn)
    If Not Neighbours Or Lever Or RelRank >= 5 Then
      Backward = False
    Else
      If Board(Square + SQ_DOWN) = WPAWN Then
        Backward = True
      Else
        Backward = (PawnsBMax(FileNum + 1) < RankNum And PawnsBMax(FileNum - 1) < RankNum)
      End If
    End If
    Passed = False
    If Not Stoppers And Not Lever And Not LeverPush Then
      If DefByPawn >= AttByPawn Or Not bWhiteToMove Then
        If Phalanx >= LeverPush Then Passed = True
      End If
    End If
    If Not Passed And Supported And RankNum <= 4 Then ' sacrify supporter pawn to create passer?
      If PawnsWMin(FileNum) = RankNum - 1 Then
        If PawnsWMin(FileNum - 1) > RankNum Then
          If CBool(BAttack(Square) And PRAttackBit) Then ' left side supporter pawn
            If Board(Square + SQ_LEFT) >= NO_PIECE Then  ' can move forward to attack stopper
              If Not CBool(WAttack(Square + SQ_LEFT) And PRAttackBit) Then ' no second left to right attacker from file-2
                Passed = True
              End If
            End If
          End If
        End If
        If Not Passed Then
          If PawnsWMin(FileNum + 1) > RankNum Then
            If CBool(BAttack(Square) And PLAttackBit) Then ' right side supporter pawn
              If Board(Square + SQ_RIGHT) >= NO_PIECE Then  ' can move forward to attack stopper
                If Not CBool(WAttack(Square + SQ_RIGHT) And PLAttackBit) Then ' no second right to left attacker
                  Passed = True
                End If
              End If
            End If
          End If
        End If
      End If
    End If
    '--- pawn score
    If Lever Then AddScore SC, LeverBonus(RelRank)
    If Supported Or Phalanx Then ' Connected
      AddScore SC, ConnectedBonus(Abs(Opposed), Abs(Phalanx <> 0), DefByPawn, RelRank)
    Else
      If Not Neighbours Then MinusScore SC, IsolatedPenalty(Abs(Opposed))
      If Backward Then MinusScore SC, BackwardPenalty(Abs(Opposed))
      If Doubled Then MinusScore SC, DoubledPenalty
    End If
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
          Case WKING: ' ignore
          Case BKING: MobCnt = MobCnt + 1: ForkCnt = ForkCnt + 1
          Case WEP_PIECE, BEP_PIECE:
            If (Not CBool(BAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1: SC.MG = SC.MG + 3
          Case Else: If (Not CBool(BAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
        End Select

        If r < 2 Then
          If WOutpostSq(Target) Then ' Empty or opp piece: square can be occupied
            If PieceColor(Board(Target)) <> COL_WHITE Then If (Not CBool(BAttack(Target) And PAttackBit)) Then r = 2 Else r = 1
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
    If r > 0 And r < 3 Then AddScore SC, ReachableOutpostKnight(r - 1)
    If CBool(BAttack(Square) And PAttackBit) Then AddPawnThreat BThreat, COL_WHITE, PieceType(Board(Square)), Square
    AddScoreWithFactor WKSafety, KingProtector(PT_KNIGHT), MaxDistance(Square, WKingLoc) ' defends king?
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
        AddScore SC, OutpostBonusKnight(Abs(CBool(BAttack(Square) And PAttackBit)))
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
          Case BKING: ' Ignore
          Case WKING: If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
            If (Not CBool(WAttack(Target) And PAttackBit)) Then ForkCnt = ForkCnt + 1
          Case WEP_PIECE, BEP_PIECE:
            If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1: SC.MG = SC.MG + 3
          Case Else: If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1
        End Select

        If r < 2 Then
          If BOutpostSq(Target) Then ' Empty or opp piece: square can be occupied
            If PieceColor(Board(Target)) <> COL_BLACK Then If (Not CBool(WAttack(Target) And PAttackBit)) Then r = 2 Else r = 1
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
    If r > 0 And r < 3 Then AddScore SC, ReachableOutpostKnight(r - 1)
    If CBool(WAttack(Square) And PAttackBit) Then AddPawnThreat WThreat, COL_BLACK, PieceType(Board(Square)), Square
    AddScoreWithFactor BKSafety, KingProtector(PT_KNIGHT), MaxDistance(Square, BKingLoc)  ' defends king?
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
          Case WQUEEN: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            AttackBit = BXrayAttackBit  '--- Continue xray
          Case WEP_PIECE, BEP_PIECE:
            If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: If Offset > 0 Then SC.MG = SC.MG + 2
          Case Else: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            Exit Do ' own bishop or knight
        End Select

        If r < 2 Then
          If WOutpostSq(Target) Then ' Empty or opp piece: square can be occupied
            If PieceColor(Board(Target)) <> COL_WHITE Then If (Not CBool(BAttack(Target) And PAttackBit)) Then r = 2 Else r = 1
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
    If r > 0 And r < 3 Then AddScore SC, ReachableOutpostBishop(r - 1)
    If CBool(BAttack(Square) And PAttackBit) Then AddPawnThreat BThreat, COL_WHITE, PieceType(Board(Square)), Square
    AddScoreWithFactor WKSafety, KingProtector(PT_BISHOP), MaxDistance(Square, WKingLoc) ' defends king?
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
          Case BQUEEN:
            If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            AttackBit = BXrayAttackBit '--- Continue xray
          Case WEP_PIECE, BEP_PIECE:
            If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1:  If Offset < 0 Then SC.MG = SC.MG + 2
          Case Else: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
            Exit Do ' own bishop or knight
        End Select

        If r < 2 Then
          If BOutpostSq(Target) Then ' Empty or opp piece: square can be occupied
            If PieceColor(Board(Target)) <> COL_BLACK Then If (Not CBool(WAttack(Target) And PAttackBit)) Then r = 2 Else r = 1
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
    If r > 0 And r < 3 Then AddScore SC, ReachableOutpostBishop(r - 1)
    If CBool(WAttack(Square) And PAttackBit) Then AddPawnThreat WThreat, COL_BLACK, PieceType(Board(Square)), Square
    AddScoreWithFactor BKSafety, KingProtector(PT_BISHOP), MaxDistance(Square, BKingLoc) ' defends king?
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
            If Not CBool(BAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1
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
            If WKingFile < 5 Then
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
    AddScoreWithFactor WKSafety, KingProtector(PT_ROOK), MaxDistance(Square, WKingLoc) ' defends king?
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
            If Not CBool(WAttack(Target) And PBNAttackBit) Then MobCnt = MobCnt + 1
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
            If BKingFile < 5 Then
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
    AddScoreWithFactor BKSafety, KingProtector(PT_ROOK), MaxDistance(Square, BKingLoc) ' defends king?
    AddScore BPos, SC
    If bEvalTrace Then WriteTrace "BRook: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & BPos.MG & ", " & BPos.EG
  Next a

  '--------------------------------------------------------------------
  '---- WHITE QUEENs ( last - full attack info needed for mobility )  -
  '--------------------------------------------------------------------
  For a = 1 To PieceSqListCnt(WQUEEN)
    Square = PieceSqList(WQUEEN, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = RankNum: SC.MG = 0: SC.EG = 0
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
            Exit Do   'Defends pawn
          Case BPAWN:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1: If AttackBit = QAttackBit Then AddThreat COL_BLACK, PT_PAWN, PT_QUEEN, Square, Target
            SC.MG = SC.MG + 7: SC.EG = SC.EG + 7
            Exit Do   'Attack pawn
          Case BKNIGHT:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = QAttackBit Then AddThreat COL_BLACK, PieceType(Board(Target)), PT_QUEEN, Square, Target
            Exit Do
          Case BBISHOP:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = QAttackBit Then AddThreat COL_BLACK, PieceType(Board(Target)), PT_QUEEN, Square, Target
            Exit Do
          Case BROOK:
            If Not CBool(BAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = QAttackBit Then AddThreat COL_BLACK, PieceType(Board(Target)), PT_QUEEN, Square, Target
            Exit Do
          Case WKING: Exit Do ' ignore
          Case BKING: MobCnt = MobCnt + 1
            Exit Do
          Case BQUEEN: If AttackBit = QAttackBit Then AddThreat COL_BLACK, PT_QUEEN, PT_QUEEN, Square, Target: MobCnt = MobCnt + 1
            Exit Do
          Case WBISHOP: If Not CBool(BAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1: SC.MG = SC.MG + 4: SC.EG = SC.EG + 2
            If i > 3 Then AttackBit = QXrayAttackBit Else Exit Do
          Case WKNIGHT: If Not CBool(BAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
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
    AddScoreWithFactor WKSafety, KingProtector(PT_QUEEN), MaxDistance(Square, WKingLoc) ' defends king?
    AddScore WPos, SC
    If bEvalTrace Then WriteTrace "WQueen: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & WPos.MG & ", " & WPos.EG
  Next a

  '--------------------------------------------------------------------
  '---- BLACK QUEENs ( last - full attack info needed for mobility ) --
  '--------------------------------------------------------------------
  For a = 1 To PieceSqListCnt(BQUEEN)
    Square = PieceSqList(BQUEEN, a): FileNum = File(Square): RankNum = Rank(Square): RelRank = (9 - RankNum): SC.MG = 0: SC.EG = 0
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
            Exit Do   'Defends pawn
          Case WPAWN:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1: If AttackBit = QAttackBit Then AddThreat COL_WHITE, PT_PAWN, PT_QUEEN, Square, Target
            SC.MG = SC.MG + 7: SC.EG = SC.EG + 7
            Exit Do   'Attack pawn
          Case WKNIGHT:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = QAttackBit Then AddThreat COL_WHITE, PieceType(Board(Target)), PT_QUEEN, Square, Target
            Exit Do
          Case WBISHOP:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = QAttackBit Then AddThreat COL_WHITE, PieceType(Board(Target)), PT_QUEEN, Square, Target
            Exit Do
          Case WROOK:
            If Not CBool(WAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
            If AttackBit = QAttackBit Then AddThreat COL_WHITE, PieceType(Board(Target)), PT_QUEEN, Square, Target
            Exit Do
          Case BKING: Exit Do ' Ignore
          Case WKING: MobCnt = MobCnt + 1
            Exit Do
          Case WQUEEN:  If AttackBit = QAttackBit Then AddThreat COL_WHITE, PT_QUEEN, PT_QUEEN, Square, Target: MobCnt = MobCnt + 1
            Exit Do
          Case BBISHOP: If Not CBool(WAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1: SC.MG = SC.MG + 4: SC.EG = SC.EG + 2
            If i > 3 Then AttackBit = QXrayAttackBit Else Exit Do
          Case BKNIGHT: If Not CBool(WAttack(Target) And PNBRAttackBit) Then MobCnt = MobCnt + 1
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
    AddScoreWithFactor BKSafety, KingProtector(PT_QUEEN), MaxDistance(Square, BKingLoc) ' defends king?
    AddScore BPos, SC
    If bEvalTrace Then WriteTrace "BQueen: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & BPos.MG & ", " & BPos.EG
  Next a

  '--------------------------------------------------------------------
  '---- Pass for pawn push ( full attack info needed for mobility )
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

  If SC.MG > 0 Then AddScore WPawnStruct, SC
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

  If SC.MG > 0 Then AddScore BPawnStruct, SC
  '--- End pass for pawn push <<<<
  '
  '--- Global eval scores -------------------------------------------
  '
  If bEndgame Then
    WKSafety = ZeroScore: BKSafety = ZeroScore
  Else
    Dim Bonus            As Long
    Dim KingOnlyDefended As Long, bSafe As Boolean
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
          '--- Check threats at king ring
          Undefended = 0: KingOnlyDefended = 0: WKingAttPieces = 0: KingLevers = 0
          '  add the 2 or 3 squares in front of king ring: king G1 => F3+G3+H3
          If RankNum = 1 Then
            For Target = WKingLoc + 19 To WKingLoc + 21
              If Board(Target) <> FRAME Then
                If BAttack(Target) <> 0 Then
                  If WAttack(Target) = 0 Then If PieceColor(Board(Target)) <> BCOL Then Undefended = Undefended + 1
                  WKingAttPieces = WKingAttPieces Or BAttack(Target)
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
                WKingAttPieces = WKingAttPieces Or BAttack(Target)
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
                  If CBool(r And RBAttackBit) Then
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
                      If CBool(r And RAttackBit) Then
                        If Not CBool(WChecksCounted And RAttackBit) Then ' count only once!
                          If bSafe Then
                            KingDanger = KingDanger + RookCheck
                            WChecksCounted = (WChecksCounted Or RAttackBit)
                            WOtherChecksCounted = WOtherChecksCounted And Not RAttackBit ' remove in others: do not count both cases
                          Else
                            If Not CBool(WAttack(Target) And PAttackBit) Then WOtherChecksCounted = WOtherChecksCounted Or RAttackBit
                          End If
                        End If
                      End If
                    Else ' i >= 4
                      ' Bishop checks
                      If CBool(r And BAttackBit) Then
                        If Not CBool(WChecksCounted And BAttackBit) Then ' count only once!
                          If bSafe Then
                            KingDanger = KingDanger + BishopCheck
                            WChecksCounted = (WChecksCounted Or BAttackBit)
                            WOtherChecksCounted = WOtherChecksCounted And Not BAttackBit ' remove in others: do not count both cases
                          Else
                            If Not CBool(WAttack(Target) And PAttackBit) Then WOtherChecksCounted = WOtherChecksCounted Or BAttackBit
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
                  If (Board(Target) Mod 2 = WCOL) Then ' own piece
                    If i < 4 Then ' orthogonal
                      If CBool(BAttack(Target) And QRAttackBit) Then  ' rook or queen, direction not clear

                        For k = 1 To 7
                          sq = Target + Offset * k: Piece = Board(sq)
                          If Piece = FRAME Then Exit For
                          If Piece < NO_PIECE Then
                            If Piece = BQUEEN Or Piece = BROOK Then If (PieceType(Piece) <> PieceType(Board(Target))) Then WPinnedCnt = WPinnedCnt + 1
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
                            If Piece = BQUEEN Or Piece = BBISHOP Then If (PieceType(Piece) <> PieceType(Board(Target))) Then WPinnedCnt = WPinnedCnt + 1
                            Exit For
                          Else
                            If Not (CBool(BAttack(sq) And QBAttackBit)) Then Exit For
                          End If
                        Next k

                      End If
                    End If
                  End If
                  ' --- Piece found - exit direction loop
                  If Board(Target) <> WQUEEN Then Exit Do
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
                  If Not CBool(WChecksCounted And NAttackBit) Then ' count only once!
                    If bSafe Then
                      KingDanger = KingDanger + KnightCheck
                      WChecksCounted = (WChecksCounted Or NAttackBit)
                      WOtherChecksCounted = WOtherChecksCounted And Not NAttackBit ' remove in others: do not count both cases
                    Else
                      If Not CBool(WAttack(Target) And PAttackBit) Then WOtherChecksCounted = WOtherChecksCounted Or NAttackBit
                    End If
                  End If
                  ' Knight check fork?
                  If WAttack(Target) = 0 Or (WAttack(Target) = QAttackBit And (BAttack(Target) <> NAttackBit)) Then ' no attack

                    For k = 0 To 7

                      Select Case Board(Target + KnightOffsets(k))
                        Case WQUEEN: If bWhiteToMove Then AddScoreVal BThreat, 25, 35 Else AddScoreVal BThreat, 45, 55
                        Case WROOK: If bWhiteToMove Then AddScoreVal BThreat, 15, 20 Else AddScoreVal BThreat, 30, 40
                      End Select

                    Next

                  End If
                End If '<<< CBool(BAttack(Target) And NAttackBit)
              End If '<<<  Board(Target) <> FRAME
            End If '<<< PieceCnt(BKNIGHT) > 0
          Next i '<<< direction

          If WKingAttPieces <> 0 Then AddWKingAttackers WKingAttPieces

          If WKingAttackersCount > 1 - PieceCnt(BQUEEN) Then
            ' Calc Other Checks
            If WOtherChecksCounted > 0 Then
              i = AttackBitCnt(WOtherChecksCounted): WKSafety.MG = WKSafety.MG - i * OtherCheck.MG: WKSafety.EG = WKSafety.EG - i * OtherCheck.EG
            End If
                      
            ' total KingDanger
            KingDanger = KingDanger + WKingAttackersCount * WKingAttackersWeight + 102 * WKingAdjacentZoneAttCnt + Abs(KingLevers > 0) * 64 _
                         + 191 * (KingOnlyDefended + Undefended) + 143 * (Abs(WPinnedCnt > 0)) - 848 * Abs(PieceCnt(BQUEEN) = 0) - 9 * Bonus \ 8 + 4 * (BNonPawnPieces + PieceCnt(BPAWN) + 1)
            
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
        
        ' King tropism: firstly, find squares that opponent attacks in our king flank
        ' Secondly, add the squares which are attacked twice in that flank and
        ' which are not defended by our pawns.
        GetKingFlankFiles WKingLoc, r, rr
        Bonus = 0

        For k = SQ_A1 - 1 To SQ_A1 - 1 + 40 Step 10 ' start square - 1 of rank 1-5
          For Square = k + r To k + rr     ' files king flank
            If BAttack(Square) <> 0 Then
              Bonus = Bonus + 1
              If Not CBool(WAttack(Square) And PAttackBit) Then ' not protected by pawn
                If AttackBitCnt(BAttack(Square)) > 1 Then Bonus = Bonus + 1   ' Attacked twice?
              End If
            End If
          Next
        Next

        If Bonus > 0 Then WKSafety.MG = WKSafety.MG - 7 * Bonus
        
        ' Bonus for a dangerous pawn in the center near the opponent king, for instance pawn e5 against king g8.
        If FileNum >= 4 Then If Board(SQ_E5) = WPAWN Then BKSafety.MG = BKSafety.MG - 5
        If FileNum <= 5 Then If Board(SQ_D5) = WPAWN Then BKSafety.MG = BKSafety.MG - 5 ' both possible if king centered
        
        ' Pawnless king flank penalty
        k = 0
        For i = r To rr
          If WPawns(i) + BPawns(i) > 0 Then k = 1: Exit For
        Next
        If k = 0 Then MinusScore WKSafety, PawnlessFlank
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
          '--- Check threats at king ring
          Undefended = 0: KingOnlyDefended = 0: BKingAttPieces = 0
          '  add the 2 or 3 squares in front of king ring: king G8 => F6+G6+H6
          If RankNum = 8 Then

            For Target = BKingLoc - 21 To BKingLoc - 19
              If Board(Target) <> FRAME Then
                If WAttack(Target) <> 0 Then
                  If BAttack(Target) = 0 Then If PieceColor(Board(Target)) <> WCOL Then Undefended = Undefended + 1
                  BKingAttPieces = BKingAttPieces Or WAttack(Target)
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
                BKingAdjacentZoneAttCnt = BKingAdjacentZoneAttCnt + AttackBitCnt(BAttack(Target) And Not PAttackBit)
                BKingAttPieces = BKingAttPieces Or WAttack(Target)
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
                  If CBool(r And RBAttackBit) Then
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
                      If CBool(r And RAttackBit) Then
                        If Not CBool(BChecksCounted And RAttackBit) Then ' count only once!
                          If bSafe Then
                            KingDanger = KingDanger + RookCheck
                            BChecksCounted = (BChecksCounted Or RAttackBit)
                            BOtherChecksCounted = BOtherChecksCounted And Not RAttackBit ' remove in others: do not count both cases
                          Else
                            If Not CBool(BAttack(Target) And PAttackBit) Then BOtherChecksCounted = BOtherChecksCounted Or RAttackBit
                          End If
                        End If
                      End If
                    Else ' i >= 4
                      ' Bishop checks
                      If CBool(r And BAttackBit) Then
                        If Not CBool(BChecksCounted And BAttackBit) Then ' count only once!
                          If bSafe Then
                            KingDanger = KingDanger + BishopCheck
                            BChecksCounted = (BChecksCounted Or BAttackBit)
                            BOtherChecksCounted = BOtherChecksCounted And Not BAttackBit ' remove in others: do not count both cases
                          Else
                            If Not CBool(BAttack(Target) And PAttackBit) Then BOtherChecksCounted = BOtherChecksCounted Or BAttackBit
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
                            If Piece = WQUEEN Or Piece = WROOK Then If (PieceType(Piece) <> PieceType(Board(Target))) Then BPinnedCnt = BPinnedCnt + 1
                            Exit For
                          Else
                            If Not (CBool(WAttack(sq) And QRAttackBit)) Then Exit For
                          End If
                        Next k

                      End If
                    Else ' i>4 diagonal
                      If CBool(WAttack(Target) And QBAttackBit) Then  ' bishop or queen, direction not clear

                        For k = 1 To 7
                          sq = Target + Offset * k: Piece = Board(sq)
                          If Piece = FRAME Then Exit For
                          If Piece < NO_PIECE Then
                            If Piece = WQUEEN Or Piece = WBISHOP Then BPinnedCnt = BPinnedCnt + 1
                            Exit For
                          Else
                            If Not (CBool(WAttack(sq) And QBAttackBit)) Then Exit For
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
                  If Not CBool(BChecksCounted And NAttackBit) Then ' count only once!
                    If bSafe Then
                      KingDanger = KingDanger + KnightCheck
                      BChecksCounted = (BChecksCounted Or NAttackBit)
                      BOtherChecksCounted = BOtherChecksCounted And Not NAttackBit ' remove in others: do not count both cases
                    Else
                      If Not CBool(BAttack(Target) And PAttackBit) Then BOtherChecksCounted = BOtherChecksCounted Or NAttackBit
                    End If
                  End If
                  ' Knight check fork?
                  If BAttack(Target) = 0 Or (BAttack(Target) = QAttackBit And (WAttack(Target) <> NAttackBit)) Then ' field not defended or by queen only but other attacker

                    For k = 0 To 7

                      Select Case Board(Target + KnightOffsets(k))
                        Case BQUEEN: If Not bWhiteToMove Then AddScoreVal WThreat, 25, 35 Else AddScoreVal WThreat, 45, 55
                        Case BROOK: If Not bWhiteToMove Then AddScoreVal WThreat, 15, 20 Else AddScoreVal WThreat, 30, 40
                      End Select

                    Next

                  End If
                End If  '<<< CBool(WAttack(Target) And NAttackBit)
              End If ' <<< Board(Target) <> FRAME
            End If '<<< PieceCnt(WKNIGHT) > 0
          Next i '<<< direction
          
          If BKingAttPieces <> 0 Then AddBKingAttackers BKingAttPieces
          
          If BKingAttackersCount > 1 - PieceCnt(WQUEEN) Then
  
            ' Calc Other Checks
            If BOtherChecksCounted > 0 Then
              i = AttackBitCnt(BOtherChecksCounted): BKSafety.MG = BKSafety.MG - i * OtherCheck.MG: BKSafety.EG = BKSafety.EG - i * OtherCheck.EG
            End If
            
  
            ' total KingDanger
            KingDanger = KingDanger + BKingAttackersCount * BKingAttackersWeight + 102 * BKingAdjacentZoneAttCnt + Abs(KingLevers > 0) * 64 _
                         + 191 * (KingOnlyDefended + Undefended) + 143 * (Abs(BPinnedCnt > 0)) - 848 * Abs(PieceCnt(WQUEEN) = 0) - 9 * Bonus \ 8 + 4 * (WNonPawnPieces + PieceCnt(WPAWN) + 1)
            
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
        ' King tropism: firstly, find squares that opponent attacks in our king flank
        ' Secondly, add the squares which are attacked twice in that flank and
        ' which are not defended by our pawns.
        GetKingFlankFiles BKingLoc, r, rr
        Bonus = 0

        For k = SQ_A1 - 1 + 30 To SQ_A1 - 1 + 70 Step 10 ' start square - 1 of rank 5-8
          For Square = k + r To k + rr     ' files king flank
            If WAttack(Square) <> 0 Then
              Bonus = Bonus + 1
              If Not CBool(BAttack(Square) And PAttackBit) Then ' not protected by pawn
                If AttackBitCnt(WAttack(Square)) > 1 Then Bonus = Bonus + 1  ' Attacked twice?
              End If
            End If
          Next
        Next

        If Bonus > 0 Then BKSafety.MG = BKSafety.MG - 7 * Bonus
        
        ' Bonus for a dangerous pawn in the center near the opponent king, for instance pawn e5 against king g8.
        If FileNum >= 4 Then If Board(SQ_E4) = BPAWN Then WKSafety.MG = WKSafety.MG - 5
        If FileNum <= 5 Then If Board(SQ_D4) = BPAWN Then WKSafety.MG = WKSafety.MG - 5 ' both possible if king centered
        
        
        ' Pawnless king flank penalty
        k = 0
        For i = r To rr
          If WPawns(i) + BPawns(i) > 0 Then k = 1: Exit For
        Next
        If k = 0 Then MinusScore BKSafety, PawnlessFlank
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
  '
  '--- Eval threats -------------------------------------------
  '
  CalcThreats  ' in WThreat and BThreat
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
          If r <> 0 Then MBonus = MBonus + r * (10 + RelRank * 3): EBonus = EBonus + r * (20 + RelRank * RelRank)
        Else
          r = Sgn(Sgn(WBishopsOnBlackSq) - Sgn(BBishopsOnBlackSq)) ' 0 if both sides have same bishop color, else +1 or -1
          If r <> 0 Then MBonus = MBonus + r * (10 + RelRank * 3): EBonus = EBonus + r * (20 + RelRank * RelRank)
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
          If r <> 0 Then MBonus = MBonus + r * (10 + RelRank * 3): EBonus = EBonus + r * (20 + RelRank * RelRank)
        Else
          r = Sgn(Sgn(BBishopsOnBlackSq) - Sgn(WBishopsOnBlackSq)) ' 0 if both sides have same bishop color, else +1 or -1
          If r <> 0 Then MBonus = MBonus + r * (10 + RelRank * 3): EBonus = EBonus + r * (20 + RelRank * RelRank)
        End If
      End If
    End If
    '--- Path to promote square blocked? => penalty
    r = RelRank - 2: rr = (r * (r - 1))
    MBonus = PassedPawnRankBonus(r).MG: EBonus = PassedPawnRankBonus(r).EG
    ' Bonus based on rank ' SF6
    If rr <> 0 Then
      BlockSq = Square + MoveUp
      If Board(BlockSq) <> FRAME Then
        '  Adjust bonus based on the king's proximity
        EBonus = EBonus + MaxDistance(BlockSq, OppKingLoc) * 5 * rr - MaxDistance(BlockSq, OwnKingLoc) * 2 * rr
        'If blockSq is not the queening square then consider also a second push
        If RelRank <> 7 Then EBonus = EBonus - MaxDistance(BlockSq + MoveUp, OwnKingLoc) * rr
        'If the pawn is free to advance, then increase the bonus
        If Board(BlockSq) >= NO_PIECE Then
          k = 0: bAllDefended = True: BlockSqDefended = True: BlockSqUnsafe = False
          ' Rook or Queen attacking/defending from behind
          If CBool(BAttack(Square) And QRAttackBit) Or CBool(WAttack(Square) And QRAttackBit) Then

            For RankPath = RelRank - 1 To 1 Step -1
              sq = Square + (RankPath - RelRank) * MoveUp

              Select Case Board(sq)
                Case NO_PIECE:
                Case BROOK, BQUEEN:
                  If OwnCol = COL_WHITE Then BlockSqUnsafe = True
                  Exit For
                Case WROOK, WQUEEN:
                  If OwnCol = COL_BLACK Then BlockSqUnsafe = True
                  Exit For
                Case Else:
                  Exit For
              End Select

            Next

          End If

          For RankPath = RelRank + 1 To 8
            sq = Square + (RankPath - RelRank) * MoveUp
            OwnAttCnt = AttackBitCnt(AttackByCol(OwnCol, sq)): OppAttCnt = AttackBitCnt(AttackByCol(OppCol, sq))
            If OwnAttCnt = 0 And sq <> OwnKingLoc Then
              bAllDefended = False: If sq = BlockSq Then BlockSqDefended = False
            End If
            If PieceColor(Board(sq)) = OppCol Then
              If sq = BlockSq Then BlockSqUnsafe = True Else UnsafeCnt = UnsafeCnt + 1
            ElseIf OppAttCnt > 0 Then
              If CBool(AttackByCol(OwnCol, sq) And PAttackBit) And Not CBool(AttackByCol(OppCol, sq) And PAttackBit) Then
                ' Own pawn support but no enemy pawn attack: square is safe ( NOT SF LOGIC )
                ' Stop
              Else
                If sq = BlockSq Then BlockSqUnsafe = True Else UnsafeCnt = UnsafeCnt + 1
              End If
            End If
          Next RankPath

          If BlockSqUnsafe Then UnsafeCnt = UnsafeCnt + 1
          If UnsafeCnt = 0 Then
            k = 18
          ElseIf Not BlockSqUnsafe Then
            k = 8 '- UnsafeCnt
          Else
            k = 0
          End If
          If bAllDefended Then
            k = k + 6 '- UnsafeCnt \ 2
          ElseIf BlockSqDefended Then
            k = k + 4 '- UnsafeCnt \ 2
          End If
          If k <> 0 Then MBonus = MBonus + k * rr: EBonus = EBonus + k * rr
        Else
          If PieceColor(Board(BlockSq)) = OwnCol Then MBonus = MBonus + rr + r * 2: EBonus = EBonus + rr + r * 2
        End If
      End If
    End If ' rr>0
    If OwnCol = COL_WHITE Then
      If WPawnCnt < BPawnCnt Then EBonus = EBonus + EBonus \ 4
      MBonus = MBonus + PassedPawnFileBonus(FileNum).MG: EBonus = EBonus + PassedPawnFileBonus(FileNum).EG
      If BNonPawnMaterial = 0 Then EBonus = EBonus + 20
      If Board(Square + SQ_UP) = BPAWN Then MBonus = MBonus \ 2: EBonus = EBonus \ 2 ' supporter sacrify needed
      If bWhiteToMove Then MBonus = (MBonus * 105) \ 100:   EBonus = (EBonus * 105) \ 100
      AddScoreVal WPassed, MBonus, EBonus
      If 1000 + EBonus > WBestPawnVal Then WBestPawn = Square: WBestPawnVal = 1000 + EBonus ' new best pawn
      If bEvalTrace Then WriteTrace "WPassed: " & LocCoord(Square) & ">" & MBonus & ", " & EBonus
    ElseIf OwnCol = COL_BLACK Then
      If BPawnCnt < WPawnCnt Then EBonus = EBonus + EBonus \ 4
      MBonus = MBonus + PassedPawnFileBonus(FileNum).MG: EBonus = EBonus + PassedPawnFileBonus(FileNum).EG
      If WNonPawnMaterial = 0 Then EBonus = EBonus + 20
      If Board(Square + SQ_DOWN) = WPAWN Then MBonus = MBonus \ 2: EBonus = EBonus \ 2 ' supporter sacrify needed
      If Not bWhiteToMove Then MBonus = (MBonus * 105) \ 100:   EBonus = (EBonus * 105) \ 100
      AddScoreVal BPassed, MBonus, EBonus
      If 1000 + EBonus > BBestPawnVal Then BBestPawn = Square: BBestPawnVal = 1000 + EBonus ' new best pawn
      If bEvalTrace Then WriteTrace "BPassed: " & LocCoord(Square) & ">" & MBonus & ", " & EBonus
    End If
  Next a

  '---<<< end  Passed pawn
  '--- If both sides have only pawns, score for potential unstoppable pawns
  If WNonPawnMaterial + BNonPawnMaterial = 0 Then
    If WFrontMostPassedPawnRank > 0 Then AddScoreVal WPassed, 0, WFrontMostPassedPawnRank * 20
    If BFrontMostPassedPawnRank > 0 Then AddScoreVal BPassed, 0, WFrontMostPassedPawnRank * 20
  End If
  '---  Penalty for pawns on same color square of bishop
  If PieceCnt(WBISHOP) > 0 Then
    r = WPawnCntOnWhiteSq * WBishopsOnWhiteSq + (WPawnCnt - WPawnCntOnWhiteSq) * WBishopsOnBlackSq - WPawnCnt
    If r > 0 Then
      AddScoreVal WPos, -8 * r, -12 * r
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
    r = BPawnCntOnWhiteSq * BBishopsOnWhiteSq + (BPawnCnt - BPawnCntOnWhiteSq) * BBishopsOnBlackSq - BPawnCnt
    If r > 0 Then
      AddScoreVal BPos, -8 * r, -12 * r
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
  '--->>> Pawn Islands (groups of pawns) ---
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
  '--- Material total ---
  ' Piece values were set in SetGamePhase
  Dim AllTotal As TScore, MatEval As Long
  AllTotal.MG = (PieceCnt(WQUEEN) - PieceCnt(BQUEEN)) * ScoreQueen.MG + (PieceCnt(WROOK) - PieceCnt(BROOK)) * ScoreRook.MG + (PieceCnt(WBISHOP) - PieceCnt(BBISHOP)) * ScoreBishop.MG + (PieceCnt(WKNIGHT) - PieceCnt(BKNIGHT)) * ScoreKnight.MG + (WPawnCnt - BPawnCnt) * ScorePawn.MG
  AllTotal.EG = (PieceCnt(WQUEEN) - PieceCnt(BQUEEN)) * ScoreQueen.EG + (PieceCnt(WROOK) - PieceCnt(BROOK)) * ScoreRook.EG + (PieceCnt(WBISHOP) - PieceCnt(BBISHOP)) * ScoreBishop.EG + (PieceCnt(WKNIGHT) - PieceCnt(BKNIGHT)) * ScoreKnight.EG + (WPawnCnt - BPawnCnt) * ScorePawn.EG
  MatEval = ScaleScore(AllTotal)
  If bEvalTrace Then
    Debug.Print "Material: " & EvalSFTo100(AllTotal.MG) & "," & EvalSFTo100(AllTotal.EG)
  End If
  '
  '--- Scale Factor ---
  '
  ScaleFactor = 64 ' Normal ScaleFactor, scales EG value only
  If GamePhase < PHASE_MIDGAME Then
    '- just one pawn makes it difficult to win
    If WMaterial > BMaterial Then
      If WPawnCnt = 1 Then If WNonPawnMaterial - BNonPawnMaterial <= ScoreBishop.MG Then ScaleFactor = 48
    ElseIf BMaterial > WMaterial Then
      If BPawnCnt = 1 Then If BNonPawnMaterial - WNonPawnMaterial <= ScoreBishop.MG Then ScaleFactor = 48
    End If
    '- Endgame with opposite-colored bishops and no other pieces (ignoring pawns)
    '- is almost a draw, in case of KBP vs KB, it is even more a draw.
    If PieceCnt(WBISHOP) = 1 And PieceCnt(BBISHOP) = 1 And WBishopsOnWhiteSq = BBishopsOnBlackSq Then ' opposite-colored bishops
      If WNonPawnMaterial = ScoreBishop.MG And BNonPawnMaterial = ScoreBishop.MG Then
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
      If ScaleFactor = 64 Or ScaleFactor = 48 Then ' SCALE_FACTOR_NORMAL or SCALE_FACTOR_ONEPAWN
        If Abs(AllTotal.EG) < ScoreBishop.EG And WMaterial <> BMaterial Then
          'Endings where weaker side can place his king in front of the opponent's pawns are drawish.
          If WMaterial > BMaterial Then ' White is strong side
            If BBestPawnVal > 500 And WPawnCnt <= 2 Then ' passed pawns only
              If Rank(WKingLoc) < BFrontMostPassedPawnRank And MaxDistance(BBestPawn, WKingLoc) < BFrontMostPassedPawnRank Then
                ScaleFactor = 37 + 7 * BPawnCnt
              End If
            End If
          ElseIf BMaterial > WMaterial Then ' Black is strong side
            If WBestPawnVal > 500 And BPawnCnt <= 2 Then  ' passed pawns only
              If Rank(BKingLoc) > WFrontMostPassedPawnRank And MaxDistance(WBestPawn, BKingLoc) < 9 - WFrontMostPassedPawnRank Then
                ScaleFactor = 37 + 7 * WPawnCnt
              End If
            End If
          End If
        End If ' Abs(AllTotal.EG
      End If
    End If
  End If
  
  ' >>> Removed , bad results, too slow ???
  '--- Calculate space in opening phase for safe squares in center
  '
'  If NonPawnMaterial > SPACE_THRESHOLD Then
'    r = 0: rr = 0
'    For k = 3 To 6 ' files 3-6
'      For RankNum = 2 To 4 ' WHITE
'        Target = 10 + RankNum * 10 + k
'        If Board(Target) <> WPAWN Then
'          If Not CBool(BAttack(Target) And PAttackBit) Then
'            If RankNum >= PawnsWMin(k) - 3 Then ' at most three squares behind some friendly pawn
'              If RankNum < PawnsWMin(k) Then  ' at most three squares behind some friendly pawn
'                If WAttack(Target) <> 0 Or BAttack(Target) = 0 Then r = r + 1
'              End If
'            End If
'          End If
'        End If
'      Next
'      For RankNum = 5 To 7 ' BLACK
'        Target = 10 + RankNum * 10 + k
'        If Board(Target) <> BPAWN Then
'          If Not CBool(WAttack(Target) And PAttackBit) Then
'            If RankNum <= PawnsBMax(k) + 3 Then ' at most three squares behind some friendly pawn
'             If RankNum > PawnsBMax(k) Then
'              If BAttack(Target) <> 0 Or WAttack(Target) = 0 Then rr = rr + 1
'             End If
'            End If
'          End If
'        End If
'      Next
'    Next
'
'    If r + rr <> 0 Then
'      ' weight for space
'      k = 0
'      For i = 1 To 8 ' count open files
'        If WPawns(i) + BPawns(i) = 0 Then k = k + 1
'      Next
'      If r > 0 Then
'       a = WNonPawnPieces + WPawnCnt - 2 * k
'       WPos.MG = WPos.MG + r * r * a \ 16
'      End If
'      If rr > 0 Then
'       a = BNonPawnPieces + BPawnCnt - 2 * k
'       BPos.MG = BPos.MG + rr * rr * a \ 16
'      End If
'    End If
'
'  End If
'

  '
  '--- Calculate weights and total eval
  '
  Dim TradeEval       As Long, PosEval As Long, PawnStructEval As Long
  Dim PassedPawnsEval As Long, MobilityEval As Long, KingSafetyEval As Long, ThreatEval As Long
  '--- evaluate_initiative() computes the initiative correction value for the
  '--- position, i.e., second order bonus/malus based on the known attacking/defending status of the players.
  r = Abs(WKingFile - BKingFile) - Abs(Rank(WKingLoc) - Rank(BKingLoc)) ' King distance
  k = 8 * (Abs(WSemiOpenFiles - BSemiOpenFiles) \ 12 + r - 15) + 12 * (WPawnCnt + BPawnCnt)
  rr = MatEval + (WPos.EG - BPos.EG) + (WPassed.EG - BPassed.EG) ' strong side?
  If rr > 0 Then
    WPos.EG = WPos.EG + GetMax(k, -Abs(rr \ 2))
  ElseIf rr < 0 Then
    BPos.EG = BPos.EG + GetMax(k, -Abs(rr \ 2))
  End If
  '--- Material Imbalance / Score trades
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
  '
  '--- Added all to eval score (SF based scaling:  Eval*100/SFPawnEndGameValue= 100 centipawns =1 pawn)
  '--- Example: Eval=240 => 1.00 pawn
  Eval = AllTotal.MG * GamePhase + AllTotal.EG * CLng(PHASE_MIDGAME - GamePhase) * ScaleFactor \ 64 '  * SF6 / 64=SCALE_FACTOR_NORMAL
  Eval = Eval \ PHASE_MIDGAME
  If bEvalTrace Then
    Debug.Print "Mat:" & EvalSFTo100(MatEval) & ", Mob:" & EvalSFTo100(MobilityEval) & ", KSafety:" & EvalSFTo100(KingSafetyEval) & ", Threat:" & EvalSFTo100(ThreatEval)
    Debug.Print "Total: " & EvalSFTo100(AllTotal.MG) & "," & EvalSFTo100(AllTotal.EG) & " = " & EvalSFTo100(ScaleScore(AllTotal))
    'Stop
  End If
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
    WriteTrace "Position: " & EvalSFTo100(PosEval) & "  => W" & ShowScore(WPos) & ", B" & ShowScore(BPos)
    WriteTrace "PawnStru: " & EvalSFTo100(PawnStructEval) & " => W" & ShowScore(WPawnStruct) & ", B" & ShowScore(BPawnStruct)
    WriteTrace "PassedPw: " & EvalSFTo100(PassedPawnsEval) & " => W" & ShowScore(WPassed) & ", B" & ShowScore(BPassed)
    WriteTrace "Mobility: " & EvalSFTo100(MobilityEval) & " => W(" & EvalSFTo100(WMobility.MG) & "," & EvalSFTo100(WMobility.EG) & "), B(" & EvalSFTo100(BMobility.MG) & "," & EvalSFTo100(BMobility.EG) & ")"
    WriteTrace "KSafety : " & EvalSFTo100(KingSafetyEval) & " => W(" & EvalSFTo100(WKSafety.MG) & "," & EvalSFTo100(WKSafety.EG) & "), B(" & EvalSFTo100(BKSafety.MG) & "," & EvalSFTo100(BKSafety.EG) & ")"
    WriteTrace "Threats : " & EvalSFTo100(ThreatEval) & " => W(" & EvalSFTo100(WThreat.MG) & "," & EvalSFTo100(WThreat.EG) & "), B(" & EvalSFTo100(BThreat.MG) & "," & EvalSFTo100(BThreat.EG) & ")"
    WriteTrace "Eval    : " & Eval & "  (" & EvalSFTo100(Eval) & "cp)"
    WriteTrace "-----------------"
    bTimeExit = True
  End If
  
  If Not bWhiteToMove Then Eval = -Eval '--- Invert for black
  
  'Eval = Eval + TEMPO_BONUS ' Tempo for side to move
  Eval = Eval + (16 + NonPawnMaterial \ ScoreKnight.MG \ 2)
  
  If Eval = DrawContempt Then Eval = Eval + 1
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
        If Not bWhiteToMove Then AddScore Score, ThreatenedByHangingPawn Else AddScoreVal Score, ThreatenedByHangingPawn.MG \ 2, ThreatenedByHangingPawn.EG \ 2 ' escape option?
      End If
    ElseIf Board(Square + SQ_UP_RIGHT) = BPAWN Then
      If Board(Square + SQ_UP_RIGHT + SQ_UP_LEFT) = BPAWN Or Board(Square + SQ_UP_RIGHT + SQ_UP_RIGHT) = BPAWN Then
        AddScore Score, ThreatBySafePawn(PieceType)
      Else
        If Not bWhiteToMove Then AddScore Score, ThreatenedByHangingPawn Else AddScoreVal Score, ThreatenedByHangingPawn.MG \ 2, ThreatenedByHangingPawn.EG \ 2 ' escape option?
      End If
    End If
  Else ' attack by white pawn?
    If Board(Square + SQ_DOWN_LEFT) = WPAWN Then
      If Board(Square + SQ_DOWN_LEFT + SQ_DOWN_LEFT) = WPAWN Or Board(Square + SQ_DOWN_LEFT + SQ_DOWN_RIGHT) = WPAWN Then
        AddScore Score, ThreatBySafePawn(PieceType)
      Else
        If bWhiteToMove Then AddScore Score, ThreatenedByHangingPawn Else AddScoreVal Score, ThreatenedByHangingPawn.MG \ 2, ThreatenedByHangingPawn.EG \ 2  ' escape option?
      End If
    ElseIf Board(Square + SQ_DOWN_RIGHT) = WPAWN Then
      If Board(Square + SQ_DOWN_RIGHT + SQ_DOWN_LEFT) = WPAWN Or Board(Square + SQ_DOWN_RIGHT + SQ_DOWN_RIGHT) = WPAWN Then
        AddScore Score, ThreatBySafePawn(PieceType)
      Else
        If bWhiteToMove Then AddScore Score, ThreatenedByHangingPawn Else AddScoreVal Score, ThreatenedByHangingPawn.MG \ 2, ThreatenedByHangingPawn.EG \ 2  ' escape option?
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

  For i = 1 To ThreatCnt

    With ThreatList(i)
      ' Add a bonus according to the kind of attacking pieces
      If .HangCol = COL_WHITE Then
        ' StronglyProtected: by pawn or by more defenders then attackers
        StronglyProtected = CBool(WAttack(.AttackedSquare) And PAttackBit) Or AttackBitCnt(WAttack(.AttackedSquare)) > AttackBitCnt(BAttack(.AttackedSquare))
        ' Non-pawn enemies strongly defended
        Defended = .HangPieceType <> PT_PAWN And StronglyProtected
        ' Enemies not strongly defended and under our attack
        Weak = Not StronglyProtected
        If Defended Or Weak Then
          If .AttackerPieceType = PT_BISHOP Or .AttackerPieceType = PT_KNIGHT Then
            AddScore BThreat, ThreatByMinor(.HangPieceType)
            If .HangPieceType <> PT_PAWN Then
              AddScoreVal BThreat, ThreatByRank.MG * Rank(.AttackedSquare), ThreatByRank.EG * Rank(.AttackedSquare)
            End If
          End If
        End If
        If (.HangPieceType = PT_QUEEN Or Weak) And .AttackerPieceType = PT_ROOK Then
          AddScore BThreat, ThreatByRook(.HangPieceType)
          If .HangPieceType <> PT_PAWN Then
            AddScoreVal BThreat, ThreatByRank.MG * Rank(.AttackedSquare), ThreatByRank.EG * Rank(.AttackedSquare)
          End If
        End If
        If Weak Then If WAttack(.AttackedSquare) = 0 Then AddScore BThreat, Hanging
      Else ' Black
        ' StronglyProtected: by pawn or by more defenders then attackers
        StronglyProtected = CBool(BAttack(.AttackedSquare) And PAttackBit) Or AttackBitCnt(BAttack(.AttackedSquare)) > AttackBitCnt(WAttack(.AttackedSquare))
        ' Non-pawn enemies strongly defended
        Defended = .HangPieceType <> PT_PAWN And StronglyProtected
        ' Enemies not strongly defended and under our attack
        Weak = Not StronglyProtected
        If Defended Or Weak Then
          If .AttackerPieceType = PT_BISHOP Or .AttackerPieceType = PT_KNIGHT Then
            AddScore WThreat, ThreatByMinor(.HangPieceType)
            If .HangPieceType <> PT_PAWN Then
              AddScoreVal WThreat, ThreatByRank.MG * (9 - Rank(.AttackedSquare)), ThreatByRank.EG * (9 - Rank(.AttackedSquare))
            End If
          End If
        End If
        If (.HangPieceType = PT_QUEEN Or Weak) And .AttackerPieceType = PT_ROOK Then
          AddScore WThreat, ThreatByRook(.HangPieceType)
          If .HangPieceType <> PT_PAWN Then
            AddScoreVal WThreat, ThreatByRank.MG * (9 - Rank(.AttackedSquare)), ThreatByRank.EG * (9 - Rank(.AttackedSquare))
          End If
        End If
        If Weak Then If BAttack(.AttackedSquare) = 0 Then AddScore WThreat, Hanging
      End If
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
 If AttackBit And BXrayAttackBit Then AddWKingAttack PT_BISHOP
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
 If AttackBit And BXrayAttackBit Then AddBKingAttack PT_BISHOP
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
          ConnectedBonus(Opposed, Phalanx, Support, r).EG = v * (r - 2) \ 4
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
  Safety = 258 ' MaxSafetyBonus
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

  WKingShelterStorm = Safety
End Function

Private Function BKingShelterStorm(ByVal ShelterKingLoc As Long) As Long
  Dim Center As Long, k As Long, r As Long, RelFile As Long, Safety As Long, RankUs As Long, RankThem As Long, RankNum As Long
  Safety = 258 ' MaxSafetyBonus
  '--- Pawn shelter
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

  BKingShelterStorm = Safety
End Function

Private Sub GetKingFlankFiles(ByVal KingLoc As Long, FileFrom As Long, FileTo As Long)

  Select Case File(KingLoc)
    Case 1 To 3: FileFrom = 1: FileTo = 4  ' File A-C
    Case 4 To 5: FileFrom = 3: FileTo = 6  ' File D-E
    Case 6 To 8: FileFrom = 6: FileTo = 8  ' File F-H
  End Select

End Sub

Public Function PinnedPieceW(ByVal PinnedLoc As Long, ByVal Direction As Long) As Boolean
  ' white pieces it threatend by pinned pieces and slider attack?
  Dim k As Long, sq As Long, Offset As Long, AttackBit As Long, Piece As Long
  PinnedPieceW = False
  If Direction < 4 Then ' Queen or rook orthogonal
    If Not CBool(BAttack(PinnedLoc) And QRAttackBit) Then Exit Function
    AttackBit = QRAttackBit
  Else ' Queen or bishop diagonal
    If Not CBool(BAttack(PinnedLoc) And QBAttackBit) Then Exit Function
    AttackBit = QBAttackBit
  End If
  Offset = QueenOffsets(Direction)

  For k = 1 To 8
    sq = PinnedLoc + Offset * k: Piece = Board(sq)
    If Piece = FRAME Then Exit For
    If Piece < NO_PIECE Then
      If Piece = BQUEEN Then PinnedPieceW = True: Exit Function
      If Piece = BROOK Then If Direction < 4 Then PinnedPieceW = True: Exit Function
      If Piece = BBISHOP Then If Direction >= 4 Then PinnedPieceW = True: Exit Function
      Exit For
    Else
      If Not (CBool(BAttack(sq) And AttackBit)) Then Exit For
    End If
  Next k

End Function

Public Function PinnedPieceB(ByVal PinnedLoc As Long, ByVal Direction As Long) As Boolean
  ' black pieces it threatend by pinned pieces and slider attack?
  Dim k As Long, sq As Long, Offset As Long, AttackBit As Long, Piece As Long
  PinnedPieceB = False
  If Direction < 4 Then ' Queen or rook orthogonal
    If Not CBool(WAttack(PinnedLoc) And QRAttackBit) Then Exit Function
    AttackBit = QRAttackBit
  Else ' Queen or bishop diagonal
    If Not CBool(WAttack(PinnedLoc) And QBAttackBit) Then Exit Function
    AttackBit = QBAttackBit
  End If
  Offset = QueenOffsets(Direction)

  For k = 1 To 8
    sq = PinnedLoc + Offset * k: Piece = Board(sq)
    If Piece = FRAME Then Exit For
    If Piece < NO_PIECE Then
      If Piece = WQUEEN Then PinnedPieceB = True: Exit Function
      If Piece = WROOK Then If Direction < 4 Then PinnedPieceB = True: Exit Function
      If Piece = WBISHOP Then If Direction >= 4 Then PinnedPieceB = True: Exit Function
      Exit For
    Else
      If Not (CBool(WAttack(sq) And AttackBit)) Then Exit For
    End If
  Next k

End Function

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
