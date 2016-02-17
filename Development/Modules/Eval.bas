Attribute VB_Name = "EvalBas"
'========================================
'=  EVAL : Evaluation of board position =
'========================================

Option Explicit
Const PHASE_MIDGAME = 128
Const PHASE_ENDGAME = 0

'SF6: Penalties for enemy's safe checks
Const QueenContactCheck           As Integer = 89
Const RookContactCheck            As Integer = 71
Const QueenCheck                  As Integer = 50
Const RookCheck                   As Integer = 37
Const BishopCheck                 As Integer = 6
Const KnightCheck                 As Integer = 14

Public bTimeExit                  As Boolean
Public StartThinkingTime          As Single
Public TimeStart                  As Single
Public SearchStart                As Single
Public SearchTime                 As Single
Public TimeForIteration           As Single
Public ExtraTimeForMove           As Single
Public TimeLeft                   As Single
Public OpponentTime               As Single
Public TimeIncrement              As Long
Public MovesToTC                  As Long
Public SecondsPerGame             As Long
Public FixedDepth                 As Integer  '=NO_FIXED_DEPTH if time limit is used
Public FixedTime                  As Single
Public LastChangeDepth            As Integer, LastChangeMove As String, bExtraTime As Boolean
Public TotalTimeGiven             As Single
Public bAddExtraTime              As Boolean
Public bResearching               As Boolean '--- out of aspiration windows: more time

Public BestMoveChanges            As Single ' More time if best move changes often
Public UnstablePvFactor           As Single
Public MaximumTime                As Single
Public OptimalTime                As Single

Public DoubledPenalty(8)          As TScore
Public IsolatedPenalty(1, 8)      As TScore
Public IsolatedNotPassed          As TScore
Public BackwardPenalty(1)         As TScore
Public ConnectedBonus(1, 1, 1, 8) As TScore
Public ShelterWeakness(4, 8)      As Integer
Public StormDanger(4, 4, 8)       As Integer

Public ThreatenedByHangingPawn    As TScore
Public Hanging                    As TScore
Public Checked                    As TScore

Public ValueP                     As Integer
Public ValueN                     As Integer
Public ValueB                     As Integer
Public ValueR                     As Integer
Public ValueQ                     As Integer

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

Public WKRelLoc                   As Integer, BKRelLoc As Integer ' relative king location for king ring: file A>B, file H<G

Dim WAttack(MAX_BOARD)            As Integer   '- Fields around king: count attacks ' public is 2x faster than local+Erase in Eval function !
Dim BAttack(MAX_BOARD)            As Integer   '- Fields around king: count attacks
Dim WThreat                       As TScore, BThreat As TScore
 
Public PiecePosScaleFactor        As Long  ' set in INI file
Public CompKingDefScaleFactor     As Long ' set in INI file
Public OppKingAttScaleFactor      As Long ' set in INI file
Public PawnStructScaleFactor      As Long ' set in INI file
Public PassedPawnsScaleFactor     As Long ' set in INI file
Public MobilityScaleFactor        As Long ' set in INI file
Public ThreatsScaleFactor         As Long ' set in INI file

Public WKingScaleFactor           As Long, BKingScaleFactor As Long

Public PawnsWMax(9)               As Integer  '--- Pawn max rank (2-7) for file A-H
Public PawnsWMin(9)               As Integer  '--- Pawn min rank (2-7) for file A-H
Public WPawns(9)                  As Integer  '--- number of pawns for file A-H
Public PawnsBMax(9)               As Integer
Public PawnsBMin(9)               As Integer
Public BPawns(9)                  As Integer

Public RootMove                   As TMove
Public LastNodesCnt               As Long

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
Public PsqVal(1, 16, MAX_BOARD) ' piece square score for piece piece: (endgame,piece,square)

'--- Mobility values for pieces
Public MobilityN(9)               As TScore
Public MobilityB(15)              As TScore
Public MobilityR(15)              As TScore
Public MobilityQ(29)              As TScore

Public ThreatenedByPawn(5)        As TScore
Public OutpostBonusKnight(1)      As TScore
Public OutpostBonusBishop(1)      As TScore

Public KingAttackWeights(6)       As Integer
'--- Counter for piece types
Public WPawnCnt                   As Integer
Public BPawnCnt                   As Integer
Public WQueenCnt                  As Integer
Public BQueenCnt                  As Integer
Public WRookCnt                   As Integer
Public BRookCnt                   As Integer
Public WBishopCnt                 As Integer
Public BBishopCnt                 As Integer
Public WKnightCnt                 As Integer
Public BKnightCnt                 As Integer

Public KingDanger(512)            As TScore ' Lookup table for King danger scoring

Public WBestPawnVal               As Long, BBestPawnVal As Long, WBestPawn As Integer, BBestPawn As Integer
Public GamePhase                  As Long

Public MaxPosCore                 As Long, MaxKsScore As Long '--- for debug scaling
Public WKingAttackersWeight       As Integer, WKingAttackersCount As Integer, BKingAttackersWeight As Integer, BKingAttackersCount As Integer
Public WKingDefendersWeight       As Integer, WKingDefendersCount As Integer, BKingDefendersWeight As Integer, BKingDefendersCount As Integer

Public bEvalTrace                 As Boolean
Public bTimeTrace                 As Boolean
Public bHashTrace                 As Boolean
Public bWinboardTrace             As Boolean

Dim PassedPawns(16)               As Integer ' List of passed pawns (Square)
Dim PassedPawnsCnt                As Integer
Dim WPassedPawnAttack             As Integer, BPassedPawnAttack As Integer

Public Enum enumWeight
  Mobility_Weight = 1
  PawnStructure_Weight = 2
  PassedPawns_Weight = 3
  Space_Weight = 4
  KingSafety_Weight = 5
  Threats_Weight = 6
End Enum

Public Weights(6) As TScore

Public PushClose(8) As Integer
Public PushAway(8) As Integer
Public PushToEdges(MAX_BOARD) As Integer

'--- Threat list
Dim ThreatCnt     As Integer

Public Type TThreatList
  HangCol As enumColor
  HangPieceType     As Integer
  AttPieceType    As Integer
  AttackingSquare As Integer
  AttackedSquare As Integer
End Type
Dim ThreatList(32)            As TThreatList

' Pawn Eval
Dim Passed                    As Boolean, Isolated As Boolean, Opposed As Boolean, Backward As Boolean
Dim Neighbours                As Boolean, Doubled As Boolean, Lever As Boolean, Connected As Boolean, Supported As Boolean, Phalanx As Boolean
Public PassedPawnFileBonus(8) As TScore
Public PassedPawnRankBonus(8) As TScore

' Threats
Public ThreatDefendedMinor(6) As TScore ' Attacker is defended minor (B/N)
Public ThreatDefendedMajor(6) As TScore
Public ThreatWeakMinor(6)     As TScore
Public ThreatWeakMajor(6)     As TScore
Public KingOnOneBonus         As TScore
   
' Material imbalance (SF6)
Public Linear(5)              As Integer
Public QuadraticOurs(5, 5)    As Integer
Public QuadraticTheirs(5, 5)  As Integer
Public ImbPieceCount(2, 5)    As Integer
   
Private bWIsland As Boolean, bBIsland As Boolean
' temp
Private bIniReadDone          As Boolean

'---------------------------------------------------------------------------
'InitEval(ThreatMove)  Set piece values and piece square tables
'---------------------------------------------------------------------------
Public Sub InitEval()

  Dim Score As Long, bSaveEvalTrace As Boolean

  '--- Limit  high eval values ( VERY important for playing style!)
  If Not bIniReadDone Then
    bIniReadDone = True
    '--- Default used if INI file is missing
    PiecePosScaleFactor = Val(ReadINISetting("POSITION_FACTOR", "90"))
    MobilityScaleFactor = Val(ReadINISetting("MOBILITY_FACTOR", "90"))
    PawnStructScaleFactor = Val(ReadINISetting("PAWNSTRUCT_FACTOR", "90"))
    PassedPawnsScaleFactor = Val(ReadINISetting("PASSEDPAWNS_FACTOR", "130"))
    ThreatsScaleFactor = Val(ReadINISetting("THREATS_FACTOR", "250"))
    OppKingAttScaleFactor = Val(ReadINISetting("OPPKINGATT_FACTOR", "80"))
    CompKingDefScaleFactor = Val(ReadINISetting("COMPKINGDEF_FACTOR", "90"))
    
    '
    '--- Piece values  MG=midgame / EG=endgame
    '--- SF6 values  ( scale to centipawns: \256 )
    '
    ScorePawn.MG = Val(ReadINISetting("PAWN_VAL_MG", "198"))
    ScorePawn.EG = Val(ReadINISetting("PAWN_VAL_EG", "258"))
  
    ScoreKnight.MG = Val(ReadINISetting("KNIGHT_VAL_MG", "817"))
    ScoreKnight.EG = Val(ReadINISetting("KNIGHT_VAL_EG", "846"))
  
    ScoreBishop.MG = Val(ReadINISetting("BISHOP_VAL_MG", "836"))
    ScoreBishop.EG = Val(ReadINISetting("BISHOP_VAL_EG", "857"))
  
    ScoreRook.MG = Val(ReadINISetting("ROOK_VAL_MG", "1270"))
    ScoreRook.EG = Val(ReadINISetting("ROOK_VAL_EG", "1278"))
  
    ScoreQueen.MG = Val(ReadINISetting("QUEEN_VAL_MG", "2521"))
    ScoreQueen.EG = Val(ReadINISetting("QUEEN_VAL_EG", "2558"))
  
    MidGameLimit = Val(ReadINISetting("MIDGAME_LIMIT", "15581")) ' for game phase
    EndgameLimit = Val(ReadINISetting("ENDGAME_LIMIT", "3998"))  ' for game phase
  
    ' Draw contempt in centipawns > scale to SF (needs ScorePawn.EG set)
    DrawContempt = Val(ReadINISetting(CONTEMPT_KEY, "1"))
    DrawContempt = Eval100ToSF(DrawContempt) ' in centipawns

  End If

  '--- Detect endgame stage ---
  bSaveEvalTrace = bEvalTrace: bEvalTrace = False ' Save trace setting, trace not needed here before init done
  Score = Eval() ' Set material,NonPawnMaterial for GamePhase calculation
  bEvalTrace = bSaveEvalTrace

  SetGamePhase NonPawnMaterial ' Set GamePhase, PieceValues, bEndGame

  '--- Pawn values needed, so init here
  'InitRecaptureMargins ' no longer used
  InitFutilityMoveCounts
  InitReductionArray

  InitConnectedPawns

End Sub

Public Sub InitPieceValue()
  '--- Piece values, always absolut, positive value
  PieceAbsValue(FRAME) = 0
  PieceAbsValue(WPAWN) = ValueP: PieceAbsValue(BPAWN) = ValueP
  PieceAbsValue(WKNIGHT) = ValueN: PieceAbsValue(BKNIGHT) = ValueN
  PieceAbsValue(WBISHOP) = ValueB: PieceAbsValue(BBISHOP) = ValueB
  PieceAbsValue(WROOK) = ValueR: PieceAbsValue(BROOK) = ValueR
  PieceAbsValue(WQUEEN) = ValueQ: PieceAbsValue(BQUEEN) = ValueQ
  PieceAbsValue(WKING) = 5000: PieceAbsValue(BKING) = 5000
  PieceAbsValue(13) = 0: PieceAbsValue(14) = 0
  PieceAbsValue(WEP_PIECE) = ValueP: PieceAbsValue(BEP_PIECE) = ValueP

  '--- Piece SCore: positive for White, negative for Black
  PieceScore(FRAME) = 0
  PieceScore(WPAWN) = ValueP: PieceScore(BPAWN) = -ValueP
  PieceScore(WKNIGHT) = ValueN: PieceScore(BKNIGHT) = -ValueN
  PieceScore(WBISHOP) = ValueB: PieceScore(BBISHOP) = -ValueB
  PieceScore(WROOK) = ValueR: PieceScore(BROOK) = -ValueR
  PieceScore(WQUEEN) = ValueQ: PieceScore(BQUEEN) = -ValueQ
  PieceScore(WKING) = 5000: PieceScore(BKING) = -5000
  PieceScore(13) = 0: PieceScore(14) = 0
  PieceScore(WEP_PIECE) = ValueP: PieceScore(BEP_PIECE) = -ValueP

  PieceTypeValue(PT_PAWN) = ValueP
  PieceTypeValue(PT_KNIGHT) = ValueN
  PieceTypeValue(PT_BISHOP) = ValueB
  PieceTypeValue(PT_ROOK) = ValueR
  PieceTypeValue(PT_QUEEN) = ValueQ
  PieceTypeValue(PT_KING) = 5000
  
End Sub

Public Function SetGamePhase(ByVal NonPawnMaterial As Long) As Integer
  Static OldNonPawnMaterial As Long
  
  NonPawnMaterial = GetMax(EndgameLimit, GetMin(NonPawnMaterial, MidGameLimit))
  GamePhase = (((NonPawnMaterial - EndgameLimit) * PHASE_MIDGAME) / (MidGameLimit - EndgameLimit))
  bEndgame = (GamePhase <= PHASE_ENDGAME)
  
  If NonPawnMaterial <> OldNonPawnMaterial Or ValueP = 0 Then
    
    ValueP = ScaleScore(ScorePawn)
    ValueN = ScaleScore(ScoreKnight)
    ValueB = ScaleScore(ScoreBishop)
    ValueR = ScaleScore(ScoreRook)
    ValueQ = ScaleScore(ScoreQueen)
    
    InitPieceValue
    OldNonPawnMaterial = NonPawnMaterial
  End If
  
End Function

'---------------------------------------------------------------------------------------------------
'---  Eval() - Evaluation of position
'---           Returns value from view of side to move (positive if black to move and black is better)
'---           Value scale to stockfish pawn endgame value (258 = 1 pawn)
'---
'---  Steps:
'---         1. Loop over all pieces to fill pawn structure array, pawn threats,
'---            calculate material + piece square values
'---         2. Check material draw
'---         3. Loop over all pieces: evaluate each piece except kings.
'---            does a move generation to calculate mobility, attackers, defenders
'---         4. calculate king safety ( shelter, pawn storm, attacks )
'---         5. calculate trapped bishops, passed pawns, king distance to best pawn, center control
'---         6. calculate threats
'---         7. Add all evalution terms weighted by variables set in INI file:
'---             Material + Position(general) + PawnStructure + PassedPawns + Mobility +
'---             KingSafetyComputer + KingSafetyOpponent + Threats
'---         8. Add tempo value for side to move
'---         9. invert score for black to move and return evaluation value
'---------------------------------------------------------------------------------------------------
Public Function Eval() As Long

  Dim a                            As Integer, i As Integer, ForkCnt As Integer, SC As TScore, Safety As Long, bKingAtt As Boolean
  Dim WPos                         As TScore, BPos As TScore, WPassed As TScore, BPassed As TScore, WMobility As TScore, BMobility As TScore
  Dim WPawnStruct                  As TScore, BPawnStruct As TScore, Piece As Integer
  Dim WKSafety                     As TScore, BKSafety As TScore, bDoWKSafety As Boolean, bDoBKSafety As Boolean
  Dim KingAdjacentZoneAttacksCount As Integer, WKnightChecks As Integer, BKnightChecks As Integer, WBishopChecks As Integer, BBishopChecks As Integer
  Dim WRookChecks                  As Integer, BRookChecks As Integer, WQueenChecks As Integer, BQueenChecks As Integer
  Dim WAttPotential                As Long, BAttPotential As Long, AttackUnits As Integer, Undefended As Integer
  Dim Square                       As Integer, FileNum As Integer, MinWKingPawnDistance As Integer, MinBKingPawnDistance As Integer
  Dim RankNum                      As Integer, RelRank As Integer, Target As Integer, Offset As Integer, MobCnt As Integer
  Dim DefByPawn                    As Integer, AttByPawn As Integer, bAllDefended As Boolean, BlockSqDefended As Boolean, WPinnedCnt As Integer, BPinnedCnt As Integer
  Dim RankPath                     As Integer, sq As Integer, IsPossibleOutpost As Boolean

  Dim BlockSq                      As Integer, r As Integer, rr As Integer, MBonus As Long, EBonus As Long, k As Integer
  Dim OwnCol                       As Integer, OppCol As Integer, MoveUp As Integer, RelRank8 As Integer, OwnKingLoc As Integer, OppKingLoc As Integer, UnsafeSq As Boolean, BlockSqUnsafe As Boolean, bAttackedFromBehind As Boolean, bDefendedFromBehind As Boolean
  Dim WBishopsOnBlackSq             As Integer, WBishopsOnWhiteSq As Integer, BBishopsOnBlackSq As Integer, BBishopsOnWhiteSq As Integer
  Dim WPawnCntOnWhiteSq As Integer, WPawnCntOnBlackSq As Integer, BPawnCntOnWhiteSq As Integer, BPawnCntOnBlackSq As Integer

  EvalCnt = EvalCnt + 1

  If bEvalTrace Then WriteTrace "------- Start Eval ------"

  WBestPawnVal = UNKNOWN_SCORE: WBestPawn = 0
  BBestPawnVal = UNKNOWN_SCORE: BBestPawn = 0

  ThreatCnt = 0: ScoreToZero WThreat: ScoreToZero BThreat

  '--- Fill Pawn Arrays
  For a = 0 To 9
    WPawns(a) = 0: BPawns(a) = 0: PawnsWMin(a) = 9: PawnsWMax(a) = 0: PawnsBMin(a) = 9: PawnsBMax(a) = 0
  Next
  WPawns(0) = -1: BPawns(0) = -1
  WPawns(9) = -1: BPawns(9) = -1
  PassedPawnsCnt = 0

  WPawnCnt = 0: WBishopCnt = 0: WKnightCnt = 0: WRookCnt = 0: WQueenCnt = 0: WQueenLoc = 0
  BPawnCnt = 0: BBishopCnt = 0: BKnightCnt = 0: BRookCnt = 0: BQueenCnt = 0: BQueenLoc = 0

  Erase WAttack(): Erase BAttack() 'Init attack arrays  (fast)

  MinWKingPawnDistance = 9: MinBKingPawnDistance = 9
  WKingAttackersCount = 0: WKingAttackersWeight = 0: BKingAttackersCount = 0: BKingAttackersWeight = 0
  WKingDefendersCount = 0: WKingDefendersWeight = 0: BKingDefendersCount = 0: BKingDefendersWeight = 0

  Eval = 0

  '--- 1. loop over pieces: count pieces for material totals and game phase calculation. add piece square table score.
  '----                     calc pawn min/max rank positions per file; pawn attacks(for mobility used later)
  For a = 1 To NumPieces
    Square = Pieces(a): If Square = 0 Or Board(Square) = NO_PIECE Then GoTo lblNextPieceCnt
    Select Case Board(Square)
      Case WPAWN
        WPos.MG = WPos.MG + PsqtWP(Square).MG: WPos.EG = WPos.EG + PsqtWP(Square).EG
        WAttack(Square + SQ_UP_LEFT) = PAttackBit: WAttack(Square + SQ_UP_RIGHT) = PAttackBit  ' Set pawn attack here for use in pieces eval
        FileNum = File(Square): RankNum = Rank(Square): WPawnCnt = WPawnCnt + 1
        WPawns(FileNum) = WPawns(FileNum) + 1
        If RankNum < PawnsWMin(FileNum) Then PawnsWMin(FileNum) = RankNum
        If RankNum > PawnsWMax(FileNum) Then PawnsWMax(FileNum) = RankNum
        If MaxDistance(WKingLoc, Square) < MinWKingPawnDistance Then MinWKingPawnDistance = MaxDistance(WKingLoc, Square)
        If ColorSq(Square) = COL_WHITE Then WPawnCntOnWhiteSq = WPawnCntOnWhiteSq + 1 Else WPawnCntOnBlackSq = WPawnCntOnBlackSq + 1 ' for Bishop eval
      Case BPAWN
        BPos.MG = BPos.MG + PsqtBP(Square).MG: BPos.EG = BPos.EG + PsqtBP(Square).EG
        BAttack(Square + SQ_DOWN_LEFT) = PAttackBit: BAttack(Square + SQ_DOWN_RIGHT) = PAttackBit
        FileNum = File(Square): RankNum = Rank(Square): BPawnCnt = BPawnCnt + 1
        BPawns(FileNum) = BPawns(FileNum) + 1
        If RankNum < PawnsBMin(FileNum) Then PawnsBMin(FileNum) = RankNum
        If RankNum > PawnsBMax(FileNum) Then PawnsBMax(FileNum) = RankNum
        If MaxDistance(BKingLoc, Square) < MinBKingPawnDistance Then MinBKingPawnDistance = MaxDistance(BKingLoc, Square)
        If ColorSq(Square) = COL_WHITE Then BPawnCntOnWhiteSq = BPawnCntOnWhiteSq + 1 Else BPawnCntOnBlackSq = BPawnCntOnBlackSq + 1 ' for Bishop eval
      Case WKING, BKING: ' needed for better performance in P-Code
      Case WBISHOP: WBishopCnt = WBishopCnt + 1: If ColorSq(Square) = COL_WHITE Then WBishopsOnWhiteSq = WBishopsOnWhiteSq + 1 Else WBishopsOnBlackSq = WBishopsOnBlackSq + 1
        WPos.MG = WPos.MG + PsqtWB(Square).MG: WPos.EG = WPos.EG + PsqtWB(Square).EG: AddPawnThreat BThreat, COL_WHITE, PT_BISHOP, Square
      Case BBISHOP: BBishopCnt = BBishopCnt + 1: If ColorSq(Square) = COL_WHITE Then BBishopsOnWhiteSq = BBishopsOnWhiteSq + 1 Else BBishopsOnBlackSq = BBishopsOnBlackSq + 1
        BPos.MG = BPos.MG + PsqtBB(Square).MG: BPos.EG = BPos.EG + PsqtBB(Square).EG: AddPawnThreat WThreat, COL_BLACK, PT_BISHOP, Square
      Case WKNIGHT: WKnightCnt = WKnightCnt + 1
        WPos.MG = WPos.MG + PsqtWN(Square).MG: WPos.EG = WPos.EG + PsqtWN(Square).EG: AddPawnThreat BThreat, COL_WHITE, PT_KNIGHT, Square
      Case BKNIGHT: BKnightCnt = BKnightCnt + 1
        BPos.MG = BPos.MG + PsqtBN(Square).MG: BPos.EG = BPos.EG + PsqtBN(Square).EG: AddPawnThreat WThreat, COL_BLACK, PT_KNIGHT, Square
      Case WROOK: WRookCnt = WRookCnt + 1
        WPos.MG = WPos.MG + PsqtWR(Square).MG: WPos.EG = WPos.EG + PsqtWR(Square).EG: AddPawnThreat BThreat, COL_WHITE, PT_ROOK, Square
      Case BROOK: BRookCnt = BRookCnt + 1
        BPos.MG = BPos.MG + PsqtBR(Square).MG: BPos.EG = BPos.EG + PsqtBR(Square).EG: AddPawnThreat WThreat, COL_BLACK, PT_ROOK, Square
      Case WQUEEN: WQueenCnt = WQueenCnt + 1: WQueenLoc = Square
        WPos.MG = WPos.MG + PsqtWQ(Square).MG: WPos.EG = WPos.EG + PsqtWQ(Square).EG: AddPawnThreat BThreat, COL_WHITE, PT_QUEEN, Square
      Case BQUEEN: BQueenCnt = BQueenCnt + 1: BQueenLoc = Square
        BPos.MG = BPos.MG + PsqtBQ(Square).MG: BPos.EG = BPos.EG + PsqtBQ(Square).EG: AddPawnThreat WThreat, COL_BLACK, PT_QUEEN, Square
    End Select
lblNextPieceCnt:
  Next

  WNonPawnMaterial = WQueenCnt * ScoreQueen.MG + WRookCnt * ScoreRook.MG + WBishopCnt * ScoreBishop.MG + WKnightCnt * ScoreKnight.MG
  WMaterial = WNonPawnMaterial + WPawnCnt * ScorePawn.MG

  BNonPawnMaterial = BQueenCnt * ScoreQueen.MG + BRookCnt * ScoreRook.MG + BBishopCnt * ScoreBishop.MG + BKnightCnt * ScoreKnight.MG
  BMaterial = BNonPawnMaterial + BPawnCnt * ScorePawn.MG

  NonPawnMaterial = WNonPawnMaterial + BNonPawnMaterial
  Material = WMaterial - BMaterial

  SetGamePhase NonPawnMaterial

 
  '--- Endgame function available?
  Select Case WPawnCnt + BPawnCnt
  Case 0 ' no pawns
    ' KQKR
    If (WMaterial = ScoreQueen.MG And BMaterial = ScoreRook.MG) Or (BMaterial = ScoreQueen.MG And WMaterial = ScoreRook.MG) Then
       Eval = Eval_KQKR(): GoTo lblEndEval
    End If
    
    '--- Insufficent material draw?
    If IsMaterialDraw() Then Eval = 0: Exit Function '- Endgame draw: not sufficent material for mate
  
  Case 1 ' one pawn
    
    If (WMaterial = ScoreRook.MG And BMaterial = ScorePawn.MG) Or (BMaterial = ScoreRook.MG And WMaterial = ScorePawn.MG) Then
       Eval = Eval_KRKP(): GoTo lblEndEval ' KRKP
    ElseIf (WMaterial = ScoreQueen.MG And BMaterial = ScorePawn.MG) Or (BMaterial = ScoreQueen.MG And WMaterial = ScorePawn.MG) Then
       Eval = Eval_KQKP(): GoTo lblEndEval ' KQKP
    End If
  End Select


  bDoWKSafety = CBool(BNonPawnMaterial >= ScoreQueen.MG)
  bDoBKSafety = CBool(WNonPawnMaterial >= ScoreQueen.MG)

  '--- Attack potential
  WAttPotential = WNonPawnMaterial \ 100
  BAttPotential = BNonPawnMaterial \ 100

  ' Init arrays for checks
  If bDoWKSafety Then FillKingCheckW
  If bDoBKSafety Then FillKingCheckB

  'Init relative king ring center loc: H1>G1
  Select Case File(WKingLoc)
    Case 1: WKRelLoc = WKingLoc + 1
    Case 8: WKRelLoc = WKingLoc - 1
    Case Else: WKRelLoc = WKingLoc
  End Select

  Select Case File(BKingLoc)
    Case 1: BKRelLoc = BKingLoc + 1
    Case 8: BKRelLoc = BKingLoc - 1
    Case Else: BKRelLoc = BKingLoc
  End Select

  '--- Double Bishop Bonus ( now included in SF6 Imbalance )
  'If WBishopCnt = 2 Then AddScoreVal WPos, 60 + (8 - WPawnCnt) * 6, 60
  'If BBishopCnt = 2 Then AddScoreVal BPos, 60 + (8 - BPawnCnt) * 6, 60

  '--- King Position
  ScoreToZero WKSafety: ScoreToZero BKSafety
  If WNonPawnMaterial > 0 And BMaterial = 0 Then
    WPos.EG = WPos.EG + (7 - MaxDistance(BKingLoc, WKingLoc)) * 12 ' follow opp king to edge for mate (KRK, KQK)
    BPos.EG = BPos.EG + PsqtBK(BKingLoc).EG
  ElseIf BNonPawnMaterial > 0 And WMaterial = 0 Then
    BPos.EG = BPos.EG + (7 - MaxDistance(WKingLoc, BKingLoc)) * 12
    WPos.EG = WPos.EG + PsqtWK(WKingLoc).EG
  Else
    AddScore WPos, PsqtWK(WKingLoc)
    AddScore BPos, PsqtWK(WKingLoc)
  End If

  '--------------------------------------------------------------------
  '--- EVAL Loop over pieces ------------------------------------------
  '--------------------------------------------------------------------

  For a = 1 To NumPieces
    Square = Pieces(a): If Square = 0 Or Board(Square) = NO_PIECE Then GoTo lblNextPiece
    FileNum = File(Square): RankNum = Rank(Square): SC.MG = 0: SC.EG = 0: bKingAtt = False
    If Board(Square) Mod 2 = 1 Then
      ' White piece
      RelRank = RankNum
    Else
      ' Black piece
      RelRank = (9 - RankNum)
    End If

    Select Case Board(Square)
      Case WPAWN  '---- WHITE PAWN ------------------------------------
        DefByPawn = 0: AttByPawn = 0
        
        If Board(Square + SQ_DOWN_LEFT) = WPAWN Then DefByPawn = DefByPawn + 1
        If Board(Square + SQ_DOWN_RIGHT) = WPAWN Then DefByPawn = DefByPawn + 1
        If Board(Square + SQ_UP_LEFT) = BPAWN Then AttByPawn = AttByPawn + 1
        If Board(Square + SQ_UP_RIGHT) = BPAWN Then AttByPawn = AttByPawn + 1
        If MaxDistance(Square, BKingLoc) <= 2 Then If Abs(File(Square) - FileNum) <= 1 Then BKingAttackersCount = BKingAttackersCount + 1
        If bEndgame And RankNum > 4 Then If MaxDistance(Square, WKingLoc) = 1 Then SC.EG = SC.EG + 10 ' advanced pawn supported by king
        
        Neighbours = (WPawns(FileNum + 1) > 0 Or WPawns(FileNum - 1) > 0)
        Doubled = (WPawns(FileNum) > 1) And RankNum < PawnsWMax(FileNum)
        Opposed = (BPawns(FileNum) > 0) And RankNum < PawnsBMax(FileNum)
        Passed = (Not Opposed And PawnsBMax(FileNum + 1) <= RankNum And PawnsBMax(FileNum - 1) <= RankNum)
        Lever = (AttByPawn > 0)
        Isolated = Not Neighbours
        Phalanx = (Board(Square + SQ_LEFT) = WPAWN Or Board(Square + SQ_RIGHT) = WPAWN)
        Supported = (DefByPawn > AttByPawn)
        Connected = (Supported Or Phalanx)
 
        If (Passed Or Isolated Or Lever Or Connected) Or RelRank >= 5 Then
          Backward = False
        Else
          If Board(Square + SQ_UP) = BPAWN Then
            Backward = True
          Else
            Backward = (PawnsWMin(FileNum + 1) > RankNum And PawnsWMin(FileNum - 1) > RankNum)
          End If
        End If

        If Isolated Then MinusScore SC, IsolatedPenalty(Abs(Opposed), FileNum)
        If Backward Then MinusScore SC, BackwardPenalty(Abs(Opposed))
        If Not Supported Then
          If Board(Square + SQ_UP_LEFT) = WPAWN And Board(Square + SQ_UP_RIGHT) = WPAWN Then
            SC.MG = SC.MG - 25: SC.EG = SC.EG - 15 ' Unsupported pawn penalty by twice supporting
          Else
            SC.MG = SC.MG - 20: SC.EG = SC.EG - 10 ' Unsupported pawn penalty
          End If
        End If
        If Connected Then AddScore SC, ConnectedBonus(Abs(Opposed), Abs(Not Phalanx), Abs(DefByPawn > 1), RelRank)
        If Doubled Then
          r = GetMax(1, MaxDistance(Square, PawnsWMax(FileNum)))
          SC.MG = SC.MG - DoubledPenalty(FileNum).MG \ r
          SC.EG = SC.EG - DoubledPenalty(FileNum).EG \ r
        End If
        
        If Lever Then
          If Supported Then
            If RelRank = 4 Then
              SC.MG = SC.MG + 10
            ElseIf RelRank = 5 Then
              SC.MG = SC.MG + 40: SC.EG = SC.EG + 20
            ElseIf RelRank = 6 Then
              SC.MG = SC.MG + 80: SC.EG = SC.EG + 40
            End If
          Else
            If RelRank = 5 Then
              SC.MG = SC.MG + 20: SC.EG = SC.EG + 10
            ElseIf RelRank = 6 Then
              SC.MG = SC.MG + 40: SC.EG = SC.EG + 20
            End If
          End If
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
        Else
          If RelRank <= 5 Then
            If Board(Square + 19) = BPAWN Or Board(Square + 21) = BPAWN Then SC.MG = SC.MG + 25
            ' Blocked E/D pawn?
            If Square = 34 Or Square = 35 Then
              If Board(Square + SQ_UP) < NO_PIECE Then SC.MG = SC.MG - 15
            End If
          End If
        End If
        
        ' bonus if safe pawn push attacks an enemy piece
        If Board(Square + SQ_UP) = NO_PIECE Then
          If BAttack(Board(Square + SQ_UP)) = 0 Or CBool(WAttack(Square + SQ_UP) And PAttackBit) Then   ' not attacked or defend by own pawn
            SC.MG = SC.MG + 8: SC.EG = SC.EG + 8 ' Safe pawn push
            For i = 19 To 21 Step 2
              r = Board(Square + i)
              If PieceColor(r) = COL_BLACK Then
                If r = BPAWN Then
                  SC.MG = SC.MG + 15: SC.EG = SC.EG + 15 ' pawn threats enemy pawn
                Else
                  SC.MG = SC.MG + 30: SC.EG = SC.EG + 30 ' pawn threats non pawn enemy
                End If
              End If
            Next i
          End If
        End If

        ' Passed : eval later when full attack is available
        If Passed And Not Doubled Then
           PassedPawnsCnt = PassedPawnsCnt + 1: PassedPawns(PassedPawnsCnt) = Square
        ElseIf Isolated And Not Opposed Then
          MinusScore SC, IsolatedNotPassed
        End If
        
        AddScore WPawnStruct, SC
        If bEvalTrace Then WriteTrace "WPawn: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG
    
      Case BPAWN  '---- BLACK PAWN ------------------------------------
        DefByPawn = 0: AttByPawn = 0
        If Board(Square + SQ_DOWN_LEFT) = WPAWN Then AttByPawn = AttByPawn + 1
        If Board(Square + SQ_DOWN_RIGHT) = WPAWN Then AttByPawn = AttByPawn + 1
        If Board(Square + SQ_UP_LEFT) = BPAWN Then DefByPawn = DefByPawn + 1
        If Board(Square + SQ_UP_RIGHT) = BPAWN Then DefByPawn = DefByPawn + 1
        If MaxDistance(Square, WKingLoc) <= 2 Then If Abs(File(Square) - FileNum) <= 1 Then WKingAttackersCount = WKingAttackersCount + 1
        If bEndgame And RelRank > 4 Then If MaxDistance(Square, BKingLoc) = 1 Then SC.EG = SC.EG + 10  ' advanced pawn supported by king

        Neighbours = (BPawns(FileNum + 1) > 0 Or BPawns(FileNum - 1) > 0)
        Doubled = (BPawns(FileNum) > 1) And RankNum > PawnsBMin(FileNum)
        Opposed = (WPawns(FileNum) > 0) And RankNum > PawnsWMin(FileNum)
        Passed = (Not Opposed And PawnsWMin(FileNum + 1) >= RankNum And PawnsWMin(FileNum - 1) >= RankNum)
        Lever = (AttByPawn > 0)
        Isolated = Not Neighbours
        Phalanx = (Board(Square + SQ_LEFT) = BPAWN Or Board(Square + SQ_RIGHT) = BPAWN)
        Supported = (DefByPawn > AttByPawn)
        Connected = (Supported Or Phalanx)

        If (Passed Or Isolated Or Lever Or Connected) Or RelRank >= 5 Then
          Backward = False
        Else
          If Board(Square + SQ_DOWN) = WPAWN Then
            Backward = True
          Else
            Backward = (PawnsBMax(FileNum + 1) < RankNum And PawnsBMax(FileNum - 1) < RankNum)
          End If
        End If

        If Isolated Then MinusScore SC, IsolatedPenalty(Abs(Opposed), FileNum)
        If Backward Then MinusScore SC, BackwardPenalty(Abs(Opposed))
        If Not Supported Then
          If Board(Square + SQ_DOWN_LEFT) = BPAWN And Board(Square + SQ_DOWN_RIGHT) = BPAWN Then
            SC.MG = SC.MG - 25: SC.EG = SC.EG - 15 ' Unsupported pawn penalty by twice supporting
          Else
            SC.MG = SC.MG - 20: SC.EG = SC.EG - 10 ' Unsupported pawn penalty
          End If
        End If
        If Connected Then AddScore SC, ConnectedBonus(Abs(Opposed), Abs(Not Phalanx), Abs(DefByPawn > 1), RelRank)
        If Doubled Then
          r = GetMax(1, MaxDistance(Square, PawnsBMin(FileNum)))
          SC.MG = SC.MG - DoubledPenalty(FileNum).MG \ r
          SC.EG = SC.EG - DoubledPenalty(FileNum).EG \ r
        End If
        
        If Lever Then
          If Supported Then
            If RelRank = 4 Then
              SC.MG = SC.MG + 10
            ElseIf RelRank = 5 Then
              SC.MG = SC.MG + 40: SC.EG = SC.EG + 20
            ElseIf RelRank = 6 Then
              SC.MG = SC.MG + 80: SC.EG = SC.EG + 40
            End If
          Else
            If RelRank = 5 Then
              SC.MG = SC.MG + 20: SC.EG = SC.EG + 10
            ElseIf RelRank = 6 Then
              SC.MG = SC.MG + 40: SC.EG = SC.EG + 20
            End If
          End If
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
        Else
          If RankNum >= 4 Then
            If Board(Square - 19) = WPAWN Or Board(Square - 21) = WPAWN Then SC.MG = SC.MG + 25
            If Square = 84 Or Square = 85 Then
              If Board(Square + SQ_DOWN) < NO_PIECE Then SC.MG = SC.MG - 15
            End If
          End If
        End If
       
        ' bonus if safe pawn push attacks an enemy piece
        If Board(Square + SQ_DOWN) = NO_PIECE Then
          If WAttack(Board(Square + SQ_DOWN)) = 0 Or CBool(BAttack(Square + SQ_DOWN) And PAttackBit) Then ' not attacked or defend by own pawn
            SC.MG = SC.MG + 8: SC.EG = SC.EG + 8 ' Safe pawn push
            For i = 19 To 21 Step 2
              r = Board(Square - i)
              If PieceColor(r) = COL_WHITE Then
                If r = WPAWN Then
                  SC.MG = SC.MG + 15: SC.EG = SC.EG + 15 ' pawn threats enemy pawn
                Else
                  SC.MG = SC.MG + 30: SC.EG = SC.EG + 30 ' pawn threats non pawn enemy
                End If
              End If
            Next i
          End If
        End If

        ' Passed : eval later when full attack is available
        If Passed And Not Doubled Then
           PassedPawnsCnt = PassedPawnsCnt + 1: PassedPawns(PassedPawnsCnt) = Square
        ElseIf Isolated And Not Opposed Then
          MinusScore SC, IsolatedNotPassed
        End If
        
        AddScore BPawnStruct, SC
        If bEvalTrace Then WriteTrace "BPawn: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG

      Case WKING, BKING:
        ' evaluate later - King ring attack data needed ' do not remove: faster in P-Code

      Case WROOK   '--- WHITE ROOK ----

        If WPawns(FileNum) = 0 Then
          If BPawns(FileNum) = 0 Then
            SC.MG = SC.MG + 17: SC.EG = SC.EG + 8
          Else
            SC.MG = SC.MG + 8: SC.EG = SC.EG + 4
          End If
        End If


        '--- Mobility
        MobCnt = 0
        For i = 0 To 3
          Offset = QueenOffsets(i): Target = Square + Offset
          Do While Board(Target) <> FRAME
            WAttack(Target) = WAttack(Target) Xor RAttackBit
            If bDoBKSafety Then
              If MaxDistance(Target, BKingLoc) <= 2 Then If Not bKingAtt And Abs(File(Target) - FileNum) <= 1 Then AddBKingAttack PT_ROOK: bKingAtt = True
              If KingCheckB(Target) <> 0 Then
                If BAttack(Target) = 0 And MaxDistance(Target, BKingLoc) > 1 Then
                  If Board(Target) < NO_PIECE Or Board(Target) Mod 2 = 0 Then
                    Select Case Abs(KingCheckB(Target))
                    Case 1, 10: WRookChecks = WRookChecks + 1 ' Check option
                    End Select
                  End If
                End If
                If KingCheckB(Target) = -Offset And Board(Target) Mod 2 = 0 Then BPinnedCnt = BPinnedCnt + 1 ' Pinned opp piece?
              End If
            End If
                    
            Select Case Board(Target)
              Case NO_PIECE, WEP_PIECE, BEP_PIECE:
                If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: If Abs(Offset) = 10 Then SC.MG = SC.MG + 3
              Case WPAWN: SC.MG = SC.MG + 1: SC.EG = SC.EG + 2: Exit Do
              Case BPAWN: SC.MG = SC.MG + 3: SC.EG = SC.EG + 4   '--- no reattack possible
                If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_PAWN, PT_ROOK, Square, Target
                If RankNum >= 5 And Abs(Offset) = 1 Then SC.MG = SC.MG + 3: SC.EG = SC.EG + 10 ' aligned pawns
                Exit Do
              Case BBISHOP: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                AddThreat COL_BLACK, PT_BISHOP, PT_ROOK, Square, Target: Exit Do    '--- no reattack possible
              Case BKNIGHT: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                AddThreat COL_BLACK, PT_KNIGHT, PT_ROOK, Square, Target: Exit Do  '--- no reattack possible
              Case BROOK: AddThreat COL_BLACK, PT_ROOK, PT_ROOK, Square, Target: MobCnt = MobCnt + 1: Exit Do ' equal exchange, ok for mobility
              Case BKING: MobCnt = MobCnt + 1: Exit Do
              Case BQUEEN: MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_QUEEN, PT_ROOK, Square, Target: Exit Do
              Case WROOK, WQUEEN: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                If Offset = 10 Then
                  If WPawns(FileNum) = 0 Then SC.MG = SC.MG + 5: If BPawns(FileNum) = 0 Then SC.MG = SC.MG + 6
                End If
                SC.MG = SC.MG + 4: SC.EG = SC.EG + 2    '--- double lines , continue xray
              Case Else: If Not CBool(BAttack(Target) And PAttackBit) And Target <> WKingLoc Then MobCnt = MobCnt + 1  ' own bishop or knight
                Exit Do
            End Select
            Target = Target + Offset
          Loop
        Next
            
        AddScore WMobility, MobilityR(MobCnt)

        ' Trapped rook by king : worse when cannot castle
        If Not bEndgame Then
          If MobCnt <= 3 And WPawns(FileNum) > 0 Then
            If RankNum = Rank(WKingLoc) Or Rank(WKingLoc) = 1 Then
              r = 0
              If File(WKingLoc) < 5 Then
                If FileNum < File(WKingLoc) Then r = -1
              Else
                If FileNum > File(WKingLoc) Then r = 1
              End If
              If r <> 0 Then
                For k = File(WKingLoc) + r To FileNum - r Step r ' own blocking pawns on files between king an rook
                 If WPawns(k) = 0 Then r = 0: Exit For
                Next
                If r <> 0 Then SC.MG = SC.MG - EvalSFTo100(92 - MobCnt * 22) * (1 + Abs(Moved(WKING_START) > 0 Or Moved(Square) > 0))
              End If
            End If
          End If
        Else
          If WPawns(FileNum) > 0 And BPawns(FileNum) = 0 And PawnsWMin(FileNum) >= 5 Then
            SC.MG = SC.MG + (PawnsWMin(FileNum)): SC.EG = SC.EG + 2 * PawnsWMin(FileNum)
          End If
        End If
        AddScore100 WPos, SC
        If bEvalTrace Then WriteTrace "WRook: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & WPos.MG & ", " & WPos.EG

      Case BROOK   '--- BLACK ROOK ----
         
         
        If BPawns(FileNum) = 0 Then
          If WPawns(FileNum) = 0 Then
            SC.MG = SC.MG + 17: SC.EG = SC.EG + 8
          Else
            SC.MG = SC.MG + 8: SC.EG = SC.EG + 4
          End If
        End If
        
        '--- Mobility
        MobCnt = 0
        For i = 0 To 3
          Offset = QueenOffsets(i): Target = Square + Offset
          Do While Board(Target) <> FRAME
            BAttack(Target) = BAttack(Target) Xor RAttackBit
            If bDoWKSafety Then
              If MaxDistance(Target, WKingLoc) <= 2 Then If Not bKingAtt And Abs(File(Target) - FileNum) <= 1 Then AddWKingAttack PT_ROOK: bKingAtt = True
              If KingCheckW(Target) <> 0 Then
                If WAttack(Target) = 0 And MaxDistance(Target, WKingLoc) > 1 Then
                  If Board(Target) < NO_PIECE Or Board(Target) Mod 2 = 1 Then
                    Select Case Abs(KingCheckW(Target))
                    Case 1, 10: BRookChecks = BRookChecks + 1
                    End Select
                  End If
                End If
                If KingCheckW(Target) = -Offset And Board(Target) Mod 2 = 1 Then WPinnedCnt = WPinnedCnt + 1 ' Pinned opp piece?
              End If
            End If
            Select Case Board(Target)
              Case NO_PIECE, WEP_PIECE, BEP_PIECE:
                If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: If Abs(Offset) = 10 Then SC.MG = SC.MG + 3
              Case BPAWN: SC.MG = SC.MG + 1: SC.EG = SC.EG + 2: Exit Do
              Case WPAWN: SC.MG = SC.MG + 3: SC.EG = SC.EG + 4  '--- no reattack possible
                If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_PAWN, PT_ROOK, Square, Target
                If RankNum <= 4 And Abs(Offset) = 1 Then SC.MG = SC.MG + 3: SC.EG = SC.EG + 10 ' aligned pawns
                Exit Do
              Case WBISHOP: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                AddThreat COL_WHITE, PT_BISHOP, PT_ROOK, Square, Target: Exit Do   '--- no reattack possible
              Case WKNIGHT: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                AddThreat COL_WHITE, PT_KNIGHT, PT_ROOK, Square, Target: Exit Do    '--- no reattack possible
              Case WROOK: AddThreat COL_WHITE, PT_ROOK, PT_ROOK, Square, Target: MobCnt = MobCnt + 1: Exit Do  ' equal exchange ok for mobility
              Case WKING: MobCnt = MobCnt + 1: Exit Do
              Case WQUEEN: MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_QUEEN, PT_ROOK, Square, Target: Exit Do
              Case BROOK, BQUEEN: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                If Offset = -10 Then
                  If BPawns(FileNum) = 0 Then SC.MG = SC.MG + 5: If WPawns(FileNum) = 0 Then SC.MG = SC.MG + 6
                End If
                SC.MG = SC.MG + 4: SC.EG = SC.EG + 2   '--- double lines , continue xray
              Case Else: If Not CBool(WAttack(Target) And PAttackBit) And Target <> BKingLoc Then MobCnt = MobCnt + 1  ' own bishop or knight
                Exit Do
            End Select
            Target = Target + Offset
          Loop
        Next
        AddScore BMobility, MobilityR(MobCnt)

        ' Trapped rook by king : worse when cannot castle
        If Not bEndgame Then
          If MobCnt <= 3 And BPawns(FileNum) > 0 Then
            If RankNum = Rank(BKingLoc) Or Rank(BKingLoc) = 1 Then
              r = 0
              If File(BKingLoc) < 5 Then
                If FileNum < File(BKingLoc) Then r = -1
              Else
                If FileNum > File(BKingLoc) Then r = 1
              End If
              If r <> 0 Then
                For k = File(BKingLoc) + r To FileNum - r Step r ' own blocking pawns on files between king an rook
                 If BPawns(k) = 0 Then r = 0: Exit For
                Next
                If r <> 0 Then SC.MG = SC.MG - EvalSFTo100(92 - MobCnt * 22) * (1 + Abs(Moved(BKING_START) > 0 Or Moved(Square) > 0))
              End If
            End If
          End If
        Else
          If BPawns(FileNum) > 0 And WPawns(FileNum) = 0 And PawnsBMax(FileNum) <= 4 Then
            SC.MG = SC.MG + (9 - PawnsBMin(FileNum)): SC.EG = SC.EG + 2 * (9 - PawnsBMin(FileNum))
          End If
        End If
         
        AddScore100 BPos, SC
        If bEvalTrace Then WriteTrace "BRook: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & BPos.MG & ", " & BPos.EG

      Case WBISHOP    '--- WHITE BISHOP ----
        
        '--- Mobility
        MobCnt = 0
        For i = 4 To 7
          Offset = QueenOffsets(i): Target = Square + Offset
          Do While Board(Target) <> FRAME
            WAttack(Target) = WAttack(Target) Xor BAttackBit
            If bDoBKSafety Then
              If MaxDistance(Target, BKingLoc) <= 2 Then If Not bKingAtt And Abs(File(Target) - FileNum) <= 1 Then AddBKingAttack PT_BISHOP: bKingAtt = True
              If KingCheckB(Target) <> 0 Then
                If BAttack(Target) = 0 And MaxDistance(Target, BKingLoc) > 1 Then
                  If Board(Target) < NO_PIECE Or Board(Target) Mod 2 = 0 Then
                    Select Case Abs(KingCheckB(Target))
                    Case 9, 11: WBishopChecks = WBishopChecks + 1
                    End Select
                  End If
                End If
                If KingCheckB(Target) = -Offset And Board(Target) Mod 2 = 0 Then BPinnedCnt = BPinnedCnt + 1 ' Pinned opp piece?
              End If
            End If
            Select Case Board(Target)
              Case NO_PIECE, WEP_PIECE, BEP_PIECE:
                If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: If Offset > 0 Then SC.MG = SC.MG + 2
              Case WPAWN: SC.MG = SC.MG + 1: SC.EG = SC.EG + 1: Exit Do
              Case BPAWN: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_PAWN, PT_BISHOP, Square, Target
                SC.MG = SC.MG + 3: SC.EG = SC.EG + 3: Exit Do
              Case BBISHOP:  MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_BISHOP, PT_BISHOP, Square, Target: Exit Do ' Reattack: no SC because x-x=0
              Case BKNIGHT: MobCnt = MobCnt + 1:  AddThreat COL_BLACK, PT_KNIGHT, PT_BISHOP, Square, Target: Exit Do
              Case BROOK: MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_ROOK, PT_BISHOP, Square, Target: Exit Do
              Case BKING: MobCnt = MobCnt + 1:  Exit Do
              Case BQUEEN: MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_QUEEN, PT_BISHOP, Square, Target:  Exit Do
              Case WQUEEN: MobCnt = MobCnt + 1: SC.MG = SC.MG + 3: SC.EG = SC.EG + 3 '--- Continue xray
              Case Else: If Not CBool(BAttack(Target) And PAttackBit) And Target <> WKingLoc Then MobCnt = MobCnt + 1  ' own bishop or knight
                Exit Do
            End Select
            Target = Target + Offset
          Loop
        Next
        AddScore WMobility, MobilityB(MobCnt)

        AddScore100 WPos, SC
          
        ' Minor behind pawn bonus
        If RelRank < 5 Then
          If PieceType(Board(Square + SQ_UP)) = PT_PAWN Then AddScoreVal WPos, 16, 0: If bEvalTrace Then WriteTrace "WBishop: " & LocCoord(Square) & "> Behind pawn 16"
        End If
  
        ' Outpost bonus
        If RelRank >= 4 And RelRank <= 6 Then
          ' Defended by pawn?
          AddScore WPos, OutpostBonusBishop(Abs(CBool(WAttack(Square) And PAttackBit)))
          If bEvalTrace Then WriteTrace "WBishop: " & LocCoord(Square) & "> Outpost:" & OutpostBonusBishop(Abs(CBool(WAttack(Square) And PAttackBit))).MG
        End If
        
        If bEvalTrace Then WriteTrace "WBishop: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & WPos.MG & ", " & WPos.EG

      Case BBISHOP   '--- BLACK BISHOP ----
    
        '--- Mobility
        MobCnt = 0
        For i = 4 To 7
          Offset = QueenOffsets(i): Target = Square + Offset
          Do While Board(Target) <> FRAME
            BAttack(Target) = BAttack(Target) Xor BAttackBit
            If bDoWKSafety Then
              If MaxDistance(Target, WKingLoc) <= 2 Then If Not bKingAtt And Abs(File(Target) - FileNum) <= 1 Then AddWKingAttack PT_BISHOP: bKingAtt = True
              If KingCheckW(Target) <> 0 Then
                If WAttack(Target) = 0 And MaxDistance(Target, WKingLoc) > 1 Then
                  If Board(Target) < NO_PIECE Or Board(Target) Mod 2 = 1 Then
                    Select Case Abs(KingCheckW(Target))
                    Case 9, 11: BBishopChecks = BBishopChecks + 1
                    End Select
                  End If
                End If
                If KingCheckW(Target) = -Offset And Board(Target) Mod 2 = 1 Then WPinnedCnt = WPinnedCnt + 1 ' Pinned opp piece?
              End If
            End If
                   
            Select Case Board(Target)
              Case NO_PIECE, WEP_PIECE, BEP_PIECE:
                If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1:  If Offset < 0 Then SC.MG = SC.MG + 2
              Case BPAWN: SC.MG = SC.MG + 1: SC.EG = SC.EG + 1: Exit Do
              Case WPAWN:  If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_PAWN, PT_BISHOP, Square, Target
                SC.MG = SC.MG + 3: SC.EG = SC.EG + 3: Exit Do
              Case WBISHOP: MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_BISHOP, PT_BISHOP, Square, Target: Exit Do ' Reattack: no SC because x-x=0
              Case WKNIGHT: MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_KNIGHT, PT_BISHOP, Square, Target: Exit Do
              Case WROOK: MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_ROOK, PT_BISHOP, Square, Target:  Exit Do
              Case WKING: MobCnt = MobCnt + 1:  Exit Do
              Case WQUEEN: MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_QUEEN, PT_BISHOP, Square, Target:  Exit Do
              Case BQUEEN: MobCnt = MobCnt + 1: SC.MG = SC.MG + 3: SC.EG = SC.EG + 3 '--- Continue xray
              Case Else: If Not CBool(WAttack(Target) And PAttackBit) And Target <> BKingLoc Then MobCnt = MobCnt + 1  ' own bishop or knight
                Exit Do
            End Select
            Target = Target + Offset
          Loop
        Next
        AddScore BMobility, MobilityB(MobCnt)

        AddScore100 BPos, SC
    
        ' Minor behind pawn bonus
        If RelRank < 5 Then
          If PieceType(Board(Square + SQ_DOWN)) = PT_PAWN Then AddScoreVal BPos, 16, 0: If bEvalTrace Then WriteTrace "BBishop: " & LocCoord(Square) & "> Behind pawn 16"
        End If
            
        ' Outpost bonus
        If RelRank >= 4 And RelRank <= 6 Then
          ' Defended by pawn?
          AddScore BPos, OutpostBonusBishop(Abs(CBool(BAttack(Square) And PAttackBit)))
          If bEvalTrace Then WriteTrace "BBishop: " & LocCoord(Square) & "> Outpost:" & OutpostBonusBishop(Abs(CBool(BAttack(Square) And PAttackBit))).MG
        End If
        
        If bEvalTrace Then WriteTrace "BBishop: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & BPos.MG & ", " & BPos.EG

      Case WKNIGHT   '--- WHITE KNIGHT ----
        '--- Mobility
        If Moved(Square) = 0 Then AddScoreVal SC, -18, 0
        ForkCnt = 0: MobCnt = 0
        For i = 0 To 7
          Offset = KnightOffsets(i): Target = Square + Offset
          If Board(Target) <> FRAME Then
            WAttack(Target) = WAttack(Target) Xor NAttackBit
            If bDoBKSafety Then
              r = MaxDistance(Target, BKingLoc)
              If Not bKingAtt And r <= 2 Then If Abs(File(Target) - FileNum) <= 1 Then AddBKingAttack PT_KNIGHT: bKingAtt = True
  
              If KingCheckB(Target) <> 0 Then
                If r <= 4 And BAttack(Target) = 0 Then
                  If Board(Target) < NO_PIECE Or Board(Target) Mod 2 = 0 Then
                    Select Case Abs(KingCheckB(Target)):
                      Case 8, 12, 19, 21:
                        WKnightChecks = WKnightChecks + 1
                    End Select
                  End If
                End If
              End If
            End If
            Select Case Board(Target)
              Case NO_PIECE, WEP_PIECE, BEP_PIECE:
                If (Not CBool(BAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1: SC.MG = SC.MG + 2
              Case WPAWN: SC.MG = SC.MG + 1: SC.EG = SC.EG + 1
              Case BPAWN: SC.MG = SC.MG + 3: SC.EG = SC.EG + 3: If Rank(Target) >= 6 Then SC.MG = SC.MG + 4
                If (Not CBool(BAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_PAWN, PT_KNIGHT, Square, Target
              Case BKNIGHT: MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_KNIGHT, PT_KNIGHT, Square, Target   '-- no Score for WKnight : total is zero
              Case BBISHOP: MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_BISHOP, PT_KNIGHT, Square, Target
              Case BROOK: MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_ROOK, PT_KNIGHT, Square, Target: ForkCnt = ForkCnt + 1:
              Case BKING: MobCnt = MobCnt + 1: ForkCnt = ForkCnt + 1:
              Case BQUEEN: MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_QUEEN, PT_KNIGHT, Square, Target: ForkCnt = ForkCnt + 1:
              Case Else: If Not CBool(BAttack(Target) And PAttackBit) And Target <> WKingLoc Then MobCnt = MobCnt + 1 ' other own piece
            End Select
          End If
        Next
        If ForkCnt > 1 Then AddScoreVal SC, 3 * ForkCnt * ForkCnt, 2 * ForkCnt * ForkCnt: If bWhiteToMove Then AddScoreVal SC, 15, 15
        AddScore WMobility, MobilityN(MobCnt)

        AddScore100 WPos, SC
  
        ' Minor behind pawn bonus
        If RelRank < 5 Then
          If PieceType(Board(Square + SQ_UP)) = PT_PAWN Then AddScoreVal WPos, 16, 0: If bEvalTrace Then WriteTrace "WKnight: " & LocCoord(Square) & "> Behind pawn 16"
        End If
          
        ' Outpost bonus
        If RelRank >= 4 And RelRank <= 6 Then
          ' Defended by pawn?
          AddScore WPos, OutpostBonusKnight(Abs(CBool(WAttack(Square) And PAttackBit)))
          If bEvalTrace Then WriteTrace "WKight: " & LocCoord(Square) & "> Outpost:" & OutpostBonusKnight(Abs(CBool(WAttack(Square) And PAttackBit))).MG
        End If
        If bEvalTrace Then WriteTrace "WKnight: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & WPos.MG & ", " & WPos.EG

      Case BKNIGHT   '--- BLACK KNIGHT ----

        If Moved(Square) = 0 Then AddScoreVal SC, -18, 0
        '--- Mobility
        ForkCnt = 0: MobCnt = 0
        For i = 0 To 7
          Offset = KnightOffsets(i)
          Target = Square + Offset
          If Board(Target) <> FRAME Then
            BAttack(Target) = BAttack(Target) Xor NAttackBit
            If bDoWKSafety Then
              r = MaxDistance(Target, WKingLoc)
              If Not bKingAtt And r <= 2 Then If Abs(File(Target) - FileNum) <= 1 Then AddWKingAttack PT_KNIGHT: bKingAtt = True
  
              If KingCheckW(Target) <> 0 Then
                If r <= 4 And WAttack(Target) = 0 Then
                  If Board(Target) < NO_PIECE Or Board(Target) Mod 2 = 1 Then
                    Select Case Abs(KingCheckW(Target)):
                      Case 8, 12, 19, 21:
                        BKnightChecks = BKnightChecks + 1
                    End Select
                  End If
                End If
              End If
            End If
            Select Case Board(Target)
              Case NO_PIECE, WEP_PIECE, BEP_PIECE:
                If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1: SC.MG = SC.MG + 3
              Case BPAWN: SC.MG = SC.MG + 1: SC.EG = SC.EG + 1
              Case WPAWN: SC.MG = SC.MG + 3: SC.EG = SC.EG + 3: If Rank(Target) <= 3 Then SC.MG = SC.MG + 4
                If (Not CBool(WAttack(Target) And PAttackBit)) Then MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_PAWN, PT_KNIGHT, Square, Target
              Case WKNIGHT: MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_KNIGHT, PT_KNIGHT, Square, Target   '-- no Score for WKnight : total is zero
              Case WBISHOP: MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_BISHOP, PT_KNIGHT, Square, Target   '-- no Score for WKnight : total is zero
              Case WROOK: MobCnt = MobCnt + 1:  AddThreat COL_WHITE, PT_ROOK, PT_KNIGHT, Square, Target: ForkCnt = ForkCnt + 1
              Case WKING: MobCnt = MobCnt + 1: ForkCnt = ForkCnt + 1::
              Case WQUEEN: MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_QUEEN, PT_KNIGHT, Square, Target: ForkCnt = ForkCnt + 1:
              Case Else: If Not CBool(WAttack(Target) And PAttackBit) And Target <> BKingLoc Then MobCnt = MobCnt + 1
            End Select
            Target = Target + Offset
          End If
        Next
        If ForkCnt > 1 Then AddScoreVal SC, 3 * ForkCnt * ForkCnt, 3 * ForkCnt * ForkCnt: If Not bWhiteToMove Then AddScoreVal SC, 15, 15
        AddScore BMobility, MobilityN(MobCnt)

        AddScore100 BPos, SC
  
        ' Minor behind pawn bonus
        If RelRank < 5 Then
          If PieceType(Board(Square + SQ_DOWN)) = PT_PAWN Then AddScoreVal BPos, 16, 0: If bEvalTrace Then WriteTrace "BKnight: " & LocCoord(Square) & "> Behind pawn 16"
        End If
          
        ' Outpost bonus
        If RelRank >= 4 And RelRank <= 6 Then
          ' Defended by pawn?
          AddScore BPos, OutpostBonusKnight(Abs(CBool(BAttack(Square) And PAttackBit)))
          If bEvalTrace Then WriteTrace "BKight: " & LocCoord(Square) & "> Outpost:" & OutpostBonusKnight(Abs(CBool(BAttack(Square) And PAttackBit))).MG
        End If
        If bEvalTrace Then WriteTrace "BKnight: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & BPos.MG & ", " & BPos.EG

      Case WQUEEN   '--- WHITE QUEEN ----

        '--- Mobility
        MobCnt = 0
        For i = 0 To 7
          Offset = QueenOffsets(i): Target = Square + Offset
          Do While Board(Target) <> FRAME
            WAttack(Target) = WAttack(Target) Xor QAttackBit
            If bDoBKSafety Then
              If MaxDistance(Target, BKingLoc) <= 2 Then If Not bKingAtt And Abs(File(Target) - FileNum) <= 1 Then AddBKingAttack PT_QUEEN: bKingAtt = True
              If KingCheckB(Target) <> 0 Then
                If BAttack(Target) = 0 And MaxDistance(Target, BKingLoc) > 1 Then
                  If Board(Target) < NO_PIECE Or Board(Target) Mod 2 = 0 Then
                    Select Case Abs(KingCheckB(Target))
                    Case 1, 10, 9, 11: WQueenChecks = WQueenChecks + 1
                    End Select
                  End If
                End If
                If KingCheckB(Target) = -Offset And Board(Target) Mod 2 = 0 Then BPinnedCnt = BPinnedCnt + 1 ' Pinned opp piece?
              End If
            End If
            If bDoWKSafety Then
              If MaxDistance(Square, WKingLoc) > 2 And MaxDistance(Target, WKingLoc) = 1 Then AddWKingDefend PT_QUEEN
            End If
            Select Case Board(Target)
              Case NO_PIECE, WEP_PIECE, BEP_PIECE: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
              Case WPAWN: SC.MG = SC.MG + 1: SC.EG = SC.EG + 1: Exit Do   'Defends pawn
              Case BPAWN: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: AddThreat COL_BLACK, PT_PAWN, PT_QUEEN, Square, Target
                SC.MG = SC.MG + 3: SC.EG = SC.EG + 3: Exit Do   'Attack pawn
              Case BBISHOP: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                AddThreat COL_BLACK, PT_BISHOP, PT_QUEEN, Square, Target: Exit Do
              Case BKNIGHT: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                AddThreat COL_BLACK, PT_KNIGHT, PT_QUEEN, Square, Target: Exit Do
              Case BROOK: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                AddThreat COL_BLACK, PT_ROOK, PT_QUEEN, Square, Target: Exit Do
              Case BKING: MobCnt = MobCnt + 1: Exit Do
              Case BQUEEN: AddThreat COL_BLACK, PT_QUEEN, PT_QUEEN, Square, Target: MobCnt = MobCnt + 1: Exit Do
              Case WBISHOP: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                SC.MG = SC.MG + 4: SC.EG = SC.EG + 4: Exit Do '--- double lines
              Case WKNIGHT: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                SC.MG = SC.MG + 2: SC.EG = SC.EG + 2: Exit Do   'Defends own piece
              Case WROOK: If Not CBool(BAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                If Offset = 10 Then
                  If WPawns(FileNum) = 0 Then
                    SC.MG = SC.MG + 4: SC.EG = SC.EG + 2
                  ElseIf BPawns(FileNum) = 0 Then
                    SC.MG = SC.MG + 6: SC.EG = SC.EG + 2
                  End If
                End If
                SC.MG = SC.MG + 4: SC.EG = SC.EG + 1 '--- double lines , continue xray
              Case Else: Exit Do
            End Select
            Target = Target + Offset
          Loop
        Next
        AddScore WMobility, MobilityQ(MobCnt)
        AddScore100 WPos, SC
        If bEvalTrace Then WriteTrace "WQueen: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & WPos.MG & ", " & WPos.EG

      Case BQUEEN  '--- BLACK QUEEN ----
        
        '--- Mobility
        MobCnt = 0
        For i = 0 To 7
          Offset = QueenOffsets(i): Target = Square + Offset
          Do While Board(Target) <> FRAME
            BAttack(Target) = BAttack(Target) Xor QAttackBit
            If bDoBKSafety Then
              If MaxDistance(Target, WKingLoc) <= 2 Then If Not bKingAtt And Abs(File(Target) - FileNum) <= 1 Then AddWKingAttack PT_QUEEN: bKingAtt = True
              If KingCheckW(Target) <> 0 Then
                If WAttack(Target) = 0 And MaxDistance(Target, WKingLoc) > 1 Then
                  If Board(Target) < NO_PIECE Or Board(Target) Mod 2 = 1 Then
                    Select Case Abs(KingCheckW(Target)):
                    Case 1, 10, 9, 11: BQueenChecks = BQueenChecks + 1
                    End Select
                  End If
                End If
                If KingCheckW(Target) = -Offset And Board(Target) Mod 2 = 1 Then WPinnedCnt = WPinnedCnt + 1 ' Pinned opp piece?
              End If
            End If
            If bDoBKSafety Then
              If MaxDistance(Square, BKingLoc) > 2 And MaxDistance(Target, BKingLoc) = 1 Then AddBKingDefend PT_QUEEN
            End If
            Select Case Board(Target)
              Case NO_PIECE, WEP_PIECE, BEP_PIECE: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
              Case BPAWN: SC.MG = SC.MG + 1: SC.EG = SC.EG + 1: Exit Do   'Defends pawn
              Case WPAWN:
                If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1: AddThreat COL_WHITE, PT_PAWN, PT_QUEEN, Square, Target
                SC.MG = SC.MG + 3: SC.EG = SC.EG + 3: Exit Do   'Attack pawn
              Case WBISHOP: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                AddThreat COL_WHITE, PT_BISHOP, PT_QUEEN, Square, Target: Exit Do
              Case WKNIGHT: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                AddThreat COL_WHITE, PT_KNIGHT, PT_QUEEN, Square, Target: Exit Do
              Case WROOK: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                AddThreat COL_WHITE, PT_ROOK, PT_QUEEN, Square, Target: Exit Do '- Reattacked, already score by opp piece
              Case WKING: MobCnt = MobCnt + 1: Exit Do
              Case WQUEEN:  AddThreat COL_WHITE, PT_QUEEN, PT_QUEEN, Square, Target: MobCnt = MobCnt + 1: Exit Do
              Case BBISHOP:  If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                SC.MG = SC.MG + 4: SC.EG = SC.EG + 4: Exit Do '--- double lines
              Case BKNIGHT: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                SC.MG = SC.MG + 2: SC.EG = SC.EG + 2: Exit Do   'Defends own piece
              Case BROOK: If Not CBool(WAttack(Target) And PAttackBit) Then MobCnt = MobCnt + 1
                If Offset = -10 Then
                  If BPawns(FileNum) = 0 Then
                    SC.MG = SC.MG + 4: SC.EG = SC.EG + 2
                  ElseIf WPawns(FileNum) = 0 Then
                    SC.MG = SC.MG + 6: SC.EG = SC.EG + 2
                  End If
                End If
                SC.MG = SC.MG + 4: SC.EG = SC.EG + 1   '--- double lines, continue xray
              Case Else: Exit Do
            End Select
            Target = Target + Offset
          Loop
        Next
        AddScore BMobility, MobilityQ(MobCnt)

        AddScore100 BPos, SC
        If bEvalTrace Then WriteTrace "WQueen: " & LocCoord(Square) & ">" & SC.MG & ", " & SC.EG & " / " & BPos.MG & ", " & BPos.EG
    
    End Select '<<< End of pieces select

lblNextPiece:
  Next '<<< Next piece

  '
  '--- Global eval scores -------------------------------------------
  '

  If bEndgame Then
    ScoreToZero WKSafety: ScoreToZero BKSafety
  Else
  
    Dim RankUs As Integer, RankThem As Integer, RelFile As Integer, Bonus As Integer, QDefend As Integer, ShelterKingLoc As Integer

    '----------------------------------------------
    '--- White King Safety Eval -------------------
    '----------------------------------------------
    RankNum = Rank(WKingLoc): FileNum = File(WKingLoc): Bonus = 0
    If (BQueenCnt * 2 + BRookCnt) > 1 Then
      AttackUnits = 0: Safety = 258 ' MaxSafetyBonus
      If WPawnCnt = 0 Then MinWKingPawnDistance = 0 Else MinWKingPawnDistance = MinWKingPawnDistance - 1
      
      If RankNum > 4 Then
        WKSafety.EG = WKSafety.EG - 16 * MinWKingPawnDistance
      Else
        '--- Pawn shelter
        ShelterKingLoc = WKRelLoc: k = File(ShelterKingLoc)
        If WhiteCastled = NO_CASTLE And WKingLoc = SQ_E1 Then
          If WPawns(7) > 0 And PawnsWMin(7) < 4 And WCanCastleOO() Then
            ShelterKingLoc = SQ_G1
          ElseIf WPawns(3) > 0 And PawnsWMin(3) < 4 And WCanCastleOOO() Then
            ShelterKingLoc = SQ_C1
          End If
        End If
        
        For i = -1 To 1
          k = File(ShelterKingLoc + i)
          If k >= 1 And k <= 8 Then
            ' Pawn shelter/storm
            RankUs = 1
            If WPawns(k) > 0 Then If PawnsWMin(k) > RankNum Then RankUs = PawnsWMin(k)
            RankThem = 8
            If BPawns(k) > 0 Then If PawnsBMax(k) > RankNum Then RankThem = PawnsBMax(k)
            If i = 0 And RankThem = RankNum + 1 Then
              r = 1 ' BlockedByKing
            ElseIf RankUs = 1 Then
              r = 2 ' NoFriendlyPawn
            ElseIf RankThem = RankUs + 1 Then
              r = 3 ' BlockedByPawn
            Else
              r = 4 ' Unblocked
            End If
            RelFile = GetMin(k, 9 - k)
            Safety = Safety - ShelterWeakness(RelFile, RankUs)
            Safety = Safety - StormDanger(r, RelFile, RankThem)
          End If
        Next
        Bonus = Safety
        AddScoreVal WKSafety, Bonus, -16 * MinWKingPawnDistance
      End If
      
      If bDoWKSafety Then
      
        '--- Check threats at king ring
        Undefended = 0: KingAdjacentZoneAttacksCount = 0
        For i = 0 To 7
          Target = WKingLoc + QueenOffsets(i)
          If Board(Target) <> FRAME Then
            r = BAttack(Target)
            If WAttack(Target) = 0 Then
              Undefended = Undefended + 1
              If r > 0 Then
                If CBool(r And QAttackBit) Then
                  AttackUnits = AttackUnits + 50
                ElseIf CBool(r And RAttackBit) Then
                  AttackUnits = AttackUnits + 30
                ElseIf CBool(r And BAttackBit) Then
                  AttackUnits = AttackUnits + 25
                ElseIf CBool(r And NAttackBit) Then
                  AttackUnits = AttackUnits + 40 ' high Attack weight !?
                Else ' Pawn, King
                  AttackUnits = AttackUnits + 12
                End If
              Else
                AttackUnits = AttackUnits + 5
              End If
            End If
            If r > 0 Then
              KingAdjacentZoneAttacksCount = KingAdjacentZoneAttacksCount + 1
            ElseIf bWhiteToMove And WAttack(Target) >= QAttackBit And Board(Target) = NO_PIECE And MaxDistance(WQueenLoc, WKingLoc) > 1 Then
              ' Queen can help
              QDefend = QDefend + 1
              If QDefend = 1 Then
                AttackUnits = AttackUnits - 50 - Abs(BQueenCnt <> 0) * 50 ' opp queen exchange option
              Else
                AttackUnits = AttackUnits - 15
              End If
            End If
            ' Safe contact checks
            If (r >= RAttackBit) And (WAttack(Target) = 0) Then   ' Undefend and not pawn, not an attacker piece there
              If Board(Target) = NO_PIECE Or (Board(Target) Mod 2 = 1) Then
                ' Supported check by another enemy: r <> QAttackBit
                If CBool(r And QAttackBit) Then If r <> QAttackBit Then AttackUnits = AttackUnits + QueenContactCheck: MinusScore WKSafety, Checked: GoTo lblNextWCheck
                '(i=0-3: orthogonal offset, 4-7:diagonal)
                If i < 4 And CBool(r And RAttackBit) Then If r <> RAttackBit Then AttackUnits = AttackUnits + RookContactCheck: MinusScore WKSafety, Checked: GoTo lblNextWCheck
                If i >= 4 And CBool(r And BAttackBit) Then If r <> BAttackBit Then AttackUnits = AttackUnits + BishopCheck: MinusScore WKSafety, Checked
              End If
            End If
          End If
lblNextWCheck:
        Next
        
        ' distance checks
        AttackUnits = AttackUnits + BKnightChecks * KnightCheck + BBishopChecks * BishopCheck + BRookChecks * RookCheck + BQueenChecks * QueenCheck
        
        ' total attackunits
        AttackUnits = AttackUnits + GetMin(74, WKingAttackersCount * WKingAttackersWeight) + 8 * KingAdjacentZoneAttacksCount + 11 * Abs(WPinnedCnt > 0) - Bonus \ 8 - Abs(BQueenCnt = 0) * 100
        
        ' Penalty for king on open or semi-open file
        If NonPawnMaterial > 9000 And WPawns(FileNum) = 0 And WKingLoc <> WKING_START Then
          If BPawns(FileNum) = 0 Then AttackUnits = AttackUnits + 18 Else AttackUnits = AttackUnits + 9
        End If
              
        AttackUnits = AttackUnits + BPassedPawnAttack * 8 ' passed pawn attacking king?
        WKSafety.MG = WKSafety.MG - KingDanger(GetMax(GetMin(AttackUnits, 399), 0)).MG
        
      End If
    End If

 

    '----------------------------------------------
    '--- Black King Safety Eval -------------------
    '----------------------------------------------

    RankNum = Rank(BKingLoc): RelRank = (9 - RankNum): FileNum = File(BKingLoc): Bonus = 0
    If (WQueenCnt * 2 + WRookCnt) > 1 Then
      AttackUnits = 0: Safety = 258 ' MaxSafetyBonus
      If BPawnCnt = 0 Then MinBKingPawnDistance = 0 Else MinBKingPawnDistance = MinBKingPawnDistance - 1
            
      If RelRank > 4 Then
        BKSafety.EG = BKSafety.EG - 16 * MinBKingPawnDistance
      Else
        '--- Pawn shelter
        ShelterKingLoc = BKRelLoc: k = File(ShelterKingLoc)
        If BlackCastled = NO_CASTLE And BKingLoc = SQ_E8 Then
          If BPawns(7) > 0 And PawnsBMax(7) > 5 And BCanCastleOO() Then
            ShelterKingLoc = SQ_G8
          ElseIf BPawns(3) > 0 And PawnsBMax(3) > 5 And BCanCastleOOO() Then
            ShelterKingLoc = SQ_C8
          End If
        End If
        For i = -1 To 1
          k = File(ShelterKingLoc + i)
          If k >= 1 And k <= 8 Then
            ' Pawn shelter/storm
            RankUs = 1
            If BPawns(k) > 0 Then If PawnsBMax(k) < RankNum Then RankUs = (9 - PawnsBMax(k))
            RankThem = 8
            If WPawns(k) > 0 Then If PawnsWMin(k) < RankNum Then RankThem = (9 - PawnsWMin(k))
            If i = 0 And RankThem = RelRank + 1 Then
              r = 1 ' BlockedByKing
            ElseIf RankUs = 1 Then
              r = 2 ' NoFriendlyPawn
            ElseIf RankThem = RankUs + 1 Then
              r = 3 ' BlockedByPawn
            Else
              r = 4 ' Unblocked
            End If
            RelFile = GetMin(k, 9 - k)
            Safety = Safety - ShelterWeakness(RelFile, RankUs)
            Safety = Safety - StormDanger(r, RelFile, RankThem)
          End If
        Next
        Bonus = Safety
        AddScoreVal BKSafety, Bonus, -16 * MinBKingPawnDistance
      End If
        
      If bDoBKSafety Then
        '--- Check threats at king ring
        Undefended = 0: KingAdjacentZoneAttacksCount = 0: QDefend = 0
        For i = 0 To 7
          Target = BKingLoc + QueenOffsets(i)
          If Board(Target) <> FRAME Then
            r = WAttack(Target)
            If BAttack(Target) = 0 Then
              Undefended = Undefended + 1
              If r > 0 Then
                If CBool(r And QAttackBit) Then
                  AttackUnits = AttackUnits + 50
                ElseIf CBool(r And RAttackBit) Then
                  AttackUnits = AttackUnits + 30
                ElseIf CBool(r And BAttackBit) Then
                  AttackUnits = AttackUnits + 25
                ElseIf CBool(r And NAttackBit) Then
                  AttackUnits = AttackUnits + 40 ' high Attack weight !?
                Else ' Pawn, King
                  AttackUnits = AttackUnits + 12
                End If
              Else
                AttackUnits = AttackUnits + 5
              End If
            End If
            If r > 0 Then
              KingAdjacentZoneAttacksCount = KingAdjacentZoneAttacksCount + 1
            ElseIf Not bWhiteToMove And BAttack(Target) >= QAttackBit And Board(Target) = NO_PIECE And MaxDistance(BQueenLoc, BKingLoc) > 1 Then
              ' Queen can help
              QDefend = QDefend + 1
              If QDefend = 1 Then
                AttackUnits = AttackUnits - 50 - Abs(WQueenCnt <> 0) * 50
              Else
                AttackUnits = AttackUnits - 15
              End If
            End If
          
            ' Safe contact checks
            If (r >= RAttackBit) And (BAttack(Target) = 0) Then   ' Undefend and not pawn, not an attacker piece there
              If Board(Target) = NO_PIECE Or (Board(Target) Mod 2 = 0) Then
                ' Supported check by another enemy: r <> QAttackBit
                If CBool(r And QAttackBit) Then If r <> QAttackBit Then AttackUnits = AttackUnits + QueenContactCheck: MinusScore BKSafety, Checked: GoTo lblNextBCheck
                If i < 4 And CBool(r And RAttackBit) Then If r <> RAttackBit Then AttackUnits = AttackUnits + RookContactCheck: MinusScore BKSafety, Checked: GoTo lblNextBCheck
                If i >= 4 And CBool(r And BAttackBit) Then If r <> BAttackBit Then AttackUnits = AttackUnits + BishopCheck: MinusScore BKSafety, Checked
              End If
            End If
          End If
lblNextBCheck:
        Next
        
        ' distance checks
        AttackUnits = AttackUnits + WKnightChecks * KnightCheck + WBishopChecks * BishopCheck + WRookChecks * RookCheck + WQueenChecks * QueenCheck
          
        ' total attackunits
        AttackUnits = AttackUnits + GetMin(74, BKingAttackersCount * BKingAttackersWeight) + 8 * KingAdjacentZoneAttacksCount + 11 * Abs(BPinnedCnt > 0) - Bonus \ 8 - Abs(WQueenCnt = 0) * 100
  
        ' Penalty for king on open or semi-open file
        If NonPawnMaterial > 9000 And BPawns(FileNum) = 0 And BKingLoc <> BKING_START Then
          If WPawns(FileNum) = 0 Then AttackUnits = AttackUnits + 18 Else AttackUnits = AttackUnits + 9
        End If
  
        AttackUnits = AttackUnits + WPassedPawnAttack * 8
  
        BKSafety.MG = BKSafety.MG - KingDanger(GetMax(GetMin(AttackUnits, 399), 0)).MG
        
      End If
    End If
  

  
  End If ' Endgame

  '--- Endgame King distance to best pawn
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
    '-- Bonus for double pawn control as center fields  E4,D4, E5, D5
    If Board(SQ_E4 + SQ_DOWN_LEFT) = WPAWN And Board(SQ_E4 + SQ_DOWN_RIGHT) = WPAWN Then AddScoreVal WPos, 35, 0
    If Board(SQ_E5 + SQ_DOWN_LEFT) = WPAWN And Board(SQ_E5 + SQ_DOWN_RIGHT) = WPAWN Then AddScoreVal WPos, 35, 0
    If Board(SQ_D4 + SQ_DOWN_LEFT) = WPAWN And Board(SQ_D4 + SQ_DOWN_RIGHT) = WPAWN Then AddScoreVal WPos, 35, 0
    If Board(SQ_D5 + SQ_DOWN_LEFT) = WPAWN And Board(SQ_D5 + SQ_DOWN_RIGHT) = WPAWN Then AddScoreVal WPos, 35, 0

    If Board(SQ_E4 + SQ_UP_LEFT) = BPAWN And Board(SQ_E4 + SQ_UP_RIGHT) = BPAWN Then AddScoreVal BPos, 35, 0
    If Board(SQ_E5 + SQ_UP_LEFT) = BPAWN And Board(SQ_E5 + SQ_UP_RIGHT) = BPAWN Then AddScoreVal BPos, 35, 0
    If Board(SQ_D4 + SQ_UP_LEFT) = BPAWN And Board(SQ_D4 + SQ_UP_RIGHT) = BPAWN Then AddScoreVal BPos, 35, 0
    If Board(SQ_D5 + SQ_UP_LEFT) = BPAWN And Board(SQ_D5 + SQ_UP_RIGHT) = BPAWN Then AddScoreVal BPos, 35, 0

  End If

  ' add kings to attack array
  For i = 0 To 7
    Offset = QueenOffsets(i)
    Target = WKingLoc + Offset: Piece = Board(Target)
    If Piece <> FRAME And Piece <> NO_PIECE Then
      WAttack(Target) = WAttack(Target) Xor KAttackBit
      If BAttack(Target) = 0 And Piece Mod 2 = 0 Then AddThreat COL_WHITE, PieceType(Piece), PT_KING, WKingLoc, Target
    End If
    Target = BKingLoc + Offset: Piece = Board(Target)
    If Piece <> FRAME And Piece <> NO_PIECE Then
      BAttack(Target) = BAttack(Target) Xor KAttackBit
      If WAttack(Target) = 0 And Piece Mod 2 = 1 Then AddThreat COL_BLACK, PieceType(Piece), PT_KING, BKingLoc, Target
    End If
  Next

  '
  '--- Eval threats -------------------------------------------
  '
  CalcThreats  ' in WThreat and BThreat

  ' Trapped bishops at a7/h7, a2/h2
  If WBishopCnt > 0 Then
    ' white bishop not defended trapped at A7 by black pawn B6 (or if pawn can move to B6)
    If Board(SQ_A7) = WBISHOP Then
      If BAttack(SQ_B6) > 0 And WAttack(SQ_A7) = 0 Then
       If Board(SQ_B6) = BPAWN Or (Not bWhiteToMove And Board(SQ_B6) = NO_PIECE And Board(SQ_B7) = BPAWN) Then
        AddScoreVal BThreat, ValueB \ 3, ValueB \ 4
       End If
      End If
    End If
    If Board(SQ_H7) = WBISHOP Then
     If BAttack(SQ_G6) > 0 And WAttack(SQ_H7) = 0 Then
      If Board(SQ_G6) = BPAWN Or (Not bWhiteToMove And Board(SQ_G6) = NO_PIECE And Board(SQ_G7) = BPAWN) Then
        AddScoreVal BThreat, ValueB \ 3, ValueB \ 4
      End If
     End If
    End If
  End If
  If BBishopCnt > 0 Then
    If Board(SQ_A2) = BBISHOP Then
     If WAttack(SQ_B3) > 0 And BAttack(SQ_A2) = 0 Then
      If Board(SQ_B3) = WPAWN Or (bWhiteToMove And Board(SQ_B3) = NO_PIECE And Board(SQ_B2) = WPAWN) Then
        AddScoreVal WThreat, ValueB \ 3, ValueB \ 4
      End If
     End If
    End If
    If Board(SQ_H2) = BBISHOP Then
     If WAttack(SQ_G3) > 0 And BAttack(SQ_H2) = 0 Then
      If Board(SQ_G3) = WPAWN Or (bWhiteToMove And Board(SQ_G3) = NO_PIECE And Board(SQ_G2) = WPAWN) Then
        AddScoreVal WThreat, ValueB \ 3, ValueB \ 4
      End If
     End If
    End If
  End If

  '--- Passed pawns (white and black). done here because full attack info is needed
  WPassedPawnAttack = 0: BPassedPawnAttack = 0
  For a = 1 To PassedPawnsCnt
    Square = PassedPawns(a): FileNum = File(Square): RankNum = Rank(Square)
    MBonus = 0: EBonus = 0
    
    If PieceColor(Board(Square)) = COL_WHITE Then
      ' White piece
      OwnCol = COL_WHITE: OppCol = COL_BLACK: MoveUp = 10
      RelRank = RankNum: RelRank8 = 8: OwnKingLoc = WKingLoc: OppKingLoc = BKingLoc
      ' Attack Opp King?
      If Abs(FileNum - File(OppKingLoc)) <= 2 Then WPassedPawnAttack = WPassedPawnAttack + 1
      
      If WBishopCnt > 0 Then    ' Bishop with same color as promote square? ( not SF logic )
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
      OwnCol = COL_BLACK: OppCol = COL_WHITE: MoveUp = -10
      ' Black piece
      RelRank = (9 - RankNum): RelRank8 = 1: OwnKingLoc = BKingLoc: OppKingLoc = WKingLoc
      If Abs(FileNum - File(OppKingLoc)) <= 2 Then BPassedPawnAttack = BPassedPawnAttack + 1
      
      If BBishopCnt > 0 Then  ' Bishop with same color as promote square?
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
    MBonus = PassedPawnRankBonus(r).MG
    EBonus = PassedPawnRankBonus(r).EG

    ' Bonus based in rank ' SF6
    If rr <> 0 Then
      BlockSq = Square + MoveUp
      If Board(BlockSq) <> FRAME Then
        '  Adjust bonus based on the king's proximity
        EBonus = EBonus + MaxDistance(BlockSq, OppKingLoc) * 5 * rr - MaxDistance(BlockSq, OwnKingLoc) * 2 * rr
        'If blockSq is not the queening square then consider also a second push
        If RelRank <> 7 Then EBonus = EBonus - MaxDistance(BlockSq + MoveUp, OwnKingLoc) * rr

        'If the pawn is free to advance, then increase the bonus
        If Board(BlockSq) = NO_PIECE Then
          k = 0: bAllDefended = True: BlockSqDefended = True: UnsafeSq = False: BlockSqUnsafe = False: bAttackedFromBehind = False: bDefendedFromBehind = False
          
          ' Rook or Queen attacking/defending from behind
           If CBool(BAttack(Square) And RAttackBit) Or CBool(BAttack(Square) And QAttackBit) Or CBool(WAttack(Square) And RAttackBit) Or CBool(WAttack(Square) And QAttackBit) Then
            For RankPath = RelRank - 1 To 1 Step -1
              sq = Square + (RankPath - RelRank) * MoveUp
              Select Case Board(sq)
              Case NO_PIECE:
              Case BROOK, BQUEEN:
                If OwnCol = COL_WHITE Then
                  bAttackedFromBehind = True: UnsafeSq = True: BlockSqUnsafe = True
                Else
                  bDefendedFromBehind = True: bAllDefended = True
                End If
                Exit For
              Case WROOK, WQUEEN:
                If OwnCol = COL_BLACK Then
                  bAttackedFromBehind = True: UnsafeSq = True: BlockSqUnsafe = True
                Else
                  bDefendedFromBehind = True: bAllDefended = True
                End If
                Exit For
              Case Else:
                Exit For
              End Select
            Next
          End If
          
          For RankPath = RelRank + 1 To 8
            sq = Square + (RankPath - RelRank) * MoveUp
            '--- defended? King attacks not in WAttack() > extra check
              If AttackByCol(OwnCol, sq) = 0 And MaxDistance(sq, OwnKingLoc) > 0 Then
                bAllDefended = False: If sq = BlockSq Then BlockSqDefended = False
              End If
              If PieceColor(Board(sq)) = OppCol Or AttackByCol(OppCol, sq) > 0 Or MaxDistance(sq, OppKingLoc) = 1 Then
                UnsafeSq = True: If sq = BlockSq Then BlockSqUnsafe = True
              End If
          Next RankPath
              
          If Not UnsafeSq Then
            k = 15
          ElseIf Not BlockSqUnsafe Then
            k = 9
          Else
            k = 0
          End If
          If bAllDefended Then
            k = k + 6
          ElseIf BlockSqDefended Then
            k = k + 4
          End If
          If k > 0 Then MBonus = MBonus + k * rr: EBonus = EBonus + k * rr
        Else
          If PieceColor(Board(BlockSq)) = OwnCol Then MBonus = MBonus + rr * 3 + r * 2 + 3: EBonus = EBonus + rr + r * 2
        End If
      End If
      
    End If ' rr>0

    If OwnCol = COL_WHITE Then
      If WPawnCnt < BPawnCnt Then EBonus = EBonus + EBonus \ 4
      MBonus = MBonus + PassedPawnFileBonus(FileNum).MG: EBonus = EBonus + PassedPawnFileBonus(FileNum).EG
      AddScoreVal WPassed, MBonus, EBonus
      If 1000 + EBonus > WBestPawnVal Then WBestPawn = Square: WBestPawnVal = 1000 + EBonus ' new best pawn
      If bEvalTrace Then WriteTrace "WPassed: " & LocCoord(Square) & ">" & MBonus & ", " & EBonus
      
    ElseIf OwnCol = COL_BLACK Then
      If BPawnCnt < WPawnCnt Then EBonus = EBonus + EBonus \ 4
      MBonus = MBonus + PassedPawnFileBonus(FileNum).MG: EBonus = EBonus + PassedPawnFileBonus(FileNum).EG
      AddScoreVal BPassed, MBonus, EBonus
      If 1000 + EBonus > BBestPawnVal Then BBestPawn = Square: BBestPawnVal = 1000 + EBonus ' new best pawn
      If bEvalTrace Then WriteTrace "BPassed: " & LocCoord(Square) & ">" & MBonus & ", " & EBonus
    End If
  Next a
            
  '---<<< end  Passed pawn
  
  '---  Penalty for pawns on same color square of bishop
 ' r = (WPawnCntOnWhiteSq - WPawnCntOnBlackSq) * WBishopsOnWhiteSq + (WPawnCntOnBlackSq - WPawnCntOnWhiteSq) * WBishopsOnBlackSq
 ' If r > 0 Then
 '   AddScoreVal WPos, -8 * r, -12 * r
 ' End If
  
 ' r = (BPawnCntOnWhiteSq - BPawnCntOnBlackSq) * BBishopsOnWhiteSq + (BPawnCntOnBlackSq - BPawnCntOnWhiteSq) * BBishopsOnBlackSq
 ' If r > 0 Then
 '   AddScoreVal BPos, -8 * r, -12 * r
 ' End If

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
  
  '---<<< Pawn Islands ---


  '
  '--- Calculate weights and total eval
  '
  Dim TotalScore      As TScore, MatEval As Long, TradeEval As Long, PosEval As Long, PawnStructEval As Long
  Dim PassedPawnsEval As Long, MobilityEval As Long, KingSafetyEval As Long, ThreatEval As Long

  ' Piece values were set in SetGamePhase
  MatEval = (WQueenCnt - BQueenCnt) * ValueQ + (WRookCnt - BRookCnt) * ValueR + (WKnightCnt - BKnightCnt) * ValueN + (WBishopCnt - BBishopCnt) * ValueB + (WPawnCnt - BPawnCnt) * ValueP

  '--- Material Imbalance / Score trades
  If MatEval = 0 Then TradeEval = 0 Else TradeEval = Imbalance() ' SF6
  'TradeEval = Eval100ToSF(ScaleScore(MaterialBalance()))  ' Protector

  TotalScore.MG = WPos.MG - BPos.MG: TotalScore.EG = WPos.EG - BPos.EG ' no Weights
  PosEval = (ScaleScore(TotalScore) * PiecePosScaleFactor) \ 100&

  TotalScore.MG = WPawnStruct.MG - BPawnStruct.MG: TotalScore.EG = WPawnStruct.EG - BPawnStruct.EG: CalcWeight TotalScore, Weights(PawnStructure_Weight)
  PawnStructEval = (ScaleScore(TotalScore) * PawnStructScaleFactor) \ 100&

  TotalScore.MG = WPassed.MG - BPassed.MG: TotalScore.EG = WPassed.EG - BPassed.EG: CalcWeight TotalScore, Weights(PassedPawns_Weight)
  PassedPawnsEval = (ScaleScore(TotalScore) * PassedPawnsScaleFactor) \ 100&

  TotalScore.MG = WMobility.MG - BMobility.MG: TotalScore.EG = WMobility.EG - BMobility.EG ': CalcWeight TotalScore, Weights(Mobility_Weight)
  MobilityEval = (ScaleScore(TotalScore) * MobilityScaleFactor) \ 100&

  ' King weights are in KingDanger array. different weights for defending computer king / attacking opp king
  If bCompIsWhite Then
    WKingScaleFactor = CompKingDefScaleFactor: BKingScaleFactor = OppKingAttScaleFactor
  Else
    BKingScaleFactor = CompKingDefScaleFactor: WKingScaleFactor = OppKingAttScaleFactor
  End If
  KingSafetyEval = (ScaleScore(WKSafety) * WKingScaleFactor) \ 100& - (ScaleScore(BKSafety) * BKingScaleFactor) \ 100&

  TotalScore.MG = WThreat.MG - BThreat.MG: TotalScore.EG = WThreat.EG - BThreat.EG ': CalcWeight TotalScore, Weights(Threats_Weight)
  ThreatEval = (ScaleScore(TotalScore) * ThreatsScaleFactor) \ 100&

  '--- Add all to eval score (SF based scaling:  Eval*100/SFPawnEndGameValue= 100 centipawns =1 pawn)
  '--- Example: Eval=258 => 1.00 pawn
  Eval = MatEval + TradeEval + PosEval + PawnStructEval + PassedPawnsEval + MobilityEval + KingSafetyEval + ThreatEval

  ' Keep more pawns when attacking
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
    WriteTrace "Material: " & EvalSFTo100(MatEval) & " => Q:" & (WQueenCnt - BQueenCnt) * ValueQ & ", R:" & (WRookCnt - BRookCnt) * ValueR & ", N:" & (WKnightCnt - BKnightCnt) * ValueN & ", B:" & (WBishopCnt - BBishopCnt) * ValueB & ", P:" & (WPawnCnt - BPawnCnt) * ValueP
    WriteTrace "Trades  : " & EvalSFTo100(TradeEval)
    WriteTrace "Position: " & EvalSFTo100(PosEval) & "  => W" & ShowScore(WPos) & ", B" & ShowScore(BPos)
    WriteTrace "PawnStru: " & EvalSFTo100(PawnStructEval) & " => W" & ShowScore(WPawnStruct) & ", B" & ShowScore(BPawnStruct)
    WriteTrace "PassedPw: " & EvalSFTo100(PassedPawnsEval) & " => W" & ShowScore(WPassed) & ", B" & ShowScore(BPassed)
    WriteTrace "Mobility: " & EvalSFTo100(MobilityEval) & " => W(" & WMobility.MG & "," & WMobility.EG & "), B(" & BMobility.MG & "," & BMobility.EG & ")"
    WriteTrace "KSafety : " & EvalSFTo100(KingSafetyEval) & " => W(" & WKSafety.MG & "," & WKSafety.EG & "), B(" & BKSafety.MG & "," & BKSafety.EG & ")"
    WriteTrace "Threats : " & EvalSFTo100(ThreatEval) & " => W(" & WThreat.MG & "," & WThreat.EG & "), B(" & BThreat.MG & "," & BThreat.EG & ")"
    WriteTrace "Eval    : " & Eval & "  (" & EvalSFTo100(Eval) & "cp)"
    WriteTrace "-----------------"
    bTimeExit = True
  End If

  If Not bWhiteToMove Then Eval = -Eval '--- Invert for black

  Eval = Eval + 17 ' Tempo for side to move
End Function
'---------------------------------
'-------- END OF EVAL ------------
'---------------------------------

Private Function IsMaterialDraw() As Boolean
  '( Protector logic )
  IsMaterialDraw = False
  If WPawnCnt + BPawnCnt = 0 Then ' no pawns
    '---  no heavies Q/R */
    If WRookCnt = 0 And BRookCnt = 0 And WQueenCnt = 0 And BQueenCnt = 0 Then
    
      If BBishopCnt = 0 And WBishopCnt = 0 Then
        '---  only knights */
        '---  it pretty safe to say this is a draw */
        If WKnightCnt < 3 And BKnightCnt < 3 Then IsMaterialDraw = True: Exit Function
      
      ElseIf WKnightCnt = 0 And BKnightCnt = 0 Then
        '---  only bishops */
        '---  not a draw if one side two other side zero
        '---  else its always a draw                     */
        If Abs(WBishopCnt - BBishopCnt) < 2 Then IsMaterialDraw = True: Exit Function
      
      ElseIf (WKnightCnt < 3 And WBishopCnt = 0) Or (WBishopCnt = 1 And WKnightCnt = 0) Then
        '---  we cant win, but can black? */
        If (BKnightCnt < 3 And BBishopCnt = 0) Or (BBishopCnt = 1 And BKnightCnt = 0) Then IsMaterialDraw = True: Exit Function '---  guess not */
      End If
  
    ElseIf WQueenCnt = 0 And BQueenCnt = 0 Then
    
      If WRookCnt = 1 And BRookCnt = 1 Then
        '---  rooks equal */
        '---  one minor difference max: a draw too usually */
        If (WKnightCnt + WBishopCnt) < 2 And (BKnightCnt + BBishopCnt) < 2 Then IsMaterialDraw = True: Exit Function
      ElseIf (WRookCnt = 1 And BRookCnt = 0) Then
        '---  one rook */
        '---  draw if no minors to support AND minors to defend  */
        If (WKnightCnt + WBishopCnt = 0) And ((BKnightCnt + BBishopCnt = 1) Or (BKnightCnt + BBishopCnt = 2)) Then IsMaterialDraw = True: Exit Function
      ElseIf BRookCnt = 1 And WRookCnt = 0 Then
        '---  one rook */
        '---  draw if no minors to support AND minors to defend  */
        If (BKnightCnt + BBishopCnt = 0) And ((WKnightCnt + WBishopCnt = 1) Or (WKnightCnt + WBishopCnt = 2)) Then IsMaterialDraw = True: Exit Function
      End If
    End If
  End If
 
End Function

Public Function CalcWeight(Score As TScore, Weight As TScore)
  Score.MG = (Score.MG * Weight.MG) \ 256&:  Score.EG = (Score.EG * Weight.EG) \ 256&
End Function

Public Function AdvancedPawnPush(ByVal Piece As Integer, _
                                 ByVal Target As Integer) As Boolean
  AdvancedPawnPush = False
  If Piece = WPAWN Then
    Select Case Rank(Target)
      Case 7, 8: AdvancedPawnPush = True
      Case 6:
        '--- if enemy in front and no enemy pawns left or right
        If (Board(Target + SQ_UP) = NO_PIECE Or Board(Target + SQ_UP) Mod 2 = 1) And Board(Target + SQ_UP_LEFT) <> BPAWN And Board(Target + SQ_UP_RIGHT) <> BPAWN Then AdvancedPawnPush = True
    End Select
  ElseIf Piece = BPAWN Then
    Select Case Rank(Target)
      Case 1, 2: AdvancedPawnPush = True
      Case 3:
        If (Board(Target + SQ_DOWN) = NO_PIECE Or Board(Target + SQ_DOWN) Mod 2 = 0) And Board(Target + SQ_DOWN_LEFT) <> WPAWN And Board(Target + SQ_DOWN_RIGHT) <> WPAWN Then AdvancedPawnPush = True
    End Select
  End If
End Function

Public Function PieceSquareValDiff(ByVal Piece As Integer, _
                                   ByVal From As Integer, _
                                   ByVal Target As Integer) As Integer
  '--- Score difference in piece square table for moving a piece
  PieceSquareValDiff = 0
  
  If bEndgame Then
    Select Case Piece
      Case NO_PIECE
      Case WPAWN
        PieceSquareValDiff = PsqtWP(Target).EG - PsqtWP(From).EG
      Case BPAWN
        PieceSquareValDiff = PsqtBP(Target).EG - PsqtBP(From).EG
      Case WKNIGHT
        PieceSquareValDiff = PsqtWN(Target).EG - PsqtWN(From).EG
      Case BKNIGHT
        PieceSquareValDiff = PsqtBN(Target).EG - PsqtBN(From).EG
      Case WBISHOP
        PieceSquareValDiff = PsqtWB(Target).EG - PsqtWB(From).EG
      Case BBISHOP
        PieceSquareValDiff = PsqtBB(Target).EG - PsqtBB(From).EG
      Case WROOK
        PieceSquareValDiff = PsqtWR(Target).EG - PsqtWR(From).EG
      Case BROOK
        PieceSquareValDiff = PsqtBR(Target).EG - PsqtBR(From).EG
      Case WQUEEN
        PieceSquareValDiff = PsqtWQ(Target).EG - PsqtWQ(From).EG
      Case BQUEEN
        PieceSquareValDiff = PsqtBQ(Target).EG - PsqtBQ(From).EG
      Case WKING
        PieceSquareValDiff = PsqtWK(Target).EG - PsqtWK(From).EG
      Case BKING
        PieceSquareValDiff = PsqtBK(Target).EG - PsqtBK(From).EG
    End Select
  Else
    Select Case Piece
      Case NO_PIECE
      Case WPAWN
        PieceSquareValDiff = PsqtWP(Target).MG - PsqtWP(From).MG
      Case BPAWN
        PieceSquareValDiff = PsqtBP(Target).MG - PsqtBP(From).MG
      Case WKNIGHT
        PieceSquareValDiff = PsqtWN(Target).MG - PsqtWN(From).MG
      Case BKNIGHT
        PieceSquareValDiff = PsqtBN(Target).MG - PsqtBN(From).MG
      Case WBISHOP
        PieceSquareValDiff = PsqtWB(Target).MG - PsqtWB(From).MG
      Case BBISHOP
        PieceSquareValDiff = PsqtBB(Target).MG - PsqtBB(From).MG
      Case WROOK
        PieceSquareValDiff = PsqtWR(Target).MG - PsqtWR(From).MG
      Case BROOK
        PieceSquareValDiff = PsqtBR(Target).MG - PsqtBR(From).MG
      Case WQUEEN
        PieceSquareValDiff = PsqtWQ(Target).MG - PsqtWQ(From).MG
      Case BQUEEN
        PieceSquareValDiff = PsqtBQ(Target).MG - PsqtBQ(From).MG
      Case WKING
        PieceSquareValDiff = PsqtWK(Target).MG - PsqtWK(From).MG
      Case BKING
        PieceSquareValDiff = PsqtBK(Target).MG - PsqtBK(From).MG
    End Select
  End If
End Function


Public Function PieceSquareVal(ByVal Piece As Integer, ByVal Square As Integer) As Integer
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
Dim Piece As Integer, Target As Integer
  For Piece = 1 To 16
    For Target = SQ_A1 To SQ_H8
        bEndgame = False
        PsqVal(0, Piece, Target) = PieceSquareVal(Piece, Target)
        bEndgame = True
        PsqVal(1, Piece, Target) = PieceSquareVal(Piece, Target)
    Next
  Next
End Sub

Public Function KingPressure(ByVal DefendColor As enumColor) As Long
  '--- Checks attacks at squares around king (own logic (slow), no longer used in search,  only used for root moves )
  Dim i As Long, j As Long, OwnColor As enumColor, OppColor As enumColor, Loc As Integer, StartLoc As Integer, ThreatCnt As Integer

  If DefendColor = COL_WHITE Then
    OwnColor = COL_WHITE: OppColor = COL_BLACK
  Else
    OwnColor = COL_BLACK:: OppColor = COL_WHITE
  End If
 
  KingPressure = 0
 
  If OwnColor = COL_WHITE Then
    If Rank(WKingLoc) <= 3 Then
      StartLoc = WKingLoc
      If File(StartLoc) = 8 Then StartLoc = StartLoc - 1 '-- H8 > G8 to get F6 too
      If File(StartLoc) = 1 Then StartLoc = StartLoc + 1 '-- A1 > A2 to get F6 too
      For i = -1 To 1
        For j = 0 To 20 Step 10
          Loc = StartLoc + j + i
          If Loc <> WKingLoc And Board(Loc) <> FRAME Then
            If Board(Loc) = BPAWN Then
              KingPressure = KingPressure - 10
            ElseIf j = 10 And Board(Loc + SQ_UP) = NO_PIECE And Board(Loc + 20) = NO_PIECE Then
              KingPressure = KingPressure + 10: If i = 0 Then KingPressure = KingPressure + 20
              If Board(Loc + 30) = NO_PIECE Then KingPressure = KingPressure + 20: If i = 0 Then KingPressure = KingPressure + 50
            End If
            Select Case KingAttackedCnt(Loc, OwnColor, ThreatCnt)
              Case 0:
                KingPressure = KingPressure + ThreatCnt * 3
              Case 1:
                KingPressure = KingPressure + 20 + ThreatCnt * 3: If j < 20 Then KingPressure = KingPressure + 20 + ThreatCnt * 8
              Case 2, 3, 4, 5:
                KingPressure = KingPressure + 60 + ThreatCnt * 3: If j < 20 Then KingPressure = KingPressure + 10 + ThreatCnt * 8
            End Select
          End If
        Next j
      Next i
    End If
  Else
    If Rank(BKingLoc) >= 6 Then
      StartLoc = BKingLoc
      If File(StartLoc) = 8 Then StartLoc = StartLoc - 1 '-- H8 > G8 to get F6 too
      If File(StartLoc) = 1 Then StartLoc = StartLoc + 1 '-- A1 > A2 to get F6 too
      For i = -1 To 1
        For j = 0 To -20 Step -10
          Loc = StartLoc + j + i
          If Loc <> BKingLoc And Board(Loc) <> FRAME Then
            If Board(Loc) = BPAWN Then
              KingPressure = KingPressure - 10
            ElseIf j = -10 And Board(Loc + SQ_DOWN) = NO_PIECE And Board(Loc - 20) = NO_PIECE Then
              KingPressure = KingPressure + 10: If i = 0 Then KingPressure = KingPressure + 20
              If Board(Loc - 30) = NO_PIECE Then KingPressure = KingPressure + 20: If i = 0 Then KingPressure = KingPressure + 50
            End If
            Select Case KingAttackedCnt(Loc, OwnColor, ThreatCnt)
              Case 0: KingPressure = KingPressure + ThreatCnt * 3
              Case 1:
                KingPressure = KingPressure + 20 + ThreatCnt * 3: If j < 20 Then KingPressure = KingPressure + 20 + ThreatCnt * 8
              Case 2, 3, 4, 5:
                KingPressure = KingPressure + 60 + ThreatCnt * 3: If j < 20 Then KingPressure = KingPressure + 10 + ThreatCnt * 8
            End Select
          End If
        Next j
      Next i
    End If
  End If
End Function

Private Function AttackByCol(Col As Integer, Square As Integer) As Integer
  If Col = COL_WHITE Then AttackByCol = WAttack(Square) Else AttackByCol = BAttack(Square)
End Function

Public Sub AddPawnThreat(Score As TScore, _
                         ByVal HangCol As enumColor, _
                         ByVal PieceType As enumPieceType, _
                         ByVal Square As Integer)
  'SF6:  const Score ThreatenedByPawn[PIECE_TYPE_NB] = {
  '         S(0, 0), S(0, 0), S(107, 138), S(84, 122), S(114, 203), S(121, 217)
  '      const Score ThreatenedByHangingPawn = S(40, 60);
 
  '--- attack by black pawn?
  If HangCol = COL_WHITE Then
    If Board(Square + SQ_UP_LEFT) = BPAWN Then
      If Board(Square + SQ_UP_LEFT + SQ_UP_LEFT) = BPAWN Or Board(Square + SQ_UP_LEFT + SQ_UP_RIGHT) = BPAWN Then
        AddScore Score, ThreatenedByPawn(PieceType)
      Else
        If Not bWhiteToMove Then AddScore Score, ThreatenedByHangingPawn Else AddScoreVal Score, ThreatenedByHangingPawn.MG \ 2, ThreatenedByHangingPawn.EG \ 2 ' escape option?
      End If
     
    ElseIf Board(Square + SQ_UP_RIGHT) = BPAWN Then
      If Board(Square + SQ_UP_RIGHT + SQ_UP_LEFT) = BPAWN Or Board(Square + SQ_UP_RIGHT + SQ_UP_RIGHT) = BPAWN Then
        AddScore Score, ThreatenedByPawn(PieceType)
      Else
        If Not bWhiteToMove Then AddScore Score, ThreatenedByHangingPawn Else AddScoreVal Score, ThreatenedByHangingPawn.MG \ 2, ThreatenedByHangingPawn.EG \ 2 ' escape option?
      End If
    End If
 
  Else ' attack by white pawn?
   
    If Board(Square + SQ_DOWN_LEFT) = WPAWN Then
      If Board(Square + SQ_DOWN_LEFT + SQ_DOWN_LEFT) = WPAWN Or Board(Square + SQ_DOWN_LEFT + SQ_DOWN_RIGHT) = WPAWN Then
        AddScore Score, ThreatenedByPawn(PieceType)
      Else
        If bWhiteToMove Then AddScore Score, ThreatenedByHangingPawn Else AddScoreVal Score, ThreatenedByHangingPawn.MG \ 2, ThreatenedByHangingPawn.EG \ 2  ' escape option?
      End If
    ElseIf Board(Square + SQ_DOWN_RIGHT) = WPAWN Then
      If Board(Square + SQ_DOWN_RIGHT + SQ_DOWN_LEFT) = WPAWN Or Board(Square + SQ_DOWN_RIGHT + SQ_DOWN_RIGHT) = WPAWN Then
        AddScore Score, ThreatenedByPawn(PieceType)
      Else
        If bWhiteToMove Then AddScore Score, ThreatenedByHangingPawn Else AddScoreVal Score, ThreatenedByHangingPawn.MG \ 2, ThreatenedByHangingPawn.EG \ 2  ' escape option?
      End If
    End If
  End If
End Sub

Public Sub AddThreat(ByVal HangCol As enumColor, _
                     ByVal HangPieceType As enumPieceType, _
                     ByVal AttPieceType As enumPieceType, _
                     ByVal AttackingSquare As Integer, _
                     ByVal AttackedSquare As Integer)
  ' Add threat to threat list. calculate score later when full attack array data is available
  ThreatCnt = ThreatCnt + 1
  With ThreatList(ThreatCnt)
    .HangCol = HangCol
    .HangPieceType = HangPieceType
    .AttPieceType = AttPieceType
    .AttackingSquare = AttackingSquare
    .AttackedSquare = AttackedSquare
  End With
End Sub

Public Sub CalcThreats()
  Dim i As Integer, IsMinor As Boolean, IsMajor As Boolean, Defended As Boolean, WToMoveFactor As Long, BToMoveFactor As Long
  If ThreatCnt = 0 Then Exit Sub
  
  ' if threat but side can move piece out of danger then reduced threat value > asymmetric eval result!
  If bWhiteToMove Then
    WToMoveFactor = 75: BToMoveFactor = 100
  Else
    BToMoveFactor = 75: WToMoveFactor = 100
  End If
  
  For i = 1 To ThreatCnt
    With ThreatList(i)
      IsMinor = False: IsMajor = False
      Select Case .AttPieceType
        Case PT_BISHOP, PT_KNIGHT: IsMinor = True ' minor
        Case PT_ROOK, PT_QUEEN, PT_KING: IsMajor = True ' major
        Case Else
          GoTo lblNext
      End Select
         
      ' Add a bonus according to the kind of attacking pieces
      If .HangCol = COL_WHITE Then
        If .HangPieceType = PT_PAWN And WAttack(.AttackedSquare) <> 0 Then GoTo lblNext ' ignore defended pawns
        Defended = CBool(WAttack(.AttackedSquare) And PAttackBit) ' Defended by pawn
        If Defended Then
          If IsMinor Then
            AddScore BThreat, ScaleScore100(ThreatDefendedMinor(.HangPieceType), WToMoveFactor)
          ElseIf .AttPieceType = PT_ROOK Then
            AddScore BThreat, ScaleScore100(ThreatDefendedMajor(.HangPieceType), WToMoveFactor)
          End If
        Else ' weak: Attacked and not defended by pawn
          If IsMinor Then
            AddScore BThreat, ScaleScore100(ThreatWeakMinor(.HangPieceType), WToMoveFactor)
          ElseIf IsMajor Then
            AddScore BThreat, ScaleScore100(ThreatWeakMajor(.HangPieceType), WToMoveFactor)
          End If
          If WAttack(.AttackedSquare) = 0 Then ' not defended by any piece: hanging
            AddScore BThreat, ScaleScore100(Hanging, WToMoveFactor)
          Else
            ' Defended piece less valuable than attacker? Simple SEE . reduce threat penalty
            If PieceTypeValue(.AttPieceType) > PieceTypeValue(.HangPieceType) Then
              MinusScore BThreat, ScaleScore100(ThreatWeakMinor(.HangPieceType), WToMoveFactor \ 3)
            End If
          End If
        End If
  
      Else ' Black
        If .HangPieceType = PT_PAWN And BAttack(.AttackedSquare) <> 0 Then GoTo lblNext ' ignore defended pawns
        Defended = CBool(BAttack(.AttackedSquare) And PAttackBit)
        If Defended Then
          If IsMinor Then
            AddScore WThreat, ScaleScore100(ThreatDefendedMinor(.HangPieceType), BToMoveFactor)
          ElseIf .AttPieceType = PT_ROOK Then
            AddScore WThreat, ScaleScore100(ThreatDefendedMajor(.HangPieceType), BToMoveFactor)
          End If
        Else ' weak: Attacked and not defended by pawn
          If IsMinor Then
            AddScore WThreat, ScaleScore100(ThreatWeakMinor(.HangPieceType), BToMoveFactor)
          ElseIf IsMajor Then
            AddScore WThreat, ScaleScore100(ThreatWeakMajor(.HangPieceType), BToMoveFactor)
          End If
          If BAttack(.AttackedSquare) = 0 Then
            AddScore WThreat, ScaleScore100(Hanging, BToMoveFactor)
          Else
            ' Defended piece less valuable than attacker? Simple SEE . reduce threat penalty
            If PieceTypeValue(.AttPieceType) > PieceTypeValue(.HangPieceType) Then
              MinusScore WThreat, ScaleScore100(ThreatWeakMinor(.HangPieceType), BToMoveFactor \ 3)
            End If
          End If
        End If
           
      End If
    End With
    
lblNext:
  Next

End Sub

Private Sub AddKingZoneDefendCnt(ByVal DefendBits As Integer, _
                                 ByVal AttackBits As Integer, _
                                 KingZoneDefendCnt As Integer, _
                                 DefendWeight As Integer)
  If DefendBits > 0 Then
    If CBool(DefendBits And QAttackBit) And (AttackBits = 0 Or AttackBits = QAttackBit) Then
      ' Queen move to square is save  to defend king
      KingZoneDefendCnt = KingZoneDefendCnt + 1: DefendWeight = DefendWeight + KingAttackWeights(PT_QUEEN)
    End If
  End If
End Sub

Private Sub AddKingZoneAttackCnt(ByVal AttackBits As Integer, _
                                 KingZoneAttackCnt As Integer, _
                                 AttackWeight As Integer)
  If AttackBits > 0 Then
    If CBool(AttackBits And PAttackBit) Then KingZoneAttackCnt = KingZoneAttackCnt + 1
    If CBool(AttackBits And NAttackBit) Then KingZoneAttackCnt = KingZoneAttackCnt + 1: AttackWeight = AttackWeight + KingAttackWeights(PT_KNIGHT)
    If CBool(AttackBits And BAttackBit) Then KingZoneAttackCnt = KingZoneAttackCnt + 1: AttackWeight = AttackWeight + KingAttackWeights(PT_BISHOP)
    If CBool(AttackBits And RAttackBit) Then KingZoneAttackCnt = KingZoneAttackCnt + 1: AttackWeight = AttackWeight + KingAttackWeights(PT_ROOK)
    If CBool(AttackBits And QAttackBit) Then KingZoneAttackCnt = KingZoneAttackCnt + 1: AttackWeight = AttackWeight + KingAttackWeights(PT_QUEEN)
  End If
End Sub

Public Sub AddWKingAttack(PT As enumPieceType)
  WKingAttackersCount = WKingAttackersCount + 1
  WKingAttackersWeight = WKingAttackersWeight + KingAttackWeights(PT)
End Sub

Public Sub AddBKingAttack(PT As enumPieceType)
  BKingAttackersCount = BKingAttackersCount + 1
  BKingAttackersWeight = BKingAttackersWeight + KingAttackWeights(PT)
End Sub

Public Sub AddWKingDefend(PT As enumPieceType)
  WKingDefendersCount = WKingDefendersCount + 1
  WKingDefendersWeight = WKingDefendersWeight + KingAttackWeights(PT)
End Sub

Public Sub AddBKingDefend(PT As enumPieceType)
  BKingDefendersCount = BKingDefendersCount + 1
  BKingDefendersWeight = BKingDefendersWeight + KingAttackWeights(PT)
End Sub

Public Sub InitKingDangerArr()
  Const MaxSlope As Long = 8700
  Const Peak     As Long = 1280000
  Dim t          As Long, i As Long

  For i = 0 To 399
    t = GetMin(Peak, GetMin(i * i * 27, t + MaxSlope))
    KingDanger(i).MG = t / 1000: CalcWeight KingDanger(i), Weights(KingSafety_Weight)
  Next
End Sub

Public Function InitConnectedPawns()
  ' SF6
  Dim Seed(8) As Long, Opposed As Integer, Phalanx As Integer, Apex As Integer, r As Integer, v As Long, x As Long
  
  ReadLngArr Seed(), 0, 0, 6, 15, 10, 57, 75, 135, 258

  For Opposed = 0 To 1
    For Phalanx = 0 To 1
      For Apex = 0 To 1
        For r = 2 To 7
          If Phalanx > 0 Then x = (Seed(r + 1) - Seed(r)) / 2 Else x = 0
          v = Seed(r) + x
          If Opposed > 0 Then v = v / 2 ' >>  operator for opposed in VB: /2
          If Apex > 0 Then v = v + v / 2
          ConnectedBonus(Opposed, Phalanx, Apex, r).MG = 3 * v / 2
          ConnectedBonus(Opposed, Phalanx, Apex, r).EG = v
        Next
      Next
    Next
  Next

End Function

Public Sub InitImbalance()  ' SF6
  ReadIntArr Linear(), 1852, -162, -1122, -183, 249, -154
  
  ' // pair pawn knight bishop rook queen  OUR PIECES
  ReadIntArr2 QuadraticOurs(), 0, 0 ' Bishop pair
  ReadIntArr2 QuadraticOurs(), PT_PAWN, 39, -1 ' Pawn
  ReadIntArr2 QuadraticOurs(), PT_KNIGHT, 21, 247, 8        ' Knight
  ReadIntArr2 QuadraticOurs(), PT_BISHOP, 0, 107, 11, 0        ' Bishop
  ReadIntArr2 QuadraticOurs(), PT_ROOK, -26, 6, 50, 94, -168                ' Rook
  ReadIntArr2 QuadraticOurs(), PT_QUEEN, -213, 25, 129, 136, -142, -26     ' Queen
  
  ' // pair pawn knight bishop rook queen  THEIR PIECES
  ReadIntArr2 QuadraticTheirs(), 0, 0 ' Bishop pair
  ReadIntArr2 QuadraticTheirs(), PT_PAWN, 33, 0     ' Pawn
  ReadIntArr2 QuadraticTheirs(), PT_KNIGHT, 16, 52, 0             ' Knight
  ReadIntArr2 QuadraticTheirs(), PT_BISHOP, 53, 72, 40, 0               ' Bishop
  ReadIntArr2 QuadraticTheirs(), PT_ROOK, 31, 39, 20, -34, 0                  ' Rook
  ReadIntArr2 QuadraticTheirs(), PT_QUEEN, 86, 93, -38, 137, 291, 0                  ' Queen
  
End Sub

Public Function Imbalance() As Long ' SF6
  Dim v As Long
  ImbPieceCount(COL_WHITE, 0) = Abs(WBishopCnt > 1)  ' index 0 used for bishop pair
  ImbPieceCount(COL_BLACK, 0) = Abs(BBishopCnt > 1)  ' index 0 used for bishop pair
  
  ImbPieceCount(COL_WHITE, PT_PAWN) = WPawnCnt
  ImbPieceCount(COL_BLACK, PT_PAWN) = BPawnCnt
  
  ImbPieceCount(COL_WHITE, PT_KNIGHT) = WKnightCnt
  ImbPieceCount(COL_BLACK, PT_KNIGHT) = BKnightCnt
  
  ImbPieceCount(COL_WHITE, PT_BISHOP) = WBishopCnt
  ImbPieceCount(COL_BLACK, PT_BISHOP) = BBishopCnt
 
  ImbPieceCount(COL_WHITE, PT_ROOK) = WRookCnt
  ImbPieceCount(COL_BLACK, PT_ROOK) = BRookCnt

  ImbPieceCount(COL_WHITE, PT_QUEEN) = WQueenCnt
  ImbPieceCount(COL_BLACK, PT_QUEEN) = BQueenCnt

  v = (ColImbalance(COL_WHITE) - ColImbalance(COL_BLACK)) \ 16
  Imbalance = v
End Function

Public Function ColImbalance(Col As enumColor) As Long
  Dim Bonus As Long, pt1 As Integer, pt2 As Integer, Us As Integer, Them As Integer, v As Long
  
  If Col = COL_WHITE Then
    Us = COL_WHITE: Them = COL_BLACK
  Else
    Us = COL_BLACK: Them = COL_WHITE
  End If
  
  For pt1 = 0 To PT_QUEEN
    If ImbPieceCount(Us, pt1) > 0 Then
      v = Linear(pt1)
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

Public Function EvalSFTo100(Eval As Long) As Long
  If Abs(Eval) < MATE_IN_MAX_PLY Then EvalSFTo100 = (Eval * 100&) / CLng(ScorePawn.EG) Else EvalSFTo100 = Eval
End Function

Public Function Eval100ToSF(Eval As Long) As Long
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

Public Sub ScoreToZero(Score As TScore)
  Score.MG = 0: Score.EG = 0
End Sub

Public Function ScaleScore100(Score As TScore, ScaleVal As Long) As TScore
  ScaleScore100.MG = (Score.MG * ScaleVal) \ 100&: ScaleScore100.EG = (Score.EG * ScaleVal) \ 100&
End Function

Public Function ShowScore(Score As TScore) As String
  ' show MG, EG Score as text
  ShowScore = "(" & CStr(Score.MG) & "," & CStr(Score.EG) & ")=" & ScaleScore(Score)
End Function

Public Function PieceSq(Side As enumColor, SearchPieceType As enumPieceType) As Integer
  Dim a As Integer, p As Integer
  
  For a = 1 To NumPieces
    p = Board(Pieces(a)): If PieceType(p) = SearchPieceType And PieceColor(p) = Side Then PieceSq = Pieces(a): Exit Function
  Next
End Function

Public Function Eval_KRKP() As Long
  Dim WKSq As Integer, BKSq As Integer, RookSq As Integer, PawnSq As Integer, StrongSide As enumColor, WeakSide As enumColor
  Dim StrongKingLoc As Integer, WeakKingLoc As Integer, QueeningSq As Integer, Result As Long, SideToMove As enumColor
  
  If WMaterial > BMaterial Then
    StrongSide = COL_WHITE: WeakSide = COL_BLACK: StrongKingLoc = WKingLoc: WeakKingLoc = BKingLoc
  Else
    StrongSide = COL_BLACK: WeakSide = COL_WHITE: StrongKingLoc = BKingLoc: WeakKingLoc = WKingLoc
  End If
  If bWhiteToMove Then SideToMove = COL_WHITE Else SideToMove = COL_BLACK
  
  WKSq = RelativeSq(StrongSide, StrongKingLoc)
  BKSq = RelativeSq(StrongSide, WeakKingLoc)
  RookSq = RelativeSq(StrongSide, PieceSq(StrongSide, PT_ROOK))
  PawnSq = RelativeSq(WeakSide, PieceSq(WeakSide, PT_PAWN))
  
  QueeningSq = SQ_A1 + File(PawnSq) - 1 + 7 * SQ_UP
  
  '-- If the stronger side's king is in front of the pawn, it's a win
  If WKSq < PawnSq And File(WKSq) = File(PawnSq) Then
      Result = ScoreRook.EG - MaxDistance(WKSq, PawnSq)

  '-- If the weaker side's king is too far from the pawn and the rook, it's a win.
  ElseIf MaxDistance(BKSq, PawnSq) >= (3 + Abs(SideToMove = WeakSide)) And MaxDistance(BKSq, RookSq) >= 3 Then
      Result = ScoreRook.EG - MaxDistance(WKSq, PawnSq)

  '-- If the pawn is far advanced and supported by the defending king, the position is drawish
  ElseIf Rank(BKSq) <= 3 And MaxDistance(BKSq, PawnSq) = 1 And Rank(WKSq) >= 4 _
            And MaxDistance(WKSq, PawnSq) > (2 + Abs(SideToMove = StrongSide)) Then
      Result = 80 - 8 * MaxDistance(WKSq, PawnSq)
  Else
      Result = 200 - 8 * (MaxDistance(WKSq, PawnSq + SQ_DOWN) - MaxDistance(BKSq, PawnSq + SQ_DOWN) - MaxDistance(PawnSq, QueeningSq))
  End If
  
  If StrongSide = SideToMove Then Eval_KRKP = Result Else Eval_KRKP = -Result
  
End Function

Public Function Eval_KQKP() As Long
' KQ vs KP. In general, this is a win for the stronger side, but there are a
' few important exceptions. A pawn on 7th rank and on the A,C,F or H files
' with a king positioned next to it can be a draw, so in that case, we only
' use the distance between the kings.
  Dim WinnerKSq As Integer, LoserKSq As Integer, PawnSq As Integer, StrongSide As enumColor, WeakSide As enumColor
  Dim Result As Long, SideToMove As enumColor
  
  If WMaterial > BMaterial Then
    StrongSide = COL_WHITE: WeakSide = COL_BLACK: WinnerKSq = WKingLoc: LoserKSq = BKingLoc
  Else
    StrongSide = COL_BLACK: WeakSide = COL_WHITE: WinnerKSq = BKingLoc: LoserKSq = WKingLoc
  End If
  PawnSq = PieceSq(WeakSide, PT_PAWN)
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
 
End Function
 
Public Function Eval_KQKR() As Long

  Dim WinnerKSq As Integer, LoserKSq As Integer, StrongSide As enumColor, WeakSide As enumColor
  Dim Result As Long, SideToMove As enumColor
  
  If WMaterial > BMaterial Then
    StrongSide = COL_WHITE: WeakSide = COL_BLACK: WinnerKSq = WKingLoc: LoserKSq = BKingLoc
  Else
    StrongSide = COL_BLACK: WeakSide = COL_WHITE: WinnerKSq = BKingLoc: LoserKSq = WKingLoc
  End If
  If bWhiteToMove Then SideToMove = COL_WHITE Else SideToMove = COL_BLACK
 
  Result = ScoreQueen.EG - ScoreRook.EG + PushToEdges(LoserKSq) + PushClose(MaxDistance(WinnerKSq, LoserKSq))
  
  If StrongSide = SideToMove Then Eval_KQKR = Result Else Eval_KQKR = -Result
 
End Function
 
