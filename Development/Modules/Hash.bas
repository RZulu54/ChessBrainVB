Attribute VB_Name = "HashBas"
'==================================================
'= HashBas:
'= Hash functions
'==================================================
Option Explicit

'The style of the hash table rows
Public Const TT_NO_BOUND    As Byte = 0
Public Const TT_UPPER_BOUND As Byte = 1
Public Const TT_LOWER_BOUND As Byte = 2
Public Const TT_EXACT       As Byte = 3

Private Const HASH_CLUSTER As Long = 4
Public Const TT_TB_BASE_DEPTH As Long = 222

Public Type THashKey
  ' 2x 32 bit
  HashKey1 As Long
  HashKey2 As Long
End Type

Public ZobristHash1    As Long
Public ZobristHash2    As Long
Public HashWhiteToMove As Long
Public HashWhiteToMove2 As Long
Public HashWCanCastle  As Long
Public HashWCanCastle2  As Long
Public HashBCanCastle  As Long
Public HashBCanCastle2  As Long


Public InHashCnt       As Long
Public HashUsage       As Long
Private bHashUsed      As Boolean
Public HashGeneration As Long
Public EmptyHash As THashKey

Private Type HashTableEntry
  Position1 As Long ' 2x32 bit position hash key
  Position2 As Long
  Depth As Integer ' negative values possible for QSearch
  Generation As Byte
  IsChecking As Boolean
  MoveFrom As Byte
  MoveTarget As Byte
  MovePromoted As Byte
  EvalType As Byte
  Eval As Long
  StaticEval As Long
End Type

Public HashSize                            As Long

Dim ZobristTable(SQ_A1 To SQ_H8, 0 To 16)  As Long ' key for each piece typ eand board position
Dim ZobristTable2(SQ_A1 To SQ_H8, 0 To 16) As Long

'The main array to hold the hash table
Private HashTable()                        As HashTableEntry

Public Sub InitHash()
  'Initialize the hash-table
  Static bIniReadDone As Boolean
  
  If Not bIniReadDone Then
    bHashTrace = CBool(ReadINISetting("HASHTRACE", "0") <> "0")
    HashSize = Val(ReadINISetting("HASHSIZE", "64")) ' in MB
    If bHashTrace Then WriteTrace "Init hash size start " & HashSize & "MB " & Now()
    HashSize = HashSize * 40000   ' seems to fit...? hash len = 22
    bIniReadDone = True
  End If
  
  If bHashTrace Then WriteTrace "Init hash size start " & HashSize & " entries " & Now()
  HashUsage = 0
  ReDim HashTable(HashSize + HASH_CLUSTER)
  bHashUsed = False
  If bHashTrace Then WriteTrace "Init hash size done " & HashSize & " entries " & Now()
End Sub

Public Function HashBoard() As THashKey
  Dim i As Long, sq As Long
  
  ZobristHash1 = 0: ZobristHash2 = 0
  For i = 1 To NumPieces
    sq = Pieces(i): If sq <> 0 Then ZobristHash1 = ZobristHash1 Xor ZobristTable(sq, Board(sq)): ZobristHash2 = ZobristHash2 Xor ZobristTable2(sq, Board(sq))
  Next
  If EpPosArr(Ply) > 0 Then HashSetPiece EpPosArr(Ply), Board(EpPosArr(Ply))
  If bWhiteToMove Then
    ZobristHash1 = ZobristHash1 Xor HashWhiteToMove: ZobristHash2 = ZobristHash2 Xor HashWhiteToMove2
  End If
  If WhiteCastled = NO_CASTLE Then
    If Moved(WKING_START) = 0 Then
      If Moved(SQ_H1) = 0 Then ZobristHash1 = ZobristHash1 Xor HashWCanCastle
      If Moved(SQ_A1) = 0 Then ZobristHash2 = ZobristHash2 Xor HashWCanCastle2
    End If
  End If
  If BlackCastled = NO_CASTLE Then
    If Moved(BKING_START) = 0 Then
      If Moved(SQ_H8) = 0 Then ZobristHash1 = ZobristHash1 Xor HashBCanCastle
      If Moved(SQ_A8) = 0 Then ZobristHash2 = ZobristHash2 Xor HashBCanCastle2
    End If
  End If
 
  HashBoard.HashKey1 = ZobristHash1: HashBoard.HashKey2 = ZobristHash2
End Function

Public Function HashGetKey() As THashKey
  HashGetKey.HashKey1 = ZobristHash1
  HashGetKey.HashKey2 = ZobristHash2
End Function

Public Sub NextHashGeneration()
  HashGeneration = GetMin(255, GameMovesCnt \ 2 + 1)
End Sub

Public Sub HashSetKey(ByRef HashKey As THashKey)
  ZobristHash1 = HashKey.HashKey1
  ZobristHash2 = HashKey.HashKey2
End Sub

Public Function InsertIntoHashTable(HashKey As THashKey, _
                                    ByVal Depth As Long, _
                                    HashMove As TMove, _
                                    ByVal EvalType As Long, _
                                    ByVal Eval As Long, _
                                    ByVal StaticEval As Long)
                                    
  Dim IndexKey As Long, TmpMove As TMove, i As Long, ReplaceIndex As Long, MaxReplaceValue As Long, ReplaceValue As Long, bPosFound As Boolean
    
  Debug.Assert HashMove.From = 0 Or HashMove.Piece <> NO_PIECE
  If bTimeExit Then Exit Function ' score not exact
  
  TmpMove = HashMove ' Don't overwrite
  bHashUsed = True: bPosFound = False
  MaxReplaceValue = 9999
  
  '--- Compute hash key
  ZobristHash1 = HashKey.HashKey1: ZobristHash2 = HashKey.HashKey2
  IndexKey = HashKeyCompute() * HASH_CLUSTER
  ReplaceIndex = IndexKey
  For i = 0 To HASH_CLUSTER - 1
    With HashTable(IndexKey + i)
      If .Position1 <> 0 Then
        ' Don't overwrite more valuable entry
        If (.Position1 = ZobristHash1 And .Position2 = ZobristHash2) Then
          ' Position found: Preserve hash move if no new move
          If TmpMove.From = 0 And .MoveFrom > 0 Then
            TmpMove.From = .MoveFrom: TmpMove.Target = .MoveTarget: TmpMove.Promoted = .MovePromoted: TmpMove.IsChecking = .IsChecking
          End If
          ReplaceIndex = IndexKey + i: bPosFound = True
          Exit For
        Else
          ' Other position found. Overwrite?
          ReplaceValue = .Depth - 8 * (HashGeneration - .Generation)
          If ReplaceValue < MaxReplaceValue Then
            MaxReplaceValue = ReplaceValue: ReplaceIndex = IndexKey + i
            'If HashUsage > 0 Then HashUsage = HashUsage - 1
          End If
        End If
      Else
        If MaxReplaceValue > -9000 Then MaxReplaceValue = -9000: ReplaceIndex = IndexKey + i
      End If
    End With
  Next
  
  If HashTable(ReplaceIndex).Position1 = 0 And HashUsage < 2147483646 Then HashUsage = HashUsage + 1
  
  With HashTable(ReplaceIndex)
    '--- Save hash data, preserve hash move if no new move
    If Not bPosFound Or EvalType = TT_EXACT Or Depth > .Depth - 4 Or .Generation <> HashGeneration Then
      .Position1 = ZobristHash1: .Position2 = ZobristHash2
      .MoveFrom = TmpMove.From: .MoveTarget = TmpMove.Target: .MovePromoted = TmpMove.Promoted
      .EvalType = EvalType: .Eval = ScoreToHash(Eval)
      .StaticEval = StaticEval: .Depth = Depth
      .Generation = HashGeneration
      .IsChecking = TmpMove.IsChecking
      Debug.Assert .MoveFrom = 0 Or Board(.MoveFrom) <> NO_PIECE
    End If
  End With
End Function

Public Function IsInHashTable(HashKey As THashKey, _
                              ByRef HashDepth As Long, _
                              HashMove As TMove, _
                              ByRef EvalType As Long, _
                              ByRef Eval As Long, _
                              ByRef StaticEval As Long) As Boolean
  Dim IndexKey As Long, i As Long
  IsInHashTable = False: HashMove = EmptyMove: EvalType = TT_NO_BOUND: Eval = UNKNOWN_SCORE: StaticEval = UNKNOWN_SCORE: HashDepth = -999
  ZobristHash1 = HashKey.HashKey1
  ZobristHash2 = HashKey.HashKey2
  IndexKey = HashKeyCompute() * HASH_CLUSTER
  For i = 0 To HASH_CLUSTER - 1
    If HashTable(IndexKey + i).Position1 <> 0 And ZobristHash1 <> 0 Then
      With HashTable(IndexKey + i)
        If ZobristHash1 = .Position1 And ZobristHash2 = .Position2 Then
          If .Depth > HashDepth Then
            ' entry found
            IsInHashTable = True
            If InHashCnt < 2000000 Then InHashCnt = InHashCnt + 1
            
            '--- Read hash data
            If .MoveFrom > 0 Then
              HashMove.From = .MoveFrom: HashMove.Target = .MoveTarget
              HashMove.Promoted = .MovePromoted: HashMove.IsChecking = .IsChecking
              HashMove.Captured = Board(.MoveTarget): HashMove.Piece = Board(.MoveFrom): HashMove.CapturedNumber = Squares(.MoveTarget)
              Debug.Assert HashMove.Piece <> NO_PIECE
              'If Not MovePossible(HashMove) Then Stop
              Select Case HashMove.Piece
              Case WPAWN
                If .MoveTarget - .MoveFrom = 20 Then
                  HashMove.EnPassant = 1
                ElseIf Board(.MoveTarget) = BEP_PIECE Then
                  HashMove.EnPassant = 3
                End If
              Case BPAWN
                If .MoveFrom - .MoveTarget = 20 Then
                  HashMove.EnPassant = 2
                ElseIf Board(.MoveTarget) = WEP_PIECE Then
                  HashMove.EnPassant = 3
                End If
              Case WKING
                If .MoveFrom = SQ_E1 Then
                  If .MoveTarget = SQ_G1 Then
                    HashMove.Castle = WHITEOO
                  ElseIf .MoveTarget = SQ_C1 Then
                    HashMove.Castle = WHITEOOO
                  End If
                End If
              Case BKING
                If .MoveFrom = SQ_E8 Then
                  If .MoveTarget = SQ_G8 Then
                    HashMove.Castle = BLACKOO
                  ElseIf .MoveTarget = SQ_C8 Then
                    HashMove.Castle = BLACKOOO
                  End If
                End If
              End Select
            End If
            
            EvalType = .EvalType: Eval = HashToScore(.Eval): StaticEval = .StaticEval
            HashDepth = .Depth
            .Generation = HashGeneration ' Update generation
            Exit For
          End If
        End If
      End With
    End If
  Next
End Function

Public Function LimitDouble(ByVal d As Double) As Long
  ' Prevent overflow by looping off anything beyond 31 bits
  Const MaxNumber As Double = 2 ^ 31
  LimitDouble = CLng(d - (Fix(d / MaxNumber) * MaxNumber))
End Function

Public Sub InitZobrist()
  Static bDone As Boolean
  Dim p As Long, s As Long
  
  If bDone Then Exit Sub
  bDone = True
  ZobristHash1 = 0: ZobristHash2 = 0

  Randomize 1001 ' init random generator with fix value
  For p = SQ_A1 To SQ_H8
    For s = 0 To 16
      ZobristTable(p, s) = CalcUniqueKey(): ZobristTable2(p, s) = CalcUniqueKey()
    Next
  Next
  HashWhiteToMove = CalcUniqueKey(): HashWhiteToMove2 = CalcUniqueKey()
  HashWCanCastle = CalcUniqueKey(): HashWCanCastle2 = CalcUniqueKey()
  HashBCanCastle = CalcUniqueKey(): HashBCanCastle2 = CalcUniqueKey()
End Sub

Private Function CalcUniqueKey() As Long
  Static KeyList((SQ_H8 - SQ_A1 + 1) * 17 * 2 + 8) As Long
  Static ListCnt As Long
  Dim l As Long, i As Long
  
NextTry:
  l = 65536 * (Int(Rnd * 65536) - 32768) Or Int(Rnd * 65536)
  For i = 1 To ListCnt
    If KeyList(i) = l Then GoTo NextTry
  Next
  ListCnt = ListCnt + 1: KeyList(ListCnt) = l
  CalcUniqueKey = l
End Function

Public Sub HashSetPiece(ByVal Position As Long, ByVal Piece As Long)
  If Piece = FRAME Or Piece = NO_PIECE Then Exit Sub
  ZobristHash1 = ZobristHash1 Xor ZobristTable(Position, Piece)
  ZobristHash2 = ZobristHash2 Xor ZobristTable2(Position, Piece)
End Sub

Public Sub HashDelPiece(ByVal Position As Long, ByVal Piece As Long)
  If Piece = FRAME Or Piece = NO_PIECE Then Exit Sub
  ZobristHash1 = ZobristHash1 Xor ZobristTable(Position, Piece)
  ZobristHash2 = ZobristHash2 Xor ZobristTable2(Position, Piece)
End Sub

Public Sub HashMovePiece(ByVal From As Long, Target As Long, ByVal Piece As Long)
  ZobristHash1 = ZobristHash1 Xor ZobristTable(From, Piece) Xor ZobristTable(Target, Piece)
  ZobristHash2 = ZobristHash2 Xor ZobristTable(From, Piece) Xor ZobristTable2(Target, Piece)
End Sub

Public Function HashKeyCompute() As Long
  HashKeyCompute = ZobristHash1 Xor ZobristHash2
  If HashKeyCompute = -2147483648# Then HashKeyCompute = HashKeyCompute + 1
  HashKeyCompute = Abs(HashKeyCompute) Mod (HashSize \ HASH_CLUSTER)
End Function

Public Sub SetHashToMove()
 If bWhiteToMove Then
  ZobristHash1 = ZobristHash1 Xor HashWhiteToMove: ZobristHash2 = ZobristHash2 Xor HashWhiteToMove2
 End If
End Sub

Public Sub HashSetCastle()
  If WhiteCastled = NO_CASTLE Then ZobristHash1 = ZobristHash1 Xor HashWCanCastle: ZobristHash2 = ZobristHash2 Xor HashWCanCastle2
  If BlackCastled = NO_CASTLE Then ZobristHash1 = ZobristHash1 Xor HashBCanCastle: ZobristHash2 = ZobristHash2 Xor HashBCanCastle2
End Sub

Public Function ScoreToHash(ByVal Score As Long) As Long
  If Score >= MATE_IN_MAX_PLY Then
    ScoreToHash = Score + Ply
  ElseIf Score <= -MATE_IN_MAX_PLY Then
    ScoreToHash = Score - Ply
  Else
    ScoreToHash = Score
  End If
End Function

Public Function HashToScore(ByVal Score As Long) As Long
  If Score = UNKNOWN_SCORE Then
    HashToScore = Score
  ElseIf Score >= MATE_IN_MAX_PLY Then
    HashToScore = Score - Ply
  ElseIf Score <= -MATE_IN_MAX_PLY Then
    HashToScore = Score + Ply
  Else
    HashToScore = Score
  End If
End Function

Public Function HashUsagePerc() As String

  If HashSize = 0 Then
    HashUsagePerc = ""
  Else
    HashUsagePerc = Format(HashUsage * 100& / HashSize, "0.0")
  End If

End Function

