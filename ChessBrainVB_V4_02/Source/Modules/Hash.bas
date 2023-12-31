Attribute VB_Name = "basHash"
'==================================================
'= basHash:
'= Hash functions for transposition table
'==================================================
Option Explicit

Public Const MAX_THREADS       As Long = 64
Public Const MAX_HASHSIZE_MB   As Long = 1400  ' limit by 32 bit around 1500Mb / also long datatype overflow for bytes if more / limit in VB6-Development 300MB
'The style of the hash table rows
Public Const TT_NO_BOUND       As Byte = 0
Public Const TT_UPPER_BOUND    As Byte = 1
Public Const TT_LOWER_BOUND    As Byte = 2
Public Const TT_EXACT          As Byte = 3 ' = TT_UPPER_BOUND or TT_LOWER_BOUND !
Public Const HASH_CLUSTER      As Long = 4
Public Const TT_TB_BASE_DEPTH  As Long = 222
Public Const MATERIAL_HASHSIZE As Long = 8192

Public Const HASH_SIZE_FACTOR  As Long = 34000  ' entries per MB hash

Public Type THashKey
  ' 2x 32 bit
  HashKey1 As Long
  Hashkey2 As Long
End Type

Public ZobristHash1     As Long ' for calculation of hash key
Public ZobristHash2     As Long

Public HashWhiteToMove  As Long ' hashkey to add for  white to move
Public HashWhiteToMove2 As Long

Public HashWCanCastle   As Long
Public HashWCanCastle2  As Long

Public HashBCanCastle   As Long
Public HashBCanCastle2  As Long

Public HashExcluded     As Long
Public InHashCnt        As Long
Public HashAccessCnt    As Long
Public HashUsage        As Long
Private bHashUsed       As Boolean
Public bHashVerify      As Boolean
Public HashGeneration   As Long
Public EmptyHash        As THashKey

Private Type HashTableEntry
  Position1 As Long ' 2x32 bit position hash key
  Position2 As Long
  Depth As Integer ' not Byte, negative values possible for QSearch
  Generation As Byte
  IsChecking As Boolean
  MoveFrom As Byte
  MoveTarget As Byte
  MovePromoted As Byte
  EvalType As Byte
  Eval As Long
  StaticEval As Long
  PvHit As Boolean
  ThreadNum As Byte ' used for thread hit cnt => for testing only
End Type

Private moHashMap                              As clsHashMap
Public HashSizeMB                              As Long
Public HashSizeMax                              As Long
Public HashSize                                As Long ' in bytes
Public bHashSizeIgnoreGUI                      As Boolean ' HASHSIZE_IGNORE_GUI
Dim ZobristTable(SQ_A1 To SQ_H8, 0 To 16)      As Long ' key for each piece type and board position
Dim ZobristTable2(SQ_A1 To SQ_H8, 0 To 16)     As Long
'Dim FiftyZobristTable(0 To 100)                As Long ' fifty move draw: make different hash when fifty increases> not better
'Dim FiftyZobristTable2(0 To 100)               As Long
Dim MatZobristTable(0 To 10, 0 To 12)          As Long
'The main array to hold the hash table
Private HashTable()                            As HashTableEntry
Private HashCluster(0 To HASH_CLUSTER - 1)     As HashTableEntry
' Pointer to multi-Thread map data
Public NoOfThreads                             As Long
Public ThreadNum                               As Long  ' 0 = Main Thread
Public MainThreadStatus                        As Long, LastThreadStatus  As Long ' 1 = start, 0 = stop, -1 = Exit
Public ThreadCommand                           As String

Public HashMapEnd                              As Long
Public HashMapHashSizePtr                      As Long
Public HashMapThreadStatusPtr(MAX_THREADS - 1) As Long
Public HashMapBestPVPtr(MAX_THREADS - 1)       As Long ' Best pv for 10 moves
Public HashMapBoardPtr                         As Long
Public HashMapMovedPtr                         As Long
Public HashMapWhiteToMovePtr                   As Long
Public HashMapGameMovesCntPtr                  As Long
Public HashMapGameMovesPtr                     As Long
Public HashMapGamePosHashPtr                   As Long
Public HashMapSearchPtr                        As Long

Public HashRecLen                              As Long
Public HashClusterLen                          As Long
Private BestPV(10)                             As TMOVE
Public SingleThreadStatus(MAX_THREADS - 1)     As Long ' 1 = start, 0 = stop, -1 = Stopped
Private HashMapFile As String
Public bTraceHashCollision                     As Boolean

Public HashFoundFromOtherThread As Long
Private Type TMaterialHashEntry
  Hashkey As Long
  Score As Long
End Type

Public MaterialHash(MATERIAL_HASHSIZE) As TMaterialHashEntry

Public Sub InitHash()
  'Initialize the hash-table
  ' Use maximum hash size form INI file and memory command
  Dim NewHashSize As Long
  bHashTrace = CBool(ReadINISetting("HASHTRACE", "0") <> "0")
  HashSizeMB = GetMin(MAX_HASHSIZE_MB, Val(ReadINISetting("HASHSIZE", "64"))) ' 2 GB for 32 bit ( max 1.5 GB?)
  If CBool(ReadINISetting("HASHSIZE_IGNORE_GUI", "0") = "0") Then
    HashSizeMB = GetMax(HashSizeMB, MemoryMB) ' memory command value from GUI
  End If
  HashSizeMB = GetMin(MAX_HASHSIZE_MB, HashSizeMB) ' in 1 core: vb array MB, in IDE max around 350MB, EXE 1.5 GB
  If InIDE Then HashSizeMB = GetMin(128, HashSizeMB) ' Limited in IDE, depends on local memory usage
  
'HashSizeMB = 1400
'NoOfThreads = 2
'ThreadNum = 0 ' TEST

lblHashSize:
  If bHashTrace Then WriteTrace "Init hash size start " & HashSizeMB & "MB " & Now()
  If ThreadNum <= 0 Then  ' for helper threads if hash size was changed
   If Not pbMSExcelRunning Then
     WriteINISetting "HASH_USED", CStr(HashSizeMB)
   End If
  Else
     HashSizeMB = Val(ReadINISetting("HASH_USED", "64")) ' read from main thread
  End If
  HashSize = HashSizeMB * HASH_SIZE_FACTOR   ' in Bytes, seems to fit...? hash len = 31
  HashUsage = 0
  bHashUsed = False
  #If VBA_MODE = 0 Then ' Find unique file name if more than one version is CB are running
    HashMapFile = ReadINISetting("HASH_MAP_FILE", "CBVBHash" & Trim(App.Major) & Trim(App.Minor) & Trim(App.Revision) & "_" & GetAppTimeString() & ".DAT")  ' Change in INI to run 2x CB engine
  #End If
  
  bHashVerify = CBool(ReadINISetting("HASH_VERIFY", "0") <> "0") ' verify hash read/write to avoid collisions for many cores
  If NoOfThreads < 2 Then bHashVerify = False
  bTraceHashCollision = bHashVerify And CBool(ReadINISetting("HASH_COLL_TRACE", "0") <> "0") ' trace hash read/write collisions for > 1 core
  HashRecLen = LenB(HashCluster(0)): HashClusterLen = HashRecLen * HASH_CLUSTER
  
  If bHashTrace Then WriteTrace "InitHash: HashSize:" & HashSize & ", Threads:" & NoOfThreads
  If NoOfThreads <= 1 Then
    If bHashTrace Then WriteTrace "InitHash: Redim HashTable(0) done " & Now()
    If HashSize > HashSizeMax Then
      ReDim HashTable(HashSize + HASH_CLUSTER) ' may be OutOfMemory Error here
      If bHashTrace Then WriteTrace "InitHash: Redim done HashTable Size= " & HashSize & " entries " & Now()
      HashSizeMax = HashSize
    Else
      If bHashTrace Then WriteTrace "InitHash: Keep HashTable Size= " & HashSize & " entries " & Now()
      ' REDIM HashTable > creates random error:  Out of memory  / needs unfragmented memory fo rrequested size
      Dim j As Long
      For j = 1 To HashSize: HashTable(j).Position1 = 0: Next
   End If
    'MsgBox "Hashtable " & NoOfThreads & "/ " & ThreadNum
  ElseIf NoOfThreads > 1 Then
    ' Structure for game data
    ' ThreadStatus as long ' 1 = start, 0 = stop, -1 = Exit
    ReDim HashTable(0) ' internal hash not needed
    If bHashTrace Then WriteTrace "InitHash: Init hash map " & HashSize & " Bytes " & Now()
    
    ' HashMapEnd value starts a 0, every part of memory added will increase the value to address the next one
    HashMapEnd = 0
    If bHashTrace Then WriteTrace "HashMap: " & NoOfThreads & "/ " & ThreadNum & ", HashMapEnd:" & HashMapEnd & " MB:" & HashSizeMB
    Dim i As Long

    For i = 0 To MAX_THREADS - 1
      HashMapThreadStatusPtr(i) = HashMapEnd: HashMapEnd = HashMapEnd + LenB(MainThreadStatus)
      'If bHashTrace Then WriteTrace "InitHash:HashMapThreadStatusPtr:" & i & ":" & HashMapThreadStatusPtr(i)
    Next

    For i = 0 To MAX_THREADS - 1
      HashMapBestPVPtr(i) = HashMapEnd: HashMapEnd = HashMapEnd + LenB(PV(0, 0)) * 10
    Next

    HashMapBoardPtr = HashMapEnd: HashMapEnd = HashMapEnd + LenB(Board(0)) * MAX_BOARD
    HashMapMovedPtr = HashMapEnd: HashMapEnd = HashMapEnd + LenB(Moved(0)) * MAX_BOARD
    HashMapWhiteToMovePtr = HashMapEnd: HashMapEnd = HashMapEnd + LenB(bWhiteToMove)
    HashMapGameMovesCntPtr = HashMapEnd: HashMapEnd = HashMapEnd + LenB(GameMovesCnt)
    HashMapGameMovesPtr = HashMapEnd: HashMapEnd = HashMapEnd + LenB(arGameMoves(0)) * MAX_GAME_MOVES
    HashMapGamePosHashPtr = HashMapEnd: HashMapEnd = HashMapEnd + LenB(GamePosHash(0)) * MAX_GAME_MOVES
    
    ' the real hash for search is allocated now:
    HashMapSearchPtr = HashMapEnd
    HashMapEnd = HashMapEnd + HashRecLen * (HashSize + HASH_CLUSTER)
    ' allocate hash map file for multiple threads
    If ThreadNum >= 0 Then
      If bHashTrace Then WriteTrace "InitHash:OpenHashMap: HashMapEnd " & HashMapEnd
      NewHashSize = HashMapEnd
      OpenHashMap NewHashSize
      If NewHashSize <> HashMapEnd Then
        HashSizeMB = NewHashSize \ 1024# \ 1024# ' use reduced hash size
        WriteTrace "InitHash: New HashSize: " & HashSizeMB & " / " & Now()
        GoTo lblHashSize
      End If
    End If
  End If
  If bHashTrace Then WriteTrace "Init hash size done " & HashSize & " entries " & Now()
End Sub

Public Sub HashBoard(HashKeyOut As THashKey, ExcludedMove As TMOVE)
  Dim i As Long, sq As Long
  ZobristHash1 = 0: ZobristHash2 = 0

  For i = 1 To NumPieces
    sq = Pieces(i): If sq <> 0 Then ZobristHash1 = ZobristHash1 Xor ZobristTable(sq, Board(sq)): ZobristHash2 = ZobristHash2 Xor ZobristTable2(sq, Board(sq))
  Next

  If EpPosArr(Ply) > 0 Then HashSetPiece EpPosArr(Ply), Board(EpPosArr(Ply))
  If bWhiteToMove Then
    ZobristHash1 = ZobristHash1 Xor HashWhiteToMove: ZobristHash2 = ZobristHash2 Xor HashWhiteToMove2
  End If
  If Moved(WKING_START) = 0 Then ' white can castle?
    If Moved(SQ_H1) = 0 Then ZobristHash1 = ZobristHash1 Xor HashWCanCastle
    If Moved(SQ_A1) = 0 Then ZobristHash2 = ZobristHash2 Xor HashWCanCastle2
  End If
  If Moved(BKING_START) = 0 Then ' black can castle?
    If Moved(SQ_H8) = 0 Then ZobristHash1 = ZobristHash1 Xor HashBCanCastle
    If Moved(SQ_A8) = 0 Then ZobristHash2 = ZobristHash2 Xor HashBCanCastle2
  End If
  If ExcludedMove.From > 0 Then ' use from/target sq to be different to normal position
    ZobristHash1 = ZobristHash1 Xor ZobristTable(ExcludedMove.From, ExcludedMove.Piece): ZobristHash2 = ZobristHash2 Xor ZobristTable2(ExcludedMove.Target, ExcludedMove.Piece)
  End If
  
  HashKeyOut.HashKey1 = ZobristHash1: HashKeyOut.Hashkey2 = ZobristHash2
End Sub

Public Function HashGetKey() As THashKey
  HashGetKey.HashKey1 = ZobristHash1
  HashGetKey.Hashkey2 = ZobristHash2
End Function

Public Sub NextHashGeneration()
  HashGeneration = GetMin(255, GameMovesCnt \ 2 + 1)
End Sub

Public Sub HashSetKey(ByRef Hashkey As THashKey)
  ZobristHash1 = Hashkey.HashKey1
  ZobristHash2 = Hashkey.Hashkey2
End Sub


Public Function HashTableSave(Hashkey As THashKey, _
                              Depth As Long, _
                              HashMove As TMOVE, _
                              EvalType As Long, _
                              EvalScore As Long, _
                              StaticEval As Long, _
                              PvHit As Boolean)
  'Dim FiftyHash As THashKey
  'If Fifty >= 4 Then ' fifty move draw: make different hash when fifty increases every 8 moves > prolbem with 3x draw detection using hash
  '  FiftyHash.HashKey1 = Hashkey.HashKey1 Xor FiftyZobristTable(Fifty \ 8): FiftyHash.Hashkey2 = Hashkey.Hashkey2 Xor FiftyZobristTable2(Fifty \ 8)
  'Else
  '  FiftyHash.HashKey1 = Hashkey.HashKey1: FiftyHash.Hashkey2 = Hashkey.Hashkey2
  'End If

  If ThreadNum < 0 Then ' single core using internal VB array
    InsertIntoHashTable Hashkey, Depth, HashMove, EvalType, EvalScore, StaticEval, PvHit
  Else ' multiple cores using global hash map
    InsertIntoHashMap Hashkey, Depth, HashMove, EvalType, EvalScore, StaticEval, PvHit
  End If
End Function
                              


Public Function InsertIntoHashTable(Hashkey As THashKey, _
                                    ByVal Depth As Long, _
                                    HashMove As TMOVE, _
                                    ByVal EvalType As Long, _
                                    ByVal EvalScore As Long, _
                                    ByVal StaticEval As Long, _
                                    ByVal PvHit As Boolean)
  '--- Insert hash entry into hash array if only one thread (faster than access to global mapped memory)
  Dim ClusterIndex As Long, NewHashMove As TMOVE, i As Long, ReplaceIndex As Long, MaxReplaceValue As Long, ReplaceValue As Long, bPosFound As Boolean
  Debug.Assert HashMove.From = 0 Or (HashMove.Piece <> NO_PIECE And Board(HashMove.From) <> NO_PIECE)
  If bTimeExit Then Exit Function ' score not exact
  SetMove NewHashMove, HashMove  ' Don't overwrite move of caller function
  bHashUsed = True: bPosFound = False
  MaxReplaceValue = 9999
  '--- Compute hash key
  ZobristHash1 = Hashkey.HashKey1: ZobristHash2 = Hashkey.Hashkey2
  ClusterIndex = HashKeyCompute() * HASH_CLUSTER
  ReplaceIndex = ClusterIndex
  If HashAccessCnt < 2100000000 Then HashAccessCnt = HashAccessCnt + 1

  For i = 0 To HASH_CLUSTER - 1
    With HashTable(ClusterIndex + i)
      If .Position1 = 0 Then ReplaceIndex = ClusterIndex + i: Exit For ' use empty entry
      If HashGeneration = .Generation Then If HashUsage < 2100000000 Then HashUsage = HashUsage + 1
      ' Don't overwrite more valuable entry
      If (.Position1 = ZobristHash1 And .Position2 = ZobristHash2) Then
        ' Position found: Preserve hash move if no new move
        If .MoveFrom > 0 And NewHashMove.From = 0 Then ' old hash move exists
          NewHashMove.From = .MoveFrom: NewHashMove.Target = .MoveTarget: NewHashMove.Promoted = .MovePromoted: NewHashMove.IsChecking = .IsChecking
        End If
        ReplaceIndex = ClusterIndex + i: bPosFound = True
        Exit For
      Else
        ' Other position found. Overwrite?
        ReplaceValue = .Depth - 8 * (HashGeneration - .Generation)
        If ReplaceValue < MaxReplaceValue Then
          MaxReplaceValue = ReplaceValue: ReplaceIndex = ClusterIndex + i
        End If
      End If
    End With
  Next
 
 
  With HashTable(ReplaceIndex)
    '--- Save hash data, preserve hash move if no new move
    If Not bPosFound Or EvalType = TT_EXACT Or Depth > .Depth - 4 Then
      .Position1 = ZobristHash1: .Position2 = ZobristHash2
      .MoveFrom = NewHashMove.From: .MoveTarget = NewHashMove.Target: .MovePromoted = NewHashMove.Promoted
      .EvalType = EvalType: .Eval = ScoreToHash(EvalScore)
      .StaticEval = StaticEval: .Depth = Depth
      .Generation = HashGeneration
      .IsChecking = NewHashMove.IsChecking
      .PvHit = PvHit
      Debug.Assert .MoveFrom = 0 Or Board(.MoveFrom) <> NO_PIECE
    End If
  End With

End Function

Public Function HashTableRead(Hashkey As THashKey, _
                              ByRef HashDepth As Long, _
                              HashMove As TMOVE, _
                              ByRef EvalType As Long, _
                              ByRef EvalScore As Long, _
                              ByRef StaticEval As Long, _
                              ByRef PvHit As Boolean, ByRef HashThreadNum As Long) As Boolean
 ' Dim FiftyHash As THashKey
 ' If Fifty >= 4 Then ' fifty move draw: make different hash when fifty increases every 8 moves > prolbem with 3x draw detection using hash
 '   FiftyHash.HashKey1 = Hashkey.HashKey1 Xor FiftyZobristTable(Fifty \ 8): FiftyHash.Hashkey2 = Hashkey.Hashkey2 Xor FiftyZobristTable2(Fifty \ 8)
 ' Else
 '   FiftyHash.HashKey1 = Hashkey.HashKey1: FiftyHash.Hashkey2 = Hashkey.Hashkey2
 ' End If
                              
  If ThreadNum < 0 Then ' single core using internal VB array
    HashTableRead = IsInHashTable(Hashkey, HashDepth, HashMove, EvalType, EvalScore, StaticEval, PvHit)
    HashThreadNum = -1
  Else ' multiple cores using global hash map
    HashTableRead = IsInHashMap(Hashkey, HashDepth, HashMove, EvalType, EvalScore, StaticEval, PvHit, HashThreadNum)
  End If
End Function

Public Function IsInHashTable(Hashkey As THashKey, _
                              ByRef HashDepth As Long, _
                              HashMove As TMOVE, _
                              ByRef EvalType As Long, _
                              ByRef EvalScore As Long, _
                              ByRef StaticEval As Long, _
                              ByRef PvHit As Boolean) As Boolean
  '--- Search for hash entry into hash array if one thread
  Dim IndexKey As Long, i As Long
  
  IsInHashTable = False: ClearMove HashMove: EvalType = TT_NO_BOUND: EvalScore = VALUE_NONE: StaticEval = VALUE_NONE: HashDepth = -MAX_GAME_MOVES
  ZobristHash1 = Hashkey.HashKey1: ZobristHash2 = Hashkey.Hashkey2
  IndexKey = HashKeyCompute() * HASH_CLUSTER

  For i = 0 To HASH_CLUSTER - 1
    If HashTable(IndexKey + i).Position1 = 0 Then If ZobristHash1 <> 0 Then Exit Function '--- empty entry, not found
      With HashTable(IndexKey + i)
        If ZobristHash1 = .Position1 And ZobristHash2 = .Position2 Then
          If .Depth > HashDepth Then
            ' entry found
            IsInHashTable = True: PvHit = False
            If InHashCnt < 2000000 Then InHashCnt = InHashCnt + 1
            '--- Read hash data
            If .MoveFrom > 0 Then
              HashMove.From = .MoveFrom: HashMove.Target = .MoveTarget: HashMove.IsChecking = .IsChecking
              If Board(.MoveTarget) <= NO_PIECE Then HashMove.Captured = Board(.MoveTarget)
              HashMove.Piece = Board(.MoveFrom): HashMove.CapturedNumber = Squares(.MoveTarget)
              HashMove.Promoted = .MovePromoted: If HashMove.Promoted <> 0 Then HashMove.Piece = HashMove.Promoted
              Debug.Assert HashMove.Piece <> NO_PIECE
              HashMove.IsLegal = True

              'If Not MovePossible(HashMove) Then Stop
              Select Case HashMove.Piece
                Case WPAWN
                  If .MoveTarget - .MoveFrom = 20 Then
                    HashMove.EnPassant = ENPASSANT_WMOVE
                  ElseIf Board(.MoveTarget) = BEP_PIECE Then
                    HashMove.EnPassant = ENPASSANT_CAPTURE
                    HashMove.Captured = BEP_PIECE
                  End If
                Case BPAWN
                  If .MoveFrom - .MoveTarget = 20 Then
                    HashMove.EnPassant = ENPASSANT_BMOVE
                  ElseIf Board(.MoveTarget) = WEP_PIECE Then
                    HashMove.EnPassant = ENPASSANT_CAPTURE
                    HashMove.Captured = WEP_PIECE
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
            EvalType = .EvalType: EvalScore = HashToScore(.Eval): StaticEval = .StaticEval
            HashDepth = .Depth
            PvHit = .PvHit
            .Generation = HashGeneration ' Update generation> still valid in this game
            Exit For
          End If
        End If
      End With
  Next

End Function

Public Function LimitDouble(ByVal d As Double) As Long
  ' Prevent overflow by looping off anything beyond 31 bits
  Const MaxNumber As Double = 2 ^ 31
  LimitDouble = CLng(d - (Fix(d / MaxNumber) * MaxNumber))
End Function

Public Sub InitZobrist()
  ' init values for hash calculation. 2x32 bit for 64 bit key
  Static bDone As Boolean
  Dim p        As Long, s As Long
  If bDone Then Exit Sub
  bDone = True
  ZobristHash1 = 0: ZobristHash2 = 0
  Randomize 1001 ' init random generator with fix value

  ' create hash value for each piece type and each board position
  For s = SQ_A1 To SQ_H8
    For p = 0 To 16
      ZobristTable(s, p) = CalcUniqueKey(): ZobristTable2(s, p) = CalcUniqueKey()
    Next
  Next

  HashWhiteToMove = CalcUniqueKey(): HashWhiteToMove2 = CalcUniqueKey()
  HashWCanCastle = CalcUniqueKey(): HashWCanCastle2 = CalcUniqueKey()
  HashBCanCastle = CalcUniqueKey(): HashBCanCastle2 = CalcUniqueKey()

'  ' for rule: draw after fifty quiet moves , make a different hash key when fifty counter increases
'  For s = 1 To 100
'    FiftyZobristTable(s) = CalcUniqueKey(): FiftyZobristTable2(s) = CalcUniqueKey()
'  Next

  ' keys for material values total
  For s = 0 To 10 ' Material hash: Piece cnt
    For p = 0 To 12 ' Piece
      MatZobristTable(s, p) = CalcUniqueKey()
    Next
  Next

End Sub

Public Function CalcMaterialKey() As Long
  CalcMaterialKey = MatZobristTable(PieceCnt(WQUEEN), WQUEEN) Xor MatZobristTable(PieceCnt(BQUEEN), BQUEEN) Xor MatZobristTable(PieceCnt(WROOK), WROOK) Xor MatZobristTable(PieceCnt(BROOK), BROOK) Xor MatZobristTable(PieceCnt(WBISHOP), WBISHOP) Xor MatZobristTable(PieceCnt(BBISHOP), BBISHOP) Xor MatZobristTable(PieceCnt(WKNIGHT), WKNIGHT) Xor MatZobristTable(PieceCnt(BKNIGHT), BKNIGHT) Xor MatZobristTable(PieceCnt(WPAWN), WPAWN) Xor MatZobristTable(PieceCnt(BPAWN), BPAWN)
End Function

Private Function CalcUniqueKey() As Long
  Static KeyList((SQ_H8 + 1) * 17 * 2 + 8) As Long
  Static ListCnt                           As Long
  Dim l                                    As Long, i As Long
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

Public Function HashKeyComputeMap() As Long
  HashKeyComputeMap = ZobristHash1 Xor ZobristHash2
  If HashKeyComputeMap = -2147483648# Then HashKeyComputeMap = HashKeyComputeMap + 1
  HashKeyComputeMap = Abs(HashKeyComputeMap) Mod (HashSize \ HASH_CLUSTER)
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
  If Score = VALUE_NONE Then
    ScoreToHash = Score
  ElseIf Score >= MATE_IN_MAX_PLY Then
    ScoreToHash = Score + Ply
  ElseIf Score <= -MATE_IN_MAX_PLY Then
    ScoreToHash = Score - Ply
  Else
    ScoreToHash = Score
  End If
End Function

Public Function HashToScore(ByVal Score As Long) As Long
  If Score = VALUE_NONE Then
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
    If HashUsage > HashSize Then HashUsage = HashSize
    HashUsagePerc = Format$(CDbl(HashUsage) * 100& / HashSize, "0.0")
  End If
End Function

Public Function HashUsageUCI() As Long
  Dim x As Single
  If HashSize = 0 Or HashUsage <= 0 Then
    HashUsageUCI = 0
  Else
    x = HashUsage: x = x * CSng(1000) / CSng(1 + HashAccessCnt)
    HashUsageUCI = GetMin(1000, CLng(x))
  End If
End Function

Public Function OpenHashMap(ByRef TotalSize As Long) As Long
  '--- init global mapped memory if more then one thread, used by all threads!
  Static OldHashSize As Long
  If OldHashSize = 0 Then
    Set moHashMap = New clsHashMap
  End If
  If OldHashSize = 0 Or OldHashSize <> TotalSize Then
    If ThreadNum = 0 Then
      If OldHashSize = 0 Then
        Set moHashMap = New clsHashMap
        If bThreadTrace Then WriteTrace "OpenHashMap: New clsHashMap: " & TotalSize
      Else
        If bThreadTrace Then WriteTrace "OpenHashMap: CloseMap"
        moHashMap.CloseMap
      End If
      moHashMap.CreateMap HashMapFile, TotalSize ' TotalSize may be reduced if not enough memory !!!
      If bThreadTrace Then WriteTrace "OpenHashMap: CreateMap: Size " & TotalSize
    ElseIf ThreadNum > 0 Then
      moHashMap.OpenMap HashMapFile, TotalSize
      If bThreadTrace Then WriteTrace "OpenHashMap: OpenMap: Size " & TotalSize
    End If
    OldHashSize = TotalSize
  Else
    If ThreadNum = 0 Then moHashMap.ClearMap TotalSize
  End If
End Function

Public Function CloseHashMap() As Long
  moHashMap.CloseMap
End Function
 
Public Function InsertIntoHashMap(Hashkey As THashKey, _
                                  ByVal Depth As Long, _
                                  HashMove As TMOVE, _
                                  ByVal EvalType As Long, _
                                  ByVal Eval As Long, _
                                  ByVal StaticEval As Long, _
                                  ByVal PvHit As Boolean)
  '--- Insert hash entry into global mapped memory if more then one thread, used by all threads!
  Dim ClusterIndex As Long, NewHashMove As TMOVE, i As Long, ReplaceIndex As Long, MaxReplaceValue As Long, ReplaceValue As Long, bPosFound As Boolean
  Debug.Assert HashMove.From = 0 Or HashMove.Piece <> NO_PIECE
  Debug.Assert NoOfThreads > 1
  If bTimeExit Then Exit Function ' score not exact
  
  'If ThreadNum > 0 Then Exit Function '####################TESTc2
  
  SetMove NewHashMove, HashMove  ' Don't overwrite
  bHashUsed = True: bPosFound = False
  MaxReplaceValue = 9999
  '--- Compute hash key
  ZobristHash1 = Hashkey.HashKey1: ZobristHash2 = Hashkey.Hashkey2
  ClusterIndex = HashKeyComputeMap() * HASH_CLUSTER
  ReplaceIndex = 0
  moHashMap.ReadMapHashCluster ClusterIndex, CLng(VarPtr(HashCluster(0))), HashClusterLen ' read this cluster only
  If HashAccessCnt < 2100000000 Then HashAccessCnt = HashAccessCnt + 1
  
  For i = 0 To HASH_CLUSTER - 1
    With HashCluster(i) ' search in retrieved cluster
      If .Position1 = 0 Then ReplaceIndex = ClusterIndex + i: Exit For ' empty entry found
      If HashGeneration = .Generation Then If HashUsage < 2100000000 Then HashUsage = HashUsage + 1
      ' Don't overwrite more valuable entry
      If (.Position1 = ZobristHash1 And .Position2 = ZobristHash2) Then
        ' Position found: Preserve hash move if no new move
        If NewHashMove.From = 0 And .MoveFrom > 0 Then
          NewHashMove.From = .MoveFrom: NewHashMove.Target = .MoveTarget: NewHashMove.Promoted = .MovePromoted: NewHashMove.IsChecking = .IsChecking
        End If
        ReplaceIndex = ClusterIndex + i: bPosFound = True
        Exit For
      Else
        ' Other position found. Find least valuable entry
        ReplaceValue = .Depth - 8 * (HashGeneration - .Generation)
        If ReplaceValue < MaxReplaceValue Then
          MaxReplaceValue = ReplaceValue: ReplaceIndex = ClusterIndex + i
        End If
      End If
    End With
  Next

  With HashCluster(ReplaceIndex - ClusterIndex)
    '--- Save hash data, preserve hash move if no new move
     If Not bPosFound Or EvalType = TT_EXACT Or Depth > .Depth - 4 Then
      .Position1 = ZobristHash1: .Position2 = ZobristHash2
      .MoveFrom = NewHashMove.From: .MoveTarget = NewHashMove.Target: .MovePromoted = NewHashMove.Promoted
      .EvalType = EvalType: .Eval = ScoreToHash(Eval)
      .StaticEval = StaticEval: .Depth = Depth
      .Generation = HashGeneration
      .IsChecking = NewHashMove.IsChecking
      .PvHit = PvHit
      If ThreadNum >= 0 Then .ThreadNum = ThreadNum
      '--- Write Hash Map: replace index in Cluster only
      moHashMap.WriteMapHashEntry ReplaceIndex, CLng(VarPtr(HashCluster(ReplaceIndex - ClusterIndex)))
      Debug.Assert .MoveFrom = 0 Or Board(.MoveFrom) <> NO_PIECE
    End If
  End With

End Function

Public Function IsInHashMap(Hashkey As THashKey, _
                            ByRef HashDepth As Long, _
                            HashMove As TMOVE, _
                            ByRef EvalType As Long, _
                            ByRef Eval As Long, _
                            ByRef StaticEval As Long, _
                            ByRef PvHit As Boolean, ByRef HashThreadNum As Long) As Boolean
  '--- search for hash entry in global mapped memory if more then one thread
  Dim IndexKey As Long, i As Long
  Debug.Assert NoOfThreads > 1
  IsInHashMap = False: ClearMove HashMove: EvalType = TT_NO_BOUND: Eval = VALUE_NONE: StaticEval = VALUE_NONE: HashDepth = -MAX_GAME_MOVES
  ZobristHash1 = Hashkey.HashKey1: ZobristHash2 = Hashkey.Hashkey2
  IndexKey = HashKeyComputeMap() * HASH_CLUSTER
  moHashMap.ReadMapHashCluster IndexKey, CLng(VarPtr(HashCluster(0))), HashClusterLen

  For i = 0 To HASH_CLUSTER - 1

    With HashCluster(i)
      If .Position1 = 0 Then If ZobristHash1 <> 0 Then Exit Function '--- empty entry, not found
        If ZobristHash1 = .Position1 And ZobristHash2 = .Position2 Then
          If .Depth > HashDepth Then
            ' entry found
            IsInHashMap = True: PvHit = False
            If InHashCnt < 2000000 Then InHashCnt = InHashCnt + 1
            '--- Read hash data
            If .MoveFrom > 0 Then
              HashMove.From = .MoveFrom: HashMove.Target = .MoveTarget: HashMove.IsChecking = .IsChecking
              If Board(.MoveTarget) <= NO_PIECE Then HashMove.Captured = Board(.MoveTarget)
              HashMove.Piece = Board(.MoveFrom): HashMove.CapturedNumber = Squares(.MoveTarget)
              HashMove.Promoted = .MovePromoted: If HashMove.Promoted <> 0 Then HashMove.Piece = HashMove.Promoted
              Debug.Assert HashMove.Piece <> NO_PIECE
              HashMove.IsLegal = True

              'If Not MovePossible(HashMove) Then Stop
              Select Case HashMove.Piece
                Case WPAWN
                  If .MoveTarget - .MoveFrom = 20 Then
                    HashMove.EnPassant = ENPASSANT_WMOVE
                  ElseIf Board(.MoveTarget) = BEP_PIECE Then
                    HashMove.EnPassant = ENPASSANT_CAPTURE
                    HashMove.Captured = BEP_PIECE
                  End If
                Case BPAWN
                  If .MoveFrom - .MoveTarget = 20 Then
                    HashMove.EnPassant = ENPASSANT_BMOVE
                  ElseIf Board(.MoveTarget) = WEP_PIECE Then
                    HashMove.EnPassant = ENPASSANT_CAPTURE
                    HashMove.Captured = WEP_PIECE
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
            ' Read values for entry
            EvalType = .EvalType: Eval = HashToScore(.Eval): StaticEval = .StaticEval
            HashDepth = .Depth
            PvHit = .PvHit
            HashThreadNum = .ThreadNum
            If .Generation <> HashGeneration Then
              .Generation = HashGeneration ' Update generation, each game move is a new generation
              '--- Write Hash Map: replace index in Cluster only
              moHashMap.WriteMapHashEntry IndexKey + i, CLng(VarPtr(HashCluster(i)))
            End If
            'If ThreadNum >= 0 Then If .ThreadNum <> GetMax(0, ThreadNum) Then HashFoundFromOtherThread = HashFoundFromOtherThread + 1
            Exit For
          End If
        End If
    End With
  Next i

End Function

Public Function InitThreads()
  Static bInitDone As Boolean
  Dim i            As Long
  DoEvents
  #If VBA_MODE = 0 Then
    If Not bInitDone And NoOfThreads > 1 Then
      If CreateAppLockFile() Then ' Already started?
        If bThreadTrace Then WriteTrace "InitThreads: NoOfThreads=" & NoOfThreads
        MainThreadStatus = 0: WriteMainThreadStatus 0 ' idle
  '  Dim tStart As Single, tEnd As Single
  '  tStart = Timer
  '  Dim sCmd As String
        For i = 2 To NoOfThreads
          StartProcess App.Path & "\ChessBrainVB.exe thread" & Trim$(CStr(i - 1)) ' Much faster
        Next
          '---Shell App.Path & "\ChessBrainVB.exe thread" & Trim$(CStr(i - 1)), vbMinimizedNoFocus  ' SHELL is MUCH slower ( 1 sec per call?!?)
   
   '   tEnd = Timer()
   '   WriteTrace "Threads started:" & ", Time:" & Format$(tEnd - tStart, "0.00000")
 
        Sleep 500
      End If
    End If
  #End If
  bInitDone = True
End Function

Public Function CreateAppLockFile() As Boolean
  ' for main thread: create a locked file that gets unlocked when main thread end/crashed
  ' this file is checked by the helper threads: if file is unlocked also exit helper threads
  Static lLOCK_FILEHANDLE As Long
  Sleep 200 ' wait for end of previous exe run
  #If VBA_MODE = 0 Then
    Debug.Assert NoOfThreads > 1
    lLOCK_FILEHANDLE = FreeFile()
    On Error GoTo lblLockErr
    Open App.Path & "\CB_THREAD0.TXT" For Append Access Write Lock Write As #lLOCK_FILEHANDLE
    Print #lLOCK_FILEHANDLE, "Temporary lock file. Main thread started:" & Now()
    CreateAppLockFile = True
  #End If
lblExit:
  Exit Function
lblLockErr:
  CreateAppLockFile = False
  WriteTrace "Already started? Cannot open Application lock file: CB_THREAD0.TXT " & Now()
  Resume lblExit
End Function

Public Function CheckAppLockFile() As Boolean
  ' this file is checked is used by the helper threads: returns true if file is unlocked > also exit helper threads
  Dim lLOCK_FILEHANDLE2 As Long
  On Error GoTo lblErr
  CheckAppLockFile = False
  #If VBA_MODE = 0 Then
    lLOCK_FILEHANDLE2 = FreeFile()
    Open App.Path & "\CB_THREAD0.TXT" For Append Access Write Lock Write As #lLOCK_FILEHANDLE2
    CheckAppLockFile = False ' File unlocked-> main thread was terminated-> exit helper threads too
    Close #lLOCK_FILEHANDLE2
  #End If
  Exit Function
lblErr:
  CheckAppLockFile = True
End Function

Public Function WriteLog(isLine As String) As Boolean
  ' write debug log
  Dim lLOCK_FILEHANDLE3 As Long
  #If VBA_MODE = 0 Then
    lLOCK_FILEHANDLE3 = FreeFile()
    Open psEnginePath & "\DEBUG_LOG.TXT" For Append As #lLOCK_FILEHANDLE3
    Print #lLOCK_FILEHANDLE3, isLine
    Close #lLOCK_FILEHANDLE3
  #End If
End Function



Public Sub CheckThreadTermination(ByVal bCheckAlways As Boolean)
  Debug.Assert NoOfThreads > 1
  If ThreadNum >= 1 Then
    If bCheckAlways Or (Nodes > LastThreadCheckNodesCnt + (GUICheckIntervalNodes * 50)) Then
      LastThreadCheckNodesCnt = Nodes
      If Not CheckAppLockFile() Then
        '>>> END of program here because main thread was terminated
        CloseHashMap
        If bThreadTrace Then WriteTrace "!!! Main Thread terminated: Stop helper thread! " & Now()
        End '<<<<
      End If
    End If
  End If
End Sub

Public Function WriteMainThreadStatus(ByVal ilNewThreadStatus As Long) As Long
  Debug.Assert NoOfThreads > 1
  SingleThreadStatus(0) = ilNewThreadStatus
  moHashMap.WriteMapPos HashMapThreadStatusPtr(0), CLng(VarPtr(ilNewThreadStatus)), CLng(LenB(ilNewThreadStatus))
  If bThreadTrace Then WriteTrace "WriteMainThreadStatus: " & HashMapThreadStatusPtr(0)
End Function

Public Function ReadMainThreadStatus() As Long
  Static LastRead      As Long
  Dim MainThreadStatus As Long
  Debug.Assert NoOfThreads > 1
  moHashMap.ReadMapPos HashMapThreadStatusPtr(0), CLng(VarPtr(MainThreadStatus)), CLng(LenB(MainThreadStatus))
  SingleThreadStatus(0) = MainThreadStatus
  ReadMainThreadStatus = MainThreadStatus
  If bThreadTrace Then If LastRead <> ReadMainThreadStatus Then WriteTrace "ReadMainThreadStatus:Threadnum=" & ThreadNum & ", Ptr:" & HashMapThreadStatusPtr(0) & ", MainStatus:" & ReadMainThreadStatus & " / " & Now()
  LastRead = ReadMainThreadStatus
End Function

Public Function WriteHelperThreadStatus(ByVal ilThreadNum As Long, _
                                        ByVal ilNewThreadStatus As Long) As Long
  ' Write run status for current thread
  Debug.Assert NoOfThreads > 1 And ilThreadNum > 0
  SingleThreadStatus(ilThreadNum) = ilNewThreadStatus
  moHashMap.WriteMapPos HashMapThreadStatusPtr(ilThreadNum), CLng(VarPtr(ilNewThreadStatus)), CLng(LenB(ilNewThreadStatus))
End Function

Public Function ReadHelperThreadStatus(ByVal ilThreadNum As Long) As Long
  ' Write run status for current thread
  Dim HelperThreadStatus As Long
  Debug.Assert NoOfThreads > 1 And ilThreadNum > 0
  moHashMap.ReadMapPos HashMapThreadStatusPtr(ilThreadNum), CLng(VarPtr(HelperThreadStatus)), CLng(LenB(HelperThreadStatus))
  SingleThreadStatus(ilThreadNum) = HelperThreadStatus
  ReadHelperThreadStatus = HelperThreadStatus
End Function

Public Function WriteMapGameData() As Long
  ' Write game moves to map for other threads
  Debug.Assert NoOfThreads > 1
  moHashMap.WriteMapPos HashMapBoardPtr, CLng(VarPtr(Board(0))), CLng(LenB(Board(0)) * MAX_BOARD)
  moHashMap.WriteMapPos HashMapMovedPtr, CLng(VarPtr(Moved(0))), CLng(LenB(Moved(0)) * MAX_BOARD)
  moHashMap.WriteMapPos HashMapWhiteToMovePtr, CLng(VarPtr(bWhiteToMove)), CLng(LenB(bWhiteToMove))
  moHashMap.WriteMapPos HashMapGameMovesCntPtr, CLng(VarPtr(GameMovesCnt)), CLng(LenB(GameMovesCnt))
  arGameMoves(MAX_GAME_MOVES - 1).Target = Fifty ' tricky fix to avoid new map size
  moHashMap.WriteMapPos HashMapGameMovesPtr, CLng(VarPtr(arGameMoves(0))), CLng(LenB(arGameMoves(0)) * MAX_GAME_MOVES)
  moHashMap.WriteMapPos HashMapGamePosHashPtr, CLng(VarPtr(GamePosHash(0))), CLng(LenB(GamePosHash(0)) * MAX_GAME_MOVES)
End Function

Public Function ReadMapGameData() As Long
  ' Read game moves to map for other threads
  Dim bToMove As Boolean
  Debug.Assert NoOfThreads > 1
  moHashMap.ReadMapPos HashMapBoardPtr, CLng(VarPtr(Board(0))), CLng(LenB(Board(0)) * MAX_BOARD)
  InitEpArr
  moHashMap.ReadMapPos HashMapMovedPtr, CLng(VarPtr(Moved(0))), CLng(LenB(Moved(0)) * MAX_BOARD)
  moHashMap.ReadMapPos HashMapWhiteToMovePtr, CLng(VarPtr(bToMove)), CLng(LenB(bToMove))
  bWhiteToMove = bToMove: bCompIsWhite = bWhiteToMove
  moHashMap.ReadMapPos HashMapGameMovesCntPtr, CLng(VarPtr(GameMovesCnt)), CLng(LenB(GameMovesCnt))
  moHashMap.ReadMapPos HashMapGameMovesPtr, CLng(VarPtr(arGameMoves(0))), CLng(LenB(arGameMoves(0)) * MAX_GAME_MOVES)
  Fifty = arGameMoves(MAX_GAME_MOVES - 1).Target ' tricky fix to avoid new map size
  moHashMap.ReadMapPos HashMapGamePosHashPtr, CLng(VarPtr(GamePosHash(0))), CLng(LenB(GamePosHash(0)) * MAX_GAME_MOVES)
  InitPieceSquares
End Function

Public Function ClearMapBestPVforThread() As Long
  Dim th As Long
  Erase BestPV()

  For th = 0 To MAX_THREADS - 1
    moHashMap.WriteMapPos HashMapBestPVPtr(th), CLng(VarPtr(BestPV(0))), CLng(LenB(BestPV(0)) * 10)
  Next

End Function

Public Function WriteMapBestPVforThread(ByVal CompletedDepth As Long, _
                                        ByVal BestScore As Long, _
                                        BestMove As TMOVE) As Long
  ' Write PV from helper thread for main thread
  Dim i As Long
  Debug.Assert NoOfThreads > 1
  Debug.Assert HashMapBestPVPtr(ThreadNum) + CLng(LenB(PV(0, 0)) * 10) < HashMapBoardPtr
  ' Use PV0 to store some values... not nice...
  Erase BestPV
  If CompletedDepth > 0 Then

    For i = 0 To GetMin(9, PVLength(1)): BestPV(i) = PV(1, i): Next
    If BestPV(1).From = 0 Then
      ' use BestMove instead
      BestPV(1) = BestMove: BestPV(0).From = 1
    End If
  End If
  BestPV(0).Target = CompletedDepth: BestPV(0).SeeValue = BestScore: BestPV(0).From = GetMin(9, PVLength(1)): BestPV(0).OrderValue = Nodes
  If bThreadTrace Then WriteTrace "WriteMapBestPVforThread: D:" & CompletedDepth & ", PV:" & MoveText(BestPV(1)) & " / " & Now()
  moHashMap.WriteMapPos HashMapBestPVPtr(ThreadNum), CLng(VarPtr(BestPV(0))), CLng(LenB(BestPV(0)) * 10)
End Function

Public Function ReadMapBestPVforThread(ByVal SelThread As Long, _
                                       ByRef CompletedDepth As Long, _
                                       ByRef BestScore As Long, _
                                       ByRef BestPVLength As Long, _
                                       ByRef HelperNodes As Long, _
                                       BestPV() As TMOVE) As Boolean
  ' Write PV from helper thread for main thread
  Debug.Assert NoOfThreads > 1
  Debug.Assert HashMapBestPVPtr(SelThread) + CLng(LenB(BestPV(0)) * 10) < HashMapBoardPtr
  ReadMapBestPVforThread = False
  Erase BestPV
  ' Use PV0 to get some values... not nice...
  moHashMap.ReadMapPos HashMapBestPVPtr(SelThread), CLng(VarPtr(BestPV(0))), CLng(LenB(BestPV(0)) * 10)
  CompletedDepth = BestPV(0).Target: BestScore = BestPV(0).SeeValue: BestPVLength = BestPV(0).From: HelperNodes = BestPV(0).OrderValue
  If BestPV(1).From = 0 Or BestPV(1).Target = 0 Then
    If bThreadTrace Then WriteTrace "!!!???ReadMapBestPVforThread:PV Empty Thread:" & SelThread & ", Completed Depth:" & CompletedDepth
  End If
  If bThreadTrace Then WriteTrace "ReadMapBestPVforThread: PV:" & MoveText(BestPV(1)) & " / " & Now()
  ReadMapBestPVforThread = (BestPVLength > 0)
End Function

Public Function SetThreads(ByVal iMaxThreads As Long)
  ' set thread numbers: 1-4
  NoOfThreads = GetMax(1, Val("0" & Trim$(ReadINISetting("THREADS", "1"))))
  NoOfThreads = GetMax(NoOfThreads, iMaxThreads)
  NoOfThreads = GetMin(NoOfThreads, MAX_THREADS)
'NoOfThreads = 2 '#######testc2
  If NoOfThreads <= 1 Then
    ThreadNum = -1 ' Single core mode
  Else
    ThreadNum = 0 ' main thread
  End If
  'WriteTrace "SetThreads= " & NoOfThreads & " / " & Now()
End Function

Public Function MaterialHashCompute(ByVal Key As Long) As Long
  If Key = -2147483648# Then Key = Key + 1
  MaterialHashCompute = Abs(Key) Mod MATERIAL_HASHSIZE
End Function

Public Function SaveMaterialHash(ByVal Key As Long, ByVal Score As Long)
  Dim Index As Long
  Index = MaterialHashCompute(Key)

  With MaterialHash(Index)
    .Hashkey = Key
    .Score = Score
  End With

End Function

Public Function ProbeMaterialHash(ByVal Key As Long) As Long
  Dim Index As Long
  Index = MaterialHashCompute(Key)

  With MaterialHash(Index)
    If .Hashkey = Key Then
      ProbeMaterialHash = .Score
    Else
      ProbeMaterialHash = VALUE_NONE
    End If
  End With

End Function

Public Function InIDE() As Boolean
   ' running IDE ( VB development environment) ? if compiled EXE returns false
    Static i As Byte
    i = i + 1
    If i = 1 Then Debug.Assert Not InIDE()
    InIDE = (i = 0)
    i = 0
End Function

Public Function GetAppTimeString() As String
  ' returns exe filedatetime with digits only
  Dim p As Long, s As String
  GetAppTimeString = ""
  s = Now()
  #If VBA_MODE = 0 Then
    If Dir(App.EXEName & ".exe") <> "" Then
      s = FileDateTime(App.EXEName & ".exe")
    End If
  #End If
  For p = 1 To Len(s)
    If IsNumeric(Mid$(s, p, 1)) Then GetAppTimeString = GetAppTimeString & Mid$(s, p, 1)
  Next
End Function




