Attribute VB_Name = "DebugBas"
Option Explicit
'==================================================
'= DebugBas:
'= Debug functions
'==================================================

Public Function DEGUBPrintMoveList(MoveList() As TMOVE) As String
  Dim i        As Long
  Dim strMoves As String

  Do While Not MoveList(i).From = 0
    strMoves = strMoves & vbTab & MoveText(MoveList(i))
    i = i + 1
    If i Mod 3 = 0 Then strMoves = strMoves & vbCrLf
  Loop

  DEGUBPrintMoveList = strMoves
End Function

Public Sub DEBUGPerfTestSearch(ByVal iDepth As Long)
  Dim NumMoves As Long
  Dim i        As Long
  If iDepth = 0 Then Exit Sub
  Ply = Ply + 1
  GenerateMoves Ply, False, NumMoves

  For i = 0 To NumMoves - 1
    MakeMove Moves(Ply, i)
    If CheckLegal(Moves(Ply, i)) Then
      Nodes = Nodes + 1
      DEBUGPerfTestSearch iDepth - 1
    End If
    UnmakeMove Moves(Ply, i)
  Next

  Ply = Ply - 1
End Sub

Public Function DEBUGPerfTest(ByVal iDepth As Long) As String
  Dim strResult As String, StartTime As Single, EndTime As Single
  InitGame
  Ply = 1
  bWhiteToMove = True
  Nodes = 0
  StartTime = Timer
  DEBUGPerfTestSearch iDepth
  EndTime = Timer
  ' time for move generation
  strResult = "time: " & Format$(EndTime - StartTime, "0.00") & " nodes: "

  ' show correct move counts until depth 5
  Select Case iDepth
    Case 1
      strResult = strResult & Nodes & " (expected: 20)"
    Case 2
      strResult = strResult & Nodes - 20 & " (expected: 400)"
    Case 3
      strResult = strResult & Nodes - 400 - 20 & " (expected: 8902)"
    Case 4
      strResult = strResult & Nodes - 8902 - 400 - 20 & " (expected: 197281)"
    Case 5
      strResult = strResult & Nodes - 197281 - 8902 - 400 - 20 & " (expected: 4865609)"
  End Select

  DEBUGPerfTest = strResult
End Function

Public Sub DEBUGBench(ByVal iDepth As Long)
  ' ORIGINAL
  Dim i         As Long, StartTime As Single, EndTime As Single, x As Long, c As Long, s As String
  Dim arTime(2) As Single, EPD(10) As String
  '--- Test positions -----
  'EPD(1) = "r1b1kb1r/pppp1ppp/2n1pq2/8/2PP4/P1P2N2/4PPPP/R1BQKB1R w KQkq - 1 7 " ' SF6 problem: Too high eval until ply 7
  ' EPD(1) = "rn1q4/pbp2kp1/1p1ppn2/8/1PP5/P5Q1/3PPP1r/R1B1KBR1 b Q b3 0 11" ' too high KSafety eval
  'EPD(1) = "3r2k1/p1q1r2p/bppb2p1/6Qn/2NPp3/1PN1Pp1P/PB3PP1/2R3RK b - - 3 27 " '  King attack eval too high <<<
  ' EPD(1) = "r3k3/p2nbpp1/bpp1p3/3nP3/2NP3P/1PB4P/P1Q2PBq/R3RK2 w q - 1 20 " ' KS eval
  'EPD(1) = "r4r2/p1q1n1kp/2n1ppp1/8/3P2N1/3BPP2/2Q2P1P/R3K2R w KQ - 0 19 " ' Trapped knight h3/h4
  'EPD(1) = "r4rk1/1p2ppbp/pq1p2p1/3P4/1nP3n1/2N2N2/PP2QPPP/R1B2RK1 b - - 0 18 " ' Trapped knight a5
  'EPD(1) = "rnbq1rk1/ppp2pp1/8/2npP2Q/1P6/8/P1PN1PPP/R1B2RK1 b - b3 0 11"
  'EPD(1) = "rnbq1r2/ppp2ppk/8/2npP2Q/8/8/PPPN1PPP/R1B2RK1 b - - 1 10 "
  'EPD(1) = "3r1r1k/1b2b1p1/1p5p/2p1Pp2/q1B2P2/4P2P/1BR1Q2K/6R1 b - - 0 1 " ' Eval BEnch
  'EPD(1) = "8/p6p/4k1p1/3p4/2p4P/Pr3PK1/R5P1/8 b - - 1 41 " ' Passed Pawn eval
  'EPD(1) = "r2qr1k1/p3bppp/bpn2n2/2pp4/3P1B2/1PN2NP1/P3PPBP/2RQ1RK1 w - - 8 1 " ' SEE problem
  'EPD(1) = "r1b1k2r/1pp1q2p/p1n3p1/3QPp2/8/1BP3B1/P5PP/3R1RK1 w kq - 0 1 " ' WAC133
  'EPD(1) = "r3kbr1/1p3p1b/pq4Pp/3pp1n1/3PP1N1/PQ4pP/1P3P1B/R3KBR1 w Qq - 0 1  " 'Eval Test symmetric 2
  
  'EPD(1) = "rnbqkbnr/1pp2pp1/p6p/3pp3/3PP3/P6P/1PP2PP1/RNBQKBNR w KQkq - 0 1 " ' Eval Test symmetric 1
  'EPD(1) = "8/5K2/8/3N4/8/8/7k/8 w - - 0 4" 'endgame test
  'EPD(1) = "8/6R1/8/4k1K1/8/8/3r4/8 w - - 3 3 " ' draw test
  'EPD(1) = "r1bq3r/ppppR1p1/5n1k/3P4/6pP/3Q4/PP1N1PP1/5K1R w - - 0 1 " ' WAC138
  'EPD(1) = "8/7p/7k/8/1PK5/8/8/8 w - - 0 1  " ' endgame  pawn promote
  'EPD(1) = "8/8/8/8/6p1/6Pp/5k2/7K w - - 2 95 "  ' bug hanging movepicker => one legal move out of check
  'EPD(1) = "r2qk2r/pp1n1ppp/2p1p3/5b2/P2Pn3/BBP1P3/3N1PPP/R3QRK1 w kq - 0 14 " ' Eval ?
  'EPD(1) = "2r5/7K/k5P1/8/8/1p6/8/8 b - - 0 1 " ' Passed pawn test
  'EPD(1) = "3R4/p6r/8/1P2k3/2B5/8/4K3/8 w - - 50 103  " ' endgame king to pawn
  'EPD(1) = "r1b2rk1/p4ppp/1p1Qp3/4P2N/1P6/8/P3qPPP/3R1RK1 w - - 0 1 " ' WAC 288
  'EPD(1) = "8/8/8/Q7/8/2K3k1/7r/8 w - - 0 1 " ' KQKR
  'EPD(1) = "8/8/8/Q7/8/2K3k1/7p/8 w - - 0 1 " ' KQKP
  'EPD(1) = "8/8/8/5pk1/8/2KR4/8/8 w - - 0 1" ' KRKP
  'EPD(1) = "2qrr1n1/3b1kp1/2pBpn1p/1p2PP2/p2P4/1BP5/P3Q1PP/4RRK1 w - - 0 1" ' ; e2h5 "BWTC.0031"
  ' EPD(1) = "5rk1/1pp3bp/3p2p1/2PPp3/1P2P3/2Q1B3/4q1PP/R5K1 b - - bm Bh6; id WAC.169"
  'EPD(1) = "8/7p/1R4pk/8/6PK/7P/1p6/1r6 b - - 3 1 " ' Passed pawn attacked by rook   SF6: mg:1.14 eg:2.24 cp
  'EPD(1) = "8/7p/1R4pk/8/6PK/7P/1pr5/8 b - - 0 1  " ' Passed pawn attacked by rook, blocked by own rook  SF6: 1.38 2.36
  'EPD(1) = "8/7p/3R2pk/8/1r4PK/7P/1p6/8 w - - 0 1 " ' Passed pawn defended by rook  SF6: 2.54  3.97
  ' EPD(1) = "r3r1k1/pbq2p2/4p2p/1p1nP2Q/2pR4/2P5/PPB2PPP/4R1K1 w - - 0 20 " ' Defend
  'EPD(1) = "r3r1k1/pbq2pp1/4p2B/1p1nP2Q/2pR4/2P5/PPB2PPP/4R1K1 b - - 0 19 " ' Attack f7f5 (g7xh6 bad)
  'EPD(1) = "r1bqkbnr/ppp2ppp/2np4/4p3/2B1P3/5N2/PPPP1PPP/RNBQK2R w KQkq - 2 4 " ' KSafety/Castle eval
  'EPD(1) = "rnbq1rkr/pppp1p1p/5n2/2b1p3/4P3/2NP4/PPP2PPP/R1BQKBNR b KQ - 2 1 " ' KSafety/Castle eval- Black
  ' EPD(1) = "6k1/6p1/8/8/8/8/4P2P/6K1 b - -" ' Test Endgame Tablebase acces in search for root
  'EPD(1) = "8/6k1/6p1/8/7r/3P1KP1/8/8 w - - 0 1 "  ' Test Endgame Tablebase acces in search for ply=1
  ' EPD(1) = "r3k2r/pb3pbp/2p1p3/1q2p3/2p5/6P1/1PQ1PPBP/R1BR2K1 w kq - 0 2 "
  'EPD(1) = "2r1r1k1/4bp1p/p2pp1pP/q3n1P1/Np1Nb3/1P2B3/P1PQ4/1K2RBR1 b - - 1 21 " ' e5f3 not found
  ' EPD(1) = "2r1r1k1/4bp1p/p2pp1pP/q5P1/Np2b3/1P2BN2/P1PQ4/1K2RBR1 b - - 0 21 "
  '   EPD(1) = "4r1k1/4bp1p/p2pp1pP/q5P1/Np2b3/1P2BN2/P1rQ4/1K2RBR1 w - - 0 22"  ' d2xc2 ok, d2d4 >Rc2c4 illegal move, IsCHecking no detected
  'EPD(1) = "r2r2k1/pb3p1p/1qn1p2Q/5p2/1p1P4/1NPB4/P4PPP/2R1R1K1 b - - 0 22 " ' KSafety test
  ' EPD(1) = "8/5pk1/1p4Pp/q6P/Q7/1P6/8/6K1 b - - 0 1 " ' ShelterStorm test
  'EPD(1) = "5k2/6b1/8/4N3/8/8/3P1K2/8 w - - 3 1  " ' Scale factor 1 pawn test
  ' EPD(1) = "r1b2r1k/p5pp/2nq4/Ppp1pp2/2Bn1N1Q/2B1R3/2P2PPP/R5K1 w - b6 0 2 " ' EnPassant test
  ' EPD(1) = "r1b2r1k/pp4pp/2nq4/P1p1pp2/2Bn1N1Q/2B1R3/2P2PPP/R5K1 b - - 1 1" ' EnPassant test2 move b7b5
  ' EPD(1) = "6k1/4Q1p1/7p/8/nn6/1p3R2/5PPP/6K1 w - - 1 1   " ' mate threat
  ' EPD(1) = "8/2pp4/3kPKP1/3P4/8/8/8/8 w - - 0 1 "
  'EPD(1) = "8/8/2k5/8/5K2/3R4/8/3qR3 w - - 0 1"  ' EGTB
  'EPD(1) = "8/3PK3/8/5p1k/8/8/8/8 b - - 0 w " ' EGTB test promotion
  'EPD(1) = "8/5PK1/8/2Q5/4P1k1/8/8/8 b - - 0 14 "
  'EPD(1) = "8/4k3/8/8/5P2/5K2/8/8 b - - 4 3 "  ' EGTB KPK
  'EPD(1) = "8/3k3K/7P/1r6/5p2/8/8/8 b - - 0 1 "
  'EPD(1) = "8/8/3R4/p3npk1/P3p2p/4P3/3K1PP1/r2B4 w - - 8 39 " ' EP capture mate bug
  'EPD(1) = "8/2b5/8/4kN2/1r4K1/6N1/8/8 w - - 0 1"  ' endgame scale factor no pawns
  'EPD(1) = "8/8/7k/p1P4p/P6P/7K/8/8 w - - 0 1"  ' passed pawn test 1 rank 5
  'EPD(1) = "8/8/7k/p1P4p/P6P/7K/8/2R5 w - - 0 1"  ' passed pawn test 2 defended from behind
  'EPD(1) = "8/8/7k/p1P4p/P6P/7K/8/2r5 w - - 0 1"  ' passed pawn test 3 attacked from behind
  'EPD(1) = "8/7r/7k/p1P4p/P6P/7K/8/2R5 w - - 0 1"  ' passed pawn test 4 defended from behind + attacked path
  'EPD(1) = "8/8/2P4k/p6p/P6P/7K/8/2R5 w - - 0 1"  ' passed pawn test 5 defended from behind rank 6
  'EPD(1) = "7k/5K1p/7P/8/8/8/8/8 b - - 1 1" ' no move draw
 ' EPD(1) = "r5k1/pp4pp/2pb3r/3p2q1/P1PP1nB1/1PB1P1PP/7K/R2Q2R1 b - - 0 27" ' KSafety
  
  'EPD(1) = "5rk1/pp4pp/2pb3r/3p2q1/P1PP4/1PB1P1PB/7K/R4QR1 b - - 2 29" '
  'EPD(1) = "6k1/4b1p1/5p1Q/1p2pP2/4P3/1P6/6PP/6rK w - - 0 41 " 'only one legalmove
  'EPD(1) = " /4b1p1/8/1p2p3/8/7p/5p2/6bK w - - 0 1 " 'no legal move
  
  'EPD(1) = "8/8/7k/8/8/8/6PP/3r3K w - - 1 1  " ' mated result : bestmove (none)
  
  'EPD(1) = "r5k1/pp4pp/2pb3r/3p2q1/P1PP4/1PB1P1PB/7K/R2Q2R1 b - - 0 28 "
   'EPD(1) = "6r1/2pq2pk/1p3p1p/1P1Pp2P/Q3P1P1/p1R3K1/P7/8 w - - 98 109 " ' fifty
  ' EPD(1) = "k7/8/P7/1K6/8/8/8/8 w - - 12 1 " ' endgame kpk"
  ' EPD(1) = "2nk4/8/8/8/6b1/8/8/R3K3 w Q - 0 1 " ' checking with castling
  '------ normal test ----
  EPD(1) = "1b5k/7P/p1p2np1/2P2p2/PP3P2/4RQ1R/q2r3P/6K1 w - - 0 1"
  EPD(2) = "1rb2rk1/p3nppp/1p1qp3/3n2N1/2pP4/2P3P1/PP3PBP/R1BQR11K w - -"  'TEST 2
  EPD(3) = "r1b2rk1/p2nq1p1/1pp1p2p/5p2/2PPp3/2Q1P3/PP1N1PPP/2R1KB1R w K - 0 13" '--- quiet
  EPD(4) = "6k1/p1r5/4b1p1/R1pprp1p/7P/1P1BP3/P1P3P1/4R1K1 w - - 4 25" ' no advantage
  EPD(5) = "8/8/2R5/1p2qp1k/1P2r3/2PQ2P1/5K2/8 w - - 0 1" ' Endgame
  EPD(6) = "r7/pbk5/1pp5/4n1q1/2P5/1P6/P4BBQ/4R1K1 b - - 0 33" '
  EPD(7) = "r1bqk2r/p2p1pp1/1p2pn1p/n1pP2B1/1bP5/2N2N2/PPQ1PPPP/R3KB1R w KQkq - 0 9" '<<<<< AKT
  '-------------------------------------------------------------------------------------
  DebugMode = True

  'iDepth = 8
  ' ReadGame "Drawbug2.txt"
  'bForceMode = False
  'For x = 1 To 1 'if EPD(1) only to test
   For x = 1 To 7 ' 7
    For i = 0 To 0 ' number of time measure runs  > 1x
      'For i = 0 To 2 ' number of time measure runs > 3x
      InitGame ' Reset FixedDepth
      ReadEPD EPD(x) ' Reset FixedDepth
      If True Then
        If x = 3 Or x = 4 Or x = 5 Or x = 7 Then
          FixedDepth = iDepth + 1
        Else
          FixedDepth = iDepth
        End If
      ' Else
      ' FixedTime = 4
       MovesToTC = 0: TimeLeft = 20: TimeIncrement = 10
      End If
      If InStr(EPD(x), " w") > 0 Then
        bCompIsWhite = True   'False:
        bWhiteToMove = True   '---False
      Else
        bCompIsWhite = False ' True  'False:
        bWhiteToMove = False ' True  '---False
      End If
      ' ParseCommand "b7b5"
      bPostMode = True
      '    bPostMode = False
      'SendCommand PrintPos
      If False Then  ' Time based end of thinking
        FixedDepth = NO_FIXED_DEPTH
        LevelMovesToTC = 40
        MovesToTC = 0
        TimeLeft = 120
        TimeIncrement = 0
        GameMovesCnt = 119 ' plies, /2 for MoveCnt
        BookPly = 31
      End If
      StartTime = Timer
      StartEngine
      EndTime = Timer
      ' Test Counter
      s = ""

      For c = 1 To 19
        If TestCnt(c) <> 0 Then s = s & CStr(c) & ":" & TestCnt(c) & ","
      Next c

      arTime(i) = EndTime - StartTime
      If arTime(i) = 0 Then arTime(i) = 1
      bPostMode = True
      SendCommand vbCrLf & "time: " & Format$(arTime(i), "0.000") & " nod: " & Nodes & " qn: " & QNodes & " ev:" & EvalCnt & " sc: " & FinalScore & " EGTB:" & EGTBasesHitsCnt & " Ply:" & MaxPly & " " & s & vbCrLf
    Next

    If arTime(0) < arTime(1) Then
      If arTime(0) < arTime(2) Then
        i = 0
      Else
        i = 2
      End If
    ElseIf arTime(1) < arTime(2) Then
      i = 1
    Else
      i = 2
    End If
    ' count 3x
    If arTime(i) > 0 Then
      SendCommand "best time: " & Format$(arTime(i), "0.000") & " nps: " & Int(Nodes / arTime(i))
    Else
      SendCommand "best time: " & Format$(0, "0.000") & " nps: " & Nodes
    End If
    SendCommand "Hash usage:" & Format((CDbl(HashUsage) / CDbl(HashSize)) * 100#, "0.00")
    SendCommand "-------------------"
  Next x

End Sub

Public Sub WriteDebug(s As String)
  Debug.Print s
End Sub

Public Sub DMoves()
  ' Debug: print current move line
  Dim i As Long, s As String
  s = CStr(IterativeDepth) & "/" & CStr(Ply) & ">"

  For i = 1 To Ply - 1
    s = s & CStr(i) & ":" & MoveText(MovesList(i)) & "/"
  Next

  Debug.Print s
  DoEvents
End Sub

Public Sub DEBUGLoadGame(ByVal iDepth As Long)
  ' ORIGINAL
  Dim i         As Long, StartTime As Single, EndTime As Single, x As Long, c As Long, s As String
  Dim arTime(2) As Single
  iDepth = 8
  DEBUGReadGame "bug001.txt"
  bForceMode = False

  For i = 0 To 0 ' number of time measure runs  > 1x
    FixedDepth = iDepth
    bCompIsWhite = False ' True  'False:
    bWhiteToMove = False ' True  '---False
    bPostMode = True
    StartTime = Timer
    StartEngine
    EndTime = Timer
    arTime(i) = EndTime - StartTime
    If arTime(i) = 0 Then arTime(i) = 1
    bPostMode = True
    SendCommand vbCrLf & "time: " & Format$(arTime(i), "0.000") & " nod: " & Nodes & " qn: " & QNodes & " ev:" & EvalCnt & " sc: " & EvalSFTo100(FinalScore) & " Ply:" & MaxPly & " " & s & vbCrLf
  Next

  SendCommand "-------------------"
End Sub

Public Sub DEBUGReadGame(sFile As String)
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
    s = Trim(sInp)
    'Debug.Print s
    m1 = Trim(Left(s, 4))
    If Len(m1) = 4 Then
      'Debug.Print m1, asMoveList(m)
      ParseCommand m1 & vbLf
    End If
  Loop

  Close #h
End Sub
