Attribute VB_Name = "TimeBas"
Option Explicit
'=======================
'= TimeBas:
'= Time management
'=======================
'----------------
'- AllocateTime()
'----------------
Public TimeLeftCorr         As Single
Public LastAddTimeMoveCnt   As Long  ' minor time adjustments
Public LastExtraTimeMoveCnt As Long ' big extra time if in trouble

Public Function AllocateTime(ByVal CurrScore As Long) As Single
  Dim GameMovesDone As Long, RemainingMovesToTC As Long, TimeBase As Single, Score As Long
  If bTimeTrace Then WriteTrace ">> Start AllocateTime  MTOC:" & MovesToTC & ", MoveCnt=" & CStr(GameMovesCnt) & ", Left:" & Format$(TimeLeft, "0.00")
  InitEval
  Score = Eval()
  If bWhiteToMove And Not bCompIsWhite Then Score = -Score
  GameMovesDone = GameMovesCnt \ 2 ' Full move = 2* Half move
  If MovesToTC = 0 Then RemainingMovesToTC = 0 Else RemainingMovesToTC = MovesToTC - (GameMovesDone Mod MovesToTC)
  If UCIMode Then RemainingMovesToTC = MovesToTC ' UCI sends remaining moves to TC
  If bTimeTrace Then WriteTrace "before CalcTime: RMTOC:" & RemainingMovesToTC & " MToTC:" & MovesToTC
  ' Subtract overhead for move transfer/execution
  If MovesToTC > 0 And MovesToTC < 10 Then
    '-- Compensate move overhead ( difference WB2 time value diffenrence ,inis engine used time)
    ' Arena 0.2 sec, ChessGU 0.4 sec overhead; for thread handling extra 0.1
    TimeLeftCorr = TimeLeft - (0.5 + Abs(NoOfThreads > 1) * 0.1) * CDbl(GetMax(1, GetMin(40, RemainingMovesToTC)))
  Else
    TimeLeftCorr = TimeLeft - TimeLeft / 20 - 0.5 * CDbl(GetMax(1, GetMin(40, RemainingMovesToTC)))
  End If
  If TimeLeftCorr < 0.1 Then TimeLeftCorr = 0.2
  AllocateTime = CalcTime(RemainingMovesToTC, TimeIncrement, TimeLeftCorr, CurrScore)
  If MovesToTC > 0 Then
    TimeBase = TimeLeftCorr / CDbl(GetMax(1, RemainingMovesToTC))
    If RemainingMovesToTC < 10 Then
      AllocateTime = TimeBase * 0.9
    Else
      If AllocateTime > TimeBase * 4# Then
        If bTimeTrace Then WriteTrace "Allocate TimeBase*4 limit. " & Format$(AllocateTime, "0.00")
        AllocateTime = TimeBase * 4#
      End If
      If (TimeLeftCorr - AllocateTime) / CDbl(GetMax(1, RemainingMovesToTC)) < TimeBase \ 2# Then
        If bTimeTrace Then WriteTrace "Allocate Timebase\2 limit. " & Format$(AllocateTime, "0.00")
        AllocateTime = TimeLeftCorr / CDbl(GetMax(1, RemainingMovesToTC))
      End If
    End If
  End If
  If AllocateTime > (TimeLeftCorr + TimeIncrement) / 2# Then AllocateTime = (TimeLeftCorr + TimeIncrement) / 2#
  If TimeLeftCorr < 1.5 And TimeIncrement = 0 And AllocateTime > 0.2 Then
    AllocateTime = GetMaxSingle(TimeLeftCorr, 0.1 + TimeLeftCorr / (GetMax(1, RemainingMovesToTC)))
    AllocateTime = GetMinSingle(AllocateTime, TimeLeftCorr)
  End If
  If (TimeLeftCorr - AllocateTime) / CDbl(GetMax(1, RemainingMovesToTC)) < 0.8 Then
    AllocateTime = (TimeLeftCorr - 0.2) / CDbl(GetMax(1, RemainingMovesToTC))
    If bTimeTrace Then WriteTrace "Average < 0.5 " & Format$(AllocateTime, "0.00")
  End If
  If MovesToTC > 1 And RemainingMovesToTC = 1 And TimeLeftCorr > 0.5 And AllocateTime < TimeLeftCorr * 0.75 Then
    AllocateTime = GetMaxSingle((TimeLeftCorr - 0.9) * 0.8, 0.5)
    If bTimeTrace Then WriteTrace "RMTOC=1 < TimeLeftCorr*0.8 " & Format$(AllocateTime, "0.00")
  End If
  AllocateTime = GetMinSingle(AllocateTime, TimeLeftCorr - 0.2)
  If AllocateTime < 0.2 Then AllocateTime = 0.2
  If DebugMode Then
    AllocateTime = 90
  End If
  If bTimeTrace Then
    WriteTrace ">>>> Time allocated: " & Format$(AllocateTime, "0.00") & " MTOC:" & MovesToTC & "/RMTOC" & RemainingMovesToTC & ", MoveCnt=" & CStr(GameMovesCnt) & ", Left:" & Format$(TimeLeft, "0.00") & ", LeftCorr:" & Format$(TimeLeftCorr, "0.00")
    WriteTrace " -------------------"
  End If
End Function

Public Function AllocateExtraTime() As Boolean
  Dim GameMovesDone As Long, RemainingMovesToTC As Long, TimeBase As Single
  GameMovesDone = GameMovesCnt \ 2 ' Full move = 2* Half move
  If MovesToTC = 0 Then RemainingMovesToTC = 0 Else RemainingMovesToTC = MovesToTC - (GameMovesDone Mod MovesToTC)
  If UCIMode Then RemainingMovesToTC = MovesToTC ' UCI sends remaining moves to TC
  TimeBase = TimeLeftCorr / CDbl(GetMax(1, RemainingMovesToTC))
  If GameMovesDone <= LastExtraTimeMoveCnt + 3 Then
    If bTimeTrace Then WriteTrace "ExtraTime ignored:LastExtraTimeMoveCnt "
  End If
  If (MovesToTC >= 0 And RemainingMovesToTC < 3) Then
    If bTimeTrace Then WriteTrace "ExtraTime ignored:RemainMoves <3"
    bExtraTime = True
    Exit Function
  End If
  LastExtraTimeMoveCnt = GameMovesDone
  If MovesToTC > 0 And RemainingMovesToTC < 10 Then
    TimeBase = TimeLeftCorr / CDbl(GetMax(1, RemainingMovesToTC))
    ExtraTimeForMove = TimeBase * 0.2: AllocateExtraTime = True
    TimeForIteration = TimeForIteration + ExtraTimeForMove
    TotalTimeGiven = TotalTimeGiven + ExtraTimeForMove
    If bTimeTrace Then WriteTrace "ExtraTime RMTOC<10: TimeBAse * 0.2"
  Else
    ExtraTimeForMove = CalcExtraTime(TimeBase, TimeIncrement, TimeLeftCorr)
    If TimeForIteration + ExtraTimeForMove > TimeLeftCorr / 3 Then
      ExtraTimeForMove = 0
      If bTimeTrace Then WriteTrace "ExtraTime: set to 0 : >(TimeLeftCorr /3)"
    Else
      TimeForIteration = TimeForIteration + ExtraTimeForMove
      TotalTimeGiven = TotalTimeGiven + ExtraTimeForMove
      If bTimeTrace Then WriteTrace "ExtraTime2+ allocated "
    End If
  End If
  bExtraTime = True
  AllocateExtraTime = CBool(ExtraTimeForMove > 0#)
  If bTimeTrace Then WriteTrace "ExtraTime: " & Format$(ExtraTimeForMove, "0.00")
  AllocateExtraTime = ExtraTimeForMove
End Function

Public Function CalcTime(ByVal RemainingMovesToTC As Long, _
                         ByVal TimeIncr As Single, _
                         ByVal TimeLeftCorr As Single, _
                         ByVal CurrScore As Long) As Single
  CalcTime = CalcTimeLimit(RemainingMovesToTC, TimeIncr, TimeLeftCorr, CurrScore)
  OptimalTime = CalcTime
  MaximumTime = GetMax(OptimalTime, GetMin(OptimalTime * 3, TimeLeftCorr / 2))
End Function

Public Function CalcTimeLimit(ByVal RemainingMovesToTC As Long, _
                              ByVal TimeIncr As Single, _
                              ByVal TimeLeftIn As Single, _
                              ByVal CurrScore As Long)
  Dim TimeTarget As Single, CalcMTOC As Long, GameMovesDone As Long
  TimeLeftIn = GetMaxSingle(TimeLeftIn * 0.8 - 0.5 - (TimeIncr * 0.85), 0#) ' L
  GameMovesDone = GameMovesCnt \ 2
  CalcMTOC = RemainingMovesToTC
  If MovesToTC = 0 Then
    If TimeIncr = 0 Then
      CalcMTOC = 40
      If GameMovesCnt > 40 Then CalcMTOC = 30
    Else
      CalcMTOC = 30
      If GameMovesCnt > 40 Then CalcMTOC = 25
    End If
  Else
    If CalcMTOC > 35 Then CalcMTOC = GetMax(1, GetMin(CalcMTOC - 5, RemainingMovesToTC))
  End If
  TimeTarget = (TimeLeftIn + CDbl(TimeIncr) * CSng(CalcMTOC) * 1.5) / CSng(GetMax(CalcMTOC, 1))
  If bTimeTrace Then WriteTrace "Target:" & Format(TimeTarget, "0.00") & ", Left:" & Format(TimeLeftIn, "0.00") & ", CalcMTOC:" & CStr(CalcMTOC) & ", RMTOC:" & RemainingMovesToTC
  '--- Add time for special cases
  If HashUsage = 0 Then ' first engine move -> fill hash table
    TimeTarget = TimeTarget * 1.5
    LastAddTimeMoveCnt = GameMovesDone
    If bTimeTrace Then WriteTrace "TimeAdd- First move"
  ElseIf CurrScore < -ScoreBishop.EG And GameMovesDone >= LastAddTimeMoveCnt + 3 Then
    TimeTarget = TimeTarget * 2#   ' Win/Loss score
    LastAddTimeMoveCnt = GameMovesDone
    If bTimeTrace Then WriteTrace "TimeAdd- BigDiff"
  ElseIf CurrScore < -ScorePawn.EG * 3 \ 2 And GameMovesDone >= LastAddTimeMoveCnt + 4 Then
    TimeTarget = TimeTarget * 1.5     ' 1.5 pawn minus
    LastAddTimeMoveCnt = GameMovesDone
    If bTimeTrace Then WriteTrace "TimeAdd- <1.3 pawn"
    ' ElseIf (RemainingMovesToTC >= 25 Or TimeIncr > 0) And GameMovesDone < 12 Then
    '   TimeTarget = TimeTarget * 1.1 ' more time during opening
    '   If bTimeTrace Then WriteTrace "TimeAdd-Opening1"
    ' ElseIf (RemainingMovesToTC >= 20 Or TimeIncr > 0) And GameMovesDone < 18 Then
    '   TimeTarget = TimeTarget * 1.2 ' more time during opening
    '   If bTimeTrace Then WriteTrace "TimeAdd-Opening2"
    ' ElseIf (RemainingMovesToTC >= 10 Or TimeIncr > 0) And GameMovesDone < 23 Then
    '   TimeTarget = TimeTarget * 1.3 ' more time during midgame
    '   If bTimeTrace Then WriteTrace "TimeAdd-MidGame1"
  ElseIf (RemainingMovesToTC >= 20 And RemainingMovesToTC = MovesToTC) And GameMovesDone >= 40 Then
    TimeTarget = TimeTarget * GetMinSingle(4#, CDbl(RemainingMovesToTC \ 10)) ' more time when time control reached
    LastAddTimeMoveCnt = GameMovesDone
    If bTimeTrace Then WriteTrace "TimeAdd-TimeControl reached"
  ElseIf (RemainingMovesToTC >= 30) And GameMovesDone > 40 And GameMovesDone >= LastAddTimeMoveCnt + 5 Then
    TimeTarget = TimeTarget * 1.2 ' more time during Endgame
    LastAddTimeMoveCnt = GameMovesDone
    If bTimeTrace Then WriteTrace "TimeAdd-start endgame"
    '  ElseIf (RemainingMovesToTC < 20) And GameMovesDone > 50 Then
    '    TimeTarget = TimeTarget * 0.8 ' less time during Endgame
    '    LastAddTimeMoveCnt = GameMovesDone
    '    If bTimeTrace Then WriteTrace "TimeSubstract- endgame"
    ' ElseIf CurrScore < -ScorePawn.EG Then
    '   TimeTarget = TimeTarget * (1.25)  ' 1 pawn minus
    '   If bTimeTrace Then WriteTrace "TimeAdd-1.0 pawn"
  End If
  If TimeTarget + 0.25 >= TimeLeftCorr Then
    If bTimeTrace Then WriteTrace "Limit2:" & Format(TimeTarget, "0.00") & " : " & Format(TimeLeftCorr, "0.00") * 0.25
    TimeTarget = GetMaxSingle(0.5, TimeLeftIn * 0.75)
  End If
  If TimeTarget < 0.1 Then TimeTarget = 0.1
  CalcTimeLimit = TimeTarget
  If bTimeTrace Then WriteTrace "---TimeLimit> Target:" & Format(TimeTarget, "0.00") & ", MTOC:" & RemainingMovesToTC & ", Left:" & Format(TimeLeftIn, "0.00") & " Inc:" & Format(TimeIncr, "0.00") & " ID:" & IterativeDepth & " / " & Now()
End Function

Public Function CalcExtraTime(ByVal TimeTarget As Single, _
                              ByVal TimeIncr As Single, _
                              ByVal TimeLeftCorr As Single) As Single
  Dim GameMovesDone As Long
  If FixedDepth <> NO_FIXED_DEPTH Then
    CalcExtraTime = 0
  Else
    CalcExtraTime = 0
    GameMovesDone = GameMovesCnt \ 2 ' Full move = 2* Half move
    If (TimeIncr = 0 And TimeLeftCorr > TimeTarget * 5#) Or (TimeIncr > 0 And TimeLeftCorr > TimeTarget * 8#) Then
      CalcExtraTime = TimeTarget * 1.25
      If bTimeTrace Then WriteTrace "ExtraTime+ " & Format$(CalcExtraTime, "0.00") & ", Target:" & Format$(TimeTarget, "0.00")
    Else
      CalcExtraTime = 0
      If bTimeTrace Then WriteTrace "ExtraTime 0"
    End If
  End If
End Function

Public Sub PVInstability()
  UnstablePvFactor = 1# + BestMoveChanges
End Sub

Public Function AvailableTime() As Single
  AvailableTime = TotalTimeGiven * UnstablePvFactor * 0.71
  If bTimeTrace Then WriteTrace "AvailableTime:" & Format(AvailableTime, "0.00") & ", Given:" & Format(TotalTimeGiven, "0.00") & ", Unstable:" & Format(UnstablePvFactor, "0.00")
End Function

Public Function MoveImportance(ByVal GamePly As Long) As Single
  ' SF6: not used
  ' move_importance() is a skew-logistic function based on naive statistical
  ' analysis of "how many games are still undecided after n half-moves". Game
  ' is considered "undecided" as long as neither side has >275cp advantage.
  Const XScale As Single = 9.3
  Const XShift As Single = 59.8
  Const Skew   As Single = 0.172
  MoveImportance = (1 + Exp((GamePly - XShift) / XScale)) ^ -Skew + 0.000001 ' // Ensure non-zero
End Function

Public Function TimerDiff(ByVal StartTime As Single, ByVal EndTime As Single) As Single
  If StartTime - 0.1 > EndTime Then ' Timer resets to 0 ad midnight > EndTime > Startime
    EndTime = EndTime + CSng(60& * 60& * 24&)
  End If
  TimerDiff = EndTime - StartTime
  If TimerDiff < 0 Then TimerDiff = 0.1
End Function
