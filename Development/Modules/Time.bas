Attribute VB_Name = "TimeBas"
Option Explicit
'=======================
'= TimeBas:
'= Time management
'=======================
Public bTimeExit                  As Boolean
Public TimeStart                  As Single
Public SearchStart                As Single
Public SearchTime                 As Single
Public ExtraTimeForMove           As Single
Public TimeLeft                   As Single
Public OpponentTime               As Single
Public TimeIncrement              As Long
Public LevelMovesToTC             As Long
Public MovesToTC                  As Long
Public SecondsPerGame             As Long
Public FixedDepth                 As Long  '=NO_FIXED_DEPTH if time limit is used
Public FixedTime                  As Single
Public LastChangeDepth            As Long, LastChangeMove As String
Public bResearching               As Boolean '--- out of aspiration windows: more time
Public BestMoveChanges            As Single ' More time if best move changes often
Public MaximumTime                As Single
Public OptimalTime                As Single
Public MoveOverhead               As Single
Public MoreTimeForFirstMove       As Boolean  ' fill Hash table


'----------------
'- AllocateTime()
'----------------

Public Sub AllocateTime()
 Dim GameMovesDone As Long

  If bTimeTrace Then
    WriteTrace " ---------------------------------------------"
    WriteTrace ">> Start AllocateTime  MTOC:" & MovesToTC & ", MoveCnt=" & CStr(GameMovesCnt) & ", Left:" & Format$(TimeLeft, "0.00")
  End If
  
  GameMovesDone = (GameMovesCnt + 1) \ 2 ' Full move = 2* Half move
  
  If Not UCIMode And LevelMovesToTC > 0 Then
    MovesToTC = LevelMovesToTC - (GameMovesDone Mod LevelMovesToTC)
    If bTimeTrace Then
       WriteTrace "CalcTime WB: LevelMovesToTC=" & LevelMovesToTC & ", MovesToTC=" & MovesToTC & ", MovesDone:" & GameMovesDone
    End If
  End If
  '
  OptimalTime = CalcTime(MovesToTC, TimeIncrement, TimeLeft, MoveOverhead + 0.05 * (NoOfThreads - 1), True)
  MaximumTime = CalcTime(MovesToTC, TimeIncrement, TimeLeft, MoveOverhead + 0.05 * (NoOfThreads - 1), False)
  '
  MaximumTime = GetMinSingle(MaximumTime, TimeLeft / 2#)
  OptimalTime = GetMinSingle(MaximumTime, OptimalTime)
  
  If OptimalTime < 0.2 Then
    OptimalTime = GetMinSingle(0.2 + 0.05 * NoOfThreads, 1#): MaximumTime = OptimalTime
  End If
  MoreTimeForFirstMove = False
  If bTimeTrace Then
    WriteTrace ">>>> Time allocated Opt: " & Format$(OptimalTime, "0.00") & " / Max:" & Format$(MaximumTime, "0.00") & " MTOC:" & MovesToTC & " MoveCnt=" & CStr(GameMovesCnt) & ", Left:" & Format$(TimeLeft, "0.00")
  End If
End Sub


Public Function CalcTime(ByVal MovesToTC As Long, _
                         ByVal TimeIncr As Single, _
                         ByVal MyTime As Single, _
                         ByVal MoveOverhead As Single, _
                         ByVal TimeTypeIsOptimum As Boolean) As Single
 Dim Ratio As Single, Inc As Single, k As Single, SafetyMargin As Single
 Dim GameMovesDone As Long
 
 GameMovesDone = (GameMovesCnt + 1) \ 2 ' Full move = 2* Half move
 
 If MyTime <= 0 Then CalcTime = 0: Exit Function
  
 Inc = TimeIncr * GetMaxSingle(60#, 125# - 0.1 * CSng((GameMovesDone - 23) * (GameMovesDone - 23)))
 SafetyMargin = 1.5
 
 If MovesToTC > 0 Then
   If TimeTypeIsOptimum Then Ratio = 1# Else Ratio = 5#
   Ratio = Ratio / CSng(GetMin(45, GetMax(1, MovesToTC)))
   
   If GameMovesDone <= 40 Then
     Ratio = Ratio * (1.3 - 0.001 * CSng((GameMovesDone - 23) * (GameMovesDone - 23)))
   Else
     Ratio = Ratio * 1.55
   End If
   Ratio = Ratio * (1# + Inc / (MyTime * 8.2))
   If MovesToTC <= 3 Then SafetyMargin = 3#
 Else
   k = 1# + 21# * CSng(GameMovesDone) / CSng(500 + GameMovesDone)
   If TimeTypeIsOptimum Then Ratio = 0.019 Else Ratio = 0.072
   Ratio = Ratio * (k + Inc / MyTime)
 End If
 If MoreTimeForFirstMove Then Ratio = Ratio * 1.5
 '
 CalcTime = GetMinSingle(1#, Ratio) * GetMaxSingle(0.01, MyTime - MoveOverhead - SafetyMargin - TimeIncr / 10#)
End Function


Public Function TimerDiff(ByVal StartTime As Single, ByVal EndTime As Single) As Single
  If StartTime - 0.1 > EndTime Then ' Timer resets to 0 ad midnight > EndTime > Startime
    EndTime = EndTime + CSng(60& * 60& * 24&)
  End If
  TimerDiff = EndTime - StartTime
  If TimerDiff < 0 Then TimerDiff = 0.1
End Function

Public Function TimeElapsed() As Single
 TimeElapsed = TimerDiff(TimeStart, Timer())
End Function

Public Function CheckTime() As Boolean
  Dim Elapsed As Single, Improve As Single, Optimum2 As Single, NewScore As Long, PrevScore As Long
  CheckTime = True
  
  Elapsed = TimeElapsed()
  If FinalScore = UNKNOWN_SCORE Then NewScore = 0 Else NewScore = FinalScore
  If PrevGameMoveScore = UNKNOWN_SCORE Then PrevScore = FinalScore - 80 Else PrevScore = PrevGameMoveScore

  Improve = GetMaxSingle(229#, GetMinSingle(715#, 357# + 119# * Abs(bFailedLowAtRoot) - 6# * CSng(NewScore - PrevScore)))
  Optimum2 = (OptimalTime * (1# + BestMoveChanges) * Improve) / 628#
  
  If Elapsed >= GetMinSingle(MaximumTime, Optimum2) Then
    CheckTime = False
    If bTimeTrace Then
        WriteTrace "CheckTime D" & IterativeDepth & ": Elapsed:" & Format$(Elapsed, "0.00") & ", Opt2:" & Format$(Optimum2, "0.00") & ", Opt:" & Format$(OptimalTime, "0.00") & ", Max:" & Format$(MaximumTime, "0.00")
    End If
  End If

End Function
