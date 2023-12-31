Attribute VB_Name = "BitBoard32"
Option Explicit

' bitboard 32 bit with 2 long variables
Public Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByVal Destination As Long, _
                                       ByVal Source As Long, _
                                       ByVal Length As Long)



Public Const MIN_INTEGER  As Integer = -32768
Public Const MAX_INTEGER  As Integer = 32767

Public Const BitL_0 As Long = &H1&
Public Const BitL_1 As Long = &H2&
Public Const BitL_2 As Long = &H4&
Public Const BitL_3 As Long = &H8&
Public Const BitL_4 As Long = &H10&
Public Const BitL_5 As Long = &H20&
Public Const BitL_6 As Long = &H40&
Public Const BitL_7 As Long = &H80&
Public Const BitL_8 As Long = &H100&
Public Const BitL_9 As Long = &H200&
Public Const BitL_10 As Long = &H400&
Public Const BitL_11 As Long = &H800&
Public Const BitL_12 As Long = &H1000&
Public Const BitL_13 As Long = &H2000&
Public Const BitL_14 As Long = &H4000&
Public Const BitL_15 As Long = &H8000&
Public Const BitL_16 As Long = &H10000
Public Const BitL_17 As Long = &H20000
Public Const BitL_18 As Long = &H40000
Public Const BitL_19 As Long = &H80000
Public Const BitL_20 As Long = &H100000
Public Const BitL_21 As Long = &H200000
Public Const BitL_22 As Long = &H400000
Public Const BitL_23 As Long = &H800000
Public Const BitL_24 As Long = &H1000000
Public Const BitL_25 As Long = &H2000000
Public Const BitL_26 As Long = &H4000000
Public Const BitL_27 As Long = &H8000000
Public Const BitL_28 As Long = &H10000000
Public Const BitL_29 As Long = &H20000000
Public Const BitL_30 As Long = &H40000000
Public Const BitL_31 As Long = &H80000000

Public Const RANK1_L = BitL_0 Or BitL_1 Or BitL_2 Or BitL_3 Or BitL_4 Or BitL_5 Or BitL_6 Or BitL_7
Public Const RANK2_L = BitL_8 Or BitL_9 Or BitL_10 Or BitL_11 Or BitL_12 Or BitL_13 Or BitL_14 Or BitL_15
Public Const RANK3_L = BitL_16 Or BitL_17 Or BitL_18 Or BitL_19 Or BitL_20 Or BitL_21 Or BitL_22 Or BitL_23
Public Const RANK4_L = BitL_24 Or BitL_25 Or BitL_26 Or BitL_27 Or BitL_28 Or BitL_29 Or BitL_30 Or BitL_31


Public Type TBit64  ' emulate 64 bit, use 4x16 bit (positive values only)
  i0 As Long
  i1 As Long
End Type

Public Type TInt16x2
  i0 As Integer
  i1 As Integer
End Type

Public Type TInt16x4
  i0 As Integer
  i1 As Integer
  i2 As Integer
  i3 As Integer
End Type

Public Type TByte8x8
  i0 As Byte
  i1 As Byte
  i2 As Byte
  i3 As Byte
  i4 As Byte
  i5 As Byte
  i6 As Byte
  i7 As Byte
End Type

'-----------------------------------------------------------------------------------------

Public FILEA_BB As TBit64, FILEB_BB As TBit64, FILEC_BB As TBit64, FILED_BB As TBit64, FILEE_BB As TBit64, FILEF_BB As TBit64, FILEG_BB As TBit64, FILEH_BB As TBit64
Public RANK1_BB As TBit64, RANK2_BB As TBit64, RANK3_BB As TBit64, RANK4_BB As TBit64, RANK5_BB As TBit64, RANK6_BB As TBit64, RANK7_BB As TBit64, RANK8_BB As TBit64

Public Bit32Pos(31) As Long
Public Pop16Cnt(MIN_INTEGER To MAX_INTEGER) As Byte ' Max -int to +int
Public Int16x4 As TInt16x4
Public Int16x2 As TInt16x2
Public Byte8x8 As TByte8x8

Public Bit8Pos(7) As Integer

Public PiecesBB(WCOL, PT_ALL_PIECES) As TBit64, AllPiecesBB As TBit64, PiecesByPtBB(PT_ALL_PIECES) As TBit64, ColBB(WCOL) As TBit64, AttackedByBB(WCOL, PT_ALL_PIECES) As TBit64, AttackedBy2BB(WCOL) As TBit64
Public SquareBB(MAX_BOARD) As TBit64
Public EmptyBB As TBit64
Public FileBB(8) As TBit64
Public RankBB(8) As TBit64
Public AdjacentFilesBB(8) As TBit64
Public ForwardRanksBB(WCOL, 8) As TBit64
Public ForwardFileBB(WCOL, MAX_BOARD) As TBit64
Public PawnAttackSpanBB(WCOL, MAX_BOARD) As TBit64
Public PawnAttackSpanAllBB(WCOL) As TBit64
Public PassedPawnMaskBB(WCOL, MAX_BOARD) As TBit64
Public OutpostRanksBB(WCOL) As TBit64
Public SqToBit(MAX_BOARD) As Long
Public BitToSq(63) As Long
Public PawnAttacksFromSqBB(WCOL, SQ_H8) As TBit64
Public PseudoAttacksFromSqBB(PIECE_TYPE_NB, MAX_BOARD) As TBit64
Public AttackFromToBB(MAX_BOARD, MAX_BOARD) As TBit64
Public BetweenBB(MAX_BOARD, MAX_BOARD) As TBit64
Public KingRingBB(WCOL) As TBit64
Public LowRanksBB(WCOL) As TBit64
Public CampBB(WCOL) As TBit64
Public CenterBB As TBit64
Public CenterFilesBB As TBit64
Public DarkSquaresBB As TBit64
Public LSB16(MIN_INTEGER To MAX_INTEGER) As Integer
Public RSB16(MIN_INTEGER To MAX_INTEGER) As Integer

Public BB0 As TBit64, BB1 As TBit64

'--------------------------------------------------------------------


Public Sub Init32BitBoards()
  Dim i As Long, j As Long, k As Long, SqBB As Long
  
  For i = 0 To 7: Bit8Pos(i) = 2 ^ i: Next
  For i = 0 To 31: Bit32Pos(i) = BitMask32(i): Next
  
  For j = MIN_INTEGER To MAX_INTEGER
    Pop16Cnt(j) = Pop16CountFkt(j)
    LSB16(j) = -1
    For i = 0 To 15
       If CBool(j And Bit32Pos(i)) Then LSB16(j) = i: Exit For
    Next
    RSB16(j) = -1
    For i = 15 To 0 Step -1
       If CBool(j And Bit32Pos(i)) Then RSB16(j) = i: Exit For
    Next
  Next
  
  SqBB = 0
  For i = 0 To 119
    SqToBit(i) = -1
    If Board(i) <> FRAME Then
        SqToBit(i) = SqBB
        BitToSq(SqBB) = i
        
        '--- set ranks
        Select Case Rank(i)
        Case 1: SetBit64 RANK1_BB, SqBB
        Case 2: SetBit64 RANK2_BB, SqBB
        Case 3: SetBit64 RANK3_BB, SqBB
        Case 4: SetBit64 RANK4_BB, SqBB
        Case 5: SetBit64 RANK5_BB, SqBB
        Case 6: SetBit64 RANK6_BB, SqBB
        Case 7: SetBit64 RANK7_BB, SqBB
        Case 8: SetBit64 RANK8_BB, SqBB
        End Select

        '--- set Files
        Select Case File(i)
        Case 1: SetBit64 FILEA_BB, SqBB
        Case 2: SetBit64 FILEB_BB, SqBB
        Case 3: SetBit64 FILEC_BB, SqBB
        Case 4: SetBit64 FILED_BB, SqBB
        Case 5: SetBit64 FILEE_BB, SqBB
        Case 6: SetBit64 FILEF_BB, SqBB
        Case 7: SetBit64 FILEG_BB, SqBB
        Case 8: SetBit64 FILEH_BB, SqBB
        End Select
        
      SetBit64 SquareBB(i), SqBB
      If ColorSq(i) = BCOL Then SetBit64 DarkSquaresBB, SqBB
      '
      SqBB = SqBB + 1
    End If
  Next i

  FileBB(FILE_A) = FILEA_BB
  FileBB(FILE_B) = FILEB_BB
  FileBB(FILE_C) = FILEC_BB
  FileBB(FILE_D) = FILED_BB
  FileBB(FILE_E) = FILEE_BB
  FileBB(FILE_F) = FILEF_BB
  FileBB(FILE_G) = FILEG_BB
  FileBB(FILE_H) = FILEH_BB
  
  RankBB(1) = RANK1_BB
  RankBB(2) = RANK2_BB
  RankBB(3) = RANK3_BB
  RankBB(4) = RANK4_BB
  RankBB(5) = RANK5_BB
  RankBB(6) = RANK6_BB
  RankBB(7) = RANK7_BB
  RankBB(8) = RANK8_BB
  
  OR64 LowRanksBB(WCOL), RANK2_BB, RANK3_BB
  OR64 LowRanksBB(BCOL), RANK6_BB, RANK7_BB
  OR64 CenterFilesBB, FILED_BB, FILEE_BB
  OR64 BB0, RANK4_BB, RANK5_BB
  AND64 CenterBB, CenterFilesBB, BB0
  OR64 OutpostRanksBB(WCOL), RANK4_BB, RANK5_BB: OR64 OutpostRanksBB(WCOL), OutpostRanksBB(WCOL), RANK6_BB
  OR64 OutpostRanksBB(BCOL), RANK3_BB, RANK4_BB: OR64 OutpostRanksBB(BCOL), OutpostRanksBB(BCOL), RANK5_BB
  
  AdjacentFilesBB(FILE_A) = FILEB_BB
  OR64 AdjacentFilesBB(FILE_B), FILEA_BB, FILEC_BB
  OR64 AdjacentFilesBB(FILE_C), FILEA_BB, FILED_BB
  OR64 AdjacentFilesBB(FILE_D), FILEA_BB, FILEE_BB
  OR64 AdjacentFilesBB(FILE_E), FILEA_BB, FILEF_BB
  OR64 AdjacentFilesBB(FILE_F), FILEA_BB, FILEG_BB
  OR64 AdjacentFilesBB(FILE_G), FILEA_BB, FILEH_BB
  AdjacentFilesBB(FILE_H) = FILEG_BB
  
  ForwardRanksBB(WCOL, 7) = RANK8_BB
  OR64 ForwardRanksBB(WCOL, 6), ForwardRanksBB(WCOL, 7), RANK7_BB
  OR64 ForwardRanksBB(WCOL, 5), ForwardRanksBB(WCOL, 6), RANK6_BB
  OR64 ForwardRanksBB(WCOL, 4), ForwardRanksBB(WCOL, 5), RANK5_BB
  OR64 ForwardRanksBB(WCOL, 3), ForwardRanksBB(WCOL, 4), RANK4_BB
  OR64 ForwardRanksBB(WCOL, 2), ForwardRanksBB(WCOL, 3), RANK3_BB
  OR64 ForwardRanksBB(WCOL, 1), ForwardRanksBB(WCOL, 2), RANK2_BB
  
  ForwardRanksBB(BCOL, 2) = RANK1_BB
  OR64 ForwardRanksBB(BCOL, 3), ForwardRanksBB(BCOL, 2), RANK2_BB
  OR64 ForwardRanksBB(BCOL, 4), ForwardRanksBB(BCOL, 3), RANK3_BB
  OR64 ForwardRanksBB(BCOL, 5), ForwardRanksBB(BCOL, 4), RANK4_BB
  OR64 ForwardRanksBB(BCOL, 6), ForwardRanksBB(BCOL, 5), RANK5_BB
  OR64 ForwardRanksBB(BCOL, 7), ForwardRanksBB(BCOL, 6), RANK6_BB
  OR64 ForwardRanksBB(BCOL, 8), ForwardRanksBB(BCOL, 7), RANK7_BB
  
  CampBB(WCOL) = ForwardRanksBB(BCOL, 6)
  CampBB(BCOL) = ForwardRanksBB(WCOL, 3)
  
  ' Init SqFrom attacks
    Dim d As Long, Col As Long, Offset As Long

 For Col = BCOL To WCOL
  For i = SQ_A1 To SQ_H8
    If Board(i) <> FRAME Then
        ' Pawn attacks
        If Col = WCOL Then j = i + 9 Else j = i - 9
        If Board(j) <> FRAME Then
           SetBit64 PawnAttacksFromSqBB(Col, i), SqToBit(j)
        End If
        If Col = WCOL Then j = i + 11 Else j = i - 11
        If Board(j) <> FRAME Then
           SetBit64 PawnAttacksFromSqBB(Col, i), SqToBit(j)
        End If
        
        AND64 ForwardFileBB(Col, i), ForwardRanksBB(Col, Rank(i)), FileBB(File(i))
        AND64 PawnAttackSpanBB(Col, i), ForwardRanksBB(Col, Rank(i)), AdjacentFilesBB(File(i))
        OR64 PassedPawnMaskBB(Col, i), ForwardFileBB(Col, i), PawnAttackSpanBB(Col, i)
        
        If Col = WCOL Then ' same for black
          ' King/Knight attacks
          For d = 0 To 7
            Offset = QueenOffsets(d)
            j = i + Offset
            If Board(j) <> FRAME Then
              SetBit64 PseudoAttacksFromSqBB(PT_KING, i), SqToBit(j)  ' King
              
              '
              Do While Board(j) <> FRAME
                SetBit64 PseudoAttacksFromSqBB(PT_QUEEN, i), SqToBit(j)  ' Queen
                
                If Board(j) <> FRAME Then
                  If j <> i + Offset Then
                    SetBit64 BetweenBB(i, j), SqToBit(j - Offset)   ' between 2 squares, current square
                    SetOR64 BetweenBB(i, j), BetweenBB(i, j - Offset)  ' previous squares in line
                  End If
                  AttackFromToBB(i, j) = BetweenBB(i, j): SetBit64 AttackFromToBB(i, j), SqToBit(j) ' includes target square
                End If
                
                If d < 4 Then
                  SetBit64 PseudoAttacksFromSqBB(PT_ROOK, i), SqToBit(j)  ' Rook
                Else
                  SetBit64 PseudoAttacksFromSqBB(PT_BISHOP, i), SqToBit(j)  ' Bishop
                End If
                j = j + Offset
              Loop
            End If
            
            '---
            j = i + KnightOffsets(d)
            If Board(j) <> FRAME Then
              SetBit64 PseudoAttacksFromSqBB(PT_KNIGHT, i), SqToBit(j) ' Knight
            End If
          Next d
        End If
     End If
   Next i
 Next Col
End Sub

Function BitMask32(ByVal BitPos As Long) As Long ' 32 bit
  'If BitPos < 0 Or BitPos > 31 Then Err.Raise 6 ' overflow
  If BitPos < 31 Then
   BitMask32 = 2 ^ BitPos
  Else
   BitMask32 = BitL_31
  End If
End Function

Public Function Pop16CountFkt(ByVal x As Long) As Long
  ' for positive values only
  Pop16CountFkt = 0: If x = 0 Then Exit Function
  If x < 0 Then Pop16CountFkt = Pop16CountFkt + 1: x = x And Not &H8000
  Do While x > 0
    Pop16CountFkt = Pop16CountFkt + 1: x = x And (x - 1)
  Loop
End Function

Public Sub AND64(Result As TBit64, Op1 As TBit64, Op2 As TBit64)
  Result.i0 = Op1.i0 And Op2.i0: Result.i1 = Op1.i1 And Op2.i1
End Sub

Public Sub SetAND64(Op1 As TBit64, Op2 As TBit64) ' returns Op1
  Op1.i0 = Op1.i0 And Op2.i0: Op1.i1 = Op1.i1 And Op2.i1
End Sub

Public Sub SetANDNOT64(Op1 As TBit64, Op2 As TBit64) ' returns Op1
  Op1.i0 = Op1.i0 And Not Op2.i0: Op1.i1 = Op1.i1 And Not Op2.i1
End Sub

Public Sub OR64(Result As TBit64, Op1 As TBit64, Op2 As TBit64)
  Result.i0 = Op1.i0 Or Op2.i0: Result.i1 = Op1.i1 Or Op2.i1
End Sub

Public Sub SetOR64(Op1 As TBit64, Op2 As TBit64) ' returns Op1
  Op1.i0 = Op1.i0 Or Op2.i0: Op1.i1 = Op1.i1 Or Op2.i1
End Sub

Public Sub XOr64(Result As TBit64, Op1 As TBit64, Op2 As TBit64)
  Result.i0 = Op1.i0 Xor Op2.i0: Result.i1 = Op1.i1 Xor Op2.i1
End Sub

Public Sub ANDNOT64(Result As TBit64, Op1 As TBit64, Op2 As TBit64)
  Result.i0 = Op1.i0 And Not Op2.i0: Result.i1 = Op1.i1 And Not Op2.i1
End Sub

Public Sub SetNOT64(Result As TBit64, Op1 As TBit64)
  Result.i0 = Not Op1.i0: Result.i1 = Not Op1.i1
End Sub

Public Sub Set64(Result As TBit64, Op1 As TBit64)
  Result.i0 = Op1.i0: Result.i1 = Op1.i1  ' much faster then  Result=Op1 !!!!
End Sub

Public Function EQUAL64(Op1 As TBit64, Op2 As TBit64) As Boolean
  If Op1.i0 = Op2.i0 Then
    If Op1.i1 = Op2.i1 Then EQUAL64 = True Else EQUAL64 = False
  Else
    EQUAL64 = False
  End If
End Function

Public Sub Clear64(Op1 As TBit64)
  Op1.i0 = 0: Op1.i1 = 0
End Sub

Public Function ShiftDown64(Op1 As TBit64) As TBit64
   ' shift right 8 bits
'  LSet Byte8x8 = Op1
'  Byte8x8.i0 = Byte8x8.i1
'  Byte8x8.i1 = Byte8x8.i2
'  Byte8x8.i2 = Byte8x8.i3
'  Byte8x8.i3 = Byte8x8.i4
'  Byte8x8.i4 = Byte8x8.i5
'  Byte8x8.i5 = Byte8x8.i6
'  Byte8x8.i6 = Byte8x8.i7
'  Byte8x8.i7 = 0
'  LSet ShiftDown64 = Byte8x8
 
   ' i1
   If Op1.i1 And BitL_31 Then
     ShiftDown64.i1 = (((Op1.i1 And Not BitL_31) \ &H100&) Or BitL_23) And Not RANK4_L   ' shift 8 bits down (=&H100& ),remove rank 4 and add sign bit 31 as bit 23
   Else
     ShiftDown64.i1 = (Op1.i1 \ &H100&) And Not RANK4_L
   End If

   ' Copy RANK5 to RANK4 > copy to i0 bits 0-6 (= &H7F&) of rank1 and shift 24 bits (=&H1000000) up
   If Op1.i1 And BitL_7 Then
     ShiftDown64.i0 = ((Op1.i1 And &H7F&) * &H1000000) Or BitL_31 ' copy bit 7 from I1 to I0 sign bit
   Else
     ShiftDown64.i0 = (Op1.i1 And &H7F&) * &H1000000
   End If
   
   ' i0
   If Op1.i0 And BitL_31 Then
     ShiftDown64.i0 = ShiftDown64.i0 Or ((((Op1.i0 And Not BitL_31) \ &H100&) Or BitL_23) And Not RANK4_L)  ' shift 8 bits down (=&H100& ),remove rank 4 and add sign bit 31 as bit 23
   Else
     ShiftDown64.i0 = ShiftDown64.i0 Or ((Op1.i0 \ &H100&) And Not RANK4_L)
   End If
   'ShowLBB ShiftDown64
End Function
      
Public Function ShiftUp64(Op1 As TBit64) As TBit64
  ' shift left 8 bits
'  LSet Byte8x8 = Op1
'  Byte8x8.i7 = Byte8x8.i6
'  Byte8x8.i6 = Byte8x8.i5
'  Byte8x8.i5 = Byte8x8.i4
'  Byte8x8.i4 = Byte8x8.i3
'  Byte8x8.i3 = Byte8x8.i2
'  Byte8x8.i2 = Byte8x8.i1
'  Byte8x8.i1 = Byte8x8.i0
'  Byte8x8.i0 = 0
'  LSet ShiftUp64 = Byte8x8
  
  ' i0
  If Op1.i0 And BitL_23 Then
    ShiftUp64.i0 = ((Op1.i0 And Not RANK4_L And Not BitL_23) * &H100& Or BitL_31)
  Else
    ShiftUp64.i0 = (Op1.i0 And Not RANK4_L) * &H100&
  End If
  ' i1
  If Op1.i1 And BitL_23 Then
    ShiftUp64.i1 = ((Op1.i1 And Not RANK4_L And Not BitL_23) * &H100&) Or BitL_31
  Else
    ShiftUp64.i1 = (Op1.i1 And Not RANK4_L) * &H100&
  End If
  ' Copy RANK5 to RANK4 > copy to i1 bits 24-30 (= &H7F000000) of rank4  and shift 24 bits (=&H1000000) down
  If Op1.i0 And BitL_31 Then
    ShiftUp64.i1 = ShiftUp64.i1 Or ((Op1.i0 And &H7F000000) \ &H1000000) Or BitL_7 ' copy sign bit 31 from I0 to I1 bit 7
  Else
    ShiftUp64.i1 = ShiftUp64.i1 Or ((Op1.i0 And &H7F000000) \ &H1000000)
  End If
End Function

Public Function ShiftLeft64(Op1 As TBit64) As TBit64
  ShiftLeft64.i0 = Op1.i0 And Not FILEA_BB.i0: ShiftLeft64.i1 = Op1.i1 And Not FILEA_BB.i1 ' remove file A
  ShiftLeft64.i0 = ShiftLeft64.i0 \ &H2& And Not BitL_31
  ShiftLeft64.i1 = ShiftLeft64.i1 \ &H2& And Not BitL_31
End Function

Public Function ShiftRight64(Op1 As TBit64) As TBit64
  ShiftRight64.i0 = Op1.i0 And Not FILEH_BB.i0: ShiftRight64.i1 = Op1.i1 And Not FILEH_BB.i1 ' remove file H
  If ShiftRight64.i0 And BitL_30 Then
    ShiftRight64.i0 = ((ShiftRight64.i0 And Not BitL_30) * &H2&) Or BitL_31 ' move bit30 to bit31 else overflow
  Else
    ShiftRight64.i0 = (ShiftRight64.i0 And &HFFFFFFFF) * &H2&
  End If
  If ShiftRight64.i1 And BitL_30 Then
    ShiftRight64.i1 = ((ShiftRight64.i1 And Not BitL_30) * &H2&) Or BitL_31
  Else
    ShiftRight64.i1 = (ShiftRight64.i1 And &HFFFFFFFF) * &H2&
  End If
End Function

Public Function ShiftUpOrDown64(ByVal UpDown As Long, Op1 As TBit64) As TBit64
  If UpDown = SQ_UP Then
    ShiftUpOrDown64 = ShiftUp64(Op1)
  ElseIf UpDown = SQ_DOWN Then
    ShiftUpOrDown64 = ShiftDown64(Op1)
  End If
End Function


Public Sub SetBit64(Op1 As TBit64, ByVal BitPos As Long)
  ' bitPos 0 to 63
  Debug.Assert BitPos >= 0 And BitPos < 64
  'If BitPos < 0 Then Exit Sub
  'If BitPos > 63 Then Exit Sub
    
  If BitPos < 32 Then
    If BitPos = 31 Then Op1.i0 = Op1.i0 Or BitL_31 Else Op1.i0 = Op1.i0 Or Bit32Pos(BitPos)
  Else
    BitPos = BitPos - 32
    If BitPos = 31 Then Op1.i1 = Op1.i1 Or BitL_31 Else Op1.i1 = Op1.i1 Or Bit32Pos(BitPos)
  End If
End Sub

Public Function IsSetBit64(Op1 As TBit64, ByVal BitPos As Long) As Boolean
  ' bitPos 0 to 63
  Debug.Assert BitPos >= 0 And BitPos < 64
  'If BitPos < 0 Then Exit Sub
  'If BitPos > 63 Then Exit Sub
      
  If BitPos < 32 Then
    IsSetBit64 = CBool(Op1.i0 And Bit32Pos(BitPos))
  Else
    IsSetBit64 = CBool(Op1.i1 And Bit32Pos(BitPos - 32))
  End If
End Function

Public Function IsSet64(Op1 As TBit64) As Boolean
  If Op1.i0 <> 0 Then IsSet64 = True: Exit Function
  If Op1.i1 <> 0 Then IsSet64 = True: Exit Function
  IsSet64 = False
End Function


Public Sub ClearBit64(Op1 As TBit64, ByVal BitPos As Long)
  ' bitPos 0 to 63
  Debug.Assert BitPos < 64
  If BitPos < 32 Then
    Op1.i0 = Op1.i0 And Not Bit32Pos(BitPos)
  Else
    Op1.i1 = Op1.i1 And Not Bit32Pos(BitPos - 32)
  End If
End Sub


Public Function PopCnt64(Op1 As TBit64) As Long
  LSet Int16x4 = Op1
  PopCnt64 = Pop16Cnt(Int16x4.i0) + Pop16Cnt(Int16x4.i1) + Pop16Cnt(Int16x4.i2) + Pop16Cnt(Int16x4.i3)
End Function

Public Sub ShowBB(bb As TBit64)
 Dim i As Long, s As String
 Debug.Print
 Debug.Print " ------------------"
 s = ""
 For i = 63 To 0 Step -1
   If (i + 1) Mod 8 = 0 And s <> "" Then
     Debug.Print CStr((i + 9) \ 8) & "|" & s & "|"
     s = ""
   End If
   If IsSetBit64(bb, ByVal i) Then s = " X" & s Else s = " ." & s
 Next
 Debug.Print CStr((i + 9) \ 8) & "|" & s & "|"
 Debug.Print " ------------------"
 Debug.Print "   A B C D E F G H"
 Debug.Print
 
End Sub

Public Function Lsb64(Op1 As TBit64) As Long
 ' returns position of first bit set
 Lsb64 = -1
 If Op1.i0 <> 0 Then
   LSet Int16x4 = Op1
   Lsb64 = LSB16(Int16x4.i0): If Lsb64 >= 0 Then Exit Function
   Lsb64 = LSB16(Int16x4.i1): If Lsb64 >= 0 Then Lsb64 = Lsb64 + 16: Exit Function
 ElseIf Op1.i1 <> 0 Then
   LSet Int16x4 = Op1
   Lsb64 = LSB16(Int16x4.i2): If Lsb64 >= 0 Then Lsb64 = Lsb64 + 32: Exit Function
   Lsb64 = LSB16(Int16x4.i3): If Lsb64 >= 0 Then Lsb64 = Lsb64 + 48
 End If
End Function


Public Function Rsb64(Op1 As TBit64) As Long
 ' returns position of last bit set
 Rsb64 = -1
 If Op1.i1 <> 0 Then
   LSet Int16x4 = Op1
   Rsb64 = RSB16(Int16x4.i3): If Rsb64 >= 0 Then Rsb64 = Rsb64 + 48: Exit Function
   Rsb64 = RSB16(Int16x4.i2): If Rsb64 >= 0 Then Rsb64 = Rsb64 + 32: Exit Function
 ElseIf Op1.i0 <> 0 Then
   LSet Int16x4 = Op1
   Rsb64 = RSB16(Int16x4.i1): If Rsb64 >= 0 Then Rsb64 = Rsb64 + 16: Exit Function
   Rsb64 = RSB16(Int16x4.i0)
 End If
End Function

Public Function PopLsb64(Op1 As TBit64) As Long
  PopLsb64 = Lsb64(Op1)
  If PopLsb64 >= 0 Then ClearBit64 Op1, PopLsb64
End Function


Public Function PawnAttacksBB(ByRef Col As enumColor, Op1 As TBit64) As TBit64
  If Col = WCOL Then
    PawnAttacksBB = ShiftUp64(Op1)
    OR64 PawnAttacksBB, ShiftLeft64(PawnAttacksBB), ShiftRight64(PawnAttacksBB)
  ElseIf Col = BCOL Then
    PawnAttacksBB = ShiftDown64(Op1)
    OR64 PawnAttacksBB, ShiftLeft64(PawnAttacksBB), ShiftRight64(PawnAttacksBB)
  End If
End Function

Public Function AttacksBoardBB(ByVal pt As enumPieceType, ByVal sq As Long) As TBit64
      Dim LastTarget As Long, Target As Long, Offset As Long, d As Long, DirStart As Long, DirEnd As Long
      AttacksBoardBB = EmptyBB
      
      Select Case pt
      Case PT_ROOK: DirStart = 0: DirEnd = 3
      Case PT_BISHOP: DirStart = 4: DirEnd = 7
      Case PT_QUEEN: DirStart = 0: DirEnd = 7
      Case Else
        Exit Function
      End Select

      For d = DirStart To DirEnd
        Offset = QueenOffsets(d): Target = sq + Offset: LastTarget = sq
        Do While Board(Target) <> FRAME
          LastTarget = Target
          If Board(Target) >= NO_PIECE Then Exit Do
          Target = Target + Offset
        Loop
        If sq <> LastTarget Then
          If MaxDistance(sq, LastTarget) = 1 Then
            SetBit64 AttacksBoardBB, SqToBit(LastTarget)
          Else
            ' --- add bitboards for direction
            AttacksBoardBB.i0 = AttacksBoardBB.i0 Or AttackFromToBB(sq, LastTarget).i0: AttacksBoardBB.i1 = AttacksBoardBB.i1 Or AttackFromToBB(sq, LastTarget).i1
          End If
        End If
      Next
End Function

Public Function AttacksBB(ByVal pt As enumPieceType, ByVal sq As Long, occupied As TBit64) As TBit64
      Dim LastTarget As Long, Target As Long, Offset As Long, d As Long, DirStart As Long, DirEnd As Long
      
      AttacksBB = EmptyBB
      
      Select Case pt
      Case PT_ROOK: DirStart = 0: DirEnd = 3
      Case PT_BISHOP: DirStart = 4: DirEnd = 7
      Case PT_QUEEN: DirStart = 0: DirEnd = 7
      Case Else
        Exit Function
      End Select
      
      For d = DirStart To DirEnd
        Offset = QueenOffsets(d): Target = sq + Offset: LastTarget = sq

        Do While Board(Target) <> FRAME
          LastTarget = Target: If IsSetBit64(occupied, SqToBit(Target)) Then Exit Do
          Target = Target + Offset
        Loop
        If sq <> LastTarget Then SetOR64 AttacksBB, AttackFromToBB(sq, LastTarget) ' --- add bitboards for direction
      Next
End Function

Public Function MoreThanOne(op1BB As TBit64) As Boolean
  MoreThanOne = CBool(PopCnt64(op1BB) > 1)
End Function

Public Function FrontMostSq(Us As enumColor, Op1 As TBit64) As Long
 ' returns first square board position relative for color
 If Us = WCOL Then FrontMostSq = Rsb64(Op1) Else FrontMostSq = Lsb64(Op1)
 If FrontMostSq >= 0 Then FrontMostSq = BitToSq(FrontMostSq) Else FrontMostSq = 0
End Function


Public Function BackMostSq(Us As enumColor, Op1 As TBit64) As Long
 ' returns first square board position relative for color
 If Us = WCOL Then BackMostSq = Lsb64(Op1) Else BackMostSq = Rsb64(Op1)
 If BackMostSq >= 0 Then BackMostSq = BitToSq(BackMostSq) Else BackMostSq = 0
End Function

Public Sub Or64To(Op1 As TBit64, Op2 As TBit64, Result As TBit64)
  Result.i0 = Op1.i0 Or Op2.i0: Result.i1 = Op1.i1 Or Op2.i1
End Sub


'--------------------------------------------------------
Public Sub SetMove(m1 As TMOVE, m2 As TMOVE)
 With m1
  .Captured = m2.Captured
  .CapturedNumber = m2.CapturedNumber
  .Castle = m2.Castle
  .EnPassant = m2.EnPassant
  .From = m2.From
  .IsChecking = m2.IsChecking
  .IsLegal = m2.IsLegal
  .OrderValue = m2.OrderValue
  .Piece = m2.Piece
  .Promoted = m2.Promoted
  .SeeValue = m2.SeeValue
  .Target = m2.Target
 End With
End Sub

Public Sub SwapMove(m1 As TMOVE, m2 As TMOVE)
 Dim t As TMOVE
 With t
  .Captured = m2.Captured: m2.Captured = m1.Captured: m1.Captured = .Captured
  .CapturedNumber = m2.CapturedNumber: m2.CapturedNumber = m1.CapturedNumber: m1.CapturedNumber = .CapturedNumber
  .Castle = m2.Castle: m2.Castle = m1.Castle: m1.Castle = .Castle
  .EnPassant = m2.EnPassant: m2.EnPassant = m1.EnPassant: m1.EnPassant = .EnPassant
  .From = m2.From: m2.From = m1.From: m1.From = .From
  .IsChecking = m2.IsChecking: m2.IsChecking = m1.IsChecking: m1.IsChecking = .IsChecking
  .IsLegal = m2.IsLegal: m2.IsLegal = m1.IsLegal: m1.IsLegal = .IsLegal
  .OrderValue = m2.OrderValue: m2.OrderValue = m1.OrderValue: m1.OrderValue = .OrderValue
  .Piece = m2.Piece: m2.Piece = m1.Piece: m1.Piece = .Piece
  .Promoted = m2.Promoted: m2.Promoted = m1.Promoted: m1.Promoted = .Promoted
  .SeeValue = m2.SeeValue: m2.SeeValue = m1.SeeValue: m1.SeeValue = .SeeValue
  .Target = m2.Target: m2.Target = m1.Target: m1.Target = .Target
 End With
 
End Sub

Public Sub ClearMove(m1 As TMOVE)
  With m1
    .From = 0: .Target = 0: .Piece = NO_PIECE: .Castle = NO_CASTLE: .Promoted = 0: .Captured = NO_PIECE: .CapturedNumber = 0
    .EnPassant = 0: .IsChecking = False: .IsLegal = False: .OrderValue = 0: .SeeValue = UNKNOWN_SCORE
  End With
End Sub


Public Function Test64()
 Dim b As TBit64, bb As TBit64, i As Long, t As TBit64, x As Long
 InitEngine
 
 Init32BitBoards
 
'----
Dim StartTime As Single, EndTime As Single, y As Long, z As Long, sq As Long, j As Long
Dim m1 As TMOVE, m2 As TMOVE, m3 As TMOVE


StartTime = Timer

b = EmptyBB: x = Len(m1)
t.i0 = 123
m1.From = 2: m2.From = 12: m3.Target = 34
For i = 1 To 50000000
         
   SetMove m1, EmptyMove ' 2x schneller
   SetMove m2, m3
   SetMove m3, m1
   
 '  m1 = EmptyMove
 '  m2 = m3
 '  m3 = m1

   '1. ---
  'x = 23 + i Mod 2
   'SetBit64 b, x
  'If x < 32& Then If x = 31 Then b.i0 = b.i0 Or BitL_31 Else b.i0 = b.i0 Or Bit32Pos(x) Else If x = 63 Then b.i1 = b.i1 Or BitL_31 Else b.i1 = b.i1 Or Bit32Pos(x - 32)
  
  'x = x + 20
  'SetBit64 b, x
  'If x < 32& Then If x = 31 Then b.i0 = b.i0 Or BitL_31 Else b.i0 = b.i0 Or Bit32Pos(x) Else If x = 63 Then b.i1 = b.i1 Or BitL_31 Else b.i1 = b.i1 Or Bit32Pos(x - 32)
  
  
  
  'For j = 1 To 7
  '  bb.i0 = AttackedByBB(1, j).i0: bb.i1 = AttackedByBB(1, j).i1
  '  b.i0 = bb.i0: b.i1 = bb.i1
    'b.i0 = bb.i0 Or t.i0: b.i1 = bb.i1 Or t.i1
    
    
    'bb = AttackedByBB(1, j)
    'Set64 b, bb
    'b = bb
    'OR64 b, bb, t
  'Next
  '------------------------------------------------
   'z = 31 + i Mod 2
  '2.---
  'AttackedByBB(1, PT_PAWN).i0 = AttackedByBB(1, PT_PAWN).i0 Or SquareBB(z).i0: AttackedByBB(1, PT_PAWN).i1 = AttackedByBB(1, PT_PAWN).i1 Or SquareBB(z).i1
  'SetOR64 AttackedByBB(1, PT_PAWN), SquareBB(z)
  
  '3.---
  'bb.i0 = b.i0 Or SquareBB(z).i0: bb.i1 = b.i1 Or SquareBB(z).i1
  ' bb = Or64(b, SquareBB(z)) ' 24,6
  'Or64To b, SquareBB(z), bb ' 3,7
  'Or64ToP b, SquareBB(z)    ' 2,6
  
  
Next

EndTime = Timer
Debug.Print Format$(EndTime - StartTime, "0.000")
Debug.Print y
MsgBox Format$(EndTime - StartTime, "0.000") & "         " & bb.i0 & bb.i1 & x & m1.Target & m2.Target & b.i1 & t.i1 & b.i0 & AttackedByBB(1, PT_PAWN).i0

End Function
