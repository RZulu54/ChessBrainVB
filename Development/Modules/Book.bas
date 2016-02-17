Attribute VB_Name = "BookBas"
'==================================================
'BookBas:
'
'opening book functions (unchanged from LarsenVB)
'==================================================
Option Explicit

Public Book()         As String   '4999
Public BookPly        As Integer
Public bUseBook       As Boolean
Public OpeningHistory As String

Private dicBookIndex  As Scripting.Dictionary
Private NumBookLines  As Long

Private BookWhite1()  As String
Private NumBookWhite1 As Integer

'---------------------------------------------------------------------------
'ChooseBookMove()
'---------------------------------------------------------------------------
Public Function ChooseBookMove() As TMove

  Dim i                As Long, j As Long, iRandom As Integer, iReplies As Integer
  Dim sPossibleMove    As String, sCoordMove As String, sPreviousMove As String, From As Integer, Target As Integer
  Dim iNumMoves        As Integer
  Dim BookReplies()    As TMove
  Dim BookCandidates() As String
  Ply = 0

  GenerateMoves Ply, False, iNumMoves
  If BookPly = 0 Then
    If NumBookWhite1 > 0 Then
       iRandom = RndInt(0, 100)
       If iRandom < 43 Then  ' 43% for e2-e4
         From = SQ_E2: Target = SQ_E4
       ElseIf iRandom < 43 + 37 Then ' 37% for d2-d4
         From = SQ_D2: Target = SQ_D4
       ElseIf iRandom < 43 + 37 + 10 Then ' g1-f3
         From = SQ_G1: Target = SQ_F3
       ElseIf iRandom < 43 + 37 + 10 + 8 Then ' c2-c4
         From = SQ_C2: Target = SQ_C4
       Else  ' g2-g3
         From = SQ_G2: Target = SQ_G3
       End If
       For i = 0 To iNumMoves - 1
         If Moves(Ply, i).From = From And Moves(Ply, i).Target = Target Then
           ChooseBookMove = Moves(Ply, i)
           Exit Function
         End If
       Next

'      iRandom = RndInt(0, NumBookWhite1 - 1)
'      For i = 0 To iNumMoves - 1
'        sCoordMove = CompToCoord(Moves(Ply, i))
'        If BookWhite1(iRandom) = sCoordMove Then
'          RemoveEpPiece
'          MakeMove Moves(Ply, i)
'          If CheckLegal(Moves(Ply, i)) Then
'            ChooseBookMove = Moves(Ply, i)
'          End If
'          UnmakeMove Moves(Ply, i)
'          ResetEpPiece
'          Exit For
'        End If
'      Next
    End If
  Else
    BookCandidates = Filter(Book, OpeningHistory)
    For i = 0 To UBound(BookCandidates)
      sPossibleMove = Mid$(BookCandidates(i), 2 + (BookPly * 4), 4)
      If sPreviousMove <> sPossibleMove Then
        For j = 0 To iNumMoves - 1
          sCoordMove = CompToCoord(Moves(Ply, j))
          If sPossibleMove = sCoordMove Then
            RemoveEpPiece
            MakeMove Moves(Ply, j)
            If CheckLegal(Moves(Ply, j)) And CheckBookBit(BookCandidates(i), (Len(OpeningHistory) \ 4) + 1) Then
              ReDim Preserve BookReplies(iReplies)
              BookReplies(iReplies) = Moves(Ply, j)
              iReplies = iReplies + 1
            End If
            UnmakeMove Moves(Ply, j)
            ResetEpPiece
            Exit For
          End If
        Next
        sPreviousMove = sPossibleMove
      End If
    Next
    
    Select Case iReplies
      Case 0
        Exit Function
      Case 1
        ChooseBookMove = BookReplies(0)
      Case Else
        iRandom = RndInt(0, iReplies - 1)
        ChooseBookMove = BookReplies(iRandom)
    End Select
  End If

End Function

'---------------------------------------------------------------------------
'InitBook()
'---------------------------------------------------------------------------
Public Function InitBook() As Boolean

  Dim sBookName       As String, sIndexFile As String
  Dim iFBook          As Integer, iFIndex As Integer
  Dim sBookLine       As String
  Dim lBookIndex      As Long, lAllSet As Long
  Dim sLastWhite1Move As String

  sBookName = ReadINISetting(USE_BOOK_KEY, "")
  If sBookName <> "" Then
    NumBookLines = 0
    Erase Book
    Set dicBookIndex = New Scripting.Dictionary
    
    iFBook = FreeFile
    If InStr(1, sBookName, "\") = 0 Then
      sBookName = psEnginePath & "\" & sBookName
    End If
    On Local Error GoTo BookErr:
    Open sBookName For Input As iFBook
    
    sIndexFile = GetIndexFileName(sBookName)
    If Dir(sIndexFile) <> "" Then
      iFIndex = FreeFile
      Open sIndexFile For Binary Lock Write As iFIndex
    Else
      lAllSet = SetAllBits
    End If
    On Local Error GoTo 0

    Do While Not EOF(iFBook)
      Line Input #iFBook, sBookLine
        
      ReDim Preserve Book(NumBookLines)
      Book(NumBookLines) = sBookLine
      If iFIndex <> 0 Then
        Get iFIndex, , lBookIndex
      Else
        lBookIndex = lAllSet
      End If
      dicBookIndex.Add sBookLine, lBookIndex
    
      If Trim$(Left$(sBookLine, 5)) <> sLastWhite1Move Then
        sLastWhite1Move = Trim$(Left$(sBookLine, 5))
        If lBookIndex And 1 Then
          ReDim Preserve BookWhite1(NumBookWhite1)
          BookWhite1(NumBookWhite1) = sLastWhite1Move
          NumBookWhite1 = NumBookWhite1 + 1
        End If
      End If
        
      NumBookLines = NumBookLines + 1
    Loop
    Close iFBook
    If iFIndex <> 0 Then Close iFIndex

    InitBook = CBool(NumBookLines > 0)
  End If
  Exit Function

BookErr:
  LogWrite "Opening book error: " & Error, , True

End Function
'---------------------------------------------------------------------------
'SetAllBits()
'---------------------------------------------------------------------------
Private Function SetAllBits() As Long

  Dim i As Integer, lBit As Long

  For i = 0 To BOOK_MAX_PLY - 1
    lBit = lBit Or 2 ^ i
  Next
  SetAllBits = lBit

End Function
'---------------------------------------------------------------------------
'CheckBookBit()
'---------------------------------------------------------------------------
Private Function CheckBookBit(ByVal sKey As String, ByVal iPos As Integer) As Boolean

  Dim lIndex As Long, lBit As Long

  lBit = 2 ^ (iPos - 1)
  lIndex = dicBookIndex(sKey)

  CheckBookBit = lIndex And lBit

End Function

'---------------------------------------------------------------------------
'GetIndexFileName()
'---------------------------------------------------------------------------
Private Function GetIndexFileName(ByVal sBook As String) As String

  Dim iPos As Integer

  iPos = InStrRev(sBook, ".")

  If iPos <> 0 Then GetIndexFileName = Left$(sBook, iPos) & "opi"

End Function
