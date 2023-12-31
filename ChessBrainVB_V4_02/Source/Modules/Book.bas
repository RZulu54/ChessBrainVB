Attribute VB_Name = "basBook"
'==================================================
'= basBook:
'= chess opening book functions
'==================================================
Option Explicit
Public bUseBook       As Boolean

Public UCIBook() As String
Public UCIBookMax As Long, UCIBookCnt As Long
Public BookMovePossible As Boolean

'---------------------------------------------------------------------------
'ChooseBookMove()
'---------------------------------------------------------------------------
Public Function ChooseBookMove() As TMOVE
  ' game has to be started from startup position, FEN/EPD loaded position not supported
  Dim i                As Long
  Dim sPossibleMove    As String, sCoordMove As String
  Dim iNumMoves        As Long
  
  SetMove ChooseBookMove, EmptyMove
  
  sPossibleMove = GetUCIBookMove()
  
  ' check for legal move
  Ply = 1
  GenerateMoves Ply, False, iNumMoves

  For i = 0 To iNumMoves - 1
    sCoordMove = CompToCoord(Moves(Ply, i)) ' format "e4d5"
    If sCoordMove = sPossibleMove Then
      SetMove ChooseBookMove, Moves(Ply, i)
      Exit Function
    End If
  Next

End Function


'---------------------------------------------------------------------------
'InitBook()
'---------------------------------------------------------------------------
Public Function InitBook() As Boolean
 Static bInitBookDone As Boolean
 Static bUseBookOk As Boolean
 Dim sBookFile As String
 
 If bInitBookDone Then ' read only once
  InitBook = bUseBookOk
  Exit Function
 End If
 
 If pbMSExcelRunning Then ' set in SetVBAPathes
   InitBook = ReadExcelBook()
 End If
 If Not InitBook Then
   sBookFile = ReadINISetting(USE_BOOK_KEY, "CB_BOOK.TXT")
   If pbIsOfficeMode And Trim(sBookFile) = "" Then
     ' Always use default book if not set in INI file
     sBookFile = "CB_BOOK.TXT"
   End If
   InitBook = ReadUCIBook(sBookFile)
 End If
 bUseBookOk = InitBook
 bInitBookDone = True
End Function

'---------------------------------------------------------------------------
' MS Excel: read book from internal worksheet
'---------------------------------------------------------------------------
Public Function ReadExcelBook() As Boolean
  On Error GoTo lblError

   #If VBA_MODE = 1 Then
      ' read opening book lines from Excel worksheet CB_BOOK
      Dim Sheet As Object, lNum As Long, i As Long, sInp As String
      
      Set Sheet = ActiveWorkbook.Sheets("CB_BOOK")
      
      ReDim UCIBook(0)
      UCIBookMax = 0: UCIBookCnt = 0
      
      With Sheet
        lNum = .Cells(.Rows.Count, 1).End(xlUp).Row
        For i = 1 To lNum
          sInp = Trim$(.Cells(i, 1))
          If Left(sInp, 1) <> "#" And sInp <> "" Then ' # : comment line
            UCIBookCnt = UCIBookCnt + 1: If UCIBookCnt > UCIBookMax Then UCIBookMax = UCIBookMax + 1000: ReDim Preserve UCIBook(UCIBookMax)
            UCIBook(UCIBookCnt) = sInp
          End If
        Next i
      End With ' sheet
  
      ReadExcelBook = (UCIBookCnt > 0)
      If ReadExcelBook Then
        SendCommand "opening book found in Excel sheet CB_BOOK. Lines found:" & UCIBookCnt
      End If
      Exit Function
   #End If
lblError:
      ReadExcelBook = False
End Function


Public Function GetUCIGameLine() As String
  Dim i As Long, h As Long, s As String, MoveCnt As Long, Cnt As Long
  
  GetUCIGameLine = ""
  
  Cnt = GameMovesCnt
  If Cnt = 0 Then Exit Function
  s = "": MoveCnt = 0

  For i = 1 To Cnt Step 2
    MoveCnt = MoveCnt + 1
    s = s & CompToCoord(arGameMoves(i))
    If i + 1 <= Cnt Then s = s & " " & CompToCoord(arGameMoves(i + 1)) & " "
  Next i
  GetUCIGameLine = Trim$(s)
  
End Function

Public Function ReadUCIBook(isFile As String) As Boolean
 ' Read PGN File
  Dim h As Long, sInp As String, sBookFile As String

  ReadUCIBook = False
  
  h = 10 'FreeFile()
  ReDim UCIBook(0)
  UCIBookMax = 0: UCIBookCnt = 0
  
  sBookFile = psEnginePath & "\" & isFile
  
  On Error GoTo lblError
  If Dir(sBookFile) = "" Or isFile = "" Then
    Dim sDefault As String
    If pbIsOfficeMode Then sDefault = "1" Else sDefault = "0"
    
    If ReadINISetting("USE_INTERNAL_BOOK", sDefault) = "1" Then
      InitInternalBook
      ReadUCIBook = True
      If pbIsOfficeMode Then
        SendCommand "internal opening book active"
      ElseIf UCIMode Then
        SendCommand "info string internal opening book active"
      End If
    End If
    Exit Function
  End If
  
  Open sBookFile For Input As #h

  Do Until EOF(h)
    Line Input #h, sInp: sInp = Trim(sInp)
    If Left(sInp, 1) <> "#" And sInp <> "" Then ' # : comment line
      UCIBookCnt = UCIBookCnt + 1: If UCIBookCnt > UCIBookMax Then UCIBookMax = UCIBookMax + 1000: ReDim Preserve UCIBook(UCIBookMax)
      UCIBook(UCIBookCnt) = sInp
    End If
  Loop
  ReadUCIBook = (UCIBookCnt > 0)
  If ReadUCIBook Then
      If pbIsOfficeMode Then
        SendCommand "opening book found: " & isFile
      ElseIf UCIMode Then
        SendCommand "info string opening book found: " & isFile
      End If
  End If
  
  Close #h
  Exit Function
lblError:
  ReadUCIBook = False
End Function

Public Function GetUCIBookMove() As String
'--- input file ist sorted, lowercase, UCi format e4d5 (not e4xd5 or Bxd5)
' ---- create book file from PGN
' pgn-extract.exe -Wuci --notags --noresults -C -N -V --output book.txt test.pgn
' sort out.txt /o book.txt

Dim sUCIGame As String, sBookLine As String, r As Double
Dim i As Long, lStart As Long, lEnd As Long, lUCILen As Long, x As Long

GetUCIBookMove = ""
sUCIGame = GetUCIGameLine()

lUCILen = Len(sUCIGame)

If lUCILen >= 4 Then
  lStart = 0
  For i = 1 To UCIBookCnt
   If Left$(UCIBook(i), lUCILen) = sUCIGame Then
     lEnd = i: If lStart = 0 Then lStart = i
   End If
  Next
Else
  ' first game move
  lStart = 1: lEnd = UCIBookCnt
End If

' get a random move in the range found
Randomize
r = Rnd
If lEnd > lStart Then lStart = lStart + Int(((lEnd - lStart + 1) * r))
sBookLine = Trim$(Mid$(UCIBook(lStart), lUCILen + 1))
If Len(sBookLine) >= 4 Then
  sBookLine = Trim$(Left$(sBookLine, 4)) ' no promotion moves supported
  If Len(sBookLine) = 4 Then GetUCIBookMove = sBookLine
End If

'Debug.Print lStart, lEnd; r, GetUCIBookMove
End Function

Public Function InitInternalBook()
  ' Read internal book, just for fun - if external book is missing
  Dim BookArr As Variant ' extra array because Variant type needed for ARRAY()
  Dim i As Long
  BookArr = Array("a2a3 g7g6 g2g3 f8g7 f1g2", "a2a3 g8f6 g1f3 d7d5 d2d4", "b1c3 c7c5 d2d4 c5d4 d1d4", "b1c3 c7c5 e2e3 g7g6 d2d4", "b1c3 d7d5 e2e4 d5d4 c3e2", _
                  "b2b3 d7d5 c1b2 c8g4 g2g3", "b2b3 e7e5 c1b2 b8c6 e2e3", "c2c4 b8c6 g2g3 e7e5 f1g2", "c2c4 c7c5 b1c3 b7b6 e2e3", "c2c4 c7c5 b1c3 b7b6 e2e4", _
                  "c2c4 e7e5 b1c3 b8c6 g2g3", "c2c4 e7e6 g1f3 g8f6 b2b3", "c2c4 e7e6 g2g3 g8f6 f1g2", "c2c4 f7f5 b1c3 g8f6 d2d3", "c2c4 f7f5 b1c3 g8f6 d2d4", _
                  "c2c4 g8f6 b1c3 e7e6 g1f3", "d2d4 d7d5 c2c4 c7c6 b1c3", "d2d4 d7d5 g1f3 g8f6 c2c4", "d2d4 d7d5 g1f3 g8f6 e2e3", "d2d4 d7d6 c1g5 b8d7 e2e4", _
                  "d2d4 d7d6 c1g5 f7f6 g5h4", "d2d4 d7d6 c1g5 g7g6 c2c4", "d2d4 d7d6 c2c3 g8f6 c1g5", "d2d4 d7d6 c2c4 e7e5 b1c3", "d2d4 d7d6 c2c4 f7f5 g2g3", _
                  "d2d4 d7d6 c2c4 g7g6 b1c3", "d2d4 d7d6 e2e4 c7c5 d4d5", "d2d4 d7d6 e2e4 e7e5 g1f3", "d2d4 d7d6 e2e4 g7g6 b1c3", "d2d4 d7d6 e2e4 g7g6 c2c4", _
                  "d2d4 d7d6 e2e4 g8f6 b1c3", "d2d4 g8f6 c2c4 e7e6 g1f3", "d2d4 g8f6 g1f3 g7g6 g2g3", "e2e4 b7b6 g2g3 c8b7 f1g2", "e2e4 b8c6 b1c3 e7e5 f1c4", _
                  "e2e4 b8c6 d2d4 e7e5 g1f3", "e2e4 b8c6 f1b5 g8f6 d2d3", "e2e4 b8c6 g1f3 d7d6 d2d4", "e2e4 c7c5 b1c3 a7a6 g2g4", "e2e4 c7c5 b1c3 b8c6 d2d3", _
                  "e2e4 c7c5 c2c3 e7e6 d2d4", "e2e4 c7c5 f2f4 d7d5 d2d3", "e2e4 c7c5 f2f4 d7d5 e4d5", "e2e4 c7c5 g1f3 a7a6 b1c3", "e2e4 c7c5 g1f3 d8c7 d2d4", _
                  "e2e4 c7c5 g1f3 e7e6 b1c3", "e2e4 c7c6 g1f3 d7d5 e4d5", "e2e4 d7d6 d2d4 g8f6 b1c3", "f2f4 b7b6 g1f3 c8b7 e2e3", "f2f4 c7c5 b2b3 g8f6 c1b2", _
                  "f2f4 d7d5 e2e3 g8f6 g1f3", "g1f3 c7c5 c2c3 g8f6 g2g3", "g1f3 c7c5 c2c4 b7b6 b1c3", "g1f3 d7d5 c2c4 c7c6 g2g3", "g1f3 d7d5 c2c4 d5c4 b1a3", _
                  "g2g3 d7d5 f1g2 c7c6 g1f3", "g2g3 d7d5 f1g2 e7e5 c2c3", "g2g3 e7e5 f1g2 d7d5 d2d3", "g2g3 g7g6 f1g2 f8g7 c2c4", "g2g3 g8f6 f1g2 e7e5 d2d3")
  UCIBookCnt = UBound(BookArr) + 1
  ReDim UCIBook(UCIBookCnt)
  For i = 1 To UCIBookCnt: UCIBook(i) = BookArr(i - 1): Next
 
End Function




