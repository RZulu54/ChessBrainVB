VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHashMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
'==========
' HashMap =
'==========
'
'- class for sharing a read/write memory-mapped file
'- backed by the Windows paging file rather than a specific
'- disk file between processes running under the same User account
'- on the same system.
'
'- The Name values can optionally be prefixed "Global\" or "Local\"
'- (see the documentation) and the rest can consist of any
'- characters except the "\" character.
'
'- After we obtain a handle to the object, we'll create a single
'- "view" containing the entire object as one BLOB.
'
'- When all handles to the mapped object have been closed, it disappears.
'
Private Const API_NULL             As Long = 0
Private Const API_FALSE            As Long = 0
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const PAGE_READWRITE       As Long = 4
Private Const SECTION_MAP_WRITE = &H2
Private Const FILE_MAP_WRITE = SECTION_MAP_WRITE
Private Const ERROR_ALREADY_EXISTS As Long = 183

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByVal Destination As Long, _
                                       ByVal Source As Long, _
                                       ByVal Length As Long)
Private Declare Function RtlCompareMemory _
                Lib "ntdll" (ByRef Source1 As Any, _
                             ByRef Source2 As Any, _
                             ByVal Length As Long) As Long
                             
Private Declare Function CreateFileMapping _
                Lib "kernel32" _
                Alias "CreateFileMappingA" (ByVal hFile As Long, _
                                            ByVal lpFileMappigAttributes As Long, _
                                            ByVal flProtect As Long, _
                                            ByVal dwMaximumSizeHigh As Long, _
                                            ByVal dwMaximumSizeLow As Long, _
                                            ByVal lpName As String) As Long
Private Declare Function MapViewOfFile _
                Lib "kernel32" (ByVal hFileMappingObject As Long, _
                                ByVal dwDesiredAccess As Long, _
                                ByVal dwFileOffsetHigh As Long, _
                                ByVal dwFileOffsetLow As Long, _
                                ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function OpenFileMapping _
                Lib "kernel32" _
                Alias "OpenFileMappingA" (ByVal dwDesiredAccess As Long, _
                                          ByVal bInheritHandle As Long, _
                                          ByVal lpName As String) As Long
Private Declare Function UnmapViewOfFile _
                Lib "kernel32" (ByVal lpBaseAddress As Long) As Long
Private Declare Sub ZeroMemory2 _
                Lib "kernel32.dll" _
                Alias "RtlZeroMemory" (Destination As Any, _
                                       ByVal Length As Long)
Private hObj           As Long
Private lpMap          As Long
Private mSize          As Long
Private Ptr            As Long
Private CheckData(128) As Byte, CheckDataPtr As Long
Private VerifyArr(500) As Byte


Public Sub CloseMap()
  UnmapViewOfFile lpMap
  CloseHandle hObj
  hObj = 0
End Sub

Public Function CreateMap(ByVal Name As String, ByRef Size As Long) As Boolean
  'Returns True if the memory mapped file already exists.  If so CAUTION,
  'the size will be its previously-created size.
  ' Size may be reduced if not enough memory - and returned to caller !!!
  Dim i As Long, ErrCode As Long
  If Size < 1 Then Err.Raise 5, TypeName(Me), "Size must be at least 1 byte"
  For i = 1 To 10
    DoEvents
    Err.Clear
    hObj = CreateFileMapping(INVALID_HANDLE_VALUE, API_NULL, PAGE_READWRITE, 0, Size, Name)
    ErrCode = Err.LastDllError
    If hObj = API_NULL Then
      Err.Raise &H80049300, TypeName(Me), "CreateFileMapping system error " & CStr(Err.LastDllError)
    End If
    CreateMap = (Err.LastDllError = ERROR_ALREADY_EXISTS)
    
    DoEvents
    lpMap = MapViewOfFile(hObj, FILE_MAP_WRITE, 0, 0, 0)
    If lpMap = API_NULL Then
      Me.CloseMap
      If i = 10 Then
        Err.Raise &H80049302, TypeName(Me), "MapViewOfFile system error " & CStr(Err.LastDllError)
        Exit Function
      ElseIf i >= 3 Then ' try reducing size
        Size = (Size / 3) * 2
      End If
        
      WriteTrace "***error> CreateMap: " & i & " / " & Name & " / Size= " & Size & " / Err:" & CStr(Err.LastDllError) & " / " & Now()
     ' Err.Raise &H80049302, TypeName(Me), "MapViewOfFile system error " & CStr(Err.LastDllError)
      Sleep 100
    Else
      Exit For
    End If
  Next
  ZeroMemory2 ByVal lpMap, Size
  If bThreadTrace Then WriteTrace "---Creat map:ZERO_MAP / " & Now()
  mSize = Size
End Function

Public Sub ClearMap(ByRef Size As Long)
  ZeroMemory2 ByVal lpMap, Size
  If bThreadTrace Then WriteTrace "--- Clear_MAP / " & Now()
End Sub


Public Sub OpenMap(ByVal Name As String, ByVal Size As Long)
  Dim i As Long
  If Size < 1 Then Err.Raise 5, TypeName(Me), "Size must be at least 1 byte"
  DoEvents
  hObj = OpenFileMapping(FILE_MAP_WRITE, API_FALSE, Name)
  If hObj = API_NULL Then
    Err.Raise &H80049304, TypeName(Me), "OpenFileMapping system error " & CStr(Err.LastDllError)
  End If
  For i = 1 To 5
    DoEvents
    lpMap = MapViewOfFile(hObj, FILE_MAP_WRITE, 0, 0, Size)
    If lpMap = API_NULL Then
      Err.Raise &H80049306, TypeName(Me), "MapViewOfFile system error " & CStr(Err.LastDllError)
    Else
      Exit For
    End If
  Next
  mSize = Size
  CheckDataPtr = VarPtr(CheckData(0))
End Sub

Public Sub ReadMapHashCluster(ByVal Index As Long, _
                              ByVal lpData As Long, _
                              ByVal Size As Long)
  'Pass a pointer lpData and a length in bytes Size.
  Dim i As Long
  If Index * HashRecLen + Size > mSize Then Err.Raise 5, TypeName(Me), "Size must not exceed mapped size"
  If hObj = API_NULL Then Err.Raise &H80049308, TypeName(Me), "ReadMap: Map not open"
  
  Ptr = lpMap + HashMapSearchPtr + Index * HashRecLen
  'If bHashTrace Then WriteTrace "ReadMapHashCluster: " & Index & "/" & Index & "Ptr:" & Ptr & "/ Nodes:" & Nodes & " / " & Now()
  For i = 1 To 3 ' about 1 hash collision for 1.000.000.000 endgame nodes measured
    CopyMemory ByVal lpData, ByVal Ptr, ByVal Size
    If bHashVerify Then
      ' If bHashTrace Then WriteTrace "ReadMapHashCluster:Verify :" & VarPtr(VerifyArr(0)) & "Ptr:" & Ptr & Now()
      CopyMemory ByVal VarPtr(VerifyArr(0)), ByVal Ptr, ByVal Size
      ' If bHashTrace Then WriteTrace "ReadMapHashCluster: Compare " & VarPtr(VerifyArr(0)) & "Ptr:" & Ptr & Now()
      If RtlCompareMemory(ByVal Ptr, ByVal VarPtr(VerifyArr(0)), ByVal Size) <> Size Then
        ' Difference found => try again
        If bTraceHashCollision Then WriteTrace "HashMapDifference: Read " & Index & "/" & i & "/ Nodes:" & Nodes & " / " & Now()
      Else
        Exit For
      End If
    Else
      Exit For
    End If
  Next
  'If bHashTrace Then WriteTrace "ReadMapHashCluster:End "
End Sub

Public Sub WriteMapHashEntry(ByVal ReplaceIndex As Long, ByVal lpData As Long)
  'Pass a pointer lpData and a length in bytes Size.
  Dim i As Long
  If (ReplaceIndex + 1) * HashRecLen > mSize Then Err.Raise 5, TypeName(Me), "Size must not exceed mapped size"
  If hObj = API_NULL Then Err.Raise &H8004930A, TypeName(Me), "WriteMap: Map not open"
  Ptr = lpMap + HashMapSearchPtr + ReplaceIndex * HashRecLen

  For i = 1 To 3  ' about 1 hash collision for 1.000.000.000 endgame nodes measured
    CopyMemory ByVal Ptr, ByVal lpData, ByVal HashRecLen
    '--- Reread the written entry to verify that there was no parallel write from other thread that mixed up the data
    '--- Try max 3 times
    If RtlCompareMemory(ByVal Ptr, ByVal lpData, ByVal HashRecLen) <> HashRecLen Then
      ' Difference found => try again
      If bTraceHashCollision Then WriteTrace "HashMapDifference: Write " & ReplaceIndex & "/" & i & "/ Nodes:" & Nodes & " / " & Now()
    Else
      Exit For
    End If
  Next

End Sub

Public Sub WriteMapPos(ByVal StartPos As Long, ByVal lpData As Long, ByVal Size As Long)
  'Pass a pointer lpData and a length in bytes Size.
  If StartPos + Size > mSize Then Err.Raise 5, TypeName(Me), "Size must not exceed mapped size"
  If hObj = API_NULL Then Err.Raise &H8004930A, TypeName(Me), "WriteMap: Map not open"
  Ptr = lpMap + StartPos
  CopyMemory ByVal Ptr, ByVal lpData, ByVal Size
End Sub

Public Sub ReadMapPos(ByVal StartPos As Long, ByVal lpData As Long, ByVal Size As Long)
  'Pass a pointer lpData and a length in bytes Size.
  If StartPos + Size > mSize Then Err.Raise 5, TypeName(Me), "Size must not exceed mapped size"
  If hObj = API_NULL Then Err.Raise &H80049308, TypeName(Me), "ReadMap: Map not open"
  Ptr = lpMap + StartPos
  CopyMemory ByVal lpData, ByVal Ptr, ByVal Size
End Sub

Private Sub Class_Terminate()
  If hObj <> 0 Then CloseMap
End Sub
