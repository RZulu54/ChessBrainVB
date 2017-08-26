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
Private Const API_NULL As Long = 0
Private Const API_FALSE As Long = 0
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const PAGE_READWRITE As Long = 4
Private Const SECTION_MAP_WRITE = &H2
Private Const FILE_MAP_WRITE = SECTION_MAP_WRITE
Private Const ERROR_ALREADY_EXISTS As Long = 183

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long)
    
Private Declare Function RtlCompareMemory Lib "ntdll" ( _
    ByRef Source1 As Any, _
    ByRef Source2 As Any, _
    ByVal Length As Long) As Long

Private Declare Function CreateFileMapping Lib "kernel32" Alias "CreateFileMappingA" ( _
    ByVal hFile As Long, _
    ByVal lpFileMappigAttributes As Long, _
    ByVal flProtect As Long, _
    ByVal dwMaximumSizeHigh As Long, _
    ByVal dwMaximumSizeLow As Long, _
    ByVal lpName As String) As Long
Private Declare Function MapViewOfFile Lib "kernel32" ( _
    ByVal hFileMappingObject As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwFileOffsetHigh As Long, _
    ByVal dwFileOffsetLow As Long, _
    ByVal dwNumberOfBytesToMap As Long) As Long
Private Declare Function OpenFileMapping Lib "kernel32" Alias "OpenFileMappingA" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal lpName As String) As Long
Private Declare Function UnmapViewOfFile Lib "kernel32" (ByVal lpBaseAddress As Long) As Long

Private Declare Sub ZeroMemory2 Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Private hObj As Long
Private lpMap As Long
Private mSize As Long
Private Ptr As Long

Private CheckData(128) As Byte, CheckDataPtr As Long


Public Sub CloseMap()
    UnmapViewOfFile lpMap
    CloseHandle hObj
    hObj = 0
End Sub

Public Function CreateMap(ByVal Name As String, ByVal Size As Long) As Boolean
    'Returns True if the memory mapped file already exists.  If so CAUTION,
    'the size will be its previously-created size.
    
    If Size < 1 Then Err.Raise 5, TypeName(Me), "Size must be at least 1 byte"
    hObj = CreateFileMapping(INVALID_HANDLE_VALUE, _
                             API_NULL, _
                             PAGE_READWRITE, _
                             0, _
                             Size, _
                             Name)
    If hObj = API_NULL Then
        Err.Raise &H80049300, _
                  TypeName(Me), _
                  "CreateFileMapping system error " & CStr(Err.LastDllError)
    End If
    CreateMap = (Err.LastDllError = ERROR_ALREADY_EXISTS)
    lpMap = MapViewOfFile(hObj, FILE_MAP_WRITE, 0, 0, 0)
    If lpMap = API_NULL Then
        Err.Raise &H80049302, _
                  TypeName(Me), _
                  "MapViewOfFile system error " & CStr(Err.LastDllError)
    End If
    ZeroMemory2 ByVal lpMap, Size
    mSize = Size
End Function

Public Sub OpenMap(ByVal Name As String, ByVal Size As Long)
    If Size < 1 Then Err.Raise 5, TypeName(Me), "Size must be at least 1 byte"
    hObj = OpenFileMapping(FILE_MAP_WRITE, API_FALSE, Name)
    If hObj = API_NULL Then
        Err.Raise &H80049304, _
                  TypeName(Me), _
                  "OpenFileMapping system error " & CStr(Err.LastDllError)
    End If
    lpMap = MapViewOfFile(hObj, FILE_MAP_WRITE, 0, 0, Size)
    If lpMap = API_NULL Then
        Err.Raise &H80049306, _
                  TypeName(Me), _
                  "MapViewOfFile system error " & CStr(Err.LastDllError)
    End If
    mSize = Size
    CheckDataPtr = VarPtr(CheckData(0))
End Sub

Public Sub ReadMapHashCluster(ByVal Index As Long, ByVal lpData As Long, ByVal Size As Long)
    'Pass a pointer lpData and a length in bytes Size.
    If Index * HashRecLen + Size > mSize Then Err.Raise 5, TypeName(Me), "Size must not exceed mapped size"
    If hObj = API_NULL Then Err.Raise &H80049308, TypeName(Me), "ReadMap: Map not open"
    Ptr = lpMap + Index * HashRecLen
    CopyMemory ByVal lpData, ByVal Ptr, ByVal Size
End Sub

Public Sub WriteMapHashEntry(ByVal ReplaceIndex As Long, ByVal lpData As Long)
    'Pass a pointer lpData and a length in bytes Size.
    Dim i As Long
    If (ReplaceIndex + 1) * HashRecLen > mSize Then Err.Raise 5, TypeName(Me), "Size must not exceed mapped size"
    If hObj = API_NULL Then Err.Raise &H8004930A, TypeName(Me), "WriteMap: Map not open"
    Ptr = lpMap + ReplaceIndex * HashRecLen
    For i = 1 To 3
      CopyMemory ByVal Ptr, ByVal lpData, ByVal HashRecLen
      '--- Reread the written entry to verify that there was no parallel write from other thread that mixed up the data
      '--- Try max 3 times
      If RtlCompareMemory(ByVal Ptr, ByVal lpData, ByVal HashRecLen) <> HashRecLen Then
        ' Difference found => try again
        'WriteTrace "HashMapDifference " & Index & "/" & i & " / " & Now()
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
