Attribute VB_Name = "CmdOutput"
Option Explicit
''''''''''''''''''''''''''''''''''''''''
' Joacim Andersson, Brixoft Software
' http://www.brixoft.net
''''''''''''''''''''''''''''''''''''''''
' STARTUPINFO flags
Private Const STARTF_USESHOWWINDOW = &H1
Private Const STARTF_USESTDHANDLES = &H100
' ShowWindow flags
Private Const SW_HIDE = 0
' DuplicateHandle flags
Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2
' Error codes
Private Const ERROR_BROKEN_PIPE = 109

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Type STARTUPINFO
  cb As Long
  lpReserved As String
  lpDesktop As String
  lpTitle As String
  dwX As Long
  dwY As Long
  dwXSize As Long
  dwYSize As Long
  dwXCountChars As Long
  dwYCountChars As Long
  dwFillAttribute As Long
  dwFlags As Long
  wShowWindow As Integer
  cbReserved2 As Integer
  lpReserved2 As Long
  hStdInput As Long
  hStdOutput As Long
  hStdError As Long
End Type

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadId As Long
End Type

Private Declare Function CreatePipe _
                Lib "kernel32" (phReadPipe As Long, _
                                phWritePipe As Long, _
                                lpPipeAttributes As Any, _
                                ByVal nSize As Long) As Long
Private Declare Function ReadFile _
                Lib "kernel32" (ByVal hFile As Long, _
                                lpBuffer As Any, _
                                ByVal nNumberOfBytesToRead As Long, _
                                lpNumberOfBytesRead As Long, _
                                lpOverlapped As Any) As Long
Private Declare Function CreateProcess _
                Lib "kernel32" _
                Alias "CreateProcessA" (ByVal lpApplicationName As String, _
                                        ByVal lpCommandLine As String, _
                                        lpProcessAttributes As Any, _
                                        lpThreadAttributes As Any, _
                                        ByVal bInheritHandles As Long, _
                                        ByVal dwCreationFlags As Long, _
                                        lpEnvironment As Any, _
                                        ByVal lpCurrentDriectory As String, _
                                        lpStartupInfo As STARTUPINFO, _
                                        lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function DuplicateHandle _
                Lib "kernel32" (ByVal hSourceProcessHandle As Long, _
                                ByVal hSourceHandle As Long, _
                                ByVal hTargetProcessHandle As Long, _
                                lpTargetHandle As Long, _
                                ByVal dwDesiredAccess As Long, _
                                ByVal bInheritHandle As Long, _
                                ByVal dwOptions As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OemToCharBuff _
                Lib "user32" _
                Alias "OemToCharBuffA" (lpszSrc As Any, _
                                        ByVal lpszDst As String, _
                                        ByVal cchDstLength As Long) As Long

' Function GetCommandOutput
'
' sCommandLine:  [in] Command line to launch
' blnStdOut        [in,opt] True (defualt) to capture output to STDOUT
' blnStdErr        [in,opt] True to capture output to STDERR. False is default.
' blnOEMConvert:   [in,opt] True (default) to convert DOS characters to Windows, False to skip conversion
'
' Returns:       String with STDOUT and/or STDERR output
'
Public Function GetCommandOutput(sCommandLine As String, _
                                 Optional blnStdOut As Boolean = True, _
                                 Optional blnStdErr As Boolean = False, _
                                 Optional blnOEMConvert As Boolean = True) As String
  Dim hPipeRead   As Long, hPipeWrite1 As Long, hPipeWrite2 As Long
  Dim hCurProcess As Long
  Dim sa          As SECURITY_ATTRIBUTES
  Dim si          As STARTUPINFO
  Dim pi          As PROCESS_INFORMATION
  Dim baOutput()  As Byte
  Dim sNewOutput  As String
  Dim lBytesRead  As Long
  Dim fTwoHandles As Boolean
  Dim lRet        As Long
  Const BUFSIZE = 1024      ' pipe buffer size
  ' At least one of them should be True, otherwise there's no point in calling the function
  If (Not blnStdOut) And (Not blnStdErr) Then
    Err.Raise 5         ' Invalid Procedure call or Argument
  End If
  ' If both are true, we need two write handles. If not, one is enough.
  fTwoHandles = blnStdOut And blnStdErr
  ReDim baOutput(BUFSIZE - 1) As Byte

  With sa
    .nLength = Len(sa)
    .bInheritHandle = 1    ' get inheritable pipe handles
  End With

  If CreatePipe(hPipeRead, hPipeWrite1, sa, BUFSIZE) = 0 Then
    Exit Function
  End If
  hCurProcess = GetCurrentProcess()
  ' Replace our inheritable read handle with an non-inheritable. Not that it
  ' seems to be necessary in this case, but the docs say we should.
  Call DuplicateHandle(hCurProcess, hPipeRead, hCurProcess, hPipeRead, 0&, 0&, DUPLICATE_SAME_ACCESS Or DUPLICATE_CLOSE_SOURCE)
  ' If both STDOUT and STDERR should be redirected, get an extra handle.
  If fTwoHandles Then
    Call DuplicateHandle(hCurProcess, hPipeWrite1, hCurProcess, hPipeWrite2, 0&, 1&, DUPLICATE_SAME_ACCESS)
  End If

  With si
    .cb = Len(si)
    .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
    .wShowWindow = SW_HIDE          ' hide the window
    If fTwoHandles Then
      .hStdOutput = hPipeWrite1
      .hStdError = hPipeWrite2
    ElseIf blnStdOut Then
      .hStdOutput = hPipeWrite1
    Else
      .hStdError = hPipeWrite1
    End If
  End With

  If CreateProcess(vbNullString, sCommandLine, ByVal 0&, ByVal 0&, 1, 0&, ByVal 0&, vbNullString, si, pi) Then
    ' Close thread handle - we don't need it
    Call CloseHandle(pi.hThread)
    ' Also close our handle(s) to the write end of the pipe. This is important, since
    ' ReadFile will *not* return until all write handles are closed or the buffer is full.
    Call CloseHandle(hPipeWrite1)
    hPipeWrite1 = 0
    If hPipeWrite2 Then
      Call CloseHandle(hPipeWrite2)
      hPipeWrite2 = 0
    End If

    Do
      ' Add a DoEvents to allow more data to be written to the buffer for each call.
      ' This results in fewer, larger chunks to be read.
      'DoEvents
      If ReadFile(hPipeRead, baOutput(0), BUFSIZE, lBytesRead, ByVal 0&) = 0 Then
        Exit Do
      End If
      If blnOEMConvert Then
        ' convert from "DOS" to "Windows" characters
        sNewOutput = String$(lBytesRead, 0)
        Call OemToCharBuff(baOutput(0), sNewOutput, lBytesRead)
      Else
        ' perform no conversion (except to Unicode)
        sNewOutput = Left$(StrConv(baOutput(), vbUnicode), lBytesRead)
      End If
      GetCommandOutput = GetCommandOutput & sNewOutput
      ' If you are executing an application that outputs data during a long time,
      ' and don't want to lock up your application, it might be a better idea to
      ' wrap this code in a class module in an ActiveX EXE and execute it asynchronously.
      ' Then you can raise an event here each time more data is available.
      'RaiseEvent OutputAvailabele(sNewOutput)
    Loop

    ' When the process terminates successfully, Err.LastDllError will be
    ' ERROR_BROKEN_PIPE (109). Other values indicates an error.
    Call CloseHandle(pi.hProcess)
  Else
    GetCommandOutput = "Failed to create process, check the path of the command line."
  End If
  ' clean up
  Call CloseHandle(hPipeRead)
  If hPipeWrite1 Then
    Call CloseHandle(hPipeWrite1)
  End If
  If hPipeWrite2 Then
    Call CloseHandle(hPipeWrite2)
  End If
End Function

Public Function ExecuteCommand(ByVal CommandLine As String, _
                               Optional bShowWindow As Boolean = False, _
                               Optional sCurrentDir As String) As String
  Dim proc         As PROCESS_INFORMATION     'Process info filled by CreateProcessA
  Dim ret          As Long                     'long variable for get the return value of the
  'API functions
  Dim start        As STARTUPINFO            'StartUp Info passed to the CreateProceeeA
  'function
  Dim sa           As SECURITY_ATTRIBUTES       'Security Attributes passeed to the
  'CreateProcessA function
  Dim hReadPipe    As Long               'Read Pipe handle created by CreatePipe
  Dim hWritePipe   As Long              'Write Pite handle created by CreatePipe
  Dim lngBytesRead As Long            'Amount of byte read from the Read Pipe handle
  Dim strBuff      As String * 256         'String buffer reading the Pipe
  Dim mCommand     As String, mOutputs As String
  'if the parameter is not empty update the CommandLine property
  If Len(CommandLine) > 0 Then
    mCommand = CommandLine
  End If
  'if the command line is empty then exit whit a error message
  If Len(mCommand) = 0 Then
    ' msgbox "command line empty"
    Exit Function
  End If
  'Create the Pipe
  sa.nLength = Len(sa)
  sa.bInheritHandle = 1&
  sa.lpSecurityDescriptor = 0&
  ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
  If ret = 0 Then
    'If an error occur during the Pipe creation exit
    Debug.Print "CreatePipe failed. Error: " & Err.LastDllError
    Exit Function
  End If
  'Launch the command line application
  start.cb = Len(start)
  start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
  'set the StdOutput and the StdError output to the same Write Pipe handle
  start.hStdOutput = hWritePipe
  start.hStdError = hWritePipe
  '    start.hStdInput = hInReadPipe
  ' If bShowWindow Then
  '     start.wShowWindow = SW_SHOWNORMAL
  ' Else
  start.wShowWindow = SW_HIDE
  ' End If
  'Execute the command
  If Len(sCurrentDir) = 0 Then
    ret& = CreateProcess(vbNullString, mCommand, sa, sa, 1, 0&, ByVal 0&, vbNullString, start, proc)
  Else
    ret& = CreateProcess(0&, mCommand, sa, sa, 1&, 0&, 0&, sCurrentDir, start, proc)
  End If
  If ret <> 1 Then
    'if the command is not found ....
    Debug.Print "File or command not found in procedure ExecuteCommand"
    Exit Function
  End If
  'Now We can ... must close the hWritePipe
  ret = CloseHandle(hWritePipe)
  '    ret = CloseHandle(hInReadPipe)
  mOutputs = vbNullString

  'Read the ReadPipe handle
  Do
    ret = ReadFile(hReadPipe, strBuff, 256, lngBytesRead, 0&)
    mOutputs = mOutputs & Left$(strBuff, lngBytesRead)
    'Send data to the object via ReceiveOutputs event
  Loop While ret <> 0

  'Close the opened handles
  Call CloseHandle(proc.hProcess)
  Call CloseHandle(proc.hThread)
  Call CloseHandle(hReadPipe)
  'Return the Outputs property with the entire DOS output
  ExecuteCommand = mOutputs
End Function
