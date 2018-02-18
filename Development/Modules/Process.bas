Attribute VB_Name = "Process"
Option Explicit

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
 dwProcessID As Long
 dwThreadID As Long
End Type

Const STARTF_USESHOWWINDOW = &H1&
Const NORMAL_PRIORITY_CLASS = &H20&
Const SW_HIDE = 3

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'--- Shells the passed command line and waits for the process to finish
'--- Returns the exit code of the shelled process
Function StartProcess(strCmdLine As String) As Long
  Dim udtProc As PROCESS_INFORMATION, udtStart As STARTUPINFO

  'initialize the STARTUPINFO structure
  udtStart.cb = Len(udtStart) 'size
  udtStart.dwFlags = STARTF_USESHOWWINDOW 'uses show window command
  udtStart.wShowWindow = SW_HIDE 'the hide window command

  'Launch the application
  CreateProcess vbNullString, strCmdLine, ByVal 0&, ByVal 0&, 0, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, udtStart, udtProc
  
End Function



