Attribute VB_Name = "exec"
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
   '===============
Private Type OVERLAPPED
  Internal As Long
  InternalHigh As Long
  offset As Long
  OffsetHigh As Long
  hEvent As Long
End Type
Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Long
End Type

Private Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, lpBuffer As Any, ByVal nBufferSize As Long, lpBytesRead As Long, lpTotalBytesAvail As Long, lpBytesLeftThisMessage As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
'========================

   Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long

   Private Declare Function GetExitCodeProcess Lib "kernel32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long

    Private Const NORMAL_PRIORITY_CLASS = &H20&
    Private Const STARTF_USESHOWWINDOW = 1
    Private Const STARTF_USESTDHANDLES = &H100
    Private Const APITRUE = 1
    Private Const pNull = 0

Public proc              As PROCESS_INFORMATION
Public Function ExecCmd(cmdline$)
    Dim start As STARTUPINFO
    Dim hReadStdOut As Long
    Dim hWriteStdOut As Long
    Dim ret         As Long
    Dim txtOut      As String
    Dim sChunk      As String
    Dim lblStatus   As String
    Dim hReadPipe   As Long
    Dim hWritePipe As Long
    
    Dim f As Long, cGot As Long, cPeek As Long, abChunk() As Byte, i
    Dim over As OVERLAPPED, saPipe As SECURITY_ATTRIBUTES, c As Long
    Dim cWant As Long
    
    cWant = 32000
    
    saPipe.nLength = LenB(saPipe)
    saPipe.bInheritHandle = APITRUE
    saPipe.lpSecurityDescriptor = pNull
    
    f = CreatePipe(hReadStdOut, hWriteStdOut, saPipe, 0)
    
    start.hStdOutput = hWriteStdOut
    start.cb = Len(start)
    start.lpTitle = "Update"
    start.wShowWindow = 0
    start.dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        
    ret& = CreateProcessA(cmdline$, 0&, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
'    ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
'    NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    i = 1
    txtOut = ""
    Do
'===========================


        f = PeekNamedPipe(hReadStdOut, pNull, 0, 0, cPeek, 0)
        If (f <> 0) And (cPeek <> 0) Then
            ReDim abChunk(0 To cWant - 1)
            Call ReadFile(hReadStdOut, abChunk(0), cWant, cGot, ByVal pNull)
            sChunk = LeftBytes(abChunk, cGot)
            txtOut = txtOut & sChunk
        End If
         
'===========================
        Call GetExitCodeProcess(proc.hProcess, ret&)
        DoEvents
        lblStatus = "Working: " & i
        i = i + 1
    Loop Until ret <> 259
         
    Call CloseHandle(proc.hThread)
    Call CloseHandle(proc.hProcess)
    Call CloseHandle(hReadPipe)
    Call CloseHandle(hWritePipe)
    ExecCmd = ret&
End Function
Private Function LeftBytes(ab() As Byte, ByVal iLen As Long) As String
    Dim s As String
    For i = 0 To iLen
        s = s & Chr(ab(i))
    Next
    LeftBytes = s
End Function



