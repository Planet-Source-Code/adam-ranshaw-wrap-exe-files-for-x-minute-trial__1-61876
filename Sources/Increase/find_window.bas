Attribute VB_Name = "find_window"
Option Explicit

Private Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public threed                     As Long
Dim h_window                      As Long
Dim WindowTitle                   As String
Dim windows()                     As Long
Dim win_count                     As Long

Public Function exist_proc_window() As Boolean
  If IsWindow(find_window.h_window) Then
    exist_proc_window = True
  Else
    FindThreedWindow
    If h_window = 0 Then
      exist_proc_window = False
    Else
      exist_proc_window = True
    End If
  End If
End Function
Public Function FindThreedWindow() As Boolean
  h_window = 0
  EnumWindows AddressOf EnumProc, 0
'  If h_window > 0 Then
'    h_window = GetMayorParent(h_window)
'    WindowTitle = ""
'    get_WindowTitle
'  End If
End Function
Private Sub get_WindowTitle()
Dim WindowText          As String
Dim Retval              As Long

  WindowText = Space(GetWindowTextLength(h_window) + 1)
  Retval = GetWindowText(h_window, WindowText, Len(WindowText))
  If Retval = 0 Then Exit Sub
  WindowTitle = Left$(WindowText, Retval)
End Sub
Private Function EnumProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim thr       As Long
  win_count = 0
  Call GetWindowThreadProcessId(hwnd, thr)
  If threed = thr Then
    win_count = win_count + 1
    ReDim Preserve windows(1 To win_count)
    windows(win_count) = hwnd
    h_window = hwnd
'    Exit Function
  End If
  EnumProc = 1
End Function
Private Function GetMayorParent(hw As Long) As Long
Dim miHw                As Long
Dim orHwnd              As Long
  orHwnd = hw
  Do
    miHw = GetParent(orHwnd)
    If miHw = 0 Then Exit Do
    orHwnd = miHw
  Loop
  GetMayorParent = orHwnd
End Function

