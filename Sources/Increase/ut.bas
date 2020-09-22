Attribute VB_Name = "ut"
Option Explicit


Public Const bkgr_name            As String = "fnd.bmp"

Public Type RGB_tp
  r               As Long
  g               As Long
  b               As Long
  clr             As Long
End Type

Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const WS_CLIPCHILDREN = &H2000000

Private Const SWP_ASYNCWINDOWPOS = &H4000
Private Const SWP_DEFERERASE = &H2000
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const PROCESS_ALL_ACCESS = &H1F0FFF
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Private Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const INVALID_HANDLE_VALUE = -1&

Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800


Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public fso                        As FileSystemObject

Private Type t_text_file
  matrix()            As Byte
  u_bound             As Long
End Type

Public destinate_name           As String

'************************** NOT CHANGE !!*****************
Private Const cs_password_beg           As String = "pswrdb"
Private Const cs_password_fin           As String = "pswrdf"
Private Const cs_allowed_minutes_beg    As String = "almnb"
Private Const cs_allowed_minutes_fin    As String = "almnf"
Private Const cs_caption_beg            As String = "captb"
Private Const cs_caption_fin            As String = "captf"
Private Const cs_icon_size_beg          As String = "icsb"
Private Const cs_icon_size_fin          As String = "icsf"
Private Const cs_message_beg            As String = "msgb"
Private Const cs_message_fin            As String = "msgf"

Private Const txt_size                  As Long = 5000
Private Const trial_exe_fin             As Long = 200

'********* Const increase_size is the size of the Increase.exe ***********
Private Const increase_size             As Long = 65536
'**************************************************************

Public Enum e_form
  e_Form_Init = 1
  e_Form_show_time
End Enum

Public Enum Acc_Mouse
  down = 1
  Move
  Up
End Enum

Public Enum e_trial_fin
  
  first_byte_version = 1
  end_byte_version = 2
  
  first_byte_increase_size = 3
  end_byte_increase_size = 8
  
  first_byte_txt_size = 9
  end_byte_txt_size = 13
  
  first_byte_for_detect = 14
  
  vers1_txt_size = 5000
  vers1_trial_exe_fin = 200
  vers1_increase_size = 61440

End Enum

Public Type t_sizes
  message                         As String
  total_size                      As Long
  program_size                    As Long
  icon_file_size                  As Long
  password                        As String
  minutes                         As Long
  caption                         As String
  time_play                       As Long
  version                         As Long
  full_matrix()                   As Byte
  full_size                       As Long
  
  n_txt_size                      As Long
  n_increase_size                 As Long

End Type
Public sizes                  As t_sizes
Dim text_file                 As t_text_file

Public fnd_name               As String

Public with_VB                    As Boolean
Public ok_password                As Boolean
Dim id_app                        As Long
Dim h_proc_window                 As Long
Dim begin_time                    As Long

Dim program_file_handle           As Long
Dim program_file_handle_2         As Long

Public Sub get_text_file()
Dim full_name         As String
Dim my_file           As File
Dim my_matrix()       As Byte
Dim nm                As Long
Dim txt               As String
Dim ub_sec            As Long
Dim ub_txt            As Long
Dim ub_fin            As Long
  
  If fso Is Nothing Then Set fso = New FileSystemObject
  
  full_name = get_my_name
  Set my_file = fso.GetFile(full_name)
  
  my_matrix = get_file_matrix(full_name)
  For nm = increase_size + 1 To increase_size + txt_size
    txt = txt & Chr$(my_matrix(nm))
  Next
  
  sizes.total_size = my_file.Size
  sizes.icon_file_size = Val(get_content_by_beg_fin(cs_icon_size_beg, cs_icon_size_fin, txt))
  sizes.program_size = sizes.total_size - increase_size - txt_size - sizes.icon_file_size - trial_exe_fin
  
  sizes.password = get_content_by_beg_fin(cs_password_beg, cs_password_fin, txt)
  sizes.minutes = Val(get_content_by_beg_fin(cs_allowed_minutes_beg, cs_allowed_minutes_fin, txt))
  sizes.time_play = sizes.minutes * 60 * 1000
  sizes.caption = get_content_by_beg_fin(cs_caption_beg, cs_caption_fin, txt)
  sizes.message = get_content_by_beg_fin(cs_message_beg, cs_message_fin, txt)

  Erase my_matrix
End Sub
Public Sub set_form_icon()
Dim icon_matrix()     As Byte
Dim my_matrix()       As Byte
Dim full_name         As String
Dim icon_file_name    As String
Const icon_ext        As String = "bmp"


  full_name = get_my_name
  my_matrix = get_file_matrix(full_name)

  ReDim icon_matrix(1 To sizes.icon_file_size)
  CopyMemory icon_matrix(1), my_matrix(1 + increase_size + txt_size), sizes.icon_file_size
  icon_file_name = get_temp_name(VB.App.path & "/", icon_ext)
  put_file_by_matrix icon_matrix, icon_file_name
  
  Form_Init.Picture_Icon.AutoRedraw = True
  Form_Init.Picture_Icon.AutoSize = True
  
  On Error Resume Next
  Set Form_Init.Picture_Icon = LoadPicture(icon_file_name)
  Err.Clear

  On Error Resume Next
  Kill icon_file_name
  Err.Clear

End Sub
Private Function get_temp_name(folder As String, ext As String) As String
Dim f_name          As String
Dim full_name       As String
  If fso Is Nothing Then Set fso = New FileSystemObject
  f_name = "1!"
  Do
    full_name = folder & "/" & f_name & "." & ext
    If Not fso.FileExists(full_name) Then
      Exit Do
    End If
    f_name = f_name & "_!"
  Loop
  get_temp_name = full_name
End Function

Private Function get_my_name()
Dim my_file       As Scripting.File
  If fso Is Nothing Then Set fso = New FileSystemObject
  If with_VB Then
    get_my_name = VB.App.path & "/" & VB.App.EXEName & ".exe"
    Set my_file = fso.GetFile(get_my_name)
    If destinate_name = "" Then
      Stop
      destinate_name = "C:\Documents and Settings\MOUSE\Mis documentos\Trabajos mios\RentaCoder\Actuales\Registration Trial Security\spider.exe"
    End If
'    get_my_name = my_file.ParentFolder.ParentFolder & "/" & destinate_name
    get_my_name = destinate_name
  Else
    get_my_name = VB.App.path & "/" & VB.App.EXEName & ".exe"
  End If
End Function
Public Sub put_program_exe_copy()
Dim three_matrix()          As Byte
Dim my_matrix()             As Byte
Dim full_name               As String
Dim begin_copy              As Long

Dim exe_file                As Scripting.File

  full_name = get_my_name
  my_matrix = get_file_matrix(full_name)
  ReDim three_matrix(1 To sizes.program_size)
  begin_copy = 1 + increase_size + txt_size + sizes.icon_file_size
  CopyMemory three_matrix(1), my_matrix(begin_copy), sizes.program_size
  
  On Error Resume Next
  Kill get_program_exe_copy_name
  Err.Clear
  
  put_file_by_matrix three_matrix, get_program_exe_copy_name
  DoEvents
  
  If fso Is Nothing Then Set fso = New FileSystemObject
  
  On Error Resume Next
  Set exe_file = fso.GetFile(get_program_exe_copy_name)
  If Err.Number <> 0 Then
    Err.Clear
    Exit Sub
  Else
    exe_file.Attributes = Hidden + System
  End If

End Sub
Public Function get_program_exe_copy_name() As String
Static f_name         As String
Dim folder            As Scripting.folder
Dim folder_name       As String

  If f_name = "" Then
    If with_VB Then
      If fso Is Nothing Then Set fso = New FileSystemObject
      Set folder = fso.GetFolder(VB.App.path)
      folder_name = folder.ParentFolder
    Else
      folder_name = VB.App.path
    End If
    f_name = get_temp_name(folder_name, "exe")
  
  End If
  get_program_exe_copy_name = f_name
'  get_program_exe_copy_name = VB.App.Path & "/" & "tercer1.exe"
End Function

Public Sub execute_program_exe_copy()
Dim shell_str                   As String
  get_text_file
  put_program_exe_copy
  shell_str = get_program_exe_copy_name
  
  id_app = Shell(shell_str, vbNormalFocus)
  my_lock_file

  get_process_window
            
  Form_Init.Visible = False
  Form_show_time.Show
  Form_show_time.Label_red.Width = 0
  begin_time = timeGetTime
  Form_Init.Timer_proc.Interval = 400
  Form_Init.Timer_proc.Enabled = True

End Sub
Public Sub my_Timer_proc_Timer()
Dim now_time        As Long
Dim term            As Boolean
  
  now_time = timeGetTime
  If process_not_exist Then
    term = True
  ElseIf now_time - begin_time > sizes.time_play Then
    If Not ok_password Then
      term = True
    End If
  Else
    If Not ok_password Then
      Form_show_time.Label_red.Width = Form_show_time.Label_verde.Width * ((now_time - begin_time) / sizes.time_play)
    End If
  End If
  If term Then
    my_TerminateProcess
    Unload Form_Init
    my_end "term"
 End If
End Sub
Private Function process_not_exist() As Boolean
  If exist_proc_window Then
    process_not_exist = False
  Else
    process_not_exist = True
  End If
End Function
Public Sub my_TerminateProcess()
Dim hProcess        As Long
Dim exe_file        As Scripting.File

  hProcess = OpenProcess(PROCESS_ALL_ACCESS, 1, id_app)
  TerminateProcess hProcess, 0
  
  If Not file_exist(get_program_exe_copy_name) Then Exit Sub
  my_unlock_file
  Set exe_file = fso.GetFile(get_program_exe_copy_name)
  exe_file.Attributes = Normal
  Do
    DoEvents
    
    On Error Resume Next
    Kill get_program_exe_copy_name
    Err.Clear
    
    If Not file_exist(get_program_exe_copy_name) Then Exit Do
  Loop
  Erase sizes.full_matrix

End Sub
Private Sub get_process_window()
  find_window.threed = id_app
  FindThreedWindow
End Sub
Private Function get_content_by_beg_fin(beg_str As String, fin_str As String, txt As String) As String
Dim beg       As Long
Dim fin       As Long
  beg = InStr(1, txt, beg_str)
  fin = InStr(1, txt, fin_str)
  If beg = 0 Or fin = 0 Then Exit Function
  beg = beg + Len(beg_str)
  fin = fin - beg
  get_content_by_beg_fin = Mid$(txt, beg, fin)
'  Stop
End Function

Public Function set_with_VB() As Boolean
  with_VB = True
  set_with_VB = True
End Function
Private Sub put_file_by_matrix(matrix() As Byte, full_file_name As String, Optional begin As Long = 1)
Dim filenum           As Long
Dim file_size         As Long
  
  filenum = FreeFile
  Open full_file_name For Binary As filenum
  Put filenum, begin, matrix
  Close filenum

End Sub
Private Function get_file_matrix(full_file_name As String) As Variant
Dim m_file()            As Byte
Dim file_size           As Long
Dim filenum             As Long

  file_size = FileLen(full_file_name)
  ReDim m_file(1 To file_size)
  filenum = FreeFile
  Open full_file_name For Binary As filenum
  Get filenum, 1, m_file
  Close filenum

  get_file_matrix = m_file
End Function
Public Sub my_SetWindowPos()
  Call SetWindowPos(Form_Init.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOMOVE)
End Sub

Public Sub Processing_mouse_Move(ac As Long, X As Single, Y As Single, frm As e_form)
Static Mover          As Boolean
Static y_In           As Long
Static x_In           As Long
Static down           As Boolean

  If ac = Acc_Mouse.down Then
    Mover = True
    y_In = Y
    x_In = X
    down = True
  ElseIf ac = Acc_Mouse.Move Then
    If X = x_In And Y = y_In Then Exit Sub
    
    If frm = e_Form_Init Then
      Form_Init.Move Form_Init.Left + X - x_In, Form_Init.Top + Y - y_In
    ElseIf frm = e_Form_show_time Then
      Form_show_time.Move Form_show_time.Left + X - x_In, Form_show_time.Top + Y - y_In
    End If
    
  ElseIf ac = Acc_Mouse.Up Then
    Mover = False
  End If
End Sub
  

Public Sub veryfi_password(pasw As String)
  If pasw = raschi_1(8, sizes.password) Then
    ok_password = True
  Else
    ok_password = False
  End If
End Sub
Private Function EncryptPassword_1(Number As Byte, DecryptedPassword As String) As String
Dim Counter     As Long

  Counter = 1
  Do Until Counter = Len(DecryptedPassword) + 1
    EncryptPassword_1 = EncryptPassword_1 & Chr$(Choose((Counter Mod 2) + 1, Asc(Mid(DecryptedPassword, Counter, 1)) - Number, Asc(Mid(DecryptedPassword, Counter, 1)) + Number) Xor (10 - Number))
    Counter = Counter + 1
  Loop
End Function

Private Function raschi_1(Number As Byte, EncryptedPassword As String) As String
Dim Counter       As Byte
  Counter = 1
  Do Until Counter = Len(EncryptedPassword) + 1
    raschi_1 = raschi_1 & Chr$(Choose((Counter Mod 2) + 1, (Asc(Mid(EncryptedPassword, Counter, 1)) Xor (10 - Number)) + Number, (Asc(Mid(EncryptedPassword, Counter, 1)) Xor (10 - Number)) - Number))
    Counter = Counter + 1
  Loop
End Function


Public Function Color_Entre(mColor1 As Long, mColor2 As Long, proc As Long) As Long
Dim VR            As Single
Dim VG            As Single
Dim VB            As Single

Dim RGB_1         As RGB_tp
Dim RGB_2         As RGB_tp

  RGB_1.clr = mColor1
  Call Determinar_RGB(RGB_1)

  RGB_2.clr = mColor2
  Call Determinar_RGB(RGB_2)

  VR = Abs(RGB_1.r - RGB_2.r) / 100
  VG = Abs(RGB_1.g - RGB_2.g) / 100
  VB = Abs(RGB_1.b - RGB_2.b) / 100

  If RGB_2.r < RGB_1.r Then VR = -VR
  If RGB_2.g < RGB_1.g Then VG = -VG
  If RGB_2.b < RGB_1.b Then VB = -VB

  Color_Entre = RGB(RGB_1.r + VR * proc, RGB_1.g + VG * proc, RGB_1.b + VB * proc)
End Function
Public Sub Determinar_RGB(miRGB As RGB_tp)
'**********************************
  With miRGB
    .r = (.clr And 255) And 255
    .g = Int(.clr / 256) And 255
    .b = Int(.clr / 65536) And 255
  End With
End Sub

Public Function file_exist(file_name As String) As Boolean
  If fso Is Nothing Then Set fso = New FileSystemObject
  file_exist = fso.FileExists(file_name)
End Function
Private Sub my_lock_file()
Dim hDevice               As Long
Dim path                  As String
Dim fl                    As Long
Dim gen                   As Long
Dim file_size             As Long
Dim ret                   As Long

  path = get_program_exe_copy_name
  If fso Is Nothing Then Set fso = New FileSystemObject
  file_size = fso.GetFile(path).Size
  
  fl = FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_SYSTEM Or FILE_ATTRIBUTE_NORMAL
  gen = GENERIC_READ
  program_file_handle = CreateFile(path, gen, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, fl, 0)
  program_file_handle_2 = CreateFile(path, gen, FILE_SHARE_READ, ByVal 0&, OPEN_EXISTING, fl, 0)

  ret = LockFile(program_file_handle, LoWord(0), HiWord(0), LoWord(file_size), HiWord(file_size))
End Sub
Private Sub my_unlock_file()
Dim hDevice               As Long
Dim path                  As String
Dim fl                    As Long
Dim file_size             As Long
Dim ret                   As Long
'
  path = get_program_exe_copy_name
  If fso Is Nothing Then Set fso = New FileSystemObject
  file_size = fso.GetFile(path).Size
  
  ret = UnlockFile(program_file_handle, LoWord(0), HiWord(0), LoWord(file_size), HiWord(file_size))
  CloseHandle program_file_handle
  CloseHandle program_file_handle_2

End Sub
Public Function LoWord(ByVal dw As Long) As Integer
  If dw And &H8000& Then
    LoWord = dw Or &HFFFF0000
  Else
    LoWord = dw And &HFFFF&
  End If
End Function

Public Function HiWord(ByVal dw As Long) As Integer
  HiWord = (dw And &HFFFF0000) \ 65536
End Function

Public Sub my_end(txt As String)
'  MsgBox txt
  End
End Sub
