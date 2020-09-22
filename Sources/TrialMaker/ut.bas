Attribute VB_Name = "ut"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public fso                        As FileSystemObject
Public Const bkgr_name            As String = "fnd.bmp"

Private Type t_text_file
  matrix()            As Byte
  u_bound             As Long
  second_file_name    As String
End Type
Public Type RGB_tp
  r               As Long
  g               As Long
  b               As Long
  clr             As Long
End Type

'************************** !!!! NOT CHANGE !!!! *****************
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

Dim trial_exe_matrix()                  As Byte
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
Public pr_caption             As String
Public with_VB                As Boolean

Public is_trial_exe           As Boolean
Public program_file_name      As String

Private Const current_version As Long = 2


Public Sub make_trial_exe()
Dim first_matrix()              As Byte
Dim second_matrix()             As Byte
Dim icon_matrix()               As Byte
Dim dest_matrix()               As Byte
Dim text_file_matrix()          As Byte
Dim u_bound_first               As Long
Dim u_bound_second              As Long
Dim u_bound_text_file           As Long
Dim u_bound_icon_file           As Long
Dim u_bound_trial_exe           As Long

  icon_matrix = get_icon_matrix(program_file_name)
  first_matrix = get_first_file_matrix
  second_matrix = get_file_matrix(program_file_name)
  
  u_bound_icon_file = UBound(icon_matrix)
  u_bound_first = CLng(increase_size)
  u_bound_second = UBound(second_matrix)
  u_bound_trial_exe = UBound(trial_exe_matrix)
  
  text_file.second_file_name = program_file_name
  sizes.icon_file_size = u_bound_icon_file
  make_text_file text_file, program_file_name
  
  ReDim dest_matrix(1 To u_bound_first + u_bound_second + u_bound_icon_file + text_file.u_bound + u_bound_trial_exe)
  
  CopyMemory dest_matrix(1), first_matrix(1), u_bound_first
  CopyMemory dest_matrix(u_bound_first + 1), text_file.matrix(1), text_file.u_bound
  CopyMemory dest_matrix(text_file.u_bound + u_bound_first + 1), icon_matrix(1), u_bound_icon_file
  CopyMemory dest_matrix(u_bound_icon_file + u_bound_first + text_file.u_bound + 1), second_matrix(1), u_bound_second
  CopyMemory dest_matrix(u_bound_icon_file + u_bound_first + text_file.u_bound + u_bound_second + 1), trial_exe_matrix(1), u_bound_trial_exe
  
  On Error Resume Next
  Kill program_file_name
  Err.Clear
  
  put_file_by_matrix dest_matrix, program_file_name
  
'  copy_back_ground

End Sub
Private Sub copy_back_ground()
Dim file_name_orig        As String
Dim file_name_dest        As String
Dim program_file          As Scripting.File

  file_name_orig = App.Path & "/" & bkgr_name
  If fso Is Nothing Then Set fso = New FileSystemObject
  If Not fso.FileExists(file_name_orig) Then Exit Sub
  Set program_file = fso.GetFile(program_file_name)
  file_name_dest = program_file.ParentFolder & "/" & bkgr_name
  
  If file_name_orig = file_name_dest Then Exit Sub
  
  If fso.FileExists(file_name_dest) Then
    On Error Resume Next
    Kill file_name_dest
    Err.Clear
  End If
  
  On Error Resume Next
  FileCopy file_name_orig, file_name_dest
  Err.Clear
End Sub
Public Sub make_not_trial()
Dim my_matrix()             As Byte
Dim exe_matrix()            As Byte
Dim begin_copy              As Long

  my_matrix = get_file_matrix(program_file_name)
  
  get_text_file
  sizes.program_size = sizes.program_size
  begin_copy = 1 + increase_size + txt_size + sizes.icon_file_size
  
  ReDim exe_matrix(1 To sizes.program_size)
  CopyMemory exe_matrix(1), my_matrix(begin_copy), sizes.program_size
  
  On Error Resume Next
  Kill program_file_name
  Err.Clear
  
  put_file_by_matrix exe_matrix, program_file_name

End Sub
Public Sub detect_trial_exe()

  sizes.full_size = FileLen(program_file_name)
  sizes.full_matrix = get_file_matrix(program_file_name)
  
  is_trial_exe = False
  If detect_version_1_trial_exe Then
    sizes.version = 1
    sizes.n_increase_size = e_trial_fin.vers1_increase_size
    sizes.n_txt_size = e_trial_fin.vers1_txt_size
    is_trial_exe = True
  ElseIf detect_version_1_trial_exe Then
    sizes.version = 2
    is_trial_exe = True
  End If
  
  
End Sub
Public Function detect_version_2_trial_exe() As Boolean
Dim nm                      As Long
Dim sz                      As Long
'e_trial_fin
  Stop
  detect_version_2_trial_exe = True
  For nm = e_trial_fin.first_byte_for_detect To trial_exe_fin
    If sizes.full_matrix(sizes.full_size - trial_exe_fin + nm) <> trial_exe_matrix(nm) Then
      detect_version_2_trial_exe = False
      Exit For
    End If
  Next
  If detect_version_2_trial_exe Then
    Stop
    
    sz = 0
    For nm = e_trial_fin.first_byte_version To e_trial_fin.end_byte_version
      sz = sz & trial_exe_matrix(nm)
    Next
    sizes.version = CLng(sz)
    
    sz = 0
    For nm = e_trial_fin.first_byte_increase_size To e_trial_fin.end_byte_increase_size
      sz = sz & trial_exe_matrix(nm)
    Next
    sizes.n_increase_size = CLng(sz)
    
    sz = 0
    For nm = e_trial_fin.first_byte_txt_size To e_trial_fin.first_byte_txt_size
      sz = sz & trial_exe_matrix(nm)
    Next
    sizes.n_txt_size = CLng(sz)
    
    
  End If
  
End Function
Public Sub set_trial_exe_fin()
Dim nm              As Long
Dim txt             As String
'Const new_vers      As Boolean = True
Const new_vers      As Boolean = False

  If new_vers Then
    ReDim trial_exe_matrix(1 To trial_exe_fin)
    make_one_item CStr(current_version), e_trial_fin.first_byte_version, e_trial_fin.end_byte_version
    make_one_item CStr(sizes.n_increase_size), e_trial_fin.vers1_increase_size, e_trial_fin.end_byte_increase_size
    make_one_item CStr(sizes.n_txt_size), e_trial_fin.first_byte_txt_size, e_trial_fin.end_byte_txt_size
  
    For nm = e_trial_fin.first_byte_for_detect To trial_exe_fin / 2
      trial_exe_matrix(nm) = nm
    Next
    For nm = trial_exe_fin / 2 To 1 Step -1
      trial_exe_matrix(nm) = nm
    Next
  Else
    ReDim trial_exe_matrix(1 To trial_exe_fin)
    For nm = 1 To trial_exe_fin / 2
      trial_exe_matrix(nm) = nm
    Next
    For nm = trial_exe_fin / 2 To 1 Step -1
      trial_exe_matrix(nm) = nm
    Next
  End If
  
End Sub
Private Sub make_one_item(value As String, first_byte As Long, end_byte As Long)
Dim txt       As String
Dim nm          As Long
  txt = value
  txt = txt & String(end_byte - first_byte + 1 - Len(txt), "0")
  For nm = first_byte To end_byte
    trial_exe_matrix(nm) = VBA.Left$(txt, nm - first_byte + 1)
  Next
  
End Sub

Public Function detect_version_1_trial_exe() As Boolean
Dim nm                      As Long

  detect_version_1_trial_exe = True
  For nm = 1 To e_trial_fin.vers1_trial_exe_fin
    If sizes.full_matrix(sizes.full_size - trial_exe_fin + nm) <> trial_exe_matrix(nm) Then
      detect_version_1_trial_exe = False
      Exit For
    End If
  Next
  
End Function

Private Function get_first_file_matrix() As Variant
Dim first_file_matrix()     As Byte
Dim my_matrix()             As Byte
Dim full_name               As String
Dim icon_file_name          As String
Dim my_size                 As Long
Const icon_ext              As String = "bmp"

  full_name = get_my_name
  my_size = FileLen(full_name)
  my_matrix = get_file_matrix(full_name)

  ReDim first_file_matrix(1 To CLng(increase_size))
  CopyMemory first_file_matrix(1), my_matrix(1 + my_size - CLng(increase_size)), CLng(increase_size)
  get_first_file_matrix = first_file_matrix
  Erase my_matrix
End Function

Private Function get_my_name()
  get_my_name = VB.App.Path & "/" & VB.App.EXEName & ".exe"
End Function

Private Function get_icon_matrix(first_file As String) As Variant
Dim icon_file             As Scripting.File
Dim icon_file_name        As String

Const icon_ext            As String = "bmp"

  If fso Is Nothing Then Set fso = New FileSystemObject
  icon_file_name = get_temp_name(VB.App.Path & "/", icon_ext)
  put_icon first_file, icon_file_name
  get_icon_matrix = get_file_matrix(icon_file_name)
  
  On Error Resume Next
  Kill icon_file_name
  Err.Clear
End Function

Private Sub put_icon(path_origen As String, icon_file_path As String)
  Form_Init.Picture_icon.AutoSize = True
  Set Form_Init.Picture_icon.Picture = GetIconFromFile(path_origen, 0, True)
  SavePicture Form_Init.Picture_icon.Image, icon_file_path
End Sub
Private Function get_temp_name(folder As String, ext As String) As String
Dim f_name          As String
Dim full_name       As String
  If fso Is Nothing Then Set fso = New FileSystemObject
  f_name = "1!"
  Do
    full_name = folder & f_name & "." & ext
    If Not fso.FileExists(full_name) Then
      Exit Do
    End If
    f_name = f_name & "_!"
  Loop
  get_temp_name = full_name
End Function
Private Function make_text_file(text_file As t_text_file, second_file_name As String) As Variant
Dim txt                     As String
Dim second_file_size        As Long
Dim second_file             As Scripting.File
Dim txt_file                As Scripting.File
Dim dest_file               As Scripting.File
Dim ts                      As Scripting.TextStream
Dim nm                      As Long
Dim pasw                    As String

  If fso Is Nothing Then Set fso = New FileSystemObject
  Set second_file = fso.GetFile(text_file.second_file_name)
  second_file_size = second_file.Size
  Set dest_file = fso.GetFile(second_file)
  pr_caption = Mid$(dest_file.Name, 1, Len(dest_file.Name) - 4)
  
  pasw = Trim(Form_Init.Text_Password.Text)
  pasw = EncryptPassword_1(8, pasw)
  txt = ""
  txt = txt & cs_password_beg & pasw & cs_password_fin
  txt = txt & cs_allowed_minutes_beg & Val(Trim(Form_Init.Text_Allowed_minutes.Text)) & cs_allowed_minutes_fin
  txt = txt & cs_caption_beg & Trim(pr_caption) & cs_caption_fin
  txt = txt & cs_icon_size_beg & CStr(sizes.icon_file_size) & cs_icon_size_fin
  txt = txt & cs_message_beg & sizes.message & cs_message_fin
   
  txt = txt & Space(txt_size - Len(txt))
  text_file.u_bound = Len(txt)
  ReDim text_file.matrix(1 To text_file.u_bound)
  For nm = 1 To text_file.u_bound
    text_file.matrix(nm) = Asc(Mid$(txt, nm, 1))
  Next
  
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

Public Function set_with_VB() As Boolean
  with_VB = True
  set_with_VB = True
End Function

Public Sub crypt()
Dim pasw        As String
Dim encr_pasw   As String
Dim decr_pasw   As String
  pasw = "123321"
  encr_pasw = EncryptPassword_1(8, pasw)
  decr_pasw = raschi_1(8, encr_pasw)
  
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

Public Sub get_text_file()
Dim txt               As String
  
  txt = get_text
  
  sizes.icon_file_size = Val(get_content_by_beg_fin(cs_icon_size_beg, cs_icon_size_fin, txt))
  sizes.program_size = sizes.total_size - increase_size - txt_size - sizes.icon_file_size - trial_exe_fin
  
  sizes.password = get_content_by_beg_fin(cs_password_beg, cs_password_fin, txt)
  sizes.minutes = Val(get_content_by_beg_fin(cs_allowed_minutes_beg, cs_allowed_minutes_fin, txt))
  sizes.time_play = sizes.minutes * 60 * 1000
  sizes.caption = get_content_by_beg_fin(cs_caption_beg, cs_caption_fin, txt)
  sizes.message = get_content_by_beg_fin(cs_message_beg, cs_message_fin, txt)
  
  sizes.password = raschi_1(8, sizes.password)
End Sub
Private Function get_text() As String
Dim my_file           As File
Dim my_matrix()       As Byte
Dim nm                As Long
Dim txt               As String
  
  If fso Is Nothing Then Set fso = New FileSystemObject
  
  Set my_file = fso.GetFile(program_file_name)
  
  my_matrix = get_file_matrix(program_file_name)
  For nm = CLng(increase_size) + 1 To CLng(increase_size) + txt_size
    txt = txt & Chr$(my_matrix(nm))
  Next
  get_text = txt
  sizes.total_size = my_file.Size

  Erase my_matrix
End Function
Public Sub set_new_text_file()
Dim txt       As String
  text_file.second_file_name = program_file_name
  make_text_file text_file, program_file_name
  put_file_by_matrix text_file.matrix, program_file_name, increase_size + 1
  get_text_file
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

