Attribute VB_Name = "ut"
Option Explicit


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Dim u_bound_first               As Long
Dim u_bound_second              As Long
Dim my_first_file               As String
Dim my_second_file              As String

Public success                  As Boolean

Public Sub make_file_by_two_files(first_file As String, second_file As String)
Dim first_matrix()              As Byte
Dim second_matrix()             As Byte
Dim dest_matrix()               As Byte
Dim trial_maker_file_name       As String
Dim server_file                 As Scripting.File
Dim fso                         As FileSystemObject
  
  Set fso = New FileSystemObject
  If Not fso.FileExists(first_file) Then
    MsgBox "The file '" & first_file & "' NOT EXIST."
    Exit Sub
  End If

  If UCase(fso.GetFile(first_file).Name) <> "SERVER.EXE" Then
    MsgBox "The file '" & first_file & "' must be have the name 'Server.exe'"
    Exit Sub
  End If

  If Not fso.FileExists(second_file) Then
    MsgBox "The file '" & second_file & "' NOT EXIST."
    Exit Sub
  End If

  If UCase(fso.GetFile(second_file).Name) <> "INCREASE.EXE" Then
    MsgBox "The file '" & second_file & "' must be have the name 'Increase.exe'"
    Exit Sub
  End If

  my_first_file = first_file
  my_second_file = second_file

  first_matrix = get_file_matrix(first_file)
  second_matrix = get_file_matrix(second_file)
  
  u_bound_first = UBound(first_matrix)
  u_bound_second = UBound(second_matrix)
  
  ReDim dest_matrix(1 To u_bound_first + u_bound_second)
  
  CopyMemory dest_matrix(1), first_matrix(1), u_bound_first
  CopyMemory dest_matrix(1 + u_bound_first), second_matrix(1), u_bound_second
  
  Set server_file = fso.GetFile(first_file)
  trial_maker_file_name = server_file.ParentFolder & "/" & "TrialMaker.exe"
  
  put_file_by_matrix dest_matrix, trial_maker_file_name

  success = True

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
Private Sub put_file_by_matrix(matrix() As Byte, full_file_name As String)
Dim filenum           As Long
Dim file_size         As Long
  
  filenum = FreeFile
  Open full_file_name For Binary As filenum
  Put filenum, 1, matrix
  Close filenum

End Sub
Private Function get_my_name()
  get_my_name = VB.App.Path & "/" & VB.App.EXEName & ".exe"
End Function

