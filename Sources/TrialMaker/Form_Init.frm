VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_Init 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Init"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   Icon            =   "Form_Init.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form_Init.frx":08CA
   ScaleHeight     =   3825
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer_SUCCES 
      Left            =   480
      Top             =   3360
   End
   Begin VB.PictureBox Picture_icon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   9960
      Picture         =   "Form_Init.frx":3038C
      ScaleHeight     =   495
      ScaleWidth      =   975
      TabIndex        =   6
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text_Allowed_minutes 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3960
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox Text_Password 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3960
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1785
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1200
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label_change_message_2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change message"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   9480
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label_change_message 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change message"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   9360
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label_value_message 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   600
      Left            =   1200
      TabIndex        =   18
      Top             =   1080
      Width           =   7920
   End
   Begin VB.Label Label_lbl_message 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MESSAGE:"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   240
      TabIndex        =   17
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label_double 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   270
      Left            =   8400
      TabIndex        =   16
      Top             =   3360
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label_Make_NOT_TRIAL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Make NOT TRIAL"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   270
      Left            =   5520
      TabIndex        =   15
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label_Make_TRIAL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Make TRIAL"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   270
      Left            =   3720
      TabIndex        =   14
      Top             =   3240
      Width           =   1485
   End
   Begin VB.Label Label_Change_allowed_minutes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change allowed minutes"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   270
      Left            =   4920
      TabIndex        =   13
      Top             =   2520
      Width           =   2835
   End
   Begin VB.Label Label_Change_password 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   270
      Left            =   6840
      TabIndex        =   12
      Top             =   1860
      Width           =   2085
   End
   Begin VB.Label Label_value_minutes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   8880
      TabIndex        =   11
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label_lb_minutes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   6480
      TabIndex        =   10
      Top             =   735
      Width           =   585
   End
   Begin VB.Label Label_value_password 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   4200
      TabIndex        =   9
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label_Lb_password 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   2760
      TabIndex        =   8
      Top             =   735
      Width           =   570
   End
   Begin VB.Label Label_warn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   690
   End
   Begin VB.Label Label_Minutes 
      BackStyle       =   0  'Transparent
      Caption         =   "ALLOWED MINUTES"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2535
      Width           =   3615
   End
   Begin VB.Label Label_Password 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1875
      Width           =   3615
   End
   Begin VB.Label Label_Second_file 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label4"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   10815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PROGRAM.EXE file"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form_Init"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const password_len_min As Long = 8

Dim old_caption                 As String
Dim bt                          As Label
Dim capt                        As String


Private Const my_caption        As String = "Trial maker"

Private Sub Form_Unload(Cancel As Integer)
  Erase sizes.full_matrix
End Sub

Private Sub Label_Change_allowed_minutes_Click()

  If Not minutes_correct Then Exit Sub
  If Not my_file_exist Then Exit Sub
  detect_trial_exe
  
  If is_trial_exe Then
'    get_text_file
    sizes.minutes = Val(Trim(Me.Text_Allowed_minutes.Text))
    
    set_new_text_file
  End If

  set_warn

End Sub

Private Sub Label_change_message_Click()
  If Not minutes_correct Then Exit Sub
  If Not my_file_exist Then Exit Sub
  
  detect_trial_exe
  
  If is_trial_exe Then
'    get_text_file
    sizes.message = Me.Label_value_message
    
    set_new_text_file
  End If

  set_warn
  
End Sub
Private Sub Label_change_message_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  label_mouse_move Me.Label_change_message
End Sub

Private Sub Label_Change_password_Click()
  If Not password_correct Then Exit Sub
  If Not my_file_exist Then Exit Sub
  detect_trial_exe

  Label_Password_Click
  
  If is_trial_exe Then
'    get_text_file
    sizes.password = Me.Text_Password.Text
    set_new_text_file
  End If

  set_warn
End Sub

Private Sub Label_lbl_message_Click()
  Form_Message.Show vbModal
  Unload Form_Message
  Set Form_Message = Nothing
  Me.Label_value_message = sizes.message
  
  Label_change_message_Click
End Sub

Private Sub Label_Make_TRIAL_Click()
Dim minutes         As String

  If Val(Me.Text_Allowed_minutes.Text) = 0 Then
    minutes = Get_Settings("Allowed_minutes", "3")
    Me.Text_Allowed_minutes.Text = minutes
  End If
  If Me.Text_Password.Text = "" Then
    Label_Password_Click
  End If
  If Not is_correct Then Exit Sub

  ut.make_trial_exe
  set_warn
  
  Save_Settings "Allowed_minutes", Text_Allowed_minutes.Text
  
  Me.Timer_SUCCES.Interval = 1500
  Set bt = Me.Label_Make_TRIAL
  capt = bt.caption
  bt.caption = "SUCCESS"
  Me.caption = "SUCCESS"
  Timer_SUCCES.Enabled = True
  
  Me.Label_double.Visible = False
  Me.Label_change_message_2.Visible = False

End Sub
Private Function is_correct() As Boolean
  If Not password_correct Then Exit Function
  If Not minutes_correct Then Exit Function
  If Not my_file_exist Then Exit Function
  
  is_correct = True
End Function
Private Function my_file_exist() As Boolean
  If fso Is Nothing Then Set fso = New FileSystemObject
  
  If Not fso.FileExists(Me.Label_Second_file.caption) Then
    MsgBox "The file 'PROGRAM.EXE' not exists."
    Exit Function
  End If
  my_file_exist = True
End Function
Private Function password_correct() As Boolean
Dim ms          As String
  If Len(Me.Text_Password) < password_len_min Then
    ms = "Password must have " & CStr(password_len_min) & " letters or more."
    ms = ms & vbNewLine & "For make new passowrd you can to key it or click on the label '" & Me.Label_Password.caption & "'"
    MsgBox ms
    Me.Text_Password.SetFocus
    Exit Function
  End If
  password_correct = True
End Function
Private Function minutes_correct() As Boolean
  If Val(Me.Text_Allowed_minutes.Text) <= 0 Then
    MsgBox "Correct, please the Allowed minutes"
    Me.Text_Allowed_minutes.SetFocus
    Exit Function
  End If
  minutes_correct = True
End Function
Private Sub Label_Make_NOT_TRIAL_Click()
  If Not my_file_exist Then Exit Sub

  make_not_trial
  set_warn
  
  Me.Timer_SUCCES.Interval = 1500
  Set bt = Me.Label_Make_NOT_TRIAL
  capt = bt.caption
  bt.caption = "SUCCESS"
  Me.caption = "SUCCESS"
  Timer_SUCCES.Enabled = True
  
  Me.Label_double.Visible = False
  Me.Label_change_message_2.Visible = False
  
End Sub

Private Sub Form_Load()
  Me.caption = my_caption
  Me.Timer_SUCCES.Enabled = False
  Me.AutoRedraw = True
  Me.PaintPicture Me.Picture, -100, 0, Me.Width + 100, Me.Height + 300
  
  Debug.Assert set_with_VB
  Picture_icon.PaintPicture Me.Picture, -100, 0, Me.Width + 100
  Set Picture_icon.Picture = Picture_icon.Image
  set_trial_exe_fin
  
  If with_VB Then
  Else
    Me.Picture_icon.Visible = False
  End If

  set_zero
End Sub

Private Sub Label_Password_Click()
  Me.Text_Password = get_password
End Sub
Private Sub set_zero()
  program_file_name = ""
  Label_warn.caption = ""
  Me.Text_Password.Locked = True
  Me.Text_Allowed_minutes.Locked = True
  Me.Text_Password.Text = ""
  Me.Text_Allowed_minutes.Text = ""
  
  Me.Label_Make_TRIAL.Enabled = False
  Me.Label_Make_NOT_TRIAL.Enabled = False
  Me.Label_Change_password.Enabled = False
  Me.Label_Change_allowed_minutes.Enabled = False
  Me.Label_change_message.Enabled = False
  
  Me.Label_Second_file.caption = "Click here for locate PROGRAM.EXE file"
  
  Me.Label_warn.caption = ""
  Me.Label_Lb_password.caption = ""
  Me.Label_value_password.caption = ""
  Me.Label_lb_minutes.caption = ""
  Me.Label_value_minutes.caption = ""
  Me.Label_value_message.caption = ""
  
End Sub
Private Sub Label_Second_file_Click()

  CommonDialog1.Filter = "Exe File (*.exe)|*.exe"
  Me.CommonDialog1.ShowOpen
  old_caption = Me.Label_Second_file
  Me.Label_Second_file = Me.CommonDialog1.FileName
  program_file_name = Me.CommonDialog1.FileName
  set_warn
End Sub
Private Sub set_warn()
Dim f_name        As String
  If fso Is Nothing Then Set fso = New FileSystemObject
  If Not fso.FileExists(program_file_name) Then
    set_zero
    Exit Sub
  End If
  detect_trial_exe
  f_name = fso.GetFile(program_file_name).Name
  If is_trial_exe Then
    get_text_file
    
    Me.Label_warn.caption = "IT IS TRIAL EXE"
    
    Me.Label_Lb_password.caption = "PASSWORD:"
    Me.Label_value_password.caption = sizes.password
    Label_value_password.Left = Label_Lb_password.Left + Label_Lb_password.Width + 100
    
    Me.Label_lb_minutes.caption = "ALLOWED MINUTES:"
    Me.Label_value_minutes.caption = sizes.minutes
    Label_value_minutes.Left = Label_lb_minutes.Left + Label_lb_minutes.Width + 100
    
    Me.Label_Password.caption = "Password in " & f_name
    Me.Label_Minutes.caption = "Allowed minutes in " & f_name
    
    Me.Text_Password.Text = sizes.password
    Me.Text_Allowed_minutes.Text = sizes.minutes
    
    Me.Label_Make_TRIAL.Enabled = False
    Me.Label_Make_NOT_TRIAL.Enabled = True
    
    Me.Label_Change_allowed_minutes.Enabled = True
    Me.Label_Change_password.Enabled = True
    Me.Label_change_message.Enabled = True
    
    Me.Label_value_message = sizes.message
  Else
    Me.Label_warn.caption = "IT IS NOT TRIAL EXE"
    Me.Label_Lb_password.caption = ""
    Me.Label_value_password.caption = ""
    Me.Label_lb_minutes.caption = ""
    Me.Label_value_minutes.caption = ""
    Me.Label_value_message.caption = ""
    
    Me.Label_Password.caption = "Password for " & f_name
    Me.Label_Minutes.caption = "Allowed minutes for " & f_name
    
    Me.Label_Make_TRIAL.Enabled = True
    Me.Label_Make_NOT_TRIAL.Enabled = False

    Me.Label_Change_allowed_minutes.Enabled = False
    Me.Label_Change_password.Enabled = False
    Me.Label_change_message.Enabled = False
  End If
  
  Me.Text_Allowed_minutes.Locked = False
'  Me.Text_Password.Locked = False

End Sub

Private Function get_password() As String
Dim nm        As Long
Const ln      As Long = 10
Dim num       As Long
Dim letter    As String
Dim psw       As String
Dim num_min   As Long
Dim num_Max   As Long

  Randomize Timer
  psw = ""
  num_min = Asc("A")
  num_Max = Asc("Z")
  
  For nm = 1 To ln
    num = Int(((num_Max - num_min) * Rnd) + num_min)
'    Debug.Print num, Chr(num)
    letter = Chr$(num)
    psw = psw & letter
    
  Next
  get_password = psw
End Function

Private Sub Timer_SUCCES_Timer()
  Timer_SUCCES.Enabled = False
'  MsgBox capt
  bt.caption = capt
  Me.caption = my_caption
  bt.Refresh
End Sub
Private Sub Label_Change_allowed_minutes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 label_mouse_move Me.Label_Change_allowed_minutes
End Sub
Private Sub Label_Change_password_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 label_mouse_move Me.Label_Change_password
End Sub
Private Sub Label_Make_NOT_TRIAL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 label_mouse_move Me.Label_Make_NOT_TRIAL
End Sub
Private Sub Label_Make_TRIAL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 label_mouse_move Me.Label_Make_TRIAL
End Sub
Private Sub label_mouse_move(lb As Label)
Dim lbl_2           As Label

  If lb Is Me.Label_change_message Then
    Set lbl_2 = Me.Label_change_message_2
  Else
    Set lbl_2 = Me.Label_double
  End If
  
  If Not lbl_2.Visible Then
    lbl_2.Left = lb.Left - Screen.TwipsPerPixelX
    lbl_2.Top = lb.Top - Screen.TwipsPerPixelY
    lbl_2.Width = lb.Width
    lbl_2.Height = lb.Height
    lbl_2.caption = lb.caption
    lbl_2.AutoSize = lb.AutoSize
    
    lbl_2.ForeColor = Color_Entre(lb.ForeColor, vbWhite, 90)
    Set lbl_2.Font = lb.Font
    lbl_2.ZOrder 1
    lbl_2.Visible = True
    lbl_2.Refresh
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.Label_double.Visible = False
  Me.Label_change_message_2.Visible = False
End Sub

