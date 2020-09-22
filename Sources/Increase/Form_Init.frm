VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_Init 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Protected EXE File"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   ClipControls    =   0   'False
   Icon            =   "Form_Init.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command_Show_sizes 
      Caption         =   "Show sizes"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command_Execute 
      Caption         =   "Execute"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command_Ok_Password 
      Caption         =   "Verify  Password"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   8280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text_Password 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6720
      TabIndex        =   2
      Top             =   5880
      Width           =   2535
   End
   Begin VB.CommandButton Command_put_exe_copy 
      Caption         =   "Put exe copy"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.PictureBox Picture_Icon 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   10320
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer_proc 
      Left            =   6360
      Top             =   8520
   End
   Begin VB.Timer Timer_Password 
      Left            =   6360
      Top             =   8040
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1920
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label label_caption_2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "If you have paid the registration fee please enter your serial code:"
      Height          =   495
      Left            =   6720
      TabIndex        =   23
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label_caption 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "EXE Name"
      BeginProperty Font 
         Name            =   "Levenim MT"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   5400
      TabIndex        =   21
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label_Execute 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Play Free Trial"
      BeginProperty Font 
         Name            =   "Aharoni"
         Size            =   24
         Charset         =   177
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   6120
      TabIndex        =   20
      Top             =   2760
      Width           =   3735
   End
   Begin VB.Label Label_Exit 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Width           =   375
   End
   Begin VB.Label Label_verify_password 
      BackStyle       =   0  'Transparent
      Caption         =   "Verify Serial"
      BeginProperty Font 
         Name            =   "Aharoni"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label_double 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Execute"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   8880
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label_Message 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4395
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   5325
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Protected EXE Name:"
      BeginProperty Font 
         Name            =   "Levenim MT"
         Size            =   24
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   120
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Height          =   975
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      Top             =   3720
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Upgrade to Full"
      BeginProperty Font 
         Name            =   "Aharoni"
         Size            =   24
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   6000
      TabIndex        =   13
      Top             =   4560
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0000C000&
      BorderWidth     =   2
      Height          =   975
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form_Init.frx":08CA
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   615
      Left            =   240
      TabIndex        =   12
      Top             =   720
      Width           =   10215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form_Init.frx":096E
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   9855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Software Description:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Protected by Clock Lock from Adranix"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   7200
      Width           =   3615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form_Init.frx":0A1E
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7560
      Width           =   10695
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "www.ADRANIX.co.uk"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   6
      Top             =   7080
      Width           =   3615
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFF80&
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   7080
      Width           =   10815
   End
   Begin VB.Shape Shape3 
      Height          =   1335
      Left            =   6480
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6480
      TabIndex        =   22
      Top             =   5280
      Width           =   3015
   End
End
Attribute VB_Name = "Form_Init"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim e_me                As Long

Dim set_caption         As Boolean

Private Sub Save(Key As String, Value As String)
    SaveSetting Label_caption, "Registered", Key, Trim(Value)
End Sub

Private Sub StoreInfo()
Save Label_caption, (Text_Password.Text)
End Sub
Private Function Load(Key As String)
    Load = GetSetting(Label_caption, "Registered", Key, "0")
End Function
Private Sub LoadInfo()
Text_Password.Text = Load(Label_caption)
End Sub

Private Sub Command_Execute_Click()
  execute_program_exe_copy
End Sub


Private Sub Command_put_exe_copy_Click()
  put_program_exe_copy
End Sub

Private Sub Command_Show_sizes_Click()
  Me.Print "caption = " & sizes.caption
  Me.Print "program_size = " & sizes.program_size
  Me.Print "total_size = " & sizes.total_size
  Me.Print "password = " & sizes.password
  Me.Print "minutes = " & sizes.minutes
End Sub

Private Sub Form_Load()
LoadInfo
veryfi_password Trim(Me.Text_Password.Text)
If Not ok_password Then
'Do Nothing
Else
Command_Execute_Click
End If

  e_me = e_form.e_Form_Init

  If VB.App.PrevInstance Then End
  
  Debug.Assert set_with_VB
  
  Picture_Icon.Top = 45
  Picture_Icon.Left = 270
  
  Me.Timer_proc.Enabled = False
  Me.Timer_Password.Enabled = False
  
  Me.Text_Password.Text = ""
  Me.AutoRedraw = True
  
  fnd_name = App.path & "/" & bkgr_name

  get_text_file
  Label4.caption = "This use is limited to " & sizes.minutes & " minutes, after this time the software will close.  By upgrading to the full version you can use it for an unlimated ammount of minutes."
  Me.Label_caption.caption = sizes.caption
  Me.Label_Message = sizes.message
  
  set_form_icon
  If Form_Init.Picture_Icon.Picture = 0 Then
    Form_Init.Picture_Icon.Visible = False
  End If
  
  set_caption = True
  label_mouse_move Me.Label_caption
  set_caption = False
  Text_Password.Text = ""

'  If file_exist(fnd_name) Then
'    Set Me.Picture = LoadPicture(fnd_name)
'    Me.PaintPicture Me.Picture, 0, 0, Me.Width + 3000
'  Else
'    Me.BackColor = vbBlack
'    Picture_Icon.BackColor = Me.BackColor
'  End If
'  Set sh_ut.fr = Me
'  sh_ut.make_frm

End Sub


Private Sub Form_Resize()
Dim x_centr           As Long
Static my_exit        As Boolean

Const Width_min       As Long = 5000
Const Height_min      As Long = 3000


  If my_exit Then Exit Sub
  If Me.Width < Width_min And Me.WindowState <> vbMinimized Then
    my_exit = True
    Me.Width = Width_min
    my_exit = False
    DoEvents
  End If
  If Me.Height < Height_min And Me.WindowState <> vbMinimized Then
    my_exit = True
    Me.Height = Height_min
    my_exit = False
    DoEvents
  End If
  
  If file_exist(fnd_name) Then
    Set Me.Picture = LoadPicture(fnd_name)
    Me.PaintPicture Me.Picture, 0, 0, Me.Width + 3000, Me.Height + 3000
  Else
    
    Picture_Icon.BackColor = Me.BackColor
  End If

  
  x_centr = Me.Width / 2
  Label_verify_password.Left = Text_Password.Left
  Me.Label_double.Visible = False
End Sub
Private Sub make_centr(lb As Label)
  lb.Left = (Me.Width - lb.Width) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Me.Timer_proc.Enabled = False
End Sub

Private Sub Label_Execute_Click()
  Command_Execute_Click
End Sub

Private Sub Label_Execute_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  label_mouse_move Label_Execute
End Sub
Private Sub label_mouse_move(lb As Label)
Dim lbl_2           As Label
Dim proc            As Long

  If set_caption Then
    Set lbl_2 = Me.label_caption_2
  Else
    Set lbl_2 = Me.Label_double
  End If
  If Not lbl_2.Visible Then
    If lb Is Me.Label_Exit Then
      lbl_2.Left = lb.Left - 5 * Screen.TwipsPerPixelX
      lbl_2.Top = lb.Top - 2 * Screen.TwipsPerPixelY
      proc = 70
    Else
      lbl_2.Left = lb.Left - Screen.TwipsPerPixelX
      lbl_2.Top = lb.Top - Screen.TwipsPerPixelY
      proc = 90
    End If
    lbl_2.Width = lb.Width
    lbl_2.Height = lb.Height
    lbl_2.caption = lb.caption
    lbl_2.ForeColor = Color_Entre(lb.ForeColor, vbWhite, proc)
    Set lbl_2.Font = lb.Font
    lbl_2.ZOrder 1
    lbl_2.Visible = True
  End If
End Sub
Private Sub Label_Exit_Click()
  Hide
  Unload Me
  my_end "Label_Exit_Click"
End Sub

Private Sub Label_Exit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  label_mouse_move Me.Label_Exit
End Sub



Private Sub Label_verify_password_Click()
  veryfi_password Trim(Me.Text_Password.Text)
  If Not ok_password Then
    Me.Text_Password.PasswordChar = ""
    Me.Text_Password.Text = "NOT CORRECT"
    Me.Text_Password.Locked = True
    Me.Text_Password.ForeColor = vbRed
    Me.Timer_Password.Interval = 3000
    Me.Timer_Password.Enabled = True
  Else
  StoreInfo
    Command_Execute_Click
  End If

End Sub

Private Sub Label_verify_password_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  label_mouse_move Me.Label_verify_password
End Sub

Private Sub Label3_Click()
MsgBox "To buy the full version from the software author please read the software description or if you already have a Serial Code from the software author you can enter it in the box below below."
End Sub

Private Sub Text_Password_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    Label_verify_password_Click
  End If
End Sub

Private Sub Timer_Password_Timer()
  Me.Timer_Password.Enabled = False
  Me.Text_Password.Locked = False
  Me.Text_Password.Text = ""
  Me.Text_Password.ForeColor = vbBlack

  
  On Error Resume Next
  Me.Text_Password.SetFocus
  Err.Clear
  
End Sub

Private Sub Timer_proc_Timer()
  my_Timer_proc_Timer
End Sub


Private Sub Label_caption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Processing_mouse_Move Acc_Mouse.down, X, Y, e_me
End Sub
Private Sub Label_caption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label_double.Visible = False
  If Button = VBRUN.MouseButtonConstants.vbLeftButton Then
    Processing_mouse_Move Acc_Mouse.Move, X, Y, e_me
  End If
End Sub
Private Sub Label_caption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Processing_mouse_Move Acc_Mouse.Up, X, Y, e_me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Processing_mouse_Move Acc_Mouse.down, X, Y, e_me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label_double.Visible = False
  If Button = VBRUN.MouseButtonConstants.vbLeftButton Then
    Processing_mouse_Move Acc_Mouse.Move, X, Y, e_me
  End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Processing_mouse_Move Acc_Mouse.Up, X, Y, e_me
End Sub

Private Sub Timer1_Timer()
LoadInfo
veryfi_password Trim(Me.Text_Password.Text)
If Not ok_password Then
'Do Nothing
Label_verify_password.Enabled = True
Label_Execute.Enabled = True
Timer1.Enabled = False
Else
Command_Execute_Click
Timer1.Enabled = False
End If
End Sub
