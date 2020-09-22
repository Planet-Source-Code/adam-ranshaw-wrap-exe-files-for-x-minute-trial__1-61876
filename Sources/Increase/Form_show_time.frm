VERSION 5.00
Begin VB.Form Form_show_time 
   BorderStyle     =   0  'None
   Caption         =   "Show time"
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   975
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command_exit 
      BackColor       =   &H000000C0&
      Caption         =   "Exit"
      Height          =   255
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label_red 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label_verde 
      BackColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   4935
   End
End
Attribute VB_Name = "Form_show_time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim he        As Long
Dim wi        As Long
Dim e_me      As Long

Private Sub Command_exit_Click()
  my_TerminateProcess
  Hide
  Unload Me
  my_end "Command_exit_Click"

End Sub

Private Sub Form_Activate()
  If ok_password Then
  Else
    Me.Left = 300
    Me.Top = 400
  End If
End Sub

Private Sub Form_Load()
  e_me = e_form.e_Form_show_time
  If ok_password Then
    Me.Left = 300
    Me.Top = -1000
  Else
    Me.Left = 300
    Me.Top = 400
  End If

  Me.AutoRedraw = True
  
  If file_exist(fnd_name) Then
    Set Me.Picture = LoadPicture(fnd_name)
    Me.PaintPicture Me.Picture, 0, 0, Me.Width + 1000
  Else
    Me.BackColor = vbBlack
  End If
  
  Me.Label_verde.Visible = False

  wi = 4000
  he = 240
  Me.Label_red.Left = 0
  Me.Label_verde.Left = 0
  
  Label_red.Top = 0
  Label_verde.Top = 0
  
  Label_red.Width = wi
  Label_verde.Width = wi
  
  Label_red.Height = he
  Label_verde.Height = he
  
  Me.Command_exit.Left = Me.Label_red.Left + Me.Label_red.Width
  Me.Command_exit.Top = 0
  Me.Command_exit.Height = he - 2 * Screen.TwipsPerPixelY
  
  Me.Height = he
  Me.Width = wi + Command_exit.Width
  my_SetWindowPos

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Processing_mouse_Move Acc_Mouse.down, X, Y, e_me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = VBRUN.MouseButtonConstants.vbLeftButton Then
    Processing_mouse_Move Acc_Mouse.Move, X, Y, e_me
  End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Processing_mouse_Move Acc_Mouse.Up, X, Y, e_me
End Sub

Private Sub Form_Terminate()
Command_exit_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
Command_exit_Click
End Sub

Private Sub Label_red_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Processing_mouse_Move Acc_Mouse.down, X, Y, e_me
End Sub
Private Sub Label_red_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = VBRUN.MouseButtonConstants.vbLeftButton Then
    Processing_mouse_Move Acc_Mouse.Move, X, Y, e_me
  End If
End Sub
Private Sub Label_red_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Processing_mouse_Move Acc_Mouse.Up, X, Y, e_me
End Sub
Private Sub Label_verde_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Processing_mouse_Move Acc_Mouse.down, X, Y, e_me
End Sub
Private Sub Label_verde_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = VBRUN.MouseButtonConstants.vbLeftButton Then
    Processing_mouse_Move Acc_Mouse.Move, X, Y, e_me
  End If
End Sub
Private Sub Label_verde_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Processing_mouse_Move Acc_Mouse.Up, X, Y, e_me
End Sub
