VERSION 5.00
Begin VB.Form Form_Message 
   Caption         =   "New Message"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6405
   Icon            =   "Form_Message.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6405
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text_Message 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   2055
      Left            =   255
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form_Message.frx":08CA
      Top             =   360
      Width           =   5895
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
      Left            =   4440
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label_OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OK"
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
      Left            =   3022
      TabIndex        =   1
      Top             =   2640
      Width           =   360
   End
End
Attribute VB_Name = "Form_Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Me.AutoRedraw = True
  Me.PaintPicture Form_Init.Picture, -100, 0, Me.Width + 100, Me.Height + 300
  Me.Text_Message.Text = sizes.message
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.Label_double.Visible = False
End Sub

Private Sub Label_OK_Click()
  sizes.message = Me.Text_Message.Text
  Hide
End Sub
Private Sub label_mouse_move(lb As Label)
Dim lbl_2           As Label

  Set lbl_2 = Me.Label_double
  If Not lbl_2.Visible Then
    lbl_2.Left = lb.Left - Screen.TwipsPerPixelX
    lbl_2.Top = lb.Top - Screen.TwipsPerPixelY
    lbl_2.Width = lb.Width
    lbl_2.Height = lb.Height
    lbl_2.caption = lb.caption
    lbl_2.ForeColor = Color_Entre(lb.ForeColor, vbWhite, 90)
    Set lbl_2.Font = lb.Font
    lbl_2.ZOrder 1
    lbl_2.Visible = True
  End If
End Sub

Private Sub Label_OK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  label_mouse_move Me.Label_OK
End Sub

Private Sub Text_Message_Change()
Const len_max          As Long = 1000

  If Len(Me.Text_Message.Text) > len_max Then
    Text_Message.Text = VBA.Left$(Text_Message.Text, len_max)
  End If
  
End Sub

Private Sub Text_Message_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Me.Label_double.Visible = False
End Sub
