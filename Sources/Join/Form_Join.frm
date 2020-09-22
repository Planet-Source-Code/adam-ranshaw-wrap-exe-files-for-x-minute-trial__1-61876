VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_Join 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Join"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   Icon            =   "Form_Join.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command_make_destination 
      Caption         =   "Make TrialMaker.exe"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "SERVER.EXE file"
      Height          =   255
      Left            =   170
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "INCREASE.EXE file"
      Height          =   255
      Left            =   165
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label_First_file 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label3"
      Height          =   615
      Left            =   170
      TabIndex        =   1
      Top             =   600
      Width           =   8000
   End
   Begin VB.Label Label_Second_file 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label4"
      Height          =   615
      Left            =   165
      TabIndex        =   0
      Top             =   1800
      Width           =   7995
   End
End
Attribute VB_Name = "Form_Join"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cs_click              As String = "Click here for ubicate the file"

Private Const cs_first              As String = "First_file"
Private Const cs_second             As String = "Second_file"

Private Sub Command_make_destination_Click()

  ut.make_file_by_two_files Me.Label_First_file.Caption, Me.Label_Second_file.Caption
  If success Then MsgBox ("SUCCESS !")
End Sub

Private Sub Form_Load()
  Me.Label_First_file.Caption = Get_Settings(cs_first, cs_click)
  Me.Label_Second_file.Caption = Get_Settings(cs_second, cs_click)
  
End Sub

Private Sub Label_First_file_Click()
Dim file_name         As String

  CommonDialog1.Filter = "Server.Exe File (Server.exe)|Server.exe"
  Me.CommonDialog1.ShowOpen
  file_name = Me.CommonDialog1.FileName
  
  If file_name = "" Then
    file_name = cs_click
  End If
  
  Me.Label_First_file = file_name
  Save_Settings cs_first, Me.Label_First_file.Caption
  
End Sub

Private Sub Label_Second_file_Click()
Dim file_name         As String
  CommonDialog1.Filter = "Increase.Exe File (Increase.exe)|Increase.exe"
  Me.CommonDialog1.ShowOpen
  file_name = Me.CommonDialog1.FileName
  
  If file_name = "" Then
    file_name = cs_click
  End If

  Me.Label_Second_file = file_name
  Save_Settings cs_second, Me.Label_Second_file.Caption
End Sub
