Attribute VB_Name = "sh_ut"
Option Explicit

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Dim poly(1 To 6)    As POINTAPI

Private Const RGN_AND = 1
Private Const RGN_COPY = 5
Private Const RGN_DIFF = 4
Private Const RGN_MAX = RGN_COPY
Private Const RGN_MIN = RGN_AND
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const ALTERNATE = 1
Private Const WINDING = 2

Private Type Rect
  Left        As Long
  Top         As Long
  Right       As Long
  Bottom      As Long
End Type

Private Const LOGPIXELSY = 90
Private Declare Function SetTextCharacterExtra Lib "gdi32" (ByVal hdc As Long, ByVal nCharExtra As Long) As Long

Private Const DEFAULT_CHARSET = 1
Private Const OUT_DEFAULT_PRECIS = 0
Private Const CLIP_DEFAULT_PRECIS = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const PROOF_QUALITY = 2
Private Const FW_DONTCARE = 0
Private Const FW_THIN = 100
Private Const FW_EXTRALIGHT = 200
Private Const FW_LIGHT = 300
Private Const FW_NORMAL = 400
Private Const FW_MEDIUM = 500
Private Const FW_SEMIBOLD = 600
Private Const FW_BOLD = 700
Private Const FW_EXTRABOLD = 800
Private Const FW_HEAVY = 900
Private Const OPAQUE = 2
Private Const TRANSPARENT = 1



Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetWindowRgn& Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long)

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As Rect, ByVal wFormat As Long) As Long
Private Declare Function TabbedTextOut Lib "user32" Alias "TabbedTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long, ByVal nTabPositions As Long, lpnTabStopPositions As Long, ByVal nTabOrigin As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long

Private Declare Function GetDC& Lib "user32" (ByVal hwnd As Long)
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Declare Function SetTextColor& Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long)
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long

'***************************************************************
Private Type t_txt
  lft               As Long
  Top               As Long
  fForeColor        As Long
  fName             As String
  fSize             As Single
  Grad              As Long
  caption           As String
End Type

Public fr               As Form

Public caption_height         As Long
Public caption_width          As Long

Public content_height         As Long
Public content_width          As Long
Public dif_gor                As Long

Public Sub make_frm()
  fr.AutoRedraw = True
  make_rgn False
'  make_caption fr.caption
End Sub
Private Sub make_caption(cap As String)
Dim ret               As Long
Dim clr               As Long
Dim my_fore_color     As Long
Dim txt               As t_txt


  Call SetTextCharacterExtra(fr.hdc, 3)
  SetBkMode fr.hdc, TRANSPARENT
  my_fore_color = 9474080
  my_fore_color = vbRed
  
  txt.fName = "Georgia"
  txt.fSize = 19
  txt.caption = cap
  
  txt.lft = 90
  txt.Top = 6
  
  DeleteObject SelectObject(fr.hdc, CreateMyFont(txt.fSize, txt.fName, FW_HEAVY, txt.Grad))
  
  txt.fForeColor = Color_Entre(my_fore_color, vbBlack, 30)
  print_text txt, 2, 2
  txt.fForeColor = Color_Entre(my_fore_color, vbBlack, 15)
  print_text txt, 1, 1
  txt.fForeColor = my_fore_color
  print_text txt, 0, 0

  fr.Refresh
End Sub
Private Sub print_text(txt As t_txt, dif_x As Long, dif_y As Long)
Dim ret         As Long
  SetTextColor& fr.hdc, txt.fForeColor
  ret = TextOut(fr.hdc, txt.lft + dif_x, txt.Top + dif_y, txt.caption, Len(txt.caption))
End Sub
Private Sub make_rgn(color_paint As Boolean)
Dim cap_height            As Long
Dim cap_width             As Long
Dim cont_height           As Long
Dim cont_width            As Long
Dim frm_rgn               As Long
Dim frm_rgn_inter         As Long
Dim rgn_lim               As Long
Dim br                    As Long
Dim color_bas             As Long
Dim color_fill            As Long


  rgn_lim = 15
  cap_height = 50
  cap_width = fr.ScaleX(fr.Width, fr.ScaleMode, vbPixels)
  
  cont_height = 170
  cont_width = cap_width - 2 * rgn_lim
  
  color_bas = 9474080
  color_bas = 16753194
  color_bas = 14852664
  color_bas = 15582595

  
  frm_rgn = make_rgn_window(cap_height, cap_width, cont_height, cont_width)
  SetWindowRgn fr.hwnd, frm_rgn, False
  DeleteObject frm_rgn
  
  frm_rgn = make_rgn_window(cap_height, cap_width, cont_height, cont_width)
  frm_rgn_inter = make_rgn_window(cap_height, cap_width, cont_height, cont_width)
  OffsetRgn frm_rgn_inter, 8, 6
  rgn_lim = CreateRectRgn(0, 0, 0, 0)
  CombineRgn rgn_lim, frm_rgn, frm_rgn_inter, RGN_DIFF

  br = CreateSolidBrush(color_bas)
  If color_paint Then
    FillRgn fr.hdc, frm_rgn, br
  End If
  DeleteObject br
  
  color_fill = Color_Entre(color_bas, vbWhite, 30)
  br = CreateSolidBrush(color_fill)
  If color_paint Then
    FillRgn fr.hdc, rgn_lim, br
  End If
  DeleteObject br
  
  OffsetRgn frm_rgn_inter, -16, -12
  CombineRgn rgn_lim, frm_rgn, frm_rgn_inter, RGN_DIFF
  color_fill = Color_Entre(color_bas, vbBlack, 20)
  br = CreateSolidBrush(color_fill)
  If color_paint Then
    FillRgn fr.hdc, rgn_lim, br
  End If
  DeleteObject br
  
  
  DeleteObject rgn_lim
  DeleteObject frm_rgn_inter
  DeleteObject frm_rgn
  DeleteObject br
  
End Sub
Private Function make_rgn_window(cap_height As Long, cap_width As Long, cont_height As Long, cont_width As Long) As Long
Dim capt_rgn        As Long
Dim cont_rgn        As Long
Dim frm_rgn         As Long
Dim ret             As Long
  
  dif_gor = (cap_width - cont_width) / 2
  
  frm_rgn = CreateRectRgn(0, 0, 0, 0)
  capt_rgn = CreateRoundRectRgn(0, 0, cap_width, 2 * cap_height, cap_height * 0.9, cap_width * 0.4)
  cont_rgn = CreateRoundRectRgn(dif_gor, cap_height - 2, dif_gor + cont_width, cont_height, cont_height * 0.4, cont_width * 0.4)

  CombineRgn frm_rgn, capt_rgn, cont_rgn, RGN_OR
  
  DeleteObject capt_rgn
  DeleteObject cont_rgn
  make_rgn_window = frm_rgn
  
End Function

Private Function CreateMyFont(nSize As Single, fName As String, Tens As Long, Grad As Long) As Long
  CreateMyFont = CreateFont(-MulDiv(nSize, GetDeviceCaps(GetDC(0), LOGPIXELSY), 72), 0, Grad * 10, 0, Tens, False, False, False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, fName)
End Function


