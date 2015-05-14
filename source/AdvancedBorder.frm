VERSION 5.00
Begin VB.Form AdvancedBorder 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Border Edit"
   ClientHeight    =   465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   690
   Icon            =   "AdvancedBorder.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   465
   ScaleWidth      =   690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBorder 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      Height          =   280
      Left            =   120
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   120
      Width           =   280
   End
End
Attribute VB_Name = "AdvancedBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10

Private Sub AdvancedBorder_Resize()
  If border_sizeX > 16 Then GoTo too_big_unload
  If border_sizeY > 16 Then GoTo too_big_unload
  picBorder.Width = 280 + (border_sizeX - 1) * 245
  picBorder.Height = 280 + (border_sizeY - 1) * 245
  AdvancedBorder.Width = picBorder.Left + picBorder.Width + 375
  If AdvancedBorder.Width < 1290 Then AdvancedBorder.Width = 1290
  AdvancedBorder.Height = picBorder.Top + picBorder.Height + 428
  Exit Sub
too_big_unload:
  MsgBox "Bordersize is too big!", vbCritical + vbOKOnly, "Error…"
  Unload Me
End Sub

Private Sub Form_Load()
  SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 5, 5, SWP_NOSIZE Or SWP_NOMOVE
  AdvancedBorder_Resize
End Sub

Private Sub picBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mx = (X \ 16)
  my = Y \ 16
  If Shift = 1 Then GoTo shift_event1
  If X > &H0 And Y > &H0 And X < picBorder.ScaleWidth + 1 And Y < picBorder.ScaleHeight + 1 Then
  dirty2 = True
    If Button = vbLeftButton Then
      Borderitems(my * border_sizeX + mx) = seltile(0)
      Form1.drawtilehdc seltile(0), picBorder.hdc, mx, my
      picBorder.Refresh
    ElseIf Button = vbRightButton Then
      Borderitems(my * border_sizeX + mx) = seltile(1)
      Form1.drawtilehdc seltile(1), picBorder.hdc, mx, my
      picBorder.Refresh
    ElseIf Button = vbMiddleButton Then
      Borderitems(my * border_sizeX + mx) = seltile(2)
      Form1.drawtilehdc seltile(2), picBorder.hdc, mx, my
      picBorder.Refresh
    End If
  End If
  Exit Sub
shift_event1:
  mx = (X) \ 16
  my = (Y) \ 16
  If X > &H0 And Y > &H0 And X < picBorder.ScaleWidth + 1 And Y < picBorder.ScaleHeight + 1 Then
    If Button = vbLeftButton Then
      seltile(0) = Borderitems(my * border_sizeX + mx) Mod &H400
      Form1.drawtilehdc seltile(0), Form1.picSel(0).hdc, 0, 0
      Form1.picSel(0).Refresh
    ElseIf Button = vbRightButton Then
      seltile(1) = Borderitems(my * border_sizeX + mx) Mod &H400
      Form1.drawtilehdc seltile(1), Form1.picSel(1).hdc, 0, 0
      Form1.picSel(1).Refresh
    ElseIf Button = vbMiddleButton Then
      seltile(2) = Borderitems(my * border_sizeX + mx) Mod &H400
      Form1.drawtilehdc seltile(2), Form1.picSel(2).hdc, 0, 0
      Form1.picSel(2).Refresh
    End If
  End If
End Sub
