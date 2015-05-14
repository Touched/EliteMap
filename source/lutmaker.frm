VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picTile2 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   6000
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   4680
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox picTileBox2 
      Height          =   1980
      Left            =   120
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   3
      Top             =   2280
      Width           =   4155
      Begin VB.PictureBox picTileset2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   15360
         Left            =   0
         MouseIcon       =   "lutmaker.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "lutmaker.frx":0152
         ScaleHeight     =   1024
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   257
         TabIndex        =   5
         Top             =   0
         Width           =   3855
         Begin VB.Shape shpTileset2 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            Height          =   255
            Left            =   0
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.VScrollBar vsbTileset2 
         Height          =   1920
         LargeChange     =   4
         Left            =   3840
         Max             =   56
         TabIndex        =   4
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picTileBox 
      Height          =   1980
      Left            =   120
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   0
      Top             =   120
      Width           =   4155
      Begin VB.VScrollBar vsbTileset 
         Height          =   1920
         LargeChange     =   4
         Left            =   3840
         Max             =   56
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picTileset 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   15360
         Left            =   0
         MouseIcon       =   "lutmaker.frx":C1194
         MousePointer    =   99  'Custom
         Picture         =   "lutmaker.frx":C12E6
         ScaleHeight     =   1024
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   257
         TabIndex        =   1
         Top             =   0
         Width           =   3855
         Begin VB.Shape shpTileset 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            Height          =   255
            Left            =   0
            Top             =   0
            Width           =   255
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "->"
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  mx = (x \ 16)
  my = y \ 16
  'Caption = Shift
  If x > 0 And y > 0 And x < (&H1000) And y < (&H4000) Then
    If Button = vbLeftButton Then
      ltile = (my * CLng(&H10)) + mx
      DrawTile2 ltile, picTile.hdc
      picTile.Refresh
    End If
  End If
End Sub

Private Sub picTileset_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  mx = (x \ 16)
  my = y \ 16
  If mx = 16 Then mx = 15
  shpTileset.Move mx * 16, my * 16
  shpTileset.Visible = True
End Sub

Private Sub vsbTileset_Change()
  On Error Resume Next
  picTileset.Move picTileset.Left, -(vsbTileset * &H10) '* &H40) ' &H80)
  picTileset.SetFocus
End Sub

Private Sub picTileset2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  mx = (x \ 16)
  my = y \ 16
  'Caption = Shift
  If x > 0 And y > 0 And x < (&H1000) And y < (&H4000) Then
    If Button = vbLeftButton Then
      rtile = (my * CLng(&H10)) + mx
      DrawTile2 rtile, picTile2.hdc
      picTile2.Refresh
    End If
  End If
End Sub

Private Sub picTileset2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  mx = (x \ 16)
  my = y \ 16
  If mx = 16 Then mx = 15
  shpTileset2.Move mx * 16, my * 16
  shpTileset2.Visible = True
End Sub

Private Sub vsbTileset2_Change()
  On Error Resume Next
  picTileset2.Move picTileset2.Left, -(vsbTileset2 * &H10) '* &H40) ' &H80)
  picTileset2.SetFocus
End Sub

Public Sub DrawTile(ByVal tileno, ByVal hdc)
  StretchBlt hdc, 0, 0, 32, 32, picTileset.hdc, (tileno Mod 16) * 16, (tileno \ 16) * 16, 16, 16, SRCCOPY
End Sub

Public Sub DrawTile2(ByVal tileno, ByVal hdc)
  StretchBlt hdc, 0, 0, 32, 32, picTileset2.hdc, (tileno Mod 16) * 16, (tileno \ 16) * 16, 16, 16, SRCCOPY
End Sub
