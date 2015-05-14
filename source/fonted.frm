VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EM FontEd"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fonted.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   224
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbGlyph 
      Height          =   255
      Left            =   1800
      Max             =   253
      TabIndex        =   9
      Top             =   600
      Value           =   213
      Width           =   1335
   End
   Begin VB.TextBox txtGlyph 
      Height          =   285
      Left            =   2640
      TabIndex        =   8
      Text            =   "213"
      Top             =   240
      Width           =   495
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   0
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   14
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   15
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.CommandButton cmdSaveGlyph 
      Caption         =   "[1] &Save this glyph"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.PictureBox guipal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2280
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox picCut 
      BackColor       =   &H00D69896&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   2880
      Left            =   1185
      ScaleHeight     =   2880
      ScaleWidth      =   375
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.HScrollBar hsbWidth 
      Height          =   255
      Left            =   120
      Max             =   8
      TabIndex        =   1
      Top             =   3000
      Value           =   6
      Width           =   1440
   End
   Begin Project1.GBATileEditor tedEdit 
      Height          =   2880
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   5080
      DotSize         =   12
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   3165
      Left            =   105
      Top             =   105
      Width           =   1470
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   615
      Left            =   1680
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "[2] Glyph #"
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1335
      Left            =   1680
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image imgBack 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnuRMB 
      Caption         =   "rmb"
      Visible         =   0   'False
      Begin VB.Menu mnuRMBColors 
         Caption         =   "[3] &Set theme..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TheFile As String
Private TheRomIndex As Long

Public MyTheme As Integer

Private Sub mnuRMBColors_Click()
  frmThemes.Show 1
End Sub

Private Sub imgBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  If Button = 2 Then PopupMenu mnuRMB
End Sub

Private Sub cmdSaveGlyph_Click()
  tedEdit.SaveTileData
  Open TheFile For Binary As #1
    Put #1, &H1E6594 + 1 + txtGlyph, CByte(hsbWidth.Value)
  Close #1
End Sub

Private Sub Form_Load()
  SetIcon Me.hwnd, "AAA", True
  Dim headr As String * 4
  InitDatabase
  If Command <> "" Then
    TheFile = Command
  Else
    frmRomSelect.Show 1
    If frmRomSelect.Tag = "Cancelled" Then End
    TheFile = frmRomSelect.Tag
    Unload frmRomSelect
  End If
  If TheFile = "" Then End
  Caption = Caption & " - " & TheFile
  
  On Error Resume Next
  MyTheme = Val(INIRead("elitemap", "Shared", "Theme"))
  If MyTheme = 0 Then MyTheme = 10
  imgBack.Picture = LoadResPicture(MyTheme, 0)
  'And now, we go through ALL controls available to recolor...
  guipal.Picture = LoadResPicture(MyTheme + 2, 0)
  Dim ColorRemap(1 To 2, 1 To 4) As Long
  Dim x As Integer, y As Integer
  For x = 1 To 2
    For y = 1 To 4
      ColorRemap(x, y) = guipal.Point((y - 1) * 8, (x - 1) * 8)
      guipal.PSet ((y - 1) * 8, (x - 1) * 8), vbRed
      'Debug.Print X & " x " & Y & " = " & Hex(ColorRemap(X, Y))
    Next y
  Next x
  Dim Ctl As Control
  For Each Ctl In Me.Controls
    For y = 1 To 4
      If Ctl.BorderColor = ColorRemap(1, y) Then Ctl.BorderColor = ColorRemap(2, y)
      If Ctl.BackColor = ColorRemap(1, y) Then Ctl.BackColor = ColorRemap(2, y)
    Next y
    If MyTheme = 40 And (TypeOf Ctl Is Label Or TypeOf Ctl Is CheckBox) Then Ctl.ForeColor = vbWhite
  Next Ctl
  If MyTheme = 40 Then
    For Each Ctl In Me.Controls
      If TypeOf Ctl Is Label Then Ctl.ForeColor = vbWhite
    Next
  End If
  For Each Ctl In Me.Controls
    If Left(Ctl.Caption, 1) = "[" Then
      i = Val(Mid(Ctl.Caption, 2, 4))
      Ctl.Caption = LoadResString(i)
    End If
  Next
  On Error GoTo 0
  
  Open TheFile For Binary As #1
    Get #1, &HAD, headr
  Close #1
  TheRomIndex = FindRom(headr)
  If TheRomIndex = -1 Then
    MsgBox "Unsupported ROM."
    End
  End If
  'Debug.Print Hex(Roms(TheRomIndex).FontGFX)
  'Debug.Print Hex(Roms(TheRomIndex).FontWidths)
  If Roms(TheRomIndex).FontGFX = 0 Then
    MsgBox LoadResString(100)
    End
  End If
  If Roms(TheRomIndex).FontWidths = 0 Then
    MsgBox LoadResString(101)
    End
  End If

  tedEdit.FileName = TheFile
  
  tedEdit.Colors(0) = RGB(254, 254, 254)
  tedEdit.Colors(14) = RGB(192, 192, 192)
  tedEdit.Colors(15) = RGB(64, 64, 64)
  
  optColor(0).BackColor = RGB(254, 254, 254)
  optColor(14).BackColor = RGB(192, 192, 192)
  optColor(15).BackColor = RGB(64, 64, 64)
  
  optColor(15).Value = True
  
  hsbGlyph_Change
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Close #1
  INIWrite "EliteMap", "Shared", "Theme", Str(MyTheme)
End Sub

Private Sub Form_Resize()
  imgBack.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub hsbGlyph_Change()
  txtGlyph = hsbGlyph.Value
  tedEdit.RomAddress = Roms(TheRomIndex).FontGFX + (txtGlyph * 64)
  tedEdit.LoadTileData

  Dim b As Byte
  Open TheFile For Binary As #1
    Get #1, &H1E6594 + 1 + txtGlyph, b
  Close #1
  hsbWidth = b
End Sub

Private Sub hsbGlyph_Scroll()
  hsbGlyph_Change
End Sub

Private Sub hsbWidth_Change()
  picCut.Left = 8 + (hsbWidth * 12)
  picCut.Width = 104 - picCut.Left
End Sub

Private Sub optColor_Click(Index As Integer)
  tedEdit.PenColor = Index
End Sub
