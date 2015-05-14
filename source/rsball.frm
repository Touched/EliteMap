VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "RS Ball"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rsball.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   402
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   241
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picBar 
      BackColor       =   &H80000010&
      Height          =   5415
      Left            =   120
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   221
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CommandButton cmdPanel 
         Caption         =   "[4] Miscellaneous"
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   36
         Top             =   5040
         Width           =   3315
      End
      Begin VB.CommandButton cmdPanel 
         Caption         =   "[3] Bitmap Import"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   35
         Top             =   4560
         Width           =   3315
      End
      Begin VB.CommandButton cmdPanel 
         Caption         =   "[2] Pointer Management"
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   34
         Tag             =   "^_^"
         Top             =   2895
         Width           =   3315
      End
      Begin VB.CommandButton cmdPanel 
         Caption         =   "[1] General"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   33
         Tag             =   "^_^"
         Top             =   0
         WhatsThisHelpID =   1000
         Width           =   3315
      End
      Begin VB.PictureBox picPanel 
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   1
         Left            =   0
         ScaleHeight     =   1335
         ScaleWidth      =   3300
         TabIndex        =   10
         Top             =   3270
         Width           =   3300
         Begin VB.TextBox txtGFX 
            Height          =   285
            Left            =   1680
            TabIndex        =   13
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox txtPAL 
            Height          =   285
            Left            =   1680
            TabIndex        =   12
            Top             =   480
            Width           =   1455
         End
         Begin VB.CommandButton cmdRepoint 
            Caption         =   "[22] Apply Pointer Changes"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label Label1 
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   2760
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label3 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2760
            TabIndex        =   16
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "[20] GFX"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "[21] PAL"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   1455
         End
         Begin VB.Image imgPanBack 
            Height          =   960
            Index           =   1
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox picPanel 
         BorderStyle     =   0  'None
         Height          =   615
         Index           =   3
         Left            =   0
         ScaleHeight     =   615
         ScaleWidth      =   3300
         TabIndex        =   24
         Top             =   315
         Width           =   3300
         Begin VB.TextBox txtUnown 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            MaxLength       =   1
            TabIndex        =   26
            Text            =   "A"
            Top             =   120
            Width           =   495
         End
         Begin VB.CommandButton cmdUnown 
            Caption         =   "[40] ASCII to Unown"
            Height          =   375
            Left            =   720
            TabIndex        =   25
            Top             =   120
            Width           =   2415
         End
         Begin VB.Image imgPanBack 
            Height          =   960
            Index           =   3
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox picPanel 
         BorderStyle     =   0  'None
         Height          =   2535
         Index           =   0
         Left            =   0
         ScaleHeight     =   169
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   220
         TabIndex        =   1
         Top             =   360
         Width           =   3300
         Begin VB.CommandButton cmdExpNext 
            Caption         =   "4"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   12
               Charset         =   2
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   31
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdExport 
            Caption         =   "[12] Export"
            Height          =   375
            Left            =   480
            TabIndex        =   5
            Top             =   1080
            Width           =   1095
         End
         Begin VB.CommandButton cmdExpPrev 
            Caption         =   "3"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   12
               Charset         =   2
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   375
         End
         Begin VB.ComboBox txtRom 
            Height          =   315
            Left            =   960
            TabIndex        =   29
            Top             =   120
            Width           =   2175
         End
         Begin VB.ComboBox cboBank 
            Height          =   315
            ItemData        =   "rsball.frx":030A
            Left            =   960
            List            =   "rsball.frx":031A
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtBMP 
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Text            =   "poke.bmp"
            Top             =   1560
            Width           =   1815
         End
         Begin VB.TextBox txtImage 
            Height          =   285
            Left            =   2640
            TabIndex        =   6
            Text            =   "0"
            Top             =   600
            Width           =   495
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000010&
            Height          =   1020
            Left            =   2085
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   64
            TabIndex        =   4
            Top             =   1080
            Width           =   1020
         End
         Begin VB.TextBox txtBank 
            Height          =   285
            Left            =   2040
            TabIndex        =   3
            Text            =   "0"
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton cmdDumpAll 
            Caption         =   "[13] Dump GFX Bank"
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   1920
            Width           =   1815
         End
         Begin VB.FileListBox filRoms 
            Height          =   285
            Left            =   1080
            Pattern         =   "*.gba;*.agb"
            TabIndex        =   30
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "[11] Image"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "[10] ROM"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   855
         End
         Begin VB.Image imgPanBack 
            Height          =   960
            Index           =   0
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.PictureBox picPanel 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1335
         Index           =   2
         Left            =   0
         ScaleHeight     =   1335
         ScaleWidth      =   3300
         TabIndex        =   18
         Top             =   1560
         Width           =   3300
         Begin VB.TextBox txtNewBMP 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   21
            Top             =   120
            Width           =   1695
         End
         Begin VB.CommandButton cmdImport 
            Caption         =   "[32] Import bitmap"
            Enabled         =   0   'False
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   3015
         End
         Begin VB.TextBox txtNewGFX 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            TabIndex        =   19
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "[30] File name"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "[31] GFX location"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   1215
         End
         Begin VB.Image imgPanBack 
            Height          =   960
            Index           =   2
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   960
         End
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "RS Ball 2000 by Kyoufu Kawa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   5640
      Width           =   3375
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
         Caption         =   "&Set theme..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type bitmapheader
  FileHeader As String * 2
  FileSize As Long
  Reserved As Long
  BitmapOffset As Long
End Type

Private Type BITMAPINFOHEADER
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private palettes(0 To &HF) As Long
Private bitmap(0 To &H7FFF) As Byte
Private obitmap(0 To &H7FFF) As Byte

Private oldbitsize As Long
Private oldpalsize As Long

Private a As Long
Private b As Long
Private f As String
Private n As Long
Private pc As Integer

Private bitheader As bitmapheader
Private bitheader2 As BITMAPINFOHEADER

'Private WithEvents m_c As cSizeMoveHelper
Public mytheme As Integer

Private Sub RefoldAll()
  'th = chkPanel(0).Height + picPanel(0).Height
  For i = 0 To picPanel.UBound
    cmdPanel(i).Top = th
    th = th + cmdPanel(i).Height
    If cmdPanel(i).Tag = "^_^" Then
      picPanel(i).Visible = True
      picPanel(i).Top = th
      th = th + picPanel(i).Height
    Else
      picPanel(i).Visible = False
    End If
  Next i
End Sub

Private Sub cboBank_Click()
  txtBank = cboBank.ListIndex
  txtImage_LostFocus
End Sub

Private Sub chkPanel_Click(Index As Integer)
  RefoldAll
End Sub

Private Sub cmdExpNext_Click()
  txtImage = txtImage + 1
  txtImage_LostFocus
  cmdExport_Click
End Sub

Private Sub cmdExpPrev_Click()
  If txtImage > 0 Then txtImage = txtImage - 1
  txtImage_LostFocus
  cmdExport_Click
End Sub

Private Sub cmdPanel_Click(Index As Integer)
  If cmdPanel(Index).Tag = "" Then
    cmdPanel(Index).Tag = "^_^"
  Else
    cmdPanel(Index).Tag = ""
  End If
  RefoldAll
End Sub

Private Sub cmdUnown_Click()
  txtUnown = UCase(txtUnown)
  If txtUnown = "A" Then
    txtImage = 201
  ElseIf txtUnown = "!" Then
    txtImage = 438
  ElseIf txtUnown = "?" Then
    txtImage = 439
  Else
    txtImage = (Asc(txtUnown) - Asc("A")) + 412
  End If
  
  If cboBank.ListIndex = 0 Or cboBank.ListIndex = 2 Then
    cboBank.ListIndex = cboBank.ListIndex + 1
  End If
End Sub

Private Sub cmdExport_Click()
  Dim inbuffer(0 To 65535) As Byte
  Dim databuffer(0 To 65535) As Byte
  Dim longon As Long
  Dim osize As Long
  Dim abyte As Byte
  
  Dim headr As String * 4
  CheckLock txtRom
  Open txtRom For Binary As #256
  Get #256, &HAD, headr
  Close #256
  i = FindRom(headr)
  If i = -1 Then
    MsgBox LoadResString(100)
    Exit Sub
  End If
  'MsgBox headr & " " & i
  
  pc = 64
  
  Select Case txtBank
    Case 0
      If Roms(i).TrainerPics = 0 Then
        MsgBox LoadResString(101), vbInformation
        Exit Sub
      End If
      a = Roms(i).TrainerPics '&H1EC53C
      b = Roms(i).TrainerPals '&H1EC7D4
      'f = "trainer-"
      n = Roms(i).TrainerPicCount '83
    Case 1
      If Roms(i).MonsterPics = 0 Then
        MsgBox LoadResString(102), vbInformation
        Exit Sub
      End If
      a = Roms(i).MonsterPics '&H1E8354
      b = Roms(i).MonsterPals '&H1EA5B4
      'f = "pkmn-"
      n = Roms(i).MonsterPicCount '440
      If Left(Roms(i).Code, 3) = "BPE" Then pc = 64 * 2
    Case 2
      If Roms(i).TrainerBackPics = -1 Then
        MsgBox LoadResString(103), vbInformation
        Exit Sub
      End If
      If Roms(i).TrainerBackPics = 0 Then
        MsgBox LoadResString(104), vbInformation
        Exit Sub
      End If
      a = Roms(i).TrainerBackPics '&H1ECAE4
      b = Roms(i).TrainerBackPals '&H1ECAFC
      'f = "tback-"
      n = Roms(i).TrainerBackPicCount
      pc = 64 * 4
    Case 3
      If Roms(i).MonsterBackPics = 0 Then
        MsgBox LoadResString(105), vbInformation
        Exit Sub
      End If
      a = Roms(i).MonsterBackPics '&H1E97F4
      b = Roms(i).MonsterShinyPals '&H1EB374
      'f = "pback-"
      n = Roms(i).MonsterPicCount '440
  End Select
  
  Open txtRom For Binary As #1
  
  Get #1, a + (txtImage * 8) + 1, longon
  'Label1 = Hex(a + (txtImage * 8))
  bp = longon - &H8000000
  txtGFX = Hex(bp)
  txtNewGFX = Hex(bp)
  Get #1, b + (txtImage * 8) + 1, longon
  'Label3 = Hex(b + (txtImage * 8))
  pp = longon - &H8000000
  txtPAL = Hex(pp)
  
  Close #1
  
  LunarOpenFile txtRom, LC_READONLY
  LunarReadFile inbuffer(0), 16384, bp, LC_SEEK
  bsize = LZ77UnComp(inbuffer(), bitmap())
  oldbitsize = laststructsize
 
  LunarReadFile inbuffer(0), 16384, pp, LC_SEEK
  osize = LZ77UnComp(inbuffer(), databuffer())
  oldpalsize = laststructsize
  LunarCloseFile
  
  For i = 0 To &HF
    Y = databuffer((i * 2) + 1)
    X = databuffer((i * 2))
    c = (Y * &H100) + X
    'MsgBox Hex(c)
    c2 = (c \ &H400) Mod &H20
    c1 = (c \ &H20) Mod &H20
    c0 = c Mod &H20
    palettes(i) = RGB(c2 * 8, c1 * 8, c0 * 8)
  Next i
  'palettes(0) = RGB(112, 192, 160)
  
  For i = 0 To bsize - 1
    p4 = (((((bsize - 1) - i) \ &H100) * &H100))
    p3 = (((((bsize - 1) - i) \ &H20) * &H4)) Mod &H20
    p2 = ((i \ 4) * &H20) Mod &H100
    p1 = i Mod 4
    p = p1 + p2 + p3 + p4
    obitmap(i) = (bitmap(p) \ &H10) + ((bitmap(p) Mod &H10) * &H10)
  Next i
  
  Open txtBMP For Binary As #4
  fsz = &H76 + (&H800)
  Put #4, , CByte(&H42)
  Put #4, , CByte(&H4D)
  Put #4, , CLng(fsz)
  Put #4, , CLng(0)
  Put #4, , CLng(&H76)
  Put #4, , CLng(&H28)
  Put #4, , CLng(64)
  Put #4, , CLng(pc) 'CLng(128) 'Put #4, , CLng(64)
  Put #4, , CInt(1)
  Put #4, , CInt(4)
  Put #4, , CLng(0)
  Put #4, , (CLng(((&H1000) \ 4) * 4))
  Put #4, , CLng(0)
  Put #4, , CLng(0)
  Put #4, , CLng(16)
  Put #4, , CLng(16)
  For i = 0 To &HF
    Put #4, , palettes(i)
  Next i
  Put #4, , obitmap()
  Close #4
  
  Picture1.Picture = LoadPicture(txtBMP)
End Sub

Private Sub cmdDumpAll_Click()
  Dim headr As String * 4
  Open txtRom For Binary As #256
  Get #256, &HAD, headr
  Close #256
  i = FindRom(headr)
  If i = -1 Then
    MsgBox LoadResString(100)
    Exit Sub
  End If
  'MsgBox headr & " " & i
    
  If MsgBox(LoadResString(106), vbYesNo, LoadResString(107)) = vbNo Then Exit Sub
    
  On Error Resume Next
  MkDir "pics"
  On Error GoTo 0
    
  Select Case txtBank
    Case 0
      If Roms(i).TrainerPics = 0 Then
        MsgBox LoadResString(101), vbInformation
        Exit Sub
      End If
      a = Roms(i).TrainerPics '&H1EC53C
      b = Roms(i).TrainerPals '&H1EC7D4
      f = "pics\trainer-"
      n = Roms(i).TrainerPicCount '83
    Case 1
      If Roms(i).MonsterPics = 0 Then
        MsgBox LoadResString(102), vbInformation
        Exit Sub
      End If
      a = Roms(i).MonsterPics '&H1E8354
      b = Roms(i).MonsterPals '&H1EA5B4
      f = "pics\pkmn-"
      n = Roms(i).MonsterPicCount '440
    Case 2
      If Roms(i).TrainerBackPics = -1 Then
        MsgBox LoadResString(103), vbInformation
        Exit Sub
      End If
      If Roms(i).TrainerBackPics = 0 Then
        MsgBox LoadResString(104), vbInformation
        Exit Sub
      End If
      a = Roms(i).TrainerBackPics '&H1ECAE4
      b = Roms(i).TrainerBackPals '&H1ECAFC
      f = "pics\tback-"
      n = Roms(i).TrainerBackPicCount
    Case 3
      If Roms(i).MonsterBackPics = 0 Then
        MsgBox LoadResString(105), vbInformation
        Exit Sub
      End If
      a = Roms(i).MonsterBackPics '&H1E97F4
      b = Roms(i).MonsterShinyPals '&H1EB374
      f = "pics\pback-"
      n = Roms(i).MonsterPicCount '440
  End Select
  
  OldName = txtBMP
  
  MousePointer = 11
  For Each c In Form1.Controls
    c.Enabled = False
  Next
  
  For i = 0 To n - 1
    txtImage = i
    j = Hex(i)
    Do While Len(j) < 3
      j = "0" & j
    Loop
    txtBMP = f & j & ".bmp"
    DoEvents
    cmdExport_Click
  Next i
  
  txtBMP = OldName
  
  For Each c In Form1.Controls
    c.Enabled = True
  Next
  MousePointer = 0
End Sub

Private Sub cmdRepoint_Click()
  Dim longon As Long
  
  Dim headr As String * 4
  Open txtRom For Binary As #256
  Get #256, &HAD, headr
  Close #256
  i = FindRom(headr)
  If i = -1 Then
    MsgBox LoadResString(100)
    Exit Sub
  End If
  'MsgBox headr & " " & i
  
  Select Case txtBank
    Case 0
      If Roms(i).TrainerPics = 0 Then
        MsgBox LoadResString(101), vbInformation
        Exit Sub
      End If
      a = Roms(i).TrainerPics '&H1EC53C
      b = Roms(i).TrainerPals '&H1EC7D4
      f = "trainer-"
      n = Roms(i).TrainerPicCount '83
    Case 1
      If Roms(i).MonsterPics = 0 Then
        MsgBox LoadResString(102), vbInformation
        Exit Sub
      End If
      a = Roms(i).MonsterPics '&H1E8354
      b = Roms(i).MonsterPals '&H1EA5B4
      f = "pkmn-"
      n = Roms(i).MonsterPicCount '440
    Case 2
      If Roms(i).TrainerBackPics = -1 Then
        MsgBox LoadResString(103), vbInformation
        Exit Sub
      End If
      If Roms(i).TrainerBackPics = 0 Then
        MsgBox LoadResString(104), vbInformation
        Exit Sub
      End If
      a = Roms(i).TrainerBackPics '&H1ECAE4
      b = Roms(i).TrainerBackPals '&H1ECAFC
      f = "tback-"
      n = Roms(i).TrainerBackPicCount
    Case 3
      If Roms(i).MonsterBackPics = 0 Then
        MsgBox LoadResString(105), vbInformation
        Exit Sub
      End If
      a = Roms(i).MonsterBackPics '&H1E97F4
      b = Roms(i).MonsterShinyPals '&H1EB374
      f = "pback-"
      n = Roms(i).MonsterPicCount '440
  End Select

  Open txtRom For Binary As #1
  longon = &H8000000 + CLng(Val("&H" & txtGFX))
  Put #1, a + (txtImage * 8) + 1, longon
  longon = &H8000000 + CLng(Val("&H" & txtPAL))
  Put #1, b + (txtImage * 8) + 1, longon
  Close #1
  cmdExport_Click
End Sub

Private Sub cmdImport_Click()
'  Dim wite As Byte
'  Dim compbuffer(0 To &H7FFF) As Byte
'  Dim BlankMeg(0 To &HFFFFFF) As Byte
'
'  Dim headr As String * 4
'  CheckLock txtRom
'  Open txtRom For Binary As #256
'  Get #256, &HAD, headr
'  Close #256
'  i = FindRom(headr)
'  If i = -1 Then
'   MsgBox LoadResString(100)
'   Exit Sub
'  End If
'
'  Open txtNewBMP For Binary As #5
'  Get #5, , bitheader
'  If bitheader.FileHeader <> "BM" Then
'    MsgBox LoadResString(123)
'    GoTo endsub
'  End If
'  Get #5, , bitheader2
'  If bitheader2.biBitCount <> 4 Then
'    MsgBox LoadResString(108)
'    GoTo endsub
'  End If
'
'  'WARNING
'  Select Case txtBank
'    Case 0  'TrainerFront - Always 64 high
'      If bitheader2.biHeight <> 64 Or bitheader2.biWidth <> 64 Then
'        MsgBox LoadResString(109)
'        GoTo endsub
'      End If
'    Case 1  'PokeFront - 64 high for all but Emerald (128)
'      If Left(Roms(i).Code, 3) = "BPE" Then
'        If bitheader2.biHeight <> 128 Or bitheader2.biWidth <> 64 Then
'          MsgBox LoadResString(109)
'          GoTo endsub
'        End If
'      Else
'        If bitheader2.biHeight <> 64 Or bitheader2.biWidth <> 64 Then
'          MsgBox LoadResString(109)
'          GoTo endsub
'        End If
'      End If
'    Case 2  'TrainerBack - Always 256 high or uncompressed
'      If bitheader2.biHeight = 256 Then
'        MsgBox LoadResString(125)
'        GoTo endsub
'      End If
'    Case 3  'PokeBack - Always 64 high?
'      If bitheader2.biHeight <> 64 Or bitheader2.biWidth <> 64 Then
'        MsgBox LoadResString(109)
'        GoTo endsub
'      End If
'  End Select
'
'
'  If bitheader2.biCompression <> 0 Then
'    MsgBox LoadResString(110)
'    GoTo endsub
'  End If
'  If bitheader2.biPlanes <> 1 Then
'    MsgBox LoadResString(111)
'    GoTo endsub
'  End If
'  If bitheader2.biClrUsed > 16 Then
'    MsgBox LoadResString(112)
'    GoTo endsub
'  End If
'  For i = 0 To bitheader2.biClrUsed - 1
'    Get #5, , palettes(i)
'  Next i
'
'  Get #5, bitheader.BitmapOffset, wite
'  For i = 0 To bitheader2.biSizeImage - 1
'    Get #5, , obitmap(i)
'    p4 = (((((bitheader2.biSizeImage - 1) - i) \ &H100) * &H100))
'    p3 = (((((bitheader2.biSizeImage - 1) - i) \ &H20) * &H4)) Mod &H20
'    p2 = ((i \ 4) * &H20) Mod &H100
'    p1 = i Mod 4
'    p = p1 + p2 + p3 + p4
'    bitmap(p) = (obitmap(i) \ &H10) + ((obitmap(i) Mod &H10) * &H10)
'  Next i
'  Close #5
'
'  'Open "bitmap.bin" For Binary As #6
'  'For i = 0 To bitheader2.biSizeImage - 1
'  'Put #6, , bitmap(i)
'  'Next i
'  'Close #6
'
'  'Open "bitmap.biz" For Output As #7
'  'Print " "
'  'Close #7
'  'Shell "dcmp.exe bitmap.bin bitmap.biz", 6
'  'MsgBox "Press OK To Insert Bitmap into ROM"
'
'  newbitsize = LZ77Comp(bitheader2.biSizeImage, bitmap(), compbuffer())
'
'  'Open "bitmap.biz" For Binary As #7
'  'newbitsize = LOF(7)
'  xyz = Val("&H" & txtNewGFX)
'  If xyz <> 0 Then
'    newbitloc = xyz
'    MsgBox LoadResString(113) & Hex(newbitloc)
'    txtGFX = Hex(newbitloc)
'  ElseIf newbitsize > oldbitsize Then
'    MsgBox LoadResString(114) & vbCrLf & LoadResString(115) & Hex(oldbitsize) & vbCrLf & LoadResString(116) & newbitsize & vbCrLf & LoadResString(117) & (newbitsize - oldbitsize), vbExclamation
'    GoTo endsub
'    'LunarOpenFile txtROM, LC_READONLY
'    'If LunarGetFileSize < &H2000000 Then
'    '  LunarCloseFile
'    '  Open txtROM For Binary As #67
'    '  Put #67, &H1000001, BlankMeg()
'    '  Close #67
'    '  LunarOpenFile txtROM, LC_READONLY
'    '  MsgBox "Expanded ROM to 32MB"
'    'End If
'    '' MsgBox LunarGetFileSize
'    'newbitloc = LunarVerifyFreeSpace(&H1100000, &H1FFFFFF, newbitsize, LC_NOBANK)
'    'LunarCloseFile
'    'If newbitloc = 0 Then
'    '  MsgBox "No space left in ROM!"
'    '  GoTo endsub
'    'End If
'    'MsgBox "Moving bitmap to new location: " & Hex(newbitloc)
'    'txtGFX = Hex(newbitloc)
'  Else
'    newbitloc = Val("&H" & txtGFX)
'  End If
'
'  Open txtRom For Binary As #256
'  For i = 0 To newbitsize - 1
'    'Get #7, i + 1, wite
'    Put #256, newbitloc + i + 1, compbuffer(i)
'  Next i
'  Close #256
'  'Close #7
'
'  cmdRepoint_Click
'  MsgBox LoadResString(118) & vbCrLf & LoadResString(119) & " 0x" & Hex(newbitsize)
'endsub:
'  Close #5
'  'Close #7
End Sub

Private Sub Form_Load()
  SetIcon Me.hwnd, "AAA", True
  InitDatabase
  
  On Error Resume Next
  mytheme = Val(INIRead("elitemap", "Shared", "Theme"))
  If mytheme = 0 Then mytheme = 10
  imgBack.Picture = LoadResPicture(mytheme, 0)
  For i = 0 To imgPanBack.UBound
    imgPanBack(i).Picture = LoadResPicture(mytheme + 1, 0)
    imgPanBack(i).Move 0, 0, picPanel(i).ScaleWidth, picPanel(i).ScaleHeight
  Next i
  If mytheme = 40 Then
    For Each ctl In Me.Controls
      If TypeOf ctl Is Label Then ctl.ForeColor = vbWhite
    Next
  End If
  On Error GoTo 0

  If filRoms.ListCount = 0 Then
    MsgBox LoadResString(120)
    End
  End If
  
  On Error Resume Next
  For Each ctl In Me.Controls
    If Left(ctl.Caption, 1) = "[" Then
      i = Val(Mid(ctl.Caption, 2, 4))
      ctl.Caption = LoadResString(i)
    End If
  Next
  For i = 0 To 3
    cboBank.List(i) = LoadResString(14 + i)
  Next i
  On Error GoTo 0
  
  For i = 0 To filRoms.ListCount - 1
    txtRom.AddItem filRoms.List(i)
  Next i
  If Command <> "" Then
    txtRom.Text = Command
  Else
    txtRom.ListIndex = 0
  End If
  
  For i = 1 To Len(Label2)
    j = j + Asc(Mid(Label2, i, 1))
  Next i
  
  For i = 0 To picPanel.UBound
    picPanel(i).Tag = picPanel(i).Height
  Next i
  RefoldAll
  
  Tag = Width
  cboBank.ListIndex = 1

  If Int(j / 2) <> 1074 Then
    MsgBox LoadResString(121), vbCritical, LoadResString(122)
    End
  End If

  '--RESET--
  'Set m_c = New cSizeMoveHelper
  'm_c.Attach Me.hwnd
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  INIWrite "EliteMap", "Shared", "Theme", Str(mytheme)
End Sub

Private Sub Form_Resize()
  'Debug.Print Width, Height
  If Height < 3960 Then Height = 3960
  Width = 3735
  Label2.Top = ScaleHeight - 8 - 16
  picBar.Height = ScaleHeight - 8 - 16 - 8 - 8
  imgBack.Move 0, 0, ScaleWidth, ScaleHeight + 32
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'm_c.Detach
End Sub

Private Sub imgBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuRMB
End Sub

'Private Sub m_c_Sizing(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)
'  If lHeight < 178 Then lHeight = 178
'  lWidth = Val(Tag) / Screen.TwipsPerPixelY
'  Label2.Top = ScaleHeight - 8 - 16
'  picBar.Height = ScaleHeight - 8 - 16 - 8 - 8
'  imgBack.Move 0, 0, ScaleWidth, ScaleHeight + 32
'End Sub

Private Sub mnuRMBColors_Click()
  frmThemes.Show 1
End Sub

Private Sub picPanel_Paint(Index As Integer)
  picPanel(Index).Cls
  picPanel(Index).Line (0, 0)-(picPanel(Index).ScaleWidth, 0), &H80000014
  picPanel(Index).Line (0, 0)-(0, picPanel(Index).ScaleHeight), &H80000014
  picPanel(Index).Line (picPanel(Index).ScaleWidth - 1, 0)-(picPanel(Index).ScaleWidth - 1, picPanel(Index).ScaleHeight - 1), &H80000010
  picPanel(Index).Line (0, picPanel(Index).ScaleHeight - 1)-(picPanel(Index).ScaleWidth, picPanel(Index).ScaleHeight - 1), &H80000010
End Sub

Private Sub txtImage_LostFocus()
  On Error GoTo Hell
  Dim headr As String * 4
  Open txtRom For Binary As #256
  Get #256, &HAD, headr
  Close #256
  i = FindRom(headr)
  If i = -1 Then
    MsgBox "Unsupported rom."
    Exit Sub
  End If
  Select Case txtBank
    Case 0: n = Roms(i).TrainerPicCount
    Case 1: n = Roms(i).MonsterPicCount
    Case 2: n = Roms(i).TrainerBackPicCount
    Case 3: n = Roms(i).MonsterPicCount
  End Select
  If Val(txtImage) > n Then txtImage = n
Hell:
  Exit Sub
End Sub
