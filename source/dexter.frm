VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dexter"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "dexter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   489
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   21
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame fraSizeComp 
      Caption         =   "&Size Comparison"
      Height          =   1215
      Left            =   2160
      TabIndex        =   7
      Top             =   1200
      Width           =   3735
      Begin VB.TextBox txtTrnrOff 
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtTrnrScale 
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txtPkmnOff 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtPkmnScale 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "against"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblOffset 
         Caption         =   "Offset"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblScale 
         Caption         =   "Scale"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraDexText 
      Caption         =   "&Description"
      Height          =   1695
      Left            =   2160
      TabIndex        =   15
      Top             =   2520
      Width           =   5055
      Begin VB.CommandButton cmdDexRep 
         Caption         =   "&Repoint"
         Height          =   285
         Left            =   3240
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDex 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtDexPtr2 
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDexPtr1 
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblDexPtr 
         Caption         =   "&Pointers"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtWeight 
      Height          =   285
      Left            =   3240
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtType 
      Height          =   285
      Left            =   3240
      MaxLength       =   11
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox lstSpecies 
      Height          =   4095
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblWeight 
      Caption         =   "&Weight"
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblHeight 
      Caption         =   "&Height"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblType 
      Caption         =   "&Type"
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuFind 
         Caption         =   "&Find in descriptions..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tDexInfo
  sPkmnType As String * 12
  iHeight As Integer
  iWeight As Integer
  pPokeDex1 As Long
  pPokeDex2 As Long
  iUnused1 As Integer
  iPkmnSize As Integer
  iPkmnVOffset As Integer
  iTrnrSize As Integer
  iTrnrVOffset As Integer
  iUnused2 As Integer
End Type

Dim DexInfo(512) As tDexInfo
Dim RomData As Integer
Dim sel As Long

Private Sub Form_Load()
  Dim i As Long
  Dim TheFile As String
  Dim headr As String * 4
  'Dim pokename As String * 11
  'Dim b As String
  
  SetIcon Me.hwnd, "AAA", True
  
  InitDatabase
  
  If Command <> "" Then
    TheFile = Command
  Else
    'TheFile = InputBox("Enter file name", , "Ruby.gba")
    frmRomSelect.Show 1
    If frmRomSelect.Tag = "Cancelled" Then End
    TheFile = frmRomSelect.Tag
    Unload frmRomSelect
  End If
  If TheFile = "" Then End
  Caption = Caption & " - " & TheFile
  
  Open TheFile For Binary As #1
  Get #1, &HAD, headr
  RomData = FindRom(headr)
  If RomData = -1 Then
    MsgBox "Unsupported ROM."
    End
  End If
  If Roms(RomData).MonsterDexData = 0 Then
    MsgBox "Supported ROM, but PokéDex base is unspecified."
    End
  End If
  
  ''Pokémon name loading code risen from the ashes of TrainEd -_-
  'Seek #1, Roms(RomData).MonsterNames + 1
  'Get #1, , i
  'Seek #1, i - &H8000000 + 1
  'For i = 0 To 410
  '  Get #1, , pokename
  '  b = Sapp2Asc(pokename)
  '  While InStr(1, b, "\x"): b = Left(b, Len(b$) - 1): Wend
  '  b = Left(b, Len(b) - 1)
  '  If b <> "?" Then lstSpecies.AddItem b
  'Next i
  
  Seek #1, Roms(RomData).MonsterDexData + 1 '&H3B1858 + 1
  For i = 0 To &H182
    Get #1, , DexInfo(i)
    lstSpecies.AddItem Right("000" & Hex(i), 3) & ". " & Replace(Sapp2Asc(DexInfo(i).sPkmnType, IIf(Roms(RomData).RomType = 1, True, False)), "\x", "")
  Next i
End Sub

Private Sub cmdSave_Click()
  Dim i As Integer
  Seek #1, Roms(RomData).MonsterDexData + 1 '&H3B1858 + 1
  For i = 0 To &H182
    Put #1, , DexInfo(i)
  Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Close #1
End Sub

Private Sub lstSpecies_Click()
  txtType = Trim(Replace(Replace(Sapp2Asc(DexInfo(lstSpecies.ListIndex).sPkmnType), "\x", ""), "î", ""))
  txtHeight = DexInfo(lstSpecies.ListIndex).iHeight
  txtWeight = DexInfo(lstSpecies.ListIndex).iWeight
  
  'TODO - Convert scale bytes to percentages
  txtPkmnScale = DexInfo(lstSpecies.ListIndex).iPkmnSize
  txtPkmnOff = DexInfo(lstSpecies.ListIndex).iPkmnVOffset
  txtTrnrScale = DexInfo(lstSpecies.ListIndex).iTrnrSize
  txtTrnrOff = DexInfo(lstSpecies.ListIndex).iTrnrVOffset
  
  txtDexPtr1 = "&H" & Hex(DexInfo(lstSpecies.ListIndex).pPokeDex1 - &H8000000)
  txtDexPtr2 = "&H" & Hex(DexInfo(lstSpecies.ListIndex).pPokeDex2 - &H8000000)
  UpdateDexSample
  
  sel = lstSpecies.ListIndex
End Sub

Private Sub lstSpecies_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuFind_Click()
  Dim ss As String
  Dim i As Integer
  On Error Resume Next
  ss = UCase(InputBox("Search the PokéDex descriptions for the following string:"))
  For i = lstSpecies.ListIndex + 1 To lstSpecies.ListCount - 1
    lstSpecies.ListIndex = i
    DoEvents
    If InStr(txtDex, ss) Then
      txtDex.SetFocus
      txtDex.SelStart = InStr(txtDex, ss) - 1
      txtDex.SelLength = InStr(txtDex.SelStart + 1, txtDex, " ") - txtDex.SelStart - 1
      Exit Sub
    End If
  Next i
  MsgBox "Not found."
End Sub

'------------------------------------------------------'

Private Sub txtType_LostFocus()
  DexInfo(sel).sPkmnType = Asc2Sapp(UCase(Replace(txtType, "î", "")) & "\x") & String(20, Chr$(0))
  lstSpecies.List(sel) = Right("000" & Hex(sel), 3) & ". " & Replace(Replace(Sapp2Asc(DexInfo(sel).sPkmnType), "î", ""), "\x", "")
End Sub

Private Sub txtHeight_LostFocus()
  DexInfo(sel).iHeight = Val(txtHeight)
End Sub

Private Sub txtWeight_LostFocus()
  DexInfo(sel).iWeight = Val(txtWeight)
End Sub

'------------------------------------------------------'

Private Sub cmdDexRep_Click()
  DexInfo(sel).pPokeDex1 = Val(txtDexPtr1) + &H8000000
  DexInfo(sel).pPokeDex2 = Val(txtDexPtr2) + &H8000000
  UpdateDexSample
End Sub

Private Sub UpdateDexSample()
  Dim s As String * 256
  Dim t As String
  Get #1, txtDexPtr1 + 1, s
  t = Sapp2Asc(s, IIf(Roms(RomData).RomType = 1, True, False))
  t = Left(t, InStr(1, s, Chr(255), vbBinaryCompare) + 1)
  txtDex = t
  Get #1, txtDexPtr2 + 1, s
  t = Sapp2Asc(s, IIf(Roms(RomData).RomType = 1, True, False))
  t = Left(t, InStr(1, s, Chr(255), vbBinaryCompare) + 1)
  txtDex = txtDex & " " & t
  txtDex = Replace(txtDex, "\n", " ")
  txtDex = Replace(txtDex, "\", "")
End Sub

'------------------------------------------------------'

Private Sub txtPkmnOff_Change()
  DexInfo(sel).iPkmnVOffset = Val(txtPkmnOff)
End Sub

Private Sub txtTrnrOff_Change()
  DexInfo(sel).iTrnrVOffset = Val(txtTrnrOff)
End Sub

Private Sub txtPkmnScale_LostFocus()
  'TODO - Convert scale bytes to percentages
  DexInfo(sel).iPkmnSize = Val(txtPkmnScale)
End Sub

Private Sub txtTrnrScale_LostFocus()
  'TODO - Convert scale bytes to percentages
  DexInfo(sel).iTrnrSize = Val(txtTrnrScale)
End Sub


