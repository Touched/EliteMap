VERSION 5.00
Begin VB.Form frmBewildered 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bewildered"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBewildered.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFSave 
      Caption         =   "[1] Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   38
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdTSave 
      Caption         =   "[1] Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   37
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdWSave 
      Caption         =   "[1] Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   36
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdGSave 
      Caption         =   "[1] Save"
      Height          =   375
      Left            =   5520
      TabIndex        =   35
      Top             =   360
      Width           =   975
   End
   Begin VB.Frame fraTree 
      Caption         =   "[12] Trees"
      Height          =   1455
      Left            =   2040
      TabIndex        =   22
      Top             =   3240
      Width           =   4575
      Begin VB.ComboBox cboTSpecies 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   960
         Width           =   2655
      End
      Begin VB.HScrollBar hsbTree 
         Height          =   255
         Left            =   120
         Max             =   4
         TabIndex        =   25
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtTMinLevel 
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtTMaxLevel 
         Height          =   285
         Left            =   960
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "[2] Min level"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "[3] Max level"
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "[4] Species"
         Height          =   255
         Left            =   1800
         TabIndex        =   26
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame fraFish 
      Caption         =   "[13] Water (fishing)"
      Height          =   1455
      Left            =   2040
      TabIndex        =   15
      Top             =   4800
      Width           =   4575
      Begin VB.ComboBox cboFSpecies 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtFMaxLevel 
         Height          =   285
         Left            =   960
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtFMinLevel 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   615
      End
      Begin VB.HScrollBar hsbFish 
         Height          =   255
         Left            =   120
         Max             =   7
         TabIndex        =   16
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label10 
         Caption         =   "[4] Species"
         Height          =   255
         Left            =   1800
         TabIndex        =   21
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "[3] Max level"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "[2] Min level"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.Frame fraWater 
      Caption         =   "[11] Water (surfing)"
      Height          =   1455
      Left            =   2040
      TabIndex        =   8
      Top             =   1680
      Width           =   4575
      Begin VB.ComboBox cboWSpecies 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   960
         Width           =   2655
      End
      Begin VB.HScrollBar hsbWater 
         Height          =   255
         Left            =   120
         Max             =   4
         TabIndex        =   11
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtWMinLevel 
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtWMaxLevel 
         Height          =   285
         Left            =   960
         TabIndex        =   9
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "[2] Min level"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "[3] Max level"
         Height          =   255
         Left            =   960
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "[4] Species"
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame fraGrass 
      Caption         =   "[10] Grass"
      Height          =   1455
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtGMaxLevel 
         Height          =   285
         Left            =   960
         TabIndex        =   33
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox cboGSpecies 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtGMinLevel 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.HScrollBar hsbGrass 
         Height          =   255
         Left            =   120
         Max             =   11
         TabIndex        =   3
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label4 
         Caption         =   "[4] Species"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "[3] Max level"
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "[2] Min level"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblVer 
      Caption         =   "Version"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   6360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "[15] Index       Bank  Level"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmBewildered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tMapPiece
  BankNumber As Byte
  MapNumber As Byte
  Filler As Integer
  GrassPtr As Long
  WaterPtr As Long
  TreePtr As Long
  FishPtr As Long
End Type
  
Private Type tEncounter
  LevelMin As Byte
  LevelMax As Byte
  Species As Integer
End Type

Dim MapPieces(1024) As tMapPiece
Dim GrassEncounter(12) As tEncounter
Dim WaterEncounter(5) As tEncounter
Dim TreeEncounter(5) As tEncounter
Dim FishEncounter(8) As tEncounter
Dim RomData As Integer

Private Sub cmdFSave_Click()
  Dim i As Long
  Seek #1, MapPieces(List1.ListIndex).FishPtr - &H8000000 + 1
  Get #1, , i
  Get #1, , i
  Seek #1, i - &H8000000 + 1
  For i = 0 To 7
    Put #1, , FishEncounter(i)
  Next i
End Sub

Private Sub cmdGSave_Click()
  Dim i As Long
  Seek #1, MapPieces(List1.ListIndex).GrassPtr - &H8000000 + 1
  Get #1, , i
  Get #1, , i
  Seek #1, i - &H8000000 + 1
  For i = 0 To 11
    Put #1, , GrassEncounter(i)
  Next i
End Sub

Private Sub cmdTSave_Click()
  Dim i As Long
  Seek #1, MapPieces(List1.ListIndex).TreePtr - &H8000000 + 1
  Get #1, , i
  Get #1, , i
  Seek #1, i - &H8000000 + 1
  For i = 0 To 4
    Put #1, , TreeEncounter(i)
  Next i
End Sub

Private Sub cmdWSave_Click()
  Dim i As Long
  Seek #1, MapPieces(List1.ListIndex).WaterPtr - &H8000000 + 1
  Get #1, , i
  Get #1, , i
  Seek #1, i - &H8000000 + 1
  For i = 0 To 4
    Put #1, , WaterEncounter(i)
  Next i
End Sub

Private Sub Form_Load()
  Dim i As Long
  Dim pokename As String * 11
  Dim b As String
  Dim headr As String * 4
  Dim TheFile As String
  Dim j As Long
  
  SetIcon Me.hwnd, "AAA", True
  
  On Error Resume Next
  Dim ctl As Control
  For Each ctl In Me.Controls
    If Left(ctl.Caption, 1) = "[" Then
      i = Val(Mid(ctl.Caption, 2, 4))
      ctl.Caption = LoadResString(i)
    End If
  Next
  On Error GoTo 0
  
  lblVer = "Bewildered " & App.Major & "." & App.Minor & " by Kyoufu Kawa"
  For i = 15 To 25
    j = j + Asc(Mid(lblVer, i, 1))
  Next i
  'MsgBox Int(j / 2)
  If Int(j / 2) <> 479 Then
    MsgBox LoadResString(20), vbCritical, LoadResString(21)
    End
  End If
  
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
    MsgBox LoadResString(22)
    End
  End If
  If Roms(RomData).WildPokemon = 0 Then
    MsgBox LoadResString(23)
    End
  End If
  
  frmLoading.Show
  frmLoading.Refresh
  
  'Pokémon name loading code risen from the ashes of TrainEd -_-
  Seek #1, Roms(RomData).MonsterNames + 1
  Get #1, , i
  Seek #1, i - &H8000000 + 1
  For i = 0 To 410
    Get #1, , pokename
    b = Sapp2Asc(pokename)
    While InStr(1, b, "\x"): b = Left(b, Len(b$) - 1): Wend
    b = Left(b, Len(b) - 1)
    cboGSpecies.AddItem b
    cboWSpecies.AddItem b
    cboTSpecies.AddItem b
    cboFSpecies.AddItem b
  Next i
  
  Seek #1, Roms(RomData).WildPokemon + 1
  For i = 0 To 512
    Get #1, , MapPieces(i)
    If MapPieces(i).BankNumber = &HFF And MapPieces(i).MapNumber = &HFF Then
      Exit For
    End If
    List1.AddItem PadHex(i, 2) & " - " & PadHex(MapPieces(i).BankNumber, 2) & "." & PadHex(MapPieces(i).MapNumber, 2)
  Next i
  List1.ListIndex = 0
  List1_Click
  
  Unload frmLoading
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Close #1
End Sub

Private Sub hsbFish_Change()
  txtFMinLevel = FishEncounter(hsbFish).LevelMin
  txtFMaxLevel = FishEncounter(hsbFish).LevelMax
  cboFSpecies.ListIndex = FishEncounter(hsbFish).Species
End Sub

Private Sub hsbGrass_Change()
  txtGMinLevel = GrassEncounter(hsbGrass).LevelMin
  txtGMaxLevel = GrassEncounter(hsbGrass).LevelMax
  cboGSpecies.ListIndex = GrassEncounter(hsbGrass).Species
End Sub

Private Sub hsbTree_Change()
  txtTMinLevel = TreeEncounter(hsbTree).LevelMin
  txtTMaxLevel = TreeEncounter(hsbTree).LevelMax
  cboTSpecies.ListIndex = TreeEncounter(hsbTree).Species
End Sub

Private Sub hsbWater_Change()
  txtWMinLevel = WaterEncounter(hsbWater).LevelMin
  txtWMaxLevel = WaterEncounter(hsbWater).LevelMax
  cboWSpecies.ListIndex = WaterEncounter(hsbWater).Species
End Sub

Private Sub List1_Click()
  Dim i As Long
  If MapPieces(List1.ListIndex).GrassPtr > 0 Then
    Seek #1, MapPieces(List1.ListIndex).GrassPtr - &H8000000 + 1
    Get #1, , i
    Get #1, , i
    Seek #1, i - &H8000000 + 1
    For i = 0 To 11
      Get #1, , GrassEncounter(i)
    Next i
    hsbGrass.Value = 0
    hsbGrass_Change
    fraGrass.Enabled = True
  Else
    fraGrass.Enabled = False
  End If

  If MapPieces(List1.ListIndex).WaterPtr > 0 Then
    Seek #1, MapPieces(List1.ListIndex).WaterPtr - &H8000000 + 1
    Get #1, , i
    Get #1, , i
    Seek #1, i - &H8000000 + 1
    For i = 0 To 4
      Get #1, , WaterEncounter(i)
    Next i
    hsbWater.Value = 0
    hsbWater_Change
    fraWater.Enabled = True
  Else
    fraWater.Enabled = False
  End If

  If MapPieces(List1.ListIndex).TreePtr > 0 Then
    Seek #1, MapPieces(List1.ListIndex).TreePtr - &H8000000 + 1
    Get #1, , i
    Get #1, , i
    Seek #1, i - &H8000000 + 1
    For i = 0 To 4
      Get #1, , TreeEncounter(i)
    Next i
    hsbTree.Value = 0
    hsbTree_Change
    fraTree.Enabled = True
  Else
    fraTree.Enabled = False
  End If
  
  If MapPieces(List1.ListIndex).FishPtr > 0 Then
    Seek #1, MapPieces(List1.ListIndex).FishPtr - &H8000000 + 1
    Get #1, , i
    Get #1, , i
    Seek #1, i - &H8000000 + 1
    For i = 0 To 7
      Get #1, , FishEncounter(i)
    Next i
    hsbFish.Value = 0
    hsbFish_Change
    fraFish.Enabled = True
  Else
    fraFish.Enabled = False
  End If
End Sub

Private Sub txtGMaxLevel_Change()
  GrassEncounter(hsbGrass).LevelMax = Val(txtGMaxLevel)
End Sub

Private Sub txtGMinLevel_Change()
  GrassEncounter(hsbGrass).LevelMin = Val(txtGMinLevel)
End Sub

Private Sub txtWMaxLevel_Change()
  WaterEncounter(hsbWater).LevelMax = Val(txtWMaxLevel)
End Sub

Private Sub txtWMinLevel_Change()
  WaterEncounter(hsbWater).LevelMin = Val(txtWMinLevel)
End Sub

Private Sub txtTMaxLevel_Change()
  TreeEncounter(hsbTree).LevelMax = Val(txtTMaxLevel)
End Sub

Private Sub txtTMinLevel_Change()
  TreeEncounter(hsbTree).LevelMin = Val(txtTMinLevel)
End Sub

Private Sub txtFMaxLevel_Change()
  FishEncounter(hsbFish).LevelMax = Val(txtFMaxLevel)
End Sub

Private Sub txtFMinLevel_Change()
  FishEncounter(hsbFish).LevelMin = Val(txtFMinLevel)
End Sub

Private Sub cboGSpecies_Click()
  GrassEncounter(hsbGrass).Species = cboGSpecies.ListIndex
End Sub

Private Sub cboWSpecies_Click()
  WaterEncounter(hsbWater).Species = cboWSpecies.ListIndex
End Sub

Private Sub cboTSpecies_Click()
  TreeEncounter(hsbTree).Species = cboTSpecies.ListIndex
End Sub

Private Sub cboFSpecies_Click()
  FishEncounter(hsbFish).Species = cboFSpecies.ListIndex
End Sub


