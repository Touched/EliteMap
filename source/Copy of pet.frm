VERSION 5.00
Begin VB.Form frmPet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P.E.T"
   ClientHeight    =   7425
   ClientLeft      =   1230
   ClientTop       =   3360
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "pet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   6960
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPokemon 
      Caption         =   "Pokémon"
      Height          =   3135
      Left            =   1920
      TabIndex        =   26
      Top             =   4200
      Width           =   4935
      Begin VB.ComboBox cboSpecies 
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   255
         Left            =   1320
         TabIndex        =   34
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmdPrev 
         Caption         =   "<"
         Height          =   255
         Left            =   840
         TabIndex        =   35
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox picPokemon 
         AutoRedraw      =   -1  'True
         Height          =   975
         Left            =   840
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
      Begin VB.Frame fraAttack 
         Caption         =   "Attacks "
         Height          =   1335
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   4695
         Begin VB.ComboBox cboAttack1 
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox cboAttack2 
            Height          =   315
            Left            =   2520
            TabIndex        =   39
            Top             =   360
            Width           =   1935
         End
         Begin VB.ComboBox cboAttack3 
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   840
            Width           =   2055
         End
         Begin VB.ComboBox cboAttack4 
            Height          =   315
            Left            =   2520
            TabIndex        =   37
            Top             =   840
            Width           =   1935
         End
      End
      Begin VB.HScrollBar hsbPokemon 
         Height          =   255
         Left            =   2520
         Max             =   6
         Min             =   1
         TabIndex        =   29
         Top             =   720
         Value           =   1
         Width           =   1215
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   120
         MaxLength       =   3
         TabIndex        =   28
         Text            =   "0"
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save Pokémon"
         Height          =   495
         Left            =   2520
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Pokémon"
         Height          =   255
         Left            =   840
         TabIndex        =   36
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Level"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Scroll to next or previous Pokémon"
         Height          =   495
         Left            =   2520
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmTrainer 
      Caption         =   "Trainer Status"
      Height          =   4095
      Left            =   1920
      TabIndex        =   6
      Top             =   0
      Width           =   4935
      Begin VB.TextBox txtMusic 
         Height          =   285
         Left            =   240
         TabIndex        =   47
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtOffset 
         Height          =   285
         Left            =   3360
         TabIndex        =   45
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtPokemonHeld 
         Height          =   285
         Left            =   3360
         TabIndex        =   43
         Text            =   "1"
         Top             =   720
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "Items"
         Height          =   1455
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   4695
         Begin VB.ComboBox cboItem1 
            Height          =   315
            ItemData        =   "pet.frx":030A
            Left            =   120
            List            =   "pet.frx":030C
            TabIndex        =   21
            Top             =   480
            Width           =   2175
         End
         Begin VB.ComboBox cboItem2 
            Height          =   315
            Left            =   2400
            TabIndex        =   20
            Top             =   480
            Width           =   2175
         End
         Begin VB.ComboBox cboItem3 
            Height          =   315
            Left            =   120
            TabIndex        =   19
            Top             =   1080
            Width           =   2175
         End
         Begin VB.ComboBox cboItem4 
            Height          =   315
            Left            =   2400
            TabIndex        =   18
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label9 
            Caption         =   "Item 1"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Item 2"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   24
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Item 3"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Item 4"
            Height          =   255
            Left            =   2400
            TabIndex        =   22
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmdSaveTrainer 
         Caption         =   "Save Trainer Data"
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox txtTrnName 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cboTrnGrp 
         Height          =   315
         ItemData        =   "pet.frx":030E
         Left            =   240
         List            =   "pet.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox chkDuo 
         Caption         =   "Double Battle"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Width           =   1575
      End
      Begin VB.PictureBox picTrainer 
         Height          =   975
         Left            =   2280
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdPlusOne 
         Caption         =   ">"
         Height          =   255
         Left            =   2760
         TabIndex        =   8
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdMinusone 
         Caption         =   "<"
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Music"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Pokémon Offset "
         Height          =   255
         Left            =   3360
         TabIndex        =   44
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Number of Pokémon"
         Height          =   255
         Left            =   3360
         TabIndex        =   42
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label13 
         Height          =   375
         Left            =   3480
         TabIndex        =   41
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Trainer Class"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Trainer name"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Label8"
         Height          =   15
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.TextBox txtTheFile 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   8880
      Width           =   1335
   End
   Begin VB.ComboBox cboUnknown21 
      Height          =   315
      Left            =   4440
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   9000
      Width           =   1095
   End
   Begin VB.ComboBox cboBlank 
      Height          =   315
      Left            =   6720
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   8880
      Width           =   1815
   End
   Begin VB.ComboBox cboPKMNUnknown 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   8880
      Width           =   1215
   End
   Begin VB.ListBox lstTrainerLoad 
      Height          =   7455
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblVer 
      Caption         =   "Version"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   9000
      Width           =   3615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuClrm 
         Caption         =   "Close Rom"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuTrainer 
      Caption         =   "Trainer"
      Begin VB.Menu mnuDTTF 
         Caption         =   "Dump To text file"
      End
      Begin VB.Menu mnuImport 
         Caption         =   "Import Trainer Data"
      End
      Begin VB.Menu mnuETD 
         Caption         =   "Export Trainer Data"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find Trainer"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear Data"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuFindpk 
         Caption         =   "Find Pokemon"
      End
      Begin VB.Menu mnuFindAt1 
         Caption         =   "Find attack1"
      End
      Begin VB.Menu mnuFindat2 
         Caption         =   "Find attack2"
      End
      Begin VB.Menu mnuFindAt3 
         Caption         =   "Find Attack3"
      End
      Begin VB.Menu mnuFindAt4 
         Caption         =   "Find Attack4"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuTNE 
         Caption         =   "Trainer Name Editor"
      End
      Begin VB.Menu mnuPT 
         Caption         =   "Show as text"
      End
      Begin VB.Menu mnuGV 
         Caption         =   "Pokemon View"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMS 
         Caption         =   "Make Script"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'KAWA: Tidied up the code a bit, fixed some boo-boos...

Private Const PICSDIR = "pics"
Private Const MAXTRAINERS = &H2B4

Private Type tTrainer
  bGroup As Byte
  bMusic As Byte
  bSprite As Byte
  sName As String * 12
  iItem1 As Integer
  iItem2 As Integer
  iItem3 As Integer
  iItem4 As Integer
  lDuo As Long
  lFilller As Long
  lHeld As Long
  lPntr As Long
End Type
Private Type tpokemon
  iUnknown1 As Integer
  iLevel As Integer
  iSpecies As Integer
  iUnknown2 As Integer
End Type
Private Type Epokemon
  iUnknown1 As Integer
  iLevel As Integer
  iSpecies As Integer
  iUnknown2 As Byte
  iUknown3 As Byte
  iAttack1 As Integer
  iAttack2 As Integer
  iAttack3 As Integer
  iAttack4 As Integer
End Type
Private Type tsPokemon
  iUnknown1 As Integer
  iLevel As Integer
  iSpecies As Integer
  iAttack1 As Integer
  iAttack2 As Integer
  iAttack3 As Integer
  iAttack4 As Integer
  iUnknown2 As Integer
End Type
Private Type Classes
  Class As String * 13
End Type

Private Class(107) As Classes
Private pokemon(1 To 6) As tpokemon
Private Gpokemon(1 To 6) As tsPokemon
Private Epokemon(1 To 6) As Epokemon
Private Trainer(MAXTRAINERS) As tTrainer
Private import() As tTrainer
Private Editing As Integer

Dim RomData As Integer
Dim SetDebug As Integer

Private Sub cboAttack1_LostFocus()
  Gpokemon(hsbPokemon).iAttack1 = cboAttack1.ListIndex
  Epokemon(hsbPokemon).iAttack1 = cboAttack1.ListIndex
End Sub

Private Sub cboAttack2_LostFocus()
  Gpokemon(hsbPokemon).iAttack2 = cboAttack2.ListIndex
  Epokemon(hsbPokemon).iAttack2 = cboAttack2.ListIndex
End Sub

Private Sub cboAttack3_LostFocus()
  Gpokemon(hsbPokemon).iAttack3 = cboAttack3.ListIndex
  Epokemon(hsbPokemon).iAttack3 = cboAttack3.ListIndex
End Sub

Private Sub cboAttack4_LostFocus()
  Gpokemon(hsbPokemon).iAttack4 = cboAttack4.ListIndex
  Epokemon(hsbPokemon).iAttack4 = cboAttack4.ListIndex
End Sub

Private Sub cboItem1_LostFocus()
  Trainer(Editing).iItem1 = cboItem1.ListIndex
End Sub

Private Sub cboItem2_LostFocus()
  Trainer(Editing).iItem2 = cboItem2.ListIndex
End Sub

Private Sub cboItem3_LostFocus()
  Trainer(Editing).iItem3 = cboItem3.ListIndex
End Sub

Private Sub cboItem4_LostFocus()
  Trainer(Editing).iItem4 = cboItem4.ListIndex
End Sub

Private Sub cboMusic_LostFocus()
  Trainer(Editing).bMusic = cboMusic.ListIndex
End Sub

Private Sub cboSpecies_LostFocus()
  'KAWA: Removed multiple calls to PokemonPics(). Waste of cycles.
  pokemon(hsbPokemon).iSpecies = cboSpecies.ListIndex
  'PokemonPics
  Epokemon(hsbPokemon).iSpecies = cboSpecies.ListIndex
  'PokemonPics
  Gpokemon(hsbPokemon).iSpecies = cboSpecies.ListIndex
  PokemonPics
End Sub

Private Sub chkDuo_Click()
  Trainer(Editing).lDuo = chkDuo
End Sub

Private Sub Form_Load()
  Dim b As Boolean
  SetIcon hwnd, "AAA", True
  
  On Error GoTo Hell
  'Open "PET.cfg" For Input As #1
  'Input #1, b
  'Close #1
  b = INIRead("elitemap", "PET", "Show as Text") 'KAWA: Changed to EliteMap.INI
  If b Then
    mnuPT.Checked = True
    cmdPrev.Visible = False
    cmdNext.Visible = False
    picPokemon.Visible = False
    cboSpecies.Visible = True
  End If
  
  'Kawa: Fucked up here, Matt. I'm rewriting this now...
  'If PICSDIR = Null Then
  '  MsgBox "Don't blame inter for what you did. Don't call PET stupid. It's your own fault. Dump the pics."
  '  Unload Me
  'End If
  If Dir(PICSDIR & "\trainer-024.bmp") = "" Then
    MsgBox "Maybe you should consider dumping the trainer " & vbCrLf & "and pokemon pictures from RS-Ball first...", vbExclamation
    End
  End If
Hell:
  '"Oh crap, singing. Mind if I smoke?" -- Bender
End Sub

Private Sub mnuDTTF_Click()
  'With cdlOpen
  '  .DialogTitle = "Save PET Datafile"
  '  .Filter = "Text files| *.txt"
  '  .ShowSave
  '  TheFile = .Filename
  '  If Len(.Filename) = 0 Then
  '    MsgBox "Canceled"
  '    Unload Me
  '  End If
  'End With
  Dim MyCDL As New cCommonDialog
  Dim TheFile As String
  If Not MyCDL.VBGetSaveFileName(TheFile, , , "Text files|*.txt") Then Exit Sub
  
  Open TheFile For Output As #2
  Print #2, "Trainer Data"
  Print #2, "¯¯¯¯¯¯¯¯¯¯¯¯"
  Print #2, "Name:" & vbTab & txtTrnName.Text
  Print #2, "Trainer Class:" & vbTab & cboTrnGrp.Text
  'a = chkDuo.Value
  'If a = 1 Then b = "Yes"
  'If a = 0 Then b = "No"
  'Print #2, , "Double Battle:" & b
  'KAWA: Here's a nice trick...
  Print #2, "Double Battle: " & vbTab & IIf(chkDuo.Value, "Yes", "No")
  Print #2, "Pokemon Held:" & txtPokemonHeld
  Print #2, "Item 1:" & vbTab & cboItem1.Text & vbTab & "Item2:" & vbTab & cboItem2.Text
  Print #2, "Item 3:" & vbTab & cboItem3.Text & vbTab & "Item4:" & vbTab & cboItem4.Text
  Print #2, ""
  Print #2, "Pokemon"
  Print #2, "¯¯¯¯¯¯¯"
  For i = 1 To txtPokemonHeld
    hsbPokemon.Value = i
    hsbPokemon_Change
    If Trainer(Editing).bGroup = &H17 Or Trainer(Editing).bGroup = &H18 Or Trainer(Editing).bGroup = &H1E Or Trainer(Editing).bGroup = &H5A Or Trainer(Editing).bGroup = &H54 Or Trainer(Editing).bGroup = &H19 Or Trainer(Editing).bGroup = &H57 Then
      Print #2, "Pokemon in slot #" & i
      Print #2, "Pokemon:" & vbTab & cboSpecies.Text & vbTab & "Level:" & vbTab & txtLevel
      Print #2, "Attack 1:" & vbTab & cboAttack1.Text & vbTab & "Attack2:" & vbTab & cboAttack2.Text
      Print #2, "Attack 3:" & vbTab & cboAttack3.Text & vbTab & "Attack4:" & vbTab & cboAttack4.Text
      Print #2, ""
    Else
      Print #2, "Pokemon in slot #" & i
      Print #2, "Pokemon:" & vbTab & cboSpecies.Text & vbTab & "Level:" & vbTab & txtLevel
      Print #2, ""
    End If
  Next i
  hsbPokemon.Value = 1
  hsbPokemon_Change
  Close #2
End Sub

Private Sub mnuGV_Click()
  cmdPrev.Visible = True
  cmdNext.Visible = True
  picPokemon.Visible = True
  cboSpecies.Visible = False
End Sub

Private Sub mnuImport_Click()
  Dim headr As String * 3
  Dim import As tTrainer
  Dim Gipokemon(1 To 6) As tsPokemon
  Dim Nipokemon(1 To 6) As tpokemon
  Dim e4pokemon(1 To 6) As Epokemon
  Dim check As Byte
  '  With cdlOpen
  '  .DialogTitle = "Open PET Datafile"
  '  .Filter = "PET datafiles| *.PET"
  '  .ShowOpen
  '  TheFile = .Filename
  '   If Len(.Filename) = 0 Then
  'MsgBox "No rom loaded"
  'Unload Me
  'End If
  '  End With
  
  Dim MyCDL As New cCommonDialog
  Dim TheFile As String
  If Not MyCDL.VBGetOpenFileName(TheFile, , , , , , "PET datafiles (*.pet)|*.pet") Then Exit Sub
  
  Open TheFile For Binary As #2
  Get #2, , headr
  If headr = "PET" Then
    Get #2, , import
    Trainer(Editing) = import
    cmdSaveTrainer_Click
    
    Get #2, &H4, check
    Seek #2, &H55
    If Trainer(Editing).bGroup = &H17 Or Trainer(Editing).bGroup = &H18 Or Trainer(Editing).bGroup = &H19 Or Trainer(Editing).bGroup = &H20 Or Trainer(Editing).bGroup = &H1E Or Trainer(Editing).bGroup = &H5A Or Trainer(Editing).bGroup = &H54 Then
      For i = 1 To Trainer(Editing).lHeld
        Get #2, , Gipokemon(i)
        hsbPokemon.Value = i
        hsbPokemon_Change
        Gpokemon(hsbPokemon) = Gipokemon(i)
      Next i
      cmdSave_Click
      ElseIf check = &H57 Then
        For i = 1 To Trainer(Editing).lHeld
          Get #2, , Epokemon(i)
          hsbPokemon.Value = i
          hsbPokemon_Change
          Epokemon(hsbPokemon) = e4pokemon(i)
        Next i
      Else
      For i = 1 To Trainer(Editing).lHeld
        Get #2, , Nipokemon(i)
        hsbPokemon.Value = i
        hsbPokemon_Change
        pokemon(hsbPokemon) = Nipokemon(i)
      Next i
    cmdSave_Click
    End If
  Else
    MsgBox "This is not a PET datafile."
  End If
  Close #2
End Sub

Private Sub cmdMinusone_Click()
  Trainer(Editing).bSprite = Trainer(Editing).bSprite - 1
  TrainerPics
End Sub

Private Sub cmdNext_Click()
  If Trainer(Editing).bGroup = &H17 Or Trainer(Editing).bGroup = &H18 Or Trainer(Editing).bGroup = &H19 Or Trainer(Editing).bGroup = &H20 Or Trainer(Editing).bGroup = &H1E Or Trainer(Editing).bGroup = &H5A Or Trainer(Editing).bGroup = &H54 Then
    Gpokemon(hsbPokemon).iSpecies = Gpokemon(hsbPokemon).iSpecies + 1
    PokemonPics
  ElseIf Trainer(Editing).bGroup = &H57 Then
    Epokemon(hsbPokemon).iSpecies = Epokemon(hsbPokemon).iSpecies + 1
    PokemonPics
  Else
    pokemon(hsbPokemon).iSpecies = pokemon(hsbPokemon).iSpecies + 1
    PokemonPics
  End If
End Sub

Private Sub cmdPlusOne_Click()
  Trainer(Editing).bSprite = Trainer(Editing).bSprite + 1
  TrainerPics
End Sub

Private Sub cboTrnGrp_LostFocus()
  Trainer(Editing).bGroup = cboTrnGrp.ListIndex
End Sub

Private Sub cbotrngrp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuName
End Sub

Private Sub cmdPrev_Click()
  pokemon(hsbPokemon.Value).iSpecies = pokemon(hsbPokemon.Value).iSpecies - 1
  PokemonPics
End Sub

Private Sub cmdSave_Click()
  If Trainer(Editing).bGroup = &H17 Or Trainer(Editing).bGroup = &H18 Or Trainer(Editing).bGroup = &H19 Or Trainer(Editing).bGroup = &H20 Or Trainer(Editing).bGroup = &H1E Or Trainer(Editing).bGroup = &H5A Or Trainer(Editing).bGroup = &H54 Then
    Seek #1, Trainer(Editing).lPntr - &H8000000 + 1
    For i = 1 To Trainer(Editing).lHeld
      Put #1, , Gpokemon(i)
    Next i
  ElseIf Trainer(Editing).bGroup = &H57 Then
    Seek #1, Trainer(Editing).lPntr - &H8000000 + 1
    For i = 1 To Trainer(Editing).lHeld
      Put #1, , Epokemon(i)
    Next i
  Else
    Seek #1, Trainer(Editing).lPntr - &H8000000 + 1
    For i = 1 To Trainer(Editing).lHeld
      Put #1, , pokemon(i)
    Next i
  End If
End Sub

Private Sub cmdSaveTrainer_Click()
  For i = 0 To MAXTRAINERS
    Put #1, Roms(RomData).TrainerData + (i * 40) + 1, Trainer(i)
  Next i
End Sub

Private Sub hsbPokemon_Change()
  On Error Resume Next
  If Epokemon(hsbPokemon).iAttack1 = &HFFFF Then
    Epokemon(hsbPokemon).iAttack1 = &H0
  ElseIf Epokemon(hsbPokemon).iAttack2 = &HFFFF Then
    Epokemon(hsbPokemon).iAttack2 = &H0
  ElseIf Epokemon(hsbPokemon).iAttack3 = &HFFFF Then
    Epokemon(hsbPokemon).iAttack3 = &H0
  ElseIf Epokemon(hsbPokemon).iAttack4 = &HFFFF Then
    Epokemon(hsbPokemon).iAttack4 = &H0
  End If
  If Gpokemon(hsbPokemon).iAttack1 = &HFFFF Then
    Gpokemon(hsbPokemon).iAttack1 = &H0
  ElseIf Gpokemon(hsbPokemon).iAttack2 = &HFFFF Then
    Gpokemon(hsbPokemon).iAttack2 = &H0
  ElseIf Gpokemon(hsbPokemon).iAttack3 = &HFFFF Then
    Gpokemon(hsbPokemon).iAttack3 = &H0
  ElseIf Gpokemon(hsbPokemon).iAttack4 = &HFFFF Then
    Gpokemon(hsbPokemon).iAttack4 = &H0
  End If
  hsbPokemon.Max = Trainer(Editing).lHeld
  If Trainer(Editing).bGroup = &H17 Or Trainer(Editing).bGroup = &H19 Or Trainer(Editing).bGroup = &H18 Or Trainer(Editing).bGroup = &H20 Or Trainer(Editing).bGroup = &H1E Or Trainer(Editing).bGroup = &H5A Or Trainer(Editing).bGroup = &H54 Or Trainer(Editing).bGroup = &H19 Or Trainer(Editing).bGroup = &H2E Then
    fraAttack.Enabled = True
    txtLevel.Text = Gpokemon(hsbPokemon).iLevel
    cboSpecies.ListIndex = Gpokemon(hsbPokemon).iSpecies
    cboAttack1.ListIndex = Gpokemon(hsbPokemon).iAttack1
    cboAttack2.ListIndex = Gpokemon(hsbPokemon).iAttack2
    cboAttack3.ListIndex = Gpokemon(hsbPokemon).iAttack3
    cboAttack4.ListIndex = Gpokemon(hsbPokemon).iAttack4
  ElseIf Trainer(Editing).bGroup = &H57 Then
    fraAttack.Enabled = True
    txtLevel.Text = Epokemon(hsbPokemon).iLevel
    cboSpecies.ListIndex = Epokemon(hsbPokemon).iSpecies
    cboAttack1.ListIndex = Epokemon(hsbPokemon).iAttack1
    cboAttack2.ListIndex = Epokemon(hsbPokemon).iAttack2
    cboAttack3.ListIndex = Epokemon(hsbPokemon).iAttack3
    cboAttack4.ListIndex = Epokemon(hsbPokemon).iAttack4
  Else
    txtLevel.Text = pokemon(hsbPokemon).iLevel
    cboSpecies.ListIndex = pokemon(hsbPokemon).iSpecies
    fraAttack.Enabled = False
  End If
  PokemonPics
End Sub

Private Function NameSap2Asc(SappName As String)
  NameSap2Asc = Replace(Replace(Replace(Replace(Sapp2Asc(SappName), "\x", ""), "\h2D", "&"), "î", ""), """", "")
End Function

Private Sub fraPokemon_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub lstTrainerLoad_Click()
  On Error Resume Next
  Editing = lstTrainerLoad.ListIndex
  'Read Offset a which is the beginning of the data!
  
  ' Label13.Caption = Hex(Trainer(Editing).lPntr)
  cboTrnGrp.ListIndex = Trainer(Editing).bGroup
  'cboMusic.ListIndex = Trainer(Editing).bMusic
  txtMusic = Trainer(Editing).bMusic
  txtTrnName = NameSap2Asc(Trainer(Editing).sName)
  chkDuo.Value = Trainer(Editing).lDuo
  txtPokemonHeld = Trainer(Editing).lHeld
  TrainerPics
  
  cboItem1.ListIndex = Trainer(Editing).iItem1
  cboItem2.ListIndex = Trainer(Editing).iItem2
  cboItem3.ListIndex = Trainer(Editing).iItem3
  cboItem4.ListIndex = Trainer(Editing).iItem4
  held = txtPokemonHeld
  txtOffset = "&H" & Hex(Trainer(Editing).lPntr - &H8000000)
  
  Seek #1, Trainer(Editing).lPntr - &H8000000 + 1
  For i = 1 To Trainer(Editing).lHeld
    Get #1, , Gpokemon(i)
  Next i
  
  Seek #1, Trainer(Editing).lPntr - &H8000000 + 1
  For i = 1 To Trainer(Editing).lHeld  '- 1
    Get #1, , pokemon(i)
  Next i
  
  Seek #1, Trainer(Editing).lPntr - &H8000000 + 1
  For i = 1 To Trainer(Editing).lHeld  '- 1
    Get #1, , Epokemon(i)
  Next i
  
  hsbPokemon.Value = 1
  hsbPokemon_Change
End Sub

Private Sub mnuClear_Click()
  Trainer(Editing).bGroup = &H0
  Trainer(Editing).bMusic = &H0
  Trainer(Editing).bSprite = &H0
  Trainer(Editing).iItem1 = &H0 '* 2 'KAWA: Because zero times two is still zero.
  Trainer(Editing).iItem2 = &H0 '* 2
  Trainer(Editing).iItem3 = &H0 '* 2
  Trainer(Editing).iItem4 = &H0 '* 2
  Trainer(Editing).lDuo = &H0 '* 4
  Trainer(Editing).lHeld = &H0 '* 4
  Trainer(Editing).sName = "" '&H0 * 12
  cmdSaveTrainer_Click
End Sub

Private Sub mnuClrm_Click()
  Close #1
  lstTrainerLoad.Clear
  cboTrnGrp.Clear
  cboItem1.Clear
  cboItem2.Clear
  cboItem3.Clear
  cboSpecies.Clear
  cboAttack1.Clear
  cboAttack2.Clear
  cboAttack3.Clear
  cboAttack4.Clear
  MsgBox "Rom closed"
End Sub

Private Sub mnuETD_Click()
  Dim headr As String
  headr = "PET"
  'With cdlOpen
  '.DialogTitle = "Save PET Datafile"
  '.Filter = "PET datafiles| *.PET"
  '.ShowSave
  'TheFile = .Filename
  ' If Len(.Filename) = 0 Then
  'MsgBox "No rom loaded"
  'Unload Me
  'End If
  'End With
  Dim MyCDL As New cCommonDialog
  Dim TheFile As String
  If Not MyCDL.VBGetSaveFileName(TheFile, , , "PET datafiles (*.pet)|*.pet") Then Exit Sub
  
  Open TheFile For Binary As #2
  Put #2, , headr
  Put #2, , Trainer(Editing)
  Seek #2, &H55
  If Trainer(Editing).bGroup = &H17 Or Trainer(Editing).bGroup = &H18 Or Trainer(Editing).bGroup = &H19 Or Trainer(Editing).bGroup = &H20 Or Trainer(Editing).bGroup = &H1E Or Trainer(Editing).bGroup = &H5A Or Trainer(Editing).bGroup = &H54 Then
    For i = 1 To Trainer(Editing).lHeld
      Put #2, , Gpokemon(i)
    Next i
  ElseIf Trainer(Editing).bGroup = &H57 Then
    For i = 1 To Trainer(Editing).lHeld
      Put #2, , Epokemon(i)
    Next
  Else
    For i = 1 To Trainer(Editing).lHeld
      Put #2, , pokemon(i)
    Next i
  End If
  Close #2
End Sub

Private Sub mnuFind_Click()
  a = UCase(InputBox("Find trainers whose names contain the following, starting the search at " & lstTrainerLoad.List(lstTrainerLoad.ListIndex)))
  For i = lstTrainerLoad.ListIndex To lstTrainerLoad.ListCount - 1
    b = lstTrainerLoad.List(i)
    If InStr(b, a) Then
      lstTrainerLoad.ListIndex = i
      Exit Sub
    End If
  Next i
End Sub

'KAWA: Ever considered combining this into one mnuFindAt(Index) sub?
Private Sub mnuFindAt1_Click()
  a = UCase(InputBox("Find the attack with the following for attack1 " & cboAttack1.List(cboAttack1.ListIndex)))
  For i = cboAttack1.ListIndex To cboAttack1.ListCount - 1
    b = cboAttack1.List(i)
    'Trace b
    If InStr(b, a) Then
      cboAttack1.ListIndex = i
      Gpokemon(hsbPokemon).iAttack1 = i
      Exit Sub
    End If
  Next i
End Sub

Private Sub mnuFindat2_Click()
  a = UCase(InputBox("Find the attack with the following for attack2 " & cboAttack2.List(cboAttack2.ListIndex)))
  For i = cboAttack2.ListIndex To cboAttack2.ListCount - 1
    b = cboAttack2.List(i)
    'Trace b
    If InStr(b, a) Then
      cboAttack2.ListIndex = i
      Gpokemon(hsbPokemon).iAttack2 = i
      Exit Sub
    End If
  Next i
End Sub

Private Sub mnuFindAt3_Click()
  a = UCase(InputBox("Find the attack with the following for attack3 " & cboAttack3.List(cboAttack3.ListIndex)))
  For i = cboAttack3.ListIndex To cboAttack3.ListCount - 1
    b = cboAttack3.List(i)
    'Trace b
    If InStr(b, a) Then
      cboAttack3.ListIndex = i
      Gpokemon(hsbPokemon).iAttack3 = i
      Exit Sub
    End If
  Next i
End Sub

Private Sub mnuFindAt4_Click()
  a = UCase(InputBox("Find the attack with the following for Attack4 " & cboAttack4.List(cboAttack4.ListIndex)))
  For i = cboAttack4.ListIndex To cboAttack4.ListCount - 1
    b = cboAttack4.List(i)
    'Trace b
    If InStr(b, a) Then
      cboAttack4.ListIndex = i
      Gpokemon(hsbPokemon).iAttack4 = i
      Exit Sub
    End If
  Next i
End Sub

Private Sub mnuFindpk_Click()
  a = UCase(InputBox("Find a Pokemon" & cboSpecies.List(cboSpecies.ListIndex)))
  For i = cboSpecies.ListIndex To cboSpecies.ListCount - 1
    b = cboSpecies.List(i)
    'Trace b
    If InStr(b, a) Then
      cboSpecies.ListIndex = i
      pokemon(hsbPokemon).iSpecies = i
      Gpokemon(hsbPokemon).iSpecies = i
      Epokemon(hsbPokemon).iSpecies = i
      PokemonPics
      Exit Sub
    End If
  Next i
  'pokemon(hsbPokemon).iSpecies = A
  'Gpokemon(hsbPokemon).iSpecies = A
  'pokemonPics
  'PokemonPics
End Sub

Private Sub mnuMS_Click()
  Dim a As Long
  Dim c As Integer
  Dim d0 As Integer
  Dim f As Long
  Dim g As Long
  Dim message As Long
  TheFile = txtTheFile
  Open TheFile For Binary As #245
  a = Val(InputBox("Enter the offset for the TrainerBattle if the address is in hex please put &H before addres"))
  
  If a = 0 Then Exit Sub  'KAWA -- Added CANCEL trap
  
  i = lstTrainerLoad.ListIndex
  Seek #245, a + 1
  Put #245, , &H5C
  b = Val(InputBox("Enter either a 0 1 2 or 3 for the battle type"))
  Put #245, , b
  Put #245, , i
  d0 = &H0
  Put #245, , d0
  f = Val(InputBox("Enter a offset for intro text Message will be asked for later if the address is in hex please put &H before address")) + &H8000000
  Put #245, , f
  g = Val(InputBox("Enter a offset for defeat text Message will be asked for later if the address is in hex please put &H before address")) + &H8000000
  Put #245, , g
  Put #245, , &HF
  Put #245, , &H0
  message = Val(InputBox("Enter an offset for after battle if the address is in hex please put &H before address ")) + &H8000000
  Put #245, , message
  Put #245, , &H9
  Put #245, , &H5
  Put #245, , &H2
  Seek #245, f - &H8000000
  X = InputBox("Enter a message for the introduction message")
  Put #245, , Asc2Sapp(X) & &HFF
  Seek #245, g - &H8000000
  d = InputBox("Enter a message for the introduction message")
  Put #245, , Asc2Sapp(d) & &HFF
  Seek #245, message - &H8000000
  asa = InputBox("Enter the after battle text")
  Put #245, , Asc2Sapp(asa) & &HFF
  MsgBox "The script that was inserted is located at" & Hex(a) & " The Trainer you inserted the script for is " & lstTrainerLoad.ListIndex
  Close #245
End Sub

Private Sub mnuOpen_Click()
  Dim i As Long
  Dim pokename As String * 11
  Dim headr As String * 4
  Dim TheFile As String
  Dim trainername As String * 12
  Dim ItemNames As Long
  Dim ItemName As String * 14
  Dim Filler As String * 30
  Dim AttackName As String * 13
  Dim AttackNameList As Long
  InitDatabase
  On Error Resume Next
  
  Dim MyCDL As New cCommonDialog
  If Not MyCDL.VBGetOpenFileName(TheFile, , , , , , "GBA roms|*.gba") Then Exit Sub
  
  CheckLock TheFile
  
  Open TheFile For Binary As #1
  Get #1, &HAD, headr
  RomData = FindRom(headr)
  If RomData = -1 Then
    MsgBox "Unsupported ROM."
    Unload Me
  End If
  txtTheFile = TheFile
  SetDebug = 0
  If SetDebug = 1 Then
    frmPet.Caption = "P.E.T: Debug Version"
  Else
    frmPet.Caption = "P.E.T - " & Roms(RomData).Name
  End If
  'Trainer name loading code
  
  lstTrainerLoad.Clear
  cboItem1.Clear
  cboItem2.Clear
  cboItem3.Clear
  cboItem4.Clear
  cboTrnGrp.Clear
  cboMusic.Clear
  cboSpecies.Clear
  cboAttack1.Clear
  cboAttack2.Clear
  cboAttack3.Clear
  cboAttack4.Clear
  
  For i = 0 To MAXTRAINERS
    Get #1, Roms(RomData).TrainerData + (i * 40) + 1, Trainer(i)
    lstTrainerLoad.AddItem Right("000" & Hex(i + 1), 3) & ". " & NameSap2Asc(Trainer(i).sName)
  Next i
  
  
  Seek #1, Roms(RomData).ItemNames + 1 'ItemNames + 1
  For i = 0 To &H15B
    Get #1, , ItemName
    Get #1, , Filler
    b$ = Replace(Replace(Sapp2Asc(ItemName), "\x", ""), "\h2D", "&")
    While InStr(1, b$, "\x"): b$ = Left(b$, Len(b$) - 1): Wend
    b$ = Left(b$, Len(b$) - 1)
    'If i = 0 Then b$ = "Nothing"
    cboItem1.AddItem b$
    cboItem2.AddItem b$
    cboItem3.AddItem b$
    cboItem4.AddItem b$
  Next i
  'Read Trainer Classes Names
  'Read Trainer Classes Names
  
  Seek #1, Roms(RomData).TrainerClasses + 1
  If (headr <> "BPRE") And (headr <> "BPGE") Then
    For i = 0 To &H39
      Get #1, , Class(i)
      b$ = NameSap2Asc(Class(i).Class)
      cboTrnGrp.AddItem b$
      cboMusic.AddItem b$
    Next i
  Else
    For i = 0 To &H6A
      Get #1, , Class(i)
      b$ = NameSap2Asc(Class(i).Class)
      cboTrnGrp.AddItem b$
      cboMusic.AddItem b$
    Next i
  End If
  
  Seek #1, Roms(RomData).MonsterNames + 1
  Get #1, , i
  Seek #1, i - &H8000000 + 1
  For i = 0 To 410
    Get #1, , pokename
    c$ = Sapp2Asc(pokename)
    While InStr(1, c$, "\x"): c$ = Left(c$, Len(c$) - 1): Wend
    c$ = Left(c$, Len(c$) - 1)
    cboSpecies.AddItem c$
  Next i
  'Read AttackNames
  
  Seek #1, Roms(RomData).AttackNameList + 1
  For i = 0 To 353
    Get #1, , AttackName
    z = Replace(Sapp2Asc(AttackName), "\x", "")
    cboAttack1.AddItem z
    cboAttack2.AddItem z
    cboAttack3.AddItem z
    cboAttack4.AddItem z
  Next i
End Sub

Private Function NameAsc2Sap(AsciiName As String)
  NameAsc2Sap = Asc2Sapp(UCase(AsciiName)) & Chr$(255)
End Function

Private Sub mnuPT_Click()
  mnuPT.Checked = Not mnuPT.Checked
  cmdPrev.Visible = Not cmdPrev.Visible
  cmdNext.Visible = Not cmdNext.Visible
  picPokemon.Visible = Not picPokemon.Visible
  cboSpecies.Visible = Not cboSpecies.Visible
  INIWrite "elitemap", "PET", "Show as Text", mnuPT.Checked 'KAWA: Changed to EliteMap.INI
  'Open "PET.cfg" For Output As #69
  'Print #69, mnuPT.Checked
  'Close #69
End Sub

Private Sub mnuQuit_Click()
  Unload Me
End Sub

Private Sub mnuTNE_Click()
  MsgBox cboTrnGrp.ListIndex
  Dim a As String * 12
  a = InputBox("What do you want the new name to be it must be a limit of 12 chars sorry.")
  Class(cboTrnGrp.ListIndex).Class = Asc2Sapp(a)
  cboTrnGrp.List(cboTrnGrp.ListIndex) = a
  cboMusic.List(cboTrnGrp.ListIndex) = a
  Seek #1, Roms(RomData).TrainerClasses + 1
  For i = 0 To cboTrnGrp.ListIndex
    Put #1, , Class(i).Class
  Next i
End Sub

Private Sub txtLevel_LostFocus()
  Gpokemon(hsbPokemon).iLevel = txtLevel
  pokemon(hsbPokemon).iLevel = txtLevel
  Epokemon(hsbPokemon).iLevel = txtLevel
End Sub

Private Sub txtOffset_LostFocus()
  Trainer(Editing).lPntr = Val(txtOffset) + &H8000000
  Seek #1, Val(txtOffset) + 1
  If Trainer(Editing).bGroup = &H17 Or Trainer(Editing).bGroup = &H18 Or Trainer(Editing).bGroup = &H1E Or Trainer(Editing).bGroup = &H5A Or Trainer(Editing).bGroup = &H54 Then
    For i = 1 To Trainer(Editing).lHeld
      Put #1, , Gpokemon(i)
    Next i
  ElseIf Trainer(Editing).bGroup = &H57 Then
    For i = 1 To Trainer(Editing).lHeld
      Put #1, , Epokemon(i)
    Next i
  Else
    For i = 1 To Trainer(Editing).lHeld
      Put #1, , pokemon(i)
    Next i
  End If
End Sub

Private Sub txtMusic_Change()
  Trainer(Editing).bMusic = Val(txtMusic)
End Sub

Private Sub txtPokemonHeld_Change()
  'Repoint
  hsbPokemon.Max = txtPokemonHeld
  Trainer(Editing).lHeld = txtPokemonHeld
End Sub
Private Sub txtTrnName_Change()
  Trainer(Editing).sName = NameAsc2Sap(txtTrnName)
  lstTrainerLoad.List(Editing) = Right("000" & Hex(Editing + 1), 3) & ". " & NameSap2Asc(Trainer(Editing).sName)
End Sub
Private Function TrainerPics()
  picTrainer = LoadPicture(PICSDIR & "\trainer-" & Right("000" & Hex(Trainer(Editing).bSprite), 3) & ".bmp")
End Function

Private Function PokemonPics()
  If pokemon(hsbPokemon).iSpecies = &HFFFF Then
    picPokemon = LoadPicture(PICSDIR & "\pkmn-000.bmp")
  Else
    picPokemon = LoadPicture(PICSDIR & "\pkmn-" & Right("000" & Hex(pokemon(hsbPokemon).iSpecies), 3) & ".bmp")
  End If
End Function

