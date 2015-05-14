VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPhoenix 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Phoenix"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TrainEd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   350
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraItems 
      Caption         =   "&Items"
      Height          =   1095
      Left            =   2160
      TabIndex        =   20
      Top             =   3240
      Width           =   4335
      Begin VB.ComboBox cboItem 
         Height          =   315
         Index           =   3
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cboItem 
         Height          =   315
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cboItem 
         Height          =   315
         Index           =   1
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   1935
      End
      Begin VB.ComboBox cboItem 
         Height          =   315
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraPokemon 
      Caption         =   "&Pokémon"
      Height          =   1455
      Left            =   2160
      TabIndex        =   10
      Top             =   1680
      Width           =   4335
      Begin VB.CommandButton cmdSaveP 
         Caption         =   "Save PkMn"
         Height          =   375
         Left            =   2880
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Text            =   "0"
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox cboSpecies 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtAmount 
         Height          =   285
         Left            =   960
         TabIndex        =   13
         Text            =   "0"
         Top             =   240
         Width           =   615
      End
      Begin MSComctlLib.TabStrip tabPokemon 
         Height          =   735
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1296
         MultiRow        =   -1  'True
         Style           =   1
         TabFixedWidth   =   697
         TabFixedHeight  =   526
         HotTracking     =   -1  'True
         TabMinWidth     =   0
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   6
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblPokePtr 
         Caption         =   "[ 0xFFFFFF ]"
         Height          =   255
         Left            =   960
         TabIndex        =   29
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblLevel 
         Caption         =   "&Level"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "&Species"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblAmount 
         Caption         =   "&Amount"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame fraGeneral 
      Caption         =   "&General"
      Height          =   1455
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtMusic 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Text            =   "0"
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox chkDual 
         Caption         =   "&Dual Battle"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin MSComctlLib.ImageCombo icbImage 
         Height          =   1050
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1852
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   10535024
         ImageList       =   "ilsImages"
      End
      Begin VB.ComboBox cboClass 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   960
         MaxLength       =   11
         TabIndex        =   3
         Text            =   "NONAME"
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblMusic 
         Caption         =   "&Music"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblClass 
         Caption         =   "&Class"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblName 
         Caption         =   "&Name"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.ListBox lstTrainers 
      Height          =   4695
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ilsImages 
      Left            =   5880
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   10535024
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   10535024
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TrainEd.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLoader 
      Height          =   855
      Left            =   4920
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblRaw 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000017&
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   5280
      Width           =   6375
   End
   Begin VB.Label lblVersion 
      Caption         =   "Booyaka, bitch. Try to replace my name (""Kawa"") and be sorry."
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4920
      Width           =   6375
   End
   Begin VB.Label lblDontForget 
      Caption         =   "Remember: Changes are NOT remembered until you change focus."
      Height          =   495
      Left            =   2160
      TabIndex        =   26
      Top             =   4440
      Width           =   2895
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupFind 
         Caption         =   "&Find..."
      End
      Begin VB.Menu mnuPopupGoto 
         Caption         =   "&Goto..."
      End
      Begin VB.Menu mnuPopupDump 
         Caption         =   "&Dump to file"
      End
   End
End
Attribute VB_Name = "frmPhoenix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const PICSDIR = "Pics"
Private Const MAXTRAINERS = &H2B4

Private Type tTrainer
  bClass As Byte
  bMusic As Byte
  bImage As Byte
  sName As String * 12
  iItem1 As Integer
  iItem2 As Integer
  iItem3 As Integer
  iItem4 As Integer
  lDual As Long
  lFiller As Long
  lNumPokemon As Long
  lPokemonPtr As Long
End Type
Private Trainers(MAXTRAINERS) As tTrainer

Private Type tPokemon
  iUnknown1 As Integer
  iLevel As Integer
  iSpecies As Integer
  iUnknown2 As Integer
End Type
Private Pokemon(5) As tPokemon

Private ThisRom As Integer
Private Editing As Integer

Private Function NameSap2Asc(SappName As String)
  NameSap2Asc = Replace(Replace(Sapp2Asc(SappName), "\x", ""), "î", "")
End Function

Private Function NameAsc2Sap(AsciiName As String)
  NameAsc2Sap = Asc2Sapp(UCase(AsciiName)) & Chr$(255)
End Function

Private Sub chkDual_LostFocus()
  Trainers(Editing).lDual = chkDual.Value
End Sub

Private Sub cmdSaveP_Click()
  Seek #1, Trainers(Editing).lPokemonPtr - &H8000000 + 1
  For i = 0 To Trainers(Editing).lNumPokemon - 1
    Put #1, , Pokemon(i)
  Next i
  lblDontForget = ""
End Sub

Private Sub Form_Load()
  Dim TrainerClass As String * 13
  Dim PokeName As String * 11
  Dim ItemName As String * 14
  Dim ItemData As String * 30
  Dim Headr As String * 4
  Dim APointer As Long
  
  lblVersion = "Phoenix version " & App.Major & "." & App.Minor & " by Kyoufu Kawa"
  
  SetIcon Me.hwnd, "AAA", True
  
  InitDatabase
  
  For i = 1 To Len(lblVersion)
    c = c + Asc(Mid(lblVersion, i, 1))
  Next i
  
  If Dir(PICSDIR & "\trainer-024.bmp") <> "" Then
  Else
    MsgBox "You must have extracted all trainer images with RS-Ball and put them in the " & PICSDIR & " folder.", vbExclamation, "Trainer images not found"
    End
  End If
  
  Trace Int(c / 2)
  If Int(c / 2) <> 1531 Then
    MsgBox "This program has been hacked and will not run.", vbCritical, "Checksum error"
    End
  End If
  
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
  Get #1, &HAD, Headr
  ThisRom = FindRom(Headr)
  If ThisRom = -1 Then
    MsgBox "Unsupported ROM."
    End
  End If
  'Trace Hex(Roms(ThisRom).TrainerData)
  If Roms(ThisRom).TrainerData = 0 Then missing = missing & "> TrainerData" & vbCrLf
  If Roms(ThisRom).TrainerClasses = 0 Then missing = missing & "> TrainerClasses" & vbCrLf
  If Roms(ThisRom).MonsterNames = 0 Then missing = missing & "> MonsterNames" & vbCrLf
  If Roms(ThisRom).ItemNames = 0 Then missing = missing & "> ItemNames" & vbCrLf
  If missing <> "" Then
    MsgBox "Supported ROM, but the following data is unspecified:" & vbCrLf & vbCrLf & missing
    End
  End If

  For i = 0 To MAXTRAINERS
    Get #1, Roms(ThisRom).TrainerData + (i * 40) + 1, Trainers(i)
    lstTrainers.AddItem Right("000" & Hex(i + 1), 3) & ". " & NameSap2Asc(Trainers(i).sName)
  Next i
  
  Seek #1, Roms(ThisRom).TrainerClasses + 1
  For i = 0 To 58
    Get #1, , TrainerClass
    b$ = Sapp2Asc(TrainerClass)
    While InStr(1, b$, "\x"): b$ = Left(b$, Len(b$) - 1): Wend
    b$ = Left(b$, Len(b$) - 1)
    b$ = Replace(b$, "[PK][MN]", "PkMn")
    cboClass.AddItem Right("00" & Hex(i), 2) & ". " & b$
  Next i

  Seek #1, Roms(ThisRom).MonsterNames + 1
  Get #1, , APointer
  Seek #1, APointer - &H8000000 + 1
  For i = 0 To 410
    Get #1, , PokeName
    b$ = Sapp2Asc(PokeName)
    While InStr(1, b$, "\x"): b$ = Left(b$, Len(b$) - 1): Wend
    b$ = Left(b$, Len(b$) - 1)
    cboSpecies.AddItem Right("000" & Hex(i), 3) & ". " & b$
  Next i

  Seek #1, Roms(ThisRom).ItemNames + 1
  For i = 0 To &H15B
    Get #1, , ItemName
    Get #1, , ItemData
    b$ = Sapp2Asc(ItemName)
    While InStr(1, b$, "\x"): b$ = Left(b$, Len(b$) - 1): Wend
    b$ = Left(b$, Len(b$) - 1)
    If i = 0 Then b$ = "Nothing"
    cboItem(0).AddItem Right("000" & Hex(i), 3) & ". " & b$
    cboItem(1).AddItem Right("000" & Hex(i), 3) & ". " & b$
    cboItem(2).AddItem Right("000" & Hex(i), 3) & ". " & b$
    cboItem(3).AddItem Right("000" & Hex(i), 3) & ". " & b$
  Next i
  
  ilsImages.ListImages.Clear
  For i = 0 To &H52
    picLoader.Picture = LoadPicture(PICSDIR & "\trainer-" & Right("000" & Hex(i), 3) & ".bmp")
    ilsImages.ListImages.Add i + 1, , picLoader.Picture
    icbImage.ComboItems.Add i + 1, , "", i + 1
  Next i
  
  lstTrainers.ListIndex = 0
  lstTrainers_Click
End Sub

Private Sub cmdSave_Click()
  For i = 0 To MAXTRAINERS
    Put #1, &H1F0525 + (i * 40) + 1, Trainers(i)
  Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Close #1
End Sub

Private Sub lblPokePtr_DblClick()
  Debug.Print "blPokePtr_DblClick() triggered!"
  If MsgBox("Do you want to repoint this team?", vbYesNo) = vbYes Then
    Debug.Print "     Old location: " & Hex(Trainers(Editing).lPokemonPtr)
    na = Val(InputBox("Enter new location. Make sure that there are is at least " & (txtAmount * 8) & " bytes of free space.", , "&H" & Hex(Trainers(Editing).lPokemonPtr - &H8000000)))
    Debug.Print "     New location: " & Hex(na)
    If na = 0 Then Exit Sub
    Trainers(Editing).lPokemonPtr = na + &H8000000
    If MsgBox("Do you want to copy the current team to it's new location?", vbYesNo) = vbYes Then cmdSaveP_Click
    lstTrainers_Click
    Debug.Print "     Final: " & Hex(Trainers(Editing).lPokemonPtr)
  End If
End Sub

Private Sub lblRaw_Click()
  lblRaw = Right("00" & Hex(Trainers(Editing).bClass), 2)
  lblRaw = lblRaw & Right("00" & Hex(Trainers(Editing).bMusic), 2)
  lblRaw = lblRaw & Right("00" & Hex(Trainers(Editing).bImage), 2) & ","
  For i = 1 To 12
    lblRaw = lblRaw & Hex(Asc(Mid(Trainers(Editing).sName, i, 1)))
  Next i
  lblRaw = lblRaw & "," & Right("0000" & Hex(Trainers(Editing).iItem1), 4)
  lblRaw = lblRaw & Right("0000" & Hex(Trainers(Editing).iItem2), 4)
  lblRaw = lblRaw & Right("0000" & Hex(Trainers(Editing).iItem3), 4)
  lblRaw = lblRaw & Right("0000" & Hex(Trainers(Editing).iItem4), 4) & ","
  lblRaw = lblRaw & Right("00000000" & Hex(Trainers(Editing).lDual), 8) & ","
  lblRaw = lblRaw & Right("00000000" & Hex(Trainers(Editing).lFiller), 8) & ","
  lblRaw = lblRaw & Right("00000000" & Hex(Trainers(Editing).lNumPokemon), 8) & ","
  lblRaw = lblRaw & Right("00000000" & Hex(Trainers(Editing).lPokemonPtr), 8)
End Sub

Private Sub lstTrainers_Click()
  'On Error Resume Next
  Editing = lstTrainers.ListIndex
    
  lblRaw_Click
    
  txtName = NameSap2Asc(Trainers(Editing).sName)
  cboClass.ListIndex = Trainers(Editing).bClass
  txtMusic = Trainers(Editing).bMusic
  icbImage.ComboItems(Trainers(Editing).bImage + 1).Selected = True
  chkDual.Value = Trainers(Editing).lDual
  
  txtAmount = Trainers(Editing).lNumPokemon
  
  cboItem(0).ListIndex = Trainers(Editing).iItem1
  cboItem(1).ListIndex = Trainers(Editing).iItem2
  cboItem(2).ListIndex = Trainers(Editing).iItem3
  cboItem(3).ListIndex = Trainers(Editing).iItem4
  
  'Force tab update
  txtAmount_LostFocus
  
  'Load Pokemon data
  lblPokePtr = "[ 0x" & Right("000000" & Hex((Trainers(Editing).lPokemonPtr - &H8000000)), 6) & " ]"
  Seek #1, Trainers(Editing).lPokemonPtr - &H8000000 + 1
  For i = 0 To Trainers(Editing).lNumPokemon - 1
    Get #1, , Pokemon(i)
  Next i
  tabPokemon.Tabs(1).Selected = True
  tabPokemon_Click
End Sub

Private Sub lstTrainers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuPopupDump_Click()
  If MsgBox("This will create a tab-delimited text file named ""Trainer list.txt"" that you can import in programs such as Excel.", vbOKCancel) = vbCancel Then Exit Sub
  Open "Trainer list.txt" For Output As #2
  Print #2, "Idx" & vbTab & _
            "Class" & vbTab & _
            "Name" & vbTab & _
            "Team ptr" & vbTab & _
            "PkMn" & vbTab
  For i = 0 To lstTrainers.ListCount - 1
    lstTrainers.ListIndex = i
    DoEvents
    Print #2, Right("   " & i, 3) & vbTab & _
              Trim(Mid(cboClass.List(cboClass.ListIndex), 5)) & vbTab & Trim(txtName) & vbTab & _
              "0x" & Right("000000" & (Hex(Trainers(i).lPokemonPtr - &H8000000)), 6) & vbTab & _
              txtAmount
  Next i
  Close #2
End Sub

Private Sub mnuPopupFind_Click()
  a = UCase(InputBox("Find trainers whose names contain the following, starting the search at " & Mid(lstTrainers.List(lstTrainers.ListIndex), 6) & "..."))
  For i = lstTrainers.ListIndex To lstTrainers.ListCount - 1
    b = Mid(lstTrainers.List(i), 6)
    'Trace b
    If InStr(b, a) Then
      lstTrainers.ListIndex = i
      Exit Sub
    End If
  Next i
End Sub

Private Sub mnuPopupGoto_Click()
  On Error Resume Next
  i = InputBox("Go to trainer #", , "&H" & Hex(Editing + 1))
  lstTrainers.ListIndex = i - 1
End Sub

Private Sub tabPokemon_Click()
  cboSpecies.ListIndex = Pokemon(tabPokemon.SelectedItem.Index - 1).iSpecies
  txtLevel = Pokemon(tabPokemon.SelectedItem.Index - 1).iLevel
End Sub

Private Sub txtName_LostFocus()
  Trainers(Editing).sName = NameAsc2Sap(txtName)
  lstTrainers.List(Editing) = Right("000" & Hex(Editing + 1), 3) & ". " & NameSap2Asc(Trainers(Editing).sName)
End Sub

Private Sub cboClass_LostFocus()
  Trainers(Editing).bClass = cboClass.ListIndex
End Sub

Private Sub txtMusic_LostFocus()
  txtMusic = CByte(Val(txtMusic))
  Trainers(Editing).bMusic = CByte(Val(txtMusic))
End Sub

Private Sub icbImage_LostFocus()
  Trainers(Editing).bImage = icbImage.SelectedItem.Index - 1
End Sub

Private Sub txtAmount_LostFocus()
  txtAmount = Val(txtAmount)
  If txtAmount > 6 Then txtAmount = 6
  Trainers(Editing).lNumPokemon = Val(txtAmount)
  tabPokemon.Tabs.Clear
  For i = 1 To Val(txtAmount)
    tabPokemon.Tabs.Add i, Chr(i), i
  Next i
End Sub

Private Sub cboSpecies_LostFocus()
  Pokemon(tabPokemon.SelectedItem.Index - 1).iSpecies = cboSpecies.ListIndex
  lblDontForget = "Don't forget to save the Pokémon data!"
End Sub

Private Sub txtLevel_LostFocus()
  Pokemon(tabPokemon.SelectedItem.Index - 1).iLevel = txtLevel
  lblDontForget = "Don't forget to save the Pokémon data!"
End Sub

Private Sub cboItem_LostFocus(Index As Integer)
  Select Case Index
    Case 0: Trainers(Editing).iItem1 = cboItem(Index).ListCount
    Case 1: Trainers(Editing).iItem2 = cboItem(Index).ListCount
    Case 2: Trainers(Editing).iItem3 = cboItem(Index).ListCount
    Case 3: Trainers(Editing).iItem4 = cboItem(Index).ListCount
  End Select
End Sub

