VERSION 5.00
Begin VB.Form frmPattEd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAttEd"
   ClientHeight    =   1965
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "patted.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      Picture         =   "patted.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton mnuOpen 
      Height          =   375
      Left            =   3360
      Picture         =   "patted.frx":03C4
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "I'm not really a menu item! Honest! My parents named me that!"
      Top             =   1440
      Width           =   375
   End
   Begin VB.ListBox lstOffsets 
      Height          =   1035
      Left            =   5280
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtRepoint 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      MaxLength       =   8
      TabIndex        =   7
      Text            =   "&HFFFFFF"
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtLevel 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.ComboBox cboAttack 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2520
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.HScrollBar hsbAttack 
      Height          =   255
      Left            =   1920
      Max             =   2
      TabIndex        =   1
      Tag             =   "Did you know that disabled scrollbars cannot be selected in Design Mode?"
      Top             =   360
      Width           =   2295
   End
   Begin VB.ListBox lstSpecies 
      Enabled         =   0   'False
      Height          =   1725
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Repoint:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Attack"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Level"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblAttack 
      Caption         =   "Select an attack"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPattEd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'KAWA - Code cleanup finished at 13:29, Thursday, October 21th.
'KAWA - That'll be ten dollars, thank you drive through fuck you.

'KAWA - Nifty Icon support coming soon.

Private Type tAttack
  Move As Byte
  Level As Byte
End Type

'KAWA - Why did you encapsulate a single element in a type?
Private Type tAp
  lOffset As Long
End Type

Private AttackAtt(255) As tAttack
Private Selected As Integer
Private attackpnt() As Long
Private AttackPnts(&H19A) As tAp

Dim RomData As Integer

Private Sub cboAttack_Click()
  'KAWA - Used to be cboAttack_Change but VB is bugged on that.
  'KAWA - Use Click, which triggers even on mousewheel usage!
  AttackAtt(hsbAttack).Move = cboAttack.ListIndex + 1
  
  'KAWA - 30 seconds later I find out PATTED -finally- saves my Tackle->Trash change
  'KAWA - for Bulbasaur and ??????.
End Sub

Private Sub cmdSave_Click()
  Dim headr As String * 4
  
  'KAWA - I left the original code in place for nostalgic purposes. Do what you must.
'  Get #1, &HAD, headr
'  If headr = "AXVE" Then
'    AttackNameList = &H1F832D
'    AttackTable = &H207BC8
'  ElseIf headr = "AXPE" Then
'    AttackTable = &H207B58
'    AttackNameList = &H1F82BD
'  ElseIf headr = "BPGE" Then
'    AttackTable = &H257470
'    AttackNameList = &H24707D
'  ElseIf headr = "BPRE" Then
'    AttackTable = &H25D7B4
'    AttackNameList = &H2470A1
'  End If
  Seek #1, Roms(RomData).AttackTable + 1
  For i = 0 To 410
    Put #1, , AttackPnts(i).lOffset
  Next i
  Seek #1, txtRepoint.Text + 1
  For i = 0 To hsbAttack.Max
    If AttackAtt(i).Move = &HFF & AttackAtt(i).Level = &HFF Then Exit For
    Put #1, , AttackAtt(i)
  Next i
End Sub

Private Sub Form_Load()
  hsbAttack.Enabled = False 'because you DON'T want to do that in Design mode. VB bug.
  mnuOpen_Click
  SetIcon Me.hwnd, "AAA", True
End Sub

Private Sub hsbAttack_Change()
  If AttackAtt(hsbAttack).Level = &HA Then
    txtLevel = AttackAtt(hsbAttack).Level = 2
  Else
    txtLevel = AttackAtt(hsbAttack).Level / 2
  End If
  cboAttack.ListIndex = AttackAtt(hsbAttack).Move - 1
End Sub

Private Sub lstOffsets_Click()
  Seek #1, txtRepoint.Text + 1
  For i = 0 To 255
    Get #1, , AttackAtt(i)
    If AttackAtt(i).Move = &HFF Then Exit For
  Next i
  hsbAttack.Max = i - 1
  hsbAttack.Value = 0
  hsbAttack_Change
End Sub

Private Sub lstSpecies_Click()
  Selected = lstSpecies.ListIndex
  txtRepoint = "&H" & Hex(AttackPnts(Selected).lOffset - &H8000000)
  lstOffsets.ListIndex = Selected
End Sub

Private Sub mnuOpen_Click()
  Dim headr As String * 4
  Dim TheFile As String
  Dim Filler As String * 30
  Dim AttackName As String * 13
  Dim AttackNameList As Long
  Dim Attacks As Long
  Dim AttackTable As Long
  Dim pokename As String * 11
  Dim i As Long
  
  InitDatabase
  On Error Resume Next
  
  'KAWA - I could rewrite this to the VBAccel CDlg Class...
'  With cdlOpen
'    .DialogTitle = "Open Romfile"
'    .Filter = "GBA roms| *.GBA"
'    .ShowOpen
'    TheFile = .Filename
'    If Len(.Filename) = 0 Then
'      MsgBox "No rom loaded"
'      Unload Me
'    End If
'  End With

  'KAWA - Here we gooooo! </mario>
  Dim cc As cCommonDialog
  Set cc = New cCommonDialog
  If cc.VBGetOpenFileName(TheFile, , , , , , "GBA roms (*.gba)|*.gba", , App.Path, , , Me.hwnd, OFN_HIDEREADONLY) = False Then
    MsgBox "No ROM loaded"
    End
  End If

  'KAWA - Now ALL EM Toolchain apps can check the Lockout!
  CheckLock TheFile

  Open TheFile For Binary As #1
  Get #1, &HAD, headr
  RomData = FindRom(headr)
  If RomData = -1 Then
    MsgBox "Unsupported ROM."
  End If
  txtTheFile = TheFile
  
  'KAWA - I could rewrite this to 100% PokeRoms...
'  If headr = "AXVE" Then
'    AttackNameList = &H1F832D
'    AttackTable = &H207BC8
'  ElseIf headr = "AXPE" Then
'    AttackTable = &H207B58
'    AttackNameList = &H1F82BD
'  ElseIf headr = "BPGE" Then
'    AttackTable = &H257470
'    AttackNameList = &H24707D
'  ElseIf headr = "BPRE" Then
'    AttackTable = &H25D7B4
'    AttackNameList = &H2470A1
'  End If
  
  Seek #1, Roms(RomData).MonsterNames + 1
  Get #1, , i
  Seek #1, i - &H8000000 + 1
  lstSpecies.Clear 'KAWA - In all your editors, the list keeps growing and growing...
  For i = 0 To 410
    Get #1, , pokename
    z = Replace(Sapp2Asc(pokename), "\x", "")
    lstSpecies.AddItem z
  Next i
  Seek #1, Roms(RomData).AttackNameList + 1 'AttackNameList + 1
  cboAttack.Clear
  For i = 1 To &HFE
    Get #1, , AttackName
    z = Replace(Sapp2Asc(AttackName), "\x", "")
    cboAttack.AddItem z
  Next i
  Seek #1, Roms(RomData).AttackTable + 1 'AttackTable + 1
  For i = 0 To 410
    Get #1, , AttackPnts(i)
    lstOffsets.AddItem Hex(i)
  Next i
  
  'KAWA - Now that we have data to edit, we re-enable EVERY control and auto-select Bulbasaur.
  For Each Control In Me.Controls
    Control.Enabled = True
  Next
  txtRepoint.Enabled = False
  lstSpecies.ListIndex = 1
End Sub

Private Sub txtLevel_Change()
  'KAWA - How rare...
  AttackAtt(hsbAttack).Level = txtLevel * 2
End Sub

Private Sub txtRepoint_LostFocus()
  txtRepoint.Text = "&H" & Right("000000" & Hex(Val(txtRepoint)), 6)
  'KAWA - "If vbYes" is always true. Object Browser says it's 6. 6 is nonzero, so true.
  'KAWA - Therefore, even if you click No, like I did, it'll DO yes.
  'KAWA - The proper form would be "If MsgBox("bla", vbYesNo) = vbYes then".
  'If vbYes Then
  a = Val(InputBox("What is the new value of attacks for this pokemon?"))
  b = Val(InputBox("What is the new offset?"))
  Seek #1, b + 1
  Put #1, , &H2 * a & &HFF & &HFF
  AttackPnts(lstSpecies.ListIndex).lOffset = b + &H8000000
  'End If
End Sub
