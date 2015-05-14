VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBoxMan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Boxman"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
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
   MaxButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   473
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdPaste 
      Caption         =   "&Paste"
      Height          =   375
      Left            =   4920
      TabIndex        =   19
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "&Kill"
      Height          =   375
      Left            =   6000
      TabIndex        =   18
      Top             =   3480
      Width           =   975
   End
   Begin VB.Frame fraBoxes 
      Caption         =   "&Boxes"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   3615
      Begin VB.ComboBox cboBoxWallpaper 
         Height          =   315
         ItemData        =   "boxman.frx":0000
         Left            =   1440
         List            =   "boxman.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtBoxName 
         Height          =   285
         Left            =   1440
         MaxLength       =   8
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblWallpaper 
         Caption         =   "&Wallpaper"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "&Name"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame fraPokemon 
      Caption         =   "&Pokémon"
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   5775
      Begin VB.CheckBox chkMark 
         Height          =   255
         Index           =   3
         Left            =   2160
         TabIndex        =   11
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkMark 
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   10
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkMark 
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkMark 
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   8
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtOTrainer 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   6
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblMarks 
         Caption         =   "&Marks"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblOTrainer 
         Caption         =   "&Original Trainer"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblName 
         Caption         =   "&Name"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox lstBoxes 
      Height          =   2790
      ItemData        =   "boxman.frx":00B6
      Left            =   6000
      List            =   "boxman.frx":00B8
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid flexGrid 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2143
      _Version        =   393216
      Rows            =   5
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmBoxMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tBoxMon
  lUnknown1 As Long
  iOTID As Integer
  iOTSID As Integer
  sName As String * 11
  bUnknown1 As Byte
  sOTName As String * 7
  bMarks As Byte
  sFiller As String * 52
End Type

Private BoxNames(13) As String * 9
Private BoxWallpapers(13) As Byte

Private BoxMon(13, 5, 4) As tBoxMon
Private Fist As tBoxMon
Private Killer As tBoxMon

Private CB As Integer
Private CC As Integer
Private CR As Integer

Private Sub chkMark_Click(Index As Integer)
  If chkMark(Index).Value = 0 Then
    BoxMon(CB, CC, CR).bMarks = BitClear(CLng(BoxMon(CB, CC, CR).bMarks), Index)
  Else
    BoxMon(CB, CC, CR).bMarks = BitSet(CLng(BoxMon(CB, CC, CR).bMarks), Index)
  End If
End Sub

Private Sub cmdCopy_Click()
  Fist = BoxMon(CB, CC, CR)
End Sub

Private Sub cmdKill_Click()
  BoxMon(CB, CC, CR) = Killer
End Sub

Private Sub cmdPaste_Click()
  BoxMon(CB, CC, CR) = Fist
  lstBoxes_Click
End Sub

Private Sub cmdSave_Click()
  Dim Box As Integer
  Dim Col As Integer
  Dim Row As Integer
  
  Seek #1, &H300A4 + 1
  For Box = 0 To 13
    For Row = 0 To 4
      For Col = 0 To 5
        Put #1, , BoxMon(Box, Col, Row)
      Next Col
    Next Row
  Next Box
  Seek #1, &H383E4 + 1
  For Box = 0 To 13
    Put #1, , BoxNames(Box)
  Next Box
  Seek #1, &H38462 + 1
  For Box = 0 To 13
    Put #1, , BoxWallpapers(Box)
  Next Box
End Sub

Private Sub flexGrid_Click()
  CC = flexGrid.Col
  CR = flexGrid.Row
  txtName = Sapp2Asc(Left(BoxMon(CB, CC, CR).sName, InStr(BoxMon(CB, CC, CR).sName & Chr$(255), Chr$(255)) - 1))
  txtOTrainer = Sapp2Asc(Left(BoxMon(CB, CC, CR).sOTName, InStr(BoxMon(CB, CC, CR).sOTName & Chr$(255), Chr$(255)) - 1))
  
  chkMark(0).Value = IIf(BitIsSet(BoxMon(CB, CC, CR).bMarks, 0), 1, 0)
  chkMark(1).Value = IIf(BitIsSet(BoxMon(CB, CC, CR).bMarks, 1), 1, 0)
  chkMark(2).Value = IIf(BitIsSet(BoxMon(CB, CC, CR).bMarks, 2), 1, 0)
  chkMark(3).Value = IIf(BitIsSet(BoxMon(CB, CC, CR).bMarks, 3), 1, 0)
End Sub

Private Sub Form_Load()
  Dim Box As Integer
  Dim Col As Integer
  Dim Row As Integer
  
  Open "ram.dmp" For Binary As #1
  Seek #1, &H300A4 + 1
  For Box = 0 To 13
    For Row = 0 To 4
      For Col = 0 To 5
        Get #1, , BoxMon(Box, Col, Row)
      Next Col
    Next Row
  Next Box
  Seek #1, &H383E4 + 1
  For Box = 0 To 13
    Get #1, , BoxNames(Box)
    lstBoxes.AddItem Sapp2Asc(Left(BoxNames(Box), InStr(BoxNames(Box) & Chr$(255), Chr$(255)) - 1))
  Next Box
  Seek #1, &H38462 + 1
  For Box = 0 To 13
    Get #1, , BoxWallpapers(Box)
  Next Box
  lstBoxes.ListIndex = 0
  flexGrid.Col = 0
  flexGrid.Row = 0
  flexGrid_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Close #1
End Sub

Private Sub lstBoxes_Click()
  Dim Row As Integer
  Dim Col As Integer
  CB = lstBoxes.ListIndex
  For Row = 0 To 4
    For Col = 0 To 5
      flexGrid.Row = Row
      flexGrid.Col = Col
      flexGrid.Clip = Sapp2Asc(Left(BoxMon(CB, Col, Row).sName, InStr(BoxMon(CB, Col, Row).sName & Chr$(255), Chr$(255)) - 1))
    Next Col
  Next Row
  flexGrid.Row = CR
  flexGrid.Col = CC
  txtBoxName = Sapp2Asc(Left(BoxNames(CB), InStr(BoxNames(CB) & Chr$(255), Chr$(255)) - 1))
  cboBoxWallpaper.ListIndex = BoxWallpapers(CB)
End Sub

Private Sub txtBoxName_LostFocus()
  If txtBoxName = "" Then txtBoxName = "BOX" & (CB + 1)
  BoxNames(CB) = Asc2Sapp(txtBoxName & "\x")
  lstBoxes.List(CB) = Sapp2Asc(Left(BoxNames(CB), InStr(BoxNames(CB) & Chr$(255), Chr$(255)) - 1))
End Sub

Private Sub txtName_LostFocus()
  BoxMon(CB, CC, CR).sName = Asc2Sapp(txtName & "\x")
  flexGrid.Row = CR
  flexGrid.Col = CC
  flexGrid.Clip = Sapp2Asc(Left(BoxMon(CB, CC, CR).sName, InStr(BoxMon(CB, CC, CR).sName & Chr$(255), Chr$(255)) - 1))
End Sub
