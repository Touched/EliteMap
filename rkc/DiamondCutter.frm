VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diamond Cutter"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "DiamondCutter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   249
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   241
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox filFiles 
      Height          =   1455
      IntegralHeight  =   0   'False
      ItemData        =   "DiamondCutter.frx":030A
      Left            =   120
      List            =   "DiamondCutter.frx":030C
      MultiSelect     =   2  'Extended
      TabIndex        =   10
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   360
      Width           =   1215
   End
   Begin VB.FileListBox filFiles2 
      Appearance      =   0  'Flat
      Height          =   420
      Left            =   240
      MultiSelect     =   2  'Extended
      Pattern         =   "*.RBC"
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fraOptions 
      Caption         =   "&Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   2055
      Begin VB.CheckBox chkWithSTDPoke 
         Caption         =   "Include STD&Poke"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chkWithSTD 
         Caption         =   "Include &STD"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CheckBox chkWithSTDItems 
         Caption         =   "Include STD&Items"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.TextBox txtROM 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblFiles 
      Caption         =   "&Code file(s):"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblROM 
      Caption         =   "&ROM:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  Open "DiamondCutter.cfg" For Output As #1
    Print #1, txtROM.Text
    Print #1, chkWithSTD.Value
    Print #1, chkWithSTDItems.Value
    Print #1, chkWithSTDPoke.Value
  Close #1
  End
End Sub

Private Sub cmdOK_Click()
  Dim MyShell As String
  Dim Files As String
  Dim i As Integer
  Dim Lol As String
  
  'Create space-delimited list of filenames
  For i = 0 To filFiles.ListCount - 1
    If filFiles.Selected(i) = True Then
      Files = Files & " " & filFiles.List(i)
    End If
  Next i
  'If empty list, no files selected
  If Trim(Files) = "" Then
    MsgBox "No files selected.", vbInformation
    Exit Sub
  End If
  
  Screen.MousePointer = 11
  Caption = "DC - Working..."
  
  'Prepare string
  MyShell = "rkc.exe /o " & Trim(txtROM.Text)
 
  If chkWithSTD.Value = 1 Then MyShell = MyShell & " std.rbh"
  If chkWithSTDItems.Value = 1 Then MyShell = MyShell & " stditems.rbh"
  If chkWithSTDPoke.Value = 1 Then MyShell = MyShell & " stdpoke.rbh"
 
  MyShell = MyShell & Files
  i = Shell(MyShell, vbHide)

  Screen.MousePointer = 0
  Caption = "Diamond Cutter"
End Sub

Private Sub filFiles_DblClick()
  cmdOK_Click
End Sub

Private Sub Form_Load()
  SetIcon Me.hWnd, "AAA", True
  
  Dim InBuff As String
  Dim i As Long
  
  For i = 0 To filFiles2.ListCount - 1
    filFiles.AddItem filFiles2.List(i)
  Next i
  
  'This block checks for Rubikon to be present
  On Error GoTo NoRKC
  Open "rkc.exe" For Input As #1
  'Must use Input instead of Binary. Binary would
  'instantly create an empty file if it's not here,
  'with VERY stupid results.
  Close #1
  
  'This block checks for the stdlib to be present
  On Error GoTo NoSTD
  Open "std.rbh" For Input As #1
  Close #1
  On Error GoTo NoSTDItems
  Open "stditems.rbh" For Input As #1
  Close #1
  On Error GoTo NoSTDPoke
  Open "stdpoke.rbh" For Input As #1
  Close #1
  
  On Error GoTo NoConfig
  Open "DiamondCutter.cfg" For Input As #1
    Input #1, InBuff: txtROM.Text = InBuff
    Input #1, InBuff: chkWithSTD.Value = Val(InBuff)
    Input #1, InBuff: chkWithSTDItems.Value = Val(InBuff)
    Input #1, InBuff: chkWithSTDPoke.Value = Val(InBuff)
  Close #1
  If txtROM.Text = "" Then txtROM.Text = "rkc.bin"

NoConfig:
  Exit Sub

NoRKC:
  MsgBox "Diamond Cutter cannot be run without the Rubikon compiler." & vbCrLf & _
         "Make sure the two are in the same directory and try again.", vbExclamation, _
         "RKC.EXE not found"
  End
  
NoSTD:
  chkWithSTD.Value = 0
  chkWithSTD.Enabled = False
  Resume Next
NoSTDItems:
  chkWithSTDItems.Value = 0
  chkWithSTDItems.Enabled = False
  Resume Next
NoSTDPoke:
  chkWithSTDPoke.Value = 0
  chkWithSTDPoke.Enabled = False
  Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  cmdExit_Click
End Sub
