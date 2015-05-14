VERSION 5.00
Begin VB.Form frmRomSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "<appname>"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRomSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   3600
      Pattern         =   "*.gba"
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboRom 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "<romdata>"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Select ROM"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmRomSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboRom_Click()
  Dim header As String * 4
  Dim r As Integer
  Open cboRom.List(cboRom.ListIndex) For Binary As #1
  Get #1, &HAD, header
  Close #1
  r = FindRom(header)
  Label2 = Roms(r).Code & " - " & Roms(r).Name
End Sub

Private Sub Command1_Click()
  Hide
  CheckLock cboRom.List(cboRom.ListIndex)
  Tag = cboRom.List(cboRom.ListIndex)
End Sub

Private Sub Command2_Click()
  Hide
  Tag = "Cancelled"
End Sub

Private Sub Form_Load()
  Dim header As String * 4
  Dim i As Integer
  Dim r As Integer

  Caption = App.Title
  InitDatabase
  If File1.ListCount = 0 Then
    MsgBox "No GBA roms in current directory at all.", vbExclamation
    End
  End If
  For i = 0 To File1.ListCount - 1
    Open File1.List(i) For Binary As #1
    Get #1, &HAD, header
    Close #1
    r = FindRom(header)
    If r >= 0 Then
      cboRom.AddItem File1.List(i) 'Roms(r).Code & " - " & Roms(r).Name
    End If
  Next i
  If cboRom.ListCount = 0 Then
    MsgBox "No Pokémon roms in current directory.", vbExclamation
    End
  End If
  cboRom.ListIndex = 0
End Sub
