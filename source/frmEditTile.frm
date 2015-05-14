VERSION 5.00
Begin VB.Form frmEditTile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Highly Experimental Map16 editor v2.1"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Behavior"
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   4695
      Begin VB.ComboBox cboBehavior 
         Height          =   315
         ItemData        =   "frmEditTile.frx":0000
         Left            =   120
         List            =   "frmEditTile.frx":0301
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   4455
      End
      Begin VB.CheckBox chkDrawOnTop 
         Caption         =   "Top overlaps player"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Top layer"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   2535
      Begin EliteMap.Map16TileEd Map16TileEd1 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
      End
      Begin EliteMap.Map16TileEd Map16TileEd1 
         Height          =   615
         Index           =   5
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
      End
      Begin EliteMap.Map16TileEd Map16TileEd1 
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
      End
      Begin EliteMap.Map16TileEd Map16TileEd1 
         Height          =   615
         Index           =   7
         Left            =   1320
         TabIndex        =   12
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bottom layer"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin EliteMap.Map16TileEd Map16TileEd1 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
      End
      Begin EliteMap.Map16TileEd Map16TileEd1 
         Height          =   615
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
      End
      Begin EliteMap.Map16TileEd Map16TileEd1 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
      End
      Begin EliteMap.Map16TileEd Map16TileEd1 
         Height          =   615
         Index           =   3
         Left            =   1320
         TabIndex        =   8
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Fixed a LOT of Behavior stuff here, adding some that were missing. I was sooo sloppy back then ^_^               --- Kawa"
      Height          =   1695
      Left            =   2760
      TabIndex        =   16
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "TODO: Add preview window?"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
End
Attribute VB_Name = "frmEditTile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TileAddress As Long
Public BehaviorAddress As Long
Public Filename As String

Public beDrawOnTop As Byte
Public beHavior As Byte

Public Sub LoadTile()
  Dim tile As Integer
  
  'Debug.Print "Behavior: " & Hex(BehaviorAddress)
  
  Open Filename For Binary As #69
  Seek #69, TileAddress + 1
  For i = 0 To 7
    Get #69, , tile
    'txtTile(i).Text = "&H" & Right("0000" & Hex(tile), 4)
    Map16TileEd1(i).Value = tile
  Next i
  'Label3 = Label3 & TileAddress
  Seek #69, BehaviorAddress + 1
  Get #69, , beHavior
  Get #69, , beDrawOnTop
  Close #69
  cboBehavior.ListIndex = beHavior
  If beDrawOnTop = &H0 Then
    chkDrawOnTop.Value = 1
  ElseIf beDrawOnTop = &H10 Then
    chkDrawOnTop.Value = 0
  Else
    chkDrawOnTop.Value = 2
  End If
End Sub

Private Sub Command1_Click()
  Open Filename For Binary As #69
  Seek #69, TileAddress + 1
  For i = 0 To 7
    Put #69, , CInt(Val(Map16TileEd1(i).Value))
    'txtTile(i).Text = "&H" & Right("0000" & Hex(tile), 4)
  Next i
  Seek #69, BehaviorAddress + 1
  Put #69, , CByte(cboBehavior.ListIndex)
  If chkDrawOnTop.Value = 1 Then
    Put #69, , CByte(0)
  ElseIf chkDrawOnTop.Value = 0 Then
    Put #69, , CByte(&H10)
  Else
   'Put #69, , CByte(whatever the fuck it used to be)
  End If
  Close #69
  Unload Me
End Sub

Private Sub Command2_Click()
  Form1.lblRom.Caption = "cancelled"
  Unload Me
End Sub

