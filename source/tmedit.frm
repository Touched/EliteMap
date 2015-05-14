VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Experimental TM viewer/editor thingy"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
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
   ScaleHeight     =   360
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBits 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5040
      Width           =   6135
   End
   Begin VB.Frame fraRawValues 
      Caption         =   "Raw values"
      Height          =   1215
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1095
      Begin VB.Label lblRawValues 
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.ListBox lstPokemon 
      Height          =   4815
      IntegralHeight  =   0   'False
      ItemData        =   "tmedit.frx":0000
      Left            =   120
      List            =   "tmedit.frx":04D8
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.ListBox lstAttacks 
      Height          =   4815
      IntegralHeight  =   0   'False
      ItemData        =   "tmedit.frx":13C1
      Left            =   3840
      List            =   "tmedit.frx":1473
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  lstPokemon.ListIndex = 1
End Sub

'Ruby:      1FD0F8
'Sapphire:  1FD088
'Fire Red:  252BD0 of 252BD8
'Bulbasaur: 2007 3584 081E E400
'10000000 00011100 11010110 00010000 00100000 01111011 10010000 000000

Private Sub lstPokemon_Click()
  Dim TMList(1 To 8) As Byte
  Dim i As Integer, j As Integer, k As Integer
  
  Open "Ruby.gba" For Binary As #1
  Seek #1, &H1FD0F0 + (lstPokemon.ListIndex * 8) + 1
  Get #1, , TMList
  Close #1
  
  'lblRawValues = PadHex(TMList(1), 8) & vbCrLf & PadHex(TMList(2), 9)
  
  txtBits = ""
  For j = 1 To 8
    For i = 0 To 7
      If BitIsSet(TMList(j), i) = True Then
        'lstAttacks.Selected(i) = True
        txtBits = txtBits & "1"
      Else
        'lstAttacks.Selected(i) = False
        txtBits = txtBits & "0"
      End If
    Next i
  Next j
  txtBits = Mid(txtBits, 3)
  For i = 1 To 58
    If Mid(txtBits, i, 1) = "1" Then
      lstAttacks.Selected(i - 1) = True
    Else
      lstAttacks.Selected(i - 1) = False
    End If
  Next i
  lstAttacks.ListIndex = 0
End Sub
