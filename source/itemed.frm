VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Create STDITEMS.RBH"
      Height          =   735
      Left            =   2760
      TabIndex        =   15
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   6240
      MaxLength       =   11
      TabIndex        =   12
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtPrice 
      Height          =   285
      Left            =   6240
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   120
      Width           =   1935
   End
   Begin VB.ComboBox cboBag 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "itemed.frx":0000
      Left            =   3360
      List            =   "itemed.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtPointer2 
      Height          =   285
      Left            =   6240
      MaxLength       =   11
      TabIndex        =   6
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox txtPointer1 
      Height          =   285
      Left            =   3360
      MaxLength       =   11
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ListBox lstItems 
      Height          =   3015
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3360
      MaxLength       =   14
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblDescription 
      Caption         =   "Label6"
      Height          =   735
      Left            =   6240
      TabIndex        =   14
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Description"
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Price"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Pocket"
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "???"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "???"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblType 
      Caption         =   "Name"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type tItem
  sName As String * 14
  b15 As Byte
  b16 As Byte
  iPrice As Integer
  b19 As Byte
  b20 As Byte
  pDescription As Long
  b25 As Byte
  b26 As Byte
  bBagPocket As Byte
  b28 As Byte
  pUnknown1 As Long
  b33 As Byte
  b34 As Byte
  b35 As Byte
  b36 As Byte
  pUnknown2 As Long
  b41 As Byte
  b42 As Byte
  b43 As Byte
  b44 As Byte
End Type

Private Items(&H15C) As tItem
Private sel As Integer

Private Sub Command1_Click()
  Dim i As Integer
  Open "Ruby.gba" For Binary As #3
  Seek #3, &H3C5564 + 3
  For i = 0 To &H15C
    Put #3, , Items(i)
  Next i
End Sub

Private Sub Command2_Click()
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim l As String
  Open "stditems.rbh" For Output As #16
  For i = 0 To lstItems.ListCount - 1
    l = Mid(lstItems.List(i), 6)
    l = Replace(l, "é", "E")
    l = Replace(l, " ", "")
    
    If l <> "????????" Then
      Print #16, "#define ITEM_" & l & vbTab & IIf(Len(l) <= 10, vbTab, "") & "0x" & Hex(i)
      j = j + 1
      If j = 9 Then
        Print #16, ""
        j = 0
      End If
    End If
  Next i
  Close #16
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Open "Obsidian.gba" For Binary As #1
  Seek #1, &H3C5564 + 1
  For i = 0 To &H15C
    Get #1, , Items(i)
    lstItems.AddItem Right("00" & Hex(i), 3) & ". " & Replace(Sapp2Asc(Items(i).sName), "\x", "")
  Next i
  lstItems.ListIndex = 0
  'lstItems_Click
End Sub

Private Sub lstItems_Click()
  sel = lstItems.ListIndex
  txtName = Trim(Replace(Replace(Sapp2Asc(Items(sel).sName), "\x", ""), "î", ""))
  txtPointer1 = "&H" & Right("00000000" & Hex(Items(sel).pUnknown1), 8)
  txtPointer2 = "&H" & Right("00000000" & Hex(Items(sel).pUnknown2), 8)
  cboBag.ListIndex = Items(sel).bBagPocket - 1
  txtPrice = Items(sel).iPrice
  txtDescription = "&H" & Right("00000000" & Hex(Items(sel).pDescription), 8)
  
  Dim s As String * 256
  Dim t As String
  Open "Ruby.gba" For Binary As #2
  Get #2, Items(sel).pDescription - &H8000000 + 1, s
  t = Sapp2Asc(s)
  t = Left(t, InStr(1, s, Chr(255), vbBinaryCompare) + 1)
  lblDescription = Replace(t, "\n", vbCrLf)
  Close #2
End Sub

Private Sub lstItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Dim ss As String
    Dim i As Integer
    On Error Resume Next
    ss = UCase(InputBox("Find..."))
    For i = lstItems.ListIndex + 1 To lstItems.ListCount - 1
      If InStr(lstItems.List(i), ss) Then
        lstItems.ListIndex = i
        Exit Sub
      End If
    Next i
    MsgBox "Not found."
  End If
End Sub

Private Sub txtName_LostFocus()
  Items(sel).sName = Asc2Sapp(UCase(Replace(txtName, "î", "")) & "\x") & String(20, Chr$(0))
  lstItems.List(sel) = Right("00" & Hex(sel), 3) & ". " & Replace(Sapp2Asc(Items(sel).sName), "\x", "")
End Sub

Private Sub txtPointer1_LostFocus()
  txtPointer1 = "&H" & Right("00000000" & Hex(Val(txtPointer1)), 8)
  Items(sel).pUnknown1 = txtPointer1
End Sub

Private Sub txtPointer2_LostFocus()
  txtPointer2 = "&H" & Right("00000000" & Hex(Val(txtPointer2)), 8)
  Items(sel).pUnknown2 = txtPointer2
End Sub

Private Sub txtDescription_LostFocus()
  txtDescription = "&H" & Right("00000000" & Hex(Val(txtDescription)), 8)
  Items(sel).pDescription = txtDescription
End Sub

Private Sub txtPrice_Change()
  txtPrice = CInt(Val(txtPrice))
  Items(sel).iPrice = txtPrice
End Sub
