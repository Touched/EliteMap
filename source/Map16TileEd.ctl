VERSION 5.00
Begin VB.UserControl Map16TileEd 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HitBehavior     =   0  'None
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   Begin VB.CheckBox chkVFlip 
      Caption         =   "V"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.CheckBox chkHFlip 
      Caption         =   "H"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox txtTile 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Text            =   "FFF"
      Top             =   0
      Width           =   495
   End
   Begin VB.ComboBox cboPal 
      Height          =   315
      ItemData        =   "Map16TileEd.ctx":0000
      Left            =   0
      List            =   "Map16TileEd.ctx":0034
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Map16TileEd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_Value As Integer
Event Change()
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."

Public Property Get Value() As Integer
  Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
  m_Value = New_Value
  PropertyChanged "Value"
  RaiseEvent Change
  haha = Right("0000" & Hex(m_Value), 4)
  cboPal.ListIndex = Val("&H" & Mid(haha, 1, 1))
  If BitIsSet(Val("&H" & haha), 10) Then
    chkHFlip.Value = 1
  Else
    chkHFlip.Value = 0
  End If
  If BitIsSet(Val("&H" & haha), 11) Then
    chkVFlip.Value = 1
  Else
    chkVFlip.Value = 0
  End If
  txtTile.Text = Mid(haha, 2, 3)
End Property

Private Sub cboPal_LostFocus()
  If Val("&H" & txtTile) > &HFFF Then txtTile = "FFF"
  haha = cboPal.List(cboPal.ListIndex) & _
         txtTile.Text
  m_Value = Val("&H" & haha)
  PropertyChanged "Value"
  RaiseEvent Change
End Sub

Private Sub chkHFlip_Click()
  If chkHFlip.Value = 1 Then
    m_Value = BitSet(CLng(m_Value), 10)
  Else
    m_Value = BitClear(CLng(m_Value), 10)
  End If
  haha = Right("0000" & Hex(m_Value), 4)
  txtTile.Text = Mid(haha, 2, 3)
  PropertyChanged "Value"
End Sub

Private Sub chkVFlip_Click()
  If chkVFlip.Value = 1 Then
    m_Value = BitSet(CLng(m_Value), 11)
  Else
    m_Value = BitClear(CLng(m_Value), 11)
  End If
  haha = Right("0000" & Hex(m_Value), 4)
  txtTile.Text = Mid(haha, 2, 3)
  PropertyChanged "Value"
End Sub

Private Sub txtTile_LostFocus()
  If Val("&H" & txtTile) > &HFFF Then txtTile = "FFF"
  haha = cboPal.List(cboPal.ListIndex) & _
         txtTile.Text
  m_Value = Val("&H" & haha)
  PropertyChanged "Value"
  RaiseEvent Change
End Sub

Private Sub UserControl_InitProperties()
  m_Value = 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  m_Value = PropBag.ReadProperty("Value", m_def_Value)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
End Sub


Private Function BitSet(Number As Long, ByVal Bit As Long) As Long
  If Bit = 31 Then
    Number = &H80000000 Or Number
  Else
    Number = (2 ^ Bit) Or Number
  End If
  BitSet = Number
End Function

Private Function BitClear(Number As Long, ByVal Bit As Long) As Long
  If Bit = 31 Then
    Number = &H7FFFFFFF And Number
  Else
    Number = ((2 ^ Bit) Xor &HFFFFFFFF) And Number
  End If
  BitClear = Number
End Function

Private Function BitIsSet(ByVal Number As Long, ByVal Bit As Long) As Boolean
  BitIsSet = False
  If Bit = 31 Then
    If Number And &H80000000 Then BitIsSet = True
  Else
    If Number And (2 ^ Bit) Then BitIsSet = True
  End If
End Function
