VERSION 5.00
Begin VB.UserControl GBATileEditor 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1995
   ScaleHeight     =   133
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   133
   ToolboxBitmap   =   "kawatileedit.ctx":0000
End
Attribute VB_Name = "GBATileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim TileData(8, 8) As Byte

Public Enum teBorder
  teNone = 0
  teInset = 1
End Enum
Event Changed()
Attribute Changed.VB_Description = "Sent whenever the user draws a pixel."
Const m_def_ShowGrid = 0
Const m_def_Filename = ""
Const m_def_PenColor = 15
Const m_def_RomAddress = 0
Const m_def_DotSize = 16
Dim m_ShowGrid As Boolean
Dim m_Filename As String
Dim m_PenColor As Byte
Dim m_RomAddress As Long
Dim m_DotSize As Integer
Dim Palette(0 To 15) As Long
 
Public Property Get Colors(ByVal pali As Integer) As OLE_COLOR
Attribute Colors.VB_Description = "An array of colors to draw with."
  If pali > 15 Then Err.Raise 1601, "Kawa's Tile Editor", "You only have 16 colors!"
  Colors = Palette(pali)
End Property
Public Property Let Colors(ByVal pali As Integer, newColor As OLE_COLOR)
  If pali > 15 Then Err.Raise 1601, "Kawa's Tile Editor", "You only have 16 colors!"
  Palette(pali) = newColor
  PropertyChanged "Colors"
End Property
 
Public Property Get BorderStyle() As teBorder
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As teBorder)
  If New_BorderStyle > teInset Then Err.Raise 1600, "Kawa's Tile Editor", "Borderstyle can't be higher than 1."
  UserControl.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

Public Property Get DotSize() As Integer
Attribute DotSize.VB_Description = "Returns/sets the scale of drawing."
  DotSize = m_DotSize
End Property
Public Property Let DotSize(ByVal New_DotSize As Integer)
  m_DotSize = New_DotSize
  PropertyChanged "DotSize"
  UserControl_Resize
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
  Dim x As Integer, y As Integer
  For x = 0 To 7
    For y = 0 To 7
      Line (x * m_DotSize, y * m_DotSize)-((x + 1) * m_DotSize, (y + 1) * m_DotSize), Palette(TileData(x, y)), BF
    Next y
  Next x
  DoGrid
End Sub

Private Sub DoGrid()
  Dim x As Integer, y As Integer
  If m_ShowGrid = True Then
    For x = 1 To 7
      For y = 1 To 7
        Line (x * m_DotSize, 0)-(x * m_DotSize, 10024), QBColor(15)
        Line (0, y * m_DotSize)-(10024, y * m_DotSize), QBColor(15)
      Next y
    Next x
  End If
End Sub

Public Sub SetPalette(NewPal() As Long)
Attribute SetPalette.VB_Description = "Set all 16 colors at once."
  Dim i As Integer
  For i = 0 To 15
    Palette(i) = NewPal(i)
  Next i
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
  Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
  UserControl.Enabled() = New_Enabled
  PropertyChanged "Enabled"
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Tag = "^_^"
  UserControl_MouseMove Button, Shift, x, y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim px As Integer, py As Integer
  Dim ax As Integer, ay As Integer
  
  If Tag <> "^_^" Then Exit Sub
  px = Int(x / m_DotSize)
  py = Int(y / m_DotSize)
  If px < 0 Or px > 7 Then Exit Sub
  If py < 0 Or py > 7 Then Exit Sub
  ax = px * m_DotSize
  ay = ay * m_DotSize
  TileData(px, py) = IIf(Button = 2, 0, PenColor)
      
  Line (px * m_DotSize, py * m_DotSize)-((px + 1) * m_DotSize, (py + 1) * m_DotSize), Palette(TileData(px, py)), BF
  DoGrid
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Tag = "-_-"
End Sub

Private Sub UserControl_Paint()
  Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
  m_RomAddress = PropBag.ReadProperty("RomAddress", m_def_RomAddress)
  m_PenColor = PropBag.ReadProperty("PenColor", m_def_PenColor)
  m_Filename = PropBag.ReadProperty("Filename", m_def_Filename)
  m_ShowGrid = PropBag.ReadProperty("ShowGrid", m_def_ShowGrid)
  m_DotSize = PropBag.ReadProperty("DotSize", m_def_DotSize)
End Sub

Private Sub UserControl_Resize()
  'Width = 1995
  'Height = 1995
  Width = ((m_DotSize * 8) + IIf(UserControl.BorderStyle = 1, 4, 0)) * Screen.TwipsPerPixelX
  Height = ((m_DotSize * 8) + IIf(UserControl.BorderStyle = 1, 4, 0)) * Screen.TwipsPerPixelY
  Refresh
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
  Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
  Call PropBag.WriteProperty("RomAddress", m_RomAddress, m_def_RomAddress)
  Call PropBag.WriteProperty("PenColor", m_PenColor, m_def_PenColor)
  Call PropBag.WriteProperty("Filename", m_Filename, m_def_Filename)
  Call PropBag.WriteProperty("ShowGrid", m_ShowGrid, m_def_ShowGrid)
  Call PropBag.WriteProperty("DotSize", m_DotSize, m_def_DotSize)
End Sub

Public Sub LoadTileData()
Attribute LoadTileData.VB_Description = "Load data from the file and address specified and show it."
  Dim RawData(32) As Byte
  Dim ff As Long
  Dim i As Integer
  Dim color1 As Integer
  Dim color2 As Integer
  Dim s As Integer
  Dim x As Integer, y As Integer
  
  ff = FreeFile
  Open m_Filename For Binary As ff
    Get ff, m_RomAddress + 1, RawData
  Close ff
  
  For i = 0 To 31
    color1 = RawData(i)
    For s = 0 To 3: color1 = color1 \ 2: Next s 'simulates color1 >> 4
    color1 = color1 And &HF
    color2 = RawData(i) And &HF
    
    TileData(x + 1, y) = color1
    TileData(x, y) = color2
    x = x + 2
    If x > 7 Then
      x = 0
      y = y + 1
    End If
  Next i
  
  Refresh
End Sub

Public Sub SaveTileData()
Attribute SaveTileData.VB_Description = "Save the tile as-is at the specified location in the specified file."
  Dim RawData(32) As Byte
  Dim i As Integer
  Dim color1 As Integer
  Dim color2 As Integer
  Dim x As Integer, y As Integer
  Dim ff As Long
  
  For i = 0 To 31
    color1 = TileData(x + 1, y)
    color2 = TileData(x, y)
    RawData(i) = CByte(Val("&H" & Hex(color1) & Hex(color2)))
    Debug.Print Hex(RawData(i))
    x = x + 2
    If x > 7 Then
      x = 0
      y = y + 1
    End If
  Next i
  
  ff = FreeFile
  Open m_Filename For Binary As ff
    Put ff, m_RomAddress + 1, RawData
  Close ff
  
End Sub

Public Property Get RomAddress() As Long
Attribute RomAddress.VB_Description = "The address in the file stored in Filename to read from."
  RomAddress = m_RomAddress
End Property
Public Property Let RomAddress(ByVal New_RomAddress As Long)
  m_RomAddress = New_RomAddress
  PropertyChanged "RomAddress"
End Property

Private Sub UserControl_InitProperties()
  m_RomAddress = m_def_RomAddress
  m_PenColor = m_def_PenColor
  m_Filename = m_def_Filename
  m_ShowGrid = m_def_ShowGrid
End Sub

Public Property Get PenColor() As Byte
Attribute PenColor.VB_Description = "The palette index to draw with."
Attribute PenColor.VB_UserMemId = 0
  PenColor = m_PenColor
End Property
Public Property Let PenColor(ByVal New_PenColor As Byte)
  m_PenColor = New_PenColor
  PropertyChanged "PenColor"
End Property

Public Property Get Filename() As String
Attribute Filename.VB_Description = "The file name to read from."
  Filename = m_Filename
End Property
Public Property Let Filename(ByVal New_Filename As String)
  m_Filename = New_Filename
  PropertyChanged "Filename"
End Property

Public Property Get ShowGrid() As Boolean
Attribute ShowGrid.VB_Description = "Toggle drawing a white grid."
  ShowGrid = m_ShowGrid
End Property
Public Property Let ShowGrid(ByVal New_ShowGrid As Boolean)
  m_ShowGrid = New_ShowGrid
  PropertyChanged "ShowGrid"
End Property

