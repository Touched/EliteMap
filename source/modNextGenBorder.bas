Attribute VB_Name = "modNextGenBorder"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020
Public Borderitems(64) As Long
Public border_sizeX As Byte
Public border_sizeY As Byte
Public border(0 To 1, 0 To 1) As Long
Public NextGen As Boolean
Public dirty2 As Boolean
Public seltile(0 To 2) As Long
Public selattr(0 To 2) As Long
Public txtbank As String
Public txtlevel As String
Public worldentry As Variant
Public exit2 As Boolean
Public allheadersize As Variant
Public newheadersize As Variant
Public noresize As Boolean
Public newwm As Variant
Public newhm As Variant
Public newwb As Variant
Public newhb As Variant
Public lheight As Variant
Public thismap As MapHeader
Public oldmap As MapHeader
Public Type MapHeader
  wWidth As Long
  wHeight As Long
  pBorder As Long
  pMap As Long
  pTilesetA As Long
  pTilesetB As Long
  bBorderX As Byte
  bBorderY As Byte
End Type

Public Sub Set_Border(tempX, tempY)
border_sizeX = tempX
border_sizeY = tempY
Form1.cmdFullBRD.Enabled = IIf(NextGen = True, True, False)
If border_sizeX = 2 Then
    If border_sizeY = 2 Then GoTo use_old_style
End If
If NextGen = False Then GoTo use_old_style
Form1.picBorder.Visible = False
Form1.cmdFullBRD.Enabled = False
Form1.cmdFullBRD.Visible = True
Form1.cmdFullBRD.Enabled = True
Exit Sub
use_old_style:
Form1.picBorder.Visible = True
Form1.cmdFullBRD.Visible = False
Form1.cmdFullBRD.Enabled = False
  Form1.drawtilehdc border(0, 0), Form1.picBorder.hdc, 0, 0
  Form1.drawtilehdc border(0, 1), Form1.picBorder.hdc, 1, 0
  Form1.drawtilehdc border(1, 0), Form1.picBorder.hdc, 0, 1
  Form1.drawtilehdc border(1, 1), Form1.picBorder.hdc, 1, 1
End Sub

Public Sub draw_ng_border()
yind = 0
For i = 0 To border_sizeY - 1
xind = 0
    For ii = 0 To border_sizeX - 1
        Form1.drawtilehdc Borderitems(yind * border_sizeX + xind) Mod &H400, AdvancedBorder.picBorder.hdc, xind, yind 'read variables out and draw them
        xind = xind + 1
    Next ii
    yind = yind + 1
Next i
AdvancedBorder.picBorder.Refresh
End Sub
