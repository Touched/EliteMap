Attribute VB_Name = "modGBAGfx"
'WARNING! THIS MODULE CONTAINS COARSE BITSHIFT HACKS AND USES API CALLS
Option Explicit

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Sub BlitTile(ByRef TileData() As Byte, ByVal destHDC As Long, ByVal destX As Long, ByVal destY As Long, ByRef pal() As Long)
  Dim i As Byte
  Dim color1 As Byte
  Dim color2 As Byte
  Dim x As Byte
  Dim y As Byte
  Dim s As Byte
  
  'This uses the "proper" decoding method, as far as possible in VB.
  For i = 0 To 31
    color1 = TileData(i)
    
    'You can't do bitshifts in VB. But since a bitshift to the right (as used here) is the same as
    'dividing, four divisions by 2 equal a bitshift by 4.
    For s = 0 To 3: color1 = color1 \ 2: Next s 'IOW, color1 = color1 >> 4
    
    color1 = color1 And &HF 'Having shifted color 1 INTO color2, we clear the old high-nibble.
    color2 = TileData(i) And &HF 'Color 2's easier because it needs no shifting.
    
    'Call the SetPixel API if you need speed. It's like the old QB45 days: Poke the video memory,
    'don't just PSET. Same goes here, VB's too damn slow so I prefer SetPixel above PSet any day.
    SetPixel destHDC, destX + x + 1, destY + y, pal(color1)
    SetPixel destHDC, destX + x, destY + y, pal(color2)

    'Set the next drawing coordinate accordingly
    x = x + 2
    If x > 7 Then
      x = 0
      y = y + 1
    End If
  Next i
End Sub

Public Sub UnPackPalette(ByRef palGBA() As Integer, ByRef palPC() As Long, Optional numcols As Integer = 15)
  Dim r As Integer
  Dim g As Integer
  Dim b As Integer
  Dim i As Integer
  Dim s As Integer

  'Whereas unpacking 4BPP graphics data requires only ONE bitshift (albeit a >>4), this one's SICK!!

  'Code taken from PokePic for conversion to hackish VB
  'r[i] = ((palbuf & 0x1F) <<3);
  'g[i] = ((palbuf >> 5) & 0x1F) <<3;
  'b[i] = (((palbuf >> 10) & 0x1f) <<3);

  For i = 0 To numcols
    r = palGBA(i) And &H1F
    For s = 0 To 2: r = r * 2: Next s 'simulates r << 3
    
    g = palGBA(i)
    For s = 0 To 4: g = g / 2: Next s 'simulates g >> 5
    g = g And &H1F
    For s = 0 To 2: g = g * 2: Next s 'simulates g << 3
    
    b = palGBA(i)
    
    For s = 0 To 9: b = b / 2: Next s 'simulates b >> 10
    b = b And &H1F
    For s = 0 To 2: b = b * 2: Next s 'simulates b << 3
    palPC(i) = RGB(r, g, b)
    
    If palGBA(i) = &H7FFF Then palPC(i) = &HFFFFFF
    'If i Mod 16 = 14 Then palPC(i) = RGB(249, 249, 249) 'UGLY hack to fix Red/White mixups.
  Next i
End Sub

'...I should write a PACKpalette function ^o^
