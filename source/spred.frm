VERSION 5.00
Begin VB.Form frmSprEd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spread"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "spred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picExport 
      AutoRedraw      =   -1  'True
      Height          =   1020
      Left            =   120
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6976
      TabIndex        =   26
      Top             =   3120
      Visible         =   0   'False
      Width           =   1.04700e5
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2760
      TabIndex        =   25
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      ItemData        =   "spred.frx":1C7A
      Left            =   1320
      List            =   "spred.frx":1C87
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdPal 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1950
      TabIndex        =   22
      Top             =   480
      Width           =   285
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2760
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   16
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtGraphic 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9.75
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox txtPal 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtOffBot 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtOffTop 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   615
   End
   Begin VB.TextBox txtSlot 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtIndex 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      DrawStyle       =   5  'Transparent
      FillStyle       =   0  'Solid
      Height          =   1020
      Left            =   90
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   90
      Width           =   540
   End
   Begin VB.TextBox CS 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3120
      TabIndex        =   13
      Text            =   "0"
      Top             =   1080
      Width           =   495
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H0000000F&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   1320
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   20
      Top             =   780
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Size"
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label lblVer 
      Caption         =   "<ver>"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label lblTrueGFX 
      Height          =   255
      Left            =   3000
      TabIndex        =   17
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   40
      X2              =   264
      Y1              =   61
      Y2              =   61
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   40
      X2              =   264
      Y1              =   60
      Y2              =   60
   End
   Begin VB.Label lblGraphic 
      Caption         =   "Graphic"
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblPal 
      Caption         =   "Palette"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblOffset 
      Caption         =   "Offset"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label lblSlot 
      Caption         =   "Slot"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblIndex 
      Caption         =   "Index"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "frmSprEd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This forces us to explicitly declare ("Dim") our variables. This prevents stupid typos from breaking the program.
Option Explicit

'Some research data copied from the Doggie Doo's SPRITES.TXT file

'        32x16     16x16     32x32
'------------------------------------
'Index?  1104FFFF  1106FFFF  1104FFFF
'Slot?   010011FF  008011FF  020011FF
'Offset  00200010  00100010  00200020
'Pal     0000....  0000....  0000....
'???     083711EC  083711D4  083711F4
'Size    083712BC  08371244  08371334
'Anim    08370F60  08370F60  08370F60
'Sprite  08......  08......  08......
'???     081E2910  081E2910  081E2910

'A reproduction of the sprite entry format. Some fields aren't
'understood at all and should not be edited.
Private Type tSprite
  lIndex As Long 'Primary Key in sprite database?
  lSlot As Long
  iOffTop As Integer 'Hotspot's top area. Normally 16.
  iOffBot As Integer 'Hotspot's bottom area. Normally 32.
  iPal As Integer 'Palette to use. Probably contains more than that, only last digit actually matters.
  filler As Integer
  u1 As Long 'An unknown pointer.
  ptrSize As Long 'Pointer to unknown data. Determines sprite size.
  ptrAnim As Long 'Pointer to unknown data. Determines sprite mobility: just one sprite, can only turn (gym leaders) or fully mobile.
  ptrGraphic As Long 'Pointer to pointer to graphics <- not a typo ;)
  u2 As Long 'Another unknown pointer.
End Type

Dim Sprite As tSprite

Dim RomData As Integer 'PokeRoms index

Dim myPal(255) As Long 'Palette data
Dim TileData(31) As Byte '4BPP tile data
Dim gfxptr As Long 'Pointer to the graphics
Dim TheFile As String 'The file name we're using

Private Sub DrawSprite()
  Dim tempal(15) As Long
  Dim pal2use As Integer
  Dim i As Integer
  
  On Error GoTo Hell
  
  'Take the last digit from the Palette value and get absolute index in 16x16 palette grid.
  pal2use = (Val("&H" & Right(txtPal, 1)) - 2) * 16
  If pal2use < 0 Then pal2use = 0
  'Copy the required colors into temporary palette for drawing
  For i = 0 To 15
    tempal(i) = myPal(pal2use + i)
  Next i
  'cboPal is no longer used. I wonder what took me so long to remove it.
  'Probably the fact it was hidden deep below the form's total height.
  ''cboPal.ListIndex = Val("&H" & Right(txtPal, 1))
  'Move the dropdown's hilite accordingly
  Label1.Top = (Val("&H" & Right(txtPal, 1)) - 2) * 10
  
  'Get the pointer to the graphics
  Seek #1, Sprite.ptrGraphic + 1 - &H8000000
  Get #1, , gfxptr
  'Reflect this to the user
  lblTrueGFX = "&&H" & Hex(gfxptr - &H8000000)
  'Prepare to load!
  Seek #1, gfxptr + 1 - &H8000000
  
  'Set the sample window's filling color with Pokémon-style Transparent Green
  Picture1.FillColor = RGB(112, 192, 168)
    
  'Developers can use this to find the SpriteXXXSet values for PokéRoms.
  'lblVer = Hex(Sprite.ptrSize - &H8000000)
  
  'This is VERY sloppy coding here. Should be optimized, can't be bothered.
  'NORMAL SPRITES
  If Sprite.ptrSize - &H8000000 = Roms(RomData).SpriteNormalSet Then
    'Read 32 bytes of 4BPP graphics data. That's worth 64 pixels -> 8 * 8 = ...
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    'Call BlitTile to draw the top left block of the sprite
    BlitTile TileData(), Picture1.hdc, 0, 0, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 8, 0, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 0, 8, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 8, 8, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 0, 16, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 8, 16, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 0, 24, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 8, 24, tempal()
    
    'Scale this to 200% ;)
    StretchBlt Picture1.hdc, 0, 0, 64, 64, Picture1.hdc, 0, 0, 32, 32, vbSrcCopy
  
  'SMALL SPRITES
  ElseIf Sprite.ptrSize - &H8000000 = Roms(RomData).SpriteSmallSet Then
    'Note how this one has only HALF the amount of tiles as Normal sprites.
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 0, 0, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 8, 0, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 0, 8, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 8, 8, tempal()

    StretchBlt Picture1.hdc, 0, 32, 64, 32, Picture1.hdc, 0, 0, 32, 16, vbSrcCopy
    Rectangle Picture1.hdc, 0, 0, 64, 33
  
  'LARGE SPRITES
  ElseIf Sprite.ptrSize - &H8000000 = Roms(RomData).SpriteLargeSet Then
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 0, 0, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 8, 0, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 16, 0, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 24, 0, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 0, 8, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 8, 8, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 16, 8, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 24, 8, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 0, 16, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 8, 16, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 16, 16, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 24, 16, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 0, 24, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 8, 24, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 16, 24, tempal()
    For i = 0 To 31
      Get #1, , TileData(i)
    Next i
    BlitTile TileData(), Picture1.hdc, 24, 24, tempal()
    
    'Copy the image to the BOTTOM half of the picture box
    'Why am I stretching anyway? A normal BitBlt would suffice!
    StretchBlt Picture1.hdc, 0, 32, 32, 32, Picture1.hdc, 0, 0, 32, 32, vbSrcCopy
    'Answer: Because today's computers are fast enough not to notice ;)
    'Fill in the top half
    Rectangle Picture1.hdc, 0, 0, 64, 33
    
  Else
    'We can't draw anything with the (lack of) data we have.
    'Rectangle() is an API function similar to VB's Line(),BF only MUCH faster.
    'It's color must be set beforehand. We do this the VB way: .Fillcolor property.
    Rectangle Picture1.hdc, 0, 0, 64, 64
  End If
  
Hell:
  'BLOOP!
  Picture1.Refresh
End Sub

'Private Sub cboPal_Click()
'  txtPal = Left(txtPal, Len(txtPal) - 1) & Hex(cboPal.ListIndex)
'  DrawSprite
'End Sub

Private Sub cboSize_Click()
  'This basically reflects the right Size pointer into the sprite's database entry,
  'and sets the hotspot accordingly. It's easier than typing the clumsy offsets by hand.
  If cboSize.ListIndex = 0 Then
    Sprite.ptrSize = Roms(RomData).SpriteNormalSet + &H8000000
    txtOffTop = 16
    txtOffBot = 32
  ElseIf cboSize.ListIndex = 1 Then
    Sprite.ptrSize = Roms(RomData).SpriteSmallSet + &H8000000
    txtOffTop = 16
    txtOffBot = 16
  ElseIf cboSize.ListIndex = 2 Then
    Sprite.ptrSize = Roms(RomData).SpriteLargeSet + &H8000000
    txtOffTop = 32
    txtOffBot = 32
  End If
  'Ofcourse, instantly reflect the change.
  DrawSprite
End Sub

Private Sub cmdEdit_Click()
  'This one's new. It launches the built-in sprite gfx editor.
  'First, we copy some DrawSprite() code to upload the palette info to the gfx controls.
  Dim tempal(15) As Long
  Dim pal2use As Integer
  Dim i As Integer
  
  'Remember me, rat fans?
  pal2use = (Val("&H" & Right(txtPal, 1)) - 2) * 16
  If pal2use < 0 Then pal2use = 0
  For i = 0 To 15
    tempal(i) = myPal(pal2use + i)
  Next i
  
  Load frmSprEditor
  With frmSprEditor 'I like shorthands. They're comfy and easy to use.
    'frmSpriteEditor is intelligent like that. Given ONE starting offset
    'it'll set ALL 16 tile editor controls accordingly!
    .StartAddress = Val(Mid(lblTrueGFX, 2))
    'First we seemingly remove all editors (mo' like move offscreen)
    'and set their palettes and file names. We ALSO color the option buttons ;)
    For i = 0 To 15
      .ted(i).Left = -5000
      .ted(i).Top = -5000
      .ted(i).SetPalette tempal()
      .ted(i).Filename = TheFile
      .optColor(i).BackColor = tempal(i)
    Next i
    'What've we got Jimmy?
    Select Case cboSize.ListIndex
      Case 0 'Normal
        'Move EIGHT editors around to form a two-by-four structure
        .ted(0).Move 32, 24
        .ted(1).Move 96, 24
        .ted(2).Move 32, 88
        .ted(3).Move 96, 88
        .ted(4).Move 32, 152
        .ted(5).Move 96, 152
        .ted(6).Move 32, 216
        .ted(7).Move 96, 216
        .NumTiles = 7 'because I'm a zero-based bastard.
        .Command2.Value = True 'Simulate a commandbutton click
      Case 1 'Small
        .ted(0).Move 32, 24
        .ted(1).Move 96, 24
        .ted(2).Move 32, 88
        .ted(3).Move 96, 88
        .NumTiles = 3
        .Command2.Value = True
      Case 2 'Large
        .ted(0).Move 32, 24
        .ted(1).Move 96, 24
        .ted(2).Move 160, 24
        .ted(3).Move 224, 24
        .ted(4).Move 32, 88
        .ted(5).Move 96, 88
        .ted(6).Move 160, 88
        .ted(7).Move 224, 88
        .ted(8).Move 32, 152
        .ted(9).Move 96, 152
        .ted(10).Move 160, 152
        .ted(11).Move 224, 152
        .ted(12).Move 32, 216
        .ted(13).Move 96, 216
        .ted(14).Move 160, 216
        .ted(15).Move 224, 216
        .NumTiles = 15
        .Command2.Value = True
    End Select
    'Finally, SHOW the fucker!
    .Show 1
  End With
End Sub

Private Sub cmdLoad_Click()
  On Error Resume Next
  'CS = current selection, used to be a variable, is now a textbox. Same diff to the code.
  Seek #1, (Roms(RomData).SpriteBase + 1) + (36 * CS)
  'SpriteBase + 1 due to VB's half-assed one-based file handling, plus 36 bytes per sprite
  'TIMES the current selection.
  
  'Get the whole damn record structure in one go.
  Get #1, , Sprite
    
  'Neatly format some text fields
  txtIndex = "&H" & PadHex(Sprite.lIndex, 8)
  txtSlot = "&H" & PadHex(Sprite.lSlot, 8)
  txtOffTop = Sprite.iOffTop
  txtOffBot = Sprite.iOffBot
  txtPal = "&H" & Hex(Sprite.iPal)  'Sprite.lPal  'Yes it used to be parsed as a Long.
  txtGraphic = "&H" & PadHex(Sprite.ptrGraphic - &H8000000, 8)
  
  'lblVer = Hex(Sprite.ptrSize)
  
  cboSize.Enabled = True
  'Find out which size we got and set the combobox accordingly.
  If Sprite.ptrSize - &H8000000 = Roms(RomData).SpriteNormalSet Then
    cboSize.ListIndex = 0
  ElseIf Sprite.ptrSize - &H8000000 = Roms(RomData).SpriteSmallSet Then
    cboSize.ListIndex = 1
  ElseIf Sprite.ptrSize - &H8000000 = Roms(RomData).SpriteLargeSet Then
    cboSize.ListIndex = 2
  Else
    'Because if it's not a proper value, it shall not be broken even further.
    cboSize.Enabled = False
  End If
  
  'Draw the fucker
  DrawSprite
End Sub

Private Sub cmdNext_Click()
  'Woah.
  CS = CS + 1
  CS.SetFocus
  cmdLoad_Click
End Sub

Private Sub cmdOK_Click()
  'I know Kung-Fu.
  Sprite.ptrGraphic = Val(txtGraphic) + &H8000000
  txtGraphic.SetFocus
  DrawSprite
End Sub

Private Sub cmdPal_Click()
  'This is a crazy hack from the old days.
  If cmdPal.Caption = "6" Then
    'Make the Marlett glyph point UP
    cmdPal.Caption = "5"
    'Pop up the palette box
    Picture2.Visible = True
    'Give it the focus
    Picture1.SetFocus
  Else
    'Point DOWN, hide the box and give focus to the TEXTBOX instead!
    cmdPal.Caption = "6"
    Picture2.Visible = False
    txtPal.SetFocus
  End If
End Sub

Private Sub cmdPrev_Click()
  'I'm a very sensitive woman. I won't tolerate your repeating.
  If CS > 0 Then CS = CS - 1
  CS.SetFocus
  cmdLoad_Click
End Sub

Private Sub cmdSave_Click()
  'OOOOOOOPS!
  'Seek #1, (&H3718D4 + 1) + (36 * CS)
  Seek #1, (Roms(RomData).SpriteBase + 1) + (36 * CS)
  Put #1, , Sprite
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer
  Dim j As Integer
  j = CS
  If Shift = 2 Then
    If MsgBox("Do you want to mass-export to a picture strip?", vbOKCancel) = vbOK Then
      MousePointer = 11
      For i = 0 To 255
        CS = i
        cmdLoad_Click
        StretchBlt picExport.hdc, i * 32, 0, 32, 64, Picture1.hdc, 0, 0, 32, 64, vbSrcCopy
      Next i
      picExport.Picture = picExport.Image
      SavePicture picExport.Picture, Left(TheFile, Len(TheFile) - 4) & " sprites.bmp"
      MousePointer = 0
      CS = j
      cmdLoad_Click
    End If
  End If
End Sub

Private Sub Form_Load()
  'Yes, we load the whole palette in this sub, for the rest of the program's runcycle.
  Dim tempal(255) As Integer
  'General porpoise counters. Woah.
  Dim i As Integer, j As Long
  'Remember me, rat fans? You WANKERS!
  Dim headr As String * 4
  
  'Anti Name Change Protection. Patent pending.
  lblVer = "Spread " & App.Major & "." & App.Minor & " by Kyoufu Kawa"
  For i = 15 To 25
    j = j + Asc(Mid(lblVer, i, 1))
  Next i
  'As always, uncomment this to get the RIGHT checksum.
  'MsgBox Int(j / 2)
  If Int(j / 2) <> 531 Then
    MsgBox "This program has been hacked and will not run.", vbCritical, "Checksum error"
    End
  End If

  InitDatabase
  
  'If we DID specify a file at startup, use that...
  If Command <> "" Then
    TheFile = Command
  Else
    '...if not, get one from my patented Rom Selector.
    frmRomSelect.Show 1
    If frmRomSelect.Tag = "Cancelled" Then End 'GAME OVER MAN! GAME OVER!
    TheFile = frmRomSelect.Tag
    Unload frmRomSelect
  End If
  'If we somehow STILL lack a file name, give up and die
  If TheFile = "" Then End
  'Reflect the file name in the title bar
  Caption = Caption & " - " & TheFile
  
  'Oooh, the first version's Standard EGA Palette! ^_^
  ''For i = 0 To 15
  ''  myPal(i) = QBColor(i)
  ''Next i
  
  Open TheFile For Binary As #1
  Get #1, &HAD, headr
  RomData = FindRom(headr)
  If RomData = -1 Then
    MsgBox "Unsupported ROM."
    End
  End If
  If Roms(RomData).SpriteBase = 0 Then
    MsgBox "Supported ROM, but Sprite Base is unspecified."
    End
  End If
  
  'Load all palette data
  Seek #1, Roms(RomData).SpriteColors + 1 '3292072 + 1
  For i = 0 To 255
    Get #1, , tempal(i)
  Next i
  'Convert it from SNES to WIN format, from the TEMPORARY palette array to the FINAL array
  UnPackPalette tempal(), myPal, 255
  
  'Show the first entry
  cmdLoad_Click
  
  'You'll notice in the Form Editor that the Palette box is BEHIND the other controls ;)
  'This puts it in front.
  Picture2.ZOrder
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Properly close the edited file
  Close #1
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Set a flag: Yes, the button's down.
  Picture2.Tag = "^_^"
  'Call MouseMove. If you don't, things don't happen until you actually move.
  Picture2_MouseMove Button, Shift, x, y
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim ry As Integer 'Real Y, pixel to row
  Dim ry2 As Integer 'Real Y + 2, don't ask, don't know
  Dim ny As Integer 'New Y
  If Picture2.Tag <> "^_^" Then Exit Sub 'Button not down? Bail out!
  ry = Int(y / 10) 'Int rounds it down to an integral value, chops off the decimals...
  ry2 = ry + 2
  ny = ry * 10 '...so we end up with a grid-to-screen reconversion
  If ry2 < 0 Then Exit Sub
  If ry2 > 15 Then Exit Sub 'Only 16 values kthxdie
  Label1.Top = ny 'Move the hilite
  txtPal.Text = Left(txtPal.Text, Len(txtPal.Text) - 1) & Hex(ry2)
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  'Remove the tag
  Picture2.Tag = "-_-"
  'Hide the box
  Picture2.Visible = False
  'Point up
  cmdPal.Caption = "6"
  'Give focus to the textbox
  txtPal.SetFocus
End Sub

Private Sub Picture2_Paint()
  'This basically draws a nice grid of colors. Don't bother.
  'The only interesting thing here is how to convert a ONE dimensional array to a
  'TWO dimensional grid: myPal((j * 16) + i
  Dim i As Integer, j As Integer
  For i = 0 To 15
    For j = 0 To 6
      Picture2.Line ((i * 10) + 1, (j * 10) + 1)-((i * 10) + 9, (j * 10) + 9), myPal((j * 16) + i), BF
    Next j
  Next i
End Sub

Private Sub txtPal_Change()
  'Set the sprite's property
  Sprite.iPal = Val(txtPal)
  'Reflect
  DrawSprite
End Sub
