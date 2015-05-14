VERSION 5.00
Begin VB.Form frmSprEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editing sprite"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hsbFrame 
      Height          =   255
      Left            =   2400
      Max             =   8
      TabIndex        =   35
      Top             =   45
      Width           =   1815
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   15
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   14
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2520
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   13
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   12
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   11
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   10
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   9
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   720
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   8
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   360
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   7
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2880
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   6
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2520
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   5
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2160
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   4
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1800
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   3
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1440
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   2
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1080
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   720
      Width           =   255
   End
   Begin VB.OptionButton optColor 
      Height          =   375
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Load"
      Height          =   375
      Left            =   840
      TabIndex        =   17
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   855
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   5
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   6
      Left            =   480
      TabIndex        =   6
      Top             =   3240
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   7
      Left            =   1440
      TabIndex        =   7
      Top             =   3240
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   8
      Left            =   2400
      TabIndex        =   8
      Top             =   360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   9
      Left            =   3360
      TabIndex        =   9
      Top             =   360
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   10
      Left            =   2400
      TabIndex        =   10
      Top             =   1320
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   11
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   12
      Left            =   2400
      TabIndex        =   12
      Top             =   2280
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   13
      Left            =   3360
      TabIndex        =   13
      Top             =   2280
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   14
      Left            =   2400
      TabIndex        =   14
      Top             =   3240
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin Spread.GBATileEditor ted 
      Height          =   960
      Index           =   15
      Left            =   3360
      TabIndex        =   15
      Top             =   3240
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1693
      DotSize         =   8
   End
   Begin VB.Label Label1 
      Caption         =   "&Frame"
      Height          =   255
      Left            =   1800
      TabIndex        =   34
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "frmSprEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'The tile editors are very specific tools. What YOU want is the
'main form's VIEWING code. I'll comment this form all the same.

Option Explicit

Public NumFrames As Byte
Public NumTiles As Byte
Public StartAddress As Long
Public i As Byte

Private Sub Command1_Click()
  For i = 0 To NumTiles
    ted(i).SaveTileData
  Next i
End Sub

Private Sub Command2_Click()
  'Set the first editor's offset
  ted(0).RomAddress = StartAddress + ((hsbFrame * ((NumTiles + 1) * 32)) * 1)
    
  'Set all the others accordingly, one 32 bytes higher than the last
  For i = 1 To NumTiles
    ted(i).RomAddress = ted(i - 1).RomAddress + 32
  Next i
  
  'Load all the data
  For i = 0 To NumTiles
    ted(i).LoadTileData
  Next i
End Sub

Private Sub hsbFrame_Change()
  'If we change the scrollbar's value, load the right sprite
  Command2_Click
End Sub

Private Sub hsbFrame_Scroll()
  'Mirror of Change(). One drags, the other clicks.
  Command2_Click
End Sub

Private Sub optColor_Click(Index As Integer)
  'Set all editor's pen color at once.
  For i = 0 To 15
    ted(i).PenColor = Index
  Next i
End Sub
