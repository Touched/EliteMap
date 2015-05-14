VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pokémon R/S Base Stat Editor"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "baseedit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   322
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboEggType2 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "baseedit.frx":0442
      Left            =   4680
      List            =   "baseedit.frx":0444
      TabIndex        =   61
      Text            =   "Combo1"
      Top             =   5280
      Width           =   2175
   End
   Begin VB.ComboBox cboEggType1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "baseedit.frx":0446
      Left            =   4680
      List            =   "baseedit.frx":0448
      TabIndex        =   60
      Text            =   "Combo1"
      Top             =   4920
      Width           =   2175
   End
   Begin VB.PictureBox guipal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   59
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtUnknown 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   8
      Left            =   4680
      TabIndex        =   51
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtUnknown 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   4680
      TabIndex        =   54
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtUnknown 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   57
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtUnknown 
      Height          =   285
      Index           =   9
      Left            =   5160
      TabIndex        =   50
      ToolTipText     =   "Egg type - Need value list"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtUnknown 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   5160
      TabIndex        =   53
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtUnknown 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   5160
      TabIndex        =   56
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtUnknown 
      Height          =   285
      Index           =   10
      Left            =   5640
      TabIndex        =   49
      ToolTipText     =   "Egg type - Need value list"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtUnknown 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   5640
      TabIndex        =   52
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtUnknown 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   5640
      TabIndex        =   55
      Top             =   3840
      Width           =   495
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   46
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Caption         =   "HP"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   45
      Top             =   3480
      Width           =   735
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   2400
      TabIndex        =   44
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Caption         =   "ATK"
      Enabled         =   0   'False
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   43
      Top             =   3720
      Width           =   735
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   42
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Caption         =   "DEF"
      Enabled         =   0   'False
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   41
      Top             =   3960
      Width           =   735
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   40
      Top             =   3480
      Width           =   255
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Caption         =   "SPD"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   3720
      TabIndex        =   39
      Top             =   3480
      Width           =   735
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   8
      Left            =   3480
      TabIndex        =   38
      Top             =   3720
      Width           =   255
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Caption         =   "SATK"
      Enabled         =   0   'False
      Height          =   255
      Index           =   9
      Left            =   3720
      TabIndex        =   37
      Top             =   3720
      Width           =   735
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Enabled         =   0   'False
      Height          =   255
      Index           =   10
      Left            =   3480
      TabIndex        =   36
      Top             =   3960
      Width           =   255
   End
   Begin VB.CheckBox chkEffort 
      BackColor       =   &H00F4E0E0&
      Caption         =   "SDEF"
      Enabled         =   0   'False
      Height          =   255
      Index           =   11
      Left            =   3720
      TabIndex        =   35
      Top             =   3960
      Width           =   735
   End
   Begin VB.ComboBox cboAbility1 
      Height          =   315
      ItemData        =   "baseedit.frx":044A
      Left            =   5880
      List            =   "baseedit.frx":0538
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1680
      Width           =   1695
   End
   Begin VB.ComboBox cboAbility2 
      Height          =   315
      ItemData        =   "baseedit.frx":098D
      Left            =   5880
      List            =   "baseedit.frx":0A7B
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   2040
      Width           =   1695
   End
   Begin VB.ComboBox cboType1 
      Height          =   315
      ItemData        =   "baseedit.frx":0ED0
      Left            =   5880
      List            =   "baseedit.frx":0F0A
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   480
      Width           =   1695
   End
   Begin VB.ComboBox cboType2 
      Height          =   315
      ItemData        =   "baseedit.frx":0FC8
      Left            =   5880
      List            =   "baseedit.frx":1002
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   840
      Width           =   1695
   End
   Begin VB.HScrollBar hsbGender 
      Height          =   255
      LargeChange     =   20
      Left            =   2400
      Max             =   255
      TabIndex        =   23
      Top             =   1920
      Value           =   127
      Width           =   2655
   End
   Begin VB.TextBox txtBaseHP 
      Height          =   285
      Left            =   3360
      TabIndex        =   15
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtBaseATK 
      Height          =   285
      Left            =   3360
      TabIndex        =   14
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtBaseDEF 
      Height          =   285
      Left            =   3360
      TabIndex        =   13
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtBaseSPD 
      Height          =   285
      Left            =   5040
      TabIndex        =   12
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtBaseSAT 
      Height          =   285
      Left            =   5040
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox txtBaseSDF 
      Height          =   285
      Left            =   5040
      TabIndex        =   10
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtExpGain 
      Height          =   285
      Left            =   5040
      TabIndex        =   3
      Top             =   2760
      Width           =   615
   End
   Begin VB.TextBox txtRarity 
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   2760
      Width           =   615
   End
   Begin VB.ListBox lstMon 
      Height          =   4575
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdDump 
      Caption         =   "Dump"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Unknown Values"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4680
      TabIndex        =   58
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1215
      Index           =   5
      Left            =   4560
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Effort Values"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2400
      TabIndex        =   48
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00F4E0E0&
      Caption         =   "Under Construction"
      Height          =   255
      Left            =   2400
      TabIndex        =   47
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1575
      Index           =   6
      Left            =   2280
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Abilities"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5880
      TabIndex        =   34
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1215
      Index           =   2
      Left            =   5760
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Types"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   31
      Top             =   240
      Width           =   1695
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1215
      Index           =   1
      Left            =   5760
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Male/Female Ratio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   28
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Male"
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F4E0E0&
      Caption         =   "Female"
      Height          =   255
      Left            =   4200
      TabIndex        =   26
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblGender 
      Alignment       =   2  'Center
      BackColor       =   &H00F4E0E0&
      Caption         =   "50%"
      Height          =   255
      Left            =   3480
      TabIndex        =   25
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label14 
      BackColor       =   &H00F4E0E0&
      Caption         =   "None"
      Height          =   255
      Left            =   5160
      TabIndex        =   24
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label lblHeader 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Base Stats"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   22
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Hit points"
      Height          =   255
      Left            =   2400
      TabIndex        =   21
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Attack"
      Height          =   255
      Left            =   2400
      TabIndex        =   20
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Defence"
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Speed"
      Height          =   255
      Left            =   4080
      TabIndex        =   18
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Sp.Attack"
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00F4E0E0&
      Caption         =   "Sp.Defence"
      Height          =   255
      Left            =   4080
      TabIndex        =   16
      Top             =   1200
      Width           =   855
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1455
      Index           =   0
      Left            =   2280
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00F4E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pokémon Advance Base Stat Editor version 2.8 by Kyoufu Kawa"
      Height          =   615
      Left            =   5880
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblBST 
      Alignment       =   2  'Center
      BackColor       =   &H00F4E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "999"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      BackColor       =   &H00F4E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Base Stat Total"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exp.Gain"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rarity"
      Height          =   255
      Left            =   2400
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1095
      Index           =   8
      Left            =   2280
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   495
      Index           =   7
      Left            =   2280
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   855
      Index           =   3
      Left            =   5760
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1335
      Index           =   4
      Left            =   6240
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Image imgBack 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
   Begin VB.Menu mnuRMB 
      Caption         =   "rmb"
      Visible         =   0   'False
      Begin VB.Menu mnuRMBColors 
         Caption         =   "&Set theme..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Names(512) As String

Dim TheFile As String
Dim RomColor As Integer
Dim phase As String

Dim todoEfforts1 As Byte
Dim todoEfforts2 As Byte

Public mytheme As Integer

Private Sub chkEffort_Click(Index As Integer)
  Dim i As Byte
  For i = 0 To 7
    If chkEffort(i).Value = 1 Then
      todoEfforts1 = CByte(BitSet(CLng(todoEfforts1), i))
    Else
      todoEfforts1 = CByte(BitClear(CLng(todoEfforts1), i))
    End If
  Next i
  For i = 0 To 3
    If chkEffort(i + 8).Value = 1 Then
      todoEfforts2 = CByte(BitSet(CLng(todoEfforts2), i))
    Else
      todoEfforts2 = CByte(BitClear(CLng(todoEfforts2), i))
    End If
  Next i
  chkEffort(1).Caption = todoEfforts1
  chkEffort(3).Caption = todoEfforts2
End Sub

Private Sub cmdDump_Click()
  Dim i As Integer
  Dim s As String
  Caption = "Easter egg activated. Dumping list..."
  Open "All your base.html" For Output As #8
  Print #8, "<head><title>Base stat dump</title></head><body>"
  Print #8, "<h2>Base stat dump for <i>" & TheFile & "</i></h2><hr>"
  For i = 0 To 411 Step 10
        Print #8, "[<a href=" & Chr(34) & "#" & Trim(Str(i)) & Chr(34) & ">" & Trim(Str(i)) & "]"
  Next i
  For i = 0 To 411
    lstMon.ListIndex = i
    lstMon_Click
    DoEvents
          Print #8, "<hr><a name=" & Chr(34) & Trim(Str(i)) & Chr(34) & "><b>" & lstMon.List(i) & "</b><br>"
    Print #8, "HP: " & txtBaseHP & "<br>"
    Print #8, "ATK: " & txtBaseATK & "<br>"
    Print #8, "DEF: " & txtBaseDEF & "<br>"
    Print #8, "SAT: " & txtBaseSAT & "<br>"
    Print #8, "SDF: " & txtBaseSDF & "<br>"
    Print #8, "SPD: " & txtBaseSPD & "<br>"
    Print #8, "<i>BST: " & lblBST & "</i><p>"
    If cboType2.ListIndex = cboType1.ListIndex Then
      Print #8, "Type: <i>" & cboType1.List(cboType1.ListIndex) & "</i><br>"
    Else
      Print #8, "Type: <i>" & cboType1.List(cboType1.ListIndex) & "</i> / <i>" & cboType2.List(cboType2.ListIndex) & "</i><br>"
    End If
    If cboAbility1.ListIndex > 0 Then
      If cboAbility2.ListIndex > 0 Then
        Print #8, "Abilities: <i>" & cboAbility1.List(cboAbility1.ListIndex) & ", " & cboAbility2.List(cboAbility2.ListIndex) & "</i><br>"
      Else
        Print #8, "Ability: <i>" & cboAbility1.List(cboAbility1.ListIndex) & "</i><br>"
      End If
    End If
  Next i
  Close #8
  Caption = Tag
End Sub

Private Sub cmdDump_GotFocus()
  Tag = Caption
  Caption = "Hey hey hey, easter egg..."
End Sub

Private Sub cmdDump_LostFocus()
  Caption = Tag
End Sub

Private Sub Form_Load()
  Dim i As Integer
  Dim d As Long
  Dim nt As String * 11
  Dim header As String * 4
  
  'SetIcon Me.hWnd, "APP", True
  
  For i = Len(Label13) - 10 To Len(Label13)
    d = d + Asc(Mid(Label13, i, 1))
  Next i
  If Int(d / 2) <> 531 Then
    MsgBox "This program has been hacked and will not run.", vbCritical, "Checksum error"
    End
  End If

  On Error Resume Next
  mytheme = Val(INIRead("elitemap", "Shared", "Theme"))
  If mytheme = 0 Then mytheme = 10
  imgBack.Picture = LoadResPicture(mytheme, 0)
  'And now, we go through ALL controls available to recolor...
  guipal.Picture = LoadResPicture(mytheme + 2, 0)
  Dim ColorRemap(1 To 2, 1 To 4) As Long
  Dim X As Integer, Y As Integer
  For X = 1 To 2
    For Y = 1 To 4
      ColorRemap(X, Y) = guipal.Point((Y - 1) * 8, (X - 1) * 8)
      guipal.PSet ((Y - 1) * 8, (X - 1) * 8), vbRed
      'Debug.Print X & " x " & Y & " = " & Hex(ColorRemap(X, Y))
    Next Y
  Next X
  Dim Ctl As Control
  For Each Ctl In Me.Controls
    For Y = 1 To 4
      If Ctl.BorderColor = ColorRemap(1, Y) Then Ctl.BorderColor = ColorRemap(2, Y)
      If Ctl.BackColor = ColorRemap(1, Y) Then Ctl.BackColor = ColorRemap(2, Y)
    Next Y
    If mytheme = 40 And (TypeOf Ctl Is Label Or TypeOf Ctl Is CheckBox) Then Ctl.ForeColor = vbWhite
  Next Ctl
  If mytheme = 40 Then
    For Each Ctl In Me.Controls
      If TypeOf Ctl Is Label Then Ctl.ForeColor = vbWhite
    Next
  End If
  On Error GoTo 0

  
  'On Error GoTo DamnItAll
  If Command <> "" Then
    TheFile = Command
  Else
    'TheFile = InputBox("Enter file name", , "Pokémon Ruby.gba")
    frmRomSelect.Show 1
    If frmRomSelect.Tag = "Cancelled" Then End
    TheFile = frmRomSelect.Tag
    Unload frmRomSelect
  End If
  'TheFile = "D:\My Downloads\se098043\Obsidian.gba"
  If TheFile = "" Then End
  
  'Open "D:\My Downloads\se098043\Obsidian.gba" For Binary As #1
  'RomColor = CheckHeader(TheFile)
  'If RomColor <= 0 Then
  '  MsgBox "That's not a Pokémon Ruby or Sapphire ROM.", vbExclamation, "RomColor = -1"
  '  End
  'End If
  phase = "Opening file"
  Open TheFile For Binary As #1
  
  phase = "Color checking"
  InitDatabase
  Get #1, &HAD, header
  RomColor = FindRom(header)
  If RomColor = -1 Then
    MsgBox "Unsupported rom.", vbExclamation
    End
  End If
  If Roms(RomColor).MonsterNames = 0 Then
    MsgBox "Supported rom, but monster names are not specified in database.", vbExclamation
    End
  End If
  If Roms(RomColor).MonsterBaseStats = 0 Then
    MsgBox "Supported rom, but monster base stats are not specified in database.", vbExclamation
    End
  End If
  
  phase = "Reading name pointer"
  Get #1, Roms(RomColor).MonsterNames + 1, d
  If d < 0 Then
    MsgBox "Incorrect offset in database." & vbCrLf & vbCrLf & "Offset points to " & Hex(d), vbExclamation
    End
  End If
  phase = "Seeking name pointer"
  Seek #1, d + 1 - &H8000000
  phase = "Reading names"
  For i = 0 To 411
    Get #1, , nt
    Names(i) = Sapp2Asc(Replace(nt, Chr$(255), Chr$(0), , , vbBinaryCompare))
    'Trace PadHex(i, 3) & " - " & Names(i)
    lstMon.AddItem PadHex(i, 3) & " - " & Names(i)
    DoEvents
  Next i
  
  Exit Sub
  
DamnItAll:
  Select Case MsgBox("Command$ = " & Command$ & vbCrLf & _
          "TheFile = " & TheFile & vbCrLf & _
          "RomColor = " & RomColor & vbCrLf & _
          "i = " & i & vbCrLf & _
          "Phase = " & phase & vbCrLf & _
          Err.Description, vbAbortRetryIgnore, "Damn it all to hell.")
    Case vbAbort: End
    Case vbRetry: Resume
    Case vbIgnore: Resume Next
  End Select
  Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Close #1
  INIWrite "EliteMap", "Shared", "Theme", Str(mytheme)
End Sub

Private Sub Form_Resize()
  imgBack.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub hsbGender_Change()
  'lblGender = Int((0.255 * hsbGender)) & "%" '(255/100)*Value
End Sub

Private Sub imgBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuRMB
End Sub

Private Sub lblHeader_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuRMB
End Sub

Private Sub lstMon_Click()
  Dim OneByte As Byte
  Dim offset As Long
  
  offset = Roms(RomColor).MonsterBaseStats
  
  Seek #1, (offset + (&H1C * (lstMon.ListIndex - 1))) + 1
  
  Get #1, , OneByte: txtBaseHP = OneByte '0
  Get #1, , OneByte: txtBaseATK = OneByte '1
  Get #1, , OneByte: txtBaseDEF = OneByte '2
  Get #1, , OneByte: txtBaseSPD = OneByte '3
  Get #1, , OneByte: txtBaseSAT = OneByte '4
  Get #1, , OneByte: txtBaseSDF = OneByte '5
  Get #1, , OneByte: cboType1.ListIndex = OneByte '6
  Get #1, , OneByte: cboType2.ListIndex = OneByte '7
  Get #1, , OneByte: txtRarity = OneByte '8
  Get #1, , OneByte: txtExpGain = OneByte '9
  
  Get #1, , OneByte: todoEfforts1 = OneByte 'A
  Get #1, , OneByte: todoEfforts2 = OneByte 'B
  
  Get #1, , OneByte: txtUnknown(2) = OneByte 'C
  Get #1, , OneByte: txtUnknown(3) = OneByte 'D
  Get #1, , OneByte: txtUnknown(4) = OneByte 'E
  Get #1, , OneByte: txtUnknown(5) = OneByte 'F
  
  Get #1, , OneByte: hsbGender.Value = OneByte '10

  Get #1, , OneByte: txtUnknown(6) = OneByte '11
  Get #1, , OneByte: txtUnknown(7) = OneByte '12

  Get #1, , OneByte: txtUnknown(8) = OneByte '13 ': txtLevelUp = OneByte
  Get #1, , OneByte: cboEggType1.Text = OneByte: txtUnknown(9) = OneByte
  Get #1, , OneByte: cboEggType2.Text = OneByte: txtUnknown(10) = OneByte
  
  Get #1, , OneByte: cboAbility1.ListIndex = OneByte '16
  Get #1, , OneByte: cboAbility2.ListIndex = OneByte '17
  
  'Get #1, , OneByte
  'Get #1, , OneByte
  'Get #1, , OneByte
  'Get #1, , OneByte
  
  For OneByte = 0 To 11
    chkEffort(OneByte).Value = 0
  Next OneByte
  For OneByte = 0 To 7
    If BitIsSet(CLng(todoEfforts1), OneByte) = True Then
      chkEffort(OneByte).Value = 1
    Else
      chkEffort(OneByte).Value = 0
    End If
  Next OneByte
  For OneByte = 0 To 3
    If BitIsSet(CLng(todoEfforts2), OneByte) = True Then
      chkEffort(OneByte + 8).Value = 1
    Else
      chkEffort(OneByte + 8).Value = 0
    End If
  Next OneByte
  chkEffort(1).Caption = todoEfforts1
  chkEffort(3).Caption = todoEfforts2
End Sub

Private Sub cmdSave_Click()
  Dim OneByte As Byte
  Dim offset As Long
  
  offset = Roms(RomColor).MonsterBaseStats
  
  Seek #1, (offset + (&H1C * (lstMon.ListIndex - 1))) + 1
  
  OneByte = txtBaseHP: Put #1, , OneByte
  OneByte = txtBaseATK: Put #1, , OneByte
  OneByte = txtBaseDEF: Put #1, , OneByte
  OneByte = txtBaseSPD: Put #1, , OneByte
  OneByte = txtBaseSAT: Put #1, , OneByte
  OneByte = txtBaseSDF: Put #1, , OneByte
  OneByte = cboType1.ListIndex: Put #1, , OneByte
  OneByte = cboType2.ListIndex: Put #1, , OneByte
  OneByte = txtRarity: Put #1, , OneByte
  OneByte = txtExpGain: Put #1, , OneByte
  
  OneByte = todoEfforts1: Put #1, , OneByte
  OneByte = todoEfforts2: Put #1, , OneByte
  OneByte = txtUnknown(2): Put #1, , OneByte
  OneByte = txtUnknown(3): Put #1, , OneByte
  OneByte = txtUnknown(4): Put #1, , OneByte
  OneByte = txtUnknown(5): Put #1, , OneByte

  OneByte = hsbGender.Value: Put #1, , OneByte
  
  OneByte = txtUnknown(6): Put #1, , OneByte
  OneByte = txtUnknown(7): Put #1, , OneByte

  OneByte = txtUnknown(8): Put #1, , OneByte
  OneByte = txtUnknown(9): Put #1, , OneByte
  OneByte = txtUnknown(10): Put #1, , OneByte
  
  OneByte = cboAbility1.ListIndex: Put #1, , OneByte
  OneByte = cboAbility2.ListIndex: Put #1, , OneByte

End Sub

Private Sub mnuRMBColors_Click()
  frmThemes.Show 1
End Sub

Private Sub txtBaseATK_Change()
  If Val(txtBaseATK) > 255 Then txtBaseATK = 255
  lblBST = Val(txtBaseHP) + Val(txtBaseATK) + Val(txtBaseDEF) + Val(txtBaseSPD) + Val(txtBaseSAT) + Val(txtBaseSDF)
End Sub

Private Sub txtBaseDEF_Change()
  If Val(txtBaseDEF) > 255 Then txtBaseDEF = 255
  lblBST = Val(txtBaseHP) + Val(txtBaseATK) + Val(txtBaseDEF) + Val(txtBaseSPD) + Val(txtBaseSAT) + Val(txtBaseSDF)
End Sub

Private Sub txtBaseHP_Change()
  If Val(txtBaseHP) > 255 Then txtBaseHP = 255
  lblBST = Val(txtBaseHP) + Val(txtBaseATK) + Val(txtBaseDEF) + Val(txtBaseSPD) + Val(txtBaseSAT) + Val(txtBaseSDF)
End Sub

Private Sub txtBaseSAT_Change()
  If Val(txtBaseSAT) > 255 Then txtBaseSAT = 255
  lblBST = Val(txtBaseHP) + Val(txtBaseATK) + Val(txtBaseDEF) + Val(txtBaseSPD) + Val(txtBaseSAT) + Val(txtBaseSDF)
End Sub

Private Sub txtBaseSDF_Change()
  If Val(txtBaseSDF) > 255 Then txtBaseSDF = 255
  lblBST = Val(txtBaseHP) + Val(txtBaseATK) + Val(txtBaseDEF) + Val(txtBaseSPD) + Val(txtBaseSAT) + Val(txtBaseSDF)
End Sub

Private Sub txtBaseSPD_Change()
  If Val(txtBaseSPD) > 255 Then txtBaseSPD = 255
  lblBST = Val(txtBaseHP) + Val(txtBaseATK) + Val(txtBaseDEF) + Val(txtBaseSPD) + Val(txtBaseSAT) + Val(txtBaseSDF)
End Sub

Private Sub txtExpGain_Change()
  If Val(txtExpGain) > 255 Then txtExpGain = 255
  lblBST = Val(txtBaseHP) + Val(txtBaseATK) + Val(txtBaseDEF) + Val(txtBaseSPD) + Val(txtBaseSAT) + Val(txtBaseSDF)
End Sub

Private Sub txtRarity_Change()
  If Val(txtRarity) > 255 Then txtRarity = 255
  lblBST = Val(txtBaseHP) + Val(txtBaseATK) + Val(txtBaseDEF) + Val(txtBaseSPD) + Val(txtBaseSAT) + Val(txtBaseSDF)
End Sub

Private Sub txtUnknown_Change(Index As Integer)
  If Val(txtUnknown(Index)) > 255 Then txtUnknown(Index) = 255
End Sub
