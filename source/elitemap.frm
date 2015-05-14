VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00F4E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EliteMap"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "elitemap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   552
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox guipal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9120
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   182
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer timThrobber 
      Interval        =   100
      Left            =   9720
      Top             =   360
   End
   Begin VB.PictureBox picThrobberPics 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   10200
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   205
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox picThrobber 
      AutoRedraw      =   -1  'True
      Height          =   315
      Left            =   10200
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   204
      Tag             =   "hi"
      Top             =   30
      Width           =   1635
   End
   Begin VB.OptionButton edittab 
      Caption         =   "[14] About"
      Height          =   255
      Index           =   4
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   186
      Top             =   390
      Width           =   975
   End
   Begin VB.OptionButton edittab 
      Caption         =   "[13] Patches"
      Height          =   255
      Index           =   3
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   173
      Top             =   390
      Width           =   975
   End
   Begin VB.OptionButton edittab 
      Caption         =   "[12] Objects"
      Height          =   255
      Index           =   2
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   172
      Top             =   390
      Width           =   975
   End
   Begin VB.OptionButton edittab 
      Caption         =   "[11] Header"
      Height          =   255
      Index           =   1
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   171
      Top             =   390
      Width           =   975
   End
   Begin VB.OptionButton edittab 
      Caption         =   "[10] Map"
      Height          =   255
      Index           =   0
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   170
      Top             =   390
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CheckBox chkSSigns 
      BackColor       =   &H00F4E0E0&
      Caption         =   "[9] Signs"
      Height          =   255
      Left            =   1200
      TabIndex        =   169
      Top             =   4920
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkSTraps 
      BackColor       =   &H00F4E0E0&
      Caption         =   "[7] Traps"
      Height          =   255
      Left            =   1200
      TabIndex        =   168
      Top             =   4680
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkSExits 
      BackColor       =   &H00F4E0E0&
      Caption         =   "[8] Exits"
      Height          =   255
      Left            =   240
      TabIndex        =   167
      Top             =   4920
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkSPeople 
      BackColor       =   &H00F4E0E0&
      Caption         =   "[6] People"
      Height          =   255
      Left            =   240
      TabIndex        =   166
      Top             =   4680
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CheckBox chkSprites 
      BackColor       =   &H00F4E0E0&
      Caption         =   "[5] Show Objects"
      Height          =   255
      Left            =   240
      TabIndex        =   165
      Top             =   4320
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.ComboBox cboBanks 
      Height          =   315
      Left            =   3480
      TabIndex        =   164
      Text            =   "0"
      Top             =   15
      Width           =   735
   End
   Begin VB.ComboBox cboLevels 
      Height          =   315
      Left            =   4320
      TabIndex        =   163
      Text            =   "0"
      Top             =   15
      Width           =   735
   End
   Begin VB.TextBox txtOldskoolChooser 
      Height          =   315
      Left            =   5160
      TabIndex        =   162
      Text            =   "0000"
      Top             =   15
      Width           =   615
   End
   Begin VB.TextBox txtRom 
      Height          =   315
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   160
      Top             =   15
      Width           =   2295
   End
   Begin VB.PictureBox picTileBox 
      BorderStyle     =   0  'None
      Height          =   1920
      Left            =   45
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   273
      TabIndex        =   151
      Top             =   615
      Width           =   4095
      Begin VB.PictureBox picTileset 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   15360
         Left            =   0
         MouseIcon       =   "elitemap.frx":030A
         MousePointer    =   99  'Custom
         ScaleHeight     =   1024
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   257
         TabIndex        =   153
         Top             =   0
         Width           =   3855
         Begin VB.Shape shpTileset 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   3
            Height          =   255
            Left            =   0
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.VScrollBar vsbTileset 
         Height          =   1920
         LargeChange     =   4
         Left            =   3840
         Max             =   56
         TabIndex        =   152
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picAttributes 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   120
      MouseIcon       =   "elitemap.frx":045C
      MousePointer    =   99  'Custom
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   150
      Top             =   2880
      Width           =   3840
      Begin VB.Shape shpAttributes 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Height          =   255
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.CommandButton cmdSetLAtts 
      Caption         =   "Set L tiles to L attribs"
      Height          =   255
      Left            =   2280
      TabIndex        =   149
      Top             =   4980
      Width           =   1695
   End
   Begin VB.CommandButton cmdSwapRL 
      Caption         =   "Swap R with L"
      Height          =   255
      Left            =   2280
      TabIndex        =   148
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdReplaceRL 
      Caption         =   "Replace R with L"
      Height          =   255
      Left            =   2280
      TabIndex        =   147
      Top             =   4380
      Width           =   1695
   End
   Begin VB.PictureBox picWorldMap 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   120
      ScaleHeight     =   160
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   146
      Top             =   5760
      Width           =   3840
      Begin VB.Shape shMap 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   135
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Shape shLoc 
         BackColor       =   &H0080FFFF&
         BorderColor     =   &H0000FFFF&
         Height          =   135
         Left            =   0
         Shape           =   3  'Circle
         Top             =   0
         Visible         =   0   'False
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdAdjMap 
      Caption         =   "Adj. 7"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   1320
      TabIndex        =   4
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdjMap 
      Caption         =   "Adj. 6"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   3
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox tattr 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   1
      Text            =   "&Hc"
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox tattr 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Text            =   "&H1"
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox tattr 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   2
      Text            =   "&H4"
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picMainTab 
      BackColor       =   &H00F4E0E0&
      BorderStyle     =   0  'None
      Height          =   7455
      Index           =   4
      Left            =   4200
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   103
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.PictureBox picTeam 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2025
         Left            =   5640
         ScaleHeight     =   135
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   116
         TabIndex        =   124
         Top             =   1200
         Width           =   1740
         Begin VB.Image imgHRguy 
            Height          =   975
            Index           =   3
            Left            =   720
            ToolTipText     =   "Kyoufu Kawa(-oneechan) - Project Management/Lead Coder/Artist/Hummus"
            Top             =   360
            Width           =   375
         End
         Begin VB.Image imgHRguy 
            Height          =   855
            Index           =   0
            Left            =   480
            ToolTipText     =   "Baro - Design and ...well, stuff"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Not pictured due to lack of sprites: Interdepth!"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   240
            TabIndex        =   227
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Image imgHRguy 
            Height          =   255
            Index           =   8
            Left            =   360
            ToolTipText     =   "Markus ""D-Kiddy"" - Assistant Hummus (missing in roleplay)"
            Top             =   0
            Width           =   255
         End
         Begin VB.Image imgHRguy 
            Height          =   855
            Index           =   7
            Left            =   1080
            ToolTipText     =   "Ranko Saotome - General Hummus"
            Top             =   360
            Width           =   375
         End
         Begin VB.Image imgHRguy 
            Height          =   375
            Index           =   5
            Left            =   840
            ToolTipText     =   "Andrew ""DJ ßouché"" Lim  - Coding/Music (missing in action)"
            Top             =   120
            Width           =   375
         End
         Begin VB.Image imgHRguy 
            Height          =   615
            Index           =   6
            Left            =   1200
            ToolTipText     =   "Rick ""The Trasher"" Trimble - General Hummus/Comics"
            Top             =   0
            Width           =   495
         End
         Begin VB.Image imgHRguy 
            Height          =   375
            Index           =   4
            Left            =   600
            ToolTipText     =   "Tony ""Majin BlueDragon"" - Serious FTP Leech™"
            Top             =   0
            Width           =   375
         End
         Begin VB.Image imgHRguy 
            Height          =   855
            Index           =   2
            Left            =   0
            ToolTipText     =   "Tauwasser - Coding/Math Issues"
            Top             =   240
            Width           =   375
         End
         Begin VB.Image imgHRguy 
            Height          =   975
            Index           =   1
            Left            =   240
            ToolTipText     =   "TJ ""Hiryuu"" Chastain - General Hummus"
            Top             =   360
            Width           =   375
         End
         Begin VB.Shape Shape9 
            BorderColor     =   &H00D69896&
            Height          =   1665
            Left            =   0
            Top             =   0
            Width           =   1740
         End
      End
      Begin VB.TextBox txtCredits 
         Appearance      =   0  'Flat
         BackColor       =   &H00E6BBBA&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   6255
         Left            =   120
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   105
         Text            =   "elitemap.frx":05AE
         Top             =   1080
         Width           =   7455
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00D69896&
         Height          =   6285
         Left            =   105
         Top             =   1065
         Width           =   7485
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "<version number here>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   104
         Top             =   600
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   840
         Left            =   120
         Picture         =   "elitemap.frx":086E
         Top             =   120
         Width           =   2610
      End
   End
   Begin VB.PictureBox picMainTab 
      BackColor       =   &H00F4E0E0&
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   1
      Left            =   4200
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CheckBox chkAllowFlash 
         BackColor       =   &H00F4E0E0&
         Caption         =   "[400] Allow FLASH usage"
         Height          =   255
         Left            =   120
         TabIndex        =   218
         Top             =   2520
         Width           =   3615
      End
      Begin VB.ComboBox cboHackLanguage 
         Height          =   315
         ItemData        =   "elitemap.frx":10C7
         Left            =   5160
         List            =   "elitemap.frx":10C9
         Style           =   2  'Dropdown List
         TabIndex        =   217
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Timer timWorkTimer 
         Enabled         =   0   'False
         Interval        =   60000
         Left            =   7080
         Top             =   1680
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   375
         Left            =   6600
         TabIndex        =   214
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtGroupName 
         Height          =   285
         Left            =   5160
         MaxLength       =   16
         TabIndex        =   211
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txtAuthorName 
         Height          =   285
         Left            =   5160
         MaxLength       =   16
         TabIndex        =   209
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtHackName 
         Height          =   285
         Left            =   5160
         MaxLength       =   16
         TabIndex        =   207
         Top             =   120
         Width           =   2415
      End
      Begin VB.ComboBox cboSong 
         Height          =   315
         ItemData        =   "elitemap.frx":10CB
         Left            =   1320
         List            =   "elitemap.frx":10CD
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox chkShowLabel 
         BackColor       =   &H00F4E0E0&
         Caption         =   "[31] Show Label on Entry"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2280
         Width           =   3615
      End
      Begin VB.ComboBox cboLabelID 
         Height          =   315
         ItemData        =   "elitemap.frx":10CF
         Left            =   1320
         List            =   "elitemap.frx":10D1
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   120
         Width           =   2415
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "elitemap.frx":10D3
         Left            =   1320
         List            =   "elitemap.frx":10D5
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cboWeather 
         Height          =   315
         ItemData        =   "elitemap.frx":10D7
         Left            =   1320
         List            =   "elitemap.frx":10D9
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtLevelHeight 
         Height          =   285
         Left            =   1320
         TabIndex        =   25
         ToolTipText     =   "Level height"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtLevelWidth 
         Height          =   285
         Left            =   1320
         TabIndex        =   24
         ToolTipText     =   "Level width"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "[404] Language"
         Height          =   255
         Left            =   3960
         TabIndex        =   216
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblGUID 
         BackStyle       =   0  'Transparent
         Caption         =   "GUID: {00000000-0000-0000-0000-0000-00000000}"
         Height          =   495
         Left            =   3960
         TabIndex        =   215
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label lblWorkTime 
         BackStyle       =   0  'Transparent
         Caption         =   "[406] Not keeping track"
         Height          =   495
         Left            =   5160
         TabIndex        =   213
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label60 
         BackStyle       =   0  'Transparent
         Caption         =   "[405] Working time"
         Height          =   255
         Left            =   3960
         TabIndex        =   212
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "[403] Group name"
         Height          =   255
         Left            =   3960
         TabIndex        =   210
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label58 
         BackStyle       =   0  'Transparent
         Caption         =   "[402] Author"
         Height          =   255
         Left            =   3960
         TabIndex        =   208
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "[401] Hack name"
         Height          =   255
         Left            =   3960
         TabIndex        =   206
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00D69896&
         X1              =   256
         X2              =   256
         Y1              =   192
         Y2              =   -8
      End
      Begin VB.Label lblSongWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "This map has a special song value."
         Height          =   255
         Left            =   1320
         TabIndex        =   138
         Top             =   1200
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "[28] Song"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "[25] Label"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "[27] Type"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "[26] Weather"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label37 
         BackStyle       =   0  'Transparent
         Caption         =   "[30] Height"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label36 
         BackStyle       =   0  'Transparent
         Caption         =   "[29] Width"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.PictureBox picMainTab 
      BackColor       =   &H00F4E0E0&
      BorderStyle     =   0  'None
      Height          =   7455
      Index           =   0
      Left            =   4200
      ScaleHeight     =   497
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CheckBox chkAttribsOnly 
         BackColor       =   &H00F4E0E0&
         Caption         =   "[409] Only set attributes"
         Height          =   375
         Left            =   1920
         TabIndex        =   222
         Top             =   6600
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "5"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3450
         TabIndex        =   131
         Top             =   330
         Width           =   1545
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "5"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   1905
         TabIndex        =   130
         Top             =   330
         Width           =   1545
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00F4E0E0&
         Caption         =   "[20] Stamp"
         Height          =   855
         Left            =   1080
         TabIndex        =   144
         Top             =   6240
         Width           =   735
         Begin VB.PictureBox Picture3 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00D69896&
            Height          =   525
            Left            =   120
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   145
            Top             =   240
            Width           =   525
         End
      End
      Begin VB.CommandButton cmdShiftDown 
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
         Height          =   255
         Left            =   6600
         TabIndex        =   143
         Top             =   5520
         Width           =   255
      End
      Begin VB.CommandButton cmdShiftRight 
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
         Left            =   6840
         TabIndex        =   142
         Top             =   5280
         Width           =   255
      End
      Begin VB.CommandButton cmdShiftLeft 
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
         Left            =   6360
         TabIndex        =   141
         Top             =   5280
         Width           =   255
      End
      Begin VB.CommandButton cmdShiftUp 
         Caption         =   "5"
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
         Left            =   6600
         TabIndex        =   140
         Top             =   5040
         Width           =   255
      End
      Begin VB.CommandButton cmdFullBRD 
         Caption         =   "View"
         Enabled         =   0   'False
         Height          =   465
         Left            =   390
         TabIndex        =   139
         Top             =   6510
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "6"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   5880
         Width           =   1725
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "4"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Index           =   12
         Left            =   5550
         TabIndex        =   134
         Top             =   2265
         Width           =   255
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "6"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   3570
         TabIndex        =   137
         Top             =   5880
         Width           =   1725
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "6"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   1845
         TabIndex        =   136
         Top             =   5880
         Width           =   1725
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "4"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Index           =   13
         Left            =   5550
         TabIndex        =   135
         Top             =   3945
         Width           =   255
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "3"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Index           =   11
         Left            =   120
         TabIndex        =   133
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "3"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Index           =   10
         Left            =   120
         TabIndex        =   132
         Top             =   2280
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E6BBBA&
         Height          =   1110
         Left            =   3840
         ScaleHeight     =   70
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   125
         Top             =   6240
         Width           =   1935
         Begin VB.Label Label2 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000017&
            Height          =   255
            Left            =   30
            TabIndex        =   129
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000017&
            Height          =   255
            Left            =   30
            TabIndex        =   128
            Top             =   30
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000017&
            Height          =   255
            Left            =   30
            TabIndex        =   127
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000018&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000017&
            Height          =   255
            Left            =   30
            TabIndex        =   126
            Top             =   525
            Width           =   1815
         End
      End
      Begin VB.VScrollBar vsbScroll 
         Height          =   5040
         Left            =   5295
         Max             =   0
         TabIndex        =   8
         Top             =   585
         Width           =   255
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "4"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Index           =   3
         Left            =   5550
         TabIndex        =   9
         Top             =   585
         Width           =   255
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "Surface"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   4980
         TabIndex        =   10
         Top             =   330
         Width           =   825
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00F4E0E0&
         Caption         =   "[19] Border"
         Height          =   855
         Left            =   240
         TabIndex        =   116
         Top             =   6240
         Width           =   735
         Begin VB.PictureBox picBorder 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00D69896&
            Height          =   525
            Left            =   120
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   117
            Top             =   240
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F4E0E0&
         Caption         =   "[17] Object Info"
         Height          =   4335
         Left            =   6000
         TabIndex        =   113
         Top             =   360
         Width           =   1575
         Begin VB.Label lblInfo 
            BackColor       =   &H00F4E0E0&
            Height          =   3975
            Left            =   120
            TabIndex        =   114
            Top             =   240
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00F4E0E0&
         Caption         =   "[18] Shift Level"
         Height          =   1095
         Left            =   6000
         TabIndex        =   112
         Top             =   4800
         Width           =   1575
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   7080
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   108
         Top             =   6000
         Width           =   480
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   7080
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   107
         Top             =   6240
         Width           =   480
      End
      Begin VB.PictureBox picSel 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   7080
         ScaleHeight     =   14
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   30
         TabIndex        =   106
         Top             =   6480
         Width           =   480
      End
      Begin VB.CheckBox chkNoDraw 
         BackColor       =   &H00F4E0E0&
         Caption         =   "[21] Disable map editing"
         Height          =   375
         Left            =   1920
         TabIndex        =   23
         Top             =   6240
         Width           =   1815
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "5"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   330
         Width           =   1545
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "3"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   12
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "Dive"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   5295
         TabIndex        =   7
         Top             =   5880
         Width           =   510
      End
      Begin VB.HScrollBar hsbScroll 
         Height          =   255
         Left            =   120
         Max             =   0
         TabIndex        =   6
         Top             =   5625
         Width           =   5685
      End
      Begin VB.PictureBox p 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   5055
         Left            =   375
         ScaleHeight     =   335
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   327
         TabIndex        =   12
         Top             =   585
         Width           =   4935
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   3945
            Left            =   0
            MousePointer    =   99  'Custom
            ScaleHeight     =   263
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   284
            TabIndex        =   13
            Top             =   0
            Width           =   4260
            Begin VB.Image sExit 
               Height          =   240
               Index           =   0
               Left            =   960
               MousePointer    =   99  'Custom
               Picture         =   "elitemap.frx":10DB
               Top             =   480
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image sTrap 
               Height          =   240
               Index           =   0
               Left            =   720
               MousePointer    =   99  'Custom
               Picture         =   "elitemap.frx":117C
               Top             =   480
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image sSign 
               Height          =   240
               Index           =   0
               Left            =   480
               MousePointer    =   99  'Custom
               Picture         =   "elitemap.frx":1220
               Top             =   480
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image sPeople 
               Height          =   240
               Index           =   0
               Left            =   1200
               MousePointer    =   99  'Custom
               Picture         =   "elitemap.frx":12BF
               Top             =   480
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.Image imgGirlStart 
               Enabled         =   0   'False
               Height          =   240
               Left            =   960
               Picture         =   "elitemap.frx":136B
               Top             =   720
               Width           =   360
            End
            Begin VB.Image imgBoyStart 
               Enabled         =   0   'False
               Height          =   240
               Left            =   480
               Picture         =   "elitemap.frx":1412
               Top             =   720
               Width           =   360
            End
            Begin VB.Shape Shape1 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   3
               Height          =   255
               Left            =   120
               Top             =   600
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label sOldPeople 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "P"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   0
               Left            =   1200
               MousePointer    =   99  'Custom
               TabIndex        =   17
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label sOldExit 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "E"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   0
               Left            =   960
               MousePointer    =   99  'Custom
               TabIndex        =   16
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label sOldTrap 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "T"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF00FF&
               Height          =   255
               Index           =   0
               Left            =   720
               MousePointer    =   99  'Custom
               TabIndex        =   15
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label sOldSign 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "S"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FFFF&
               Height          =   255
               Index           =   0
               Left            =   480
               MousePointer    =   99  'Custom
               TabIndex        =   14
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
         End
      End
      Begin VB.Label lblLvlScript 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
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
         Left            =   6720
         TabIndex        =   115
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "[22] Left"
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   111
         Top             =   6000
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "[23] Right"
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   110
         Top             =   6240
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "[24] Middle"
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   109
         Top             =   6480
         Width           =   975
      End
      Begin VB.Label lblRom 
         BackStyle       =   0  'Transparent
         Caption         =   "[15] No ROM loaded"
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
         Left            =   120
         TabIndex        =   21
         Top             =   60
         Width           =   3135
      End
      Begin VB.Label lblLevelName 
         BackStyle       =   0  'Transparent
         Caption         =   "[16 ] No level loaded"
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
         Left            =   3240
         TabIndex        =   20
         Top             =   60
         Width           =   3255
      End
   End
   Begin VB.PictureBox picMainTab 
      BackColor       =   &H00F4E0E0&
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   3
      Left            =   4200
      ScaleHeight     =   233
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   187
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ListBox lstPatches 
         Height          =   3255
         IntegralHeight  =   0   'False
         ItemData        =   "elitemap.frx":14C0
         Left            =   120
         List            =   "elitemap.frx":14CD
         TabIndex        =   188
         Top             =   120
         Width           =   1935
      End
      Begin VB.PictureBox picPatches 
         BackColor       =   &H00E6BBBA&
         BorderStyle     =   0  'None
         Height          =   2775
         Index           =   1
         Left            =   2280
         ScaleHeight     =   2775
         ScaleWidth      =   5295
         TabIndex        =   200
         Top             =   480
         Width           =   5295
         Begin VB.CommandButton cmdApplyInsertCredit 
            Caption         =   "Apply"
            Height          =   375
            Left            =   3960
            TabIndex        =   201
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00E6BBBA&
            BorderStyle     =   0  'None
            Height          =   2295
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   223
            Text            =   "elitemap.frx":1504
            Top             =   0
            Width           =   5055
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Patch made by Kyoufu Kawa."
            Height          =   255
            Left            =   120
            TabIndex        =   202
            Top             =   2400
            Width           =   3735
         End
      End
      Begin VB.PictureBox picPatches 
         BackColor       =   &H00E6BBBA&
         BorderStyle     =   0  'None
         Height          =   2775
         Index           =   100
         Left            =   2280
         ScaleHeight     =   2775
         ScaleWidth      =   5295
         TabIndex        =   191
         Top             =   480
         Width           =   5295
         Begin VB.CommandButton cmdApplyBlankSlate 
            Caption         =   "Apply"
            Height          =   375
            Left            =   3960
            TabIndex        =   192
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00E6BBBA&
            BorderStyle     =   0  'None
            Height          =   2175
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   224
            Text            =   "elitemap.frx":174C
            Top             =   0
            Width           =   5055
         End
         Begin VB.Label Label53 
            BackStyle       =   0  'Transparent
            Caption         =   "Patch made by Interdpth."
            Height          =   255
            Left            =   120
            TabIndex        =   195
            Top             =   2400
            Width           =   3735
         End
      End
      Begin VB.PictureBox picPatches 
         BackColor       =   &H00E6BBBA&
         BorderStyle     =   0  'None
         Height          =   2775
         Index           =   0
         Left            =   2280
         ScaleHeight     =   2775
         ScaleWidth      =   5295
         TabIndex        =   189
         Top             =   480
         Width           =   5295
         Begin VB.TextBox Text3 
            BackColor       =   &H00E6BBBA&
            BorderStyle     =   0  'None
            Height          =   2175
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   225
            Text            =   "elitemap.frx":1913
            Top             =   0
            Width           =   5055
         End
         Begin VB.CommandButton cmdApplyNoIntro 
            Caption         =   "Apply"
            Height          =   375
            Left            =   3960
            TabIndex        =   190
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label52 
            BackStyle       =   0  'Transparent
            Caption         =   "Patch made by Kyoufu Kawa."
            Height          =   255
            Left            =   120
            TabIndex        =   194
            Top             =   2400
            Width           =   3735
         End
      End
      Begin VB.PictureBox picPatches 
         BackColor       =   &H00E6BBBA&
         BorderStyle     =   0  'None
         Height          =   2775
         Index           =   99
         Left            =   2280
         ScaleHeight     =   2775
         ScaleWidth      =   5295
         TabIndex        =   193
         Top             =   480
         Width           =   5295
         Begin VB.TextBox Text4 
            BackColor       =   &H00E6BBBA&
            BorderStyle     =   0  'None
            Height          =   2175
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   226
            Text            =   "elitemap.frx":1A48
            Top             =   0
            Width           =   5055
         End
         Begin VB.CommandButton cmdSubmitPatch2 
            Caption         =   "Send E-Mail"
            Height          =   375
            Left            =   3840
            TabIndex        =   197
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CommandButton cmdSubmitPatch 
            Caption         =   "Visit the Protoboard"
            Height          =   375
            Left            =   1920
            TabIndex        =   196
            Top             =   2280
            Width           =   1815
         End
      End
      Begin VB.Label lblPatchName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<patchname>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   198
         Top             =   150
         Width           =   1740
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H00E6BBBA&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   255
         Left            =   2280
         Top             =   360
         Width           =   4575
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00D69896&
         X1              =   505
         X2              =   505
         Y1              =   24
         Y2              =   218
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00D69896&
         X1              =   505
         X2              =   152
         Y1              =   217
         Y2              =   217
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00D69896&
         X1              =   151
         X2              =   151
         Y1              =   23
         Y2              =   218
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00D69896&
         X1              =   369
         X2              =   152
         Y1              =   23
         Y2              =   23
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00D69896&
         X1              =   144
         X2              =   144
         Y1              =   0
         Y2              =   232
      End
      Begin VB.Shape shpPatchName 
         BackColor       =   &H00E6BBBA&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00D69896&
         Height          =   375
         Left            =   5520
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   2070
      End
   End
   Begin VB.PictureBox picMainTab 
      BackColor       =   &H00F4E0E0&
      BorderStyle     =   0  'None
      Height          =   2895
      Index           =   2
      Left            =   4200
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   37
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.OptionButton subtab 
         Caption         =   "[37] Signs"
         Height          =   255
         Index           =   5
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   179
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton subtab 
         Caption         =   "[36] Traps"
         Height          =   255
         Index           =   4
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton subtab 
         Caption         =   "[35] Exits"
         Height          =   255
         Index           =   3
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton subtab 
         Caption         =   "[34] People"
         Height          =   255
         Index           =   2
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton subtab 
         Caption         =   "[33] Connections"
         Height          =   255
         Index           =   1
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   120
         Width           =   1095
      End
      Begin VB.OptionButton subtab 
         Caption         =   "[32] Labels"
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   120
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.PictureBox picSubEditor 
         BackColor       =   &H00F4E0E0&
         Height          =   1455
         Index           =   4
         Left            =   120
         ScaleHeight     =   93
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   493
         TabIndex        =   48
         Top             =   600
         Visible         =   0   'False
         Width           =   7455
         Begin VB.HScrollBar vsbTraps 
            Height          =   255
            Left            =   0
            Max             =   0
            TabIndex        =   119
            Top             =   0
            Width           =   7455
         End
         Begin VB.CommandButton cmdWipeTraps 
            Caption         =   "[97] Wipe"
            Height          =   375
            Left            =   6480
            TabIndex        =   55
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtTrapY 
            Height          =   285
            Left            =   2400
            TabIndex        =   54
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtTrapX 
            Height          =   285
            Left            =   1440
            TabIndex        =   53
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtTrapValue 
            Height          =   285
            Left            =   2400
            TabIndex        =   52
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtTrapFlag 
            Height          =   285
            Left            =   1440
            TabIndex        =   51
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtTrapScript 
            Height          =   285
            Left            =   1440
            TabIndex        =   50
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton cmdRepointTraps 
            Caption         =   "[98] Repoint"
            Height          =   375
            Left            =   6480
            TabIndex        =   49
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "[49] Location"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label40 
            BackStyle       =   0  'Transparent
            Caption         =   "[50] Flags"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label39 
            BackStyle       =   0  'Transparent
            Caption         =   "[47] Script"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.PictureBox picSubEditor 
         BackColor       =   &H00F4E0E0&
         Height          =   1215
         Index           =   5
         Left            =   120
         ScaleHeight     =   77
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   493
         TabIndex        =   59
         Top             =   600
         Visible         =   0   'False
         Width           =   7455
         Begin VB.HScrollBar vsbSigns 
            Height          =   255
            Left            =   0
            Max             =   0
            TabIndex        =   118
            Top             =   0
            Width           =   7455
         End
         Begin VB.CommandButton cmdWipeSigns 
            Caption         =   "[97] Wipe"
            Height          =   375
            Left            =   6480
            TabIndex        =   64
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSignScript 
            Height          =   285
            Left            =   1440
            TabIndex        =   63
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtSignX 
            Height          =   285
            Left            =   1440
            TabIndex        =   62
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtSignY 
            Height          =   285
            Left            =   2400
            TabIndex        =   61
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton cmdRepointSigns 
            Caption         =   "[98] Repoint"
            Height          =   375
            Left            =   6480
            TabIndex        =   60
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label46 
            BackStyle       =   0  'Transparent
            Caption         =   "[47] Script"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label44 
            BackStyle       =   0  'Transparent
            Caption         =   "[49] Location"
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.PictureBox picSubEditor 
         BackColor       =   &H00F4E0E0&
         CausesValidation=   0   'False
         Height          =   1455
         Index           =   3
         Left            =   120
         ScaleHeight     =   93
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   493
         TabIndex        =   38
         Top             =   600
         Visible         =   0   'False
         Width           =   7455
         Begin VB.HScrollBar vsbExits 
            Height          =   255
            Left            =   0
            Max             =   0
            TabIndex        =   120
            Top             =   0
            Width           =   7455
         End
         Begin VB.CommandButton cmdRepointExits 
            Caption         =   "[98] Repoint"
            Height          =   375
            Left            =   6480
            TabIndex        =   39
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmdWipeExits 
            Caption         =   "[97] Wipe"
            Height          =   375
            Left            =   6480
            TabIndex        =   44
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtExitY 
            Height          =   285
            Left            =   2400
            TabIndex        =   43
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtExitX 
            Height          =   285
            Left            =   1440
            TabIndex        =   42
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtExitTarget 
            Height          =   285
            Left            =   1440
            TabIndex        =   41
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtExitLevel 
            Height          =   285
            Left            =   1440
            TabIndex        =   40
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   "[49] Location"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "[53] Exit #"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label28 
            BackStyle       =   0  'Transparent
            Caption         =   "[45] Level"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.PictureBox picSubEditor 
         BackColor       =   &H00F4E0E0&
         Height          =   1575
         Index           =   1
         Left            =   120
         ScaleHeight     =   101
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   493
         TabIndex        =   82
         Top             =   600
         Visible         =   0   'False
         Width           =   7455
         Begin VB.CommandButton cmdConnRepoint 
            Caption         =   "[98] Repoint"
            Height          =   375
            Left            =   6480
            TabIndex        =   185
            Top             =   360
            Width           =   975
         End
         Begin VB.HScrollBar vsbConn 
            Height          =   255
            Left            =   0
            Max             =   0
            TabIndex        =   122
            Top             =   0
            Width           =   7455
         End
         Begin VB.TextBox txtConnLevel 
            Height          =   285
            Left            =   1440
            TabIndex        =   85
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtConnOffset 
            Height          =   285
            Left            =   1440
            TabIndex        =   84
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox cboConnDir 
            Height          =   315
            ItemData        =   "elitemap.frx":1C0B
            Left            =   1440
            List            =   "elitemap.frx":1C0D
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "[45] Level"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label25 
            BackStyle       =   0  'Transparent
            Caption         =   "[44] Offset"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "[43] Direction"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.PictureBox picSubEditor 
         BackColor       =   &H00F4E0E0&
         Height          =   2175
         Index           =   2
         Left            =   120
         ScaleHeight     =   141
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   493
         TabIndex        =   67
         Top             =   600
         Visible         =   0   'False
         Width           =   7455
         Begin VB.PictureBox picSpritestrip 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            Height          =   375
            Left            =   5520
            ScaleHeight     =   21
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   45
            TabIndex        =   220
            Top             =   1440
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.PictureBox picSprite 
            Height          =   1020
            Left            =   5520
            ScaleHeight     =   64
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   49
            TabIndex        =   219
            Top             =   360
            Visible         =   0   'False
            Width           =   795
            Begin VB.VScrollBar vsbPeepSprite 
               Height          =   960
               LargeChange     =   16
               Left            =   480
               Max             =   255
               TabIndex        =   221
               Top             =   0
               Width           =   255
            End
         End
         Begin VB.CheckBox chkPeepIsTrainer 
            BackColor       =   &H00F4E0E0&
            Caption         =   "[51] Is a trainer"
            Height          =   495
            Left            =   0
            TabIndex        =   184
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtPeepLOS 
            Height          =   285
            Left            =   2400
            TabIndex        =   183
            Top             =   1800
            Width           =   855
         End
         Begin VB.HScrollBar vsbPeeps 
            Height          =   255
            Left            =   0
            Max             =   0
            TabIndex        =   121
            Top             =   0
            Width           =   7455
         End
         Begin VB.CommandButton cmdWipePeople 
            Caption         =   "[97] Wipe"
            Height          =   375
            Left            =   6480
            TabIndex        =   76
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtPeepY 
            Height          =   285
            Left            =   2400
            TabIndex        =   75
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtPeepX 
            Height          =   285
            Left            =   1440
            TabIndex        =   74
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtPeepFlag 
            Height          =   285
            Left            =   4440
            TabIndex        =   73
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboPeepBehave 
            Height          =   315
            ItemData        =   "elitemap.frx":1C0F
            Left            =   1440
            List            =   "elitemap.frx":1CFA
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   1080
            Width           =   3855
         End
         Begin VB.TextBox txtPeepScript 
            Height          =   285
            Left            =   1440
            TabIndex        =   71
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtPeepSprite 
            Height          =   285
            Left            =   1440
            TabIndex        =   70
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton cmdRepointPeople 
            Caption         =   "[98] Repoint"
            Height          =   375
            Left            =   6480
            TabIndex        =   69
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmdPeepBecome 
            Caption         =   "[96] Become..."
            Height          =   375
            Left            =   6480
            TabIndex        =   68
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label57 
            BackStyle       =   0  'Transparent
            Caption         =   "[52] Range"
            Height          =   255
            Left            =   1560
            TabIndex        =   203
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label41 
            BackStyle       =   0  'Transparent
            Caption         =   "[49] Location"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "[50] Flags"
            Height          =   255
            Left            =   3120
            TabIndex        =   80
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "[48] Behavior"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label31 
            BackStyle       =   0  'Transparent
            Caption         =   "[47] Script"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label27 
            BackStyle       =   0  'Transparent
            Caption         =   "[46] Sprite"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.PictureBox picSubEditor 
         BackColor       =   &H00F4E0E0&
         Height          =   2175
         Index           =   0
         Left            =   120
         ScaleHeight     =   141
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   493
         TabIndex        =   89
         Top             =   600
         Visible         =   0   'False
         Width           =   7455
         Begin VB.CommandButton cmdSaveName 
            Caption         =   "[99] Save"
            Height          =   375
            Left            =   6480
            TabIndex        =   180
            Top             =   360
            Width           =   975
         End
         Begin VB.HScrollBar vsbDummy 
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            Max             =   0
            TabIndex        =   123
            Top             =   0
            Width           =   7455
         End
         Begin VB.TextBox txtLabelLocH 
            Height          =   285
            Left            =   5400
            TabIndex        =   96
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtLabelLocW 
            Height          =   285
            Left            =   4200
            TabIndex        =   95
            Top             =   1080
            Width           =   615
         End
         Begin VB.CommandButton cmdSaveLocs 
            Caption         =   "[99] Save"
            Height          =   375
            Left            =   6480
            TabIndex        =   94
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtLabelLocY 
            Height          =   285
            Left            =   5400
            TabIndex        =   93
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtLabelLocX 
            Height          =   285
            Left            =   4200
            TabIndex        =   92
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtLabel 
            Height          =   285
            Left            =   4200
            TabIndex        =   91
            Top             =   360
            Width           =   1815
         End
         Begin VB.ListBox lstLabelID 
            Height          =   1695
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "[41] by"
            Height          =   255
            Left            =   4800
            TabIndex        =   102
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "[40] Size"
            Height          =   255
            Left            =   2520
            TabIndex        =   101
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "[41] by"
            Height          =   255
            Left            =   4800
            TabIndex        =   100
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "[39] Location"
            Height          =   255
            Left            =   2520
            TabIndex        =   99
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "[38] Label"
            Height          =   255
            Left            =   2520
            TabIndex        =   98
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "[42] Label editing is here! Go wild, but watch out; when the text goes red, you know you're going too far."
            Height          =   495
            Left            =   2520
            TabIndex        =   97
            Top             =   1560
            Width           =   4695
         End
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00D69896&
         X1              =   0
         X2              =   512
         Y1              =   32
         Y2              =   32
      End
   End
   Begin VB.Label Label54 
      BackStyle       =   0  'Transparent
      Caption         =   "Oooh... old, hidden shit..."
      Height          =   255
      Left            =   4560
      TabIndex        =   199
      Top             =   8520
      Width           =   3135
   End
   Begin VB.Image imgBarRight 
      Height          =   360
      Left            =   11880
      Top             =   0
      Width           =   105
   End
   Begin VB.Image imgBarLeft 
      Height          =   360
      Left            =   45
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   2430
      Left            =   105
      Top             =   5745
      Width           =   3870
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   990
      Left            =   105
      Top             =   2865
      Width           =   3870
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F4E0E0&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2640
      TabIndex        =   181
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1950
      Left            =   30
      Top             =   600
      Width           =   4125
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00D69896&
      X1              =   484
      X2              =   484
      Y1              =   2
      Y2              =   20
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00D69896&
      X1              =   436
      X2              =   436
      Y1              =   2
      Y2              =   20
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00D69896&
      Height          =   7485
      Left            =   4185
      Top             =   705
      Width           =   7725
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   855
      Left            =   120
      Top             =   4440
      Width           =   2130
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1110
      Left            =   45
      Top             =   4260
      Width           =   4005
   End
   Begin VB.Image imgToolBtn 
      Height          =   240
      Index           =   11
      Left            =   8760
      Picture         =   "elitemap.frx":2925
      Tag             =   "web"
      ToolTipText     =   "[259] Visit HR website"
      Top             =   45
      Width           =   240
   End
   Begin VB.Image imgToolBtn 
      Height          =   240
      Index           =   10
      Left            =   8400
      Picture         =   "elitemap.frx":2D37
      Tag             =   "launch"
      ToolTipText     =   "[258] Toolbelt"
      Top             =   45
      Width           =   240
   End
   Begin VB.Image imgToolBtn 
      Height          =   240
      Index           =   7
      Left            =   8040
      Picture         =   "elitemap.frx":3103
      Tag             =   "viewscript"
      ToolTipText     =   "[257] View level script"
      Top             =   45
      Width           =   240
   End
   Begin VB.Image imgToolBtn 
      Height          =   240
      Index           =   6
      Left            =   7680
      Picture         =   "elitemap.frx":34B5
      Tag             =   "resize"
      ToolTipText     =   "[256] Resize level"
      Top             =   45
      Width           =   240
   End
   Begin VB.Image imgToolBtn 
      Height          =   240
      Index           =   5
      Left            =   7320
      Picture         =   "elitemap.frx":386E
      Tag             =   "clear"
      ToolTipText     =   "[255] Clear level"
      Top             =   45
      Width           =   240
   End
   Begin VB.Image imgToolBtn 
      Height          =   240
      Index           =   4
      Left            =   6960
      Picture         =   "elitemap.frx":3BE9
      Tag             =   "copytileset"
      ToolTipText     =   "[254] Copy Tileset bitmap"
      Top             =   45
      Width           =   240
   End
   Begin VB.Image imgToolBtn 
      Height          =   240
      Index           =   3
      Left            =   6600
      Picture         =   "elitemap.frx":3FCB
      Tag             =   "copylevel"
      ToolTipText     =   "[253] Copy level bitmap"
      Top             =   45
      Width           =   240
   End
   Begin VB.Image imgToolBtn 
      Height          =   240
      Index           =   2
      Left            =   6240
      Picture         =   "elitemap.frx":439F
      Tag             =   "gohome"
      ToolTipText     =   "[252] Go to Home level"
      Top             =   45
      Width           =   240
   End
   Begin VB.Image imgToolBtn 
      Height          =   240
      Index           =   1
      Left            =   5880
      Picture         =   "elitemap.frx":477F
      Tag             =   "save"
      ToolTipText     =   "[251] Save level"
      Top             =   45
      Width           =   240
   End
   Begin VB.Image imgToolBtn 
      Height          =   240
      Index           =   0
      Left            =   3120
      Picture         =   "elitemap.frx":4B2E
      Tag             =   "browse"
      ToolTipText     =   "[250] Browse for ROM"
      Top             =   45
      Width           =   240
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "ROM"
      Height          =   285
      Left            =   240
      TabIndex        =   161
      Top             =   75
      Width           =   495
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "[4] World Map"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   159
      Top             =   5520
      Width           =   3855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "[3] Miscellaneous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   158
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "[2] Attributes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   157
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00F4E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "[1] Tileset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   156
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblTilesetLoc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F4E0E0&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1560
      TabIndex        =   155
      ToolTipText     =   "Did you know that you can double click me to see various nifty offsets?"
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00F4E0E0&
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3000
      TabIndex        =   154
      Top             =   360
      Width           =   1095
   End
   Begin VB.Shape toolhilite 
      BorderColor     =   &H00D69896&
      Height          =   315
      Left            =   8730
      Shape           =   4  'Rounded Rectangle
      Top             =   15
      Width           =   315
   End
   Begin VB.Image toolbar 
      Height          =   360
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
   Begin VB.Image backdrop 
      Height          =   960
      Left            =   0
      Stretch         =   -1  'True
      Top             =   360
      Width           =   960
   End
   Begin VB.Menu mnuBecome 
      Caption         =   "become"
      Visible         =   0   'False
      Begin VB.Menu mnuBecomePerson 
         Caption         =   "[200] Person"
      End
      Begin VB.Menu mnuBecomeTrainer 
         Caption         =   "[201] Trainer"
      End
      Begin VB.Menu mnuBecomeItem 
         Caption         =   "[202] Item Ball"
      End
   End
   Begin VB.Menu mnuLaunch 
      Caption         =   "launcher"
      Visible         =   0   'False
      Begin VB.Menu mnuLauncher 
         Caption         =   "BaseEdit"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "Bewildered"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "Dexter"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "FontEd"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "PAttEd"
         Enabled         =   0   'False
         Index           =   4
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "PET"
         Enabled         =   0   'False
         Index           =   5
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "RSBall"
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "Spread"
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "ScriptEd"
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "LIPS"
         Enabled         =   0   'False
         Index           =   10
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "PokeCryGUI"
         Enabled         =   0   'False
         Index           =   11
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "PokePic"
         Enabled         =   0   'False
         Index           =   12
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "Sappy"
         Enabled         =   0   'False
         Index           =   13
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "SNESEdit"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu mnuLauncher 
         Caption         =   "TLP"
         Enabled         =   0   'False
         Index           =   15
      End
   End
   Begin VB.Menu mnuRMB 
      Caption         =   "context"
      Visible         =   0   'False
      Begin VB.Menu cmdLoadExtern 
         Caption         =   "[203] &Load ExMap"
      End
      Begin VB.Menu cmdSaveExtern 
         Caption         =   "[204] &Save ExMap"
      End
      Begin VB.Menu mnuRMBColors 
         Caption         =   "[208] Set &theme"
      End
   End
   Begin VB.Menu mnuMapRMB 
      Caption         =   "maprmb"
      Visible         =   0   'False
      Begin VB.Menu mnuSetBoyStartHere 
         Caption         =   "[205] Set boy's starting position here"
      End
      Begin VB.Menu mnuSetGirlStartHere 
         Caption         =   "[206] Set girl's starting position here"
      End
   End
   Begin VB.Menu mnuWorldMap 
      Caption         =   "worldmap"
      Visible         =   0   'False
      Begin VB.Menu mnuWorldMapChange 
         Caption         =   "[207] Load another bitmap"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBivsbTileset As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBivsbTileset As Any) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Type Connect2Header
  wConnects As Long
  pConnects As Long
End Type

Private Type ConnectHeader
  wDirection As Long
  wOffset As Long
  b1 As Byte
  hLevel As Integer
  b2 As Byte
End Type

Private Type LevelHeader
  pOldMap As Long
  pSprites As Long
  pScript As Long
  pConnect As Long
  hSong As Integer
  hMap As Integer
  bLabelID As Byte
  bFlash As Byte
  bWeather As Byte
  bType As Byte
  bUnused1 As Byte
  bUnused2 As Byte
  bLabelToggle As Byte
  bUnused3 As Byte
End Type

Private Type TilesetHeader
  b1 As Byte
  b2 As Byte
  b3 As Byte
  b4 As Byte
  pGFX As Long
  pPalettes As Long
  pMap As Long
  pBehavior As Long
  pAnimation As Long
End Type

Private Type SpriteHeader
  bPeople As Byte
  bExits As Byte
  bTraps As Byte
  bSigns As Byte
  pPeople As Long
  pExits As Long
  pTraps As Long
  pSigns As Long
End Type

Private Type PeopleHeader
  b1 As Byte
  bSpriteSet As Byte
  b3 As Byte
  b4 As Byte
  bX As Byte
  b6 As Byte
  bY As Byte
  b8 As Byte
  b9 As Byte
  bBehavior1 As Byte
  b10 As Byte
  bBehavior2 As Byte
  bIsTrainer As Byte
  b14 As Byte
  bTrainerLOS As Byte
  b16 As Byte
  pScript As Long
  iFlag As Integer
  b23 As Byte
  b24 As Byte
End Type

Private Type TrapHeader
  bX As Byte
  b2 As Byte
  bY As Byte
  b4 As Byte
  h3 As Integer
  hFlagCheck As Integer
  hFlagValue As Integer
  h6 As Integer
  pScript As Long
End Type

Private Type ExitHeader
  bX As Byte
  b2 As Byte
  bY As Byte
  b4 As Byte
  b5 As Byte
  b6 As Byte
  hLevel As Integer
End Type

Private Type SignHeader
  bX As Byte
  b2 As Byte
  bY As Byte
  b4 As Byte
  b5 As Byte
  b6 As Byte
  b7 As Byte
  b8 As Byte
  pScript As Long
End Type

Private Type maplocs
  bX As Byte
  bY As Byte
  bW As Byte
  bH As Byte
End Type

Private Const BITSPIXEL = 12
Private Const PLANES = 14
Private Const HORZRES = 8
Private Const OBJ_BITMAP = 7
Private Const VERTRES = 10
Private Const FLOODFILLSURFACE = 1

Private checkifcompa As Byte
Private checkifcompb As Byte
Private checkifpala As Byte
Private checkifpalb As Byte
Private endvalue As Variant

Private worldlocs(0 To &H1FF) As maplocs
Private cmdAdjMaps(0 To 15) As Integer
Private palettesA(0 To &HF, 0 To &HF) As Long
Private palettesB(0 To &HF, 0 To &HF) As Long
Private gfxA(0 To 32767) As Byte
Private gfxB(0 To 32767) As Byte
Private Map16A(0 To 10239) As Byte
Private Map16B(0 To 8191) As Byte
Private lastbank As Byte
Private lastlev As Byte
Private thislevel As LevelHeader
Private thissprite As SpriteHeader
Private thisconnect As Connect2Header
Private thistileseta As TilesetHeader
Private thistilesetb As TilesetHeader
Private xm As Long
Private MapLabels(0 To &H1FF) As String
Private peoples(0 To &H3F) As PeopleHeader
Private exits(0 To &H3F) As ExitHeader
Private traps(0 To &H3F) As TrapHeader
Private signs(0 To &H3F) As SignHeader
Private mapConnects(0 To &H7) As ConnectHeader
Private blankmap(0 To &H3FF, 0 To &H3FF) As Long
Private mapsize As Long
Private attribcolors(0 To &H3F) As Long
Private xd As Long
Private xp As Long
Private xlangs(0 To &HFF)
Private xvers(0 To &HFF)
Private headr As String * 4
Private pat As String
Private tx(0 To &H3FF) As Long
Private ty(0 To &H3FF) As Long
Private bmap(0 To 15) As Byte
Public lheight As Byte
Public lwidth As Byte
Private attribnames(0 To &HFF) As String
Private TileMap(0 To &H3FF, 0 To &H3FF) As Long
Private tempTileMap(0 To &H3FF, 0 To &H3FF) As Long
Private dirty As Boolean
Private HomeLevel As Integer
Private StampMap(0 To 1, 0 To 1) As Long
Private RomType As Integer
Private romisjapanese As Boolean
Private oldlbllen As Integer

Public mytheme As Integer

Private EMPath As String

Private MyHeader As tRomHackHeader

Private Sub cboConnDir_Click()
  mapConnects(vsbConn.Value).wDirection = cboConnDir.ListIndex + 1
  renderconnects
  dirty = True
End Sub

Private Sub cboHackLanguage_Click()
  MyHeader.iLanguage = cboHackLanguage.ListIndex
End Sub

Private Sub cboLabelID_Click()
  thislevel.bLabelID = cboLabelID.ListIndex
  If NextGen = True Then
    thislevel.bLabelID = cboLabelID.ListIndex + &H58
  End If
  dirty = True
End Sub

Private Sub cboPeepBehave_Click()
  peoples(vsbPeeps).bBehavior1 = cboPeepBehave.ListIndex
End Sub

Private Sub cboSong_Click()
  thislevel.hSong = cboSong.ListIndex
  dirty = True
End Sub

Private Sub cboType_Click()
  thislevel.bType = cboType.ListIndex
  dirty = True
End Sub

Private Sub cboWeather_Click()
  thislevel.bWeather = cboWeather.ListIndex
  dirty = True
End Sub

Private Sub chkNoDraw_Click()
  Shape1.Visible = IIf(chkNoDraw.Value = 1, False, True)
End Sub

Private Sub chkPeepIsTrainer_Click()
  peoples(vsbPeeps).bIsTrainer = chkPeepIsTrainer.Value
  dirty = True
End Sub

Private Sub chkShowLabel_Click()
  thislevel.bLabelToggle = chkShowLabel.Value
  dirty = True
End Sub

Private Sub cboBanks_Click()
  Open txtRom For Binary As #256
    LOADLevels
  Close #256
  txtOldskoolChooser = Right("00" & Hex(cboBanks.ListIndex), 2) & Right("00" & Hex(cboLevels.ListIndex), 2)
End Sub

Private Sub cboLevels_click()
  cboLevels.Enabled = False
  txtOldskoolChooser = Right("00" & Hex(cboBanks.ListIndex), 2) & Right("00" & Hex(cboLevels.ListIndex), 2)
  cmdLoad_Click
  cboLevels.Enabled = True
  cboLevels.Tag = cboLevels.ListIndex
End Sub

Private Sub cmdAdjMap_Click(Index As Integer)
  If dirty = True Then
    If MsgBox("Continue loading new map and lose your changes to this one?", vbYesNo, "Changes not saved") = vbNo Then Exit Sub
  End If
  dirty = False
  cmdAdjMap(Index).Enabled = False
  For i = 0 To 13
  cmdAdjMap(i).Enabled = False
  Next i
  write_bank_lev cmdAdjMaps(Index) \ 100, cmdAdjMaps(Index) Mod 100
  lastbank = cmdAdjMaps(Index) \ 100
  lastlev = cmdAdjMaps(Index) Mod 100
  cboLevels.Enabled = True
End Sub

Private Sub chkSExits_Click()
  rendersprites
End Sub

Private Sub cmdBrowse_Click()
  exit2 = False
  'tlbToolbar.Buttons(12).Enabled = False
  LoadRom False
  'If txtRom <> "" Then tlbToolbar.Buttons(12).Enabled = True
  cboLevels.Clear
  cboLevels.Text = "Level"
  If txtRom <> "" Then Exit Sub
  cboBanks.Clear
  cboBanks.Text = "Bank"
End Sub

Private Sub cmdGoHome_Click()
  write_bank_lev HomeLevel \ 256, HomeLevel Mod 256
End Sub

Private Sub cmdApplyInsertCredit_Click()
  Dim palptr As Long
  Dim gfxptr As Long
  Dim mapptr As Long
  If Roms(RomType).Code = "AXVE" Or Roms(RomType).Code = "AXPE" Or Roms(RomType).Code = "AXKW" Then
    palptr = InputBox(LoadResString(100), , "&H")
    If palptr = 0 Then Exit Sub
    gfxptr = InputBox(LoadResString(101), , "&H")
    If gfxptr = 0 Then Exit Sub
    mapptr = InputBox(LoadResString(102), , "&H")
    If mapptr = 0 Then Exit Sub
    
    Open txtRom For Binary As #16
    Seek #16, &H13BA9A + 1
    'Palette
    Put #16, , CByte(&H5)
    Put #16, , CByte(&H48)
    Put #16, , CByte(&H5)
    Put #16, , CByte(&H49)
    Put #16, , CByte(&H12)
    Put #16, , CByte(&HDF)
    'GFX
    Put #16, , CByte(&H5)
    Put #16, , CByte(&H48)
    Put #16, , CByte(&H6)
    Put #16, , CByte(&H49)
    Put #16, , CByte(&H12)
    Put #16, , CByte(&HDF)
    'Map
    Put #16, , CByte(&H6)
    Put #16, , CByte(&H48)
    Put #16, , CByte(&H6)
    Put #16, , CByte(&H49)
    Put #16, , CByte(&H12)
    Put #16, , CByte(&HDF)
    'Branch over the pool
    Put #16, , CByte(&HD)
    Put #16, , CByte(&HE0)
    'Pool
    Put #16, , CByte(&H0)
    Put #16, , CByte(&H0)
    Put #16, , palptr + &H8000000
    Put #16, , CLng(&H5000000)
    Put #16, , gfxptr + &H8000000
    Put #16, , CLng(&H6000000)
    Put #16, , mapptr + &H8000000
    Put #16, , CLng(&H60037C0)
    Close #16
    MsgBox LoadResString(103)
    '0548 0549 12DF 0548 0649 12DF 0648 0649 12DF 0DE0 0000
    '70346C08 00000005 D0346C08 00000006 00406C08 BE370006
    '  PAL               GFX               MAP
  Else
    MsgBox LoadResString(104)
  End If
End Sub

Private Sub cmdApplyNoIntro_Click()
  If Roms(RomType).Code = "AXVE" Or Roms(RomType).Code = "AXPE" Or Roms(RomType).Code = "AXKW" Then
    Open txtRom For Binary As #16
    Seek #16, &H13BA9A + 1
    For i = 0 To 360
      Put #16, , CByte(0)
    Next i
    Close #16
    MsgBox LoadResString(103)
  Else
    MsgBox LoadResString(104)
  End If
End Sub

Private Sub cmdApplyBlankSlate_Click()
  'MATT - This Patch Clears The truck event, Starts the game in littleroot, and destroys every people event. ~ Matt
  'Kawa not sure how elitemap works so I did my best.
  'KAWA -- Don't worry li'l buddy. </samnmax> I'll fix your code wherever needed.
  
  If Roms(RomType).Code = "AXVE" Or Roms(RomType).Code = "AXPE" Or Roms(RomType).Code = "AXKW" Then
    'KAWA -- Added a warning message
    If MsgBox(LoadResString(105), vbYesNo) = vbNo Then Exit Sub
    
    Open txtRom For Binary As #16
    'KAWA -- Well now, you got some serious file number assignment issues here...
    Seek #16, &H9BBCCD + 1 'TODO -- Find CORRECT offset!
    Put #16, , CByte(&H0)
    Put #16, , CByte(&H0)
    Put #16, , CByte(&H2)
    
    Put #16, &H52E0E + 1, CByte(&H0)  'KAWA -- You can hide the Seek command -in- the Put/Get ;)
    Put #16, &H52E10 + 1, CByte(&H9)
    Close #16
    
    'MsgBox "This will now kill all people events on this map."
    'KAWA -- Why tell them you're committing genocide? They don't even get a choice!
    thislevel.pScript = &H89BBCCD  'KAWA -- No need to recalculate a pointer now. Also saves cycles.
    '------------------------------
    For i = 0 To vsbPeeps.Max
      peoples(i).b1 = 0
      peoples(i).b3 = 0
      peoples(i).b4 = 0
      peoples(i).b6 = 0
      peoples(i).b8 = 0
      peoples(i).b9 = 0
      peoples(i).b10 = 0
      peoples(i).b14 = 0
      peoples(i).b16 = 0
      peoples(i).b23 = 0
      peoples(i).b24 = 0
      peoples(i).bBehavior1 = 0
      peoples(i).bBehavior2 = 0
      peoples(i).bSpriteSet = 0
      peoples(i).bIsTrainer = 0
      peoples(i).bTrainerLOS = 0
      peoples(i).bX = 0
      peoples(i).bY = 0
      peoples(i).iFlag = 0
      peoples(i).pScript = 0
    Next i
    thissprite.bPeople = 0
    vsbPeeps.Value = 0
    vsbPeeps.Max = 0
    rendersprites
    'MsgBox "Now changing starting map to little root to by pass truck routine will be bug in corner"
    'MsgBox "During Game Play can fix by placing correct tiles over not my fault"
    
    MsgBox LoadResString(103)
  Else
    MsgBox LoadResString(104)
  End If
End Sub

Private Sub cmdConnRepoint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim n As Long
  If Shift = vbCtrlMask Then
    If MsgBox(LoadResString(106), vbYesNo) = vbYes Then
      n = Val(InputBox(LoadResString(107), , "&H" & Right("000000" & Hex(thisconnect.pConnects), 6)))
      i = Val(InputBox(LoadResString(108), , thisconnect.wConnects))
      thisconnect.wConnects = i
      thisconnect.pConnects = n + &H8000000
      vsbConn.Value = 0
      vsbConn.Max = i - 1
      rendersprites
    End If
  Else
    MsgBox LoadResString(109), vbInformation
  End If
End Sub

Private Sub cmdFullBRD_Click()
  AdvancedBorder.Left = Form1.Left + 5250
  AdvancedBorder.Top = Form1.Top + 7500
  AdvancedBorder.Show
  'TODO -- Add always-on-top mode for border window.
  draw_ng_border
End Sub

'Private Sub cmdPanel_Click(Index As Integer)
'  If cmdPanel(Index).Tag = "" Then
'    cmdPanel(Index).Tag = "^_^"
'    picPanel(Index).Height = Val(picPanel(Index).Tag)
'  Else
'    cmdPanel(Index).Tag = ""
'    picPanel(Index).Height = cmdPanel(Index).Height + 1
'  End If
'  picPanel(1).Top = picPanel(0).Top + picPanel(0).Height
'  picPanel(2).Top = picPanel(1).Top + picPanel(1).Height
'  picPanel(3).Top = picPanel(2).Top + picPanel(2).Height
'End Sub

Private Sub cmdPeepBecome_Click()
  PopupMenu mnuBecome
End Sub

Private Sub cmdRepointExits_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim n As Long
  If Shift = vbCtrlMask Then
    If MsgBox(LoadResString(106), vbYesNo) = vbYes Then
      n = Val(InputBox(LoadResString(107), , "&H" & Right("000000" & Hex(thissprite.pExits), 6)))
      i = Val(InputBox(LoadResString(110), , thissprite.bExits))
      thissprite.bExits = i
      thissprite.pExits = n + &H8000000
      vsbExits.Value = 0
      vsbExits.Max = 0
      rendersprites
    End If
  Else
    MsgBox LoadResString(109), vbInformation
  End If
End Sub

Private Sub cmdRepointSigns_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim n As Long
  If Shift = vbCtrlMask Then
    If MsgBox(LoadResString(106), vbYesNo) = vbYes Then
      n = Val(InputBox(LoadResString(107), , "&H" & Right("000000" & Hex(thissprite.pSigns), 6)))
      i = Val(InputBox(LoadResString(111), , thissprite.bSigns))
      thissprite.bSigns = i
      thissprite.pSigns = n + &H8000000
      vsbSigns.Value = 0
      vsbSigns.Max = i
      rendersprites
    End If
  Else
    MsgBox LoadResString(109), vbInformation
  End If
End Sub

Private Sub cmdRepointPeople_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim n As Long
  If Shift = vbCtrlMask Then
    If MsgBox(LoadResString(106), vbYesNo) = vbYes Then
      n = Val(InputBox(LoadResString(107), , "&H" & Right("000000" & Hex(thissprite.pPeople), 6)))
      i = Val(InputBox(LoadResString(112), , thissprite.bPeople))
      thissprite.bPeople = i
      thissprite.pPeople = n + &H8000000
      vsbPeeps.Value = 0
      vsbPeeps.Max = i
      rendersprites
    End If
  Else
    MsgBox LoadResString(109), vbInformation
  End If
End Sub

Private Sub cmdRepointTraps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim n As Long
  If Shift = vbCtrlMask Then
    If MsgBox(LoadResString(106), vbYesNo) = vbYes Then
      n = Val(InputBox(LoadResString(107), , "&H" & Right("000000" & Hex(thissprite.pTraps), 6)))
      i = Val(InputBox(LoadResString(113), , thissprite.bTraps))
      thissprite.bTraps = i
      thissprite.pTraps = n + &H8000000
      vsbTraps.Value = 0
      vsbTraps.Max = i
      rendersprites
    End If
  Else
    MsgBox LoadResString(109), vbInformation
  End If
End Sub

Private Sub cmdSaveExtern_Click()
  Dim wite As Byte
  Dim wite2 As Byte
  Dim woo As Integer
  Dim hdr As String * 8
  hdr = "ELITEMAP"
  wite = &H81
  
  On Error GoTo Hell
  'cdlCommon.flags = cdlCommonOFNHideReadOnly + cdlCommonOFNLongNames + cdlCommonOFNFileMustExist
  'cdlCommon.Filter = "EliteMap exmaps (*.emap)|*.emap|All Files (*.*)|*.*"
  'cdlCommon.Tag = cdlCommon.Filename
  'cdlCommon.Filename = ""
  'cdlCommon.ShowOpen
  Dim cc As cCommonDialog
  Set cc = New cCommonDialog
  Dim heyhey As String
  If Not cc.VBGetSaveFileName(heyhey, , , "EliteMap exmaps (*.emap)|*.emap|All Files (*.*)|*.*", , , , "emap") Then Exit Sub
  
  Open heyhey For Binary As #256
  
  Put #256, , hdr
  Put #256, , wite
  
  'Put #256, , thistileseta
  'Put #256, , thistilesetb
  
  Put #256, , lwidth
  Put #256, , lheight
    
  For i = 0 To lheight - 1
    For j = 0 To lwidth - 1
      'wite = TileMap(j, i) Mod &H100
      'wite2 = TileMap(j, i) \ &H100
      'Put #256, , wite
      'Put #256, , wite2
      Put #256, , TileMap(j, i)
    Next j
  Next i
  
Hell:
  Close #256
  'cdlCommon.Filename = cdlCommon.Tag

End Sub

Private Sub cmdLoadExtern_Click()
  Dim wite As Byte
  Dim wite2 As Byte
  Dim hdr As String * 8
  
  On Error GoTo Hell
  Dim cc As cCommonDialog
  Set cc = New cCommonDialog
  Dim heyhey As String
  If Not cc.VBGetOpenFileName(heyhey, , , , , , "EliteMap exmaps (*.emap)|*.emap|All Files (*.*)|*.*", , , , "emap") Then Exit Sub
  Open heyhey For Binary As #256
  
  Get #256, , hdr
  Get #256, , wite
  If hdr = "ELITEMAP" And wite = &H81 Then
  Else
    If wite = &H80 Then
      MsgBox LoadResString(114)
    Else
      MsgBox LoadResString(115), vbExclamation, LoadResString(116)
    End If
    Exit Sub
  End If
      
  'Get #256, , thistileseta
  'Get #256, , thistilesetb
  
  Get #256, , lwidth
  Get #256, , lheight
  thismap.wWidth = lwidth
  thismap.wHeight = lheight
  txtLevelWidth = Hex(lwidth)
  txtLevelHeight = Hex(lheight)
    
  Call refreshlevel
    
  For i = 0 To lheight - 1
    For j = 0 To lwidth - 1
      'Get #256, , wite
      'Get #256, , wite2
      'TileMap(j, i) = (CLng(wite2) * CLng(&H100)) + wite
      Get #256, , TileMap(j, i)
    Next j
  Next i
    
Hell:
  Close #256
  refreshlevel
  dirty = True
End Sub

Private Sub cmdLvlScript_Click()
  CallScriptEd CLng(Val("&H" & lblLvlScript)) + &H8000000
End Sub

Private Sub refreshlevel(Optional ByVal movelevel As Boolean = False)
  If movelevel = False Then
    Picture1.Move 0, 0, lwidth * &H10, lheight * &H10
    't.Move 0, 0, lwidth, lheight
    hsbScroll = 0
    vsbScroll = 0
    hsbScroll.Max = 0
    vsbScroll.Max = 0
    If lwidth > p.Width \ &H10 Then hsbScroll.Max = (lwidth) - (p.Width \ &H10)
    If lheight > p.Height \ &H10 Then vsbScroll.Max = (lheight) - (p.Height \ &H10)
  End If
  For i = 0 To lheight - 1
    For j = 0 To lwidth - 1
      DrawTile TileMap(j, i), j, i
    Next j
  Next i
  Picture1.Refresh
End Sub

Private Sub cmdLoad_Click()
  If dirty = True Then
    If MsgBox(LoadResString(117), vbYesNo, LoadResString(118)) = vbNo Then Exit Sub
  End If
  
  If txtRom = "" Then
    MsgBox LoadResString(119), vbInformation, LoadResString(120)
    Exit Sub
  End If
  
  'tlbToolbar.Buttons(5).Enabled = False
  
  MousePointer = 11
  
  imgBoyStart.Left = -500
  imgGirlStart.Left = -500
  
  Picture1.Move 0, 0
  Picture1.Cls
  Picture1.BackColor = p.BackColor
  picThrobber.Tag = ""
  DoEvents
  Shape1.Visible = False
  
  Dim point As Long
  Dim Width As Byte
  Dim Height As Byte
  Dim wite As Byte
  Dim wite2 As Byte
  Dim headr As String * 2
  Dim ver As String * 1
  Dim lang As String * 1
  
  Open txtRom For Binary As #256
  
#If GENTLE_LOAD_ERRORS Then
  On Error GoTo Robot_Hell_Bonanza
#End If

  Text2 = ""
  get_bank_lev
  'Nine out of ten sociopaths agree
  CopyMemory TileMap(0, 0), blankmap(0, 0), 4194304
  CopyMemory tempTileMap(0, 0), blankmap(0, 0), 4194304
  point = getgbapointer((Val(txtbank) * 4) + xd)
  If point = -1 Then
    MousePointer = 0
    MsgBox LoadResString(121)
    Close #256
    Exit Sub
  End If
  
  point = getgbapointer((Val(txtlevel) * 4) + point)
  If point = -1 Then
    MousePointer = 0
    MsgBox LoadResString(122)
    Close #256
    Exit Sub
  End If
  
  Picture1.Enabled = True
  picTileset.Enabled = True
  
  lpoint = point
  Get #256, point + 1, thislevel
  If NextGen = False Then
    lblLevelName = Hex(lpoint) & ": " & MapLabels(thislevel.bLabelID)
  Else
  '  'TODO -- Add nextgen label support
    lblLevelName = Hex(lpoint) & ": " & MapLabels(thislevel.bLabelID - &H58)
  End If
  lblLvlScript = Hex(GBA2PC(thislevel.pScript))
  
  shMap.Move (worldlocs(thislevel.bLabelID - IIf(NextGen = True, 88, 0)).bX + 1) * 8, (worldlocs(thislevel.bLabelID - IIf(NextGen = True, 88, 0)).bY + 2) * 8, worldlocs(thislevel.bLabelID - IIf(NextGen = True, 88, 0)).bW * 8, worldlocs(thislevel.bLabelID - IIf(NextGen = True, 88, 0)).bH * 8
  shLoc.Move (worldlocs(thislevel.bLabelID - IIf(NextGen = True, 88, 0)).bX + 1) * 8, (worldlocs(thislevel.bLabelID - IIf(NextGen = True, 88, 0)).bY + 2) * 8, worldlocs(thislevel.bLabelID - IIf(NextGen = True, 88, 0)).bW * 8, worldlocs(thislevel.bLabelID - IIf(NextGen = True, 88, 0)).bH * 8
  shMap.Visible = True
  shLoc.Visible = True
  
  'You gotta see Hyakugoyuichi
  point = GBA2PC(thislevel.pSprites)
  Get #256, point + 1, thissprite
  point = GBA2PC(thislevel.pConnect)
  If point <> -1 Then
    Get #256, point + 1, thisconnect
  Else
    thisconnect.wConnects = 0
  End If
  If thissprite.bPeople > 0 Then
    point = GBA2PC(thissprite.pPeople)
    For d = 0 To thissprite.bPeople - 1
      Get #256, point + (d * 24) + 1, peoples(d)
    Next d
  End If
  If thissprite.bExits > 0 Then
    point = GBA2PC(thissprite.pExits)
    For d = 0 To thissprite.bExits - 1
      Get #256, point + (d * 8) + 1, exits(d)
    Next d
  End If
  If thissprite.bTraps > 0 Then
    point = GBA2PC(thissprite.pTraps)
    For d = 0 To thissprite.bTraps - 1
      Get #256, point + (d * 16) + 1, traps(d)
    Next d
  End If
  If thissprite.bSigns > 0 Then
    point = GBA2PC(thissprite.pSigns)
    For d = 0 To thissprite.bSigns - 1
      Get #256, point + (d * 12) + 1, signs(d)
    Next d
  End If
  If thisconnect.wConnects > 0 Then
    point = GBA2PC(thisconnect.pConnects)
    For d = 0 To thisconnect.wConnects - 1
      Get #256, point + (d * 12) + 1, mapConnects(d)
    Next d
  End If
  
  If thissprites <> -1 Then rendersprites
  
' KAWA --- Let me show you how ;)
  picSubEditor(1).Enabled = IIf(thisconnect.wConnects > 0, True, False)
  picSubEditor(2).Enabled = IIf(thissprite.bPeople > 0, True, False)
  picSubEditor(3).Enabled = IIf(thissprite.bExits > 0, True, False)
  picSubEditor(4).Enabled = IIf(thissprite.bTraps > 0, True, False)
  picSubEditor(5).Enabled = IIf(thissprite.bSigns > 0, True, False)
  
  'From the Moch to the Rie to the Pee to the Wee
  point = getgbapointer(((thislevel.hMap * 4) - 4) + xp)
  'Debug.Print Hex(point)
  If point = -1 Then
    MousePointer = 0
    MsgBox LoadResString(123)
    Close #256
    Exit Sub
  End If
  dpoint = point
  Get #256, point + 1, thismap
  
  'For later use in repointing if things are too big
  allheadersize = thismap.wHeight * thismap.wWidth + 28
  If NextGen = True Then allheadersize = allheadersize + 4 + thismap.bBorderX * thismap.bBorderY
  
  'point = getgbapointer(point)
  lwidth = thismap.wWidth
  lheight = thismap.wHeight
  mapsize = lwidth * CLng(lheight) * CLng(2)
  'Picture1.Visible = False
  t1 = GBA2PC(thismap.pTilesetA)
  t2 = GBA2PC(thismap.pTilesetB)
  Get #256, t1 + 1, thistileseta
  Get #256, t2 + 1, thistilesetb
  Get #256, t1 + 1, checkifcompa
  Get #256, t2 + 1, checkifcompb
  Get #256, t1 + 2, checkifpala
  Get #256, t2 + 2, checkifpalb
  lblTilesetLoc = Hex(thismap.pTilesetB)
  picTileset.BackColor = 0
  CopyMemory palettesA(0, 0), blankmap(0, 0), &H400
  CopyMemory palettesB(0, 0), blankmap(0, 0), &H400
  CopyMemory gfxA(0), blankmap(0, 0), 32768
  CopyMemory gfxB(0), blankmap(0, 0), 32768
  'Different MAP16Asize in NextGen
  CopyMemory Map16A(0), blankmap(0, 0), 10240
  CopyMemory Map16B(0), blankmap(0, 0), 8192
  
  'picTileset.Visible = False
  ApplyTileset 0, thistileseta
  ApplyTileset 1, thistilesetb
  picTileset.Refresh
  picTileset.Visible = True
  drawtilehdc seltile(0), picSel(0).hdc, 0, 0
  drawtilehdc seltile(1), picSel(1).hdc, 0, 0
  drawtilehdc seltile(2), picSel(2).hdc, 0, 0
  picSel(0).Refresh
  picSel(1).Refresh
  picSel(2).Refresh
  
  'Just take it from me, MC NC
  'MsgBox "Border pointer at 0x" & Hex(thismap.pBorder)
  point = GBA2PC(thismap.pBorder)
  If point = -1 Then
    MousePointer = 0
    MsgBox LoadResString(124)
    Close #256
    Exit Sub
  End If
  'MsgBox "Border pointer at 0x" & Hex(point)
  picBorder.ToolTipText = "Ptr: 0x" & Hex(point)
  If NextGen = False Or (thismap.bBorderX = 2 And thismap.bBorderY = 2) Then
    Get #256, point, wite
    Get #256, , wite
    Get #256, , wite2
    border(0, 0) = h2d(wite2 & Right("00" & Hex(wite), 2))
    Get #256, , wite
    Get #256, , wite2
    border(0, 1) = h2d(wite2 & Right("00" & Hex(wite), 2))
    Get #256, , wite
    Get #256, , wite2
    border(1, 0) = h2d(wite2 & Right("00" & Hex(wite), 2))
    Get #256, , wite
    Get #256, , wite2
    border(1, 1) = h2d(wite2 & Right("00" & Hex(wite), 2))
    drawtilehdc border(0, 0), picBorder.hdc, 0, 0
    drawtilehdc border(0, 1), picBorder.hdc, 1, 0
    drawtilehdc border(1, 0), picBorder.hdc, 0, 1
    drawtilehdc border(1, 1), picBorder.hdc, 1, 1
    picBorder.Refresh
  Else
    Get #256, point, wite
    If thismap.bBorderX = 0 Then thismap.bBorderX = 1 'Like in the rom
    If thismap.bBorderY = 0 Then thismap.bBorderY = 1
    yind = 0
    For i = 0 To thismap.bBorderY - 1 ' Fill NextGenArray
      xind = 0
      For ii = 0 To thismap.bBorderX - 1
        Get #256, , wite
        Get #256, , wite2
        Borderitems(yind * thismap.bBorderX + xind) = h2d(wite2 & Right("00" & Hex(wite), 2)) 'Fill up Array with variables
        xind = xind + 1
      Next ii
      yind = yind + 1
    Next i
    Set_Border thismap.bBorderX, thismap.bBorderY
    picBorder.Refresh
  End If
  
  'You won't believe your eyes you'll go insane
  point = GBA2PC(thismap.pMap)
  If point = -1 Then
    MousePointer = 0
    MsgBox LoadResString(125)
    Close #256
    Exit Sub
  End If
  Get #256, point, wite
  oldmapadd = point
  For i = 0 To lheight - 1
    For j = 0 To lwidth - 1
      Get #256, , wite
      Get #256, , wite2
      'wite2 = wite2 Mod 4
      'x = x & " " & Chr(IIf(wite < 32, wite + 32, wite)) '& Chr(IIf(wite2 < 32, wite2 + 32, wite2))
      TileMap(j, i) = (CLng(wite2) * CLng(&H100)) + wite
    Next j
    'x = x & vbCrLf
  Next i
  
  refreshlevel
  Picture1.Visible = True
  'Text2 = x
  
  If Roms(RomType).StartPosBoy <> 0 Then
    Seek #256, Roms(RomType).StartPosBoy + 1
    Get #256, , wite
    Get #256, , wite2
    If wite = cboBanks.ListIndex And wite2 = cboLevels.ListIndex Then
      Get #256, , wite
      Get #256, , wite
      Get #256, , wite2
      imgBoyStart.Left = ((CLng(wite2) * CLng(&H100)) + wite) * 16
      Get #256, , wite
      Get #256, , wite2
      imgBoyStart.Top = ((CLng(wite2) * CLng(&H100)) + wite) * 16
    End If
  End If
  If Roms(RomType).StartPosGirl <> 0 Then
    Seek #256, Roms(RomType).StartPosGirl + 1
    Get #256, , wite
    Get #256, , wite2
    If wite = cboBanks.ListIndex And wite2 = cboLevels.ListIndex Then
      Get #256, , wite
      Get #256, , wite
      Get #256, , wite2
      imgGirlStart.Left = ((CLng(wite2) * CLng(&H100)) + wite) * 16
      Get #256, , wite
      Get #256, , wite2
      imgGirlStart.Top = ((CLng(wite2) * CLng(&H100)) + wite) * 16
    End If
  End If
  
  Close #256
  
  renderconnects
  
  'I mean what's up with that plastic plane?
  txtLevelWidth = "&H" & Hex(lwidth)
  txtLevelHeight = "&H" & Hex(lheight)
  
  'tlbToolbar.Buttons("save").Enabled = True
  'tlbToolbar.Buttons("copylevel").Enabled = True
  'tlbToolbar.Buttons("copytileset").Enabled = True
  'tlbToolbar.Buttons("clear").Enabled = True
  'tlbToolbar.Buttons("resize").Enabled = True
  'tlbToolbar.Buttons("viewscript").Enabled = True
  
  'You're an idiot if you disagree
  If NextGen = False Then
    cboLabelID.ListIndex = thislevel.bLabelID
  Else
    cboLabelID.ListIndex = thislevel.bLabelID - &H58
  End If
  Trace Hex(thislevel.hSong)
  If thislevel.hSong < &H1FF Then
      cboSong.ListIndex = thislevel.hSong
      cboSong.Visible = True
      lblSongWarning.Visible = False
  Else
      cboSong.Visible = False
      lblSongWarning.Visible = True
  End If
  If thislevel.bFlash < 2 Then chkAllowFlash.Value = thislevel.bFlash
  If thislevel.bLabelToggle < 2 Then chkShowLabel.Value = thislevel.bLabelToggle
  cboWeather.ListIndex = thislevel.bWeather
  cboType.ListIndex = thislevel.bType
  vsbConn.Value = 0
  vsbConn.Max = thisconnect.wConnects - 1
  vsbConn_Change
  vsbPeeps.Value = 0
  vsbPeeps.Max = thissprite.bPeople - 1
  vsbPeeps_Change
  vsbExits.Value = 0
  vsbExits.Max = thissprite.bExits - 1
  vsbExits_Change
  vsbTraps.Value = 0
  vsbTraps.Max = thissprite.bTraps - 1
  vsbTraps_Change
  vsbSigns.Value = 0
  vsbSigns.Max = thissprite.bSigns - 1
  vsbSigns_Change
  
  dirty = False

  'You gotta see Hyakugoyuichi!
  MousePointer = 0
  'tlbToolbar.Buttons(5).Enabled = True
  picThrobber.Tag = "stop!"
  timThrobber.Tag = "0"

  Exit Sub
  
Robot_Hell_Bonanza:
  'Oh crap. Singing! Mind if I smoke?
  Open "Crud Vapors.txt" For Output As #13
  Print #13, "Data report for unloadable map"
  Print #13, "------------------------------"
  Print #13, "Bank/Level: " & txtOldskoolChooser.Text
  Print #13, ""
  Print #13, "thislevel..."
  Print #13, "  pOldMap: " & Hex(thislevel.pOldMap)
  Print #13, "  pSprites: " & Hex(thislevel.pSprites)
  Print #13, "  pScript: " & Hex(thislevel.pScript)
  Print #13, "  pConnect: " & Hex(thislevel.pConnect)
  Print #13, "  hSong: " & Hex(thislevel.hSong)
  Print #13, "  hMap: " & Hex(thislevel.hMap)
  Print #13, "  bLabelID: " & Hex(thislevel.bLabelID) & " (" & lblLevelName & ")"
  Print #13, "  bWeather: " & Hex(thislevel.bWeather)
  Print #13, "  bType: " & Hex(thislevel.bType)
  Print #13, "  bUnused1: " & Hex(thislevel.bUnused1)
  Print #13, "  bUnused2: " & Hex(thislevel.bUnused2)
  Print #13, "  bLabelToggle: " & Hex(thislevel.bLabelToggle)
  Print #13, "  bUnused3: " & Hex(thislevel.bUnused3)
  Close #13
  If MsgBox(Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
            "The map's header has been written to ""Crud Vapors.txt"" and can be sent " & _
            "to The Helmeted Rodent for assistance." & vbCrLf & vbCrLf & _
            "Ignore this error and continue at your own risk?", vbYesNo) = vbYes Then
    Resume Next
  Else
    End
  End If
End Sub

Private Function getgbapointer(ByVal offset As Long, Optional fn As Integer = 256)
  Dim a(0 To 3) As Byte
  Get fn, offset + 1, a(0)
  Get fn, offset + 2, a(1)
  Get fn, offset + 3, a(2)
  Get fn, offset + 4, a(3)
  If a(3) = 8 Then
    getgbapointer = (CLng(a(2)) * CLng(&H10000)) + (CLng(a(1)) * CLng(&H100)) + a(0)
  Else
    getgbapointer = -1
  End If
End Function

Private Function GBA2PC(ByVal gbapoint As Long) As Long
  Dim a(0 To 3) As Byte
  a(0) = gbapoint Mod 256
  a(1) = (gbapoint \ CLng(256)) Mod 256
  a(2) = (gbapoint \ CLng(256 ^ 2)) Mod 256
  a(3) = (gbapoint \ CLng(256 ^ 3)) Mod 256
  If a(3) = 8 Then
    GBA2PC = (CLng(a(2)) * CLng(&H10000)) + (CLng(a(1)) * CLng(&H100)) + a(0)
  Else
    GBA2PC = -1
  End If
End Function

Private Sub cmdClear_Click()
  i = MsgBox("ARE YOU SURE?", vbYesNo)
  If i = vbYes Then
    For i = 0 To lheight - 1
      For j = 0 To lwidth - 1
        TileMap(j, i) = seltile(1) + selattr(1)
        DrawTile seltile(1), j, i
      Next j
    Next i
    Picture1.Refresh
  End If
  dirty = True
End Sub

Private Sub cmdReplaceRL_Click()
  For i = 0 To lheight - 1
    For j = 0 To lwidth - 1
      If TileMap(j, i) Mod &H400 = seltile(1) Then TileMap(j, i) = seltile(0) + selattr(0)
    Next j
  Next i
  refreshlevel
  dirty = True
End Sub

Private Sub cmdSaveLocs_Click()
  'If MsgBox("This has just been coded. You sure? It might destroy your location table...", vbYesNo) = vbNo Then Exit Sub
  Open txtRom For Binary As #80
    Do While i < &H59
      Put #80, (xm + 1) + (i * 8), worldlocs(i)
      i = i + 1
    Loop
  Close #80
End Sub

Private Sub cmdSaveName_Click()
  Dim pc As Long
  Dim xpc As Long
  Dim i As Long
  Dim newdata As String
  
  Open txtRom For Binary As #32
  Do While i < &H58
    Get #32, (xm + 1) + (i * 8), worldlocs(i)
    pc = getgbapointer((xm + 4) + (i * 8), 32) + 1
    xpc = pc
    newdata = Asc2Sapp(MapLabels(i)) & Chr$(255)
    Put #32, pc, newdata
    i = i + 1
  Loop
  Close #32
End Sub

Private Sub cmdSubmitPatch2_Click()
  ShellExecute 0, vbNullString, "mailto:helmeted@helmetedrodent.kickassgamers.com?subject=New patch", vbNullString, "", 1
End Sub

Private Sub cmdSwapRL_Click()
  For i = 0 To lheight - 1
    For j = 0 To lwidth - 1
      If TileMap(j, i) Mod &H400 = seltile(1) Then
        TileMap(j, i) = seltile(0) + selattr(0)
      ElseIf TileMap(j, i) Mod &H400 = seltile(0) Then
        TileMap(j, i) = seltile(1) + selattr(1)
      End If
    Next j
  Next i
  refreshlevel
  dirty = True
End Sub

Private Sub cmdSetLAtts_Click()
  For i = 0 To lheight - 1
    For j = 0 To lwidth - 1
      If TileMap(j, i) Mod &H400 = seltile(0) Then TileMap(j, i) = seltile(0) + selattr(0)
    Next j
  Next i
  refreshlevel
  dirty = True
End Sub

Private Sub cmdShiftLeft_Click()
  For Y = 0 To lheight - 1
    For X = 0 To lwidth - 1
      tempTileMap(X, Y) = TileMap(((X + 1) Mod lwidth), Y)
    Next X
  Next Y
  CopyMemory TileMap(0, 0), tempTileMap(0, 0), &H400 * CLng(&H400) * CLng(4)
  refreshlevel True
  dirty = True
End Sub

Private Sub cmdShiftUp_Click()
  For Y = 0 To lheight - 1
    For X = 0 To lwidth - 1
      tempTileMap(X, Y) = TileMap(X, (Y + 1) Mod lheight)
    Next X
  Next Y
  CopyMemory TileMap(0, 0), tempTileMap(0, 0), &H400 * CLng(&H400) * CLng(4)
  refreshlevel True
  dirty = True
End Sub

Private Sub cmdShiftDown_Click()
  For Y = 0 To lheight - 1
    For X = 0 To lwidth - 1
      tempTileMap(X, Y) = TileMap(X, ((Y + lheight) - 1) Mod lheight)
    Next X
  Next Y
  CopyMemory TileMap(0, 0), tempTileMap(0, 0), &H400 * CLng(&H400) * CLng(4)
  refreshlevel True
  dirty = True
End Sub

Private Sub cmdShiftRight_Click()
  For Y = 0 To lheight - 1
    For X = 0 To lwidth - 1
      tempTileMap(X, Y) = TileMap((((X + lwidth) - 1) Mod lwidth), Y)
    Next X
  Next Y
  CopyMemory TileMap(0, 0), tempTileMap(0, 0), &H400 * CLng(&H400) * CLng(4)
  refreshlevel True
  dirty = True
End Sub

Private Sub cmdCopyLevel_Click()
  Picture1.Picture = Picture1.Image
  Clipboard.Clear
  Clipboard.SetData Picture1.Picture, 2
End Sub

Private Sub DrawTile(ByVal tileno, ByVal destX, ByVal destY)
  'MsgBox "DRAW TILE LOG [test]"
  'BitBlt Picture1.hDC, (destx + 2) * 16, (desty + 2) * 16, 16, 16, pTiles.hDC, tx(tileno) * 16, ty(tileno) * 16, SRCCOPY
  td = tileno Mod &H400
  ta = tileno \ &H400
  BitBlt Picture1.hdc, (destX) * 16, (destY) * 16, 16, 16, picTileset.hdc, (td Mod 16) * 16, (td \ 16) * 16, SRCCOPY
End Sub

Public Sub drawtilehdc(ByVal tileno, ByVal hdc, ByVal destX, ByVal destY)
  'MsgBox "DRAW TILE HDC LOG"
  'BitBlt Picture1.hDC, (destx + 2) * 16, (desty + 2) * 16, 16, 16, pTiles.hDC, tx(tileno) * 16, ty(tileno) * 16, SRCCOPY
  BitBlt hdc, (destX) * 16, (destY) * 16, 16, 16, picTileset.hdc, (tileno Mod 16) * 16, (tileno \ 16) * 16, SRCCOPY
End Sub

Private Sub LoadRom(Optional HandGrenade As Boolean = False)
  On Error GoTo Hell
  If HandGrenade = True Then GoTo SkipABitBrotherMaynard
  'cdlCommon.flags = cdlCommonOFNHideReadOnly + cdlCommonOFNLongNames + cdlCommonOFNFileMustExist
  'cdlCommon.Filter = "GBA ROM Dumps (*.gba;*.bin)|*.gba;*.bin|All Files (*.*)|*.*"
  'cdlCommon.ShowOpen
  Dim cc As cCommonDialog
  Set cc = New cCommonDialog
  Dim t As String
  Dim t2 As String
  Dim check As Integer
  t = txtRom.Text
  t2 = txtRom.Text
  If cc.VBGetOpenFileName(t, , , , , , "GBA roms (*.gba)|*.gba", , App.Path, , , Me.hwnd, OFN_HIDEREADONLY) Then
    txtRom = t
  Else
    GoTo Hell
  End If
  On Error GoTo 0
  'txtRom = cdlCommon.Filename
SkipABitBrotherMaynard:
  Close #256
  Open txtRom For Binary As #256
  Get #256, &HAD, headr
  i = FindRom(headr)
  If i = -1 Then
      MsgBox LoadResString(126) & vbCrLf & headr & ".", vbExclamation
      exit2 = True
      Close #256
      Exit Sub
  End If
  CheckLock txtRom
  Get #256, &HC1, check
  If check = &H3713 And FindRom("AXKW") Then i = FindRom("AXKW")
  
  If Roms(i).MapHeaders = 0 Or Roms(i).Maps = 0 Or Roms(i).MapLabels = 0 Then
    MsgBox LoadResString(127)
    lblRom = Roms(i).Code & " - " & Roms(i).Name & " (no info)"
    Close #256
    Exit Sub
  End If
  
  xd = getgbapointer(Roms(i).MapHeaders) 'getgbapointer(340772)
  xp = getgbapointer(Roms(i).Maps) 'getgbapointer(340588)
  xm = getgbapointer(Roms(i).MapLabels) 'getgbapointer(1032160)
  If xp = -1 Or xd = -1 Then
    MsgBox LoadResString(128) '  & vbCrLf & vbCrLf & _
           "MapHeaders points to " & Hex(xd) & "," & vbCrLf & _
           "Maps points to " & Hex(xp) & "," & vbCrLf & _
           "MapLabels points to " & Hex(xm) & "."
    Close #256
    Exit Sub
  End If
  lblRom = Roms(i).Code & " - " & Roms(i).Name
  
  Dim NewHeader As tRomHackHeader
  Dim c2 As Byte
  MyHeader = NewHeader
  txtHackName = Trim(MyHeader.sName)
  txtAuthorName = Trim(MyHeader.sAuthor)
  txtGroupName = Trim(MyHeader.sGroup)
  cboHackLanguage.ListIndex = 0
  lblGUID = "GUID: " & MakeStringFromGUID(MyHeader.gGUID)
  lblWorkTime = LoadResString(406)
  timWorkTimer.Enabled = False
  Get #256, &HCE, c2
  If (c2 = 1) And (LOF(256) > &H1000000) Then
    i = ReadHeader(txtRom, NewHeader)
    If i Then
      MyHeader = NewHeader
      txtHackName = Trim(MyHeader.sName)
      txtAuthorName = Trim(MyHeader.sAuthor)
      txtGroupName = Trim(MyHeader.sGroup)
      cboHackLanguage.ListIndex = MyHeader.iLanguage
      lblGUID = "GUID: " & MakeStringFromGUID(MyHeader.gGUID)
      lblWorkTime_Click
      timWorkTimer.Enabled = True
    End If
  End If
  
  If Roms(i).Language = rlJapanese Then
    romisjapanese = True
  Else
    romisjapanese = False
  End If
  If Roms(i).RomType > 0 Then
    NextGen = True
    'cboSong.Enabled = False
    'cboLabelID.Enabled = False
    picSubEditor(1).Enabled = False
    xm = Roms(i).MapLabels
    maplabelreadNG
    cmdSaveName.Enabled = True
    subtab(0).Enabled = False
    subtab(1).Value = True
  Else
    NextGen = False
    cboSong.Enabled = True
    cboLabelID.Enabled = True
    picSubEditor(1).Enabled = True
    maplabelread
    subtab(0).Enabled = True
    If Roms(i).Language = rlJapanese Then
      txtLabel.Locked = True
      txtLabel.Enabled = False
      cmdSaveName.Enabled = False
    Else
      txtLabel.Locked = False
      txtLabel.Enabled = True
      cmdSaveName.Enabled = True
    End If
  End If
  
  Dim j As Integer
  cboSong.Clear
  For check = 0 To &H1FF
    cboSong.AddItem Hex(check)
  Next check
  If Roms(i).MusicList <> "" Then
    If Dir(Roms(i).MusicList) <> "" Then
      Open Roms(i).MusicList For Input As #1
      Input #1, check
      j = check
      Do
        Line Input #1, t
        cboSong.List(j) = Hex(j) & ". " & t
        j = j + 1
      Loop Until EOF(1)
      Close #1
    End If
  End If
  
  If Roms(i).WorldMap <> "" Then
    On Error Resume Next 'should be NoMap but there's only one statement :P
    picWorldMap.Picture = LoadPicture(Roms(i).WorldMap)
  End If
  
  picSprite.Visible = False
  If Dir(Left(txtRom, Len(txtRom) - 4) & " sprites.bmp") <> "" Then
    picSprite.Visible = True
    picSpritestrip.Picture = LoadPicture(Left(txtRom, Len(txtRom) - 4) & " sprites.bmp")
  End If
  
  If Roms(i).HomeLevel > 0 Then
    HomeLevel = Roms(i).HomeLevel
  Else
    HomeLevel = &H9 'Default to LittleRoot Town
  End If
  
  LOADBanks (i)
  
  RomType = i
     
  Close #256
  Exit Sub

'KAWA - Added cancel button support =^.-=
Hell:
  txtRom.Text = t2
  Exit Sub
End Sub

Private Sub cmdWipePeople_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  If Shift = vbCtrlMask Then
    If MsgBox(LoadResString(106), vbYesNo) = vbYes Then
      For i = 0 To 63
        peoples(i).b1 = 0
        peoples(i).b3 = 0
        peoples(i).b4 = 0
        peoples(i).b6 = 0
        peoples(i).b8 = 0
        peoples(i).b9 = 0
        peoples(i).b10 = 0
        peoples(i).b14 = 0
        peoples(i).b16 = 0
        peoples(i).b23 = 0
        peoples(i).b24 = 0
        peoples(i).bBehavior1 = 0
        peoples(i).bBehavior2 = 0
        peoples(i).bSpriteSet = 0
        peoples(i).bIsTrainer = 0
        peoples(i).bTrainerLOS = 0
        peoples(i).bX = 0
        peoples(i).bY = 0
        peoples(i).iFlag = 0
        peoples(i).pScript = 0
      Next i
      thissprite.bPeople = 0
      vsbPeeps.Value = 0
      vsbPeeps.Max = 0
      rendersprites
    End If
  Else
    MsgBox LoadResString(129), vbInformation
  End If
End Sub

Private Sub cmdWipeExits_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  If Shift = vbCtrlMask Then
    If MsgBox(LoadResString(106), vbYesNo) = vbYes Then
      For i = 0 To 63
        exits(i).b2 = 0
        exits(i).b4 = 0
        exits(i).b5 = 0
        exits(i).b6 = 0
        exits(i).bX = 0
        exits(i).bY = 0
        exits(i).hLevel = 0
      Next i
      thissprite.bExits = 0
      vsbExits.Value = 0
      vsbExits.Max = 0
      rendersprites
    End If
  Else
    MsgBox LoadResString(129), vbInformation
  End If
End Sub

Private Sub cmdWipeTraps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  If Shift = vbCtrlMask Then
    If MsgBox(LoadResString(106), vbYesNo) = vbYes Then
      For i = 0 To 63
        traps(i).b2 = 0
        traps(i).b4 = 0
        traps(i).bX = 0
        traps(i).bY = 0
        traps(i).h3 = 0
        traps(i).h6 = 0
        traps(i).hFlagCheck = 0
        traps(i).hFlagValue = 0
        traps(i).pScript = 0
      Next i
      thissprite.bTraps = 0
      vsbTraps.Value = 0
      vsbTraps.Max = 0
      rendersprites
    End If
  Else
    MsgBox LoadResString(129), vbInformation
  End If
End Sub

Private Sub cmdWipeSigns_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  If Shift = vbCtrlMask Then
    If MsgBox(LoadResString(106), vbYesNo) = vbYes Then
      For i = 0 To 63
        signs(i).b2 = 0
        signs(i).b4 = 0
        signs(i).b5 = 0
        signs(i).b6 = 0
        signs(i).b7 = 0
        signs(i).b8 = 0
        signs(i).bX = 0
        signs(i).bY = 0
        signs(i).pScript = 0
      Next i
      thissprite.bSigns = 0
      vsbSigns.Value = 0
      vsbSigns.Max = 0
      rendersprites
    End If
  Else
    MsgBox LoadResString(129), vbInformation
  End If
End Sub

'Private Sub Command4_Click()
'  For i = 0 To 15
'    If Val(Text2) = 0 Then
'      spal(i).BackColor = palettesA(Text3, i)
'    End If
'  Next i
'End Sub

'Private Sub Command5_Click()
'  DrawMap16 picBorder.hdc, Text4, Text5, 0, 0
'  'DrawTile8 picBorder.hdc, Text4, Text5, 0, 0
'  picBorder.Refresh
'End Sub

'Private Sub Command6_Click()
'  picTileset.BackColor = 0
'  For i = 0 To &H3FF
'    X = i Mod &H20
'    Y = i \ &H20
'    DrawTile8 picTileset.hdc, 0, i + (Text3 * CLng(&H1000)), X * 8, Y * 8
'  Next i
'  picTileset.Refresh
'End Sub

Private Sub cmdResize_Click()
  If lwidth = 0 Then
    MsgBox LoadResString(130), vbInformation + vbOKOnly, LoadResString(131)
    Exit Sub
  End If
  If lheight = 0 Then
    MsgBox LoadResString(130), vbInformation + vbOKOnly, LoadResString(131)
    Exit Sub
  End If

  If txtRom.Text = "" Then Exit Sub
  newwm = Val(txtLevelWidth)
  newhm = Val(txtLevelHeight)
  Resize.Show vbModal
  If modNextGenBorder.noresize = False Then Exit Sub
  thismap.wWidth = newwm
  thismap.wHeight = newhm
  lwidth = newwm
  lheight = newhm
  txtLevelWidth = "&H" & Right("00" & Hex(newwm), 2) 'KAWA -- You missed a spot, Tau.
  txtLevelHeight = "&H" & Right("00" & Hex(newhm), 2)
  refreshlevel
  dirty = True
    
  If NextGen = False Then Exit Sub
  thismap.bBorderX = newwb
  thismap.bBorderY = newhb
End Sub

Private Sub cmdCopyTiles_Click()
  picTileset.Picture = picTileset.Image
  Clipboard.Clear
  Clipboard.SetData picTileset.Picture, 2
End Sub

'Can you find Waldo in this code?
Private Sub cmdSave_Click()
  'Tau, I'm VERY impressed indeed, but people wouldn't like sudden relocations without any warning.
  'I'm putting it in now...
  ' -- Kawa
  
  Dim searchnewplace As Boolean
  If lwidth = 0 Then
    MsgBox LoadResString(130), vbInformation + vbOKOnly, LoadResString(131)
    Exit Sub
  End If
  If lheight = 0 Then
    MsgBox LoadResString(131), vbInformation + vbOKOnly, LoadResString(131)
    Exit Sub
  End If
  If txtRom = "" Then
    MsgBox LoadResString(132), vbInformation, LoadResString(133)
    Exit Sub
  End If
  
  searchnewplace = True
  
  newheadersize = thismap.wHeight * thismap.wWidth + 28
  If NextGen = True Then newheadersize = newheadersize + 4 + thismap.bBorderX * thismap.bBorderY
  
  If newheadersize <= allheadersize Then searchnewplace = False
  
  MousePointer = 11
  
  Dim point As Long
  Dim Width As Byte
  Dim Height As Byte
  Dim wite As Byte
  Dim wite2 As Byte
  Dim headr As String * 2
  Dim ver As String * 1
  Dim lang As String * 1
  Dim c As Byte
  Dim d As Byte
  Open txtRom For Binary As #256
  Text2 = ""
  point = getgbapointer((Val(txtbank) * 4) + xd)
  If point = -1 Then
    MousePointer = 0
    MsgBox LoadResString(121)
    Close #256
    Exit Sub
  End If
  point = getgbapointer((Val(txtlevel) * 4) + point)
  If point = -1 Then
    MousePointer = 0
    MsgBox LoadResString(122)
    Close #256
    Exit Sub
  End If
  
  lpoint = point
  Put #256, point + 1, thislevel
  lblLevelName = Hex(lpoint) & ": " & MapLabels(thislevel.bLabelID)
  
  Form1.Refresh
  point = GBA2PC(thislevel.pConnect)
  If point <> -1 Then
    Put #256, point + 1, thisconnect
  Else
    thisconnect.wConnects = 0
  End If
  point = GBA2PC(thislevel.pSprites)
  Put #256, point + 1, thissprite
  
  If thissprite.bPeople > 0 Then
    point = GBA2PC(thissprite.pPeople)
    For d = 0 To thissprite.bPeople - 1
      Put #256, point + (d * 24) + 1, peoples(d)
    Next d
  End If
  If thissprite.bExits > 0 Then
    point = GBA2PC(thissprite.pExits)
    For d = 0 To thissprite.bExits - 1
      Put #256, point + (d * 8) + 1, exits(d)
    Next d
  End If
  If thissprite.bTraps > 0 Then
    point = GBA2PC(thissprite.pTraps)
    For d = 0 To thissprite.bTraps - 1
      Put #256, point + (d * 16) + 1, traps(d)
    Next d
  End If
  If thissprite.bSigns > 0 Then
    point = GBA2PC(thissprite.pSigns)
    For d = 0 To thissprite.bSigns - 1
      Put #256, point + (d * 12) + 1, signs(d)
    Next d
  End If
no_Events:
  If thisconnect.wConnects > 0 Then
    point = GBA2PC(thisconnect.pConnects)
    For d = 0 To thisconnect.wConnects - 1
      Put #256, point + (d * 12) + 1, mapConnects(d)
    Next d
  End If
  
  writeoffset = getgbapointer(((thislevel.hMap * 4) - 4) + xp)
  oldoffset = writeoffset
  
  If searchnewplace = True Then
    'Okay Tau... I love what you did to the place but those curtains, my god...
    LunarOpenFile txtRom, LC_READWRITE
    start_off = &H200000
search_more:
    writeoffset = LunarVerifyFreeSpace(start_off, &HFFFFF0, newheadersize, LC_NOBANK)
    If h2d(Right(Hex(writeoffset), 1)) <> 0 Then
      start_off = writeoffset + 16 - h2d(Right(Hex(writeoffset), 1)) 'we need an offset with 0,4,8 or 12 at the end because of the ARM7, we only take the ones with 0 at the end because we like it easy ;)
      GoTo search_more
    End If
    
    '-- start of revived code --
    message = LoadResString(134)
    message = Replace(message, "[1]", Right("00000" & Hex(writeoffset), 6))
    message = Replace(message, "[2]", Right("00000" & Hex(oldoffset), 6))
    writeoffset = Val(InputBox(message, LoadResString(135), "&H" & Right("00000" & Hex(writeoffset), 6)))
    If writeoffset = 0 Then
      MsgBox LoadResString(136), vbOKOnly, LoadResString(137) 'How dare you desecrate the Bitch Message? KAWA WILL KILL YOU!
      MousePointer = 0
      Exit Sub
    End If
    '-- end of revived code --
    
    LunarCloseFile
  End If
  
  If writeoffset = -1 Then
    MousePointer = 0
    MsgBox LoadResString(123)
    Close #256
    Exit Sub
  End If
  
  If searchnewplace = True Then GoTo putdirect
  
  Get #256, writeoffset + 1, oldmap
  
resume_searchnewplace:
  'input border
  borderoff = oldmap.pBorder - &H8000000
  If NextGen = False Then
    'Debug.Print Hex(border(0, 0)) & " " & Hex(border(0, 1))
    'Debug.Print Hex(border(1, 0)) & " " & Hex(border(1, 1))
'    Debug.Print
'    Seek #256, borderoff + 1
'    Put #256, , border(0, 1)
    'KAWA - Can't fix this du du du dum wee wee can't fix this
    Seek #256, borderoff + 1
    wite = border(0, 0) Mod &H100
    wite2 = border(0, 0) \ &H100
    Put #256, , wite
    Put #256, , wite2
    wite = border(0, 1) Mod &H100
    wite2 = border(0, 1) \ &H100
    Put #256, , wite
    Put #256, , wite2
    wite = border(1, 0) Mod &H100
    wite2 = border(1, 0) \ &H100
    Put #256, , wite
    Put #256, , wite2
    wite = border(1, 1) Mod &H100
    wite2 = border(1, 1) \ &H100
    Put #256, , wite
    Put #256, , wite2
  End If
  'input border nextgen
  If NextGen = True Then
    yind = 0
    z = 0
    For Y = 0 To thismap.bBorderY - 1
      xind = 0
      For X = 0 To thismap.bBorderX - 1
        c = h2d(Right("00" & Hex(Borderitems(yind * thismap.bBorderX + xind)), 2))
        d = h2d(Left(Right("0000" & Hex(Borderitems(yind * thismap.bBorderX + xind)), 4), 2))
     
        'Put #256, borderoff + 1 + (yind * thismap.bBorderX + xind) * 2, c
        'Put #256, borderoff + 2 + (yind * thismap.bBorderX + xind) * 2, d
        xind = xind + 1
        z = z + 1
      Next X
      yind = yind + 1
    Next Y
  End If
  mapoff = oldmap.pMap - &H8000000
  For Y = 0 To thismap.wHeight - 1
    For X = 0 To thismap.wWidth - 1
      c = h2d(Right("00" & Hex(TileMap(X, Y)), 2))
      d = h2d(Left(Right("0000" & Hex(TileMap(X, Y)), 4), 2))
      Put #256, mapoff + 1 + (Y * thismap.wWidth + X) * 2, c
      Put #256, mapoff + 2 + (Y * thismap.wWidth + X) * 2, d
    Next X
  Next Y
 
  'Save tileset changes and all...
  point = getgbapointer(((thislevel.hMap * 4) - 4) + xp)
  Put #256, point + 1, thismap
 
  Close #256
  allheadersize = newheadersize
  dirty = False
  MousePointer = 0
  MsgBox LoadResString(138)
  Exit Sub

putdirect:
  'delete old completely
  Get #256, getgbapointer(((thislevel.hMap * 4) - 4) + xp) + 1, oldmap
  ''delete old header
  'cntx = 23
  'headroff = getgbapointer(((thislevel.hMap * 4) - 4) + xp)
  'If NextGen = True Then cntx = 29
  'For X = 0 To cntx
  '  Put #256, headroff + 1 + X, 255
  'Next X
  'delete old border
  'borderoff = oldmap.pBorder - &H8000000
  'If nexgen = False Then
  '  oldmap.bBorderX = 2
  '  oldmap.bBorderY = 2
  'End If
  'For Y = 0 To (oldmap.bBorderY) * 2 - 1
  '  For X = 0 To (oldmap.bBorderX) * 2 - 1
  '    Put #256, borderoff + 1 + X + Y * oldmap.bBorderX, 255
  '  Next X
  'Next Y
  'delete old map
  mapoff = oldmap.pMap - &H8000000
  For Y = 0 To (oldmap.wHeight) * 2 - 1
    For X = 0 To (oldmap.wWidth) * 2 - 1
      Put #256, mapoff + 1 + X + Y * oldmap.wHeight, 255
    Next X
  Next Y
  'thismap.pBorder = writeoffset + &H8000000
  thismap.pMap = writeoffset + &H8000000 ' + (thismap.bBorderX * thismap.bBorderY) * 2
  If NextGen = False Then thismap.pMap = writeoffset + 4 + &H8000000
  Put #256, thismap.pMap + (thismap.wHeight * thismap.wWidth) * 2 - &H8000000 + 1, thismap
  Put #256, getgbapointer((Val(txtlevel) * 4) + getgbapointer((Val(txtbank) * 4) + xd)) + 1, GBA2PC(thismap.pMap + (thismap.wHeight * thismap.wWidth) * 2) + &H8000000
  Put #256, thislevel.hMap * 4 - 4 + xp + 1, GBA2PC(thismap.pMap + (thismap.wHeight * thismap.wWidth) * 2) + &H8000000
  oldmap = thismap
  GoTo resume_searchnewplace
End Sub

Private Sub chkSPeople_Click()
  rendersprites
  If chkSPeople.Value = 0 Then
    imgBoyStart.Visible = False
    imgGirlStart.Visible = False
  Else
    imgBoyStart.Visible = True
    imgGirlStart.Visible = True
  End If
End Sub

Private Sub chkSSigns_Click()
  rendersprites
End Sub

Private Sub chkSprites_Click()
  rendersprites
End Sub

Private Sub chkSTraps_Click()
  rendersprites
End Sub

Private Sub Command1_Click()
  If txtRom = "" Then Exit Sub
  If MyHeader.gGUID.part1 = 0 Then MakeGUID MyHeader.gGUID
  lblGUID = "GUID: " & MakeStringFromGUID(MyHeader.gGUID)
  WriteHeader txtRom, MyHeader
End Sub

Private Sub edittab_Click(Index As Integer)
  LockWindowUpdate hwnd
  On Error Resume Next
  For i = 0 To edittab.UBound
    picMainTab(i).Visible = False
  Next i
  picMainTab(Index).Visible = True
  Shape2.Height = picMainTab(Index).Height + 2
  Shape2.Width = picMainTab(Index).Width + 2
  LockWindowUpdate 0
End Sub

Private Sub Form_Load()
  SetIcon Me.hwnd, "APP", True
  lblVersion = "EliteMap™ version " & App.Major & "." & App.Minor & " by Kyoufu Kawa"
  EMPath = App.Path & "\"
  
  i = InStr(txtCredits, "[q]")
  txtCredits = Replace(txtCredits, "[q]", "")
  'MsgBox Format(Date, "ddmm")
  Select Case Format(Date, "ddmm")
    Case "0606": txtCredits = Left(txtCredits, i) & """Yes, my male sim finally got pregnant!"" -- Ailure"
    Case "2606": txtCredits = Left(txtCredits, i) & """Are you seriously leeching it?™"" -- Kawa"
    Case "2706": txtCredits = Left(txtCredits, i) & """Did you know that there's only two days" & vbCrLf & "between our birthdays, Matt?"" -- Kawa"
    Case "2806": txtCredits = Left(txtCredits, i) & """This is the day that I shall be laid hurray!"" -- Interdpth"
    Case "1507": txtCredits = Left(txtCredits, i) & """404"" -- Ranko"
    Case "1408": txtCredits = Left(txtCredits, i) & """Insert quote here"" -- Hiryuu"
    Case "1810": txtCredits = Left(txtCredits, i) & """your dick, your dick, your dick is on fiar"" -- DJ ßouché"
    Case "1811": txtCredits = Left(txtCredits, i) & """posting a meaningles thread does not" & vbCrLf & "  mean you have a bigger penis."" -- Majin Bluedragon"
   'Case "ddmm": txtCredits = Left(txtCredits, i) & """Your quote here"" -- Your name here"
  End Select
  txtCredits = Replace(txtCredits, "[t]", LoadResString(156))
  
  InitDatabase
    
  sSign(0).MouseIcon = LoadResPicture(106, vbResCursor)
  sTrap(0).MouseIcon = LoadResPicture(105, vbResCursor)
  sExit(0).MouseIcon = LoadResPicture(104, vbResCursor)
  sPeople(0).MouseIcon = LoadResPicture(103, vbResCursor)
  picTeam.Picture = LoadResPicture(1, 0)
  picWorldMap.Picture = LoadResPicture(2, 0)
  picAttributes.Picture = LoadResPicture(3, 0)
  picThrobberPics.Picture = LoadResPicture(5, 0)
  
  On Error Resume Next 'We ignore all errors here.
  mytheme = Val(INIRead("elitemap", "Shared", "Theme"))
  If mytheme = 0 Then mytheme = 10
  backdrop.Picture = LoadResPicture(mytheme, 0)
  toolbar.Picture = LoadResPicture(mytheme + 1, 0)
  imgBarLeft.Picture = LoadResPicture(mytheme + 3, 0)
  imgBarRight.Picture = LoadResPicture(mytheme + 3, 0)
  'And now, we go through ALL controls available to recolor...
  guipal.Picture = LoadResPicture(mytheme + 2, 0)
  Dim ColorRemap(1 To 2, 1 To 4) As Long
  For X = 1 To 2
    For Y = 1 To 4
      ColorRemap(X, Y) = guipal.point((Y - 1) * 8, (X - 1) * 8)
      guipal.PSet ((Y - 1) * 8, (X - 1) * 8), vbRed
      'Debug.Print X & " x " & Y & " = " & Hex(ColorRemap(X, Y))
    Next Y
  Next X
  For Each ctl In Me.Controls
    'Debug.Print Ctl.Name & " " & Hex(Ctl.BorderColor)
    'Debug.Print Ctl.Name
    For Y = 1 To 4
      If ctl.BorderColor = ColorRemap(1, Y) Then ctl.BorderColor = ColorRemap(2, Y)
      If ctl.BackColor = ColorRemap(1, Y) Then ctl.BackColor = ColorRemap(2, Y)
    Next Y
    If mytheme = 40 And (TypeOf ctl Is Label Or TypeOf ctl Is CheckBox) Then ctl.ForeColor = vbWhite
  Next ctl
  If mytheme = 40 Then txtCredits.ForeColor = vbWhite
  On Error GoTo 0 '...and turn on errors afterwards ;)
  
  On Error Resume Next
  For Each ctl In Me.Controls
    If Left(ctl.Caption, 1) = "[" Then
      i = Val(Mid(ctl.Caption, 2, 4))
      ctl.Caption = LoadResString(i)
    End If
    If Left(ctl.ToolTipText, 1) = "[" Then
      i = Val(Mid(ctl.ToolTipText, 2, 4))
      ctl.ToolTipText = LoadResString(i)
    End If
  Next

  For i = 0 To 11
    imgToolBtn(i).ToolTipText = LoadResString(250 + i)
  Next i

  toolhilite.Left = -500
  
  'Check for the presence of any of the programs in the Launcher
  '--- HINT - To add another program, open the Toolbar's property
  '           pages, 18th button. It's a dropdown. Add another
  '           ButtonMenu object to it, disable it and set Key to
  '           the file name sans extension.
'  With tlbToolbar.Buttons("launch")
    For i = 0 To mnuLauncher.UBound
      If mnuLauncher(i).Caption <> "" Then
        If Dir(mnuLauncher(i).Caption & ".exe") <> "" Then
          mnuLauncher(i).Enabled = True
        End If
      End If
    Next i
'  End With
  
  '--- KAWA - Using run-time generated objects reduces elitemap.FRM's
  '---        file size from 420kb to a mere 170-something and it
  '---        still works fine =^_^=                March 11th, 2004
  For i = 1 To 63
    Load sSign(i)
    Load sTrap(i)
    Load sExit(i)
    Load sPeople(i)
  Next i
  
  'cdlCommon.InitDir = App.Path
  'pat = IIf(App.EXEName = "Project1f", "c:\vba", App.Path)
  bmap(vbLeftButton) = 0
  bmap(vbRightButton) = 1
  bmap(vbMiddleButton) = 2
  For i = 0 To &H3F
    attribcolors(i) = RGB(255, 255, 255)
  Next i
  attribcolors(0) = RGB(255, 127, 255)
  attribcolors(1) = RGB(255, 0, 0)
  attribcolors(4) = RGB(0, 0, 255)
  attribcolors(12) = RGB(0, 255, 0)
  attribcolors(13) = RGB(255, 255, 0)
  attribcolors(16) = RGB(255, 0, 255)
  attribcolors(60) = RGB(0, 255, 255)
  
  For i = 16 To 27
    j = j + Asc(Mid(lblVersion, i, 1))
  Next i
  'Coders, uncomment this for the checksum!
  'MsgBox Int(j / 2)
  
  'vsbTileset.Max = 14
  'picTileset.Picture = LoadPicture("C:\littleroot.bmp")
  'picTileset.Refresh
  For i = 0 To 63
    Select Case i
      Case 0: attribnames(i) = LoadResString(139)
      Case 4: attribnames(i) = LoadResString(140)
      Case 16: attribnames(i) = LoadResString(141)
      Case &H30: attribnames(i) = LoadResString(142)
      Case &H34: attribnames(i) = LoadResString(143)
      Case &H40: attribnames(i) = LoadResString(144)
      Case &HF0: attribnames(i) = LoadResString(145)
      Case Else: attribnames(i) = "---"
    End Select
  Next i
  selattr(0) = &H400&
  selattr(1) = &H3000&
  selattr(2) = &H1000&
  seltile(0) = &H149
  seltile(1) = 1
  seltile(2) = &H170
  For i = 0 To 15
    cboWeather.AddItem LoadResString(500 + i)
    cboType.AddItem LoadResString(600 + i)
  Next i
  For i = 0 To 8
    cboHackLanguage.AddItem LoadResString(700 + i)
  Next i
  For i = 1 To 6
    cboConnDir.AddItem LoadResString(800 + i)
  Next i
  
  If Int(j / 2) <> 442 Then
    MsgBox LoadResString(146), vbCritical, LoadResString(147)
    End
  End If
  
  drawtilehdc &H149, picSel(0).hdc, 0, 0
  drawtilehdc 1, picSel(1).hdc, 0, 0
  drawtilehdc &H170, picSel(2).hdc, 0, 0
  BitBlt picSel(0).hdc, 16, 0, 16, 16, picAttributes.hdc, &H10, 0, SRCCOPY
  BitBlt picSel(1).hdc, 16, 0, 16, 16, picAttributes.hdc, &HC0, 0, SRCCOPY
  BitBlt picSel(2).hdc, 16, 0, 16, 16, picAttributes.hdc, &H40, 0, SRCCOPY
  picSel(0).Refresh
  picSel(1).Refresh
  picSel(2).Refresh
  
  For i = 0 To picSubEditor.UBound
    picSubEditor(i).Move picSubEditor(0).Left, picSubEditor(0).Top
    picSubEditor(i).BorderStyle = 0
  Next i
  picSubEditor(0).Visible = True
  For i = 0 To picMainTab.UBound
    picMainTab(i).Move picMainTab(0).Left, picMainTab(0).Top
  Next i
  picMainTab(0).Visible = True
  Shape9.Height = picTeam.Height
  
'  'KAWA - External Overrides
'  On Error GoTo NoSongs
'  i = 0
'  Open "songs.lst" For Input As #1
'  While Not EOF(1)
'    Line Input #1, s
'    cboSong.List(i) = s
'    i = i + 1
'  Wend
'  Close #1
'
'NoSongs:

  imgBoyStart.Left = -500
  imgGirlStart.Left = -500
  
  lstPatches.ListIndex = 0

  Show
  Refresh
  DoEvents

  If Command <> "" Then
    txtRom.Text = Command
    Call LoadRom(True)
    If exit2 = True Then Exit Sub
    Call cmdGoHome_Click
    'Call cmdLoad_Click
  End If
End Sub

Private Sub Backdrop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuRMB
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If dirty = True Then
    If MsgBox(LoadResString(148), vbYesNo, LoadResString(149)) = vbNo Then Exit Sub
  End If
  INIWrite "EliteMap", "Shared", "Theme", Str(mytheme)
  If MyHeader.gGUID.part1 Then
    WriteHeader txtRom, MyHeader
  End If
End Sub

Private Sub Form_Resize()
  backdrop.Move 0, 24, ScaleWidth, ScaleHeight
End Sub

Private Sub hsbScroll_Change()
  On Error Resume Next
  Picture1.Move -(hsbScroll * &H10)
  Picture1.SetFocus
End Sub

Private Sub cmdSubmitPatch_Click()
  ShellExecute 0, vbNullString, "http://helmetedrodent.kickassgamers.com/board", vbNullString, "", 1
End Sub

Private Sub imgToolBtn_Click(Index As Integer)
  If imgToolBtn(Index).Tag = "browse" Then cmdBrowse_Click
  If imgToolBtn(Index).Tag = "save" Then cmdSave_Click
  If imgToolBtn(Index).Tag = "gohome" Then cmdGoHome_Click
  If imgToolBtn(Index).Tag = "copylevel" Then cmdCopyLevel_Click
  If imgToolBtn(Index).Tag = "copytileset" Then cmdCopyTiles_Click
  If imgToolBtn(Index).Tag = "clear" Then cmdClear_Click
  If imgToolBtn(Index).Tag = "resize" Then cmdResize_Click
  If imgToolBtn(Index).Tag = "viewscript" Then cmdLvlScript_Click
  If imgToolBtn(Index).Tag = "loadex" Then cmdLoadExtern_Click
  If imgToolBtn(Index).Tag = "saveex" Then cmdSaveExtern_Click
  If imgToolBtn(Index).Tag = "launch" Then PopupMenu mnuLaunch, , imgToolBtn(Index).Left, imgToolBtn(Index).Top + 16
  If imgToolBtn(Index).Tag = "web" Then ShellExecute 0, vbNullString, "http://helmetedrodent.kickassgamers.com", vbNullString, "", 1
  toolhilite.Left = -500
End Sub

Private Sub imgToolBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  toolhilite.Left = imgToolBtn(Index).Left - 2
End Sub


Private Sub lblTilesetLoc_DblClick()
  Dim newt1 As Long
  Dim newt2 As Long
  If MsgBox("Tileset A" & vbCrLf & _
            "Master Data: " & Hex(thismap.pTilesetA) & vbCrLf & _
            "Map16: " & Hex(thistileseta.pMap) & vbCrLf & _
            "Graphics: " & Hex(thistileseta.pGFX) & vbCrLf & _
            "Behavior: " & Hex(thistileseta.pBehavior) & vbCrLf & _
            "Animation: " & Hex(thistileseta.pAnimation) & vbCrLf & _
            " " & vbCrLf & _
            "Tileset B" & vbCrLf & _
            "Master Data: " & Hex(thismap.pTilesetB) & vbCrLf & _
            "Map16: " & Hex(thistilesetb.pMap) & vbCrLf & _
            "Graphics: " & Hex(thistilesetb.pGFX) & vbCrLf & _
            "Behavior: " & Hex(thistilesetb.pBehavior) & vbCrLf & _
            "Animation: " & Hex(thistilesetb.pAnimation) & vbCrLf & vbCrLf & _
            "Would you like to change this?" _
            , vbInformation + vbYesNo + vbDefaultButton2, "Tileset data peek") = vbYes Then
    newt1 = Val(InputBox("Enter new offset for Tileset A", , "&H" & Right("000000" & Hex(thismap.pTilesetA), 6)))
    If newt1 = 0 Then Exit Sub
    newt2 = Val(InputBox("Enter new offset for Tileset B", , "&H" & Right("000000" & Hex(thismap.pTilesetB), 6)))
    If newt2 = 0 Then Exit Sub
    thismap.pTilesetA = newt1 + &H8000000
    thismap.pTilesetB = newt2 + &H8000000
  
    lblRom.Tag = lblRom.Caption
    lblRom.Caption = LoadResString(152)
    Open txtRom For Binary As #256
    t1 = GBA2PC(thismap.pTilesetA)
    t2 = GBA2PC(thismap.pTilesetB)
    Get #256, t1 + 1, thistileseta
    Get #256, t2 + 1, thistilesetb
    Get #256, t1 + 1, checkifcompa
    Get #256, t2 + 1, checkifcompb
    Get #256, t1 + 2, checkifpala
    Get #256, t2 + 2, checkifpalb
    lblTilesetLoc = Hex(thistilesetb.pMap)
    picTileset.BackColor = 0
    CopyMemory palettesA(0, 0), blankmap(0, 0), &H400
    CopyMemory palettesB(0, 0), blankmap(0, 0), &H400
    CopyMemory gfxA(0), blankmap(0, 0), 32768
    CopyMemory gfxB(0), blankmap(0, 0), 32768
    'Different MAP16Asize in NextGen
    CopyMemory Map16A(0), blankmap(0, 0), 10240
    CopyMemory Map16B(0), blankmap(0, 0), 8192
    'picTileset.Visible = False
    ApplyTileset 0, thistileseta
    ApplyTileset 1, thistilesetb
    picTileset.Refresh
    picTileset.Visible = True
    drawtilehdc seltile(0), picSel(0).hdc, 0, 0
    drawtilehdc seltile(1), picSel(1).hdc, 0, 0
    drawtilehdc seltile(2), picSel(2).hdc, 0, 0
    picSel(0).Refresh
    picSel(1).Refresh
    picSel(2).Refresh
    For i = 0 To lheight - 1
      For j = 0 To lwidth - 1
        DrawTile TileMap(j, i), j, i
      Next j
    Next i
    Picture1.Refresh
    Close #256
    lblRom.Caption = lblRom.Tag
  End If
End Sub

Private Sub lblWorkTime_Click()
  Dim dys As Long
  Dim hrs As Long
  Dim mns As Long
  Dim p As String
  mns = MyHeader.lWorkTime
  Do
    If mns >= 60 Then
      mns = mns - 60
      hrs = hrs + 1
      If hrs = 24 Then
        hrs = 0
        dys = dys + 1
      End If
    Else
      Exit Do
    End If
  Loop
  If dys Then p = dys & IIf(dys = 1, " day", " days") & ", "
  If hrs Then p = p & hrs & IIf(hrs = 1, " hour", " hours") & ", "
  p = p & mns & IIf(mns = 1, " minute", " minutes")
  lblWorkTime = p
End Sub

Private Sub lstLabelID_Click()
  txtLabelLocX.Text = worldlocs(lstLabelID.ListIndex).bX
  txtLabelLocY.Text = worldlocs(lstLabelID.ListIndex).bY
  txtLabelLocW.Text = worldlocs(lstLabelID.ListIndex).bW
  txtLabelLocH.Text = worldlocs(lstLabelID.ListIndex).bH
  
  oldlbllen = Len(MapLabels(lstLabelID.ListIndex))
  txtLabel.Text = MapLabels(lstLabelID.ListIndex)
  
  shMap.Move (worldlocs(lstLabelID.ListIndex).bX + 1) * 8, (worldlocs(lstLabelID.ListIndex).bY + 2) * 8, worldlocs(lstLabelID.ListIndex).bW * 8, worldlocs(lstLabelID.ListIndex).bH * 8
End Sub

Private Sub lstPatches_Click()
  On Error GoTo Hell
  For i = 0 To picPatches.UBound
    picPatches(i).Visible = False
  Next i
  If lstPatches.ListIndex = lstPatches.ListCount - 1 Then
    picPatches(99).Visible = True
  Else
    picPatches(lstPatches.ListIndex).Visible = True
  End If
  lblPatchName = lstPatches.List(lstPatches.ListIndex) & " "
  shpPatchName.Left = lblPatchName.Left - 8
  shpPatchName.Width = lblPatchName.Width + 14
  'linPatchName.X1 = lblPatchName.Left - 8
  Line5.X1 = shpPatchName.Left
Hell:
  Resume Next
End Sub

Private Sub mnuBecomeItem_Click()
  If MsgBox(LoadResString(106), vbYesNo) = vbNo Then Exit Sub
  peoples(vsbPeeps).bSpriteSet = &H3B
  peoples(vsbPeeps).b3 = 0
  peoples(vsbPeeps).b9 = 0
  peoples(vsbPeeps).b10 = 0
  peoples(vsbPeeps).bBehavior1 = 0
  peoples(vsbPeeps).bBehavior2 = 0
  peoples(vsbPeeps).bIsTrainer = 0
  peoples(vsbPeeps).b14 = 0
  peoples(vsbPeeps).bTrainerLOS = 0
  peoples(vsbPeeps).b16 = 0
  peoples(vsbPeeps).iFlag = Val(InputBox("Enter unique item flag number, 0 - 255 inclusive.")) + &H400
  peoples(vsbPeeps).b23 = 0
  peoples(vsbPeeps).b24 = 0
  MsgBox "Done. Use Rubikon to make new code for this event.", vbInformation
End Sub

Private Sub mnuBecomePerson_Click()
  If MsgBox(LoadResString(106), vbYesNo) = vbNo Then Exit Sub
  peoples(vsbPeeps).b9 = 3
  peoples(vsbPeeps).bIsTrainer = 0
  peoples(vsbPeeps).bTrainerLOS = 0
  MsgBox "Done. Use Rubikon to make new dialogue code for this event", vbInformation
End Sub

Private Sub mnuBecomeTrainer_Click()
  If MsgBox(LoadResString(106), vbYesNo) = vbNo Then Exit Sub
  peoples(vsbPeeps).bIsTrainer = 1
  peoples(vsbPeeps).bTrainerLOS = Val(InputBox("Please enter the line of sight for this trainer."))
  MsgBox "Done. Use Rubikon to make new battle code for this event", vbInformation
End Sub

Private Sub mnuLauncher_Click(Index As Integer)
  X = Shell(EMPath & mnuLauncher(Index).Caption & ".exe " & txtRom, vbNormalFocus)
End Sub

Private Sub mnuRMBColors_Click()
  frmThemes.Show 1
End Sub

Private Sub mnuSetBoyStartHere_Click()
  If Roms(RomType).StartPosBoy = 0 Then
    MsgBox LoadResString(150)
    Exit Sub
  End If
  Open txtRom For Binary As #69
  Seek #69, Roms(RomType).StartPosBoy + 1
  Put #69, , CByte(cboBanks.ListIndex)
  Put #69, , CByte(cboLevels.ListIndex)
  Put #69, , CByte(&HFF)
  Put #69, , CInt("&H" & Mid(Label2, 4, 2))
  Put #69, , CInt("&H" & Mid(Label2, InStr(Label2, "Y:") + 3))
  Close #69
  imgBoyStart.Left = CInt("&H" & Mid(Label2, 4, 2)) * 16
  imgBoyStart.Top = CInt("&H" & Mid(Label2, InStr(Label2, "Y:") + 3)) * 16
End Sub

Private Sub mnuSetGirlStartHere_Click()
  If Roms(RomType).StartPosGirl = 0 Then
    MsgBox LoadResString(151)
    Exit Sub
  End If
  Open txtRom For Binary As #69
  Seek #69, Roms(RomType).StartPosGirl + 1
  Put #69, , CByte(cboBanks.ListIndex)
  Put #69, , CByte(cboLevels.ListIndex)
  Put #69, , CByte(&HFF)
  Debug.Print CInt("&H" & Mid(Label2, 4, 2))
  Debug.Print CInt("&H" & Mid(Label2, InStr(Label2, "Y:") + 3))
  Put #69, , CInt("&H" & Mid(Label2, 4, 2))
  Put #69, , CInt("&H" & Mid(Label2, InStr(Label2, "Y:") + 3))
  Close #69
  imgGirlStart.Left = CInt("&H" & Mid(Label2, 4, 2)) * 16
  imgGirlStart.Top = CInt("&H" & Mid(Label2, InStr(Label2, "Y:") + 3)) * 16
End Sub

Private Sub mnuWorldMapChange_Click()
  Dim cc As cCommonDialog
  Set cc = New cCommonDialog
  Dim heyhey As String
  heyhey = Roms(RomType).WorldMap
  If Not cc.VBGetOpenFileName(heyhey, , , , , , "Bitmaps (*.bmp)|*.bmp|GIFs (*.gif)|*.gif|All (*.*)|*.*") Then Exit Sub
  picWorldMap.Picture = LoadPicture(heyhey)
End Sub

Private Sub picBorder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mx = (X \ 16)
  my = Y \ 16
  If Shift = 1 Then GoTo shift_event2:
  If X > &H0 And Y > &H0 And X < (&H20) And Y < (&H20) Then
  dirty2 = True
    If Button = vbLeftButton Then
      border(my, mx) = seltile(0)
      drawtilehdc seltile(0), picBorder.hdc, mx, my
      picBorder.Refresh
    ElseIf Button = vbRightButton Then
      border(my, mx) = seltile(1)
      drawtilehdc seltile(1), picBorder.hdc, mx, my
      picBorder.Refresh
    ElseIf Button = vbMiddleButton Then
      border(my, mx) = seltile(2)
      drawtilehdc seltile(2), picBorder.hdc, mx, my
      picBorder.Refresh
    End If
  End If
Exit Sub
shift_event2:
  mx = (X) \ 16
  my = (Y) \ 16
  If X > &H0 And Y > &H0 And X < (&H20) And Y < (&H20) Then
  If Button = vbLeftButton Then
      seltile(0) = border(my, mx) Mod &H400
      Form1.drawtilehdc seltile(0), Form1.picSel(0).hdc, 0, 0
      Form1.picSel(0).Refresh
    ElseIf Button = vbRightButton Then
      seltile(1) = border(my, mx) Mod &H400
      Form1.drawtilehdc seltile(1), Form1.picSel(1).hdc, 0, 0
      Form1.picSel(1).Refresh
    ElseIf Button = vbMiddleButton Then
      seltile(2) = border(my, mx) Mod &H400
      Form1.drawtilehdc seltile(2), Form1.picSel(2).hdc, 0, 0
      Form1.picSel(2).Refresh
    End If
  End If
End Sub

Private Sub picSprite_Paint()
  BitBlt picSprite.hdc, 0, 0, 32, 64, picSpritestrip.hdc, vsbPeepSprite.Value * 32, 0, vbSrcCopy
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Picture1_MouseMove Button, Shift, X, Y
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim y_long As Long
  Dim zwisp As Variant
  mx = X \ 16
  my = Y \ 16
  If X > &H0 And Y > &H0 And X < ((lwidth * &H10)) And Y < ((lheight * &H10)) Then
  Else
    Exit Sub
  End If
  
  If Shift = 1 And Button = 2 Then
    'OMG RMB MENU!
    PopupMenu mnuMapRMB
  End If
  
  Label2 = "X: " & Hex(mx) & " Y: " & Hex(my)
  Label3 = Hex(TileMap((mx), (my)) Mod &H400)
  Label9 = Hex(TileMap((mx), (my)) \ &H400) & ":" & attribnames((TileMap(mx, my) \ &H400) * 4)
  
  at = TileMap(mx, my) \ &H400
  
  Shape1.Move mx * 16, my * 16, 16, 16
  If chkNoDraw.Value = 0 Then
    Shape1.BorderColor = attribcolors(at)
  Else
    Shape1.BorderColor = 0
  End If
  Shape1.Visible = True
  
  If Shift = 0 Then GoTo Pencil
  If Shift = 1 Then GoTo Dropper
  If Shift = 3 Then GoTo Stamp
  'If chkUseStamp.value = 1 Then GoTo Stamp

Pencil:
  Picture1.MouseIcon = LoadResPicture(100, vbResCursor)
  If X > &H0 And Y > &H0 And X < ((lwidth * &H10)) And Y < ((lheight * &H10)) Then
    If chkNoDraw.Value = 0 Then
      If Button = vbLeftButton Then
        oldtile = seltile(0)
        If chkAttribsOnly.Value = 1 Then seltile(0) = TileMap((mx), (my)) Mod &H400
        TileMap((mx), (my)) = seltile(0) + selattr(0)
        DrawTile seltile(0), mx, my
        Picture1.Refresh
        dirty = True
        seltile(0) = oldtile
      ElseIf Button = vbRightButton Then
        oldtile = seltile(1)
        If chkAttribsOnly.Value = 1 Then seltile(1) = TileMap((mx), (my)) Mod &H400
        TileMap((mx), (my)) = seltile(1) + selattr(1)
        DrawTile seltile(1), mx, my
        Picture1.Refresh
        dirty = True
        seltile(1) = oldtile
      ElseIf Button = vbMiddleButton Then
        oldtile = seltile(2)
        If chkAttribsOnly.Value = 1 Then seltile(2) = TileMap((mx), (my)) Mod &H400
        TileMap((mx), (my)) = seltile(2) + selattr(2)
        DrawTile seltile(2), mx, my
        Picture1.Refresh
        dirty = True
        seltile(2) = oldtile
      End If
    End If
  End If
  Exit Sub
Dropper:
  Picture1.MouseIcon = LoadResPicture(101, vbResCursor)
  If X > &H0 And Y > &H0 And X < ((lwidth * &H10)) And Y < ((lheight * &H10)) Then
    If Button = vbLeftButton Then
      seltile(0) = TileMap((mx), (my)) Mod &H400
      a = Split((TileMap((mx), (my)) / &H400), ",", , vbTextCompare)
      selattr(0) = a(0) * &H400
      drawtilehdc seltile(0), picSel(0).hdc, 0, 0
      BitBlt picSel(0).hdc, 16, 0, 16, 16, picAttributes.hdc, ((selattr(0) / &H400) And &HF) * 16, (((selattr(0) / &H400 - 1) And &HF0) / &H10) * 16, SRCCOPY
      picSel(0).Refresh
    ElseIf Button = vbRightButton Then
      seltile(1) = TileMap((mx), (my)) Mod &H400
      a = Split((TileMap((mx), (my)) / &H400), ",", , vbTextCompare)
      selattr(1) = a(0) * &H400
      drawtilehdc seltile(1), picSel(1).hdc, 0, 0
      BitBlt picSel(1).hdc, 16, 0, 16, 16, picAttributes.hdc, ((selattr(1) / &H400) And &HF) * 16, (((selattr(1) / &H400 - 1) And &HF0) / &H10) * 16, SRCCOPY
      picSel(1).Refresh
    ElseIf Button = vbMiddleButton Then
      seltile(2) = TileMap((mx), (my)) Mod &H400
      a = Split((TileMap((mx), (my)) / &H400), ",", , vbTextCompare)
      selattr(2) = a(0) * &H400
      drawtilehdc seltile(2), picSel(2).hdc, 0, 0
      BitBlt picSel(2).hdc, 16, 0, 16, 16, picAttributes.hdc, ((selattr(2) / &H400) And &HF) * 16, (((selattr(2) / &H400 - 1) And &HF0) / &H10) * 16, SRCCOPY
      picSel(2).Refresh
    End If
  End If
  Exit Sub
Stamp:
  Picture1.MouseIcon = LoadResPicture(102, vbResCursor)
  If X > &H0 And Y > &H0 And X < (((lwidth - 1) * &H10)) And Y < (((lheight - 1) * &H10)) Then
    If Button = vbLeftButton Then
      TileMap((mx), (my)) = StampMap(0, 0) + selattr(0)
      TileMap((mx + 1), (my)) = StampMap(0, 1) + selattr(0)
      TileMap((mx), (my + 1)) = StampMap(1, 0) + selattr(0)
      TileMap((mx + 1), (my + 1)) = StampMap(1, 1) + selattr(0)
      DrawTile StampMap(0, 0), mx, my
      DrawTile StampMap(0, 1), mx + 1, my
      DrawTile StampMap(1, 0), mx, my + 1
      DrawTile StampMap(1, 1), mx + 1, my + 1
      Label3 = Hex(TileMap((mx), (my)))
      Picture1.Refresh
      dirty = True
    End If
    Shape1.Move mx * 16, my * 16, 32, 32
  End If
End Sub

Private Sub picTileset_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mx = (X \ 16)
  my = Y \ 16
  'Caption = Shift
  If X > 0 And Y > 0 And X < (&H1000) And Y < (&H4000) Then
    If Button = vbLeftButton Then
      seltile(0) = (my * CLng(&H10)) + mx
      drawtilehdc (my * CLng(&H10)) + mx, picSel(0).hdc, 0, 0
      picSel(0).Refresh
      'picAttributes.Visible = True
    ElseIf Button = vbRightButton Then
      seltile(1) = (my * CLng(&H10)) + mx
      drawtilehdc (my * CLng(&H10)) + mx, picSel(1).hdc, 0, 0
      picSel(1).Refresh
      'picAttributes.Visible = True
    ElseIf Button = vbMiddleButton Then
      seltile(2) = (my * CLng(&H10)) + mx
      drawtilehdc (my * CLng(&H10)) + mx, picSel(2).hdc, 0, 0
      picSel(2).Refresh
      'picAttributes.Visible = True
    End If
    If Shift = 2 Then
      Load frmEditTile
      With frmEditTile
        lblRom.Tag = lblRom.Caption
        tte = (my * CLng(&H10)) + mx
        If tte >= &H200 Then
          .TileAddress = GBA2PC(thistilesetb.pMap) + ((tte - &H200) * 16)
          .BehaviorAddress = GBA2PC(thistilesetb.pBehavior) + ((tte - &H200) * 2)
        Else
          .TileAddress = GBA2PC(thistileseta.pMap) + (tte * 16)
          .BehaviorAddress = GBA2PC(thistileseta.pBehavior) + (tte * 2)
        End If
        'MsgBox Hex(.TileAddress)
        .Filename = txtRom
        .LoadTile
        .Show 1
        'cmdLoad_Click
        If lblRom.Caption = "cancelled" Then
          lblRom.Caption = lblRom.Tag
          Exit Sub
        End If
        lblRom.Caption = LoadResString(152)
        Open txtRom For Binary As #256
        If tte >= &H200 Then
          t1 = GBA2PC(thismap.pTilesetB)
          Get #256, t1 + 1, thistilesetb
          CopyMemory Map16B(0), blankmap(0, 0), 10240
          ApplyTileset 1, thistilesetb
        Else
          t1 = GBA2PC(thismap.pTilesetA)
          Get #256, t1 + 1, thistileseta
          CopyMemory Map16A(0), blankmap(0, 0), 10240
          ApplyTileset 0, thistileseta
        End If
        picTileset.Refresh
        For i = 0 To lheight - 1
          For j = 0 To lwidth - 1
            DrawTile TileMap(j, i), j, i
          Next j
        Next i
        Picture1.Refresh
        Close #256
        lblRom.Caption = lblRom.Tag
      End With
    End If
  End If
End Sub

Private Sub picTileset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mx = (X \ 16)
  my = Y \ 16
  If mx = 16 Then mx = 15
  If X > 0 And Y > 0 And X < (&H1000) And Y < (&H4000) Then
    Label4 = Hex((my * CLng(&H10)) + mx)
  End If
  shpTileset.Move mx * 16, my * 16
  shpTileset.Visible = True
End Sub

Private Sub picAttributes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mx = (X \ 16)
  my = Y \ 16
  
  If X > 0 And Y > 0 And X < &H100 And Y < (&H40) Then
    If Button = vbLeftButton Then
      tattr(0) = "&H" & Hex((my * &H10 + mx))
      BitBlt picSel(0).hdc, 16, 0, 16, 16, picAttributes.hdc, mx * 16, my * 16, SRCCOPY
      picSel(0).Refresh
      'picAttributes.Visible = False
    ElseIf Button = vbRightButton Then
      tattr(1) = "&H" & Hex(my * &H10 + mx)
      BitBlt picSel(1).hdc, 16, 0, 16, 16, picAttributes.hdc, mx * 16, my * 16, SRCCOPY
      picSel(1).Refresh
      'picAttributes.Visible = False
    ElseIf Button = vbMiddleButton Then
      tattr(2) = "&H" & Hex(my * &H10 + mx)
      BitBlt picSel(2).hdc, 16, 0, 16, 16, picAttributes.hdc, mx * 16, my * 16, SRCCOPY
      picSel(2).Refresh
      'picAttributes.Visible = False
    End If
    shpAttributes.Move mx * 16, my * 16
    shpAttributes.BorderColor = attribcolors((my * &H10) + mx)
    shpAttributes.Visible = True
  End If
End Sub

'Private Sub pTiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  mx = (X \ 16)
'  my = Y \ 16
'  If X > 0 And Y > 0 And X < (&H1000) And Y < (&H1000) Then
'    Label4 = "X: " & Hex(mx) & " Y: " & Hex(my)
'  End If
'End Sub

Private Sub drawborder(ByVal lw As Byte, ByVal lh As Byte)
  For i = -2 To lh + 1
    For j = -2 To lw + 1
      DrawTile border(Abs(i Mod 2), Abs(j Mod 2)), j, i
    Next j
  Next i
End Sub


Private Sub drawstamp(ByVal lw As Byte, ByVal lh As Byte)
  For i = -2 To lh + 1
    For j = -2 To lw + 1
      DrawTile border(Abs(i Mod 2), Abs(j Mod 2)), j, i
    Next j
  Next i
End Sub

Private Sub picAttributes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mx = (X \ 16)
  my = Y \ 16
  
  If X > 0 And Y > 0 And X < &H100 And Y < (&H40) Then
    Label45 = Hex((my * &H10 + mx)) & " " & attribnames((my * &H10 + mx) * 4)
    shpAttributes.Move mx * 16, my * 16
    shpAttributes.BorderColor = attribcolors((my * &H10) + mx)
    shpAttributes.Visible = True
  End If
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mx = (X \ 16)
  my = Y \ 16
  If X > &H0 And Y > &H0 And X < (&H20) And Y < (&H20) Then
  dirty2 = True
    If Button = vbLeftButton Then
      StampMap(my, mx) = seltile(0)
      drawtilehdc seltile(0), Picture3.hdc, mx, my
      Picture3.Refresh
    ElseIf Button = vbRightButton Then
      StampMap(my, mx) = seltile(1)
      drawtilehdc seltile(1), Picture3.hdc, mx, my
      Picture3.Refresh
    ElseIf Button = vbMiddleButton Then
      StampMap(my, mx) = seltile(2)
      drawtilehdc seltile(2), Picture3.hdc, mx, my
      Picture3.Refresh
    End If
  End If
End Sub

Private Sub picWorldMap_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If picWorldMap.Tag = "yiffykittycat" Then
      If MsgBox(LoadResString(153), vbYesNo) = vbYes Then
        Open txtRom For Binary As #256
          Put #256, 1, CByte(&H31)
          Put #256, &HCF, CByte(&HA0)
          Put #256, &HD0, CByte(&HE3)
        Close #256
        MsgBox "ROM locked."
      End If
    ElseIf picWorldMap.Tag = "trashcan" Then
      Load frmSecret
      frmSecret.imgSecret.Picture = LoadResPicture(4, 0)
      Randomize Timer
      Select Case Int(Rnd * 4)
        Case 1: frmSecret.lblSecret = "BRAAAAAINS!" & vbCrLf & vbCrLf & "And love!"
        Case 2: frmSecret.lblSecret = "I'm not mad at you, babe. I just look that way! Honest!"
        Case 3: frmSecret.lblSecret = "Are you blind, girl? I'm behind you, not in -front- of you!"
        Case Else: frmSecret.lblSecret = "Catch me a fish, you stupid sexy catgirl! Don't make me use my pimp hand!"
      End Select
      frmSecret.Show 1
    ElseIf picWorldMap.Tag = "thecolorpurple" Then
      Load frmSecret
      frmSecret.imgSecret.Visible = False
      frmSecret.lblSecret.Visible = False
      frmSecret.ScaleMode = 0
      frmSecret.ScaleWidth = 16
      frmSecret.ScaleHeight = 16
      Dim X As Integer
      Dim Y As Integer
      For X = 0 To 15
        For Y = 0 To 15
          frmSecret.Line (X, Y)-(X + 1, Y + 1), palettesA(Y, X), BF
        Next Y
      Next X
      frmSecret.CurrentX = 1
      frmSecret.CurrentY = 11
      frmSecret.ForeColor = vbWhite
      frmSecret.FontSize = 16
      frmSecret.Print "W00t! Tileset pal!"
      frmSecret.Show 1
    End If
    picWorldMap.Tag = ""
  Else
    picWorldMap.Tag = picWorldMap.Tag + Chr(KeyAscii)
  End If
  'Caption = picWorldMap.Tag
End Sub

Private Sub picWorldMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuWorldMap
End Sub

Private Sub sExit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = vbCtrlMask And Button = vbLeftButton Then
    If exits(Index).hLevel = &H7F7F& Then
    write_bank_lev lastbank, lastlev
    Else
    write_bank_lev exits(Index).hLevel \ 256, exits(Index).hLevel Mod 256
      lastbank = exits(Index).hLevel \ 256
      lastlev = exits(Index).hLevel Mod 256
    End If
  End If
End Sub

Private Sub sExit_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo Hell
  If Shift <> vbCtrlMask Then
    Picture1_MouseMove 0, 0, sExit(Index).Left, sExit(Index).Top
    If Button = vbLeftButton Then
      sExit(Index).Move Int((sExit(Index).Left + (X \ Screen.TwipsPerPixelX) - 8) / 16) * 16, Int((sExit(Index).Top + (Y \ Screen.TwipsPerPixelY) - 8) / 16) * 16
      If sExit(Index).Left < 0 Then sExit(Index).Left = 0 'KAWA -- Quick fix.
      If sExit(Index).Top < 0 Then sExit(Index).Top = 0 'About time too >_<
      exits(Index).bX = sExit(Index).Left \ 16
      exits(Index).bY = sExit(Index).Top \ 16
      dirty = True
    End If
  End If
  m = "[Exit " & Hex(Index) & "]" & vbCrLf
  m = m & "X: " & Hex(exits(Index).bX) & vbCrLf
  m = m & "Y: " & Hex(exits(Index).bY) & vbCrLf
  m = m & "Level: B:" & Hex(exits(Index).hLevel \ 256) & " L:" & Hex(exits(Index).hLevel Mod 256) & vbCrLf
  m = m & "Exit: " & Hex(exits(Index).b6)
  m = m & vbCrLf & "[Ctrl+Click to Follow Exit]" & vbCrLf
  lblInfo = m
Hell:
End Sub

Private Sub sExit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo Hell
  If Shift <> vbCtrlMask Then
    If Button = vbLeftButton Then
      sExit(Index).Move Int((sExit(Index).Left + (X \ Screen.TwipsPerPixelX) - 8) / 16) * 16, Int((sExit(Index).Top + (Y \ Screen.TwipsPerPixelY) - 8) / 16) * 16
      exits(Index).bX = sExit(Index).Left \ 16
      exits(Index).bY = sExit(Index).Top \ 16
      dirty = True
    End If
  End If
Hell:
End Sub

Private Sub sPeople_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = vbCtrlMask And Button = vbLeftButton Then CallScriptEd peoples(Index).pScript
'    If Dir("scripted.exe") <> "" Then
'      scriptad = peoples(Index).pScript
'      If scriptad > -1 Then
'        Open "scripted.dat" For Output As #5
'        Print #5, txtRom
'        Print #5, scriptad - &H8000000
'        Print #5, "people"
'        Close #5
'        Shell "scripted.exe 1", vbNormalFocus
'      End If
'    Else: MsgBox "ScriptEd.exe not found. Cannot show script.": End If
'  End If
End Sub

Private Sub sPeople_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo Hell
  If Shift <> vbCtrlMask Then
    Picture1_MouseMove 0, 0, sPeople(Index).Left, sPeople(Index).Top
    If Button = vbLeftButton Then
      sPeople(Index).Move Int((sPeople(Index).Left + (X \ Screen.TwipsPerPixelX) - 8) / 16) * 16, Int((sPeople(Index).Top + (Y \ Screen.TwipsPerPixelY) - 8) / 16) * 16
      If sPeople(Index).Left < 0 Then sPeople(Index).Left = 0 'KAWA -- Quick fix.
      If sPeople(Index).Top < 0 Then sPeople(Index).Top = 0 'About time too >_<
      peoples(Index).bX = sPeople(Index).Left \ 16
      peoples(Index).bY = sPeople(Index).Top \ 16
      dirty = True
    End If
  End If
  
'  b1 As Byte
'  bSpriteSet As Byte
'  b3 As Byte
'  b4 As Byte
'  bX As Byte
'  b6 As Byte
'  bY As Byte
'  b8 As Byte
'  b9 As Byte
'  b10 As Byte
'  bBehavior1 As Byte
'  bBehavior2 As Byte
'  b13 As Byte
'  b14 As Byte
'  b15 As Byte
'  b16 As Byte
'  pScript As Long
'  iFlag As Integer
'  bValue As Byte
'  b23 As Byte
'  b24 As Byte
  m = m & "Index: " & Hex(peoples(Index).b1) & vbCrLf
  m = m & "Sprite: " & Hex(peoples(Index).bSpriteSet) & vbCrLf
  m = m & "B3: " & Hex(peoples(Index).b3) & vbCrLf
  m = m & "X: " & Hex(peoples(Index).bX) & vbCrLf
  m = m & "Y: " & Hex(peoples(Index).bY) & vbCrLf
  m = m & "B9: " & Hex(peoples(Index).b9) & vbCrLf
  m = m & "B10: " & Hex(peoples(Index).b10) & vbCrLf
  m = m & "Behave1: " & Hex(peoples(Index).bBehavior1) & vbCrLf
  m = m & "Behave2: " & Hex(peoples(Index).bBehavior2) & vbCrLf
  m = m & "IsTrainer: " & Hex(peoples(Index).bIsTrainer) & vbCrLf
  m = m & "B14: " & Hex(peoples(Index).b14) & vbCrLf
  m = m & "LOS: " & Hex(peoples(Index).bTrainerLOS) & vbCrLf
  m = m & "B16: " & Hex(peoples(Index).b16) & vbCrLf
  m = m & "Script: " & Hex(GBA2PC(peoples(Index).pScript)) & vbCrLf
  m = m & "Flag: " & Hex(peoples(Index).iFlag) & vbCrLf
  m = m & "B23: " & Hex(peoples(Index).b23) & vbCrLf
  m = m & "B24: " & Hex(peoples(Index).b24) & vbCrLf
  m = m & vbCrLf & "[Ctrl+Click to View Script]" & vbCrLf
  lblInfo = m
Hell:
End Sub

Private Sub sPeople_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift <> vbCtrlMask Then
    If Button = vbLeftButton Then
      sPeople(Index).Move Int((sPeople(Index).Left + (X \ Screen.TwipsPerPixelX) - 8) / 16) * 16, Int((sPeople(Index).Top + (Y \ Screen.TwipsPerPixelY) - 8) / 16) * 16
      peoples(Index).bX = sPeople(Index).Left \ 16
      peoples(Index).bY = sPeople(Index).Top \ 16
    End If
  End If
End Sub

Private Sub sTrap_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = vbCtrlMask And Button = vbLeftButton Then CallScriptEd traps(Index).pScript
'    If Dir("scripted.exe") <> "" Then
'      scriptad = traps(Index).pScript
'      If scriptad > -1 Then
'        Open "scripted.dat" For Output As #5
'        Print #5, txtRom
'        Print #5, scriptad - &H8000000
'        Print #5, "trap"
'        Close #5
'        Shell App.Path & "\scripted.exe 1", vbNormalFocus
'      End If
'    Else: MsgBox "ScriptEd.exe not found. Cannot show script.": End If
'  End If
End Sub

Private Sub sTrap_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error GoTo Hell
  If Shift <> vbCtrlMask Then
    Picture1_MouseMove 0, 0, sTrap(Index).Left, sTrap(Index).Top
    If Button = vbLeftButton Then
      sTrap(Index).Move Int((sTrap(Index).Left + (X \ Screen.TwipsPerPixelX) - 8) / 16) * 16, Int((sTrap(Index).Top + (Y \ Screen.TwipsPerPixelY) - 8) / 16) * 16
      If sTrap(Index).Left < 0 Then sTrap(Index).Left = 0 'KAWA -- Quick fix.
      If sTrap(Index).Top < 0 Then sTrap(Index).Top = 0 'About time too >_<
      traps(Index).bX = sTrap(Index).Left \ 16
      traps(Index).bY = sTrap(Index).Top \ 16
      dirty = True
    End If
  End If
  m = "[Trigger " & Hex(Index) & "]" & vbCrLf
  m = m & "X: " & Hex(traps(Index).bX) & vbCrLf
  m = m & "Y: " & Hex(traps(Index).bY) & vbCrLf
  m = m & "Check: " & Hex(traps(Index).hFlagCheck) & vbCrLf
  m = m & "Value: " & Hex(traps(Index).hFlagValue) & vbCrLf
  m = m & "Script: " & Hex(GBA2PC(traps(Index).pScript)) & vbCrLf
  m = m & vbCrLf & "[Ctrl+Click to View Script]" & vbCrLf
  lblInfo = m
Hell:
End Sub

Private Sub sTrap_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift <> vbCtrlMask Then
    If Button = vbLeftButton Then
      sTrap(Index).Move Int((sTrap(Index).Left + (X \ Screen.TwipsPerPixelX) - 8) / 16) * 16, Int((sTrap(Index).Top + (Y \ Screen.TwipsPerPixelY) - 8) / 16) * 16
      traps(Index).bX = sTrap(Index).Left \ 16
      traps(Index).bY = sTrap(Index).Top \ 16
      dirty = True
    End If
  End If
End Sub

Private Sub CallScriptEd(offset As Long)
  Dim mycc As New cCommonDialog
  If Dir(EMPath & "scripted.exe") <> "" Then
    scriptad = signs(Index).pScript
    If scriptad > -1 Then
      Shell "scripted.exe " & mycc.VBGetFileTitle(txtRom) & ":" & Hex(offset - &H8000000), vbNormalFocus
    End If
  Else
    MsgBox LoadResString(154)
  End If
End Sub

Private Sub sSign_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = vbCtrlMask And Button = vbLeftButton Then CallScriptEd signs(Index).pScript
End Sub

Private Sub sSign_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift <> vbCtrlMask Then
    Picture1_MouseMove 0, 0, sSign(Index).Left, sSign(Index).Top
    If Button = vbLeftButton Then
      sSign(Index).Move Int((sSign(Index).Left + (X \ Screen.TwipsPerPixelX) - 8) / 16) * 16, Int((sSign(Index).Top + (Y \ Screen.TwipsPerPixelY) - 8) / 16) * 16
      If sSign(Index).Left < 0 Then sSign(Index).Left = 0 'KAWA -- Quick fix.
      If sSign(Index).Top < 0 Then sSign(Index).Top = 0 'About time too >_<
      signs(Index).bX = sSign(Index).Left \ 16
      signs(Index).bY = sSign(Index).Top \ 16
      dirty = True
    End If
  End If
  m = "[Sign " & Hex(Index) & "]" & vbCrLf
  m = m & "X: " & Hex(signs(Index).bX) & vbCrLf
  m = m & "Y: " & Hex(signs(Index).bY) & vbCrLf
  m = m & "Script: " & Hex(GBA2PC(signs(Index).pScript)) & vbCrLf
  m = m & vbCrLf & "[Ctrl+Click to View Script]" & vbCrLf
  lblInfo = m
End Sub

Private Sub sSign_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift <> vbCtrlMask Then
    If Button = vbLeftButton Then
      sSign(Index).Move Int((sSign(Index).Left + (X \ Screen.TwipsPerPixelX) - 8) / 16) * 16, Int((sSign(Index).Top + (Y \ Screen.TwipsPerPixelY) - 8) / 16) * 16
      signs(Index).bX = sSign(Index).Left \ 16
      signs(Index).bY = sSign(Index).Top \ 16
      dirty = True
    End If
  End If
End Sub



Private Sub subtab_Click(Index As Integer)
  For i = 0 To picSubEditor.UBound
    picSubEditor(i).Visible = False
  Next i
  picSubEditor(Index).Visible = True
End Sub

Private Sub tattr_Change(Index As Integer)
  selattr(Index) = CLng(Val(tattr(Index))) * CLng(&H400)
End Sub

Private Sub timThrobber_Timer()
  picThrobber.PaintPicture picThrobberPics.Picture, 0, 0, 105, 17, 0, Val(timThrobber.Tag) * 17, 105, 17
  'BitBlt picThrobber.hdc, 0, 0, 105, 17, picThrobberPics.hdc, 0, Val(timThrobber.Tag) * 17, vbSrcCopy
  If picThrobber.Tag = "" Then
    timThrobber.Tag = Val(timThrobber.Tag) + 1
    If timThrobber.Tag = "10" Then
      timThrobber.Tag = "0"
      'picThrobber.Tag = "hammertime"
    End If
  End If
End Sub

Private Sub timWorkTimer_Timer()
  MyHeader.lWorkTime = MyHeader.lWorkTime + 1
  lblWorkTime_Click
End Sub

Private Sub toolbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuRMB
End Sub

Private Sub toolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  toolhilite.Left = -5000
End Sub

Private Sub txtAuthorName_LostFocus()
  MyHeader.sAuthor = txtAuthorName
End Sub

Private Sub txtCredits_Click()
  txtCredits.SelStart = 0
  txtCredits.SelLength = 0
  picTeam.SetFocus
End Sub

Private Sub txtCredits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtCredits_Click
End Sub

Private Sub txtExitLevel_LostFocus()
  txtExitLevel.Text = "&H" & Right("0000" & Hex(Val(txtExitLevel)), 4)
  exits(vsbExits).hLevel = Val(txtExitLevel)
  dirty = True
End Sub

Private Sub txtExitTarget_Change()
  txtExitTarget.Text = "&H" & Right("00" & Hex(Val(txtExitTarget)), 2)
  exits(vsbExits).b6 = Val(txtExitTarget)
  dirty = True
End Sub

Private Sub txtGroupName_LostFocus()
  MyHeader.sGroup = txtGroupName
End Sub

Private Sub txtHackName_LostFocus()
  MyHeader.sName = txtHackName
End Sub

Private Sub txtLevelHeight_LostFocus()
  txtLevelHeight.Text = "&H" & Right("00" & Hex(Val(txtLevelHeight)), 2)
  'KAWA -- Now that the resizer has been overhauled, I'll just haveta make this instantaneous...
  thismap.wHeight = Val(txtLevelHeight)
  lheight = Val(txtLevelHeight)
  refreshlevel
  dirty = True
End Sub

Private Sub txtLevelWidth_LostFocus()
  txtLevelWidth.Text = "&H" & Right("00" & Hex(Val(txtLevelWidth)), 2)
  'KAWA -- Now that the resizer has been overhauled, I'll just haveta make this instantaneous...
  thismap.wWidth = Val(txtLevelWidth)
  lwidth = Val(txtLevelWidth)
  refreshlevel
  dirty = True
End Sub

Private Sub txtOldskoolChooser_LostFocus()
  txtOldskoolChooser = Right("0000" & Hex(Val("&H" & txtOldskoolChooser)), 4)
  write_bank_lev Val("&H" & txtOldskoolChooser) \ 256, Val("&H" & txtOldskoolChooser) Mod 256
End Sub

Private Sub txtPeepFlag_LostFocus()
  txtPeepFlag.Text = "&H" & Right("0000" & Hex(Val(txtPeepFlag)), 4)
  peoples(vsbPeeps).iFlag = Val(txtPeepFlag)
  dirty = True
End Sub

Private Sub txtPeepLOS_Change()
  txtPeepLOS.Text = "&H" & Right("00" & Hex(Val(txtPeepLOS)), 2)
  peoples(vsbPeeps).bTrainerLOS = Val(txtPeepLOS)
  dirty = True
End Sub

Private Sub txtPeepScript_LostFocus()
  txtPeepScript.Text = "&H" & Right("000000" & Hex(Val(txtPeepScript)), 6)
  peoples(vsbPeeps).pScript = Val(txtPeepScript) + &H8000000
  dirty = True
End Sub

Private Sub txtPeepSprite_LostFocus()
  txtPeepSprite.Text = "&H" & Right("00" & Hex(Val(txtPeepSprite)), 2)
  peoples(vsbPeeps).bSpriteSet = Val(txtPeepSprite)
  vsbPeepSprite.Value = Val(txtPeepSprite)
  dirty = True
End Sub

Private Sub txtConnLevel_LostFocus()
  txtConnLevel.Text = "&H" & Right("0000" & Hex(Val(txtConnLevel)), 4)
  mapConnects(vsbConn.Value).hLevel = Val(txtConnLevel.Text)
  renderconnects
  dirty = True
End Sub

Private Sub txtConnOffset_LostFocus()
  'txtConnOffset.Text = "&H" & Right("0000" & Hex(Val(txtConnOffset)), 4)
  mapConnects(vsbConn.Value).wOffset = Val(txtConnOffset.Text)
  renderconnects
  dirty = True
End Sub

Private Sub txtLabel_Change()
  MapLabels(lstLabelID.ListIndex) = txtLabel.Text
  cboLabelID.List(lstLabelID.ListIndex) = Right("00" & Hex(lstLabelID.ListIndex), 2) & ". " & MapLabels(lstLabelID.ListIndex)
  lstLabelID.List(lstLabelID.ListIndex) = Right("00" & Hex(lstLabelID.ListIndex), 2) & ". " & MapLabels(lstLabelID.ListIndex)
  
  If Len(txtLabel) > Val(oldlbllen) Then
    txtLabel.ForeColor = vbRed
  Else
    txtLabel.ForeColor = 0
  End If
End Sub

Private Sub txtLabelLocH_Change()
  i = Val(txtLabelLocH)
  worldlocs(lstLabelID.ListIndex).bH = i
  shMap.Move (worldlocs(lstLabelID.ListIndex).bX + 1) * 8, (worldlocs(lstLabelID.ListIndex).bY + 2) * 8, worldlocs(lstLabelID.ListIndex).bW * 8, worldlocs(lstLabelID.ListIndex).bH * 8
End Sub

Private Sub txtLabelLocW_Change()
  i = Val(txtLabelLocW)
  worldlocs(lstLabelID.ListIndex).bW = i
  shMap.Move (worldlocs(lstLabelID.ListIndex).bX + 1) * 8, (worldlocs(lstLabelID.ListIndex).bY + 2) * 8, worldlocs(lstLabelID.ListIndex).bW * 8, worldlocs(lstLabelID.ListIndex).bH * 8
End Sub

Private Sub txtLabelLocX_Change()
  i = Val(txtLabelLocX)
  worldlocs(lstLabelID.ListIndex).bX = i
  shMap.Move (worldlocs(lstLabelID.ListIndex).bX + 1) * 8, (worldlocs(lstLabelID.ListIndex).bY + 2) * 8, worldlocs(lstLabelID.ListIndex).bW * 8, worldlocs(lstLabelID.ListIndex).bH * 8
End Sub

Private Sub txtLabelLocY_Change()
  i = Val(txtLabelLocY)
  worldlocs(lstLabelID.ListIndex).bY = i
  shMap.Move (worldlocs(lstLabelID.ListIndex).bX + 1) * 8, (worldlocs(lstLabelID.ListIndex).bY + 2) * 8, worldlocs(lstLabelID.ListIndex).bW * 8, worldlocs(lstLabelID.ListIndex).bH * 8
End Sub

Private Sub txtPeepX_LostFocus()
  txtPeepX.Text = "&H" & Right("00" & Hex(Val(txtPeepX)), 2)
  peoples(vsbPeeps).bX = Val(txtPeepX)
  rendersprites
  dirty = True
End Sub

Private Sub txtPeepY_LostFocus()
  txtPeepY.Text = "&H" & Right("00" & Hex(Val(txtPeepY)), 2)
  peoples(vsbPeeps).bY = Val(txtPeepY)
  rendersprites
  dirty = True
End Sub

Private Sub txtExitX_LostFocus()
  txtExitX.Text = "&H" & Right("00" & Hex(Val(txtExitX)), 2)
  exits(vsbExits).bX = Val(txtExitX)
  rendersprites
  dirty = True
End Sub

Private Sub txtExitY_LostFocus()
  txtExitY.Text = "&H" & Right("00" & Hex(Val(txtExitY)), 2)
  exits(vsbExits).bY = Val(txtExitY)
  rendersprites
  dirty = True
End Sub

Private Sub txtSignScript_LostFocus()
  txtSignScript.Text = "&H" & Right("000000" & Hex(Val(txtSignScript)), 6)
  signs(vsbSigns).pScript = Val(txtSignScript) + &H8000000
  dirty = True
End Sub

Private Sub txtTrapX_LostFocus()
  txtTrapX.Text = "&H" & Right("00" & Hex(Val(txtTrapX)), 2)
  traps(vsbTraps).bX = Val(txtTrapX)
  rendersprites
  dirty = True
End Sub

Private Sub txtTrapY_LostFocus()
  txtTrapY.Text = "&H" & Right("00" & Hex(Val(txtTrapY)), 2)
  traps(vsbTraps).bY = Val(txtTrapY)
  rendersprites
  dirty = True
End Sub

Private Sub txtSignX_LostFocus()
  txtSignX.Text = "&H" & Right("00" & Hex(Val(txtSignX)), 2)
  signs(vsbSigns).bX = Val(txtSignX)
  rendersprites
  dirty = True
End Sub

Private Sub txtSignY_LostFocus()
  txtSignY.Text = "&H" & Right("00" & Hex(Val(txtSignY)), 2)
  signs(vsbSigns).bY = Val(txtSignY)
  rendersprites
  dirty = True
End Sub

Private Sub txtTrapFlag_LostFocus()
  txtTrapFlag.Text = "&H" & Right("0000" & Hex(Val(txtTrapFlag)), 4)
  traps(vsbTraps).hFlagCheck = Val(txtTrapFlag)
  dirty = True
End Sub

Private Sub txtTrapValue_LostFocus()
  txtTrapValue.Text = "&H" & Right("0000" & Hex(Val(txtTrapValue)), 4)
  traps(vsbTraps).hFlagValue = Val(txtTrapValue)
  dirty = True
End Sub

Private Sub txtTrapScript_LostFocus()
  txtTrapScript.Text = "&H" & Right("000000" & Hex(Val(txtTrapScript)), 6)
  traps(vsbTraps).pScript = Val(txtTrapScript) + &H8000000
  dirty = True
End Sub

Private Sub vsbConn_Change()
  cboConnDir.ListIndex = mapConnects(vsbConn.Value).wDirection - 1
  txtConnOffset.Text = mapConnects(vsbConn.Value).wOffset
  txtConnLevel.Text = "&H" & Right("0000" & Hex(mapConnects(vsbConn.Value).hLevel), 4)
End Sub

Private Sub vsbPeepSprite_Change()
  txtPeepSprite.Text = "&H" & Right("00" & Hex(vsbPeepSprite.Value), 2)
  BitBlt picSprite.hdc, 0, 0, 32, 64, picSpritestrip.hdc, vsbPeepSprite.Value * 32, 0, vbSrcCopy
  'txtPeepSprite.Text = "&H" & Right("00" & Hex(Val(txtPeepSprite)), 2)
End Sub

Private Sub vsbPeepSprite_Scroll()
  vsbPeepSprite_Change
End Sub

Private Sub vsbTileset_Scroll()
  vsbTileset_Change
End Sub

Private Sub vsbTraps_Change()
  'Matt - Added this little hack to make sure it's a trap w00t for unknown bytes
  'Kawa - Bad karma Matt. You shouldn't just assume it's good like that.
'  If traps(vsbTraps).h3 <> 3 Then
'    If MsgBox(LoadResString(407), vbYesNo, LoadResString(408)) = vbYes Then
'      traps(vsbTraps).h3 = &H3
'      traps(vsbTraps).b2 = 0
'      traps(vsbTraps).b4 = 0
'    End If
'  End If
  'Kawa - See? It's much safer and friendlier like this!
  
  txtTrapScript.Text = "&H" & Right("000000" & Hex(traps(vsbTraps).pScript - &H8000000), 6)
  txtTrapX.Text = "&H" & Right("00" & Hex(traps(vsbTraps).bX), 2)
  txtTrapY.Text = "&H" & Right("00" & Hex(traps(vsbTraps).bY), 2)
  txtTrapFlag.Text = "&H" & Right("0000" & Hex(traps(vsbTraps).hFlagCheck), 4)
  txtTrapValue.Text = "&H" & Right("0000" & Hex(traps(vsbTraps).hFlagValue), 4)
End Sub

Private Sub vsbSigns_Change()
  On Error GoTo noSigns
  txtSignScript.Text = "&H" & Right("000000" & Hex(signs(vsbSigns).pScript - &H8000000), 6)
  txtSignX.Text = "&H" & Right("00" & Hex(signs(vsbSigns).bX), 2)
  txtSignY.Text = "&H" & Right("00" & Hex(signs(vsbSigns).bY), 2)
  txtSignScript.Enabled = True
  Exit Sub
noSigns:
  txtSignScript.Enabled = False
  Resume Next
End Sub

Private Sub vsbExits_Change()
  txtExitLevel.Text = "&H" & Right("0000" & Hex(exits(vsbExits).hLevel), 4)
  txtExitTarget.Text = "&H" & Right("00" & Hex(exits(vsbExits).b6), 2)
  txtExitX.Text = "&H" & Right("00" & Hex(exits(vsbExits).bX), 2)
  txtExitY.Text = "&H" & Right("00" & Hex(exits(vsbExits).bY), 2)
End Sub

Private Sub vsbPeeps_Change()
  txtPeepSprite.Text = "&H" & Right("00" & Hex(peoples(vsbPeeps).bSpriteSet), 2)
  vsbPeepSprite.Value = peoples(vsbPeeps).bSpriteSet
  vsbPeepSprite_Change
  txtPeepScript.Text = "&H" & Right("000000" & Hex(peoples(vsbPeeps).pScript - &H8000000), 6)
  cboPeepBehave.ListIndex = peoples(vsbPeeps).bBehavior1
  txtPeepFlag.Text = "&H" & Right("0000" & Hex(peoples(vsbPeeps).iFlag), 4)
  txtPeepX.Text = "&H" & Right("00" & Hex(peoples(vsbPeeps).bX), 2)
  txtPeepY.Text = "&H" & Right("00" & Hex(peoples(vsbPeeps).bY), 2)
  If peoples(vsbPeeps).bIsTrainer = 0 Then
    chkPeepIsTrainer.Value = 0
  ElseIf peoples(vsbPeeps).bIsTrainer = 1 Then
    chkPeepIsTrainer.Value = 1
  Else
    chkPeepIsTrainer.Enabled = False
  End If
  txtPeepLOS.Text = "&H" & Right("00" & Hex(peoples(vsbPeeps).bTrainerLOS), 2)
End Sub

'Why did the chicken cross the sub?
Private Sub vsbTileset_Change()
  On Error Resume Next
  picTileset.Move picTileset.Left, -(vsbTileset * &H10) '* &H40) ' &H80)
  picTileset.SetFocus
End Sub

Private Sub vsbScroll_Change()
  On Error Resume Next
  Picture1.Move Picture1.Left, -(vsbScroll * &H10)
  shLoc.Move (worldlocs(thislevel.bLabelID).bX + 1) * 8, (worldlocs(thislevel.bLabelID).bY + 2) * 8, worldlocs(thislevel.bLabelID).bW * 8, worldlocs(thislevel.bLabelID).bH * 8
  Picture1.SetFocus
End Sub

Private Sub maplabelread()
  On Error Resume Next
  Dim inbyte As Byte
  i = 0
  'Val("&H" & txtLevel)
  X = ""
  cboLabelID.Clear
  lstLabelID.Clear
  Do While i < &H59
    data = ""
    Get #256, (xm + 1) + (i * 8), worldlocs(i)
    pc = getgbapointer((xm + 4) + (i * 8)) + 1
    xpc = pc
    Do
      Get #256, pc, inbyte
      data = data & IIf(inbyte = 255, "", Chr(inbyte))
      pc = pc + 1
    Loop Until inbyte = 255
    MapLabels(i) = Replace(Replace(Sapp2Asc(data, romisjapanese), "\c\h00", ""), "\v\h08", "[TEAM]")
    cboLabelID.AddItem Right("00" & Hex(i), 2) & ". " & MapLabels(i)
    lstLabelID.AddItem Right("00" & Hex(i), 2) & ". " & MapLabels(i)
    i = i + 1
  Loop
End Sub

Private Sub maplabelreadNG()
  On Error Resume Next
  Dim inbyte As Byte
  i = 0
  'Val("&H" & txtLevel)
  X = ""
  cboLabelID.Clear
  lstLabelID.Clear
  'pc = &H3B5A48 + 1 'Red only for now...
  Dim pc As Long
  Dim safeguard As Integer
  Get #256, xm + 1, pc
  pc = pc - &H8000000 + 1
  Do While i < &H6D
    data = ""
    safeguard = 0
    Do
      Get #256, pc, inbyte
      'Trace Hex(inbyte)
      data = data & IIf(inbyte = 255, "", Chr(inbyte))
      pc = pc + 1
      safeguard = safeguard + 1
    Loop Until inbyte = 255 Or safeguard = 40
    MapLabels(i) = Sapp2Asc(data, romisjapanese)
    cboLabelID.AddItem Right("00" & Hex(i), 2) & ". " & Trim(MapLabels(i)) & IIf(safeguard = 40, "<!>", "")
    lstLabelID.AddItem Right("00" & Hex(i), 2) & ". " & Trim(MapLabels(i))
    i = i + 1
  Loop
  i = 0
  Seek #256, xm + 436 + 1
  Do While i < &H6D
    Get #256, , inbyte
    worldlocs(i).bX = inbyte
    Get #256, , inbyte
    Get #256, , inbyte
    worldlocs(i).bY = inbyte
    Get #256, , inbyte
    i = i + 1
  Loop
  i = 0
  Seek #256, xm + 1228 + 1
  Do While i < &H6D
    Get #256, , inbyte
    worldlocs(i).bW = inbyte
    Get #256, , inbyte
    Get #256, , inbyte
    worldlocs(i).bH = inbyte
    Get #256, , inbyte
    i = i + 1
  Loop
End Sub

'To get to the other side!
Private Sub rendersprites()
  For i = 0 To 63
    sPeople(i).Visible = IIf(i + 1 > thissprite.bPeople Or chkSprites.Value = vbUnchecked Or chkSPeople.Value = vbUnchecked, False, True)
    sExit(i).Visible = IIf(i + 1 > thissprite.bExits Or chkSprites.Value = vbUnchecked Or chkSExits.Value = vbUnchecked, False, True)
    sTrap(i).Visible = IIf(i + 1 > thissprite.bTraps Or chkSprites.Value = vbUnchecked Or chkSTraps.Value = vbUnchecked, False, True)
    sSign(i).Visible = IIf(i + 1 > thissprite.bSigns Or chkSprites.Value = vbUnchecked Or chkSSigns.Value = vbUnchecked, False, True)
    sPeople(i).Move peoples(i).bX * &H10, peoples(i).bY * &H10
    sExit(i).Move exits(i).bX * &H10, exits(i).bY * &H10
    sTrap(i).Move traps(i).bX * &H10, traps(i).bY * &H10
    sSign(i).Move signs(i).bX * &H10, signs(i).bY * &H10
  Next i
  imgBoyStart.Visible = IIf(chkSprites.Value = vbUnchecked, False, True)
  imgGirlStart.Visible = IIf(chkSprites.Value = vbUnchecked, False, True)
End Sub

'Render Connects, mahvellously fucked up by Tau!
Private Sub renderconnects()
Dim zwisp As Integer
Dim xtimes() As Variant
Dim xx(0 To 5) As Variant
  For i = 0 To 15
    cmdAdjMap(i).Enabled = False
  Next i
  'check how many times one direction is used
  If thisconnect.wConnects = 0 Then Exit Sub
  ReDim xtimes(thisconnect.wConnects - 1)
  For i = 0 To thisconnect.wConnects - 1
  xtimes(i) = mapConnects(i).wDirection - 1
  Next i
  xx(0) = 0
  xx(1) = 0
  xx(2) = 0
  xx(3) = 0
  xx(4) = 0
  xx(5) = 0
  For i = 0 To thisconnect.wConnects - 1
  If xtimes(i) = 0 Then xx(0) = xx(0) + 1
  If xtimes(i) = 1 Then xx(1) = xx(1) + 1
  If xtimes(i) = 2 Then xx(2) = xx(2) + 1
  If xtimes(i) = 3 Then xx(3) = xx(3) + 1
  If xtimes(i) = 4 Then xx(4) = xx(4) + 1
  If xtimes(i) = 5 Then xx(5) = xx(5) + 1
  Next i
  'move and resize buttons
  For i = 0 To 3
  If xx(i) = 1 Then
   If i = 0 Then
   cmdAdjMap(i).Width = 345
   cmdAdjMap(14).Visible = False
   cmdAdjMap(15).Visible = False
   End If
   If i = 1 Then
   cmdAdjMap(i).Width = 309
   cmdAdjMap(8).Visible = False
   cmdAdjMap(9).Visible = False
   End If
   If i = 2 Then
   cmdAdjMap(i).Height = 336
   cmdAdjMap(10).Visible = False
   cmdAdjMap(11).Visible = False
   End If
   If i = 3 Then
   cmdAdjMap(i).Height = 336
   cmdAdjMap(12).Visible = False
   cmdAdjMap(13).Visible = False
   End If
  End If
  If xx(i) = 2 Then
  If i = 0 Then
   cmdAdjMap(i).Width = 172
   cmdAdjMap(14).Width = 173
   cmdAdjMap(14).Left = cmdAdjMap(i).Left + 172
   cmdAdjMap(14).Visible = True
   cmdAdjMap(15).Visible = False
  End If
  If i = 1 Then
   cmdAdjMap(i).Width = 154
   cmdAdjMap(8).Width = 155
   cmdAdjMap(8).Left = cmdAdjMap(i).Left + 154
   cmdAdjMap(8).Visible = True
   cmdAdjMap(9).Visible = False
  End If
  If i = 2 Then
   cmdAdjMap(i).Height = 168
   cmdAdjMap(10).Height = 168
   cmdAdjMap(10).Top = cmdAdjMap(i).Top + 168
   cmdAdjMap(10).Visible = True
   cmdAdjMap(11).Visible = False
  End If
  If i = 3 Then
   cmdAdjMap(i).Height = 168
   cmdAdjMap(12).Height = 168
   cmdAdjMap(12).Top = cmdAdjMap(i).Top + 168
   cmdAdjMap(12).Visible = True
   cmdAdjMap(13).Visible = False
  End If
  End If
  If xx(i) = 3 Then
    If i = 0 Then
    cmdAdjMap(i).Width = 115
    cmdAdjMap(14).Width = 115
    cmdAdjMap(15).Width = 115
    cmdAdjMap(14).Left = cmdAdjMap(i).Left + 115
    cmdAdjMap(15).Left = cmdAdjMap(14).Left + 115
    cmdAdjMap(14).Visible = True
    cmdAdjMap(15).Visible = True
    End If
    If i = 1 Then
    cmdAdjMap(i).Width = 103
    cmdAdjMap(8).Width = 103
    cmdAdjMap(9).Width = 103
    cmdAdjMap(8).Left = cmdAdjMap(i).Left + 103
    cmdAdjMap(9).Left = cmdAdjMap(8).Left + 103
    cmdAdjMap(8).Visible = True
    cmdAdjMap(9).Visible = True
    End If
    If i = 2 Then
    cmdAdjMap(i).Height = 112
    cmdAdjMap(10).Height = 112
    cmdAdjMap(11).Height = 112
    cmdAdjMap(10).Top = cmdAdjMap(i).Top + 112
    cmdAdjMap(11).Top = cmdAdjMap(10).Top + 112
    cmdAdjMap(10).Visible = True
    cmdAdjMap(11).Visible = True
    End If
    If i = 3 Then
    cmdAdjMap(i).Height = 112
    cmdAdjMap(12).Height = 112
    cmdAdjMap(13).Height = 112
    cmdAdjMap(12).Top = cmdAdjMap(i).Top + 112
    cmdAdjMap(13).Top = cmdAdjMap(12).Top + 112
    cmdAdjMap(12).Visible = True
    cmdAdjMap(13).Visible = True
    End If
   End If
  Next i
  If xx(0) = 0 Then GoTo xxx1:
  For i = 0 To xx(0) - 1
  cmdAdjMap(0).Enabled = True
  If i = 1 Then cmdAdjMap(14).Enabled = True
  If i = 2 Then cmdAdjMap(15).Enabled = True
  z0 = 0
    For ii = 0 To thisconnect.wConnects - 1
    If mapConnects(ii).wDirection <> 1 Then GoTo search1
    If z0 = i Then Exit For
    z0 = z0 + 1
search1:
    Next ii
    zwisp = mapConnects(ii).b1
    If i = 0 Then cmdAdjMaps(0) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
    If i = 1 Then cmdAdjMaps(14) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
    If i = 2 Then cmdAdjMaps(15) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
  Next i
xxx1:
    If xx(1) = 0 Then GoTo xxx2:
  For i = 0 To xx(1) - 1
  cmdAdjMap(1).Enabled = True
  If i = 1 Then cmdAdjMap(8).Enabled = True
  If i = 2 Then cmdAdjMap(9).Enabled = True
  z1 = 0
    For ii = 0 To thisconnect.wConnects - 1
    If mapConnects(ii).wDirection <> 2 Then GoTo search2
    If z1 = i Then Exit For
    z1 = z1 + 1
search2:
    Next ii
    zwisp = mapConnects(ii).b1
    If i = 0 Then cmdAdjMaps(1) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
    If i = 1 Then cmdAdjMaps(8) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
    If i = 2 Then cmdAdjMaps(9) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
  Next i
xxx2:
    If xx(2) = 0 Then GoTo xxx3:
  For i = 0 To xx(2) - 1
  cmdAdjMap(2).Enabled = True
  If i = 1 Then cmdAdjMap(10).Enabled = True
  If i = 2 Then cmdAdjMap(11).Enabled = True
  z2 = 0
    For ii = 0 To thisconnect.wConnects - 1
    If mapConnects(ii).wDirection <> 3 Then GoTo search3
    If z2 = i Then Exit For
    z2 = z2 + 1
search3:
    Next ii
    zwisp = mapConnects(ii).b1
    If i = 0 Then cmdAdjMaps(2) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
    If i = 1 Then cmdAdjMaps(10) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
    If i = 2 Then cmdAdjMaps(11) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
  Next i
xxx3:
    If xx(3) = 0 Then GoTo xxx4:
  For i = 0 To xx(3) - 1
  cmdAdjMap(3).Enabled = True
  If i = 1 Then cmdAdjMap(12).Enabled = True
  If i = 2 Then cmdAdjMap(13).Enabled = True
  z3 = 0
    For ii = 0 To thisconnect.wConnects - 1
    If mapConnects(ii).wDirection <> 4 Then GoTo search4
    If z3 = i Then Exit For
    z3 = z3 + 1
search4:
    Next ii
    zwisp = mapConnects(ii).b1
    If i = 0 Then cmdAdjMaps(3) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
    If i = 1 Then cmdAdjMaps(12) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
    If i = 2 Then cmdAdjMaps(13) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
  Next i
xxx4:
  If xx(4) = 0 Then GoTo xxx5
  cmdAdjMap(4).Enabled = True
    For ii = 0 To thisconnect.wConnects - 1
        If mapConnects(ii).wDirection <> 5 Then GoTo search5
    Exit For
search5:
    Next ii
    zwisp = mapConnects(ii).b1
    cmdAdjMaps(4) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
xxx5:
  If xx(5) = 0 Then GoTo xxx6
  cmdAdjMap(5).Enabled = True
    For ii = 0 To thisconnect.wConnects - 1
        If mapConnects(ii).wDirection <> 6 Then GoTo search6
    Exit For
search6:
    Next ii
    zwisp = mapConnects(ii).b1
    cmdAdjMaps(5) = zwisp & Right("00" & mapConnects(ii).hLevel, 2)
xxx6:
End Sub

Private Sub ApplyTileset(ByVal bank As Byte, tileset As TilesetHeader)
  'KAWA --- ... O_O ...  HOLY MOTHER OF FUCK!
  '         And I thought Bouché wrote dirty code >_<
  '         It hurts my eyes just to read this sub!
  '         I still love you anyway.
  
  '         P.S.: At least properly indent ;)
  
  Dim DataIn(0 To 32767) As Byte
  Dim DataIn2(0 To 32767) As Byte
  Dim byteIn As Byte
  
  Dim cachefile As String
  cachefile = App.Path & "\tilecache\" & Replace(Left(Right(txtRom.Text, 12), 8), "\", "_") & txtOldskoolChooser.Text & bank & ".bmp"
  If Dir(cachefile) <> "" Then
    picTileset.Picture = LoadPicture(cachefile)
    Exit Sub
  End If
  
  Get #256, GBA2PC(thistileseta.pGFX) + 1, DataIn
  Get #256, GBA2PC(thistilesetb.pGFX) + 1, DataIn2
  
  'Decompress GFX
  If bank = 0 Then
    If checkifcompa <> 1 Then
      zahl = 204780
      If NextGen = False Then zahl = 16384
      CopyMemory gfxA(0), DataIn(0), zahl
      GoTo no_comp1
    End If
    LZ77UnComp DataIn(), gfxA()
no_comp1:
    If checkifcompb <> 1 Then
      zahl = 12288
      If NextGen = False Then zahl = 16384
      CopyMemory gfxB(0), DataIn2(0), zahl
      GoTo no_comp
    End If
    LZ77UnComp DataIn2(), gfxB()
  End If
no_comp:
  If bank <> 0 Then GoTo already_added
  If NextGen = False Then CopyMemory gfxA(16384), gfxB(0), 16384
  If NextGen = True Then CopyMemory gfxA(20480), gfxB(0), 12288
already_added:
  'If tattr(0) <> "&H30" Then GoTo nodebug1 'KAWA --- Made it an easter egg.
  Open App.Path & "\tileset.bin" For Binary As #512
  For z = 0 To 32767
    Put #512, z + 1, gfxA(z) 'KAWA --- Why was there a +336 in the offset?
  Next z
  Close #512
nodebug1:
  If NextGen = False Then
    adder = &HC0
    rounds = 5
    rounds2 = 6
  End If
  If NextGen = True Then
    adder = &HE0
    rounds = 6
    rounds2 = 7
  End If
  
  On Error GoTo PalError
  paloffsetx = GBA2PC(thistileseta.pPalettes) + adder * checkifpala
  For aa = 0 To 1
    For i = 0 To rounds
      For j = 0 To &HF
        Get #256, paloffsetx + 1 + (i * &H20) + (j * 2), byteIn
        pal = byteIn
        Get #256, paloffsetx + 2 + (i * &H20) + (j * 2), byteIn
        pal = CInt(pal + (byteIn * 256))
        c2 = (pal \ &H400) Mod &H20
        c1 = (pal \ &H20) Mod &H20
        C0 = pal Mod &H20
        pal = ((c2 * 8) * CLng(&H10000)) + ((c1 * 8) * &H100) + (C0 * 8)
        palettesA(i + rounds2 * aa, j) = pal
      Next j
    Next i
    paloffsetx = GBA2PC(thistilesetb.pPalettes) + adder * checkifpalb
  Next aa
  'Read Palettes
  On Error GoTo 0
  
  If bank = 0 Then
    Get #256, GBA2PC(tileset.pMap) + 1, Map16A
  ElseIf bank = 1 Then
    Get #256, GBA2PC(tileset.pMap) + 1, Map16B
  End If
  'Read Map16
  If NextGen = False Then endvalue = &H1FF
  If NextGen = True Then
    If bank = 0 Then endvalue = &H27F
    If bank = 1 Then endvalue = &H17F
  End If
  For i = 0 To endvalue
    DoEvents
    X = (i Mod &H10) * &H10
    If NextGen = False Then Y = ((i \ &H10) + IIf(bank = 1, &H20, 0)) * &H10
    If NextGen = True Then Y = ((i \ &H10) + IIf(bank = 1, &H28, 0)) * &H10
    DrawMap16 picTileset.hdc, bank, i, X, Y
  Next i
  
  picTileset.Picture = picTileset.Image
  On Error Resume Next
  MkDir App.Path & "\tilecache"
  SavePicture picTileset.Picture, cachefile
  Exit Sub
PalError:
  If BeenThereDoneThat = 0 Then
    MsgBox LoadResString(155)
    BeenThereDoneThat = 1
  End If
  Resume Next
End Sub

Public Sub DrawMap16(ByVal hdc As Long, ByVal bank As Byte, ByVal map16n As Long, ByVal destX As Long, ByVal destY As Long)
  For i = 0 To 1
    For Y = 0 To 1
      For X = 0 To 1
        offset = (map16n * &H10) + (i * 8) + (Y * 4) + (X * 2)
        If bank = 0 Then
          tileno = (Map16A(offset + 1) * CLng(&H100)) + Map16A(offset)
        ElseIf bank = 1 Then
          tileno = (Map16B(offset + 1) * CLng(&H100)) + Map16B(offset)
        End If
        mx = destX + (X * 8)
        my = destY + (Y * 8)
        DrawTile8 hdc, bank, tileno, mx, my
      Next X
    Next Y
  Next i
End Sub

Public Sub DrawTile8(ByVal hdc As Long, ByVal bank As Byte, ByVal map16n As Long, ByVal destX As Long, ByVal destY As Long)
  Dim tiledata(0 To 31) As Byte
  tile = map16n And &H3FF
  If bank = 0 Then
    CopyMemory tiledata(0), gfxA(tile * 32), 32
  ElseIf bank = 1 Then
    CopyMemory tiledata(0), gfxA((tile) * 32), 32
  End If
  flipx = IIf((map16n And &H400) = &H400, True, False)
  flipy = IIf((map16n And &H800) = &H800, True, False)
  pal = map16n \ &H1000
  For i = 0 To 31
    X1 = (i * 2) Mod 8
    X2 = X1 + 1
    If flipx = True Then
      X1 = 7 - X1
      X2 = 7 - X2
    End If
    Y = (i \ 4)
    If flipy = True Then
      Y = 7 - Y
    End If
    
    colA = tiledata(i) Mod &H10
    colB = tiledata(i) \ &H10
    If bank = 0 Then
      If Not (colA = 0) Then SetPixel hdc, (destX) + X1, destY + Y, palettesA(pal, colA)
      If Not (colB = 0) Then SetPixel hdc, (destX) + X2, destY + Y, palettesA(pal, colB)
    ElseIf bank = 1 Then
      If Not (colA = 0) Then SetPixel hdc, (destX) + X1, destY + Y, palettesA(pal, colA)
      If Not (colB = 0) Then SetPixel hdc, (destX) + X2, destY + Y, palettesA(pal, colB)
    End If
  Next i
End Sub

Public Function h2d(Number As String)
dec = CLng("&H" & Number)
h2d = dec
End Function

Public Function LOADBanks(entry As Variant)
  worldentry = entry
  cboBanks.Clear
  cboBanks.Text = "Bank"
  Index = 0
  For X = 0 To 1024 ' Just for security so it doesn't exploit all ;)
    headers = getgbapointer(getgbapointer(Roms(entry).MapHeaders) + X * 4)
    If headers = -1 Then Exit For
    cboBanks.AddItem "B " & Hex(Index), Index
    Index = Index + 1
  Next X
End Function

Public Function LOADLevels()
  headers = getgbapointer(getgbapointer(Roms(worldentry).MapHeaders) + cboBanks.ListIndex * 4)
  On Error GoTo lastbank
  headers2 = getgbapointer(getgbapointer(Roms(worldentry).MapHeaders) + (cboBanks.ListIndex + 1) * 4)
  GoTo go_on
lastbank:
  headers2 = getgbapointer(Roms(worldentry).MapHeaders)
go_on:
  cboLevels.Clear
  cboLevels.Text = "Level"
  Index = 0
  For X = 0 To 1024 'security
    header2 = getgbapointer(headers + X * 4)
    If header2 = -1 Then Exit For
    If headers + X * 4 = headers2 Then Exit For
    cboLevels.AddItem "L " & Hex(Index), Index
    Index = Index + 1
  Next X
End Function

Public Sub write_bank_lev(bank As Variant, lev As Variant)
  On Error GoTo Ende
  cboBanks.ListIndex = bank
  cboBanks_Click
  cboLevels.ListIndex = lev
Ende:
End Sub

Public Sub get_bank_lev()
txtbank = cboBanks.ListIndex
txtlevel = cboLevels.ListIndex
End Sub
