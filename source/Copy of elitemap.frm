VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EliteMap"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11880
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
   ScaleHeight     =   550
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timAutoSave 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   11280
      Top             =   0
   End
   Begin VB.PictureBox picPanel 
      BorderStyle     =   0  'None
      Height          =   1455
      Index           =   2
      Left            =   0
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   281
      TabIndex        =   135
      Tag             =   "97"
      Top             =   4005
      Width           =   4215
      Begin VB.CommandButton cmdReplaceRL 
         Caption         =   "Replace R with L"
         Height          =   255
         Left            =   2160
         TabIndex        =   145
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdSwapRL 
         Caption         =   "Swap R with L"
         Height          =   255
         Left            =   2160
         TabIndex        =   144
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdSetLAtts 
         Caption         =   "Set L tiles to L attribute"
         Height          =   255
         Left            =   2160
         TabIndex        =   143
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   120
         TabIndex        =   137
         Top             =   360
         Width           =   1935
         Begin VB.CheckBox chkSSigns 
            Caption         =   "Signs"
            Height          =   255
            Left            =   1080
            TabIndex        =   142
            Top             =   600
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkSTraps 
            Caption         =   "Traps"
            Height          =   255
            Left            =   1080
            TabIndex        =   141
            Top             =   360
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkSExits 
            Caption         =   "Exits"
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   600
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkSPeople 
            Caption         =   "People"
            Height          =   255
            Left            =   120
            TabIndex        =   139
            Top             =   360
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkSprites 
            Caption         =   "Show Objects"
            Height          =   255
            Left            =   120
            TabIndex        =   138
            Top             =   0
            Value           =   1  'Checked
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdPanel 
         Caption         =   "Tile Tools"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   136
         Tag             =   "^_^"
         ToolTipText     =   "Click to toggle the World Map panel"
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox picPanel 
      BorderStyle     =   0  'None
      Height          =   2715
      Index           =   3
      Left            =   0
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   44
      Tag             =   "181"
      Top             =   5460
      Width           =   4200
      Begin VB.PictureBox picWorldMap 
         Height          =   2415
         Left            =   15
         ScaleHeight     =   157
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   253
         TabIndex        =   46
         Top             =   270
         Width           =   3855
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
      End
      Begin VB.CommandButton cmdPanel 
         Caption         =   "World Map"
         Height          =   255
         Index           =   3
         Left            =   30
         TabIndex        =   45
         Tag             =   "^_^"
         ToolTipText     =   "Click to toggle the World Map panel"
         Top             =   15
         Width           =   1215
      End
   End
   Begin VB.PictureBox picPanel 
      BorderStyle     =   0  'None
      Height          =   2280
      Index           =   0
      Left            =   0
      ScaleHeight     =   152
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   37
      Tag             =   "152"
      Top             =   360
      Width           =   4200
      Begin VB.PictureBox picTileBox 
         Height          =   1980
         Left            =   15
         ScaleHeight     =   128
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   273
         TabIndex        =   39
         Top             =   270
         Width           =   4160
         Begin VB.VScrollBar vsbTileset 
            Height          =   1920
            LargeChange     =   4
            Left            =   3840
            Max             =   56
            TabIndex        =   41
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox picTileset 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            Height          =   15360
            Left            =   0
            MouseIcon       =   "elitemap.frx":030A
            MousePointer    =   99  'Custom
            ScaleHeight     =   1024
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   257
            TabIndex        =   40
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
      End
      Begin VB.CommandButton cmdPanel 
         Caption         =   "Blocks"
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   38
         Tag             =   "^_^"
         ToolTipText     =   "Click to toggle the Blocks panel"
         Top             =   15
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3000
         TabIndex        =   43
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label lblTilesetLoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1320
         TabIndex        =   42
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox picPanel 
      BorderStyle     =   0  'None
      Height          =   1320
      Index           =   1
      Left            =   0
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   280
      TabIndex        =   33
      Tag             =   "88"
      Top             =   2685
      Width           =   4200
      Begin VB.PictureBox picAttributes 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1020
         Left            =   15
         MouseIcon       =   "elitemap.frx":045C
         MousePointer    =   99  'Custom
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   35
         Top             =   240
         Width           =   3900
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
      Begin VB.CommandButton cmdPanel 
         Caption         =   "Attributes"
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   34
         ToolTipText     =   "Click to toggle the Attribute panel"
         Top             =   15
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   15
         Width           =   1575
      End
   End
   Begin MSComctlLib.ImageList ilsNormal 
      Left            =   9480
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":05AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":0900
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":0C52
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":0FA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":12F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":1648
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":199A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":1CEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":203E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":2390
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":26E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":2A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "elitemap.frx":2D86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   345
      Left            =   4200
      TabIndex        =   32
      Top             =   375
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   609
      Style           =   1
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Map"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Header"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Objects"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   10920
      TabIndex        =   4
      Text            =   "0"
      Top             =   9960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   255
      Left            =   8640
      TabIndex        =   10
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   255
      Left            =   7440
      TabIndex        =   9
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdjMap 
      Caption         =   "Adj. 7"
      Enabled         =   0   'False
      Height          =   255
      Index           =   7
      Left            =   7440
      TabIndex        =   12
      Top             =   9720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdjMap 
      Caption         =   "Adj. 6"
      Enabled         =   0   'False
      Height          =   255
      Index           =   6
      Left            =   6240
      TabIndex        =   11
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   9960
      TabIndex        =   8
      Text            =   "&H0000"
      Top             =   9960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   10920
      TabIndex        =   7
      Text            =   "0"
      Top             =   9720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox t 
      BackColor       =   &H00000000&
      Height          =   135
      Left            =   11280
      ScaleHeight     =   5
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   3
      Top             =   10080
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CheckBox cFlood 
      Caption         =   "Check1"
      Height          =   255
      Left            =   9960
      TabIndex        =   13
      Top             =   9720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox tattr 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   10080
      TabIndex        =   1
      Text            =   "&Hc"
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox tattr 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   9480
      TabIndex        =   0
      Text            =   "&H1"
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox tattr 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   10680
      TabIndex        =   2
      Text            =   "&H4"
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Palette"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   9480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Text            =   "0"
      Top             =   9480
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.Toolbar tlbToolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   167
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilsNormal"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   196
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "browse"
            Object.ToolTipText     =   "Browse for a rom"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   186
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save this level to the rom"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "gohome"
            Object.ToolTipText     =   "Go home"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copylevel"
            Object.ToolTipText     =   "Copy a bitmap of the level map"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copytileset"
            Object.ToolTipText     =   "Copy a bitmap of the map tileset"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "clear"
            Object.ToolTipText     =   "Wipe out the map"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "resize"
            Object.ToolTipText     =   "Resize the map"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "viewscript"
            Object.ToolTipText     =   "View the map's script"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "loadex"
            Object.ToolTipText     =   "Load an ExMap file"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "saveex"
            Object.ToolTipText     =   "Save this map to an ExMap file"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "launch"
            Object.ToolTipText     =   "Run other EM programs"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   11
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "baseedit"
                  Text            =   "BaseEdit"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "bewildered"
                  Text            =   "Bewildered"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "rsball"
                  Text            =   "RS-Ball"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "spread"
                  Text            =   "Spread"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "dexter"
                  Text            =   "Dexter"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "trained"
                  Text            =   "TrainEd"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "lips"
                  Text            =   "LIPS"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "pokepic"
                  Text            =   "PokePic"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "snesedit"
                  Text            =   "SnesEdit"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Key             =   "tlp"
                  Text            =   "Tile Layer Pro"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "web"
            Object.ToolTipText     =   "Visit the Helmeted Rodent page"
            ImageIndex      =   13
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox picToolbarLevelBox 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   3360
         ScaleHeight     =   300
         ScaleWidth      =   2775
         TabIndex        =   171
         Top             =   8
         Width           =   2775
         Begin VB.ComboBox cmb2 
            Height          =   315
            Left            =   1440
            TabIndex        =   189
            Text            =   "Level"
            Top             =   0
            Width           =   1335
         End
         Begin VB.ComboBox cmb1 
            Height          =   315
            Left            =   0
            TabIndex        =   188
            Text            =   "Bank"
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.PictureBox picToolbarRomBox 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   120
         ScaleHeight     =   300
         ScaleWidth      =   2775
         TabIndex        =   168
         Top             =   8
         Width           =   2775
         Begin VB.TextBox txtRom 
            Height          =   285
            Left            =   480
            Locked          =   -1  'True
            TabIndex        =   169
            Top             =   0
            Width           =   2295
         End
         Begin VB.Label Label12 
            Caption         =   "ROM"
            Height          =   165
            Left            =   0
            TabIndex        =   170
            Top             =   30
            Width           =   495
         End
      End
   End
   Begin VB.PictureBox picMainTab 
      BorderStyle     =   0  'None
      Height          =   2415
      Index           =   1
      Left            =   4200
      ScaleHeight     =   2415
      ScaleWidth      =   7695
      TabIndex        =   47
      Top             =   705
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ComboBox cboSong 
         Height          =   315
         ItemData        =   "elitemap.frx":30D8
         Left            =   960
         List            =   "elitemap.frx":323E
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox chkShowLabel 
         Caption         =   "Show Label on Entry"
         Height          =   255
         Left            =   3480
         TabIndex        =   56
         Top             =   120
         Width           =   1935
      End
      Begin VB.ComboBox cboLabelID 
         Height          =   315
         ItemData        =   "elitemap.frx":3C99
         Left            =   960
         List            =   "elitemap.frx":3C9B
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   120
         Width           =   2415
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "elitemap.frx":3C9D
         Left            =   960
         List            =   "elitemap.frx":3CD1
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cboWeather 
         Height          =   315
         ItemData        =   "elitemap.frx":3D6E
         Left            =   960
         List            =   "elitemap.frx":3DA2
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtLevelHeight 
         Height          =   285
         Left            =   960
         TabIndex        =   50
         ToolTipText     =   "Level height"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtLevelWidth 
         Height          =   285
         Left            =   960
         TabIndex        =   49
         ToolTipText     =   "Level width"
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label lblSongWarning 
         Caption         =   "This map has a special song value."
         Height          =   255
         Left            =   960
         TabIndex        =   187
         Top             =   1200
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label Label18 
         Caption         =   "Song"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Label"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Type"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Weather"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "Height"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "Width"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.PictureBox picMainTab 
      BorderStyle     =   0  'None
      Height          =   7425
      Index           =   0
      Left            =   4200
      ScaleHeight     =   495
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.Frame Frame6 
         Caption         =   "Stamp"
         Height          =   855
         Left            =   1080
         TabIndex        =   195
         Top             =   6240
         Width           =   735
         Begin VB.PictureBox Picture3 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000001&
            Height          =   525
            Left            =   120
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   196
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
         TabIndex        =   194
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
         TabIndex        =   193
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
         TabIndex        =   192
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
         TabIndex        =   191
         Top             =   5040
         Width           =   255
      End
      Begin VB.CommandButton cmdFullBRD 
         Caption         =   "View"
         Enabled         =   0   'False
         Height          =   525
         Left            =   360
         TabIndex        =   190
         Top             =   6480
         Visible         =   0   'False
         Width           =   525
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
         TabIndex        =   20
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
         TabIndex        =   183
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
         TabIndex        =   186
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
         TabIndex        =   185
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
         TabIndex        =   184
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
         TabIndex        =   182
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
         TabIndex        =   181
         Top             =   2280
         Width           =   255
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
         TabIndex        =   180
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
         TabIndex        =   179
         Top             =   330
         Width           =   1545
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000018&
         Height          =   1110
         Left            =   3840
         ScaleHeight     =   70
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   125
         TabIndex        =   174
         Top             =   6240
         Width           =   1935
         Begin VB.Label Label2 
            BackColor       =   &H80000018&
            ForeColor       =   &H80000017&
            Height          =   255
            Left            =   30
            TabIndex        =   178
            Top             =   780
            Width           =   1815
         End
         Begin VB.Label Label5 
            BackColor       =   &H80000018&
            ForeColor       =   &H80000017&
            Height          =   255
            Left            =   30
            TabIndex        =   177
            Top             =   30
            Width           =   1815
         End
         Begin VB.Label Label9 
            BackColor       =   &H80000018&
            ForeColor       =   &H80000017&
            Height          =   255
            Left            =   30
            TabIndex        =   176
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000018&
            ForeColor       =   &H80000017&
            Height          =   255
            Left            =   30
            TabIndex        =   175
            Top             =   525
            Width           =   1815
         End
      End
      Begin VB.VScrollBar vsbScroll 
         Height          =   5025
         Left            =   5295
         Max             =   0
         TabIndex        =   17
         Top             =   600
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
         TabIndex        =   18
         Top             =   585
         Width           =   255
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "Surface"
         Enabled         =   0   'False
         Height          =   255
         Index           =   5
         Left            =   4980
         TabIndex        =   19
         Top             =   330
         Width           =   825
      End
      Begin VB.Frame Frame5 
         Caption         =   "Border"
         Height          =   855
         Left            =   240
         TabIndex        =   160
         Top             =   6240
         Width           =   735
         Begin VB.PictureBox picBorder 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000001&
            Height          =   525
            Left            =   120
            ScaleHeight     =   31
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   31
            TabIndex        =   161
            Top             =   240
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Object Info"
         Height          =   4335
         Left            =   6000
         TabIndex        =   157
         Top             =   360
         Width           =   1575
         Begin VB.Label lblInfo 
            Height          =   3975
            Left            =   120
            TabIndex        =   158
            Top             =   240
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Shift Level"
         Height          =   1095
         Left            =   6000
         TabIndex        =   156
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
         TabIndex        =   152
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
         TabIndex        =   151
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
         TabIndex        =   150
         Top             =   6480
         Width           =   480
      End
      Begin VB.CheckBox chkNoDraw 
         Caption         =   "Disable map editing"
         Height          =   255
         Left            =   1920
         TabIndex        =   48
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton cmdAdjMap 
         Caption         =   "Dive"
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   5295
         TabIndex        =   16
         Top             =   5880
         Width           =   510
      End
      Begin VB.HScrollBar hsbScroll 
         Height          =   255
         Left            =   120
         Max             =   0
         TabIndex        =   15
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
         TabIndex        =   21
         Top             =   585
         Width           =   4935
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
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
            TabIndex        =   22
            Top             =   0
            Width           =   4260
            Begin VB.Shape Shape1 
               BorderColor     =   &H00FF0000&
               BorderWidth     =   3
               Height          =   255
               Left            =   120
               Top             =   600
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label sPeople 
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
               Left            =   1080
               MouseIcon       =   "elitemap.frx":3E8E
               MousePointer    =   99  'Custom
               TabIndex        =   26
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label sExit 
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
               Left            =   840
               MouseIcon       =   "elitemap.frx":3FE0
               MousePointer    =   99  'Custom
               TabIndex        =   25
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label sTrap 
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
               Left            =   600
               MouseIcon       =   "elitemap.frx":4132
               MousePointer    =   99  'Custom
               TabIndex        =   24
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label sSign 
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
               Left            =   360
               MouseIcon       =   "elitemap.frx":4284
               MousePointer    =   99  'Custom
               TabIndex        =   23
               Top             =   240
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.Label lblCreds 
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   360
            Left            =   120
            TabIndex        =   27
            Top             =   4560
            Width           =   4635
         End
      End
      Begin VB.Label Label10 
         Caption         =   "Hold shift to pick up a tile or Ctrl to stamp four."
         Height          =   495
         Left            =   1920
         TabIndex        =   198
         Top             =   6840
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "New in this version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   197
         Top             =   6600
         Width           =   1815
      End
      Begin VB.Label lblLvlScript 
         Alignment       =   1  'Right Justify
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
         Left            =   6360
         TabIndex        =   159
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "Left"
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   155
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Right"
         Height          =   255
         Index           =   1
         Left            =   6000
         TabIndex        =   154
         Top             =   6240
         Width           =   495
      End
      Begin VB.Label Label15 
         Caption         =   "Middle"
         Height          =   255
         Index           =   2
         Left            =   6000
         TabIndex        =   153
         Top             =   6480
         Width           =   495
      End
      Begin VB.Label lblRom 
         Caption         =   "No ROM loaded"
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
         TabIndex        =   31
         Top             =   60
         Width           =   3135
      End
      Begin VB.Label lblLevelName 
         Caption         =   "No level loaded"
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
         Left            =   3360
         TabIndex        =   30
         Top             =   60
         Width           =   2775
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   615
         Left            =   4200
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picMainTab 
      BorderStyle     =   0  'None
      Height          =   5295
      Index           =   2
      Left            =   4200
      ScaleHeight     =   5295
      ScaleWidth      =   7695
      TabIndex        =   62
      Top             =   705
      Visible         =   0   'False
      Width           =   7695
      Begin MSComctlLib.TabStrip tabSubEditor 
         Height          =   375
         Left            =   120
         TabIndex        =   148
         Top             =   120
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   661
         MultiRow        =   -1  'True
         Style           =   1
         HotTracking     =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   6
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Labels"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Connections"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "People"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Exits"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Traps"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Signs"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox picSubEditor 
         Height          =   1335
         Index           =   5
         Left            =   120
         ScaleHeight     =   1275
         ScaleWidth      =   7395
         TabIndex        =   84
         Top             =   3480
         Visible         =   0   'False
         Width           =   7455
         Begin VB.HScrollBar vsbSigns 
            Height          =   255
            Left            =   0
            TabIndex        =   162
            Top             =   0
            Width           =   7455
         End
         Begin VB.CommandButton cmdWipeSigns 
            Caption         =   "Wipe"
            Height          =   375
            Left            =   6480
            TabIndex        =   89
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtSignScript 
            Height          =   285
            Left            =   1080
            TabIndex        =   88
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtSignX 
            Height          =   285
            Left            =   1080
            TabIndex        =   87
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtSignY 
            Height          =   285
            Left            =   1920
            TabIndex        =   86
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton cmdRepointSigns 
            Caption         =   "Repoint"
            Height          =   375
            Left            =   6480
            TabIndex        =   85
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label46 
            Caption         =   "Script"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label44 
            Caption         =   "Location"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.PictureBox picSubEditor 
         Height          =   1575
         Index           =   4
         Left            =   120
         ScaleHeight     =   1515
         ScaleWidth      =   7395
         TabIndex        =   73
         Top             =   2880
         Visible         =   0   'False
         Width           =   7455
         Begin VB.HScrollBar vsbTraps 
            Height          =   255
            Left            =   0
            TabIndex        =   163
            Top             =   0
            Width           =   7455
         End
         Begin VB.CommandButton cmdWipeTraps 
            Caption         =   "Wipe"
            Height          =   375
            Left            =   6480
            TabIndex        =   80
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtTrapY 
            Height          =   285
            Left            =   1920
            TabIndex        =   79
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtTrapX 
            Height          =   285
            Left            =   1080
            TabIndex        =   78
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtTrapValue 
            Height          =   285
            Left            =   1920
            TabIndex        =   77
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtTrapFlag 
            Height          =   285
            Left            =   1080
            TabIndex        =   76
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtTrapScript 
            Height          =   285
            Left            =   1080
            TabIndex        =   75
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdRepointTraps 
            Caption         =   "Repoint"
            Height          =   375
            Left            =   6480
            TabIndex        =   74
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label43 
            Caption         =   "Location"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label40 
            Caption         =   "Flags"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label39 
            Caption         =   "Script"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.PictureBox picSubEditor 
         Height          =   2535
         Index           =   1
         Left            =   120
         ScaleHeight     =   165
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   493
         TabIndex        =   108
         Top             =   1080
         Visible         =   0   'False
         Width           =   7455
         Begin VB.HScrollBar vsbConn 
            Height          =   255
            Left            =   0
            TabIndex        =   166
            Top             =   0
            Width           =   7455
         End
         Begin VB.Frame Frame4 
            Caption         =   "Advanced"
            Height          =   975
            Left            =   120
            TabIndex        =   112
            Top             =   1320
            Width           =   4575
            Begin VB.TextBox txtConnCount 
               Height          =   285
               Left            =   2280
               TabIndex        =   115
               Top             =   600
               Width           =   855
            End
            Begin VB.CommandButton cmdConnPtr 
               Caption         =   "Set"
               Height          =   375
               Left            =   3720
               TabIndex        =   114
               Top             =   240
               Width           =   735
            End
            Begin VB.TextBox txtConnPtr 
               Height          =   285
               Left            =   2280
               TabIndex        =   113
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label38 
               Caption         =   "Number of connections"
               Height          =   255
               Left            =   120
               TabIndex        =   117
               Top             =   600
               Width           =   1935
            End
            Begin VB.Label Label29 
               Caption         =   "Connection table location"
               Height          =   255
               Left            =   120
               TabIndex        =   116
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.TextBox txtConnLevel 
            Height          =   285
            Left            =   2760
            TabIndex        =   111
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtConnOffset 
            Height          =   285
            Left            =   960
            TabIndex        =   110
            Top             =   840
            Width           =   975
         End
         Begin VB.ComboBox cboConnDir 
            Height          =   315
            ItemData        =   "elitemap.frx":43D6
            Left            =   1080
            List            =   "elitemap.frx":43EC
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label26 
            Caption         =   "Level"
            Height          =   255
            Left            =   120
            TabIndex        =   120
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label25 
            Caption         =   "Offset"
            Height          =   255
            Left            =   2040
            TabIndex        =   119
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label24 
            Caption         =   "Direction"
            Height          =   255
            Left            =   120
            TabIndex        =   118
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.PictureBox picSubEditor 
         Height          =   2295
         Index           =   0
         Left            =   120
         ScaleHeight     =   149
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   493
         TabIndex        =   121
         Top             =   480
         Visible         =   0   'False
         Width           =   7455
         Begin VB.HScrollBar vsbDummy 
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            Max             =   0
            TabIndex        =   172
            Top             =   0
            Width           =   7455
         End
         Begin VB.TextBox txtLabelLocH 
            Height          =   285
            Left            =   4800
            TabIndex        =   128
            Top             =   1440
            Width           =   615
         End
         Begin VB.TextBox txtLabelLocW 
            Height          =   285
            Left            =   3600
            TabIndex        =   127
            Top             =   1440
            Width           =   615
         End
         Begin VB.CommandButton cmdSaveLocs 
            Caption         =   "Save"
            Height          =   375
            Left            =   6480
            TabIndex        =   126
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtLabelLocY 
            Height          =   285
            Left            =   4800
            TabIndex        =   125
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtLabelLocX 
            Height          =   285
            Left            =   3600
            TabIndex        =   124
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox txtLabel 
            Height          =   285
            Left            =   3600
            Locked          =   -1  'True
            TabIndex        =   123
            Top             =   360
            Width           =   1815
         End
         Begin VB.ListBox lstLabelID 
            Height          =   1815
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   122
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label23 
            Caption         =   "by"
            Height          =   255
            Left            =   4320
            TabIndex        =   134
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label22 
            Caption         =   "Size"
            Height          =   255
            Left            =   2520
            TabIndex        =   133
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "by"
            Height          =   255
            Left            =   4320
            TabIndex        =   132
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "Location"
            Height          =   255
            Left            =   2520
            TabIndex        =   131
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label19 
            Caption         =   "Label"
            Height          =   255
            Left            =   2520
            TabIndex        =   130
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Still no label editing..."
            Height          =   255
            Left            =   2520
            TabIndex        =   129
            Top             =   720
            Width           =   2895
         End
      End
      Begin VB.PictureBox picSubEditor 
         CausesValidation=   0   'False
         Height          =   2055
         Index           =   3
         Left            =   120
         ScaleHeight     =   1995
         ScaleWidth      =   7395
         TabIndex        =   63
         Top             =   2280
         Visible         =   0   'False
         Width           =   7455
         Begin VB.HScrollBar vsbExits 
            Height          =   255
            Left            =   0
            TabIndex        =   164
            Top             =   0
            Width           =   7455
         End
         Begin VB.CommandButton cmdRepointExits 
            Caption         =   "Repoint"
            Height          =   375
            Left            =   6480
            TabIndex        =   64
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmdWipeExits 
            Caption         =   "Wipe"
            Height          =   375
            Left            =   6480
            TabIndex        =   69
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtExitY 
            Height          =   285
            Left            =   2040
            TabIndex        =   68
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtExitX 
            Height          =   285
            Left            =   1080
            TabIndex        =   67
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtExitTarget 
            Height          =   285
            Left            =   1080
            TabIndex        =   66
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox txtExitLevel 
            Height          =   285
            Left            =   1080
            TabIndex        =   65
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label42 
            Caption         =   "Location"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label33 
            Caption         =   "Exit #"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "Level"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.PictureBox picSubEditor 
         Height          =   2295
         Index           =   2
         Left            =   120
         ScaleHeight     =   2235
         ScaleWidth      =   7395
         TabIndex        =   92
         Top             =   1680
         Visible         =   0   'False
         Width           =   7455
         Begin VB.HScrollBar vsbPeeps 
            Height          =   255
            Left            =   0
            TabIndex        =   165
            Top             =   0
            Width           =   7455
         End
         Begin VB.CommandButton cmdWipePeople 
            Caption         =   "Wipe"
            Height          =   375
            Left            =   6480
            TabIndex        =   101
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtPeepY 
            Height          =   285
            Left            =   2040
            TabIndex        =   100
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtPeepX 
            Height          =   285
            Left            =   1080
            TabIndex        =   99
            Top             =   1440
            Width           =   855
         End
         Begin VB.TextBox txtPeepFlag 
            Height          =   285
            Left            =   3360
            TabIndex        =   98
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox cboPeepBehave 
            Height          =   315
            ItemData        =   "elitemap.frx":442E
            Left            =   1080
            List            =   "elitemap.frx":4732
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   1080
            Width           =   3255
         End
         Begin VB.TextBox txtPeepScript 
            Height          =   285
            Left            =   1080
            TabIndex        =   96
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtPeepSprite 
            Height          =   285
            Left            =   1080
            TabIndex        =   95
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton cmdRepointPeople 
            Caption         =   "Repoint"
            Height          =   375
            Left            =   6480
            TabIndex        =   94
            Top             =   840
            Width           =   975
         End
         Begin VB.CommandButton cmdPeepBecome 
            Caption         =   "Become..."
            Height          =   375
            Left            =   6480
            TabIndex        =   93
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label41 
            Caption         =   "Location"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label Label32 
            Caption         =   "Flags"
            Height          =   255
            Left            =   2640
            TabIndex        =   106
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label35 
            Caption         =   "TODO: Fix behavior list. It's not right..."
            Height          =   255
            Left            =   120
            TabIndex        =   105
            Top             =   1920
            Width           =   4815
         End
         Begin VB.Label Label34 
            Caption         =   "Behavior"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label31 
            Caption         =   "Script"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label27 
            Caption         =   "Sprite"
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   360
            Width           =   1575
         End
      End
   End
   Begin VB.PictureBox picMainTab 
      BorderStyle     =   0  'None
      Height          =   7455
      Index           =   3
      Left            =   4200
      ScaleHeight     =   7455
      ScaleWidth      =   7695
      TabIndex        =   146
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
      Begin VB.PictureBox picTeam 
         BorderStyle     =   0  'None
         Height          =   1305
         Left            =   5640
         ScaleHeight     =   87
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   115
         TabIndex        =   173
         Top             =   1200
         Width           =   1725
         Begin VB.Image imgHRguy 
            Height          =   255
            Index           =   8
            Left            =   360
            ToolTipText     =   "Markus/D-Kiddy - Assistant Hummus"
            Top             =   0
            Width           =   255
         End
         Begin VB.Image imgHRguy 
            Height          =   855
            Index           =   7
            Left            =   1320
            ToolTipText     =   "Saotome Ranko - General Hummus"
            Top             =   360
            Width           =   375
         End
         Begin VB.Image imgHRguy 
            Height          =   855
            Index           =   5
            Left            =   840
            ToolTipText     =   "DJ ßouché - Coding/Music"
            Top             =   240
            Width           =   375
         End
         Begin VB.Image imgHRguy 
            Height          =   615
            Index           =   6
            Left            =   960
            ToolTipText     =   "Trasher - General Hummus/Comics"
            Top             =   0
            Width           =   735
         End
         Begin VB.Image imgHRguy 
            Height          =   375
            Index           =   4
            Left            =   600
            ToolTipText     =   "Majin BlueDragon - Coding"
            Top             =   0
            Width           =   375
         End
         Begin VB.Image imgHRguy 
            Height          =   975
            Index           =   3
            Left            =   600
            ToolTipText     =   "Kyoufu Kawa(-oneechan) - Project Management/Lead Coder/Artist/Hummus"
            Top             =   360
            Width           =   255
         End
         Begin VB.Image imgHRguy 
            Height          =   855
            Index           =   2
            Left            =   360
            ToolTipText     =   "Tauwasser - Coding/Math Issues"
            Top             =   240
            Width           =   255
         End
         Begin VB.Image imgHRguy 
            Height          =   855
            Index           =   1
            Left            =   0
            ToolTipText     =   "Hiryuu - General Hummus"
            Top             =   360
            Width           =   375
         End
         Begin VB.Image imgHRguy 
            Height          =   255
            Index           =   0
            Left            =   120
            ToolTipText     =   "Interdpth/NekoMattchan - Spy"
            Top             =   120
            Width           =   255
         End
      End
      Begin VB.TextBox txtCredits 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   6375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   149
         Text            =   "elitemap.frx":5629
         Top             =   1080
         Width           =   7455
      End
      Begin VB.Label lblVersion 
         Caption         =   "<version number here>"
         Height          =   255
         Left            =   2400
         TabIndex        =   147
         Top             =   720
         Width           =   5175
      End
      Begin VB.Image Image1 
         Height          =   870
         Left            =   120
         Picture         =   "elitemap.frx":5905
         Top             =   120
         Width           =   2700
      End
   End
   Begin VB.Label lblStampCursor 
      Caption         =   "cursor"
      Height          =   255
      Left            =   7800
      MouseIcon       =   "elitemap.frx":6C6E
      MousePointer    =   99  'Custom
      TabIndex        =   201
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblDropperCursor 
      Caption         =   "cursor"
      Height          =   255
      Left            =   7440
      MouseIcon       =   "elitemap.frx":6DC0
      MousePointer    =   99  'Custom
      TabIndex        =   200
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblPencilCursor 
      Caption         =   "cursor"
      Height          =   255
      Left            =   7080
      MouseIcon       =   "elitemap.frx":6F12
      MousePointer    =   99  'Custom
      TabIndex        =   199
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   1800
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   2040
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   2280
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   2520
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   4
      Left            =   2760
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   5
      Left            =   3000
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   6
      Left            =   3240
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   7
      Left            =   3480
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   15
      Left            =   5400
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   14
      Left            =   5160
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   13
      Left            =   4920
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   12
      Left            =   4680
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   11
      Left            =   4440
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   10
      Left            =   4200
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   9
      Left            =   3960
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape spal 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   8
      Left            =   3720
      Top             =   9480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mnuBecome 
      Caption         =   "Become..."
      Visible         =   0   'False
      Begin VB.Menu mnuBecomePerson 
         Caption         =   "Person"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBecomeTrainer 
         Caption         =   "Trainer"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuBecomeItem 
         Caption         =   "Item Ball"
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
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public endvalue As Variant
Public checkifcompa As Byte
Public checkifcompb As Byte
Public checkifpala As Byte
Public checkifpalb As Byte

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
  bLabelID As Integer
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
  b10 As Byte
  bBehavior1 As Byte
  bBehavior2 As Byte
  b13 As Byte
  b14 As Byte
  b15 As Byte
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
Private romtype As Integer

Private Sub cboConnDir_Click()
  mapConnects(vsbConn.value).wDirection = cboConnDir.ListIndex + 1
  renderconnects
  dirty = True
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
  thislevel.hSong = cboSong.ListIndex + &H15E
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
  Shape1.Visible = IIf(chkNoDraw.value = 1, False, True)
End Sub

Private Sub chkShowLabel_Click()
  thislevel.bLabelToggle = chkShowLabel.value
  dirty = True
End Sub

Private Sub cmb1_Click()
Open txtRom For Binary As #256
LOADLevels
Close #256
End Sub


Private Sub cmb2_click()
cmb2.Enabled = False
cmdLoad_Click
cmb2.Enabled = True
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
  cmb2.Enabled = True
End Sub

Private Sub chkSExits_Click()
  rendersprites
End Sub

Private Sub cmdBrowse_Click()
  exit2 = False
  tlbToolbar.Buttons(12).Enabled = False
  LoadRom False
  If txtRom <> "" Then tlbToolbar.Buttons(12).Enabled = True
  cmb2.Clear
  cmb2.Text = "Level"
  If txtRom <> "" Then Exit Sub
  cmb1.Clear
  cmb1.Text = "Bank"
End Sub

Private Sub cmdGoHome_Click()
write_bank_lev HomeLevel \ 256, HomeLevel Mod 256
End Sub

Private Sub cmdFullBRD_Click()
  AdvancedBorder.Left = Form1.Left + 5250
  AdvancedBorder.Top = Form1.Top + 7500
  AdvancedBorder.Show
  'TODO -- Add always-on-top mode for border window.
  draw_ng_border
End Sub

Private Sub cmdPanel_Click(Index As Integer)
  If cmdPanel(Index).Tag = "" Then
    cmdPanel(Index).Tag = "^_^"
    picPanel(Index).Height = Val(picPanel(Index).Tag)
  Else
    cmdPanel(Index).Tag = ""
    picPanel(Index).Height = cmdPanel(Index).Height + 1
  End If
  picPanel(1).Top = picPanel(0).Top + picPanel(0).Height
  picPanel(2).Top = picPanel(1).Top + picPanel(1).Height
  picPanel(3).Top = picPanel(2).Top + picPanel(2).Height
End Sub

Private Sub cmdPeepBecome_Click()
  PopupMenu mnuBecome
End Sub

Private Sub cmdRepointExits_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim n As Long
  If Shift = vbCtrlMask Then
    If MsgBox("Are you sure?", vbYesNo) = vbYes Then
      n = Val(InputBox("Enter new address:", , "&H" & Right("000000" & thissprite.pExits, 6)))
      i = Val(InputBox("Enter new number of exits:", , thissprite.bExits))
      thissprite.bExits = i
      thissprite.pExits = n
      vsbExits.value = 0
      vsbExits.Max = 0
      rendersprites
    End If
  Else
    MsgBox "Hold down Ctrl while clicking the Repoint button to disable failsafe.", vbInformation
  End If
End Sub

Private Sub cmdRepointSigns_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim n As Long
  If Shift = vbCtrlMask Then
    If MsgBox("Are you sure?", vbYesNo) = vbYes Then
      n = Val(InputBox("Enter new address:", , "&H" & Right("000000" & thissprite.pSigns, 6)))
      i = Val(InputBox("Enter new number of signs:", , thissprite.bSigns))
      thissprite.bSigns = i
      thissprite.pSigns = n
      vsbSigns.value = 0
      vsbSigns.Max = i
      rendersprites
    End If
  Else
    MsgBox "Hold down Ctrl while clicking the Repoint button to disable failsafe.", vbInformation
  End If
End Sub

Private Sub cmdRepointPeople_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim n As Long
  If Shift = vbCtrlMask Then
    If MsgBox("Are you sure?", vbYesNo) = vbYes Then
      n = Val(InputBox("Enter new address:", , "&H" & Right("000000" & thissprite.pPeople, 6)))
      i = Val(InputBox("Enter new number of people:", , thissprite.bPeople))
      thissprite.bPeople = i
      thissprite.pPeople = n
      vsbPeeps.value = 0
      vsbPeeps.Max = i
      rendersprites
    End If
  Else
    MsgBox "Hold down Ctrl while clicking the Repoint button to disable failsafe.", vbInformation
  End If
End Sub

Private Sub cmdRepointTraps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  Dim n As Long
  If Shift = vbCtrlMask Then
    If MsgBox("Are you sure?", vbYesNo) = vbYes Then
      n = Val(InputBox("Enter new address:", , "&H" & Right("000000" & thissprite.pTraps, 6)))
      i = Val(InputBox("Enter new number of traps:", , thissprite.bTraps))
      thissprite.bTraps = i
      thissprite.pTraps = n
      vsbTraps.value = 0
      vsbTraps.Max = i
      rendersprites
    End If
  Else
    MsgBox "Hold down Ctrl while clicking the Repoint button to disable failsafe.", vbInformation
  End If
End Sub

Private Sub cmdSaveExtern_Click()
  Dim wite As Byte
  Dim wite2 As Byte
  Dim hdr As String * 8
  hdr = "ELITEMAP"
  wite = &H80
  
  On Error GoTo Hell
  cdlCommon.flags = cdlCommonOFNHideReadOnly + cdlCommonOFNLongNames + cdlCommonOFNFileMustExist
  cdlCommon.Filter = "EliteMap exmaps (*.emap)|*.emap|All Files (*.*)|*.*"
  cdlCommon.Tag = cdlCommon.Filename
  cdlCommon.Filename = ""
  cdlCommon.ShowOpen
  Open cdlCommon.Filename For Binary As #256
  
  Put #256, , hdr
  Put #256, , wite
  
  Put #256, , lheight
  Put #256, , lwidth
  
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
  
  For i = 0 To lheight - 1
    For j = 0 To lwidth - 1
      wite = TileMap(j, i) Mod &H100
      wite2 = TileMap(j, i) \ &H100
      Put #256, , wite
      Put #256, , wite2
    Next j
  Next i
  
Hell:
  Close #256
  cdlCommon.Filename = cdlCommon.Tag

End Sub

Private Sub cmdLoadExtern_Click()
  Dim wite As Byte
  Dim wite2 As Byte
  Dim hdr As String * 8
  
  On Error GoTo Hell
  cdlCommon.flags = cdlCommonOFNHideReadOnly + cdlCommonOFNLongNames + cdlCommonOFNFileMustExist
  cdlCommon.Filter = "EliteMap exmaps (*.emap)|*.emap|All Files (*.*)|*.*"
  cdlCommon.Tag = cdlCommon.Filename
  cdlCommon.Filename = ""
  cdlCommon.ShowOpen
  Open cdlCommon.Filename For Binary As #256
  
  Get #256, , hdr
  Get #256, , wite
  If hdr = "ELITEMAP" And wite = &H80 Then
  Else
    MsgBox "Not an EliteMap ExMap file", vbExclamation, "Invalid header."
    Exit Sub
  End If
  
  Get #256, , lwidth
  Get #256, , lheight
  txtLevelWidth = Hex(lwidth)
  txtLevelHeight = Hex(lheight)
  thismap.wWidth = lwidth
  thismap.wHeight = lheight
  
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
  
  For i = 0 To lheight - 1
    For j = 0 To lwidth - 1
      Get #256, , wite
      Get #256, , wite2
      TileMap(j, i) = (CLng(wite2) * CLng(&H100)) + wite
    Next j
  Next i
  
Hell:
  Close #256
  refreshlevel
  dirty = True
  cdlCommon.Filename = cdlCommon.Tag
End Sub

Private Sub cmdLvlScript_Click()
  scriptad = thislevel.pScript
  If scriptad > -1 Then
    Open "scripted.dat" For Output As #5
    Print #5, txtRom
    Print #5, scriptad - &H8000000
    Print #5, "level"
    Close #5
    Shell "scripted.exe 1", vbNormalFocus
  End If
End Sub

Private Sub cmdLoad_Click()
  If dirty = True Then
    If MsgBox("Continue loading new map and lose your changes to this one?", vbYesNo, "Changes not saved") = vbNo Then Exit Sub
  End If

  If txtRom = "" Then
    MsgBox "Please choose a ROM!", vbInformation, "No ROM Chosen"
    Exit Sub
  End If
  
  tlbToolbar.Buttons(5).Enabled = False
  
  MousePointer = 11
  
  Picture1.Move 0, 0
  Picture1.Cls
  Picture1.BackColor = 0
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
  Text2 = ""
  get_bank_lev
  'Nine out of ten sociopaths agree
  CopyMemory TileMap(0, 0), blankmap(0, 0), 4194304
  CopyMemory tempTileMap(0, 0), blankmap(0, 0), 4194304
  point = getgbapointer((Val(txtbank) * 4) + xd)
  If point = -1 Then
    MousePointer = 0
    MsgBox "Invalid Bank"
    Close #256
    Exit Sub
  End If
  
  point = getgbapointer((Val(txtlevel) * 4) + point)
  If point = -1 Then
    MousePointer = 0
    MsgBox "Invalid Level"
    Close #256
    Exit Sub
  End If
  
   
  lpoint = point
  Get #256, point + 1, thislevel
  If NextGen = False Then
    lblLevelName = Hex(lpoint) & ": " & MapLabels(thislevel.bLabelID)
  Else
  '  'TODO -- Add nextgen label support
    lblLevelName = Hex(lpoint) & ": " & MapLabels(thislevel.bLabelID - &H58)
  End If
  lblLvlScript = Hex(GBA2PC(thislevel.pScript))
  
  shMap.Move (worldlocs(thislevel.bLabelID).bX + 1) * 8, (worldlocs(thislevel.bLabelID).bY + 2) * 8, worldlocs(thislevel.bLabelID).bW * 8, worldlocs(thislevel.bLabelID).bH * 8
  shLoc.Move (worldlocs(thislevel.bLabelID).bX + 1) * 8, (worldlocs(thislevel.bLabelID).bY + 2) * 8, worldlocs(thislevel.bLabelID).bW * 8, worldlocs(thislevel.bLabelID).bH * 8
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
  If point = -1 Then
    MousePointer = 0
    MsgBox "Invalid Map"
    Close #256
    Exit Sub
  End If
  dpoint = point
  'Debug.Print Hex(point + 1)
  Get #256, point + 1, thismap
  
  'For later use in repointing if things are too big
  allheadersize = thismap.wHeight * thismap.wWidth + 28
  If NextGen = True Then allheadersize = allheadersize + 4 + thismap.bBorderX * thismap.bBorderY
  
  'thismap.wWidth = 48
  'Debug.Print thismap.wWidth
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
  
  'Just take it from me, MC NC
  point = GBA2PC(thismap.pBorder)
  If point = -1 Then
    MousePointer = 0
    MsgBox "Invalid Border"
    Close #256
    Exit Sub
  End If
  If NextGen = False Then 'if R/S then proceed with normal routine
old_routine1:
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
    GoTo after_ng_border 'make sure there is no loop in the code
  End If
  '
  If NextGen = True Then 'If NextGen = true then
    If thismap.bBorderX = 2 Then
     If thismap.bBorderY = 2 Then GoTo old_routine1 ' proceed with old routine when 2x2 block
    End If
  
  Get #256, point, wite
  
  If thismap.bBorderX = 0 Then thismap.bBorderX = 1 'Like in the rom
  If thismap.bBorderY = 0 Then thismap.bBorderY = 1
    yind = 0
  For i = 0 To thismap.bBorderY - 1 ' Fill NextGenArray
  xind = 0
   For ii = 0 To thismap.bBorderX - 1
    Get #256, , wite
    Get #256, , wite2
    ReDim Preserve Borderitems(i & Right("00" & ii, 2))
    Borderitems(yind * thismap.bBorderX + xind) = h2d(wite2 & Right("00" & Hex(wite), 2)) 'Fill up Array with variables
    xind = xind + 1
   Next ii
  yind = yind + 1
  Next i
  End If
after_ng_border:
  Set_Border thismap.bBorderX, thismap.bBorderY
  picBorder.Refresh
  
  'You won't believe your eyes you'll go insane
  point = GBA2PC(thismap.pMap)
  If point = -1 Then
    MousePointer = 0
    MsgBox "Invalid Mapping"
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
  Close #256
  
  renderconnects
  
  'I mean what's up with that plastic plane?
  txtLevelWidth = "&H" & Hex(lwidth)
  txtLevelHeight = "&H" & Hex(lheight)
  
  tlbToolbar.Buttons("save").Enabled = True
  tlbToolbar.Buttons("copylevel").Enabled = True
  tlbToolbar.Buttons("copytileset").Enabled = True
  tlbToolbar.Buttons("clear").Enabled = True
  tlbToolbar.Buttons("resize").Enabled = True
  tlbToolbar.Buttons("viewscript").Enabled = True
  
  'You're an idiot if you disagree
  If NextGen = False Then
    Trace Hex(thislevel.hSong)
    cboLabelID.ListIndex = thislevel.bLabelID
    If thislevel.hSong < &H2FF And thislevel.hSong > &H15E Then
        cboSong.ListIndex = thislevel.hSong - &H15E
        cboSong.Visible = True
        lblSongWarning.Visible = False
    Else
        cboSong.Visible = False
        lblSongWarning.Visible = True
    End If
  Else
    'TODO -- Add Next Gen song support
    cboLabelID.ListIndex = thislevel.bLabelID - &H58
  End If
  chkShowLabel.value = thislevel.bLabelToggle
  cboWeather.ListIndex = thislevel.bWeather
  cboType.ListIndex = thislevel.bType
  vsbConn.value = 0
  vsbConn.Max = thisconnect.wConnects - 1
  vsbConn_Change
  txtConnPtr = Right("00000000" & Hex(thisconnect.pConnects), 8)
  vsbPeeps.value = 0
  vsbPeeps.Max = thissprite.bPeople - 1
  vsbPeeps_Change
  vsbExits.value = 0
  vsbExits.Max = thissprite.bExits - 1
  vsbExits_Change
  vsbTraps.value = 0
  vsbTraps.Max = thissprite.bTraps - 1
  vsbTraps_Change
  vsbSigns.value = 0
  vsbSigns.Max = thissprite.bSigns - 1
  vsbSigns_Change
  
  dirty = False

  'You gotta see Hyakugoyuichi!
  MousePointer = 0
  tlbToolbar.Buttons(5).Enabled = True

End Sub

Private Sub refreshlevel(Optional ByVal movelevel As Boolean = False)
  If movelevel = False Then
    Picture1.Move 0, 0, lwidth * &H10, lheight * &H10
    t.Move 0, 0, lwidth, lheight
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

Private Function getgbapointer(ByVal offset As Long)
  Dim a(0 To 3) As Byte
  Get #256, offset + 1, a(0)
  Get #256, offset + 2, a(1)
  Get #256, offset + 3, a(2)
  Get #256, offset + 4, a(3)
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
  Open txtRom For Binary As #256
    Do While i < &H59
      Put #256, (xm + 1) + (i * 8), worldlocs(i)
      i = i + 1
    Loop
  Close #256
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
  t = txtRom
  t2 = txtRom
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
      MsgBox "This is not a supported Pokémon Advance rom." & vbCrLf & headr & ".", vbExclamation
      exit2 = True
      Close #256
      Exit Sub
  End If
  If Roms(i).MapHeaders = 0 Or Roms(i).Maps = 0 Or Roms(i).MapLabels = 0 Then
    MsgBox "Map pointers are missing in INI. Rom cannot be used."
    lblRom = Roms(i).Code & " - " & Roms(i).Name & " (no info)"
    Close #256
    Exit Sub
  End If
  xd = getgbapointer(Roms(i).MapHeaders) 'getgbapointer(340772)
  xp = getgbapointer(Roms(i).Maps) 'getgbapointer(340588)
  xm = getgbapointer(Roms(i).MapLabels) 'getgbapointer(1032160)
  If xp = -1 Or xd = -1 Then
    MsgBox "Invalid Map pointers in INI." & vbCrLf & vbCrLf & _
           "MapHeaders points to " & Hex(xd) & "," & vbCrLf & _
           "Maps points to " & Hex(xp) & "," & vbCrLf & _
           "MapLabels points to " & Hex(xm) & "."
    Close #256
    Exit Sub
  End If
  lblRom = Roms(i).Code & " - " & Roms(i).Name
  If Roms(i).romtype > 0 Then
    NextGen = True
    cboSong.Enabled = False
    'cboLabelID.Enabled = False
    picSubEditor(1).Enabled = False
    xm = Roms(i).MapLabels
    maplabelreadNG
  Else
    NextGen = False
    cboSong.Enabled = True
    cboLabelID.Enabled = True
    picSubEditor(1).Enabled = True
    maplabelread
  End If
  
  If Roms(i).WorldMap <> "" Then
    On Error Resume Next 'should be NoMap but there's only one statement :P
    picWorldMap.Picture = LoadPicture(Roms(i).WorldMap)
  End If
  
  If Roms(i).HomeLevel > 0 Then
    HomeLevel = Roms(i).HomeLevel
  Else
    HomeLevel = &H9 'Default to LittleRoot Town
  End If
  
  LOADBanks (i)
  
  romtype = i
  
  Close #256
  Exit Sub

'KAWA - Added cancel button support =^.-=
Hell:
  txtRom.Text = ""
  Exit Sub
End Sub

Private Sub cmdWipePeople_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  If Shift = vbCtrlMask Then
    If MsgBox("Are you sure?", vbYesNo) = vbYes Then
      For i = 0 To 63
        peoples(i).b1 = 0
        peoples(i).b3 = 0
        peoples(i).b4 = 0
        peoples(i).b6 = 0
        peoples(i).b8 = 0
        peoples(i).b9 = 0
        peoples(i).b10 = 0
        peoples(i).b13 = 0
        peoples(i).b14 = 0
        peoples(i).b15 = 0
        peoples(i).b16 = 0
        peoples(i).b23 = 0
        peoples(i).b24 = 0
        peoples(i).bBehavior1 = 0
        peoples(i).bBehavior2 = 0
        peoples(i).bSpriteSet = 0
        peoples(i).bX = 0
        peoples(i).bY = 0
        peoples(i).iFlag = 0
        peoples(i).pScript = 0
      Next i
      thissprite.bPeople = 0
      vsbPeeps.value = 0
      vsbPeeps.Max = 0
      rendersprites
    End If
  Else
    MsgBox "Hold down Ctrl while clicking the Wipe button to disable failsafe.", vbInformation
  End If
End Sub

Private Sub cmdWipeExits_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  If Shift = vbCtrlMask Then
    If MsgBox("Are you sure?", vbYesNo) = vbYes Then
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
      vsbExits.value = 0
      vsbExits.Max = 0
      rendersprites
    End If
  Else
    MsgBox "Hold down Ctrl while clicking the Wipe button to disable failsafe.", vbInformation
  End If
End Sub

Private Sub cmdWipeTraps_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  If Shift = vbCtrlMask Then
    If MsgBox("Are you sure?", vbYesNo) = vbYes Then
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
      vsbTraps.value = 0
      vsbTraps.Max = 0
      rendersprites
    End If
  Else
    MsgBox "Hold down Ctrl while clicking the Wipe button to disable failsafe.", vbInformation
  End If
End Sub

Private Sub cmdWipeSigns_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  If Shift = vbCtrlMask Then
    If MsgBox("Are you sure?", vbYesNo) = vbYes Then
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
      vsbSigns.value = 0
      vsbSigns.Max = 0
      rendersprites
    End If
  Else
    MsgBox "Hold down Ctrl while clicking the Wipe button to disable failsafe.", vbInformation
  End If
End Sub

Private Sub Command4_Click()
  For i = 0 To 15
    If Val(Text2) = 0 Then
      spal(i).BackColor = palettesA(Text3, i)
    End If
  Next i
End Sub

Private Sub Command5_Click()
  DrawMap16 picBorder.hdc, Text4, Text5, 0, 0
  'DrawTile8 picBorder.hdc, Text4, Text5, 0, 0
  picBorder.Refresh
End Sub

Private Sub Command6_Click()
  picTileset.BackColor = 0
  For i = 0 To &H3FF
    X = i Mod &H20
    Y = i \ &H20
    DrawTile8 picTileset.hdc, 0, i + (Text3 * CLng(&H1000)), X * 8, Y * 8
  Next i
  picTileset.Refresh
End Sub

Private Sub cmdResize_Click()
  If lwidth = 0 Then
    MsgBox "Please load a Level!", vbInformation + vbOKOnly, "No Level Chosen"
    Exit Sub
  End If
  If lheight = 0 Then
    MsgBox "Please load a Level!", vbInformation + vbOKOnly, "No Level Chosen"
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
  
  timAutoSave.Enabled = False 'Don't want to autosave anymore.
  Trace "---------------------------------------"
  Trace "WARNING: Autosave temporarily disabled."
  Trace "---------------------------------------"
  
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
    MsgBox "Please load a Level!", vbInformation + vbOKOnly, "No Level Chosen"
    Exit Sub
  End If
  If lheight = 0 Then
    MsgBox "Please load a Level!", vbInformation + vbOKOnly, "No Level Chosen"
    Exit Sub
  End If
  If txtRom = "" Then
    MsgBox "Please choose a ROM!", vbInformation, "No ROM Chosen"
    Exit Sub
  End If
  
  searchnewplace = True
  
  newheadersize = thismap.wHeight * thismap.wWidth + 28
  If NextGen = True Then newheadersize = newheadersize + 4 + thismap.bBorderX * thismap.bBorderY
  
  If newheadersize = allheadersize Then searchnewplace = False
  
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
    MsgBox "Invalid Bank"
    Close #256
    Exit Sub
  End If
  point = getgbapointer((Val(txtlevel) * 4) + point)
  If point = -1 Then
    MousePointer = 0
    MsgBox "Invalid Level"
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
    writeoffset = Val(InputBox("There is free space at &H" & Right("00000" & Hex(writeoffset), 6) & "." & vbCrLf & "Original address was &H" & Right("00000" & Hex(oldoffset), 6) & ".", "Map is bigger than it used to be", "&H" & Right("00000" & Hex(writeoffset), 6)))
    If writeoffset = 0 Then
      MsgBox "NOO BITCH YOU CANCELED KAWA WILL KILL YOU", vbOKOnly, "Drew says" 'How dare you desecrate the Bitch Message? KAWA WILL KILL YOU!
      MousePointer = 0
      Exit Sub
    End If
    '-- end of revived code --
    
    LunarCloseFile
  End If
  
  If writeoffset = -1 Then
    MousePointer = 0
    MsgBox "Invalid Map"
    Close #256
    Exit Sub
  End If
  
  If searchnewplace = True Then GoTo putdirect
  
  Get #256, writeoffset + 1, oldmap
  
resume_searchnewplace:
  'input border
  borderoff = oldmap.pBorder - &H8000000
  If NextGen = False Then
    Seek #256, borderoff + 1
    For Y = 0 To 1
      For X = 0 To 1
        Put #256, borderoff + 1 + X + Y * 2, border(Y, X)
      Next X
    Next Y
    'KAWA - Can't fix this du du du dum wee wee can't fix this
  End If
  'input border nextgen
  If NextGen = True Then
    yind = 0
    For Y = 0 To thismap.bBorderY - 1
      xind = 0
      For X = 0 To thismap.bBorderX - 1
        c = h2d(Right("00" & Hex(Borderitems(yind * thismap.bBorderX + xind)), 2))
        d = h2d(Left(Right("0000" & Hex(Borderitems(yind * thismap.bBorderX + xind)), 4), 2))
     
        Put #256, borderoff + 1 + (yind * thismap.bBorderX + xind) * 2, c
        Put #256, borderoff + 2 + (yind * thismap.bBorderX + xind) * 2, d
        xind = xind + 1
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
 
  Close #256
  dirty = False
  MousePointer = 0
  timAutoSave.Enabled = True 'Yes, we want to autosave now.
  If timAutoSave.Tag = "nowarn" Then Exit Sub
  MsgBox "Level Saved!"
  Exit Sub

putdirect:
  'delete old completely
  Get #256, getgbapointer(((thislevel.hMap * 4) - 4) + xp) + 1, oldmap
  'delete old header
  cntx = 23
  headroff = getgbapointer(((thislevel.hMap * 4) - 4) + xp)
  If NextGen = True Then cntx = 29
  For X = 0 To cntx
    Put #256, headroff + 1 + X, 255
  Next X
  'delete old border
  borderoff = oldmap.pBorder - &H8000000
  If nexgen = False Then
    oldmap.bBorderX = 2
    oldmap.bBorderY = 2
  End If
  For Y = 0 To (oldmap.bBorderY) * 2 - 1
    For X = 0 To (oldmap.bBorderX) * 2 - 1
      Put #256, borderoff + 1 + X + Y * oldmap.bBorderX, 255
    Next X
  Next Y
  'delete old map
  mapoff = oldmap.pMap - &H8000000
  For Y = 0 To (oldmap.wHeight) * 2 - 1
    For X = 0 To (oldmap.wWidth) * 2 - 1
      Put #256, mapoff + 1 + X + Y * oldmap.wHeight, 255
    Next X
  Next Y
  thismap.pBorder = writeoffset + &H8000000
  thismap.pMap = writeoffset + &H8000000 + (thismap.bBorderX * thismap.bBorderY) * 2
  If NextGen = False Then thismap.pMap = writeoffset + 4 + &H8000000
  Put #256, thismap.pMap + (thismap.wHeight * thismap.wWidth) * 2 - &H8000000 + 1, thismap
  Put #256, getgbapointer((Val(txtlevel) * 4) + getgbapointer((Val(txtbank) * 4) + xd)) + 1, GBA2PC(thismap.pMap + (thismap.wHeight * thismap.wWidth) * 2) + &H8000000
  Put #256, thislevel.hMap * 4 - 4 + xp + 1, GBA2PC(thismap.pMap + (thismap.wHeight * thismap.wWidth) * 2) + &H8000000
  oldmap = thismap
  GoTo resume_searchnewplace
End Sub

Private Sub chkSPeople_Click()
  rendersprites
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

Private Sub Form_Load()
  SetIcon Me.hwnd, "APP", True
  lblVersion = "EliteMap version " & App.Major & "." & App.Minor & " by Kyoufu Kawa"
  
  InitDatabase
    
  '--- KAWA - Using run-time generated objects reduces elitemap.FRM's
  '---        file size from 420kb to a mere 170-something and it
  '---        still works fine =^_^=                March 11th, 2004
  For i = 1 To 63
    Load sSign(i)
    Load sTrap(i)
    Load sExit(i)
    Load sPeople(i)
  Next i
  
  cmdPanel_Click 1
  
  picTeam.Picture = LoadResPicture(1, 0)
  picWorldMap.Picture = LoadResPicture(2, 0)
  picAttributes.Picture = LoadResPicture(3, 0)
  
  'Check for the presence of any of the programs in the Launcher
  '--- HINT - To add another program, open the Toolbar's property
  '           pages, 18th button. It's a dropdown. Add another
  '           ButtonMenu object to it, disable it and set Key to
  '           the file name sans extension.
  With tlbToolbar.Buttons("launch")
    For i = 1 To .ButtonMenus.Count
      If .ButtonMenus(i).Key <> "" Then
        If Dir(.ButtonMenus(i).Key & ".exe") <> "" Then
          .ButtonMenus(i).Enabled = True
        End If
      End If
    Next i
  End With
  
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
      Case 0: attribnames(i) = "Indoor Exit"
      Case 4: attribnames(i) = "Wall"
      Case 16: attribnames(i) = "Water"
      Case &H30: attribnames(i) = "Walkthrough"
      Case &H34: attribnames(i) = "Sign"
      Case &H40: attribnames(i) = "Bridge/Wall"
      Case &HF0: attribnames(i) = "Bridge/Walk"
      Case Else: attribnames(i) = "---"
    End Select
  Next i
  selattr(0) = &H400&
  selattr(1) = &H3000&
  selattr(2) = &H1000&
  seltile(0) = &H149
  seltile(1) = 1
  seltile(2) = &H170
    
  If Int(j / 2) <> 440 Then
    MsgBox "This program has been hacked and will not run.", vbCritical, "Checksum error"
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
  picSubEditor(tabSubEditor.SelectedItem.Index - 1).Visible = True
  For i = 0 To picMainTab.UBound
    picMainTab(i).Move picMainTab(0).Left, picMainTab(0).Top
  Next i
  picMainTab(tabMain.SelectedItem.Index - 1).Visible = True
  
  'KAWA - External Overrides
  On Error GoTo NoSongs
  i = 0
  Open "songs.lst" For Input As #1
  While Not EOF(1)
    Line Input #1, s
    cboSong.List(i) = s
    i = i + 1
  Wend
  Close #1

NoSongs:

  If Command <> "" Then
    txtRom.Text = Command
    Call LoadRom(True)
    If exit2 = True Then Exit Sub
    Call cmdGoHome_Click
    Call cmdLoad_Click
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If dirty = True Then
    If MsgBox("Lose your changes to this map?", vbYesNo, "Changes not saved") = vbNo Then Exit Sub
  End If
End Sub

Private Sub hsbScroll_Change()
  On Error Resume Next
  Picture1.Move -(hsbScroll * &H10)
  Picture1.SetFocus
End Sub

Private Sub lblTilesetLoc_DblClick()
  MsgBox "Tileset A" & vbCrLf & _
         "Map16: " & Hex(thistileseta.pMap) & vbCrLf & _
         "Graphics: " & Hex(thistileseta.pGFX) & vbCrLf & _
         "Behavior: " & Hex(thistileseta.pBehavior) & vbCrLf & _
         "Animation: " & Hex(thistileseta.pAnimation) & vbCrLf & _
         " " & vbCrLf & _
         "Tileset B" & vbCrLf & _
         "Map16: " & Hex(thistilesetb.pMap) & vbCrLf & _
         "Graphics: " & Hex(thistilesetb.pGFX) & vbCrLf & _
         "Behavior: " & Hex(thistilesetb.pBehavior) & vbCrLf & _
         "Animation: " & Hex(thistilesetb.pAnimation) _
         , vbInformation, "Tileset data peek"
End Sub

Private Sub lstLabelID_Click()
  txtLabelLocX.Text = worldlocs(lstLabelID.ListIndex).bX
  txtLabelLocY.Text = worldlocs(lstLabelID.ListIndex).bY
  txtLabelLocW.Text = worldlocs(lstLabelID.ListIndex).bW
  txtLabelLocH.Text = worldlocs(lstLabelID.ListIndex).bH
  
  txtLabel.Text = MapLabels(lstLabelID.ListIndex)
  
  shMap.Move (worldlocs(lstLabelID.ListIndex).bX + 1) * 8, (worldlocs(lstLabelID.ListIndex).bY + 2) * 8, worldlocs(lstLabelID.ListIndex).bW * 8, worldlocs(lstLabelID.ListIndex).bH * 8
End Sub

Private Sub mnuBecomeItem_Click()
  If MsgBox("Are you sure?", vbYesNo) = vbNo Then Exit Sub
  peoples(vsbPeeps).bSpriteSet = &H3B
  peoples(vsbPeeps).b3 = 0
  peoples(vsbPeeps).b9 = 0
  peoples(vsbPeeps).b10 = 0
  peoples(vsbPeeps).bBehavior1 = 0
  peoples(vsbPeeps).bBehavior2 = 0
  peoples(vsbPeeps).b13 = 0
  peoples(vsbPeeps).b14 = 0
  peoples(vsbPeeps).b15 = 0
  peoples(vsbPeeps).b16 = 0
  peoples(vsbPeeps).iFlag = Val(InputBox("Enter unique item flag number, 0 - 255 inclusive.")) + &H400
  peoples(vsbPeeps).b23 = 0
  peoples(vsbPeeps).b24 = 0
  MsgBox "Done. Use ITEMBALL.RBC in Rubikon as a base for this Item Ball's new code.", vbInformation
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
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Picture1_MouseMove Button, Shift, X, Y
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim y_long As Long
  Dim zwisp As Variant
  mx = X \ 16
  my = Y \ 16
  If X > &H0 And Y > &H0 And X < ((lwidth * &H10)) And Y < ((lheight * &H10)) Then
    Label2 = "X: " & Hex(mx) & " Y: " & Hex(my)
    Label3 = Hex(TileMap((mx), (my)) Mod &H400)
    Label9 = Hex(TileMap((mx), (my)) \ &H400) & ":" & attribnames((TileMap(mx, my) \ &H400) * 4)
  End If
  at = TileMap(mx, my) \ &H400
  
  Shape1.Move mx * 16, my * 16, 16, 16
  If chkNoDraw.value = 0 Then
    Shape1.BorderColor = attribcolors(at)
  Else
    Shape1.BorderColor = 0
  End If
  Shape1.Visible = True
  
  If Shift = 0 Then GoTo Pencil
  If Shift = 1 Then GoTo Dropper
  If Shift = 2 Then GoTo Stamp
  'If chkUseStamp.value = 1 Then GoTo Stamp

Pencil:
  Picture1.MouseIcon = lblPencilCursor.MouseIcon
  If X > &H0 And Y > &H0 And X < ((lwidth * &H10)) And Y < ((lheight * &H10)) Then
    If chkNoDraw.value = 0 Then
      If Button = vbLeftButton Then
        TileMap((mx), (my)) = seltile(0) + selattr(0)
        DrawTile seltile(0), mx, my
        Picture1.Refresh
        dirty = True
      ElseIf Button = vbRightButton Then
        TileMap((mx), (my)) = seltile(1) + selattr(1)
        DrawTile seltile(1), mx, my
        Picture1.Refresh
        dirty = True
      ElseIf Button = vbMiddleButton Then
        TileMap((mx), (my)) = seltile(2) + selattr(2)
        DrawTile seltile(2), mx, my
        Picture1.Refresh
        dirty = True
      End If
    End If
  End If
  Exit Sub
Dropper:
  Picture1.MouseIcon = lblDropperCursor.MouseIcon
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
  Picture1.MouseIcon = lblStampCursor.MouseIcon
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
    Label8 = attribnames((my * &H10 + mx) * 4)
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
  If Shift = vbCtrlMask And Button = vbLeftButton Then
    scriptad = peoples(Index).pScript
    If scriptad > -1 Then
      Open "scripted.dat" For Output As #5
      Print #5, txtRom
      Print #5, scriptad - &H8000000
      Print #5, "people"
      Close #5
      Shell "scripted.exe 1", vbNormalFocus
    End If
  End If
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
  m = m & "B13: " & Hex(peoples(Index).b13) & vbCrLf
  m = m & "B14: " & Hex(peoples(Index).b14) & vbCrLf
  m = m & "B15: " & Hex(peoples(Index).b15) & vbCrLf
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
  If Shift = vbCtrlMask And Button = vbLeftButton Then
    scriptad = traps(Index).pScript
    If scriptad > -1 Then
      Open "scripted.dat" For Output As #5
      Print #5, txtRom
      Print #5, scriptad - &H8000000
      Print #5, "trap"
      Close #5
      Shell App.Path & "\scripted.exe 1", vbNormalFocus
    End If
  End If
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

Private Sub sSign_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Shift = vbCtrlMask And Button = vbLeftButton Then
    scriptad = signs(Index).pScript
    If scriptad > -1 Then
      Open "scripted.dat" For Output As #5
      Print #5, txtRom
      Print #5, scriptad - &H8000000
      Print #5, "sign"
      Close #5
      Shell "scripted.exe 1", vbNormalFocus
    End If
  End If
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

Private Sub tabMain_Click()
  On Error Resume Next
  For i = 0 To picMainTab.UBound
    picMainTab(i).Visible = False
  Next i
  picMainTab(tabMain.SelectedItem.Index - 1).Visible = True
End Sub

Private Sub tabSubEditor_Click()
  For i = 0 To picSubEditor.UBound
    picSubEditor(i).Visible = False
  Next i
  picSubEditor(tabSubEditor.SelectedItem.Index - 1).Visible = True
End Sub

Private Sub tattr_Change(Index As Integer)
  selattr(Index) = CLng(Val(tattr(Index))) * CLng(&H400)
End Sub

Private Sub timAutoSave_Timer()
  Trace "AUTOSAVE triggered"
  timAutoSave.Tag = "nowarn"
  'cmdSave_Click
End Sub

Private Sub tlbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
  If Button.Key = "browse" Then cmdBrowse_Click
  If Button.Key = "save" Then cmdSave_Click
  If Button.Key = "gohome" Then cmdGoHome_Click
  If Button.Key = "copylevel" Then cmdCopyLevel_Click
  If Button.Key = "copytileset" Then cmdCopyTiles_Click
  If Button.Key = "clear" Then cmdClear_Click
  If Button.Key = "resize" Then cmdResize_Click
  If Button.Key = "viewscript" Then cmdLvlScript_Click
  If Button.Key = "loadex" Then cmdLoadExtern_Click
  If Button.Key = "saveex" Then cmdSaveExtern_Click
  If Button.Key = "web" Then ShellExecute 0, vbNullString, "http://helmetedrodent.kickassgamers.com", vbNullString, "", 1
End Sub

Private Sub tlbToolbar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  X = Shell(ButtonMenu.Key & ".exe " & txtRom, vbNormalFocus)
End Sub

Private Sub txtConnCount_LostFocus()
  txtConnCount = Val(txtConnCount)
  If txtConnCount = 0 Then txtConnCount = 1
  thisconnect.wConnects = txtConnCount
  vsbConn.value = 0
  vsbConn.Max = thisconnect.wConnects - 1
  For i = 0 To thisconnect.wConnects - 1
    If mapConnects(i).wDirection = 0 Then mapConnects(i).wDirection = 1
  Next i
End Sub

Private Sub txtConnPtr_LostFocus()
  txtConnPtr.Text = "&H" & Right("000000" & Hex(Val(txtConnPtr)), 8)
  thisconnect.pConnects = Val(txtConnPtr) + &H8000000
End Sub

Private Sub txtCredits_DblClick()
  If Trim(txtCredits.SelText) = "dump the TILESET" Then
  
  End If
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

Private Sub txtLevelHeight_LostFocus()
  txtLevelHeight.Text = "&H" & Right("00" & Hex(Val(txtLevelHeight)), 2)
  'KAWA -- Now that the resizer has been overhauled, I'll just haveta make this instantaneous...
  thismap.wHeight = Val(txtLevelHeight)
  lheight = Val(txtLevelHeight)
End Sub

Private Sub txtLevelWidth_LostFocus()
  txtLevelWidth.Text = "&H" & Right("00" & Hex(Val(txtLevelWidth)), 2)
  'KAWA -- Now that the resizer has been overhauled, I'll just haveta make this instantaneous...
  thismap.wWidth = Val(txtLevelWidth)
  lwidth = Val(txtLevelWidth)
End Sub

Private Sub txtPeepFlag_LostFocus()
  txtPeepFlag.Text = "&H" & Right("0000" & Hex(Val(txtPeepFlag)), 4)
  peoples(vsbPeeps).iFlag = Val(txtPeepFlag)
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
  dirty = True
End Sub

Private Sub txtConnLevel_LostFocus()
  txtConnLevel.Text = "&H" & Right("0000" & Hex(Val(txtConnLevel)), 4)
  mapConnects(vsbConn.value).hLevel = Val(txtConnLevel.Text)
  renderconnects
  dirty = True
End Sub

Private Sub txtConnOffset_LostFocus()
  txtConnOffset.Text = "&H" & Right("0000" & Hex(Val(txtConnOffset)), 4)
  mapConnects(vsbConn.value).wOffset = Val(txtConnOffset.Text)
  renderconnects
  dirty = True
End Sub

Private Sub txtLabel_Change()
  MapLabels(lstLabelID.ListIndex) = txtLabel.Text
  cboLabelID.List(lstLabelID.ListIndex) = Right("00" & Hex(lstLabelID.ListIndex), 2) & ". " & MapLabels(lstLabelID.ListIndex)
  lstLabelID.List(lstLabelID.ListIndex) = Right("00" & Hex(lstLabelID.ListIndex), 2) & ". " & MapLabels(lstLabelID.ListIndex)
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
  txtTrapScript.Text = "&H" & Right("000000" & Hex(Val(txtPeepScript)), 6)
  traps(vsbTraps).pScript = Val(txtTrapScript) + &H8000000
  dirty = True
End Sub

Private Sub vsbConn_Change()
  cboConnDir.ListIndex = mapConnects(vsbConn.value).wDirection - 1
  txtConnOffset.Text = "&H" & Right("0000" & Hex(mapConnects(vsbConn.value).wOffset), 4)
  txtConnLevel.Text = "&H" & Right("0000" & Hex(mapConnects(vsbConn.value).hLevel), 4)
End Sub

Private Sub vsbTileset_Scroll()
  vsbTileset_Change
End Sub

Private Sub vsbTraps_Change()
  txtTrapScript.Text = "&H" & Right("000000" & Hex(traps(vsbTraps).pScript - &H8000000), 6)
  txtTrapX.Text = "&H" & Right("00" & Hex(traps(vsbTraps).bX), 2)
  txtTrapY.Text = "&H" & Right("00" & Hex(traps(vsbTraps).bY), 2)
  txtTrapFlag.Text = "&H" & Right("0000" & Hex(traps(vsbTraps).hFlagCheck), 4)
  txtTrapValue.Text = "&H" & Right("0000" & Hex(traps(vsbTraps).hFlagValue), 4)
End Sub

Private Sub vsbSigns_Change()
  txtSignScript.Text = "&H" & Right("000000" & Hex(signs(vsbSigns).pScript - &H8000000), 6)
  txtSignX.Text = "&H" & Right("00" & Hex(signs(vsbSigns).bX), 2)
  txtSignY.Text = "&H" & Right("00" & Hex(signs(vsbSigns).bY), 2)
End Sub

Private Sub vsbExits_Change()
  txtExitLevel.Text = "&H" & Right("0000" & Hex(exits(vsbExits).hLevel), 4)
  txtExitTarget.Text = "&H" & Right("00" & Hex(exits(vsbExits).b6), 2)
  txtExitX.Text = "&H" & Right("00" & Hex(exits(vsbExits).bX), 2)
  txtExitY.Text = "&H" & Right("00" & Hex(exits(vsbExits).bY), 2)
End Sub

Private Sub vsbPeeps_Change()
  txtPeepSprite.Text = "&H" & Right("00" & Hex(peoples(vsbPeeps).bSpriteSet), 2)
  txtPeepScript.Text = "&H" & Right("000000" & Hex(peoples(vsbPeeps).pScript - &H8000000), 6)
  cboPeepBehave.ListIndex = peoples(vsbPeeps).bBehavior1
  txtPeepFlag.Text = "&H" & Right("0000" & Hex(peoples(vsbPeeps).iFlag), 4)
  txtPeepX.Text = "&H" & Right("00" & Hex(peoples(vsbPeeps).bX), 2)
  txtPeepY.Text = "&H" & Right("00" & Hex(peoples(vsbPeeps).bY), 2)
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
    MapLabels(i) = Replace(Replace(Sapp2Asc(data), "\c\h00", ""), "\v\h08", "[TEAM]")
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
  pc = &H3B5A48 + 1 'Red only for now...
  pc = xm + 1
  Do While i < &HFF
    data = ""
    Do
      Get #256, pc, inbyte
      'Trace Hex(inbyte)
      data = data & IIf(inbyte = 255, "", Chr(inbyte))
      pc = pc + 1
    Loop Until inbyte = 255
    MapLabels(i) = Sapp2Asc(data, True)
    cboLabelID.AddItem Right("00" & Hex(i), 2) & ". " & MapLabels(i)
    lstLabelID.AddItem Right("00" & Hex(i), 2) & ". " & MapLabels(i)
    i = i + 1
  Loop
End Sub

'To get to the other side!
Private Sub rendersprites()
  For i = 0 To 63
    sPeople(i).Visible = IIf(i + 1 > thissprite.bPeople Or chkSprites.value = vbUnchecked Or chkSPeople.value = vbUnchecked, False, True)
    sExit(i).Visible = IIf(i + 1 > thissprite.bExits Or chkSprites.value = vbUnchecked Or chkSExits.value = vbUnchecked, False, True)
    sTrap(i).Visible = IIf(i + 1 > thissprite.bTraps Or chkSprites.value = vbUnchecked Or chkSTraps.value = vbUnchecked, False, True)
    sSign(i).Visible = IIf(i + 1 > thissprite.bSigns Or chkSprites.value = vbUnchecked Or chkSSigns.value = vbUnchecked, False, True)
    sPeople(i).Move peoples(i).bX * &H10, peoples(i).bY * &H10
    sExit(i).Move exits(i).bX * &H10, exits(i).bY * &H10
    sTrap(i).Move traps(i).bX * &H10, traps(i).bY * &H10
    sSign(i).Move signs(i).bX * &H10, signs(i).bY * &H10
  Next i
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
  If tattr(0) <> "&H30" Then GoTo nodebug1 'KAWA --- Made it an easter egg.
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
  
  Exit Sub
PalError:
  If BeenThereDoneThat = 0 Then
    MsgBox "Oh dear, we seem to have some problems with palette value overflows..."
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

Public Function h2d(number As String)
dec = CLng("&H" & number)
h2d = dec
End Function

Public Function LOADBanks(entry As Variant)
worldentry = entry
cmb1.Clear
cmb1.Text = "Bank"
Index = 0
For X = 0 To 1024 ' Just for security so it doesn't exploit all ;)
headers = getgbapointer(getgbapointer(Roms(entry).MapHeaders) + X * 4)
If headers = -1 Then Exit For
cmb1.AddItem "Bank &H" & Hex(Index), Index
Index = Index + 1
Next X
End Function

Public Function LOADLevels()
headers = getgbapointer(getgbapointer(Roms(worldentry).MapHeaders) + cmb1.ListIndex * 4)
On Error GoTo lastbank
headers2 = getgbapointer(getgbapointer(Roms(worldentry).MapHeaders) + (cmb1.ListIndex + 1) * 4)
GoTo go_on
lastbank:
headers2 = getgbapointer(Roms(worldentry).MapHeaders)
go_on:
cmb2.Clear
cmb2.Text = "Level"
Index = 0
For X = 0 To 1024 'security
header2 = getgbapointer(headers + X * 4)
If header2 = -1 Then Exit For
If headers + X * 4 = headers2 Then Exit For
cmb2.AddItem "Level &H" & Hex(Index), Index
Index = Index + 1
Next X
End Function

Public Sub write_bank_lev(bank As Variant, lev As Variant)
On Error GoTo Ende
cmb1.ListIndex = bank
cmb1_Click
cmb2.ListIndex = lev
Ende:
End Sub

Public Sub get_bank_lev()
txtbank = cmb1.ListIndex
txtlevel = cmb2.ListIndex
End Sub
