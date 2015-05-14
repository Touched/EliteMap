VERSION 5.00
Begin VB.Form frmThemes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select theme"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3135
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.ListBox lstTheme 
      Height          =   1620
      ItemData        =   "frmThemes.frx":0000
      Left            =   120
      List            =   "frmThemes.frx":0013
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmThemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  If (lstTheme.ListIndex + 1) * 10 <> Form1.mytheme Then
    MsgBox Replace(LoadResString(303), "[1]", App.Title)
  End If
  Form1.mytheme = (lstTheme.ListIndex + 1) * 10
  Unload Me
End Sub

Private Sub Form_Load()
  lstTheme.ListIndex = Int(Form1.mytheme / 10) - 1
  Caption = LoadResString(300)
  cmdCancel.Caption = LoadResString(301)
  cmdOK.Caption = LoadResString(302)
End Sub
