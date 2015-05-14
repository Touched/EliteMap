VERSION 5.00
Begin VB.Form frmSecret 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "You found a secret!"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5220
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblSecret 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   2415
   End
   Begin VB.Image imgSecret 
      BorderStyle     =   1  'Fixed Single
      Height          =   3090
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5220
   End
End
Attribute VB_Name = "frmSecret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
  Unload Me
End Sub
Private Sub imgSecret_Click()
  Unload Me
End Sub

Private Sub lblSecret_Click()
  Unload Me
End Sub
