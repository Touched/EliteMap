VERSION 5.00
Begin VB.Form Resize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resize Map"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   Icon            =   "Resize.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm4 
      Caption         =   "Border : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   3615
      Begin VB.TextBox hdb3 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox hdb4 
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Text            =   "0"
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblhbnow 
         Caption         =   "0"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lbl2 
         Caption         =   "Width :                       Height :"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label lblwbnow 
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Old width:                   Old height :"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame frm3 
      Caption         =   "Map : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      Begin VB.TextBox hdb1 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox hdb2 
         Height          =   285
         Left            =   2880
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lbl2 
         Caption         =   "Width :                       Height :"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lblhmnow 
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblwmnow 
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lbl1 
         Caption         =   "Old width:                   Old height :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "Resize"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "Resize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd1_Click()
  Form1.Enabled = True
  'when empty then exit
  If hdb1.Text = "" Then GoTo no_size
  If hdb2.Text = "" Then GoTo no_size
  If hdb3.Text = "" Then GoTo no_size
  If hdb4.Text = "" Then GoTo no_size
  'when size = 0 then replace it with 1
  If hdb1.Text = "0" Then hdb1.Text = "1"
  If hdb2.Text = "0" Then hdb2.Text = "1"
  If hdb3.Text = "0" Then hdb3.Text = "1"
  If hdb4.Text = "0" Then hdb4.Text = "1"
  'write stuff into worldvariables
  modNextGenBorder.noresize = True
  newwm = Val(hdb1.Text)
  newhm = Val(hdb2.Text)
  If NextGen = False Then GoTo end3
  Unload AdvancedBorder 'hey - no instant preview
  newwb = Val(hdb3.Text)
  newhb = Val(hdb4.Text)
  'border_sizeX = newwb
  'border_sizeY = newhb
  'ReDim Preserve Borderitems(newwb * newhb)
end3:
  Unload Me
  Exit Sub
no_size:
  MsgBox "One ore more resolutions have not been entered. Aborting!", vbOKOnly, "Error…"
End Sub

Private Sub Command1_Click()
  Form1.Enabled = True
  If noresize = True Then Exit Sub
  modNextGenBorder.noresize = False
  Unload Me
End Sub

Private Sub Form_Load()
  noresize = False
  Form1.Enabled = False
  lblwmnow.Caption = newwm
  lblhmnow.Caption = newhm
  hdb1.Text = newwm
  hdb2.Text = newhm
  If NextGen = False Then
  frm4.Enabled = False
  Exit Sub
  End If
  frm4.Enabled = True
  lblwbnow.Caption = thismap.bBorderX
  lblhbnow.Caption = thismap.bBorderY
  hdb3.Text = thismap.bBorderX
  hdb4.Text = thismap.bBorderY
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Command1_Click
End Sub
