VERSION 5.00
Begin VB.Form frmReference 
   BackColor       =   &H00E6BBBA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Command Reference"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReference.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   346
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShowList 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   9
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmdFocusHider 
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblParams 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<params>"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00D69896&
      X1              =   8
      X2              =   336
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<description>"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCommand 
      BackStyle       =   0  'Transparent
      Caption         =   "<cmd>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F8DDDC&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   1815
      Left            =   120
      Top             =   120
      Width           =   4935
   End
   Begin VB.Menu mnuList 
      Caption         =   "list"
      Visible         =   0   'False
      Begin VB.Menu mnuCommand 
         Caption         =   "template"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ResizeMe()
  Shape1.Width = ScaleWidth - 16
  Line2.Y1 = lblDescription.Top + lblDescription.Height + 8
  Line2.Y2 = Line2.Y1
  Line2.X2 = Shape1.Width + 8
  cmdShowList.Left = ScaleWidth - 16 - cmdShowList.Width
  cmdFocusHider.Left = cmdShowList.Left
  lblCommand.Width = Shape1.Width - 16
  lblDescription.Width = Shape1.Width - 16
  lblParams.Width = Shape1.Width - 16
  lblParams.Top = Line2.Y1 + 8
  Shape1.Height = lblParams.Top + lblParams.Height + 8
  Height = (Shape1.Height + 40) * Screen.TwipsPerPixelY
End Sub

Private Sub cmdShowList_Click()
  cmdFocusHider.SetFocus
  PopupMenu mnuList
End Sub

Private Sub Form_Load()
  Dim i As Integer
  mnuCommand(0).Caption = RCD.RubiCommands(0).Keyword
  For i = 1 To 255
    If RCD.RubiCommands(i).Keyword <> "" And Left(RCD.RubiCommands(i).Keyword, 4) <> "#raw" Then
      Load mnuCommand(i)
      mnuCommand(i).Caption = RCD.RubiCommands(i).Keyword
    End If
  Next i
End Sub

Private Sub mnuCommand_Click(Index As Integer)
  Dim found As Boolean
  Dim i As Integer, j As Integer
  For i = 0 To 255
    If RCD.RubiCommands(i).Keyword = mnuCommand(Index).Caption Then
      found = True
      Top = frmRubIDE.Top + frmRubIDE.Height
      Left = frmRubIDE.Left
      Width = frmRubIDE.Width
      lblCommand = RCD.RubiCommands(i).Keyword
      lblDescription = RCD.RubiCommands(i).Description
      If RCD.RubiCommands(i).ParamCount = 0 Then
        lblParams = "No parameters required."
      Else
        lblParams = "Parameters:"
        For j = 0 To RCD.RubiCommands(i).ParamCount - 1
          lblParams = frmReference.lblParams & vbCrLf & _
                    RCD.GetSizeName(RCD.RubiParameters(i, j).Size) & " - " & _
                    RCD.RubiParameters(i, j).Description
        Next j
      End If
      ResizeMe
    End If
  Next i
End Sub
