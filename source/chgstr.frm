VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick and Dirty Starter Changer"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   975
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   120
      X2              =   4080
      Y1              =   850
      Y2              =   850
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   120
      X2              =   4080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Ruby only. To allow other colors, edit this to use the PokeRoms.INI Rom Database System. I leave that to you, as a challenge..."
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Dim thefile As String
  Dim i As Long
  Dim pokename As String * 11
  Dim b As String
  Dim blah As Integer
  
  thefile = Command$ 'What was dropped in Explorer or typed in Console.
  If thefile = "" Then 'If nothing given...
    thefile = InputBox("Enter file name", , "Ruby.gba") '...ask for a filename
    If thefile = "" Then End 'If still nothing, aka Cancel, end.
  End If
  
  Open thefile For Binary As #1
  Seek #1, &H1F716C + 1 'Go to the start of the names
  For i = 0 To 410
    Get #1, , pokename 'Read a name
    b = Sapp2Asc(pokename) 'Convert it to ASCII
    While InStr(1, b, "\x"): b = Left(b, Len(b$) - 1): Wend 'Do some magic
    b = Left(b, Len(b) - 1) 'Do more magic
    Combo1.AddItem b 'Add it to all three combos
    Combo2.AddItem b
    Combo3.AddItem b
  Next i
  Seek #1, &H3F76C4 + 1 'Go to the starter info
  Get #1, , blah 'Get an integer/word
  Combo1.ListIndex = blah 'Set the combo's selection to said integer
  Get #1, , blah 'Rinse
  Combo2.ListIndex = blah
  Get #1, , blah 'Repeat
  Combo3.ListIndex = blah
End Sub

Private Sub Command1_Click() 'When we click the button...
  Seek #1, &H3F76C4 + 1 '...we return to the starter info...
  Put #1, , CInt(Combo1.ListIndex) '...convert our ListIndex to an Integer...
  Put #1, , CInt(Combo2.ListIndex) '...and write it back.
  Put #1, , CInt(Combo3.ListIndex) 'Rinse, repeat.
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Close #1 'But Mike, that's amazing!
End Sub

