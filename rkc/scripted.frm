VERSION 5.00
Begin VB.Form frmRubIDE 
   Caption         =   "Rubikon ScriptEd"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "scripted.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   344
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   481
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFile 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   60
      Width           =   2655
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3495
      TabIndex        =   3
      Top             =   60
      Width           =   255
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Top             =   45
      Width           =   735
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Height          =   285
      Left            =   5520
      TabIndex        =   6
      Top             =   45
      Width           =   735
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      HideSelection   =   0   'False
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.CheckBox chkJapanese 
      Caption         =   "Jap"
      Height          =   285
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   45
      Width           =   495
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   45
      Width           =   735
   End
   Begin VB.Image imgBarGripR 
      Height          =   360
      Left            =   7080
      Picture         =   "scripted.frx":628A
      Top             =   0
      Width           =   105
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00D69896&
      Height          =   285
      Left            =   780
      Top             =   45
      Width           =   2985
   End
   Begin VB.Image imgBarGripL 
      Height          =   360
      Left            =   0
      Picture         =   "scripted.frx":678C
      Top             =   0
      Width           =   105
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "&File"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   60
      Width           =   495
   End
   Begin VB.Image imgBar 
      Height          =   375
      Left            =   120
      Picture         =   "scripted.frx":6C8E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmRubIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private TehFile As String
Private TehOffset As Long

Private Ryan As Integer
Private MyCDlg As cCommonDialog

Private movelabels(0 To &HFF) As String
Private strings(0 To &HFF) As Long
Private scripttype As String
Private numstrings As Long
Private snips(0 To &HFF) As Long
Private numsnips As Long
Private moves(0 To &HFF) As Long
Private nummoves As Long

Private Sub cmdBrowse_Click()
  Dim NewFile As String
  If Not MyCDlg.VBGetOpenFileName(NewFile, , , , , , "RubiKode (*.rbc)|*.rbc|RKC Headers (*.rbh)|*.rbh|GBA roms (*.gba)|*.gba|Binaries (*.bin)|*.bin") Then Exit Sub
  txtFile = MyCDlg.VBGetFileTitle(NewFile)
  If Right(LCase(txtFile), 3) = "gba" Or Right(LCase(txtFile), 3) = "bin" Then
    txtFile = txtFile & ":000000"
    txtFile.SelStart = Len(txtFile) - 6
    txtFile.SelLength = 6
    txtFile.SetFocus
  Else
    cmdLoad_Click
  End If
End Sub

Private Sub cmdCompile_Click()
  Dim NewFile As String
  On Error GoTo NoLastFile
  Open "lastfile.txt" For Input As #1
    Input #1, NewFile
  Close #1
  On Error GoTo 0
  If Not MyCDlg.VBGetSaveFileName(NewFile, , , "GBA roms (*.gba)|*.gba|Binaries (*.bin)|*.bin|All (*.*)|*.*") Then Exit Sub
  NewFile = MyCDlg.VBGetFileTitle(NewFile)
  Open "rubide.rbc" For Output As #2
  Print #2, "'RubIDE temporary file for recompilation."
  If Not InStr(txtCode, "#include ""std.rbh""") Then Print #2, "#include ""std.rbh"""
  Print #2, txtCode
  Close #2
  Shell "rkc.exe /o " & NewFile & " rubide.rbc"
NoLastFile:
  Open "lastfile.txt" For Output As #3
  Print #3, NewFile
  Close #3
End Sub

Private Sub cmdLoad_Click()
  Dim wite As Byte
  Dim wite2 As Byte
  Dim wite3 As Byte
  Dim wite4 As Byte
  Dim asprin As Integer
  Dim asprin2 As Integer
  Dim asprin3 As Integer
  Dim asprin4 As Integer
  Dim wow As Long
  Dim wow2 As Long
  Dim wow3 As Long
  Dim wow4 As Long
  Dim ns As String
  Dim offset As Long
  Dim break As Boolean
  Dim C As Integer
  Dim d As Integer
  Dim jump As Integer
  Dim stringcmd As String
  Dim stringarg As String
  Dim foundcmd As Boolean
  
  If txtFile = "" Then Exit Sub
  
  MousePointer = 11
  LockWindowUpdate hwnd
  
  TehFile = txtFile
  TehOffset = 1
  If InStr(LCase(txtFile), ".gba:") Or InStr(LCase(txtFile), ".bin:") Then
    TehFile = Left(txtFile, InStr(LCase(txtFile), ".") + 3)
    TehOffset = Val("&H" & Mid(txtFile, InStr(LCase(txtFile), ".") + 5)) + 1
  End If
  'MsgBox TehFile & vbCrLf & Hex(TehOffset)
  
  txtCode = ""
  
  If LCase(Right(TehFile, 3)) = "gba" Or LCase(Right(TehFile, 3)) = "bin" Then
  'It's a binary file that must be decompiled.
  
  Open TehFile For Binary As #1
  
  numstrings = 0
  nummoves = 0
  snips(0) = TehOffset
  numsnips = 1
  Ryan = 0
  
  Seek #1, TehOffset + 1
  Do While Ryan < numsnips
    txtCode = txtCode & "'-----------------------" & vbCrLf
    
    offset = snips(Ryan)
    break = False
    If offset <= 0 Then Exit Do
    txtCode = txtCode & "#org 0x" & Hex(offset - 1) & vbCrLf
    Seek #1, offset
    Do
      If offset <= 0 Then Exit Do
      'Seek #1, offset
      Get #1, , wite
      jump = 0
      stringcmd = ""
      stringarg = ""
      
      If wite = 2 Then
        txtCode = txtCode & "end" & vbCrLf
        break = True
      ElseIf wite = 3 Then
        txtCode = txtCode & "return" & vbCrLf
        break = True
      ElseIf wite = 4 Then
        txtCode = txtCode & "call"
        Get #1, , wow2
        txtCode = txtCode & " 0x" & Hex(wow2) & " "
        jump = jump + 4
        If notinsnips(wow2) = True Then
          snips(numsnips) = wow2
          numsnips = numsnips + 1
        End If
      ElseIf wite = 5 Then
        txtCode = txtCode & "goto"
        Get #1, , wow2
        txtCode = txtCode & " 0x" & Hex(wow2) & " "
        jump = jump + 4
        If notinsnips(wow2) = True Then
          snips(numsnips) = wow2
          numsnips = numsnips + 1
        End If
      ElseIf wite = 6 Then 'jumping version
        Get #1, , wite2
        Get #1, , wow2
        txtCode = txtCode & "if 0x" & Hex(wite2) & " jump 0x" & Hex(wow2 - IIf(wow2 > &H8000000, &H8000000, 0))
        jump = jump + 4
        If notinsnips(wow2 - IIf(wow2 > &H8000000, &H8000000, 0) + 1) = True Then
          snips(numsnips) = wow2 - IIf(wow2 > &H8000000, &H8000000, 0) + 1
          numsnips = numsnips + 1
        End If
      ElseIf wite = 7 Then 'calling version
        Get #1, , wite2
        Get #1, , wow2
        txtCode = txtCode & "if 0x" & Hex(wite2) & " call 0x" & Hex(wow2 - IIf(wow2 > &H8000000, &H8000000, 0))
        jump = jump + 4
        If notinsnips(wow2 - IIf(wow2 > &H8000000, &H8000000, 0) + 1) = True Then
          snips(numsnips) = wow2 - IIf(wow2 > &H8000000, &H8000000, 0) + 1
          numsnips = numsnips + 1
        End If
      ElseIf wite = &HF Then
        Get #1, , wite2
        Get #1, , wow2
        If wite2 Then
          txtCode = txtCode & "loadptr 0x" & Hex(wite2) & " " & Hex(wow2)
        Else
          txtCode = txtCode & "msgbox 0x" & Hex(wow2) & " '""FAB0BABE" & Hex(wow2 - (IIf(wow2 > &H8000000, &H8000000, 0))) & """"
          If notinstrings(wow2) = True Then
            strings(numstrings) = wow2 - (IIf(wow2 > &H8000000, &H8000000, 0))
            numstrings = numstrings + 1
          End If
        End If
  
  
      Else
        foundcmd = False
        For C = 0 To 255
          If RCD.RubiCommands(C).Bytecode = wite Then
            txtCode = txtCode & RCD.RubiCommands(C).Keyword
            jump = 1
            If RCD.RubiCommands(C).ParamCount > 0 Then
              For d = 0 To RCD.RubiCommands(C).ParamCount - 1
                Select Case RCD.RubiParameters(C, d).Size
                  Case 0 'Byte
                    Get #1, , wite2
                    txtCode = txtCode & " 0x" & Hex(wite2)
                    jump = jump + 1
                  Case 1 'Word
                    Get #1, , asprin2
                    txtCode = txtCode & " 0x" & Hex(asprin2)
                    jump = jump + 2
                  Case 2 'DWord
                    Get #1, , wow2
                    txtCode = txtCode & " 0x" & Hex(wow2)
                    jump = jump + 4
                  Case 3 'Pointer
                    Get #1, , wow2
                    txtCode = txtCode & " 0x" & Hex(wow2)
                    jump = jump + 4
                End Select
              Next d
            End If
            foundcmd = True
            Exit For
          End If
        Next C
      End If
      
      txtCode = txtCode & vbCrLf
      DoEvents
      
      offset = offset + jump
    Loop Until break = True Or offset > LOF(1)
    Ryan = Ryan + 1
  Loop
  
  txtCode = Replace(txtCode, "0x800D", "LASTRESULT")
  
  If numstrings > 0 Then
    txtCode = txtCode & vbCrLf & "'---------"
    txtCode = txtCode & vbCrLf & "' Strings"
    txtCode = txtCode & vbCrLf & "'---------" & vbCrLf

    For Ryan = 0 To numstrings - 1
      ns = ""
      txtCode = txtCode & "#org 0x" & Hex(strings(Ryan)) & vbCrLf
      wow = strings(Ryan) - (IIf(strings(Ryan) > &H8000000, &H8000000, 0))
      wow2 = wow
      Do
        If wow <= 0 Then Exit Do
        Get #1, wow + 1, wite
        ns = ns & IIf(wite = 255, "", Chr(wite))
        wow = wow + 1
      Loop Until wite = 255
      txtCode = txtCode & "= " & Sapp2Asc(ns, IIf(chkJapanese.Value = 1, True, False)) & vbCrLf
      txtCode = Replace(txtCode, "FAB0BABE" & Hex(strings(Ryan)), Left(Sapp2Asc(ns, IIf(chkJapanese.Value = 1, True, False)), 20) & IIf(Len(strings(Ryan)) < 20, "...", ""))
    Next Ryan
  End If
  
  Close #1
  MousePointer = 0
  LockWindowUpdate 0
  On Error Resume Next
  txtCode.SetFocus
  Exit Sub
  
  Else
    'It's a text file that can just be loaded as-is
    On Error Resume Next
    Open TehFile For Input As #1
    Do
      Line Input #1, ns
      txtCode = txtCode & ns & vbCrLf
    Loop Until EOF(1)
    Close #1
  
  End If
  
  MousePointer = 0
  LockWindowUpdate 0
  On Error Resume Next
  txtCode.SetFocus
End Sub

Private Function notinsnips(ByVal address As Long) As Boolean
  Dim i As Integer
  notinsnips = True
  For i = 0 To numsnips - 1
    If snips(i) = address Then notinsnips = False
  Next i
End Function
Private Function notinstrings(ByVal address As Long) As Boolean
  Dim i As Integer
  notinstrings = True
  For i = 0 To numstrings - 1
    If strings(i) = address Then notinstrings = False
  Next i
End Function
Private Function notinmoves(ByVal address As Long) As Boolean
  Dim i As Integer
  notinmoves = True
  For i = 0 To nummoves - 1
    If moves(i) = address Then notinmoves = False
  Next i
End Function

Private Sub cmdSave_Click()
  Dim NewFile As String
  If Not MyCDlg.VBGetSaveFileName(NewFile, , , "RubiKode (*.rbc)|*.rbc|RKC Headers (*.rbh)|*.rbh|All (*.*)|*.*") Then Exit Sub
  Open NewFile For Output As #1
  Print #1, txtCode
  Close #1
End Sub

Private Sub Form_Load()
  Dim retro As String
  Dim C As Long
  'modLinedTextBox.ShowLines txtCode, True, 3
  
  Set MyCDlg = New cCommonDialog
    
  RCD.LoadCommands
  
  For C = 0 To 255 'ugly hack
    If RCD.RubiCommands(C).Keyword = "" Then
      RCD.RubiCommands(C).Keyword = "#raw 0x" & Hex(C)
    End If
  Next C '/ugly hack
    
  If Dir("rkc.exe") = "" Then
    cmdCompile.Enabled = False
  End If
  
  For Ryan = 0 To &HFF
    movelabels(Ryan) = "mov" & Hex(Ryan)
  Next Ryan
  movelabels(&H0) = "Down0"
  movelabels(&H1) = "Up0"
  movelabels(&H2) = "Left0"
  movelabels(&H3) = "Right0"
  movelabels(&H4) = "Down1"
  movelabels(&H5) = "Up1"
  movelabels(&H6) = "Left1"
  movelabels(&H7) = "Right1"
  movelabels(&H8) = "Down2"
  movelabels(&H9) = "Up2"
  movelabels(&HA) = "Left2"
  movelabels(&HB) = "Right2"
  movelabels(&HC) = "HopTileDown"
  movelabels(&HD) = "HopTileUp"
  movelabels(&HE) = "HopTileLeft"
  movelabels(&HF) = "HopTileRight"
  movelabels(&H10) = "Delay0"
  movelabels(&H11) = "Delay1"
  movelabels(&H12) = "Delay2"
  movelabels(&H13) = "Delay3"
  movelabels(&H14) = "Delay4"
  movelabels(&H15) = "Down3"
  movelabels(&H16) = "Up3"
  movelabels(&H17) = "Left3"
  movelabels(&H18) = "Right3"
  movelabels(&H19) = "StDown1"
  movelabels(&H1A) = "StUp1"
  movelabels(&H1B) = "StLeft1"
  movelabels(&H1C) = "StRight1"
  movelabels(&H1D) = "StDown2"
  movelabels(&H1E) = "StUp2"
  movelabels(&H1F) = "StLeft2"
  movelabels(&H20) = "StRight2"
  movelabels(&H21) = "StDown3"
  movelabels(&H22) = "StUp3"
  movelabels(&H23) = "StLeft3"
  movelabels(&H24) = "StRight3"
  movelabels(&H25) = "StDown4"
  movelabels(&H26) = "StUp4"
  movelabels(&H27) = "StLeft4"
  movelabels(&H28) = "StRight4"
  movelabels(&H29) = "Down3"
  movelabels(&H2A) = "Up3"
  movelabels(&H2B) = "Left3"
  movelabels(&H2C) = "Right3"
  movelabels(&H2D) = "Down4"
  movelabels(&H2E) = "Up4"
  movelabels(&H2F) = "Left4"
  movelabels(&H30) = "Right4"
  movelabels(&H31) = "SlideFaceDown"
  movelabels(&H32) = "SlideFaceUp"
  movelabels(&H33) = "SlideFaceLeft"
  movelabels(&H34) = "SlideFaceRight"
  movelabels(&H35) = "RunDown"
  movelabels(&H36) = "RunUp"
  movelabels(&H37) = "RunLeft"
  movelabels(&H38) = "RunRight"
  movelabels(&H39) = "St0"
  movelabels(&H3A) = "HighHopDown"
  movelabels(&H3B) = "HighHopUp"
  movelabels(&H3C) = "HighHopLeft"
  movelabels(&H3D) = "HighHopRight"
  movelabels(&H3E) = "Up0A"
  movelabels(&H3F) = "Down0A"
  movelabels(&H40) = "mov40"
  movelabels(&H41) = "mov41"
  movelabels(&H42) = "JumpDown"
  movelabels(&H43) = "JumpUp"
  movelabels(&H44) = "JumpLeft"
  movelabels(&H45) = "JumpRight"
  movelabels(&H46) = "HopDown"
  movelabels(&H47) = "HopUp"
  movelabels(&H48) = "HopLeft"
  movelabels(&H49) = "HopRight"
  movelabels(&H4A) = "HopDown180"
  movelabels(&H4B) = "HopUp180"
  movelabels(&H4C) = "HopLeft180"
  movelabels(&H4D) = "HopRight180"
  movelabels(&H4E) = "Down0B"
  movelabels(&H4F) = "StRun"
  movelabels(&H50) = "mov50"
  movelabels(&H51) = "mov51"
  movelabels(&H52) = "mov52"
  movelabels(&H53) = "mov53"
  movelabels(&H54) = "Hide"
  movelabels(&H55) = "Show"
  movelabels(&H56) = "Alert"
  movelabels(&H57) = "Question"
  movelabels(&H58) = "Love"
  movelabels(&H59) = "mov59"
  movelabels(&H5A) = "Pokeball"
  movelabels(&H5B) = "mov5B"
  movelabels(&H5C) = "mov5C"
  movelabels(&H5D) = "mov5D"
  movelabels(&H5E) = "mov5E"
  movelabels(&H5F) = "mov5F"
  movelabels(&H60) = "mov60"
  movelabels(&H61) = "mov61"
  movelabels(&H63) = "Up0B"
  movelabels(&H64) = "mov64"
  movelabels(&H65) = "Right0A"
  movelabels(&H66) = "RunStopLoopDown"
  movelabels(&H67) = "RunStopLoopUp"
  movelabels(&H68) = "RunStopLoopLeft"
  movelabels(&H69) = "RunStopLoopRight"
  movelabels(&H6A) = "StDown1i"
  movelabels(&H6B) = "StUp1i"
  movelabels(&H6C) = "StLeft1i"
  movelabels(&H6D) = "StRight1i"
  movelabels(&H6E) = "StDown5"
  movelabels(&H6F) = "StUp5"
  movelabels(&H70) = "StLeft5"
  movelabels(&H71) = "StRight5"
  movelabels(&H72) = "Down15"
  movelabels(&H73) = "Up15"
  movelabels(&H74) = "Left15"
  movelabels(&H75) = "Right15"
  movelabels(&H76) = "mov76"
  movelabels(&H77) = "mov77"
  movelabels(&H78) = "mov78"
  movelabels(&H79) = "mov79"
  movelabels(&H7A) = "Down6"
  movelabels(&H7B) = "Up6"
  movelabels(&H7C) = "Left6"
  movelabels(&H7D) = "Right6"
  movelabels(&H7E) = "RunDown2"
  movelabels(&H7F) = "RunUp2"
  movelabels(&H80) = "RunLeft2"
  movelabels(&H81) = "RunRight2"
  movelabels(&H82) = "Down7"
  movelabels(&H83) = "Up7"
  movelabels(&H84) = "Left7"
  movelabels(&H85) = "Right7"
  movelabels(&H86) = "IceSlideDown"
  movelabels(&H87) = "IceSlideUp"
  movelabels(&H88) = "IceSlideLeft"
  movelabels(&H89) = "IceSlideRight"
  movelabels(&HFE) = "Exit"
    
  If Trim(Command$) = "" Then Exit Sub
    
  txtFile = MyCDlg.VBGetFileTitle(Command$)
  
  'Backwards compatibility for EM's before than 3.6
  If Command$ = "1" Then
    Open "scripted.dat" For Input As #1
    Line Input #1, retro
    txtFile = MyCDlg.VBGetFileTitle(retro) & ":"
    Line Input #1, retro
    txtFile = txtFile & Hex(Val(retro))
    Close #1
    Kill "scripted.dat"
    Show
    Refresh
    DoEvents
  End If

  cmdLoad_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  'Unload frmReference
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  LockWindowUpdate hwnd
  imgBar.Width = ScaleWidth - (imgBarGripR.Width * 2)
  imgBarGripR.Left = ScaleWidth - imgBarGripR.Width
  txtCode.Width = ScaleWidth
  txtCode.Height = ScaleHeight - txtCode.Top
  chkJapanese.Left = ScaleWidth - chkJapanese.Width - 16
  cmdCompile.Left = chkJapanese.Left - cmdCompile.Width - 8
  cmdSave.Left = cmdCompile.Left - cmdSave.Width - 8
  cmdLoad.Left = cmdSave.Left - cmdLoad.Width - 8
  Shape1.Width = cmdLoad.Left - cmdLoad.Width - 8
  cmdBrowse.Left = cmdLoad.Left - cmdBrowse.Width - 6
  txtFile.Width = Shape1.Width - cmdBrowse.Width - 8
  LockWindowUpdate 0
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
'  If KeyCode <> 112 Then Exit Sub
'  KeyCode = 0
'  Dim i As Integer
'  Dim j As Integer
'  Dim found As Boolean
'  Dim word As String
'  If txtCode.SelLength Then
'    word = Trim(txtCode.SelText)
'    For i = 0 To 255
'      If RCD.RubiCommands(i).Keyword = word Then
'        found = True
'        frmReference.Top = Top + Height
'        frmReference.Left = Left
'        frmReference.Width = Width
'        frmReference.lblCommand = RCD.RubiCommands(i).Keyword
'        frmReference.lblDescription = RCD.RubiCommands(i).Description
'        If RCD.RubiCommands(i).ParamCount = 0 Then
'          frmReference.lblParams = "No parameters required."
'        Else
'          frmReference.lblParams = "Parameters:"
'          For j = 0 To RCD.RubiCommands(i).ParamCount - 1
'            frmReference.lblParams = frmReference.lblParams & vbCrLf & _
'                      RCD.GetSizeName(RCD.RubiParameters(i, j).Size) & " - " & _
'                      RCD.RubiParameters(i, j).Description
'          Next j
'        End If
'        frmReference.ResizeMe
'        frmReference.Show
'        ZOrder
'      End If
'    Next i
'  End If
End Sub


Private Sub txtFile_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = 0
    cmdLoad_Click
  End If
End Sub
