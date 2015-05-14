VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRubikon 
   Caption         =   "Rubikon"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "rubikon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox filRoms 
      Height          =   675
      Left            =   6120
      Pattern         =   "*.gba;*.agb"
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6120
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Ready..."
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   6720
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "rbc"
      Filter          =   "Rubikon code (*.rbc)|*.rbc|Rubikon header (*.rbh)|*.rbh|All files|*.*"
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "save"
                  Text            =   "Save"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "saveas"
                  Text            =   "Save As..."
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Compile"
            Object.ToolTipText     =   "Compile"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ComboBox txtRom 
         Height          =   315
         Left            =   3120
         TabIndex        =   4
         Top             =   0
         Width           =   3615
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6720
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rubikon.frx":030A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rubikon.frx":041C
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rubikon.frx":052E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rubikon.frx":0640
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rubikon.frx":0752
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rubikon.frx":0864
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "rubikon.frx":0976
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRubikon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Define
  Symbol As String
  Value As Long
End Type
Private Defines(2048) As Define
Private DefCnt As Integer
Private AutoBank As Boolean
Private Dirty As Boolean
Private ln As Long

Private Sub Form_Load()
  Dim ff As Integer
  Dim a As String
  On Error GoTo Barkness
  ff = FreeFile
  Open "std.rbh" For Input As ff
  Close ff
  
  SetIcon Me.hWnd, "AAA", True
  
  For ff = 0 To filRoms.ListCount - 1
    txtRom.AddItem filRoms.List(ff)
  Next ff
  If txtRom.ListCount > 0 Then
    txtRom.ListIndex = 0
  Else
    txtRom.Text = ""
  End If
   
  Dim oneline As String
  If Command <> "" Then
    txtCode.Tag = Command
    txtCode.Text = ""
    Open Command For Input As #1
    While Not EOF(1)
      Line Input #1, oneline
      txtCode.Text = txtCode.Text & oneline & vbCrLf
    Wend
    Close #1
    Dirty = False
    Caption = "Rubikon - " & Command
    If LCase(Right(txtCode.Tag, 3)) = "rbc" Then
      Toolbar1.Buttons("Compile").Enabled = True
    Else
      Toolbar1.Buttons("Compile").Enabled = False
    End If
  End If
  
  Exit Sub
Barkness:
  MsgBox "std.rbh not found.", vbExclamation
  Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim ff As Integer
  If Dirty = True Then
    If MsgBox("You haven't saved yet. Really continue?", vbYesNo) = vbNo Then
      Cancel = -1
      Exit Sub
    End If
  End If
End Sub

Private Sub Form_Resize()
  txtRom.Left = Width - txtRom.Width - 200
  txtCode.Width = ScaleWidth
  txtCode.Height = ScaleHeight - Toolbar1.Height - StatusBar1.Height
End Sub

Private Sub StatusBar1_DblClick()
  Text1.Width = txtCode.Width
  Text1.Height = txtCode.Height
  Text1.Visible = True
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  Dim oneline As String
  On Error GoTo Hell
  If ButtonMenu.Key = "save" Then
    Open cdl.FileName For Output As #1
      Print #1, txtCode.Text
    Close #1
    Dirty = False
    txtCode.Tag = cdl.FileTitle
    Caption = "Rubikon - " & cdl.FileTitle
  ElseIf ButtonMenu.Key = "saveas" Then
    cdl.ShowSave
    Open cdl.FileName For Output As #1
      Print #1, txtCode.Text
    Close #1
    Dirty = False
    txtCode.Tag = cdl.FileTitle
    Caption = "Rubikon - " & cdl.FileTitle
  End If
Hell:
  Resume Next
End Sub

Private Sub txtCode_Change()
  Dirty = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim oneline As String
  On Error GoTo Hell
  Select Case Button.Key
    Case "New"
      If Dirty = True Then
        If MsgBox("You haven't saved yet. Really continue?", vbYesNo) = vbNo Then Exit Sub
      End If
      txtCode.Tag = ""
      txtCode.Text = ""
      Dirty = False
      Caption = "Rubikon - Untitled"
    Case "Open"
      If Dirty = True Then
        If MsgBox("You haven't saved yet. Really continue?", vbYesNo) = vbNo Then Exit Sub
      End If
      cdl.ShowOpen
      txtCode.Tag = cdl.FileTitle
      txtCode.Text = ""
      Open cdl.FileName For Input As #1
      While Not EOF(1)
        Line Input #1, oneline
        txtCode.Text = txtCode.Text & oneline & vbCrLf
      Wend
      Close #1
      Dirty = False
      Caption = "Rubikon - " & cdl.FileTitle
      If LCase(Right(txtCode.Tag, 3)) = "rbc" Then
        Toolbar1.Buttons("Compile").Enabled = True
      Else
        Toolbar1.Buttons("Compile").Enabled = False
      End If
    Case "Save"
      If txtCode.Tag = "" Then cdl.ShowSave
      Open cdl.FileName For Output As #1
        Print #1, txtCode.Text
      Close #1
      Dirty = False
      txtCode.Tag = cdl.FileTitle
      Caption = "Rubikon - " & cdl.FileTitle
    Case "Cut"
      Clipboard.Clear
      Clipboard.SetText txtCode.SelText
      txtCode.SelText = ""
    Case "Copy"
      Clipboard.Clear
      Clipboard.SetText txtCode.SelText
    Case "Paste"
      txtCode.SelText = Clipboard.GetText
    Case "Compile"
      If Dirty = True Then
        cdl.ShowSave
        Open cdl.FileName For Output As #1
          Print #1, txtCode.Text
        Close #1
        txtCode.Tag = cdl.FileTitle
        Caption = "Rubikon - " & cdl.FileTitle
      End If
      Dirty = False
      DefCnt = 0
      'txtROM_DblClick
      Text1 = "COMPILER REPORT" & vbCrLf & "---------------" & vbCrLf & vbCrLf
      Open txtRom For Binary As #1
      ln = 0
      Call Process("std.rbh")
      Call Process(cdl.FileTitle)
      Close #1
  End Select
Hell:
  Exit Sub
End Sub

Private Sub Process(f As String)
  Dim ff As Integer
  Dim i As Integer
  Dim SoT As String
  Dim OldSod As String
  Dim Lit As String
  Dim wite(32) As String
  Dim witec As Integer
  Dim b As Byte
 
  StatusBar1.SimpleText = "Processing " & f & "..."
  Text1 = Text1 & "Processing " & f & "..." & vbCrLf
  MousePointer = 11
  ff = FreeFile
  'On Error GoTo EdVenture
  Open f For Input As ff
  While Not EOF(ff)
    Line Input #ff, SoT
    ln = ln + 1
    Text1 = Text1 & Right("0000" & ln, 4) & "- " & SoT & vbCrLf
    OldSod = SoT
    'SoT = LCase(SoT)
    SoT = Replace(SoT, vbTab & vbTab & vbTab & vbTab & vbTab, " ")
    SoT = Replace(SoT, vbTab & vbTab & vbTab & vbTab, " ")
    SoT = Replace(SoT, vbTab & vbTab & vbTab, " ")
    SoT = Replace(SoT, vbTab & vbTab, " ")
    SoT = Replace(SoT, vbTab, " ")
    SoT = Replace(SoT, "0x", "&H")
    If Left(SoT, 2) <> "= " Then
      i = InStr(SoT, "'")
      If i Then SoT = Left(SoT, i - 1)
    End If
    If DefCnt > 0 Then
      For i = 0 To DefCnt
        SoT = Replace(SoT, Defines(i).Symbol, "&H" & Hex(Defines(i).Value))
      Next i
    End If
    Split SoT, " ", wite(), witec
    
    'StatusBar1.SimpleText = wite(1)
         
    Text1 = Text1 & "    - " & SoT & vbCrLf

    'OH WOW! Automajick Documentation Generator inlines!
    '---------------------------------------------------
    '--- autobr off
    '/// <html><head><title>Rubikon reference</title></head><body>
    '/// <font face=tahoma size=2>
    '/// <h2><img src=rubikonicon.gif align=left> RubiKode<sup>tm</sup> commands</h2>
    '--- autobr on

    Select Case wite(1)
      Case "#include"
      '/// <keyword>#include</keyword>
      '/// <syntax>#include sFile</syntax>
      '/// Include another file in the compilation process. Double quotes are automatically removed.
      '/// <example>#include "mydefs.rbh"</example>
        Call Process(Replace(wite(2), Chr(34), ""))
      Case "#define"
      '/// <keyword>#define</keyword>
      '/// <syntax>#define sSymbol iNumber</syntax>
      '/// Allows you to define symbols to replace numbers. Only numbers are allowed but they can be any size from byte to dword. It's good practice to use uppercase symbol names.
      '/// <example>#define MASTERBALL 1</example>
        Defines(DefCnt).Symbol = wite(2)
        Defines(DefCnt).Value = CLng(wite(3))
        'List1.AddItem wite(2) & " = 0x" & Hex(CLng(wite(3)))
        DefCnt = DefCnt + 1
      Case "#org", "#seek"
      '/// <keyword>#org</keyword>
      '/// <alias>#seek</alias>
      '/// Set the compiler's write cursor to the specified location in ROM.
      '/// <example>#org 0x605040 'continue (or start) writing somewhere in the empty space.</example>
        Text1 = Text1 & "Seeking to " & CLng(wite(2) + 1) & vbCrLf
        Seek #1, CLng(wite(2) + 1)
      Case "#autobank"
      '/// <keyword>#autobank</keyword>
      '/// <syntax>#autobank [on|off]</syntax>
      '/// Some might prefer to include the trailing 0x08 in pointers, others may not. When turned on, the extra 0x08 is automajickally added
      '/// to any pointer. Only #org is not affected since, as a file, the ROM always starts at bank 0x00.
      '/// <example>#autobank on
      '/// message 0x604020 'becomes 0F00 08604020
      '/// #autobank off
      '/// message 0x604020 'becomes 0F00 00604020, which is a bad thing.</example>
        If wite(2) = "off" Then
          AutoBank = False
          Text1 = Text1 & "-- Autobanking is OFF" & vbCrLf
        Else
          AutoBank = True
          Text1 = Text1 & "-- Autobanking is ON" & vbCrLf
        End If
      Case "#incbin"
      '/// <keyword>#incbin</keyword>
      '/// <syntax>#incbin sFile</syntax>
      '/// Same as #include, but for binary data in another file, like movement data.
        Text1 = Text1 & " -- Including binary " & wite(2) & "..."
        Dim bf As Integer
        bf = FreeFile
        Open Replace(wite(2), Chr(34), "") For Binary As bf
        For i = 0 To LOF(bf)
          StatusBar1.SimpleText = "Including binary - " & i & "/" & LOF(bf)
          Get #bf, , b
          Put #1, , b
        Next i
        Close bf
      Case "#raw", "#binary"
      '/// <keyword>#raw</keyword>
      '/// <alias>#binary</alias>
      '/// <syntax>#raw aLot</syntax>
      '/// Inserts a load of raw data into the ROM. This can be used for unsupported commands as well as movement data. To determine which data type to use, simply add in the type's name before any values that follow. The sequence <i>always</i> defaults to byte.
      '/// Possible data types are:
      '/// - byte (also char)
      '/// - word (also int or integer)
      '/// - dword (also long)
      '/// - pointer (also ptr)
      '/// <example>#binary 0x12 0x69 word 0x1234 dword 0x12345678 pointer 0x123456</example><br>
      '/// This example would output two bytes, one word, one dword and a pointer. Note how the pointer is affected by the AutoBank system and the DWord is not.
        Text1 = Text1 & "-- Starting raw byte sequence..." & vbCrLf
        b = 0 'Byte
        For i = 2 To witec
          Text1 = Text1 & " --- " & wite(i) & vbCrLf
          If wite(i) = "byte" Then
            b = 0
          ElseIf wite(i) = "char" Then
            b = 0
          ElseIf wite(i) = "word" Then
            b = 1
          ElseIf wite(i) = "int" Then
            b = 1
          ElseIf wite(i) = "integer" Then
            b = 1
          ElseIf wite(i) = "dword" Then
            b = 2
          ElseIf wite(i) = "long" Then
            b = 2
          ElseIf wite(i) = "pointer" Then
            b = 3
          ElseIf wite(i) = "ptr" Then
            b = 3
          Else 'It's a value
            If b = 0 Then 'Byte
              Put #1, , CByte(wite(i))
              Text1 = Text1 & "  -- Writing BYTE " & Hex(CByte(wite(i))) & vbCrLf
            ElseIf b = 1 Then 'Word
              Put #1, , CInt(wite(i))
              Text1 = Text1 & "  -- Writing WORD " & Hex(CInt(wite(i))) & vbCrLf
            ElseIf b = 2 Then 'DWord
              Put #1, , CLng(wite(i))
              Text1 = Text1 & "  -- Writing DWORD " & Hex(CLng(wite(i))) & vbCrLf
            ElseIf b = 3 Then 'Pointer
              Put #1, , CLng(IIf((AutoBank = True), &H8000000 + wite(i), wite(i)))
              Text1 = Text1 & "  -- Writing POINTER " & Hex(CLng(IIf((AutoBank = True), &H8000000 + wite(i), wite(i)))) & vbCrLf
            End If
          End If
        Next i
        Text1 = Text1 & "-- End of sequence" & vbCrLf
        
      Case "break"
      '/// <keyword>break</keyword>
      '/// <syntax>break</syntax>
      '/// End execution of the script.
        Put #1, , CByte(&H2)
      Case "goto"
      '/// <keyword>goto</keyword>
      '/// <syntax>goto lPointer</syntax>
      '/// Goto to another script.
      '/// <example>goto GOTO_BAGISFULL</example>
      Case "lock"
      '/// <keyword>lock</keyword>
      '/// <syntax>lock</syntax>
      '/// Locks down movement for the caller.
        Put #1, , CByte(&H6A)
      Case "faceplayer"
      '/// <keyword>faceplayer</keyword>
      '/// <syntax>faceplayer</syntax>
      '/// Turns the caller towards the player.
        Put #1, , CByte(&H5A)
      Case "unlock"
      '/// <keyword>unlock</keyword>
      '/// <syntax>unlock</syntax>
      '/// Resumes movement for the caller.
        Put #1, , CByte(&H6C)
      Case "message"
      '/// <keyword>message</keyword>
      '/// <syntax>message lPointer</syntax>
      '/// Displays a message on the screen. Requires 'boxset' to function properly.
      '/// <example>message 0x604020
      '/// boxset 0x04</example>
        Put #1, , CByte(&HF)
        Put #1, , CByte(&H0)
        Put #1, , CLng(IIf((AutoBank = True), &H8000000 + wite(2), wite(2)))
      Case "boxset"
      '/// <keyword>boxset</keyword>
      '/// <syntax>boxset bValue</syntax>
      '/// The author is uncertain of this command's use, but it allows a quick "yes/no" question to be asked if bValue is 0x05.
      '/// <example>message 0x604020 'question
      '/// boxset BOXSET_YESNO
      '/// if LASTRESULT NO 0x600020 'goto "no" handler
      '/// '"yes" handler goes here
      '/// ...
      '/// break
      '/// #org 0x600020
      '/// '"no" handler goes here
      '/// ...
      '/// break</example>
        Put #1, , CByte(&H9)
        Put #1, , CByte(wite(2))
      Case "setvar"
      '/// <keyword>setvar</keyword>
      '/// <syntax>setvar iVar iValue</syntax>
      '/// There are three versions of this command. All three (AFAIK) set a specific variable to a specific value.
      '/// This can be used for various things as some variables are constantly probed by the game, like 0x8001 to give an item.
      '/// This version's bytecode is 0x16.
        Put #1, , CByte(&H16)
        Put #1, , CInt(wite(2))
        Put #1, , CInt(wite(3))
      Case "setvar2"
      '/// <keyword>setvar2</keyword>
      '/// Another version of <i>setvar</i>. This one's 0x19
        Put #1, , CByte(&H19)
        Put #1, , CInt(wite(2))
        Put #1, , CInt(wite(3))
      Case "setvar3"
      '/// <keyword>setvar3</keyword>
      '/// Another version of <i>setvar</i>. This one's 0x1A and is used by <i>giveitem</i>.
        Put #1, , CByte(&H1A)
        Put #1, , CInt(wite(2))
        Put #1, , CInt(wite(3))
      Case "special"
      '/// <keyword>special</keyword>
      '/// <syntax>special iEvent</syntax>
      '/// Triggers a special event to occur.
      '/// <example>special SPECIAL_WALLYCATCH 'Play back movie of Wally catching a Ralts.</example>
        Put #1, , CByte(&H25)
        Put #1, , CInt(wite(2))
      Case "setflag"
      '/// <keyword>setflag</keyword>
      '/// <syntax>setflag iFlag</syntax>
      '/// Sets a flag.
      '/// <example>setflag 0x64
      '/// checkflag 0x64
      '/// if LASTRESULT YES 0x600020 'goto "yes" handler
      '/// '"no" handler goes here
      '/// ...
      '/// break
      '/// #org 0x600020
      '/// '"yes" handler goes here
      '/// ...
      '/// break</example>
        Put #1, , CByte(&H29)
        Put #1, , CInt(wite(2))
      Case "clearflag"
      '/// <keyword>clearflag</keyword>
      '/// <syntax>clearflag iFlag</keyword>
      '/// Clears a flag.
        Put #1, , CByte(&H2A)
        Put #1, , CInt(wite(2))
      Case "checkflag"
      '/// <keyword>checkflag</keyword>
      '/// <syntax>checkflag iFlag</syntax>
      '/// Checks if iFlag is set. For an example, see <i>setflag</i>.
        Put #1, , CByte(&H2B)
        Put #1, , CInt(wite(2))
      Case "compare"
      '/ <keyword>compare</keyword>
      '/ <syntax>compare iVar iValue</syntax>
      '/ Normally used with <i>if</i>.
        Put #1, , CByte(&H21)
        Put #1, , CInt(wite(2))
        Put #1, , CInt(wite(3))
      Case "if"
      '/ <keyword>if</keyword>
      '/ <syntax>if lPointer</syntax>
      '/ Used in conjunction with <i>compare</i>.
      '/ <example>if LASTRESULT YES 0x600020 'analogous to "if it = true then goto somewhere" in basic.
        Put #1, , CByte(&H6)
        Put #1, , CByte(&H1)
        Put #1, , CLng(IIf((AutoBank = True), &H8000000 + wite(4), wite(4)))
      Case "mart"
      '/// <keyword>mart</keyword>
      '/// <syntax>mart lPointer</syntax>
      '/// Opens the PokéMart shop system with the price list found at lPointer.
      '/// For more information on price lists, check the big pile of doggie doo.
        Put #1, , CByte(&H86)
        Put #1, , CLng(IIf((AutoBank = True), &H8000000 + wite(2), wite(2)))
      Case "choice"
      '/// <keyword>choice</keyword>
      '/// <syntax>choice bLeft bTop bList bCancel</syntax>
      '/// Puts up a list of choices for the player to make. Available choices depend on the value of bList. bCancel determines wether the player can press the B button to select the last item, if yes the last item should be "Cancel". As always, the player's choice is stored in LASTRESULT.
      '/// <example>message 0x604020 '"What city do you like best?"
      '/// choice 2 2 13 0 'items available are littleroot, slateport and lilycove
      '/// if LASTRESULT 0 0x600030 'goto littleroot handler
      '/// if LASTRESULT 1 0x600050 'goto slateport handler
      '/// 'lilycove handler starts right here, no "if LASTRESULT 3" needed.</example>
        Put #1, , CByte(&H6F)
        Put #1, , CByte(wite(2))
        Put #1, , CByte(wite(3))
        Put #1, , CByte(wite(4))
        Put #1, , CByte(wite(5))
      Case "cry"
      '/// <keyword>cry</keyword>
      '/// <syntax>cry iSpecies</syntax>
      '/// Plays back the cry of the specified Pokémon. Normally used for tame Pokémon such as the movers in R/S.
      '/// <example>cry 25 'pikachu!</example>
        Put #1, , CByte(&H30)
        Put #1, , CByte(&HA1)
        Put #1, , CInt(wite(2))
      Case "checkitem"
      '/// <keyword>checkitem</keyword>
      '/// <syntax>checkitem iItem</syntax>
      '/// Checks if the player carries at least one instance of the specified item.
      '/// <example>checkitem ITEM_MAXPOTION
      '/// if LASTRESULT YES 0x604020 'if we have at least one max potion...</example>
        Put #1, , CByte(&H47)
        Put #1, , CInt(wite(2))
        Put #1, , CByte(&H1)
        Put #1, , CByte(&H0)
      Case "trainerbattle"
      '/// <keyword>trainerbattle</keyword>
      '/// <syntax>trainerbattle bKind iBattle lPtrIntro lPtrDefeat</syntax>
      '/// Starts a trainer battle.
      '/// bKind is 0x00 for normal battles, 0x04 for 2-on-2 and 0x05 for rematches.
      '/// You'll need to add an extra parameter, lPtrNotEnough, for 2-on-2 battles.
      '/// <example>trainerbattle 0 MYBATTLEINTRO MYBATTLEDEFEAT
      '/// message MYBATTLEAFTERWARDS
      '/// boxset 4</example>
        Put #1, , CByte(&H5C)
        Put #1, , CByte(wite(2))
        Put #1, , CInt(wite(3))
        Put #1, , CByte(&H0)
        Put #1, , CByte(&H0)
        Put #1, , CLng(IIf((AutoBank = True), &H8000000 + wite(4), wite(4)))
        Put #1, , CLng(IIf((AutoBank = True), &H8000000 + wite(5), wite(5)))
        If CByte(wite(2)) = 4 Then Put #1, , CLng(IIf((AutoBank = True), &H8000000 + wite(6), wite(6)))
      Case "checkrematch"
      '/// <keyword>checkrematch</keyword>
      '/// <syntax>checkrematch iIndex</syntax>
      '/// Details are sketchy.
      '/// <example>trainerbattle 0 MYBATTLEIDX MYBATTLEINTRO MYBATTLEDEFEAT
      '/// checkrematch MYBATTLEREMATIDX
      '/// if LASTRESULT YES LBL_REMATCH
      '/// message MYBATTLEAFTERWARDS
      '/// boxset 6
      '/// break
      '/// #org LBL_REMATCH
      '/// trainerbattle 0 MYBATTLEIDXREDUX MYBATTLEINTROREDUX MYBATTLEDEFEATREDUX
      '/// message MYBATTLEAFTERWARDSREDUX
      '/// boxset 6
      '/// break</example>
        Put #1, , CByte(&H26)
        Put #1, , CByte(&HD)
        Put #1, , CByte(&H80)
        Put #1, , CInt(wite(2))
      Case "checkgender"
      '/// <keyword>checkgender</keyword>
      '/// <syntax>checkgender</syntax>
      '/// Simply puts 1 in in LASTRESULT if the player is a girl or 0 if he's a boy.
        Put #1, , CByte(&HA0)
      Case "givepokemon" 'NEW!
      '/// <keyword>givepokemon</keyword>
      '/// <syntax>givepokemon iSpecies iLevel iItem</syntax>
      '/// TODO --- Get off my lazy ass and write!
        Put #1, , CByte(&H79)
        Put #1, , CInt(wite(2))
        Put #1, , CByte(wite(3))
        Put #1, , CInt(wite(4))
        Put #1, , CInt(0) 'Filler word to make sure it works out.
      Case "fanfare" 'NEW!
      '/// <keyword>fanfare</keyword>
      '/// <syntax>fanfare iIndex</syntax>
        Put #1, , CByte(&H31)
        Put #1, , CInt(wite(2))
      Case "waitfanfare" 'NEW!
      '/// <keyword>waitfanfare</keyword>
      '/// <syntax>waitfanfare</syntax>
        Put #1, , CByte(&H32)
      Case "giveitem"
      '/// <keyword>giveitem</keyword>
      '/// <syntax>giveitem iItem</syntax>
      '/// Gives the player an item.
      '/// <example>giveitem ITEM_SODAPOP</example>
        Put #1, , CByte(&H1A)
        Put #1, , CByte(&H0)
        Put #1, , CByte(&H80)
        Put #1, , CInt(wite(2))
        Put #1, , CByte(&H1A)
        Put #1, , CByte(&H1)
        Put #1, , CByte(&H80)
        Put #1, , CByte(&H1)
        Put #1, , CByte(&H0)
      Case "playsound"
      '/// <keyword>playsound</keyword>
      '/// <syntax>playsound iIndex</syntax>
        Put #1, , CByte(&H33)
        Put #1, , CInt(wite(2))
      Case "fadesound"
      '/// <keyword>fadesound</keyword>
      '/// <syntax>fadesound iIndex</syntax>
        Put #1, , CByte(&H36)
        Put #1, , CInt(wite(2))
      Case "warp"
      '/// <keyword>warp</keyword>
      '/// <syntax>warp bBank bMap</syntax>
        Put #1, , CByte(&H39)
        Put #1, , CByte(wite(2))
        Put #1, , CByte(wite(3))
      Case "applymovement"
      '/// <keyword>applymovement</keyword>
      '/// <syntax>applymovement iSprite lPointer</syntax>
      '/// Applies the movement data at lPointer to the specified sprite.
      '/// <example>movement 0x5 0x604020
      '/// wait 0x0 'wait for movement to complete</example>
        Put #1, , CByte(&H47)
        Put #1, , CInt(wite(2))
        Put #1, , CLng(wite(3))
      Case "wait"
      '/// <keyword>wait</keyword>
      '/// <syntax>wait iVal</syntax>
        Put #1, , CByte(&H51)
        Put #1, , CInt(wite(2))
      
      Case "="
      '/// <keyword>=</keyword>
      '/// <syntax>= sText</syntax>
      '/// Raw text inserter! Supports four Rubikon tags along with the Sapp2Asc tags.
      '/// &lt;hero&gt; - inserts the player's name.
      '/// &lt;team&gt; - inserts the bad guys team name (aqua, magma...)
      '/// The other two are a bit buggy and shouldn't be used.
      '/// <example>= Hello &lt;hero&gt;! I'm so happy!\cHappy happy happy...</example>
        Lit = Right(SoT, Len(OldSod) - 2) & "\x"
        Lit = Replace(Lit, "<hero>", "\hfd\h01")
        Lit = Replace(Lit, "<red>", "\hfc\h04\h02\h0c")
        Lit = Replace(Lit, "</red>", "\hfc\h04\h01\h0c")
        Lit = Replace(Lit, "<team>", "\hfd\h08")
        Lit = Asc2Sapp(Lit)
        For i = 1 To Len(Lit)
          Put #1, , CByte(Asc(Mid(Lit, i, 1)))
        Next i
      Case ""
        'Skip a bit, Brother...
      Case Else
        'MsgBox "Unknown opcode " & wite(1) & vbCrLf & vbCrLf & SoT
        MsgBox "Unknown opcode " & wite(1) & " in " & f & "!"
        If f = txtCode.Tag Then
          i = InStr(txtCode, wite(1))
          txtCode.SelStart = i - 1
          i = InStr(i, txtCode, vbCrLf)
          txtCode.SelLength = i - txtCode.SelStart
        End If
        MousePointer = 0
        StatusBar1.SimpleText = "Unknown opcode " & wite(1) & " in " & f & "!"
        Exit Sub
    End Select
      
    '--- autobr off
    '/// <hr><h2><img src=rubikonicon.gif align=left> Predefined variables and constants in STD.RBH</h2>
    '--- autobr on
    '/// <keyword>0x800D - LASTRESULT</keyword>
    '/// Nearly every time something is checked, the answer is stored in this variable.
    '/// <keyword>SPRITEMOVE_INDEX<br>SPRITEMOVE_NEWXPOS<br>SPRITEMOVE_NEWYPOS</keyword>
    '/// Combine these three with <i>set</i> to move a sprite to a proper position on the map.
    '/// This is used in Littleroot to move Mom to either the left or right hand house door depending on the player's gender.
    '/// <keyword>GOTO_BAGISFULL</keyword>
    '/// A predefined label for Ruby to write ye olde standard error message when the player's inventory is completely filled. See <i>goto</i> for an example.
    '/// <hr><font size=1>Last updated: #DATE#, #TIME#</font></body></html>
    
  Wend
  Close ff
  MousePointer = 0
  StatusBar1.SimpleText = "Ready..."
  Exit Sub
EdVenture:
  'MsgBox f & " not found.", vbExclamation
  MsgBox Err.Number & " --- " & Err.Description
  Exit Sub
End Sub

Private Sub Split(ByVal sString As String, ByVal sDelim As String, ByRef sValues() As String, ByRef iCount As Integer)
' ==================================================================
' Splits sString into an array of parts which are
' delimited in the string by sDelim.  The array is
' indexed 1-iCount where iCount is the number of
' items.  If no items found iCount=1 and the array has
' one element, the original string.
'   sString : String to split
'   sDelim  : Delimiter
'   sValues : Return array of values
'   iCount  : Number of items returned in sValues()
' ==================================================================
  Dim iPos As Integer
  Dim iNextPos As Integer
  Dim iDelimLen As Integer
  iCount = 0
  Erase sValues
  iDelimLen = Len(sDelim)
  iPos = 1
  iNextPos = InStr(sString, sDelim)
  Do While iNextPos > 0
    iCount = iCount + 1
    'ReDim Preserve sValues(1 To iCount) As String
    sValues(iCount) = Mid$(sString, iPos, (iNextPos - iPos))
    iPos = iNextPos + iDelimLen
    iNextPos = InStr(iPos, sString, sDelim)
  Loop
  iCount = iCount + 1
  'ReDim Preserve sValues(1 To iCount) As String
  sValues(iCount) = Mid$(sString, iPos)
End Sub

Private Sub Text1_Click()
  Text1.Visible = False
End Sub

Private Sub txtROM_LostFocus()
  Open "Rubikon.cfg" For Output As #3
    Print #3, txtRom.Text
  Close #3
End Sub
