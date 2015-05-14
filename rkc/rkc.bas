Attribute VB_Name = "Rubikon"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public OutFile As String
Public AutoBank As Boolean
Public BreakOnError As Boolean

Private Type Define
  Symbol As String
  'Value As Long
  Value As String
End Type
Private Defines(1024) As Define
Private DefCnt As Integer

Public Sub Main()
  Dim Params(32) As String
  Dim pCount As Integer
  Dim NoShowLog As Boolean
  Dim i As Integer
  Dim cmd As String
  
  OutFile = "rkc.bin"
  
  cmd = Command$
  For i = 1 To 10
    cmd = Replace(cmd, "  ", " ")
  Next i
  
  Split cmd$, " ", Params(), pCount
  Open "rkclog.txt" For Output As #1
  Open OutFile For Binary As #2
  Scribe App.ProductName & " " & App.Major & "." & App.Minor & " by Kyoufu Kawa"
  Scribe "--------------------------------------------------------------"
  If cmd = "" Then
    Scribe "Usage: rkc.exe [/noshowlog] /o romfile rbcfile [rbcfile...]"
    Scribe "Outfile defaults to rkc.bin"
    GoTo CleanUp
  End If
  Scribe "Loading command database..."
  LoadCommands
  Randomize Timer
  If RCD.RubiCommands(1).Bytecode <> 1 Then
    Scribe "Empty command database detected. Program halted."
    GoTo CleanUp
  End If
  For i = 1 To pCount
    If Params(i) = "/noshowlog" Then
      NoShowLog = True
    ElseIf Params(i) = "/breakonerr" Then
      BreakOnError = True
    ElseIf Params(i) = "/o" Then
      OutFile = Params(i + 1)
      Scribe "Output file: " & OutFile
      Close #2
      Open OutFile For Binary As #2
      i = i + 1
    Else
      Process Params(i)
    End If
  Next i
CleanUp:
  Close #2
  Close #1
  If NoShowLog = False Then ShellExecute 0, vbNullString, "rkclog.txt", vbNullString, "", 1
End Sub

Public Sub Process(FileName As String)
  Dim RawInput As String
  Dim TextCopy As String
  Dim TextConv As String
  Dim Keywords(64) As String
  Dim KeyCount As Integer
  Dim LineNo As Long
  Dim ff As Integer
  Dim i As Long, j As Long
  Dim RawType As Byte
  
  Scribe "Processing " & FileName & "..."
  ff = FreeFile
  Scribe "File handle: " & ff
  'On Error GoTo NoOpen
  Open FileName For Input As ff
  'On Error GoTo StdErr
  While Not EOF(ff)
    LineNo = LineNo + 1
    Line Input #ff, RawInput
    'Scribe "{" & RawInput & "}"
    
    If Left(RawInput, 2) = "/*" Then
      Do
        If Right(RawInput, 2) = "*/" Then Exit Do
        LineNo = LineNo + 1
        Line Input #ff, RawInput
      Loop
      LineNo = LineNo + 1
      Line Input #ff, RawInput
    End If
    
    
    If DefCnt > 0 Then
      For i = 0 To DefCnt
        RawInput = Replace(RawInput, Defines(i).Symbol, Defines(i).Value)
      Next i
    End If
    
    TextCopy = Mid$(RawInput, 3)
    
    i = InStr(RawInput, "'")
    If i Then RawInput = Left(RawInput, i - 1)
    
    RawInput = Replace(RawInput, "0x", "&H")
    RawInput = Replace(RawInput, """", "")
    RawInput = Replace(RawInput, vbTab, " ")
    For i = 1 To 10
      RawInput = Replace(RawInput, "  ", " ")
    Next i
    RawInput = Trim(RawInput)
        
    'Scribe "{" & RawInput & "}"
    Split RawInput, " ", Keywords(), KeyCount
    Keywords(1) = LCase(Keywords(1))
    Select Case Keywords(1)
      Case "="
        Scribe LineNo & " - RAW TEXT"
        If Len(TextCopy) > 512 Then Scribe "'=' text string too long. I let you have 512 characters, that's 3 more than ANSI said I should."
        TextConv = Asc2Sapp(Left(TextCopy, 512) & "\x")
        Scribe " > sOld = """ & TextCopy & """"
        Scribe " > sNew = """ & TextConv & """"
        For i = 1 To Len(TextConv)
          Put #2, , CByte(Asc(Mid(TextConv, i, 1)))
        Next i

'----- DIRECTIVES -----
      Case "#include"
        Scribe LineNo & " - INCLUDE"
        Scribe " > sFile = " & Keywords(2)
        Process Keywords(2)
      Case "#define"
        Scribe LineNo & " - DEFINE"
        Scribe " > sSymbol = " & Keywords(2)
        Scribe " > sValue = " & Keywords(3)
        If DefCnt = 1025 Then
          Scribe "---------------------------------------------"
          Scribe "Error: Out of #defines on line " & LineNo & "."
          Scribe "       All 1024 available #defines are used."
          Scribe "---------------------------------------------"
        End If
        Defines(DefCnt).Symbol = Keywords(2)
        Defines(DefCnt).Value = Keywords(3)
        DefCnt = DefCnt + 1
      Case "#spawndefinelist"
        Scribe "------------------------------------"
        Scribe "You rang, master?"
        Scribe ""
        For i = 0 To DefCnt - 1
          Scribe i & " - [" & Defines(i).Symbol & "] > [" & Defines(i).Value & "]"
        Next i
        Scribe "------------------------------------"
      Case "#org", "#seek"
        Scribe LineNo & " - SEEK"
        Scribe " > pNewOffset = " & Hex$(CLng(Keywords(2)))
        Seek #2, CLng(Keywords(2) + 1)
      Case "#autobank"
        Scribe LineNo & " - AUTOBANK"
        If Keywords(2) = "off" Then
          AutoBank = False
          Scribe " > Autobanking is now OFF"
        ElseIf Keywords(2) = "on" Then
          AutoBank = True
          Scribe " > Autobanking is now ON"
        Else
          AutoBank = True
          Scribe " > Warning: Autobank parameter should be either ""on"" or ""off""."
          Scribe " > Autobanking is now ON anyway."
        End If
      Case "#raw", "#binary"
        Scribe LineNo & " - RAW"
        RawType = 0 'Byte
        For i = 2 To KeyCount
          If Keywords(i) = "b" Or Keywords(i) = "byte" Or Keywords(i) = "char" Then
            RawType = 0
          ElseIf Keywords(i) = "i" Or Keywords(i) = "word" Or Keywords(i) = "int" Or Keywords(i) = "integer" Then
            RawType = 1
          ElseIf Keywords(i) = "l" Or Keywords(i) = "dword" Or Keywords(i) = "long" Then
            RawType = 2
          ElseIf Keywords(i) = "p" Or Keywords(i) = "pointer" Or Keywords(i) = "ptr" Then
            RawType = 3
          Else 'It's a value
            If Left(Keywords(i), 2) <> "&H" Then Keywords(i) = "&H" & Keywords(i)
            If RawType = 0 Then 'Byte
              Put #2, , CByte(Keywords(i))
              Scribe " > bOut = " & Hex$(CByte(Keywords(i)))
            ElseIf RawType = 1 Then 'Word
              Put #2, , CInt(Keywords(i))
              Scribe " > iOut = " & Hex$(CInt(Keywords(i)))
            ElseIf RawType = 2 Then 'DWord
              Put #2, , CLng(Keywords(i))
              Scribe " > lOut = " & Hex$(CLng(Keywords(i)))
            ElseIf RawType = 3 Then 'Pointer
              Put #2, , CPtr(Keywords(i))
              Scribe " > pOut = " & Hex$(CPtr(Keywords(i)))
            End If
          End If
        Next i

'----- NATIVES -----
      'Natives have been REWRITTEN to support RCD, the RKC Command Database.
              
'----- CONSTRUCTS -----
      Case "wildbattle"
        Scribe LineNo & " - (--) WILDBATTLE"
        Scribe " > iSpecies = " & Hex(CInt(Keywords(2)))
        Scribe " > bLevel = " & Hex(CByte(Keywords(3)))
        Scribe " > bStyle = " & Hex(CByte(Keywords(4)))
        Put #2, , CByte(&HB6)
        Put #2, , CInt(Keywords(2))
        Put #2, , CByte(Keywords(3))
        Put #2, , CInt(0) 'Safety filler
        Put #2, , CByte(&H25)
        Select Case CByte(Keywords(4))
          Case 0: Put #2, , CByte(&H43)
          Case 1: Put #2, , CByte(&H37)
          Case 2: Put #2, , CByte(&H38)
          Case 3: Put #2, , CByte(&H39)
          Case Else: Put #2, , CByte(&H43)
        End Select
        Put #2, , CByte(&H1)

      Case "giveitem"
        Scribe LineNo & " - (--) GIVEITEM"
        Scribe " > iItem = " & Hex$(CInt(Keywords(2)))
        Scribe " > iQuantity = " & Hex$(CInt(Keywords(3)))
        Put #2, , CByte(&H1A)
        Put #2, , CByte(&H0)
        Put #2, , CByte(&H80)
        Put #2, , CInt(Keywords(2))
        Put #2, , CByte(&H1A)
        Put #2, , CByte(&H1)
        Put #2, , CByte(&H80)
        Put #2, , CInt(Keywords(3))
        Put #2, , CByte(&H9)
        Put #2, , CByte(&H0)
      
      Case "if"
        Scribe LineNo & " - (??) IF (native)"
        Scribe " > bCondition = " & Hex$(CByte(Keywords(2)))
        If Keywords(3) = "call" Or Keywords(3) = "gosub" Then
          Scribe " This is a calling IF, 0x06."
          Scribe " > pTarget = " & Hex$(CPtr(Keywords(4)))
          Put #2, , CByte(&H6)
          Put #2, , CByte(Keywords(2))
          Put #2, , CPtr(Keywords(4))
        ElseIf Keywords(3) = "jump" Or Keywords(3) = "goto" Then
          Scribe " This is a jumping IF, 0x07."
          Scribe " > pTarget = " & Hex$(CPtr(Keywords(4)))
          Put #2, , CByte(&H7)
          Put #2, , CByte(Keywords(2))
          Put #2, , CPtr(Keywords(4))
        Else
          Scribe " No kind specified. Assuming call."
          Scribe " > pTarget = " & Hex$(CPtr(Keywords(3)))
          Put #2, , CByte(&H6)
          Put #2, , CByte(Keywords(2))
          Put #2, , CPtr(Keywords(3))
        End If
      
      Case "message", "msgbox"
        Scribe LineNo & " - (0F) MESSAGE (native)"
        Scribe " > pText = " & Hex(CPtr(Keywords(2)))
        Put #2, , CByte(&HF)
        Put #2, , CByte(0)
        Put #2, , CPtr(Keywords(2))
        
      Case "trainerbattle"
        Scribe LineNo & " - (5C) TRAINERBATTLE (native)"
        Scribe " > bKind = " & Hex(CByte(Keywords(2)))
        Scribe " > iIndex = " & Hex(CInt(Keywords(3)))
        Scribe " > iFiller = " & Hex(CInt(Keywords(4)))
        Scribe " > pChallenge = " & Hex(CPtr(Keywords(5)))
        Scribe " > pDefeat = " & Hex(CPtr(Keywords(6)))
        Put #2, , CByte(Keywords(2))
        Put #2, , CInt(Keywords(3))
        Put #2, , CInt(Keywords(4))
        Put #2, , CPtr(Keywords(5))
        Put #2, , CPtr(Keywords(6))
        If Keywords(7) = "" Then
          Scribe " - No special third pointer specified. Pointer not written to ROM."
        Else
          Scribe " > pSpecial = " & Hex(CPtr(Keywords(7)))
          Put #2, , CPtr(Keywords(7))
        End If
        
      Case "" 'wtf??
      
      'Scan through the database in search for the keyword we found
      Case Else
        Dim FoundIt As Boolean
        
        FoundIt = False
        For i = 0 To &HFF
          If LCase(RCD.RubiCommands(i).Keyword) = Keywords(1) Then
            FoundIt = True
            Scribe LineNo & " - (" & Right("00" & Hex(RCD.RubiCommands(i).Bytecode), 2) & ") - " & UCase(RCD.RubiCommands(i).Keyword)
            Put #2, , CByte(RCD.RubiCommands(i).Bytecode)
            If RCD.RubiCommands(i).ParamCount > 0 Then
              For j = 0 To RCD.RubiCommands(i).ParamCount - 1
                Select Case RCD.RubiParameters(i, j).Size
                  Case 0
                    Scribe " > bByte = " & Hex(CByte(Keywords(2 + j)))
                    Put #2, , CByte(Keywords(2 + j))
                  Case 1
                    Scribe " > iWord = " & Hex(CInt(Keywords(2 + j)))
                    Put #2, , CInt(Keywords(2 + j))
                  Case 2
                    Scribe " > lDword = " & Hex(CLng(Keywords(2 + j)))
                    Put #2, , CLng(Keywords(2 + j))
                  Case 3
                    Scribe " > pOut = " & Hex(CPtr(Keywords(2 + j)))
                    Put #2, , CPtr(Keywords(2 + j))
                End Select
              Next j
            End If
            Exit For
          End If
        Next i
        If FoundIt = False Then
          Scribe "Unknown keyword """ & Keywords(1) & """ at line " & LineNo & "."
          If BreakOnError = True Then
            MsgBox "Unknown keyword """ & Keywords(1) & """ at line " & LineNo & "."
            GoTo CleanUp
          End If
        End If
    End Select
  Wend
  GoTo CleanUp
  
NoOpen:
  Scribe "ERROR: Failed to open file " & FileName & " for processing. Program halted."
  End
StdErr:
  Scribe "---------------------------------------------"
  Scribe "Error: """ & Err.Description & """ on line " & LineNo
  Scribe "       Processing terminated."
  Scribe "Line content: " & RawInput
  Scribe "---------------------------------------------"
  Dim niftyerr As String
  niftyerr = "Error """ & Err.Description & """ in file " & FileName & " on line " & LineNo & "." & vbCrLf
  If Err.Number = 13 Then 'Type Mismatch
    niftyerr = niftyerr & "Type mismatches often indicate a missing #define or parameter." & vbCrLf
  End If
  niftyerr = niftyerr & vbCrLf & "Line: " & RawInput
  MsgBox niftyerr
CleanUp:
  On Error Resume Next
  Scribe "Cleaning up..."
  Scribe "Closing file..."
  Close ff
  Scribe "Finished processing " & FileName & "."
  Exit Sub
End Sub

Public Function CPtr(base) As Long
  Dim b As Long
  b = CLng(base)
  If b > &H8000000 Then
    b = b - &H8000000
  End If
  If AutoBank = True Then
    CPtr = b + &H8000000
  Else
    CPtr = b
  End If
End Function

Public Sub Scribe(text As String)
  Dim ff As Integer
  Trace text
  Print #1, text
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

