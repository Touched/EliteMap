Attribute VB_Name = "modHackHeaders"
Option Explicit

Private Type tGUID
  part1 As Long
  part2 As Long
  part3 As Long
  part4 As Long
End Type

Public Type tRomHackHeader
  sHeader As String * 16
  gGUID As tGUID
  sName As String * 16
  sAuthor As String * 16
  sGroup As String * 16
  lWorkTime As Long
  lCueList As Long
  iLanguage As Integer
End Type

Public Sub WriteHeader(File As String, Header As tRomHackHeader)
  Dim ff As Integer
  ff = FreeFile
  Header.sHeader = "~EM Romhacks~"
  Open File For Binary As ff
    Put ff, &H1000001, Header
    Put ff, &HCE, CByte(1)
  Close ff
End Sub

Public Function ReadHeader(File As String, ByRef Header As tRomHackHeader) As Integer
  Dim ff As Integer
  ff = FreeFile
  ReadHeader = 1
  Open File For Binary As ff
    Get ff, &H1000001, Header
  Close ff
  If Trim(Header.sHeader) <> "~EM Romhacks~" Then ReadHeader = 0
End Function

Public Function MakeGUID(ByRef GUID As tGUID) As Integer
  Dim p As String
  p = "&H"
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  GUID.part1 = Val(p)
  
  p = "&H"
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  GUID.part2 = Val(p)
  
  p = "&H"
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  GUID.part3 = Val(p)
  
  p = "&H"
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  Randomize Timer
  p = p & Right("00" & Hex(Rnd * &HFF), 2)
  GUID.part4 = Val(p)
  
End Function

Public Function MakeStringFromGUID(ByRef GUID As tGUID) As String
'{BCB67D4D-2096-36BE-974C-A003-FC95041B}
  Dim part As String
  part = Right("00000000" & Hex(GUID.part1), 8)
  MakeStringFromGUID = "{" & part
  part = Right("00000000" & Hex(GUID.part2), 8)
  MakeStringFromGUID = MakeStringFromGUID & "-" & Left(part, 4) & "-" & Right(part, 4)
  part = Right("00000000" & Hex(GUID.part3), 8)
  MakeStringFromGUID = MakeStringFromGUID & "-" & Left(part, 4) & "-" & Right(part, 4)
  part = Right("00000000" & Hex(GUID.part4), 8)
  MakeStringFromGUID = MakeStringFromGUID & "-" & part & "}"
End Function
