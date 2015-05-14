Attribute VB_Name = "RCD"
Option Explicit

Public Type tRubiParameter
  Size As Byte
  Description As String
End Type

Public Type tRubiCommand
  ASM As Long
  Bytecode As Byte
  Keyword As String
  Description As String
  ParamCount As Integer
End Type

Public RubiCommands(&HFF) As tRubiCommand
Public RubiParameters(&HFF, 7) As tRubiParameter

Public Sub LoadCommands()
  Dim i As Integer, j As Integer
  'On Error GoTo Hell
  Open "rubikon.dat" For Binary As #255
  For i = 0 To &HFF
    Get #255, , RubiCommands(i)
  Next i
  For i = 0 To &HFF
    For j = 0 To 7
      Get #255, , RubiParameters(i, j)
    Next j
  Next i
  Close #255
  Exit Sub
'Hell:
  'MsgBox "Rubikon.dat not found. Using empty database.", vbInformation
End Sub

Public Sub SaveCommands()
  Dim i As Integer, j As Integer
  Open "rubikon.dat" For Binary As #1
  For i = 0 To &HFF
    Put #1, , RubiCommands(i)
  Next i
  For i = 0 To &HFF
    For j = 0 To 7
      Put #1, , RubiParameters(i, j)
    Next j
  Next i
  Close #1
End Sub

Public Function GetSizeName(i As Byte) As String
  Select Case i
    Case 0: GetSizeName = "Byte"
    Case 1: GetSizeName = "Word"
    Case 2: GetSizeName = "DWord"
    Case 3: GetSizeName = "Pointer"
    Case Else: GetSizeName = "???"
  End Select
End Function
