Attribute VB_Name = "modExtras"
Option Explicit

Public Function PadHex$(a, l)
  PadHex$ = Right$("00000000" & Hex$(a), l)
End Function

Public Function BitSet(Number As Long, ByVal Bit As Long) As Long
Attribute BitSet.VB_Description = "Returns Number with Bit set."
  If Bit = 31 Then
    Number = &H80000000 Or Number
  Else
    Number = (2 ^ Bit) Or Number
  End If
  BitSet = Number
End Function

Public Function BitClear(Number As Long, ByVal Bit As Long) As Long
Attribute BitClear.VB_Description = "Returns Number with Bit cleared."
  If Bit = 31 Then
    Number = &H7FFFFFFF And Number
  Else
    Number = ((2 ^ Bit) Xor &HFFFFFFFF) And Number
  End If
  BitClear = Number
End Function

Public Function BitIsSet(ByVal Number As Long, ByVal Bit As Long) As Boolean
Attribute BitIsSet.VB_Description = "Returns True if Bit is set."
  BitIsSet = False
  If Bit = 31 Then
    If Number And &H80000000 Then BitIsSet = True
  Else
    If Number And (2 ^ Bit) Then BitIsSet = True
  End If
End Function
