Attribute VB_Name = "modBCompress"
Public laststructsize As Long

Public Function DecompressGold(ByVal offset As Long, gfxBuffer() As Byte) As Long
  Dim gfxPointer As Long
  Dim byteIn As Byte
  Get #256, offset, byteIn
  Do
    Get #256, , byteIn
    If byteIn = &HFF Then Exit Do
    c = byteIn And &HE0
    X = byteIn And &H1F
recalc:
    Select Case c
      Case 0
        For i = 0 To X
          Get #256, , byteIn
          gfxBuffer(gfxPointer) = byteIn
          gfxPointer = gfxPointer + 1
        Next i
      Case &H20
        Get #256, , byteIn
          For i = 0 To X
          gfxBuffer(gfxPointer) = byteIn
          gfxPointer = gfxPointer + 1
        Next i
      Case &H40
        Get #256, , byteIn
        Y = byteIn
        Get #256, , byteIn
        z = byteIn
        For i = 0 To X
          gfxBuffer(gfxPointer) = IIf(i Mod 2 = 0, Y, z)
          gfxPointer = gfxPointer + 1
        Next i
      Case &H60
        For i = 0 To X
          gfxBuffer(gfxPointer) = 0
          gfxPointer = gfxPointer + 1
        Next i
      Case &H80
        Get #256, , byteIn
        a = byteIn
        If a And &H80 = 0 Then
          Get #256, , byteIn
          n = byteIn
          For i = 0 To X
            gfxBuffer(gfxPointer) = gfxBuffer((a * &H100) + (n + 1) + i)
            gfxPointer = gfxPointer + 1
          Next i
        Else
          a = a And &H7F
          s = gfxPointer - a - 1
          For i = 0 To X
            gfxBuffer(gfxPointer) = gfxBuffer(s + i)
            gfxPointer = gfxPointer + 1
          Next i
        End If
      Case &HA0
        Get #256, , byteIn
        a = byteIn
        If a And &H80 = 0 Then
          Get #256, , byteIn
          n = byteIn
          For i = 0 To X 'reverse bit order
            gfxBuffer(gfxPointer) = gfxBuffer((a * &H100) + (n + 1) + i)
            gfxPointer = gfxPointer + 1
          Next i
        Else
          a = a And &H7F
          s = gfxPointer - a - 1
          For i = 0 To X 'reverse bit order..
            gfxBuffer(gfxPointer) = gfxBuffer(s + i)
            gfxPointer = gfxPointer + 1
          Next i
        End If
      Case &HC0
        Get #256, , byteIn
        a = byteIn
        If a And &H80 = 0 Then
          Get #256, , byteIn
          n = byteIn
          For i = 0 To X
            gfxBuffer(gfxPointer) = gfxBuffer((a * &H100) + (n + 1) - i)
            gfxPointer = gfxPointer + 1
          Next i
        Else
          a = a And &H7F
          s = gfxPointer - a - 1
          For i = 0 To X
            gfxBuffer(gfxPointer) = gfxBuffer(s - i)
            gfxPointer = gfxPointer + 1
          Next i
        End If
      Case &HE0
        c = X And &H1C
        Get #256, , byteIn
        w = byteIn
        X = ((X And 3) * &H100) + w
        GoTo recalc
    End Select
  Loop
  Close #256
  DecompressGold = gfxPointer
End Function

'Used to simulate the suffix operator ++ in C
Public Function Inc(variable, Optional ByVal value As Long = 1)
  Inc = variable
  variable = variable + value
End Function

Public Function LZ77UnComp(source() As Byte, dest() As Byte) As Long
  On Error Resume Next
  Dim Header As Long
  Header = (source(0) Or (source(1) * CLng(256)) Or (source(2) * CLng(2 ^ 16)) Or (source(3) * CLng(2 ^ 24)))
  Dim i As Long
  Dim j As Long
  Dim xin As Long
  Dim xout As Long
  xin = 4
  xout = 0
  Dim length As Long
  Dim offset As Long
  Dim windowOffset As Long
  Dim retLen As Long
  Dim xLen As Long
  Dim d As Byte
  Dim data As Long
  xLen = Header \ 256
  retLen = xLen
  Do While xLen > 0
    d = source(Inc(xin))
    For i = 0 To 7
      If (d And &H80) <> 0 Then
        data = ((source(xin) * (2 ^ 8)) Or source(xin + 1))
        Inc xin, 2
        length = (data \ (2 ^ 12)) + 3
        offset = (data And &HFFF)
        windowOffset = xout - offset - 1
        For j = 0 To length - 1
          dest(Inc(xout)) = dest(Inc(windowOffset))
          Inc xLen, -1
          If xLen = 0 Then
            LZ77UnComp = retLen
            Exit Function
          End If
        Next j
      Else
        dest(Inc(xout)) = source(Inc(xin))
        Inc xLen, -1
        If xLen = 0 Then
          LZ77UnComp = retLen
          Exit Function
        End If
      End If
      d = (d * 2) Mod 256
    Next i
  Loop
  LZ77UnComp = retLen
  laststructsize = xin
End Function

Public Function LZ77Comp(decmpsize As Long, source() As Byte, dest() As Byte) As Long
  Dim i As Long
  Dim j As Long
  Dim xin As Long
  Dim xout As Long
  
  Dim length As Long
  Dim offset As Long
  Dim tmplen As Long
  Dim tmpoff As Long
  Dim tmpxin As Long
  Dim tmpxout As Long
  Dim bufxout As Long
  Dim ctrl As Byte
  Dim xdata(0 To 7, 0 To 1) As Byte
  
  dest(0) = &H10  'unknown byte?
  dest(1) = (decmpsize Mod 256)
  dest(2) = ((decmpsize \ 256) Mod 256)
  dest(3) = ((decmpsize \ (2 ^ 16)) Mod 256)
  Do While (decmpsize > tmpxin)
    ctrl = 0
    For i = 7 To 0 Step -1
      If (xin < &H1000) Then
        j = xin
      Else
        j = &H1000
      End If
      length = 0
      offset = 0
      Do While (j > 1)
        tmpxin = xin
        tmpxout = (xin - j)
        Do While source(Inc(tmpxin)) = source(Inc(tmpxout))
          If (tmpxin >= decmpsize) Then Exit Do
        Loop
        tmplen = (tmpxin - xin - 1)
        tmpoff = (tmpxin - tmpxout - 1)
        If (tmplen > length) Then
          length = tmplen
          offset = tmpoff
        End If
        If (length >= &H12) Then Exit Do
        Inc j, -1
      Loop
      If (length >= 3) Then
        ctrl = ctrl Or (1 * (2 ^ i))
        If (length >= &H12) Then length = &H12
        xdata(i, 0) = (((length - 3) * (2 ^ 4)) Or (offset \ 256))
        xdata(i, 1) = (offset Mod 256)
        Inc xin, length
        Inc bufxout, 2
      Else
        xdata(i, 0) = source(Inc(xin))
        Inc bufxout
      End If
    Next i
    dest(Inc(xout) + 4) = ctrl
    For i = 7 To 0 Step -1
      dest(Inc(xout) + 4) = xdata(i, 0)
      If ((ctrl And &H80) <> 0) Then dest(Inc(xout) + 4) = xdata(i, 1)
      ctrl = (ctrl * 2) Mod 256
      If (decmpsize < tmpxin) Then Exit For
    Next i
  Loop
  LZ77Comp = xout + 4
End Function

