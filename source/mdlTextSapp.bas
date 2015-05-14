Attribute VB_Name = "mdlTextSapp"
Public Function Asc2Sapp(ByVal asciistring As String) As String
  o = ""
  Dim m As Boolean
  For i = 1 To Len(asciistring)
    m = False
    If Len(asciistring) - (i - 1) > 3 Then
      Select Case Mid(asciistring, i, 4)
        Case "[Lv]": Y = &H34: m = True
        Case "[PK]": Y = &H53: m = True
        Case "[MN]": Y = &H54: m = True
        Case "[PO]": Y = &H55: m = True
        Case "[Ke]": Y = &H56: m = True
        Case "[BL]": Y = &H57: m = True
        Case "[OC]": Y = &H58: m = True
      End Select
      If m = True Then
        i = i + 3
      Else
        If Mid(asciistring, i, 2) = "\h" And IsHex(Mid(asciistring, i + 2, 2)) = True Then
          Y = Val("&H" & Mid(asciistring, i + 2, 2))
          i = i + 3
          m = True
        End If
      End If
    End If
    If Len(asciistring) - (i - 1) > 2 And m = False Then
      Select Case Mid(asciistring, i, 3)
        Case "[K]": Y = &H59: m = True
        Case "[U]": Y = &H79: m = True
        Case "[D]": Y = &H7A: m = True
        Case "[L]": Y = &H7B: m = True
        Case "[R]": Y = &H7C: m = True
        Case "[.]": Y = &HB0: m = True
        Case "[""]": Y = &HB1: m = True
        Case "[']": Y = &HB3: m = True
        Case "[m]": Y = &HB5: m = True
        Case "[f]": Y = &HB6: m = True
        Case "[p]": Y = &HB7: m = True
        Case "[x]": Y = &HB9: m = True
        Case "[>]": Y = &HEF: m = True
        Case "[u]": Y = &HF7: m = True
        Case "[d]": Y = &HF8: m = True
        Case "[l]": Y = &HF9: m = True
      End Select
      If m = True Then i = i + 2
    End If
  
    If Len(asciistring) - (i - 1) > 1 And m = False Then
      Select Case Mid(asciistring, i, 2)
        Case "\l": Y = &HFA: m = True
        Case "\p": Y = &HFB: m = True
        Case "\c": Y = &HFC: m = True
        Case "\v": Y = &HFD: m = True
        Case "\n": Y = &HFE: m = True
        Case "\x": Y = &HFF: m = True
      End Select
      If m = True Then i = i + 1
    End If
  
    If m = False Then
      Select Case Mid(asciistring, i, 1)
        Case " ": Y = &H0: m = True
        Case "À": Y = &H1: m = True
        Case "Á": Y = &H2: m = True
        Case "Â": Y = &H3: m = True
        Case "Ç": Y = &H4: m = True
        Case "È": Y = &H5: m = True
        Case "É": Y = &H6: m = True
        Case "Ê": Y = &H7: m = True
        Case "Ë": Y = &H8: m = True
        Case "Ì": Y = &H9: m = True
        Case "Î": Y = &HB: m = True
        Case "Ï": Y = &HC: m = True
        Case "Ò": Y = &HD: m = True
        Case "Ó": Y = &HE: m = True
        Case "Ô": Y = &HF: m = True
        Case "Œ": Y = &H10: m = True
        Case "Ù": Y = &H11: m = True
        Case "Ú": Y = &H12: m = True
        Case "Û": Y = &H13: m = True
        Case "ß": Y = &H15: m = True
        Case "à": Y = &H16: m = True
        Case "á": Y = &H17: m = True
        Case "ç": Y = &H19: m = True
        Case "è": Y = &H1A: m = True
        Case "é": Y = &H1B: m = True
        Case "ê": Y = &H1C: m = True
        Case "ë": Y = &H1D: m = True
        Case "ì": Y = &H1E: m = True
        Case "î": Y = &H20: m = True
        Case "ï": Y = &H21: m = True
        Case "ò": Y = &H22: m = True
        Case "ó": Y = &H23: m = True
        Case "œ": Y = &H24: m = True
        Case "ù": Y = &H25: m = True
        Case "ú": Y = &H26: m = True
        Case "°": Y = &H28: m = True
        Case "ª": Y = &H29: m = True
        Case "+": Y = &H2C: m = True
        Case "&": Y = &H2D: m = True
        Case "=": Y = &H35: m = True
        Case "¿": Y = &H51: m = True
        Case "¡": Y = &H52: m = True
        Case "Í": Y = &H5A: m = True
        Case "%": Y = &H5B: m = True
        Case "(": Y = &H5C: m = True
        Case ")": Y = &H5D: m = True
        Case "â": Y = &H68: m = True
        Case "í": Y = &H6F: m = True
        Case "0": Y = &HA1: m = True
        Case "1": Y = &HA2: m = True
        Case "2": Y = &HA3: m = True
        Case "3": Y = &HA4: m = True
        Case "4": Y = &HA5: m = True
        Case "5": Y = &HA6: m = True
        Case "6": Y = &HA7: m = True
        Case "7": Y = &HA8: m = True
        Case "8": Y = &HA9: m = True
        Case "9": Y = &HAA: m = True
        Case "!": Y = &HAB: m = True
        Case "?": Y = &HAC: m = True
        Case ".": Y = &HAD: m = True
        Case "-": Y = &HAE: m = True
        Case "·": Y = &HAF: m = True
        Case ",": Y = &HB8: m = True
        Case """": Y = &HB2: m = True
        Case "'": Y = &HB4: m = True
        Case "/": Y = &HBA: m = True
        Case "A": Y = &HBB: m = True
        Case "B": Y = &HBC: m = True
        Case "C": Y = &HBD: m = True
        Case "D": Y = &HBE: m = True
        Case "E": Y = &HBF: m = True
        Case "F": Y = &HC0: m = True
        Case "G": Y = &HC1: m = True
        Case "H": Y = &HC2: m = True
        Case "I": Y = &HC3: m = True
        Case "J": Y = &HC4: m = True
        Case "K": Y = &HC5: m = True
        Case "L": Y = &HC6: m = True
        Case "M": Y = &HC7: m = True
        Case "N": Y = &HC8: m = True
        Case "O": Y = &HC9: m = True
        Case "P": Y = &HCA: m = True
        Case "Q": Y = &HCB: m = True
        Case "R": Y = &HCC: m = True
        Case "S": Y = &HCD: m = True
        Case "T": Y = &HCE: m = True
        Case "U": Y = &HCF: m = True
        Case "V": Y = &HD0: m = True
        Case "W": Y = &HD1: m = True
        Case "X": Y = &HD2: m = True
        Case "Y": Y = &HD3: m = True
        Case "Z": Y = &HD4: m = True
        Case "a": Y = &HD5: m = True
        Case "b": Y = &HD6: m = True
        Case "c": Y = &HD7: m = True
        Case "d": Y = &HD8: m = True
        Case "e": Y = &HD9: m = True
        Case "f": Y = &HDA: m = True
        Case "g": Y = &HDB: m = True
        Case "h": Y = &HDC: m = True
        Case "i": Y = &HDD: m = True
        Case "j": Y = &HDE: m = True
        Case "k": Y = &HDF: m = True
        Case "l": Y = &HE0: m = True
        Case "m": Y = &HE1: m = True
        Case "n": Y = &HE2: m = True
        Case "o": Y = &HE3: m = True
        Case "p": Y = &HE4: m = True
        Case "q": Y = &HE5: m = True
        Case "r": Y = &HE6: m = True
        Case "s": Y = &HE7: m = True
        Case "t": Y = &HE8: m = True
        Case "u": Y = &HE9: m = True
        Case "v": Y = &HEA: m = True
        Case "w": Y = &HEB: m = True
        Case "x": Y = &HEC: m = True
        Case "y": Y = &HED: m = True
        Case "z": Y = &HEE: m = True
        Case ":": Y = &HF0: m = True
        Case "Ä": Y = &HF1: m = True
        Case "Ö": Y = &HF2: m = True
        Case "Ü": Y = &HF3: m = True
        Case "ä": Y = &HF4: m = True
        Case "ö": Y = &HF5: m = True
        Case "ü": Y = &HF6: m = True
      End Select
    End If
    If m = False Then Y = &H0
    o = o & Chr(Y)
  Next i
  Asc2Sapp = o
End Function

Public Function Sapp2Asc(ByVal sappstring As String, Optional japanese As Boolean) As String
  Dim Y As String
  Dim n As Boolean
  o = ""
  For i = 1 To Len(sappstring)
    X = IIf(Mid(sappstring, i, 1) = "", 0, Asc(Mid(sappstring, i, 1)))
    If n = True Then
      Y = "\h" & IIf(Len(Hex(X)) < 2, "0" & Hex(X), Hex(X))
      n = False
    Else
      If japanese Then
        Select Case X
        
        'This whole thing auto-converted from TBL file
          Case &H0: Y = " "
          Case &H1: Y = "a"
          Case &H2: Y = "i"
          Case &H3: Y = "u"
          Case &H4: Y = "e"
          Case &H5: Y = "o"
          Case &H6: Y = "ka"
          Case &H7: Y = "ki"
          Case &H8: Y = "ku"
          Case &H9: Y = "ke"
          Case &HA: Y = "ko"
          Case &HB: Y = "sa"
          Case &HC: Y = "shi"
          Case &HD: Y = "su"
          Case &HE: Y = "se"
          Case &HF: Y = "so"
          Case &H10: Y = "ta"
          Case &H11: Y = "chi"
          Case &H12: Y = "tsu"
          Case &H13: Y = "te"
          Case &H14: Y = "to"
          Case &H15: Y = "na"
          Case &H16: Y = "ni"
          Case &H17: Y = "nu"
          Case &H18: Y = "ne"
          Case &H19: Y = "no"
          Case &H1A: Y = "ha"
          Case &H1B: Y = "hi"
          Case &H1C: Y = "fu"
          Case &H1D: Y = "he"
          Case &H1E: Y = "ho"
          Case &H1F: Y = "ma"
          Case &H20: Y = "mi"
          Case &H21: Y = "mu"
          Case &H22: Y = "me"
          Case &H23: Y = "mo"
          Case &H27: Y = "ra"
          Case &H28: Y = "ri"
          Case &H29: Y = "ru"
          Case &H2A: Y = "re"
          Case &H2B: Y = "ro"
          Case &H2E: Y = "n"
          Case &H34: Y = "ya"
          Case &H35: Y = "yu"
          Case &H36: Y = "yo"
          Case &H37: Y = "ga"
          Case &H38: Y = "gi"
          Case &H39: Y = "gu"
          Case &H3A: Y = "ge"
          Case &H3B: Y = "go"
          Case &H3C: Y = "za"
          Case &H3D: Y = "ji"
          Case &H3E: Y = "zu"
          Case &H3F: Y = "ze"
          Case &H40: Y = "zo"
          Case &H41: Y = "da"
          Case &H42: Y = "dji"
          Case &H43: Y = "zu"
          Case &H44: Y = "de"
          Case &H45: Y = "do"
          Case &H46: Y = "ba"
          Case &H47: Y = "be"
          Case &H48: Y = "bo"
          Case &H49: Y = "pa"
          Case &H4A: Y = "pi"
          Case &H4B: Y = "pu"
          Case &H4C: Y = "pe"
          Case &H4D: Y = "po"
          Case &H51: Y = "A"
          Case &H52: Y = "I"
          Case &H53: Y = "U"
          Case &H54: Y = "E"
          Case &H55: Y = "O"
          Case &H56: Y = "KA"
          Case &H57: Y = "KI"
          Case &H58: Y = "KU"
          Case &H59: Y = "KE"
          Case &H5A: Y = "KO"
          Case &H5B: Y = "SA"
          Case &H5C: Y = "SHI"
          Case &H5D: Y = "SU"
          Case &H5E: Y = "SE"
          Case &H5F: Y = "SO"
          Case &H60: Y = "TA"
          Case &H61: Y = "CHI"
          Case &H62: Y = "TSU"
          Case &H63: Y = "TE"
          Case &H64: Y = "TO"
          Case &H65: Y = "NA"
          Case &H66: Y = "NI"
          Case &H67: Y = "NU"
          Case &H68: Y = "NE"
          Case &H69: Y = "NO"
          Case &H6A: Y = "HA"
          Case &H6B: Y = "HI"
          Case &H6C: Y = "FU"
          Case &H6D: Y = "HE"
          Case &H6E: Y = "HO"
          Case &H6F: Y = "MA"
          Case &H70: Y = "MI"
          Case &H71: Y = "MU"
          Case &H72: Y = "ME"
          Case &H73: Y = "MO"
          Case &H77: Y = "RA"
          Case &H78: Y = "RI"
          Case &H79: Y = "RU"
          Case &H7A: Y = "RE"
          Case &H7B: Y = "RO"
          Case &H7E: Y = "N"
          Case &H84: Y = "YA"
          Case &H85: Y = "YU"
          Case &H86: Y = "YO"
          Case &H87: Y = "GA"
          Case &H88: Y = "GI"
          Case &H89: Y = "GU"
          Case &H8A: Y = "GE"
          Case &H8B: Y = "GO"
          Case &H8C: Y = "ZA"
          Case &H8D: Y = "JI"
          Case &H8E: Y = "ZU"
          Case &H8F: Y = "ZE"
          Case &H90: Y = "ZO"
          Case &H91: Y = "DA"
          Case &H92: Y = "DJI"
          Case &H93: Y = "DU"
          Case &H94: Y = "DE"
          Case &H95: Y = "DO"
          Case &H96: Y = "BA"
          Case &H97: Y = "BI"
          Case &H98: Y = "BU"
          Case &H99: Y = "BE"
          Case &H9A: Y = "BO"
          Case &H9B: Y = "PA"
          Case &H9C: Y = "PI"
          Case &H9D: Y = "PU"
          Case &H9E: Y = "PE"
          Case &H9F: Y = "PO"
          Case &HA0: Y = "_"
          Case &HA1: Y = "0"
          Case &HA2: Y = "1"
          Case &HA3: Y = "2"
          Case &HA4: Y = "3"
          Case &HA5: Y = "4"
          Case &HA6: Y = "5"
          Case &HA7: Y = "6"
          Case &HA8: Y = "7"
          Case &HA9: Y = "8"
          Case &HAA: Y = "9"
          Case &HAC: Y = "?"
          Case &HAE: Y = "-"
        
          Case &HFA: Y = "\l"
          Case &HFB: Y = "\p"
          Case &HFC: Y = "\c": n = True
          Case &HFD: Y = "\v": n = True
          Case &HFE: Y = "\n"
          Case &HFF: Y = "\x"
          
          Case &H7C: Y = " "
          Case &H80: Y = ""
          Case Else: Y = "\h" & IIf(Len(Hex(X)) < 2, "0" & Hex(X), Hex(X))
        End Select
      Else
        Select Case X
          Case &H0: Y = " "
          Case &H1: Y = "À"
          Case &H2: Y = "Á"
          Case &H3: Y = "Â"
          Case &H4: Y = "Ç"
          Case &H5: Y = "È"
          Case &H6: Y = "É"
          Case &H7: Y = "Ê"
          Case &H8: Y = "Ë"
          Case &H9: Y = "Ì"
          Case &HB: Y = "Î"
          Case &HC: Y = "Ï"
          Case &HD: Y = "Ò"
          Case &HE: Y = "Ó"
          Case &HF: Y = "Ô"
          Case &H10: Y = "Œ"
          Case &H11: Y = "Ù"
          Case &H12: Y = "Ú"
          Case &H13: Y = "Û"
          Case &H15: Y = "ß"
          Case &H16: Y = "à"
          Case &H17: Y = "á"
          Case &H19: Y = "ç"
          Case &H1A: Y = "è"
          Case &H1B: Y = "é"
          Case &H1C: Y = "ê"
          Case &H1D: Y = "ë"
          Case &H1E: Y = "ì"
          Case &H20: Y = "î"
          Case &H21: Y = "ï"
          Case &H22: Y = "ò"
          Case &H23: Y = "ó"
          Case &H24: Y = "œ"
          Case &H25: Y = "ù"
          Case &H26: Y = "ú"
          Case &H28: Y = "°"
          Case &H29: Y = "ª"
          Case &H2B: Y = "&"
          Case &H2C: Y = "+"
          Case &H2D: Y = "&"
          Case &H34: Y = "[Lv]"
          Case &H35: Y = "="
          Case &H51: Y = "¿"
          Case &H52: Y = "¡"
          Case &H53: Y = "[PK]"
          Case &H54: Y = "[MN]"
          Case &H55: Y = "[PO]"
          Case &H56: Y = "[Ke]"
          Case &H57: Y = "[BL]"
          Case &H58: Y = "[OC]"
          Case &H59: Y = "[K]"
          Case &H5A: Y = "Í"
          Case &H5B: Y = "%"
          Case &H5C: Y = "("
          Case &H5D: Y = ")"
          Case &H68: Y = "â"
          Case &H6F: Y = "í"
          Case &H79: Y = "[U]"
          Case &H7A: Y = "[D]"
          Case &H7B: Y = "[L]"
          Case &H7C: Y = "[R]"
          Case &HA1: Y = "0"
          Case &HA2: Y = "1"
          Case &HA3: Y = "2"
          Case &HA4: Y = "3"
          Case &HA5: Y = "4"
          Case &HA6: Y = "5"
          Case &HA7: Y = "6"
          Case &HA8: Y = "7"
          Case &HA9: Y = "8"
          Case &HAA: Y = "9"
          Case &HAB: Y = "!"
          Case &HAC: Y = "?"
          Case &HAD: Y = "."
          Case &HAE: Y = "-"
          Case &HAF: Y = "·"
          Case &HB0: Y = "[.]"
          Case &HB1: Y = "[""]"
          Case &HB2: Y = """"
          Case &HB3: Y = "[']"
          Case &HB4: Y = "'"
          Case &HB5: Y = "[m]"
          Case &HB6: Y = "[f]"
          Case &HB7: Y = "[p]"
          Case &HB8: Y = ","
          Case &HB9: Y = "[x]"
          Case &HBA: Y = "/"
          Case &HBB: Y = "A"
          Case &HBC: Y = "B"
          Case &HBD: Y = "C"
          Case &HBE: Y = "D"
          Case &HBF: Y = "E"
          Case &HC0: Y = "F"
          Case &HC1: Y = "G"
          Case &HC2: Y = "H"
          Case &HC3: Y = "I"
          Case &HC4: Y = "J"
          Case &HC5: Y = "K"
          Case &HC6: Y = "L"
          Case &HC7: Y = "M"
          Case &HC8: Y = "N"
          Case &HC9: Y = "O"
          Case &HCA: Y = "P"
          Case &HCB: Y = "Q"
          Case &HCC: Y = "R"
          Case &HCD: Y = "S"
          Case &HCE: Y = "T"
          Case &HCF: Y = "U"
          Case &HD0: Y = "V"
          Case &HD1: Y = "W"
          Case &HD2: Y = "X"
          Case &HD3: Y = "Y"
          Case &HD4: Y = "Z"
          Case &HD5: Y = "a"
          Case &HD6: Y = "b"
          Case &HD7: Y = "c"
          Case &HD8: Y = "d"
          Case &HD9: Y = "e"
          Case &HDA: Y = "f"
          Case &HDB: Y = "g"
          Case &HDC: Y = "h"
          Case &HDD: Y = "i"
          Case &HDE: Y = "j"
          Case &HDF: Y = "k"
          Case &HE0: Y = "l"
          Case &HE1: Y = "m"
          Case &HE2: Y = "n"
          Case &HE3: Y = "o"
          Case &HE4: Y = "p"
          Case &HE5: Y = "q"
          Case &HE6: Y = "r"
          Case &HE7: Y = "s"
          Case &HE8: Y = "t"
          Case &HE9: Y = "u"
          Case &HEA: Y = "v"
          Case &HEB: Y = "w"
          Case &HEC: Y = "x"
          Case &HED: Y = "y"
          Case &HEE: Y = "z"
          Case &HEF: Y = "[>]"
          Case &HF0: Y = ":"
          Case &HF1: Y = "Ä"
          Case &HF2: Y = "Ö"
          Case &HF3: Y = "Ü"
          Case &HF4: Y = "ä"
          Case &HF5: Y = "ö"
          Case &HF6: Y = "ü"
          Case &HF7: Y = "[u]"
          Case &HF8: Y = "[d]"
          Case &HF9: Y = "[l]"
          Case &HFA: Y = "\l"
          Case &HFB: Y = "\p"
          Case &HFC: Y = "\c": n = True
          Case &HFD: Y = "\v": n = True
          Case &HFE: Y = "\n"
          Case &HFF: Y = "\x"
          Case Else: Y = "\h" & IIf(Len(Hex(X)) < 2, "0" & Hex(X), Hex(X))
        End Select
      End If
    End If
    o = o & Y
  Next i
  Sapp2Asc = o
End Function

'Public Function Sapp2AscTabled(ByVal sappstring As String, Optional japanese As Boolean) As String
'  Dim mytable(256) As String
'  Dim i As String
'  Dim a As Integer, b As Integer, c As Integer
'  Dim ff As Integer
'  ff = FreeFile
'  Open "obsidian.tbl" For Input As ff
'  While Not EOF(ff)
'    Line Input #ff, i
'    c = Val("&H" & Left(i, 2))
'    mytable(c) = Mid(i, 4)
'  Wend
'  Close #ff
'  i = ""
'  For a = 1 To Len(sappstring)
'    c = Asc(Mid(sappstring, a, 1))
'    If mytable(c) = "" Then
'      i = i & "\h" & Right("  " & Hex(c), 2)
'    Else
'      i = i & mytable(c)
'    End If
'  Next a
'  Sapp2Asc = i
'End Function

Private Function IsHex(ByVal hexstring As String) As Boolean
  Dim z As Boolean
  Dim Y As Byte
  For i = 1 To Len(hexstring)
    Y = Asc(Mid(hexstring, i, 1))
    z = IIf((Y > 47 And Y < 58) Or (Y > 64 And Y < 71) Or (Y > 96 And Y < 103), True, False)
    If z = False Then Exit For
  Next i
  IsHex = z
End Function

Private Function Hex2(ByVal indec As Long, Optional ByVal digits As Byte = 2) As String
  X = Hex(indec)
  Do While Len(X) < digits
    X = "0" & X
  Loop
  Hex2 = X
End Function
