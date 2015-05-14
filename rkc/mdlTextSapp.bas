Attribute VB_Name = "mdlTextSapp"
Public Function Asc2Sapp(ByVal asciistring As String) As String
  O = ""
  Dim m As Boolean
  
  asciistring = Replace(asciistring, "[player]", "\v\h01")
  asciistring = Replace(asciistring, "[rival]", "\v\h06")
  asciistring = Replace(asciistring, "[game]", "\v\h07")
  asciistring = Replace(asciistring, "[team]", "\v\h08")
  asciistring = Replace(asciistring, "[otherteam]", "\v\h09")
  asciistring = Replace(asciistring, "[teamleader]", "\v\h0A")
  asciistring = Replace(asciistring, "[otherteamleader]", "\v\h0B")
  asciistring = Replace(asciistring, "[legend]", "\v\h0C")
  asciistring = Replace(asciistring, "[otherlegend]", "\v\h0D")
  
  For i = 1 To Len(asciistring)
    m = False
    If Len(asciistring) - (i - 1) > 3 Then
      Select Case Mid(asciistring, i, 4)
        Case "[Lv]": y = &H34: m = True
        Case "[PK]": y = &H53: m = True
        Case "[MN]": y = &H54: m = True
        Case "[PO]": y = &H55: m = True
        Case "[Ke]": y = &H56: m = True
        Case "[BL]": y = &H57: m = True
        Case "[OC]": y = &H58: m = True
      End Select
      If m = True Then
        i = i + 3
      Else
        If Mid(asciistring, i, 2) = "\h" And IsHex(Mid(asciistring, i + 2, 2)) = True Then
          y = Val("&H" & Mid(asciistring, i + 2, 2))
          i = i + 3
          m = True
        End If
      End If
    End If
    If Len(asciistring) - (i - 1) > 2 And m = False Then
      Select Case Mid(asciistring, i, 3)
        Case "[K]": y = &H59: m = True
        Case "[U]": y = &H79: m = True
        Case "[D]": y = &H7A: m = True
        Case "[L]": y = &H7B: m = True
        Case "[R]": y = &H7C: m = True
        Case "[.]": y = &HB0: m = True
        Case "[""]": y = &HB1: m = True
        Case "[']": y = &HB3: m = True
        Case "[m]": y = &HB5: m = True
        Case "[f]": y = &HB6: m = True
        Case "[p]": y = &HB7: m = True
        Case "[x]": y = &HB9: m = True
        Case "[>]": y = &HEF: m = True
        Case "[u]": y = &HF7: m = True
        Case "[d]": y = &HF8: m = True
        Case "[l]": y = &HF9: m = True
      End Select
      If m = True Then i = i + 2
    End If
  
    If Len(asciistring) - (i - 1) > 1 And m = False Then
      Select Case Mid(asciistring, i, 2)
        Case "\l": y = &HFA: m = True
        Case "\p": y = &HFB: m = True
        Case "\c": y = &HFC: m = True
        Case "\v": y = &HFD: m = True
        Case "\n": y = &HFE: m = True
        Case "\x": y = &HFF: m = True
      End Select
      If m = True Then i = i + 1
    End If
  
    If m = False Then
      Select Case Mid(asciistring, i, 1)
        Case " ": y = &H0: m = True
        Case "À": y = &H1: m = True
        Case "Á": y = &H2: m = True
        Case "Â": y = &H3: m = True
        Case "Ç": y = &H4: m = True
        Case "È": y = &H5: m = True
        Case "É": y = &H6: m = True
        Case "Ê": y = &H7: m = True
        Case "Ë": y = &H8: m = True
        Case "Ì": y = &H9: m = True
        Case "Î": y = &HB: m = True
        Case "Ï": y = &HC: m = True
        Case "Ò": y = &HD: m = True
        Case "Ó": y = &HE: m = True
        Case "Ô": y = &HF: m = True
        Case "Œ": y = &H10: m = True
        Case "Ù": y = &H11: m = True
        Case "Ú": y = &H12: m = True
        Case "Û": y = &H13: m = True
        Case "ß": y = &H15: m = True
        Case "à": y = &H16: m = True
        Case "á": y = &H17: m = True
        Case "ç": y = &H19: m = True
        Case "è": y = &H1A: m = True
        Case "é": y = &H1B: m = True
        Case "ê": y = &H1C: m = True
        Case "ë": y = &H1D: m = True
        Case "ì": y = &H1E: m = True
        Case "î": y = &H20: m = True
        Case "ï": y = &H21: m = True
        Case "ò": y = &H22: m = True
        Case "ó": y = &H23: m = True
        Case "œ": y = &H24: m = True
        Case "ù": y = &H25: m = True
        Case "ú": y = &H26: m = True
        Case "°": y = &H28: m = True
        Case "ª": y = &H29: m = True
        Case "+": y = &H2C: m = True
        Case "&": y = &H2D: m = True
        Case "=": y = &H35: m = True
        Case "¿": y = &H51: m = True
        Case "¡": y = &H52: m = True
        Case "Í": y = &H5A: m = True
        Case "%": y = &H5B: m = True
        Case "(": y = &H5C: m = True
        Case ")": y = &H5D: m = True
        Case "â": y = &H68: m = True
        Case "í": y = &H6F: m = True
        Case "0": y = &HA1: m = True
        Case "1": y = &HA2: m = True
        Case "2": y = &HA3: m = True
        Case "3": y = &HA4: m = True
        Case "4": y = &HA5: m = True
        Case "5": y = &HA6: m = True
        Case "6": y = &HA7: m = True
        Case "7": y = &HA8: m = True
        Case "8": y = &HA9: m = True
        Case "9": y = &HAA: m = True
        Case "!": y = &HAB: m = True
        Case "?": y = &HAC: m = True
        Case ".": y = &HAD: m = True
        Case "-": y = &HAE: m = True
        Case "·": y = &HAF: m = True
        Case ",": y = &HB8: m = True
        Case """": y = &HB2: m = True
        Case "'": y = &HB4: m = True
        Case "/": y = &HBA: m = True
        Case "A": y = &HBB: m = True
        Case "B": y = &HBC: m = True
        Case "C": y = &HBD: m = True
        Case "D": y = &HBE: m = True
        Case "E": y = &HBF: m = True
        Case "F": y = &HC0: m = True
        Case "G": y = &HC1: m = True
        Case "H": y = &HC2: m = True
        Case "I": y = &HC3: m = True
        Case "J": y = &HC4: m = True
        Case "K": y = &HC5: m = True
        Case "L": y = &HC6: m = True
        Case "M": y = &HC7: m = True
        Case "N": y = &HC8: m = True
        Case "O": y = &HC9: m = True
        Case "P": y = &HCA: m = True
        Case "Q": y = &HCB: m = True
        Case "R": y = &HCC: m = True
        Case "S": y = &HCD: m = True
        Case "T": y = &HCE: m = True
        Case "U": y = &HCF: m = True
        Case "V": y = &HD0: m = True
        Case "W": y = &HD1: m = True
        Case "X": y = &HD2: m = True
        Case "Y": y = &HD3: m = True
        Case "Z": y = &HD4: m = True
        Case "a": y = &HD5: m = True
        Case "b": y = &HD6: m = True
        Case "c": y = &HD7: m = True
        Case "d": y = &HD8: m = True
        Case "e": y = &HD9: m = True
        Case "f": y = &HDA: m = True
        Case "g": y = &HDB: m = True
        Case "h": y = &HDC: m = True
        Case "i": y = &HDD: m = True
        Case "j": y = &HDE: m = True
        Case "k": y = &HDF: m = True
        Case "l": y = &HE0: m = True
        Case "m": y = &HE1: m = True
        Case "n": y = &HE2: m = True
        Case "o": y = &HE3: m = True
        Case "p": y = &HE4: m = True
        Case "q": y = &HE5: m = True
        Case "r": y = &HE6: m = True
        Case "s": y = &HE7: m = True
        Case "t": y = &HE8: m = True
        Case "u": y = &HE9: m = True
        Case "v": y = &HEA: m = True
        Case "w": y = &HEB: m = True
        Case "x": y = &HEC: m = True
        Case "y": y = &HED: m = True
        Case "z": y = &HEE: m = True
        Case ":": y = &HF0: m = True
        Case "Ä": y = &HF1: m = True
        Case "Ö": y = &HF2: m = True
        Case "Ü": y = &HF3: m = True
        Case "ä": y = &HF4: m = True
        Case "ö": y = &HF5: m = True
        Case "ü": y = &HF6: m = True
      End Select
    End If
    If m = False Then y = &H0
    O = O & Chr(y)
  Next i
  Asc2Sapp = O
End Function

Public Function Sapp2Asc(ByVal sappstring As String, Optional japanese As Boolean) As String
  Dim y As String
  Dim n As Boolean
  O = ""
    
  For i = 1 To Len(sappstring)
    x = IIf(Mid(sappstring, i, 1) = "", 0, Asc(Mid(sappstring, i, 1)))
    If n = True Then
      y = "\h" & IIf(Len(Hex(x)) < 2, "0" & Hex(x), Hex(x))
      n = False
    Else
      If japanese Then
        Select Case x
        
        'This whole thing auto-converted from TBL file
          Case &H0: y = " "
          Case &H1: y = "a"
          Case &H2: y = "i"
          Case &H3: y = "u"
          Case &H4: y = "e"
          Case &H5: y = "o"
          Case &H6: y = "ka"
          Case &H7: y = "ki"
          Case &H8: y = "ku"
          Case &H9: y = "ke"
          Case &HA: y = "ko"
          Case &HB: y = "sa"
          Case &HC: y = "shi"
          Case &HD: y = "su"
          Case &HE: y = "se"
          Case &HF: y = "so"
          Case &H10: y = "ta"
          Case &H11: y = "chi"
          Case &H12: y = "tsu"
          Case &H13: y = "te"
          Case &H14: y = "to"
          Case &H15: y = "na"
          Case &H16: y = "ni"
          Case &H17: y = "nu"
          Case &H18: y = "ne"
          Case &H19: y = "no"
          Case &H1A: y = "ha"
          Case &H1B: y = "hi"
          Case &H1C: y = "fu"
          Case &H1D: y = "he"
          Case &H1E: y = "ho"
          Case &H1F: y = "ma"
          Case &H20: y = "mi"
          Case &H21: y = "mu"
          Case &H22: y = "me"
          Case &H23: y = "mo"
          Case &H27: y = "ra"
          Case &H28: y = "ri"
          Case &H29: y = "ru"
          Case &H2A: y = "re"
          Case &H2B: y = "ro"
          Case &H2E: y = "n"
          Case &H34: y = "ya"
          Case &H35: y = "yu"
          Case &H36: y = "yo"
          Case &H37: y = "ga"
          Case &H38: y = "gi"
          Case &H39: y = "gu"
          Case &H3A: y = "ge"
          Case &H3B: y = "go"
          Case &H3C: y = "za"
          Case &H3D: y = "ji"
          Case &H3E: y = "zu"
          Case &H3F: y = "ze"
          Case &H40: y = "zo"
          Case &H41: y = "da"
          Case &H42: y = "dji"
          Case &H43: y = "zu"
          Case &H44: y = "de"
          Case &H45: y = "do"
          Case &H46: y = "ba"
          Case &H47: y = "be"
          Case &H48: y = "bo"
          Case &H49: y = "pa"
          Case &H4A: y = "pi"
          Case &H4B: y = "pu"
          Case &H4C: y = "pe"
          Case &H4D: y = "po"
          Case &H51: y = "A"
          Case &H52: y = "I"
          Case &H53: y = "U"
          Case &H54: y = "E"
          Case &H55: y = "O"
          Case &H56: y = "KA"
          Case &H57: y = "KI"
          Case &H58: y = "KU"
          Case &H59: y = "KE"
          Case &H5A: y = "KO"
          Case &H5B: y = "SA"
          Case &H5C: y = "SHI"
          Case &H5D: y = "SU"
          Case &H5E: y = "SE"
          Case &H5F: y = "SO"
          Case &H60: y = "TA"
          Case &H61: y = "CHI"
          Case &H62: y = "TSU"
          Case &H63: y = "TE"
          Case &H64: y = "TO"
          Case &H65: y = "NA"
          Case &H66: y = "NI"
          Case &H67: y = "NU"
          Case &H68: y = "NE"
          Case &H69: y = "NO"
          Case &H6A: y = "HA"
          Case &H6B: y = "HI"
          Case &H6C: y = "FU"
          Case &H6D: y = "HE"
          Case &H6E: y = "HO"
          Case &H6F: y = "MA"
          Case &H70: y = "MI"
          Case &H71: y = "MU"
          Case &H72: y = "ME"
          Case &H73: y = "MO"
          Case &H77: y = "RA"
          Case &H78: y = "RI"
          Case &H79: y = "RU"
          Case &H7A: y = "RE"
          Case &H7B: y = "RO"
          Case &H7E: y = "N"
          Case &H84: y = "YA"
          Case &H85: y = "YU"
          Case &H86: y = "YO"
          Case &H87: y = "GA"
          Case &H88: y = "GI"
          Case &H89: y = "GU"
          Case &H8A: y = "GE"
          Case &H8B: y = "GO"
          Case &H8C: y = "ZA"
          Case &H8D: y = "JI"
          Case &H8E: y = "ZU"
          Case &H8F: y = "ZE"
          Case &H90: y = "ZO"
          Case &H91: y = "DA"
          Case &H92: y = "DJI"
          Case &H93: y = "DU"
          Case &H94: y = "DE"
          Case &H95: y = "DO"
          Case &H96: y = "BA"
          Case &H97: y = "BI"
          Case &H98: y = "BU"
          Case &H99: y = "BE"
          Case &H9A: y = "BO"
          Case &H9B: y = "PA"
          Case &H9C: y = "PI"
          Case &H9D: y = "PU"
          Case &H9E: y = "PE"
          Case &H9F: y = "PO"
          Case &HA0: y = "_"
          Case &HA1: y = "0"
          Case &HA2: y = "1"
          Case &HA3: y = "2"
          Case &HA4: y = "3"
          Case &HA5: y = "4"
          Case &HA6: y = "5"
          Case &HA7: y = "6"
          Case &HA8: y = "7"
          Case &HA9: y = "8"
          Case &HAA: y = "9"
          Case &HAC: y = "?"
          Case &HAE: y = "-"
        
          Case &HFA: y = "\l"
          Case &HFB: y = "\p"
          Case &HFC: y = "\c": n = True
          Case &HFD: y = "\v": n = True
          Case &HFE: y = "\n"
          Case &HFF: y = "\x"
          
          Case &H7C: y = " "
          Case &H80: y = ""
          Case Else: y = "\h" & IIf(Len(Hex(x)) < 2, "0" & Hex(x), Hex(x))
        End Select
      Else
        Select Case x
          Case &H0: y = " "
          Case &H1: y = "À"
          Case &H2: y = "Á"
          Case &H3: y = "Â"
          Case &H4: y = "Ç"
          Case &H5: y = "È"
          Case &H6: y = "É"
          Case &H7: y = "Ê"
          Case &H8: y = "Ë"
          Case &H9: y = "Ì"
          Case &HB: y = "Î"
          Case &HC: y = "Ï"
          Case &HD: y = "Ò"
          Case &HE: y = "Ó"
          Case &HF: y = "Ô"
          Case &H10: y = "Œ"
          Case &H11: y = "Ù"
          Case &H12: y = "Ú"
          Case &H13: y = "Û"
          Case &H15: y = "ß"
          Case &H16: y = "à"
          Case &H17: y = "á"
          Case &H19: y = "ç"
          Case &H1A: y = "è"
          Case &H1B: y = "é"
          Case &H1C: y = "ê"
          Case &H1D: y = "ë"
          Case &H1E: y = "ì"
          Case &H20: y = "î"
          Case &H21: y = "ï"
          Case &H22: y = "ò"
          Case &H23: y = "ó"
          Case &H24: y = "œ"
          Case &H25: y = "ù"
          Case &H26: y = "ú"
          Case &H28: y = "°"
          Case &H29: y = "ª"
          Case &H2B: y = "&"
          Case &H2C: y = "+"
          Case &H2D: y = "&"
          Case &H34: y = "[Lv]"
          Case &H35: y = "="
          Case &H51: y = "¿"
          Case &H52: y = "¡"
          Case &H53: y = "[PK]"
          Case &H54: y = "[MN]"
          Case &H55: y = "[PO]"
          Case &H56: y = "[Ke]"
          Case &H57: y = "[BL]"
          Case &H58: y = "[OC]"
          Case &H59: y = "[K]"
          Case &H5A: y = "Í"
          Case &H5B: y = "%"
          Case &H5C: y = "("
          Case &H5D: y = ")"
          Case &H68: y = "â"
          Case &H6F: y = "í"
          Case &H79: y = "[U]"
          Case &H7A: y = "[D]"
          Case &H7B: y = "[L]"
          Case &H7C: y = "[R]"
          Case &HA1: y = "0"
          Case &HA2: y = "1"
          Case &HA3: y = "2"
          Case &HA4: y = "3"
          Case &HA5: y = "4"
          Case &HA6: y = "5"
          Case &HA7: y = "6"
          Case &HA8: y = "7"
          Case &HA9: y = "8"
          Case &HAA: y = "9"
          Case &HAB: y = "!"
          Case &HAC: y = "?"
          Case &HAD: y = "."
          Case &HAE: y = "-"
          Case &HAF: y = "·"
          Case &HB0: y = "[.]"
          Case &HB1: y = "[""]"
          Case &HB2: y = """"
          Case &HB3: y = "[']"
          Case &HB4: y = "'"
          Case &HB5: y = "[m]"
          Case &HB6: y = "[f]"
          Case &HB7: y = "[p]"
          Case &HB8: y = ","
          Case &HB9: y = "[x]"
          Case &HBA: y = "/"
          Case &HBB: y = "A"
          Case &HBC: y = "B"
          Case &HBD: y = "C"
          Case &HBE: y = "D"
          Case &HBF: y = "E"
          Case &HC0: y = "F"
          Case &HC1: y = "G"
          Case &HC2: y = "H"
          Case &HC3: y = "I"
          Case &HC4: y = "J"
          Case &HC5: y = "K"
          Case &HC6: y = "L"
          Case &HC7: y = "M"
          Case &HC8: y = "N"
          Case &HC9: y = "O"
          Case &HCA: y = "P"
          Case &HCB: y = "Q"
          Case &HCC: y = "R"
          Case &HCD: y = "S"
          Case &HCE: y = "T"
          Case &HCF: y = "U"
          Case &HD0: y = "V"
          Case &HD1: y = "W"
          Case &HD2: y = "X"
          Case &HD3: y = "Y"
          Case &HD4: y = "Z"
          Case &HD5: y = "a"
          Case &HD6: y = "b"
          Case &HD7: y = "c"
          Case &HD8: y = "d"
          Case &HD9: y = "e"
          Case &HDA: y = "f"
          Case &HDB: y = "g"
          Case &HDC: y = "h"
          Case &HDD: y = "i"
          Case &HDE: y = "j"
          Case &HDF: y = "k"
          Case &HE0: y = "l"
          Case &HE1: y = "m"
          Case &HE2: y = "n"
          Case &HE3: y = "o"
          Case &HE4: y = "p"
          Case &HE5: y = "q"
          Case &HE6: y = "r"
          Case &HE7: y = "s"
          Case &HE8: y = "t"
          Case &HE9: y = "u"
          Case &HEA: y = "v"
          Case &HEB: y = "w"
          Case &HEC: y = "x"
          Case &HED: y = "y"
          Case &HEE: y = "z"
          Case &HEF: y = "[>]"
          Case &HF0: y = ":"
          Case &HF1: y = "Ä"
          Case &HF2: y = "Ö"
          Case &HF3: y = "Ü"
          Case &HF4: y = "ä"
          Case &HF5: y = "ö"
          Case &HF6: y = "ü"
          Case &HF7: y = "[u]"
          Case &HF8: y = "[d]"
          Case &HF9: y = "[l]"
          Case &HFA: y = "\l"
          Case &HFB: y = "\p"
          Case &HFC: y = "\c": n = True
          Case &HFD: y = "\v": n = True
          Case &HFE: y = "\n"
          Case &HFF: y = "\x"
          Case Else: y = "\h" & IIf(Len(Hex(x)) < 2, "0" & Hex(x), Hex(x))
        End Select
      End If
    End If
    O = O & y
  Next i
  
  O = Replace(O, "\v\h01", "[player]")
  O = Replace(O, "\v\h06", "[rival]")
  O = Replace(O, "\v\h07", "[game]")
  O = Replace(O, "\v\h08", "[team]")
  O = Replace(O, "\v\h09", "[otherteam]")
  O = Replace(O, "\v\h0A", "[teamleader]")
  O = Replace(O, "\v\h0B", "[otherteamleader]")
  O = Replace(O, "\v\h0C", "[legend]")
  O = Replace(O, "\v\h0D", "[otherlegend]")
  
  Sapp2Asc = O

End Function

Private Function IsHex(ByVal hexstring As String) As Boolean
  Dim z As Boolean
  Dim y As Byte
  For i = 1 To Len(hexstring)
    y = Asc(Mid(hexstring, i, 1))
    z = IIf((y > 47 And y < 58) Or (y > 64 And y < 71) Or (y > 96 And y < 103), True, False)
    If z = False Then Exit For
  Next i
  IsHex = z
End Function

Private Function Hex2(ByVal indec As Long, Optional ByVal digits As Byte = 2) As String
  x = Hex(indec)
  Do While Len(x) < digits
    x = "0" & x
  Loop
  Hex2 = x
End Function
