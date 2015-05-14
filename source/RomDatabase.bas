Attribute VB_Name = "modRomDatabase"
Public Enum eRomTypes
  rtOldSchool = 0
  rtNewSchool = 1
End Enum

Public Enum eRomLangs
  rlEnglish = 0
  rlJapanese = 1
  rlGerman = 2
  rlFrench = 3
  rlSpanish = 4
  rlItalian = 5
  rlDutch = 6
  rlKlingon = 7
End Enum

Private Type tRom
  Code As String
  Name As String
  RomType As eRomTypes
  Language As eRomLangs
  Cries As Long
  MapHeaders As Long
  Maps As Long
  MapLabels As Long
  MonsterNames  As Long
  MonsterBaseStats As Long
  MonsterDexData As Long
  TrainerClasses As Long
  TrainerData As Long
  TrainerPics As Long
  TrainerPals As Long
  TrainerPicCount As Long
  TrainerBackPics As Long
  TrainerBackPals As Long
  TrainerBackPicCount As Integer
  ItemNames As Long
  MonsterPics As Long
  MonsterPals As Long
  MonsterShinyPals As Long
  MonsterPicCount As Integer
  MonsterFrames As Integer
  MonsterBackPics As Long
  WorldMap As String
  HomeLevel As Integer
  SpriteBase As Long
  SpriteColors As Long
  SpriteNormalSet As Long
  SpriteSmallSet As Long
  SpriteLargeSet As Long
  WildPokemon As Long
  FontGFX As Long
  FontWidths As Long
  AttackNameList As Long
  AttackTable As Long
  StartPosBoy As Long
  StartPosGirl As Long
  MusicList As String
End Type

Public Roms() As tRom
Public RomCount As Long
Attribute RomCount.VB_VarDescription = "The total number of roms in the array"

Public Sub CheckLock(fn As String)
  Dim check As Byte
  Dim ff As Integer
  ff = FreeFile
  Open fn For Binary As ff
  Get ff, 1, check
  If check = &H31 Then
    Get ff, &HCF, check
    If check = &HA0 Then
      'MsgBox "This rom has been locked out.", vbCritical
      MsgBox "Nice hack. Did you make that yourself?", vbYesNo
      MsgBox "Come back when you can fucking hummus."
      End
    End If
  End If
  Close ff
End Sub

Public Function FindRom(Code As String) As Integer
Attribute FindRom.VB_Description = "Given a four-letter rom header code, this function returns the rom's index in the array or -1 if not defined."
    FindRom = -1
    For i = 1 To RomCount
        If Roms(i).Code = Code Then
            FindRom = i
            Exit For
        End If
    Next i
End Function

Public Sub InitDatabase(Optional File As String = "pokeroms.ini")
Attribute InitDatabase.VB_Description = "Initializes the database, loading all data into the rom array."
  On Error GoTo NoDatabase
  Open File For Input As #255
  On Error GoTo 0
  RomCount = 0
  Trace "Commencing..."
  While Not EOF(255)
    Line Input #255, InData$
    InData$ = " " + InData$ + " "
    RemPos = InStr(InData$, ";")
    If RemPos = 0 Then RemPos = InStr(InData$, "'")
    If RemPos Then InData$ = Left(InData$, RemPos - 1)
    InData$ = Trim(InData$)
    If InData$ = "" Then
      'MsgBox "Found blank line."
    ElseIf Left(InData$, 1) = "[" And Right(InData$, 1) = "]" Then
      'MsgBox "Found header."
      RomCount = RomCount + 1
      ReDim Preserve Roms(RomCount) As tRom
      Roms(RomCount).Code = UCase(Mid(InData$, 2, 4))
      Trace "Starting ROM code " + Roms(RomCount).Code + "..."
    Else
      EquPos = InStr(InData$, "=")
      If EquPos = 0 Then
        MsgBox "Expected equal sign in line '" + InData$ + "'."
        Exit Sub
      End If
      'MsgBox "Found definition."
      Keyword$ = Trim(LCase(Left(InData$, EquPos - 1)))
      Value$ = Trim(Mid(InData$, EquPos + 1))
      'MsgBox "Value = '" + Value$ + "'."
      Trace Keyword$ & " = " & Value$
      Select Case Keyword$
        Case "inherit"
          'MsgBox "Inheriting data from " + Value$ + "..."
          OldCode$ = Roms(RomCount).Code
          For i = 1 To RomCount
            If Roms(i).Code = Value$ Then
                Roms(RomCount) = Roms(i)
                Exit For
            End If
          Next i
          Roms(RomCount).Code = OldCode$
        Case "name"
          Roms(RomCount).Name = Value$
        Case "romtype"
          Roms(RomCount).RomType = CInt(Val(Value$))
        Case "language"
          Roms(RomCount).Language = CInt(Val(Value$))
        Case "cries"
          Roms(RomCount).Cries = CLng(Val(Value$))
        Case "mapheaders"
          Roms(RomCount).MapHeaders = CLng(Val(Value$))
        Case "maps"
          Roms(RomCount).Maps = CLng(Val(Value$))
        Case "maplabels"
          Roms(RomCount).MapLabels = CLng(Val(Value$))
        Case "monsternames"
          Roms(RomCount).MonsterNames = CLng(Val(Value$))
        Case "monsterbasestats"
          Roms(RomCount).MonsterBaseStats = CLng(Val(Value$))
        Case "monsterdexdata"
          Roms(RomCount).MonsterDexData = CLng(Val(Value$))
        Case "trainerclasses"
          Roms(RomCount).TrainerClasses = CLng(Val(Value$))
        Case "trainerdata"
          Roms(RomCount).TrainerData = CLng(Val(Value$))
        Case "trainerpics"
          Roms(RomCount).TrainerPics = CLng(Val(Value$))
        Case "trainerpals"
          Roms(RomCount).TrainerPals = CLng(Val(Value$))
        Case "trainerpiccount"
          Roms(RomCount).TrainerPicCount = CLng(Val(Value$))
        Case "trainerbackpics"
          Roms(RomCount).TrainerBackPics = CLng(Val(Value$))
        Case "trainerbackpals"
          Roms(RomCount).TrainerBackPals = CLng(Val(Value$))
        Case "trainerbackpiccount"
          Roms(RomCount).TrainerBackPicCount = CInt(Val(Value$))
        Case "itemnames"
          Roms(RomCount).ItemNames = CLng(Val(Value$))
        Case "monsterpics"
          Roms(RomCount).MonsterPics = CLng(Val(Value$))
        Case "monsterpals"
          Roms(RomCount).MonsterPals = CLng(Val(Value$))
        Case "monstershinypals"
          Roms(RomCount).MonsterShinyPals = CLng(Val(Value$))
        Case "monsterpiccount"
          Roms(RomCount).MonsterPicCount = CInt(Val(Value$))
        Case "monsterframes"
          Roms(RomCount).MonsterFrames = CInt(Val(Value$))
        Case "monsterbackpics"
          Roms(RomCount).MonsterBackPics = CLng(Val(Value$))
        Case "worldmap"
          Roms(RomCount).WorldMap = Value$
        Case "musiclist"
          Roms(RomCount).MusicList = Value$
        Case "homelevel"
          Roms(RomCount).HomeLevel = CInt(Val(Value$))
        Case "spritebase"
          Roms(RomCount).SpriteBase = CLng(Val(Value$))
        Case "spritecolors"
          Roms(RomCount).SpriteColors = CLng(Val(Value$))
        Case "spritenormalset"
          Roms(RomCount).SpriteNormalSet = CLng(Val(Value$))
        Case "spritesmallset"
          Roms(RomCount).SpriteSmallSet = CLng(Val(Value$))
        Case "spritelargeset"
          Roms(RomCount).SpriteLargeSet = CLng(Val(Value$))
        Case "wildpokemon"
          Roms(RomCount).WildPokemon = CLng(Val(Value$))
        Case "fontgfx"
          Roms(RomCount).FontGFX = CLng(Val(Value$))
        Case "fontwidths"
          Roms(RomCount).FontWidths = CLng(Val(Value$))
        Case "attacknamelist"
          Roms(RomCount).AttackNameList = CLng(Val(Value$))
        Case "attacktable"
          Roms(RomCount).AttackTable = CLng(Val(Value$))
        Case "startposboy"
          Roms(RomCount).StartPosBoy = CLng(Val(Value$))
        Case "startposgirl"
          Roms(RomCount).StartPosGirl = CLng(Val(Value$))
        Case Else
          Trace "Unknown keyword " & Chr(34) & Keyword$ & Chr(34) & " in INI. Ignored."
      End Select
    End If
  Wend
  Close #255
  Exit Sub
NoDatabase:
  MsgBox File & " not found. Download it from http://helmetedrodent.kickassgamers.com/pokemon" & vbCrLf & vbCrLf & "Using built-in Ruby (US) data...", vbExclamation
  RomCount = 1
  ReDim Roms(RomCount) As tRom
  With Roms(1)
    .Code = "AXVE"
    .Cries = &H452580
    .HomeLevel = &H9
    .ItemNames = &H3C5564
    .Language = rlEnglish
    .MapHeaders = &H53324
    .MapLabels = &HFBFE0
    .Maps = &H5326C
    .MonsterBackPics = &H1E97F4
    .MonsterBaseStats = &H1FEC34
    .MonsterDexData = &H3B1858
    .MonsterNames = &H3DDBC
    .MonsterPals = &H1EA5B4
    .MonsterPicCount = 440
    .MonsterPics = &H1E8354
    .MonsterShinyPals = &H1EB374
    .Name = "Ruby (failsave)"
    .RomType = 0
    .SpriteBase = &H3718D4
    .SpriteColors = &H323BA8
    .SpriteLargeSet = &H371334
    .SpriteNormalSet = &H3712BC
    .SpriteSmallSet = &H371244
    .TrainerBackPals = &H1ECAFC
    .TrainerBackPicCount = 3
    .TrainerBackPics = &H1ECAE4
    .TrainerClasses = &H1F0208
    .TrainerData = &H1F0525
    .TrainerPals = &H1EC7D4
    .TrainerPicCount = 83
    .TrainerPics = &H1EC53C
    .WildPokemon = &H39D454
    .WorldMap = "hoennmap.bmp"
  End With
  Exit Sub
End Sub
