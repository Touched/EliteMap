!define WITHBONUSPROGRAMS

!define VER_MAJOR 3
!define VER_MINOR 7

!ifdef HAVE_UPX
  !packhdr tmp.dat "upx\upx --best --compress-icons=1 tmp.dat"
!endif

!include "MUI.nsh"

Name "EliteMap"
OutFile elitemap${VER_MAJOR}${VER_MINOR}.exe

!define MUI_HEADERIMAGE
!define MUI_ABORTWARNING
!define MUI_COMPONENTSPAGE_NODESC

!define MUI_ICON nsis_icon.ico
!define MUI_HEADERIMAGE_BITMAP nsis_header.bmp
!define MUI_WELCOMEFINISHPAGE_BITMAP nsis_sidebar.bmp
!define MUI_COMPONENTSPAGE_CHECKBITMAP nsis_checks.bmp
!define MUI_FINISHPAGE_SHOWREADME welcome.mht
!define MUI_FINISHPAGE_RUN elitemap.exe
!define MUI_FINISHPAGE_RUN_TEXT "Run EliteMap now"
!define MUI_FINISHPAGE_LINK "Visit The Helmeted Rodent website"
!define MUI_FINISHPAGE_LINK_LOCATION "http://helmetedrodent.kickassgamers.com"
!define MUI_FINISHPAGE_NOREBOOTSUPPORT
!define MUI_WELCOMEPAGE_TEXT "This wizard will guide you through the installation of EliteMap.\r\n\r\nRemember that you must place any Pokémon ROM files you intend to use in the same folder as EliteMap.\r\n\r\nClick next to continue."

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_LANGUAGE "English"

AutoCloseWindow false
ShowInstDetails show

SetOverwrite on
SetDateSave on

;BGGradient babbe6 ffffff

InstallDir $EXEDIR\EliteMap
CompletedText "Yatta ze!"

SubSection "!EliteMap"
Section "EliteMap (req'd)"
  SectionIn RO
  SetOutPath $INSTDIR
  File elitemap.exe
  File scripted.exe
  File welcome.mht
  SetOverwrite ifnewer
  File pokeroms.ini
  File rs_songs.txt
  File fr_songs.txt
  File em_songs.txt
  File "Hidden Functions.png"
  ;CreateShortCut "$FAVORITES\The Helmeted Rodent.lnk" "http://helmetedrodent.kickassgamers.com"
SectionEnd
Section "Supporting DLLs"
  File "Lunar Compress.dll"
SectionEnd
Section "Toolchain"
  File baseedit.exe
  File bewildered.exe
  File dexter.exe
  File rsball.exe
  File spread.exe
  File patted.exe
  File pet.exe
  ;File trained.exe
  File fonted.exe
SectionEnd
Section "Beta tools"
  File mapedit.exe
SectionEnd
Section "Nifty IPS stuff"
  File "GBPlayer intro.ips"
  File "No Drugs intro.ips"
  File "Edited in EliteMap.ips"
SectionEnd
Section "Media Library"
  SectionIn 1 2 
  SetOutPath $INSTDIR
  File hoennmap.bmp
  File kantomap.bmp
  SetOutPath $INSTDIR\Media
  File Media\*.*
SectionEnd
SubSectionEnd
Section "Rubikon"
  SetOutPath $INSTDIR
  File rkc.exe
  File rubikon.dat
  File commands.html
  File std.rbh
  File stditems.rbh
  File stdpoke.rbh
  File richboy.rbc
  File richboy2.rbc
  File tutorial.rbc
  File rubikonicon.gif
  File tutorial.html
  File rubikon.syn
SectionEnd

!ifdef WITHBONUSPROGRAMS
Section "Bonus Programs"
  SetOutPath $INSTDIR
  File lips.exe
  File pokepic.exe
  File snesedit.exe
  File tlp.exe
  File PokeCryGUI.exe
  File unLZ.GBA.exe
SectionEnd
!endif
