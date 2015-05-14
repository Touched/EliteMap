!define WITHSOURCE

!define VER_MAJOR 3
!define VER_MINOR 7

!ifdef HAVE_UPX
  !packhdr tmp.dat "upx\upx --best --compress-icons=1 tmp.dat"
!endif

!include "MUI.nsh"

Name "EliteMap source"
OutFile elitemap${VER_MAJOR}${VER_MINOR}src.exe

!define MUI_HEADERIMAGE
!define MUI_ABORTWARNING
!define MUI_COMPONENTSPAGE_NODESC

!define MUI_ICON nsis_icon.ico
!define MUI_HEADERIMAGE_BITMAP nsis_header.bmp
!define MUI_WELCOMEFINISHPAGE_BITMAP nsis_sidebar.bmp
!define MUI_COMPONENTSPAGE_CHECKBITMAP nsis_checks.bmp

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

AutoCloseWindow false
ShowInstDetails show

SetOverwrite on
SetDateSave on

InstallDir $EXEDIR\EliteMap
CompletedText "Yatta ze!"

SubSection "!Elitemap Source code"
Section "EM and Toolchain source"
  SetOutPath $INSTDIR\source
  File source\*.vbp
  File source\*.rc
  File source\*.res
  File source\*.frm
  File source\*.frx
  File source\*.bas
  File source\*.cls
  File source\*.ctl
  File source\*.ctx
  File source\make.bat
  File source\vbtrace6.exe
  File source\upx.exe
  SetOutPath $INSTDIR\source\assets
  File source\assets\*.*
SectionEnd
Section "Rubikon source"
  SetOutPath $INSTDIR\rkc
  File rkc\*.vbp
  File rkc\*.res
  File rkc\*.frm
  File rkc\*.frx
  File rkc\*.bas
  File rkc\*.ico
SectionEnd
Section "NSIS installer scripts"
  SetOutPath $INSTDIR
  File emsource.nsi
  File emredist.nsi
  File nsis_sidebar.bmp
  File nsis_header.bmp
  File nsis_checks.bmp
  File nsis_icon.ico
SectionEnd
SubSectionEnd
Function .onInit
  MessageBox MB_OK "This package contains the full source code. Not for redistribution."
FunctionEnd

