@echo off

echo Compiling resources...
echo - BASEEDIT
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo baseedit.res baseedit.rc
echo - BEWILDERED
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo bewildered.res bewildered.rc
echo - DEXTER
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo dexter.res dexter.rc
echo - ELITEMAP
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo elitemap.res elitemap.rc
echo - FONTED
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo fonted.res fonted.rc
echo - PATTED
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo patted.res patted.rc
echo - PET
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo pet.res pet.rc
echo - RSBALL
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo rsball.res rsball.rc
echo - SCRIPTED
"C:\Program Files\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /fo scripted.res scripted.rc

echo Compiling projects...
echo - BASEEDIT
"C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" baseedit.vbp /make
echo - BEWILDERED
"C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" bewildered.vbp /make
echo - DEXTER
"C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" dexter.vbp /make
echo - ELITEMAP
"C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" elitemap.vbp /make
echo - FONTED
"C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" fonted.vbp /make
echo - PATTED
"C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" patted.vbp /make
echo - PET
"C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" pet.vbp /make
echo - RSBALL
"C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" rsball.vbp /make
echo - SCRIPTED
"C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" scripted.vbp /make
echo - SPREAD
"C:\Program Files\Microsoft Visual Studio\VB98\vb6.exe" spred.vbp /make

echo Packing files...
echo - BASEEDIT
upx -9 -q ..\baseedit.exe > nul
echo - BEWILDERED
upx -9 -q ..\bewildered.exe > nul
echo - DEXTER
upx -9 -q ..\dexter.exe > nul
echo - ELITEMAP
upx -9 -q ..\elitemap.exe > nul
echo - FONTED
upx -9 -q ..\fonted.exe > nul
echo - PATTED
upx -9 -q ..\patted.exe > nul
echo - PET
upx -9 -q ..\pet.exe > nul
echo - RSBALL
upx -9 -q ..\rsball.exe > nul
echo - SCRIPTED
upx -9 -q ..\scripted.exe > nul
echo - SPREAD
upx -9 -q ..\spread.exe > nul

echo All done.
pause
