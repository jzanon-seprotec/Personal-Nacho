 REM SETLOCAL ENABLEDELAYEDEXPANSION
@echo off
rem Necesitamos un path con la herramienta 7zip para descomprimir los archivos.
rem set path="C:\Program Files\7-Zip";%path%
rem set path=%cd%;%path%

set path="C:\Tools";%path% rem 7zip files must be copied to this path

rem strore variables of path, name, full path,etc
set "Fullpath=%1"
set "PathOnly=%~dp1"
set "NameOnly=%~n1"
set "NameExtOnly=%~nx1"
set "ExtOnly=%~x1"


rem strores the value for English found or French found
set "French="
set "English="


rem If not a .PPF file Check if TPF file
if NOT "%ExtOnly%"==".PPF" (
		goto TPFFILE
)


rem Creamos el path temporal, si ya existe lo borra si no lo crea
set "TempPATH=%PathOnly%%NameOnly%_TEMP"
if exist "%TempPATH%" (
    @ECHO Sub folder found. Deleting . . .  
    RD /S /Q "%TempPATH%"
)


rem Extract the files .EN1 and .FR1 to a temp folder; don't care if they are .FR1 or .EN1
7z e %Fullpath% -o"%TempPATH%" *.FR1 *.EN1

REM wait 3 seconds
timeout /t 3 /nobreak >nul 

rem Renames the files in zip
FOR %%f in ("%TempPATH%"\*) DO (
if "%%~xf" == ".FR1" 7z rn %Fullpath% "%%~nxf" "%%~nf".FRA & set "French=y"
if "%%~xf" == ".EN1" 7z rn %Fullpath% "%%~nxf" "%%~nf".ENG & set "English=y"
)


rem chech if some language has been found if not exit
if "%French%"=="" if "%English%"=="" goto NOTAPPLICABLE


rem Remove all files in temp folder in silent mode
del /q "%TempPATH%\*.*"


rem Extract the Prject file .PPJ to temp folder
7z e %Fullpath% -o"%TempPATH%" *.PRJ


rem Changes the language code in project file .PRJ and project name to take control ProjectName = ProjectName_SEPROFIX
FOR %%f in ("%TempPATH%"\*) DO (
set FullPathNameUP="%TempPATH%\%%~nxf"
set PRJName="%%~nf"
fart "%TempPATH%\%%~nxf" ProjectName=%%~nf ProjectName="%%~nf_SEPROFIX"
fart "%TempPATH%\%%~nxf" SourceLanguage=31753 SourceLanguage=2057
fart "%TempPATH%\%%~nxf" SourceLanguage=31756 SourceLanguage=1036
)


rem updates the .PRJ file in zip 
7z u %Fullpath% %FullPathNameUP% 


rem Rename the .PRJ adding _SEPROFIX at the end
7z rn %Fullpath% "%PRJName%".PRJ "%PRJName%"_SEPROFIX.PRJ

rem renames original zip to _SEPROFIX
rename %Fullpath% "%NameOnly%_SEPROFIX%ExtOnly%"


rem Delete subdir
RMDIR "%TempPATH%" /S /Q


GOTO EXIT
REM FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN FIN 


:TPFFILE
if NOT "%ExtOnly%"==".tpf" (
	cls
	echo FILE CAN'T BE PROCESSED, IT IS NOT A .PPF OR .TPF FILE
	pause
	goto EXIT
)

rem El path temporal, si ya existe lo borra si no 7zip lo crea
set "TempPATH=%PathOnly%%NameOnly%_TEMP"

if exist "%TempPATH%" (
    ECHO Sub folder found. Deleting . . . 
    RD /S /Q "%TempPATH%"
)


rem Extract the Prject file .PPJ to temp folder
7z e %Fullpath% -o"%TempPATH%" *.PRJ

rem Changes the language code in project file .PRJ and removes _SEPROFIX in project name: ProjectName = ProjectName_SEPROFIX

FOR %%f in ("%TempPATH%"\*) DO (
set "NameWithSepro=%%~nf"
set "FullPathNameUP=%TempPATH%\%%~nxf"
)


rem Removes the _SEPROFIX to the name
set NameSinSepro=%NameWithSepro:_SEPROFIX=%
fart "%FullPathNameUP%" ProjectName="%NameWithSepro%" "ProjectName=%NameSinSepro%"
fart "%FullPathNameUP%" SourceLanguage=2057 SourceLanguage=31753
fart "%FullPathNameUP%" SourceLanguage=1036 SourceLanguage=31756


rem updates the file in zip
7z u %Fullpath% "%FullPathNameUP%" 


rem Rename the .PRJ REMOVING _SEPROFIX at the end
7z rn %Fullpath% "%NameWithSepro%.PRJ" "%NameSinSepro%.PRJ" 


rem renames original zip REMOVING _SEPROFIX
set NameOnly=%NameOnly:_SEPROFIX=%
rename %Fullpath% "%NameOnly%%ExtOnly%"

rem Delete subdir
RMDIR "%TempPATH%" /S /Q

GOTO EXIT

:NOTAPPLICABLE
cls
echo[
echo[
echo[
echo[
echo[
echo[
echo[
echo[
echo[
echo[
echo[
echo[
echo[
@echo  IT IS NOT NECCESARY TO CHANGE ANYTHING ON THIS FILE; WILL WORK IN STUDIO AS IS
echo[
echo[
echo[
pause


:EXIT
rem exit
