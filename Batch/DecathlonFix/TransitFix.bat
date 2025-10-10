@echo off
setlocal ENABLEDELAYEDEXPANSION
REM SETLOCAL ENABLEDELAYEDEXPANSION
@echo off
rem We need a path with the 7zip tool to extract the files.
rem set path="C:\Program Files\7-Zip";%path%
rem set path=%cd%;%path%

set path="C:\Tools";%path% rem 7zip files must be copied to this path

rem strore variables of path, name, full path,etc
set "Fullpath=%1"
set "PathOnly=%~dp1"
set "NameOnly=%~n1"
set "NameExtOnly=%~nx1"
set "ExtOnly=%~x1"

rem store the value for English found or French found
set "French="
set "English="

rem === LOCK FILE SYSTEM ===
set "LockFile=%Fullpath%.lock"
if exist "%LockFile%" (
    echo ====================================================
    echo [!] Another process is already working on:
    echo     %Fullpath%
    echo     Wait until it finishes or remove the .lock if needed.
    echo ====================================================
    pause
    goto EXIT
)
echo %DATE% %TIME% > "%LockFile%"
rem === END LOCK ===

rem If not a .PPF file Check if TPF file
if NOT "%ExtOnly%"==".PPF" (
		goto TPFFILE
)

rem Create the temp path, if it exists delete it, if not create it
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

rem check if some language has been found if not exit
if "%French%"=="" if "%English%"=="" goto NOTAPPLICABLE

rem Remove all files in temp folder in silent mode
del /q "%TempPATH%\*.*"

rem Extract the Project file .PPJ to temp folder
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
REM END OF PPF PROCESSING


:TPFFILE
if NOT "%ExtOnly%"==".tpf" (
	cls
	echo FILE CAN'T BE PROCESSED, IT IS NOT A .PPF OR .TPF FILE
	pause
	goto EXIT
)

rem === NEW: If .tpf but NAME does NOT end with _SEPROFIX -> delete lock, show message and exit ===
set "NameSinSepro=%NameOnly:_SEPROFIX=%"
if "%NameSinSepro%"=="%NameOnly%" (
    rem Remove lock file before exiting since we won't process this file
    if exist "%LockFile%" del "%LockFile%"

    cls
    rem Build expected name with .tpf explicitly
    set "ExpectedName=%NameOnly%_SEPROFIX.tpf"
    echo This is not a XXXXXXXXXX_SEPROFIX.tpf
    echo The name should be !ExpectedName! and it is %NameExtOnly%
    echo The file is not a SEPROFIX file or Seprofix macro was already run over this file.
    pause
    goto EXIT
)
rem === END NEW CHECK ===

rem The temp path, if exists delete it, if not 7zip will create it
set "TempPATH=%PathOnly%%NameOnly%_TEMP"

if exist "%TempPATH%" (
    ECHO Sub folder found. Deleting . . . 
    RD /S /Q "%TempPATH%"
)

rem Extract the Project file .PPJ to temp folder
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
rem --- REMOVE LOCK FILE IF EXISTS ---
if exist "%LockFile%" del "%LockFile%"
rem exit
endlocal
