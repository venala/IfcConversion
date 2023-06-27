<# : chooser.bat
:: Solutions for opening explorer:
:: https://stackoverflow.com/a/15885133/1683264

@echo off

setlocal

set ifcpath=""
set outputfolder=""
set filename=""
set outputpath=""
set xmlpath=""
set threads=%NUMBER_OF_PROCESSORS%
set program="C:\Users\vwell\Desktop\IfcConvert"

echo Right now it is assumed, that ifcConvert is at this location:
echo %program%
echo If this is wrong, please edit this program and enter the correct path.

pause

echo Choose the IFC file you want to convert.

for /f "delims=" %%I in ('powershell -noprofile "iex (${%~f0} | out-string)"') do (
    set ifcpath=%%~I
)

echo You chose %ifcpath%.

echo Now choose the output folder.

set "psCommand="(new-object -COM 'Shell.Application')^
.BrowseForFolder(0,'Please choose an output folder.',0,0).self.path""

for /f "usebackq delims=" %%I in (`powershell %psCommand%`) do set "outputfolder=%%I"

echo You chose %outputfolder% as output folder.

call:removePath %ifcpath%

set filename=%filename:~0,-4%

set outputpath="%outputfolder%\%filename%_obj.obj"
set xmlpath="%outputfolder%\%filename%_xml.xml"

echo %outputpath%

::Convert to .obj and .mtl
%program% -y -j %threads% --use-element-guids --y-up %ifcpath% %outputpath% --exclude entities IfcOpeningElement

::Convert to .xml
%program% -y -j %threads% --use-element-guids --y-up %ifcpath% %xmlpath% --exclude entities IfcOpeningElement

goto :EOF

:removePath
set path=%~1
set reducedPath=%path:*\=%
if %reducedPath%==%path% (set filename=%path%) else call:removePath %reducedPath%
goto:eof

: end Batch portion / begin PowerShell hybrid chimera #>

Add-Type -AssemblyName System.Windows.Forms
$f = new-object Windows.Forms.OpenFileDialog
$f.InitialDirectory = pwd
$f.Title = "Choose ifc file..."
$f.Filter = "ifc Files (*.ifc)|*.ifc|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
$f.ShowHelp = $true
$f.Multiselect = $false
[void]$f.ShowDialog()
if ($f.Multiselect) { $f.FileNames } else { $f.FileName }