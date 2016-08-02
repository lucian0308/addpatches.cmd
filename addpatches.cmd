REM addpatches.cmd
REM Copyright (C) 2016 lea2000
REM 
REM This program is free software; you can redistribute it
REM and/or modify it under the terms of the GNU Lesser General
REM Public License as published by the Free Software Foundation;
REM either version 3 of the License, or (at your option) any
REM later version.
REM 
REM This program is distributed in the hope that it will be useful,
REM but WITHOUT ANY WARRANTY; without even the implied warranty of
REM MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
REM Lesser General Public License for more details.
REM 
REM You should have received a copy of the GNU Lesser General Public
REM License along with this program; if not, write to the Free Software
REM Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

@echo off
cls
setlocal EnableDelayedExpansion


REM COMMON SETTINGS START

REM If True, the script only shows, what it would do
set dryrun=False

REM If True, the patches will be installed to the online system - see ONLINE SETTINGS
REM Otherwise the script will try to use the !install_wim! - see OFFLINE SETTINGS
set online=False

REM Location of the msu files
set patches_dir=%~dp0msu

REM Everything will be copied to this location.
REM Directory will be removed, before and afterwards
set work_dir=%SYSTEMDRIVE%\addpatches

REM If True and links.txt exists, not existing files will be downloaded
set downloadpatches=True
REM If True, all patches in the links.txt will be redownloaded
set force_downloadpatches=False

REM If True, always install all patches
REM Otherwise only KBs which are seems to be missing 
set force_install=False

REM If True and dism is used, the script will try to install all patches in a single dism session
REM Otherwise one session for each job/dir
set dism_installallatonce=True

REM If True, log it a package is already downloaded/installed
set log_found=False

REM If True, the script will pause at the end
set pause_end=True

REM COMMON SETTINGS END


REM ONLINE SETTINGS START

REM If True, this script will stop if a wusa installation exits not with 0 or 3010
set wusa_exitonerror=False

REM If True, the patches will added with dism instead of wusa
REM You wont have a history in the windows update gui
set online_dism=True

REM If True, it will start a new window, where you can follow the WindowsUpdate.log with Get-Content from powershell
REM You will see some action there, if wusa is used to install the patches (online_dism=False)
set follow_wulog=False

REM ONLINE SETTINGS END


REM OFFLINE SETTINGS START

REM The .wim-file of the offline system
set install_wim=%~dp0cd\sources\install.wim
REM The index in the .wim-file
set install_wim_index=1

REM If True and !imagex_exe! exists, the script will work with a copy of the select !install_wim_index! and afterwards append this copy to the !install_wim!
REM You need to have installed the Windows 10 Assessment and Deployment Kit (ADK) on the build system
set install_wim_copy=True

REM If True, !osdcimg_exe! exists and !cd_root!\boot\etfsboot.com found the script will try to create a ready to use iso file at the end
REM You need to have installed the Windows 10 Assessment and Deployment Kit (ADK) on the build system
set install_wim_createiso=True
REM The root of the extracted original iso file
set cdroot_dir=%~dp0cd
REM Where to save the iso
set isooutput_dir=%~dp0

REM OFFLINE SETTINGS END


net session >nul 2>&1
if not %ERRORLEVEL% == 0 (
	echo Error - this command prompt is not elevated
	goto end
)

set dism_exe=%PROGRAMFILES(x86)%\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\%PROCESSOR_ARCHITECTURE%\DISM\dism.exe
if not exist "!dism_exe!" set dism_exe=%PROGRAMFILES%\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\%PROCESSOR_ARCHITECTURE%\DISM\dism.exe
if not exist "!dism_exe!" set dism_exe=%SYSTEMROOT%\System32\dism.exe

set imagex_exe=%PROGRAMFILES(x86)%\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\%PROCESSOR_ARCHITECTURE%\DISM\imagex.exe
if not exist "!imagex_exe!" set imagex_exe=%PROGRAMFILES%\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\%PROCESSOR_ARCHITECTURE%\DISM\imagex.exe

set osdcimg_exe=%PROGRAMFILES(x86)%\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\%PROCESSOR_ARCHITECTURE%\Oscdimg\oscdimg.exe
if not exist "!osdcimg_exe!" set osdcimg_exe=%PROGRAMFILES%\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\%PROCESSOR_ARCHITECTURE%\Oscdimg\oscdimg.exe

set timeout_exe=%SYSTEMROOT%\System32\timeout.exe

set find_exe=%SYSTEMROOT%\System32\find.exe
set wmic_exe=%SYSTEMROOT%\System32\wbem\WMIC.exe

set systeminfo_exe=%SYSTEMROOT%\System32\systeminfo.exe
set bitsadmin_exe=%SYSTEMROOT%\System32\bitsadmin.exe
set expand_exe=%SYSTEMROOT%\System32\expand.exe

set wusa_exe=%SYSTEMROOT%\SysNative\wusa.exe
if not exist "!wusa_exe!" set wusa_exe=%SYSTEMROOT%\System32\wusa.exe

set powershell_exe=%SYSTEMROOT%\System32\WindowsPowerShell\v1.0\powershell.exe

set name=?
set arch=?
set producttype=?
set version=?

set wiminfo=!work_dir!\wiminfo.txt
set mount=!work_dir!\mount

set temp_install_wim=!install_wim!
set temp_install_wim_index=!install_wim_index!

set installedmsus=0

set dism_installedpackages=!work_dir!\dism_installedpackages.txt
set systeminfo_installedpackages=!work_dir!\systeminfo_installedpackages.txt
set dism_features=!work_dir!\dism_features.txt

set patches_install=!work_dir!\install
set jobs=!work_dir!\jobs.txt

set wu_reg_path=HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU
set wu_reg_name=NoAutoUpdate
set wu_reg_tempvalue=1
set wu_reg_savedvalue=?

set exitcode=0

net session >nul 2>&1
if not !ERRORLEVEL!==0 (
	echo Error - this command prompt is not elevated
	goto end
)


:deleteold

"!dism_exe!" /Get-MountedWimInfo | !find_exe! /i "!mount!" >nul 2>&1
if !ERRORLEVEL!==0 (
	echo|set /p=Discarding and unmounting old mount dir... 
	"!dism_exe!" /Unmount-Wim /MountDir:"!mount!" /Discard /English >nul 2>&1
	set exitcode=!ERRORLEVEL!
	echo !exitcode!
	
	if not !exitcode!==0 (
		echo Error discarding and unmounting old mount dir
		goto end
	)
)

if exist "!work_dir!" (
	echo|set /p=Deleting old work dir ... 
	rmdir /s /q "!work_dir!" >nul 2>&1
	set exitcode=!ERRORLEVEL!
	if not exist "!work_dir!" set exitcode=0
	echo !exitcode!
	
	if not !exitcode!==0 (
		echo Error deleting old work dir
		goto end
	)
)

md "!work_dir!" >nul 2>&1

:info

echo.
echo Start time     : !TIME:~0,2!:!TIME:~3,2!

echo Dry run        : !dryrun!
echo.

echo.
if "!online!"=="True" (

	echo Used system    : Online
	
	if "!online_dism!"=="True" (
		echo Install method : dism
	) else (
		echo Install method : wusa
		echo Exit on error  : !wusa_exitonerror!
	)
	
	echo.

) else (

	echo Used system    : Offline
	
	echo WIM file       : !install_wim!
	echo WIM Index      : !install_wim_index!
	echo Create Copy    : !install_wim_copy!
	echo Create ISO     : !install_wim_createiso!
	
	echo.

)

echo.
echo Work dir       : !work_dir!
echo Patches dir    : !patches_dir!

set downloadmethod=
if "!downloadpatches!"=="True" (
	set downloadmethod=only missing
	if !force_downloadpatches!=="True" set downloadmethod=all
)
if not "!downloadmethod!"=="" echo Download       : !downloadmethod!

set installmethod=only missing
if !force_install!=="True" set installmethod=all
echo Install        : !installmethod!

"!timeout_exe!" /t 30

:configurewu
if not "!online!"=="True" goto getofflineinfo

echo.

sc query wuauserv| !find_exe! "RUNNING" >nul 2>&1
if !ERRORLEVEL!==0 (

	echo|set /p=!TIME:~0,2!:!TIME:~3,2! Stopping wuauserv... 
	net stop wuauserv >nul 2>&1
	set exitcode=!ERRORLEVEL!
	echo !exitcode!

	if not !exitcode!==0 (
		echo Error stopping wuauserv
		goto end
	)
	
	"!timeout_exe!" /t 3 >nul 2>&1
)

if "!online_dism!"=="True" (

	sc qc wuauserv | !find_exe! "DISABLED" >nul 2>&1
	if not !ERRORLEVEL!==0 (
		echo|set /p=!TIME:~0,2!:!TIME:~3,2! Setting starttype for wuauserv to disabled... 
		sc config wuauserv start= disabled >nul 2>&1
		set exitcode=!ERRORLEVEL!
		echo !exitcode!

		if not !exitcode!==0 (
			echo Error setting starttype for wuauserv to disabled
			goto end
		)
	)
	
	goto getonlineinfo

)

if "!follow_wulog!"=="True" (
	if not exist "!powershell_exe!" (
		echo Error powershell.exe not found
		goto end
	)
	start "" "!powershell_exe!" -Command Get-Content -Wait -Path $env:SYSTEMROOT\WindowsUpdate.log
)

reg query "!wu_reg_path!" /v "!wu_reg_name!" >nul 2>&1
if !ERRORLEVEL!==0 for /f "tokens=3" %%a in ('reg query "!wu_reg_path!" /v "!wu_reg_name!" ^| !find_exe! /i "!wu_reg_name!"') do set /a wu_reg_savedvalue=%%a + 0

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Setting !wu_reg_name! temporarily to !wu_reg_tempvalue!... 
reg add "!wu_reg_path!" /v !wu_reg_name! /t REG_DWORD /d !wu_reg_tempvalue! /f >nul 2>&1
set exitcode=!ERRORLEVEL!
echo !exitcode!

if not !exitcode!==0 (
	echo Error setting !wu_reg_name! temporarily to !wu_reg_tempvalue!
	goto end
)

sc qc wuauserv | !find_exe! "DISABLED" >nul 2>&1
if !ERRORLEVEL!==0 (
	echo|set /p=!TIME:~0,2!:!TIME:~3,2! Setting starttype for wuauserv to demand... 
	sc config wuauserv start= demand >nul 2>&1
	set exitcode=!ERRORLEVEL!
	echo !exitcode!

	if not !exitcode!==0 (
		echo Error setting starttype for wuauserv to demand
		goto end
	)
)

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Starting wuauserv... 
net start wuauserv >nul 2>&1
set exitcode=!ERRORLEVEL!
echo !exitcode!

if not !exitcode!==0 (
	echo Error starting wuauserv
	goto end
)

:getonlineinfo

echo.
echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating name... 
for /f "tokens=2 delims==|" %%a in ('!wmic_exe! os get Name /value ^| !find_exe! /i "Name="') do set name=%%a
echo !name!

if "!name!"=="?" (
	echo Error determinating name
	goto end
)

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating architecture... 
set arch=x86
if exist "%PROGRAMFILES(x86)%" set arch=x64
echo !arch!

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating product type... 
set producttype=Desktop
!wmic_exe! os get ProductType | !find_exe! "2" >nul 2>&1
if !ERRORLEVEL!==0 set producttype=Server
!wmic_exe! os get ProductType | !find_exe! "3" >nul 2>&1
if !ERRORLEVEL!==0 set producttype=Server
echo !producttype!

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating version... 
for /f "tokens=2 delims==" %%a in ('!wmic_exe! os get Version /value ^| !find_exe! /i "Version="') do set version=%%a
echo !version!

if "!version!"=="?" (
	echo Error determinating version
	goto end
)

goto getenabledfeatures

:getofflineinfo

echo.

if not exist "!install_wim!" (
	echo Error couldn't find !install_wim!
	goto end
)

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Getting info of wim file... 
"!dism_exe!" /Get-WimInfo /WimFile:"!install_wim!" /Index:!install_wim_index! /English > "!wiminfo!"
set exitcode=!ERRORLEVEL!
echo !exitcode!
if not !exitcode!==0 (
	echo Error getting info from index !install_wim_index! of wim file
	goto end
)

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating name... 
for /f "tokens=2*" %%a in ('type !wiminfo! ^| !find_exe! /i "Name :"') do set name=%%b
echo !name!
if "!name!"=="?" (
	echo Error determinating name
	goto end
)

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating architecture... 
for /f "tokens=2*" %%a in ('type !wiminfo! ^| !find_exe! /i "Architecture :"') do set arch=%%b
echo !arch!
if "!arch!"=="?" (
	echo Error determinating architecture
	goto end
)

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating product type... 
for /f "tokens=2*" %%a in ('type !wiminfo! ^| !find_exe! /i "ProductType :"') do set producttype=%%b
if "!producttype!"=="?" (
	echo Error determinating product type
	goto end
)
if "!producttype!"=="WinNT" set producttype=Desktop
if "!producttype!"=="ServerNT" set producttype=Server
echo !producttype!

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating version... 
for /f "tokens=2*" %%a in ('type !wiminfo! ^| !find_exe! /i "Version :"') do set version=%%b
echo !version!
if "!version!"=="?" (
	echo Error determinating product version
	goto end
)

del /f /q "!wiminfo!" >nul 2>&1

:mountwim

echo.

md "!mount!" >nul 2>&1

if "!install_wim_copy!"=="True" (
	
	if not exist "!imagex_exe!" (
		echo Error imagex.exe not found
		goto end
	)
	
	if "!dryrun!"=="True" (
		echo|set /p=!TIME:~0,2!:!TIME:~3,2! Would export !install_wim_index! of wim file to work dir... 
	) else (
		
		set temp_install_wim=!work_dir!\temp_install.wim
		set temp_install_wim_index=1
		echo|set /p=!TIME:~0,2!:!TIME:~3,2! Exporting !install_wim_index! of wim file to work dir... 
		"!imagex_exe!" /Export "!install_wim!" !install_wim_index! !temp_install_wim! "!name!" >nul 2>&1
		set exitcode=!ERRORLEVEL!
		echo !exitcode!

		if not !exitcode!==0 (
			echo Error exporting !install_wim_index! of wim file to work dir
			goto end
		)
	)	
)

set readonly=
if "!dryrun!"=="True" set readonly=/ReadOnly
echo|set /p=!TIME:~0,2!:!TIME:~3,2! Mounting index !temp_install_wim_index! of wim file to mount dir !readonly!... 
"!dism_exe!" /Mount-Wim /WimFile:!temp_install_wim! /Index:!temp_install_wim_index! /MountDir:"!mount!" !readonly! /English >nul 2>&1
set exitcode=!ERRORLEVEL!
echo !exitcode!

if not !exitcode!==0 (
	echo Error mounting index !temp_install_wim_index! of wim file to mount dir
	goto end
)

"!timeout_exe!" /t 3 >nul 2>&1

:getenabledfeatures

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating enabled features... 
set which=/Online
if not "!online!"=="True" set which=/Image:"!mount!"
"!dism_exe!" !which! /Get-Features /Format:Table /English > "!dism_features!"
set exitcode=!ERRORLEVEL!

if not !exitcode!==0 (
	echo !exitcode!
	echo Error determinating enabled features
	goto end
)

set count=0
for /f %%f in ('type !dism_features! ^|!find_exe! "| Enabled"') do set /a count+=1
echo found !count! enabled features

:getinstalledpatches
if "!force_install!"=="True" goto getjobs

del /f /q "!dism_installedpackages!" >nul 2>&1
del /f /q "!systeminfo_installedpackages!" >nul 2>&1

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating installed KBs with dism... 
set which=/Online
if not "!online!"=="True" set which=/Image:"!mount!"
"!dism_exe!" !which! /Get-Packages > "!dism_installedpackages!" /English 2>&1
set exitcode=!ERRORLEVEL!

if not !exitcode!==0 (
	echo !exitcode!
	echo Error determinating installed KBs with dism
	goto end
)

set count=0
for /f %%c in ('type "!dism_installedpackages!" ^| !find_exe! /i /c "kb"') do set count=%%c
echo found !count! KBs

if "!online!"=="True" (

	echo|set /p=!TIME:~0,2!:!TIME:~3,2! Determinating installed KBs with systeminfo... 
	"!systeminfo_exe!" > "!systeminfo_installedpackages!" 2>&1
	set exitcode=!ERRORLEVEL!

	if not !exitcode!==0 (	
		echo !exitcode!
		echo Error determinating installed KBs with systeminfo
		goto end
	)
	
	set count=0
	for /f %%c in ('type "!systeminfo_installedpackages!" ^| !find_exe! /i /c "kb"') do set count=%%c
	echo found !count! KBs
	
)

"!timeout_exe!" /t 3 >nul 2>&1

:getjobs

set patches_dir=!patches_dir!\!arch!\!producttype!\!version!
set patches_dir=%patches_dir%

if not exist "!patches_dir!" (

	echo Error no matching dir found for !arch!\!producttype!\!version!
	
	echo|set /p=Creating dir !patches_dir!...
	md "!patches_dir!" >nul 2>&1
	echo !ERRORLEVEL!
	
	goto end
)

for /f "tokens=*" %%d in ('dir /b /a:d /o:n "!patches_dir!"') do (

	set subdir_name=%%d
	set subdir_ap=!patches_dir!\!subdir_name!
	
	if exist "!subdir_ap!\links.txt" (
		echo !subdir_name! >> "!jobs!"
	) else (
		if exist "!subdir_ap!\*.msu" echo !subdir_name! >> "!jobs!"
	)
	
	for /f %%f in ('type !dism_features! ^|!find_exe! "| Enable"') do (
		
		set feature_name=%%f
		set subdir_feature_ap=!subdir_ap!\!feature_name!
		if exist "!subdir_feature_ap!\links.txt" (
			echo !subdir_name!\!feature_name! >> "!jobs!"
		) else (
			if exist "!subdir_feature_ap!\*.msu" echo !subdir_name!\!feature_name! >> "!jobs!"
		)
	)

)

:dojobs

if not exist "!jobs!" (
	echo Error found no links.txt or msu files	
	goto end
)

if "!downloadpatches!"=="True" (
	ping -n 1 google.com 2>&1 | !find_exe! "TTL" >nul 2>&1
	if not !ERRORLEVEL!==0 (
		echo Erro no internet connection
		goto end
	)
)

set lastexitcode=-99999

for /f "tokens=*" %%j in (!jobs!) do (
	
	set job_name=%%j
	
	set job_ap=!patches_dir!\!job_name!
	pushd "!job_ap!"

	echo.
	echo --------
	echo !TIME:~0,2!:!TIME:~3,2! !job_name!
	echo --------

	if "!downloadpatches!"=="True" (
		
		set links_txt=links.txt
		if exist "!links_txt!" (
			
			echo.
			set download_method=Downloading only missing files
			if "!force_downloadpatches!"=="True" set download_method=Force downloading all files
			echo !TIME:~0,2!:!TIME:~3,2! !download_method!:
			echo.
			
			set count=0
			for /f %%c in ('type "!links_txt!" ^| !find_exe! /v /c ""') do set count=%%c

			if !count! lss 10 set count_zeroes=00!count!
			if !count! gtr 9 set count_zeroes=0!count!
			if !count! gtr 99 set count_zeroes=!count!
				
			set index=0
			for /f "usebackq tokens=*" %%l in ("!links_txt!") do (
	
				set /a index+=1
				if !index! lss 10 set index_zeroes=00!index!
				if !index! gtr 9 set index_zeroes=0!index!
				if !index! gtr 99 set index_zeroes=!index!
				
				set url=%%l
				set filename=%%~nxl
	
				set kb=!filename!
				set "kb=!kb:*-KB=!"
				set "kb=KB!kb:~0,7!"
		
				if "!force_downloadpatches!"=="True" del /f /q !filename! >nul 2>&1
				
				if exist "!filename!" (
					if "!dryrun!"=="False" if "!log_found!"=="True" echo !TIME:~0,2!:!TIME:~3,2! !index_zeroes!/!count_zeroes! !kb! seems to be already downloaded
				) else (
				
					set job_number=%RANDOM%
					set temp_file=%TEMP%\!filename!_!job_number!
					
					if "!dryrun!"=="True" (
						echo !TIME:~0,2!:!TIME:~3,2! !index_zeroes!/!count_zeroes! Would download !kb! 
					) else (
						echo|set /p=!TIME:~0,2!:!TIME:~3,2! !index_zeroes!/!count_zeroes! Downloading !kb!... 
						"!bitsadmin_exe!" /Transfer "bitsjob_!job_number!" "!url!" "!temp_file!" >nul 2>&1
						set exitcode=!ERRORLEVEL!
						echo !exitcode!
						
						if not !exitcode!==0 (
							echo Error downlading !kb!
							del /f /q "%TEMP%\!filename!" >nul 2>&1
							goto end
						)
						
						copy "!temp_file!" "!filename!" >nul 2>&1
						set exitcode=!ERRORLEVEL!
						
						del /f /q "%TEMP%\!filename!" >nul 2>&1

						if not !exitcode!==0 (
							echo Error copying downloaded "!kb!" to !job_name!
							goto end
						)
					)					
				)
			)		
		)		
	)

	"!timeout_exe!" /t 3 >nul 2>&1
	
	if exist *.msu (
				
		del /f /q "!patches_install!\*.msu" >nul 2>&1
		md "!patches_install!" >nul 2>&1
	
		echo.
		set copy_method=Copying only packages, which seems to be not installed
		if "!force_install!"=="True" set copy_method= Copying all packages for installation
		echo !TIME:~0,2!:!TIME:~3,2! !copy_method!:
		echo.
	
		set count=0
		for /f %%c in ('dir *.msu /b ^| !find_exe! /v /c ""') do set count=%%c
		
		if !count! lss 10 set count_zeroes=00!count!
		if !count! gtr 9 set count_zeroes=0!count!
		if !count! gtr 99 set count_zeroes=!count!
				
		set index=0
		for /f "tokens=*" %%f in ('dir *.msu /b /o:d /t:w') do (
			
			set /a index+=1
			if !index! lss 10 set index_zeroes=00!index!
			if !index! gtr 9 set index_zeroes=0!index!
			if !index! gtr 99 set index_zeroes=!index!
			
			set filename=%%f
			
			set kb=!filename!
			set "kb=!kb:*-KB=!"
			set "kb=KB!kb:~0,7!"
			
			set install=True
			if not "!force_install!"=="True" (
				type !dism_installedpackages! 2>&1 | !find_exe! "!kb!" >nul 2>&1
				if !ERRORLEVEL!==0 (
					set install=False
				) else (
					type !systeminfo_installedpackages! 2>&1 | !find_exe! "!kb!" >nul 2>&1
					if !ERRORLEVEL!==0 set install=False
				)				
			)
			if "!install!"=="False" (		
				if "!dryrun!"=="False" if "!log_found!"=="True" echo !TIME:~0,2!:!TIME:~3,2! !index_zeroes!/!count_zeroes! !kb! seems to be already installed
			) else (
				
				set /a installedmsus+=1
				
				if "!dryrun!"=="True" (
					echo !TIME:~0,2!:!TIME:~3,2! !index_zeroes!/!count_zeroes! Would copy !kb! to work dir
				) else (
					echo|set /p=!TIME:~0,2!:!TIME:~3,2! !index_zeroes!/!count_zeroes! Copying !kb! to work dir... 
					copy "!filename!" "!patches_install!\!filename!" >nul 2>&1
					set exitcode=!ERRORLEVEL!
					echo !exitcode!
					
					if not !exitcode!==0 (
						set installedmsus=0
						echo Error copying !kb! to work dir
						goto end
					)
				)
				
				if "!online!"=="True" if "!online_dism!"=="True" (
										
					if "!dryrun!"=="True" (
						echo !TIME:~0,2!:!TIME:~3,2! Would extract !kb! 
					) else (		
						
						echo|set /p=!TIME:~0,2!:!TIME:~3,2! Extracting !kb!... 
						"!expand_exe!" -f:* "!patches_install!\!filename!" "!patches_install!" >nul 2>&1
						set exitcode=!ERRORLEVEL!
						echo !exitcode!
						
						if not !exitcode!==0 (
							set installedmsus=0
							echo Error extracting !kb!
							goto end
						)
						
						del /f /q "!patches_install!\*.txt" >nul 2>&1
						del /f /q "!patches_install!\*.xml" >nul 2>&1
						del /f /q "!patches_install!\wsusscan.cab" >nul 2>&1
						del /f /q "!patches_install!\*.exe" >nul 2>&1
						del /f /q "!patches_install!\!filename!" >nul 2>&1
						
					)
				)							
			)
		)
	)	
	popd "!job_ap!"

	"!timeout_exe!" /t 3 >nul 2>&1
	
	pushd "!patches_install!"
	
	set foundfiles=False	
	if exist "*.msu" set foundfiles=True
	if exist "*.cab" set foundfiles=True

	if "!foundfiles!"=="True" (
	
		echo.
		set install_method=Install only missing packages
		if "!force_install!"=="True" set install_method=Install all packages
		echo !TIME:~0,2!:!TIME:~3,2! !install_method!:
		echo.
		
		if "!online!"=="True" (
			
			if "!online_dism!"=="True" (		
				
				if "!dryrun!"=="True" (
					echo !TIME:~0,2!:!TIME:~3,2! Would install online ^(dism^)
				) else (			
								
					echo !TIME:~0,2!:!TIME:~3,2! - Installing online ^(dism^):
					"!dism_exe!" /Online /Add-Package /PackagePath:"!patches_install!" /NoRestart /English
					set exitcode=!ERRORLEVEL!
					echo !exitcode!
					if not !exitcode!==3010 if not !exitcode!==0 (
						set installedmsus=0					
						echo Error installing online ^(dism^) 
						goto end
					)
				)
				
			) else (
								
				set count=0
				for /f %%c in ('dir "!patches_install!\*.msu" /b ^| !find_exe! /v /c ""') do set count=%%c

				if !count! lss 10 set count_zeroes=00!count!
				if !count! gtr 9 set count_zeroes=0!count!
				if !count! gtr 99 set count_zeroes=!count!
	
				set index=0
				for /f "tokens=*" %%f in ('dir "!patches_install!\*.msu" /b /o:d /t:w') do (
					
					set /a index+=1
					if !index! lss 10 set index_zeroes=00!index!
					if !index! gtr 9 set index_zeroes=0!index!
					if !index! gtr 99 set index_zeroes=!index!
		
					set filename=%%f
					
					set kb=!filename!
					set "kb=!kb:*-KB=!"
					set "kb=KB!kb:~0,7!"
		
					if "!dryrun!"=="True" (
						echo !TIME:~0,2!:!TIME:~3,2! Would install online ^(wusa^) !kb!
					) else (
					
						echo|set /p=!TIME:~0,2!:!TIME:~3,2! !index_zeroes!/!count_zeroes! Installing online ^(wusa^) !kb!... 
						call "!wusa_exe!" "!filename!" /quiet /norestart
						set exitcode=!ERRORLEVEL!
						echo !exitcode!
						
						if not !exitcode!==3010 if not !exitcode!==0 (
							if "!wusa_exitonerror!"=="True" (
								set installedmsus=0
								echo Error installing online ^(wusa^) !kb!
								goto end
							)
							
							set lastexitcode=!exitcode!
							set /a installedmsus-=1

						)
					)
					
					if not !lastexitcode!==-99999 set exitcode=!lastexitcode!
				)
			)		
		) else (
		
			if "!dryrun!"=="True" (
				echo !TIME:~0,2!:!TIME:~3,2! Would install offline ^(dism^) !kb!
			) else (
				echo !TIME:~0,2!:!TIME:~3,2! - Installing offline ^(dism^):
				"!dism_exe!" /Image:!mount! /Add-Package /Packagepath:"!patches_install!" /English
				set exitcode=!ERRORLEVEL!
				
				if not !exitcode!==0 (
					set installedmsus=0
					echo Error installing offline ^(dism^) 
					goto end
				)
			)
		)		
	)
	
	popd "!patches_install!"
	
	"!timeout_exe!" /t 3 >nul 2>&1
	
)

:unmount
if "!online!"=="True" goto end
if "!dryrun!"=="True" goto end
if !installedmsus!==0 (
	echo Error it seems, that no msu have been installed
	goto end
)
 
echo.

echo|set /p=!TIME:~0,2!:!TIME:~3,2! Commiting and unmounting mount dir... 
"!dism_exe!" /Unmount-Wim /MountDir:"!mount!" /Commit /English >nul 2>&1
set returncode=!ERRORLEVEL!
echo !returncode!

if not !returncode!==0 (
	set exitcode=!returncode!
	set installedmsus=0
	echo Error commiting and unmounting mount dir
	goto end
)

"!timeout_exe!" /t 3 >nul 2>&1

set sortabledate=00000000
for /f %%a in ('!wmic_exe! os get localdatetime ^| !find_exe! "."') do set sortabledate=%%a
set sortabledate=!sortabledate:~0,8!

set /a ran_numb=%RANDOM% %%100

set newname=!sortabledate!_!ran_numb!_!name!

if "!install_wim_copy!"=="True" (

	if not exist "!temp_install_wim!" (
		echo Error temp install wim file not found
		goto end
	)

	set description=!name! with !installedmsus! MSUs
	
	echo|set /p=!TIME:~0,2!:!TIME:~3,2! Setting name and description for the new image... 
	"!imagex_exe!" /Info "!temp_install_wim!" 1 "!newname!" "!description!" >nul 2>&1
	set returncode=!ERRORLEVEL!
	echo !returncode!
	
	if not !returncode!==0 (
		echo Error setting name and description for the new image
		goto end
	)
	
	echo|set /p=!TIME:~0,2!:!TIME:~3,2! Appending new image to original wim file... 
	"!imagex_exe!" /Export "!temp_install_wim!" 1 "!install_wim!" "!newname!" >nul 2>&1
	set returncode=!ERRORLEVEL!
	echo !returncode!
	
	if not !returncode!==0 (
		echo Error appending new image to original wim file
		goto end
	)
	
	"!timeout_exe!" /t 3 >nul 2>&1

)


if "!install_wim_createiso!"=="True" (
	
	if not exist "!osdcimg_exe!" ( 
		echo Error oscdimg.exe not found
		goto end
	)
	
	set etfsboot_com=!cdroot_dir!\boot\etfsboot.com 
	if not exist "!etfsboot_com!" ( 
		echo Error etfsboot.com file not found
		goto end
	)
	
	echo|set /p=!TIME:~0,2!:!TIME:~3,2! Creating !newname!.iso... 
	"!osdcimg_exe!" -b"!etfsboot_com!" -u2 -h -m -l"!newname:~32!" "!cdroot_dir!" "!isooutput_dir!\!newname!.iso" >nul 2>&1
	set returncode=!ERRORLEVEL!
	echo !returncode!
	
	if not !returncode!==0 (
		echo Error creating !newname!.iso
		goto end
	)

) 

"!timeout_exe!" /t 3 >nul 2>&1

:end

echo.

taskkill /f /im dism.exe >nul 2>&1
if !ERRORLEVEL!==0 echo Killed remaining dism.exe processes
REM taskkill /f /im powershell.exe >nul 2>&1
REM if !ERRORLEVEL!==0 echo Killed remaining powershell.exe processes

"!dism_exe!" /Get-MountedWimInfo | !find_exe! /i "!mount!" >nul 2>&1
if !ERRORLEVEL!==0 (
	echo|set /p=!TIME:~0,2!:!TIME:~3,2! Discarding and unmounting mount dir... 
	"!dism_exe!" /Unmount-Wim /MountDir:"!mount!" /Discard /English >nul 2>&1
	set returncode=!ERRORLEVEL!
	echo !returncode!
	
	if not !returncode!==0 (
		if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode! 
		echo Error discarding and unmounting mount dir
	)
	
	"!timeout_exe!" /t 3 >nul 2>&1
)

if exist "!work_dir!" (
	echo|set /p=!TIME:~0,2!:!TIME:~3,2! Deleting work dir ... 
	rmdir /s /q "!work_dir!" >nul 2>&1
	set returncode=!ERRORLEVEL!
	if not exist "!work_dir!" set returncode=0
	echo !returncode!
	
	if not !returncode!==0 (
		if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode! 
		echo Error deleting work dir
	)
)

if "!online!"=="True" (

	sc query wuauserv | !find_exe! "RUNNING" >nul 2>&1
	if !ERRORLEVEL!==0 (
		echo|set /p=!TIME:~0,2!:!TIME:~3,2! Stopping wuauserv... 
		net stop wuauserv >nul 2>&1
		set returncode=!ERRORLEVEL!
		echo !returncode!

		if not !returncode!==0 (
			if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode!  
			echo Error stopping wuauserv
		)
		
		"!timeout_exe!" /t 3 >nul 2>&1
	)
	
	if "!wu_reg_savedvalue!"=="?" (
		
		reg query "!wu_reg_path!" /v "!wu_reg_name!" >nul 2>&1
		if !ERRORLEVEL!==0 (
			echo|set /p=!TIME:~0,2!:!TIME:~3,2! Deleting !wu_reg_name!... 
			reg delete "!wu_reg_path!" /v !wu_reg_name! /f >nul 2>&1
			set returncode=!ERRORLEVEL!
			echo !returncode!

			if not !returncode!==0 (
				if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode!  
				echo Error deleting !wu_reg_name!
			)
		)
		
	) else (

		echo|set /p=!TIME:~0,2!:!TIME:~3,2! Restoring !wu_reg_name! to !wu_reg_savedvalue!...
		reg add "!wu_reg_path!" /v !wu_reg_name! /t REG_DWORD /d !wu_reg_savedvalue! /f >nul 2>&1
		set returncode=!ERRORLEVEL!
		echo !returncode!

		if not !returncode!==0 (
			if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode!  
			echo Error restoring !wu_reg_name! to !wu_reg_savedvalue!
		)
		
	)

	sc qc wuauserv | !find_exe! "DISABLED" >nul 2>&1
	if !ERRORLEVEL!==0 (
		echo|set /p=!TIME:~0,2!:!TIME:~3,2! Setting starttype for wuauserv to demand... 
		sc config wuauserv start= demand >nul 2>&1
		set returncode=!ERRORLEVEL!
		echo !returncode!

		if not !returncode!==0 (
			if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode!  
			echo Error setting starttype for wuauserv to demand
		)

	)	

	echo|set /p=!TIME:~0,2!:!TIME:~3,2! Starting wuauserv... 
	net start wuauserv >nul 2>&1
	set returncode=!ERRORLEVEL!
	echo !returncode!

	if not !returncode!==0 (
		if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode!  
		echo Error starting wuauserv
	)
	
)

echo.
echo !TIME:~0,2!:!TIME:~3,2! Installed !installedmsus! MSUs
echo !TIME:~0,2!:!TIME:~3,2! Exitcode - !exitcode!

if "!pause_end!"=="True" pause
"!timeout_exe!" /t 3 >nul 2>&1
exit !exitcode!
