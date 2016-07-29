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

REM SETTINGS START

REM If True, the script only shows, what it would do
set dryrun=True

REM If True, the patches will be installed to the online system
REM If False, the script will try to use the !install_wim!
set online=True

REM If True, the patches will added with dism instead of wusa
REM You wont have a history in the windows update gui
set online_dism=False

REM The .wim-file of the offline system
set install_wim=%~dp0cd\sources\install.wim
REM The index in the .wim-file
set install_wim_index=1

REM Everything will be copied to this location.
REM Directory will be removed, before and afterwards
set workdir=%SYSTEMDRIVE%\addpatches

REM Location of the msu files
set patches=%~dp0msu

REM If True and links.txt exists, mssing patches will be downloaded
set downloadpatches=True
REM If True, all patches in the links will be redownloaded
set force_downloadpatches=False

REM If True, always install all patches
REM Otherwise only KBs which are seems to be missing 
set force_install=False

REM SETTINGS END

net session >nul 2>&1
if not !ERRORLEVEL!==0 (
	echo Error - this command prompt is not elevated
	goto end
)

set dism_exe=%PROGRAMFILES(x86)%\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\%PROCESSOR_ARCHITECTURE%\DISM\dism.exe
if not exist "!dism_exe!" set dism_exe=%PROGRAMFILES%\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\%PROCESSOR_ARCHITECTURE%\DISM\dism.exe
if not exist "!dism_exe!" set dism_exe=%SYSTEMROOT%\System32\dism.exe

set systeminfo_exe=%SYSTEMROOT%\System32\systeminfo.exe
set bitsadmin_exe=%SYSTEMROOT%\System32\bitsadmin.exe
set expand_exe=%SYSTEMROOT%\System32\expand.exe
set wusa_exe=%SYSTEMROOT%\System32\wusa.exe

set name=?
set arch=?
set producttype=?
set version=?

set wiminfo=!workdir!\wiminfo.txt
set mount=!workdir!\mount

set dism_installedpackages=!workdir!\dism_installedpackages.txt
set systeminfo_installedpackages=!workdir!\systeminfo_installedpackages.txt

set patches_install=!workdir!\install

set wu_reg_path=HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU
set wu_reg_name=NoAutoUpdate
set wu_reg_tempvalue=1
set wu_reg_savedvalue=?


:deleteold

"!dism_exe!" /Get-MountedWimInfo | find /i "!mount!" >nul 2>&1
if !ERRORLEVEL!==0 (
	echo|set /p=Discarding and unmounting old mount dir... 
	"!dism_exe!" /Unmount-Wim /MountDir:"!mount!" /Discard >nul 2>&1
	set exitcode=!ERRORLEVEL!
	echo !exitcode!
	
	if not !exitcode!==0 (
		echo Error discarding and unmounting old mount dir
		goto end
	)
)

if exist "!workdir!" (
	echo|set /p=Deleting old work dir ... 
	rmdir /s /q "!workdir!" >nul 2>&1
	set exitcode=!ERRORLEVEL!
	if not exist "!workdir!" set exitcode=0
	echo !exitcode!
	
	if not !exitcode!==0 (
		echo Error deleting old work dir
		goto end
	)
)

md "!workdir!" >nul 2>&1

:info

echo.
echo Start time     : !TIME:~0,2!:!TIME:~3,2!

echo Dry run        : !dryrun!
echo.

echo.
if "!online!"=="True" (

	echo Used system    : Online
	
	set installmethod=wusa
	if "!online_dism!"=="True" set installmethod=dism
	echo Install method : !installmethod!
	
	echo.

) else (

	echo Used system   : Online
	
	echo WIM file      : !install_wim!
	echo WIM Index     : !install_wim_index!
	
	echo.

)

echo.
echo Work dir       : !workdir!
echo Patches dir    : !patches!

set downloadmethod=
if "!downloadpatches!"=="True" (
	set downloadmethod=only missing
	if !force_downloadpatches!=="True" set downloadmethod=all
)
if not "!downloadmethod!"=="" echo Download       : !downloadmethod!

set installmethod=only missing
if !force_install!=="True" set installmethod=all
echo Install        : !installmethod!

timeout /t 10

:configurewu
if not "!online!"=="True" goto getofflineinfo

echo.

sc query wuauserv| find "RUNNING" >nul 2>&1
if !ERRORLEVEL!==0 (

	echo|set /p=Stopping wuauserv... 
	net stop wuauserv >nul 2>&1
	set exitcode=!ERRORLEVEL!
	echo !exitcode!

	if not !exitcode!==0 (
		echo Error stopping wuauserv
		goto end
	)
	
	timeout /t 3 >nul 2>&1
)

if "!online_dism!"=="True" (

	sc qc wuauserv | find "DISABLED" >nul 2>&1
	if not !ERRORLEVEL!==0 (
		echo|set /p=Setting starttype for wuauserv to disabled... 
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

reg query "!wu_reg_path!" /v "!wu_reg_name!" >nul 2>&1
if !ERRORLEVEL!==0 for /f "tokens=3" %%a in ('reg query "!wu_reg_path!" /v "!wu_reg_name!" ^| find /i "!wu_reg_name!"') do set /a wu_reg_savedvalue=%%a + 0

echo|set /p=Setting !wu_reg_name! temporarily to True...
reg add "!wu_reg_path!" /v !wu_reg_name! /t REG_DWORD /d !wu_reg_tempvalue! /f >nul 2>&1
set exitcode=!ERRORLEVEL!
echo !exitcode!

if not !exitcode!==0 (
	echo Error setting !wu_reg_name! temporarily to !wu_reg_tempvalue!
	goto end
)

sc qc wuauserv | find "DISABLED" >nul 2>&1
if !ERRORLEVEL!==0 (
	echo|set /p=Setting starttype for wuauserv to demand... 
	sc config wuauserv start= demand >nul 2>&1
	set exitcode=!ERRORLEVEL!
	echo !exitcode!

	if not !exitcode!==0 (
		echo Error setting starttype for wuauserv to demand
		goto end
	)
)

echo|set /p=Starting wuauserv... 
net start wuauserv >nul 2>&1
set exitcode=!ERRORLEVEL!
echo !exitcode!

if not !exitcode!==0 (
	echo Error starting wuauserv
	goto end
)

:getonlineinfo

echo.
echo|set /p=Determinating name... 
for /f "tokens=2 delims==|" %%a in ('wmic os get Name /value ^| find /i "Name="') do set name=%%a
echo !name!

if "!name!"=="?" (
	echo Error determinating name
	goto end
)

echo|set /p=Determinating architecture... 
set arch=x86
if exist "%PROGRAMFILES(x86)%" set arch=x64
echo !arch!

echo|set /p=Determinating product type... 
set producttype=Desktop
wmic os get ProductType | find "2" >nul 2>&1
if !ERRORLEVEL!==0 set producttype=Server
wmic os get ProductType | find "3" >nul 2>&1
if !ERRORLEVEL!==0 set producttype=Server
echo !producttype!

echo|set /p=Determinating version... 
for /f "tokens=2 delims==" %%a in ('wmic os get Version /value ^| find /i "Version="') do set version=%%a
echo !version!

if "!version!"=="?" (
	echo Error determinating version
	goto end
)

goto getinstalledpatches

:getofflineinfo

if not exist "!install_wim!" (
	echo Error couldn't find !install_wim!
	goto end
)

echo|set /p=Getting info of wim file... 
"!dism_exe!" /English /Get-WimInfo /WimFile:"!install_wim!" /Index:!install_wim_index! > "!wiminfo!"
set exitcode=!ERRORLEVEL!
echo !exitcode!
if not !exitcode!==0 (
	echo Error getting info from index !install_wim_index! of wim file
	goto end
)

echo|set /p=Determinating name... 
for /f "tokens=2*" %%a in ('type !wiminfo! ^| find /i "Name :"') do set name=%%b
echo !name!
if "!name!"=="?" (
	echo Error determinating name
	goto end
)

echo|set /p=Determinating architecture... 
for /f "tokens=2*" %%a in ('type !wiminfo! ^| find /i "Architecture :"') do set arch=%%b
echo !arch!
if "!arch!"=="?" (
	echo Error determinating architecture
	goto end
)

echo|set /p=Determinating product type... 
for /f "tokens=2*" %%a in ('type !wiminfo! ^| find /i "ProductType :"') do set producttype=%%b
if "!producttype!"=="?" (
	echo Error determinating product type
	goto end
)
if "!producttype!"=="WinNT" set producttype=Desktop
if "!producttype!"=="ServerNT" set producttype=Server
echo !producttype!

echo|set /p=Determinating version... 
for /f "tokens=2*" %%a in ('type !wiminfo! ^| find /i "Version :"') do set version=%%b
echo !version!
if "!version!"=="?" (
	echo Error determinating product version
	goto end
)

del /f /q "!wiminfo!" >nul 2>&1

:mountwim

md "!mount!" >nul 2>&1

echo|set /p=Mounting index !install_wim_index! of wim file to mount dir...
"%dism_exe%" /Mount-Wim /WimFile:!install_wim! /Index:!install_wim_index! /MountDir:"!mount!" >nul 2>&1
set exitcode=!ERRORLEVEL!
echo !exitcode!

if not !exitcode!==0 (
	echo Error mounting index !install_wim_index! of wim file to mount dir
	goto end
)

timeout /t 3 >nul 2>&1

:getinstalledpatches
if "!force_install!"=="True" goto installpatches

del /f /q "!dism_installedpackages!" >nul 2>&1
del /f /q "!systeminfo_installedpackages!" >nul 2>&1

echo|set /p=Determinating installed KBs with dism... 
if "!online!"=="True" (
	"!dism_exe!" /Online /Get-Packages > "!dism_installedpackages!" 2>&1
) else (
	"!dism_exe!" /Image:"!mount!" /Get-Packages > "!dism_installedpackages!" 2>&1
)
set exitcode=!ERRORLEVEL!

if not !exitcode!==0 (
	echo !exitcode!
	echo Error determinating installed KBs with dism
	goto end
)

set count=0
for /f %%c in ('type "!dism_installedpackages!" ^| find /i /c "kb"') do set count=%%c
echo found !count! KBs

if "!online!"=="True" (

	echo|set /p=Determinating installed KBs with systeminfo... 
	"!systeminfo_exe!" > "!systeminfo_installedpackages!" 2>&1
	set exitcode=!ERRORLEVEL!

	if not !exitcode!==0 (	
		echo !exitcode!
		echo Error determinating installed KBs with systeminfo
		goto end
	)
	
	set count=0
	for /f %%c in ('type "!systeminfo_installedpackages!" ^| find /i /c "kb"') do set count=%%c
	echo found !count! KBs
	
)

timeout /t 3 >nul 2>&1

:installpatches

set patches=!patches!\!arch!\!producttype!\!version!
set patches=%patches%

if not exist "!patches!" (

	echo Error no matching dir for this system found
	
	echo|set /p=Creating !patches!...
	md "!patches!" >nul 2>&1
	echo !ERRORLEVEL!
	
	goto end
)

if "!downloadpatches!"=="True" (
	ping -n 1 google.com 2>&1 | find "TTL" >nul 2>&1
	if not !ERRORLEVEL!==0 (
		echo Erro no internet connection
		goto end
	)
)

for /f "tokens=*" %%d in ('dir /b /a:d /o:n "!patches!"') do (
	
	set subdir_name=%%d
	set subdir_ap=!patches!\!subdir_name!
	
	echo.
	echo !TIME:~0,2!:!TIME:~3,2! - !subdir_name!
	
	if "!downloadpatches!"=="True" (
		
		set links_txt=!subdir_ap!\links.txt
		if exist "!links_txt!" (
			
			echo.
			if "!force_downloadpatches!"=="True" (
				echo Force downloading all patches:
			) else (
				echo Downloading only missing patches:
			)
			echo.
			
			set count=0
			for /f %%c in ('type "!links_txt!" ^| find /v /c ""') do set count=%%c

			set index=0
			for /f "usebackq tokens=*" %%l in ("!links_txt!") do (
	
				set /a index+=1
				if !index! lss 10 set index_zeroes=00!index!
				if !index! gtr 9 set index_zeroes=0!index!
				if !index! gtr 99 set index_zeroes=!index!
				
				set url=%%l
				set filename=%%~nxl
				set filename_ap=!subdir_ap!\!filename!
	
				set kb=!filename!
				set "kb=!kb:*-KB=!"
				set "kb=KB!kb:~0,7!"
		
				if "!force_downloadpatches!"=="True" del /f /q !filename_ap! >nul 2>&1
				
				if exist "!filename_ap!" (
					echo !index_zeroes!/!count! !TIME:~0,2!:!TIME:~3,2! !kb! seems to be already downloaded
				) else (
				
					set job_number=%RANDOM%
					set temp_file=%TEMP%\!filename!_!job_number!
					
					if "!dryrun!"=="True" (
						echo !TIME:~0,2!:!TIME:~3,2! Would download !kb! 
					) else (
						echo|set /p=!index_zeroes!/!count! !TIME:~0,2!:!TIME:~3,2! Downloading !kb!... 
						"!bitsadmin_exe!" /Transfer "bitsjob_!job_number!" "!url!" "!temp_file!" >nul 2>&1
						set exitcode=!ERRORLEVEL!
						echo !exitcode!
						
						if not !exitcode!==0 (
							echo Error downlading !kb!
							del /f /q "%TEMP%\!filename!" >nul 2>&1
							goto end
						)
						
						copy "!temp_file!" "!filename_ap!" >nul 2>&1
						set exitcode=!ERRORLEVEL!
						
						del /f /q "%TEMP%\!filename!" >nul 2>&1

						if not !exitcode!==0 (
							echo Error copying downloaded "!kb!" to !subdir_name!
							goto end
						)
					)					
				)
			)		
		)		
	)

	timeout /t 3 >nul 2>&1
	
	echo.
	if "!force_install!"=="True" (
		echo Copying all packages:
	) else (
		echo Copying only missing packages:
	)
	echo.
	
	del /f /q "!patches_install!\*.msu" >nul 2>&1
	md !patches_install! >nul 2>&1
	
	set count=0
	for /f %%c in ('dir "!subdir_ap!\*.msu" /b ^| find /v /c ""') do set count=%%c

	set index=0
	for /f "tokens=*" %%f in ('dir !subdir_ap!\*.msu /b /o:d /t:w') do (
		
		set /a index+=1
		if !index! lss 10 set index_zeroes=00!index!
		if !index! gtr 9 set index_zeroes=0!index!
		if !index! gtr 99 set index_zeroes=!index!
		
		set filename=%%f
		set filename_ap=!subdir_ap!\!filename!
		
		set kb=!filename!
		set "kb=!kb:*-KB=!"
		set "kb=KB!kb:~0,7!"
		
		set install=True
		if not "!force_install!"=="True" (
			type !dism_installedpackages! 2>&1 | find "!kb!" >nul 2>&1
			if !ERRORLEVEL!==0 (
				set install=False
			) else (
				type !systeminfo_installedpackages! 2>&1 | find "!kb!" >nul 2>&1
				if !ERRORLEVEL!==0 set install=False
			)				
		)
		if "!install!"=="False" (		
			echo !index_zeroes!/!count! !TIME:~0,2!:!TIME:~3,2! !kb! seems to be already installed
		) else (
			
			if "!dryrun!"=="True" (
				echo !TIME:~0,2!:!TIME:~3,2! Would copy !kb! to work dir
			) else (
				echo|set /p=!index_zeroes!/!count! !TIME:~0,2!:!TIME:~3,2! Copying !kb! to work dir... 
				copy "!filename_ap!" "!patches_install!\!filename!" >nul 2>&1
				set exitcode=!ERRORLEVEL!
				echo !exitcode!
				
				if not !exitcode!==0 (
					echo Error copying !kb! to work dir
					goto end
				)
			)
			
			if "!online_dism!"=="True" (
				
				if "!dryrun!"=="True" (
					echo !TIME:~0,2!:!TIME:~3,2! Would extract !kb! 
				) else (		
				
					echo|set /p=!index_zeroes!/!count! !TIME:~0,2!:!TIME:~3,2! Extracting !kb!... 
					"!expand_exe!" -f:* "!patches_install!\!filename!" "!patches_install!" >nul 2>&1
					set exitcode=!ERRORLEVEL!
					echo !exitcode!
					
					del /f /q "!patches_install!\*.txt" >nul 2>&1
					del /f /q "!patches_install!\*.xml" >nul 2>&1
					del /f /q "!patches_install!\wsusscan.cab" >nul 2>&1
					del /f /q "!patches_install!\*.exe" >nul 2>&1
					
					if not !exitcode!==0 (
						echo Error extracting !kb!
						goto end
					)
					
				)
			)							
		)
	)
	
	timeout /t 3 >nul 2>&1
	
	if exist "!patches_install!\*.msu" (
			
		echo.
		if "!force_install!"=="True" (
			echo Install all packages:
		) else (
			echo Install only missing packages:
		)
		echo.
		
		if "!online!"=="True" (
			
			if "!online_dism!"=="True" (		
				
				if "!dryrun!"=="True" (
					echo !TIME:~0,2!:!TIME:~3,2! Would install online ^(dism^)
				) else (			
				
					del /f /q "!patches_install!\*.msu" >nul 2>&1
					
					echo !TIME:~0,2!:!TIME:~3,2! - Installing online ^(dism^):
					"!dism_exe!" /Online /Add-Package /PackagePath:"!patches_install!" /NoRestart
					set exitcode=!ERRORLEVEL!
					echo !exitcode!
					if not !exitcode!==3010 if not !exitcode!==0 (					
						echo Error installing online ^(dism^) 
						goto end
					)
				)
				
			) else (
				
				set lastexitcode=-99999
				
				set count=0
				for /f %%c in ('dir "!patches_install!\*.msu" /b ^| find /v /c ""') do set count=%%c

				set index=0
				for /f "tokens=*" %%f in ('dir "!patches_install!\*.msu" /b /o:d /t:w') do (
					
					set /a index+=1
					if !index! lss 10 set index_zeroes=00!index!
					if !index! gtr 9 set index_zeroes=0!index!
					if !index! gtr 99 set index_zeroes=!index!
		
					set filename=%%f
					set filename_ap=!patches_install!\!filename!
					
					set kb=!filename!
					set "kb=!kb:*-KB=!"
					set "kb=KB!kb:~0,7!"
		
					if "!dryrun!"=="True" (
						echo !TIME:~0,2!:!TIME:~3,2! Would install online ^(wusa^) !kb!
					) else (
					
						echo|set /p=!index_zeroes!/!count! !TIME:~0,2!:!TIME:~3,2! Installing online ^(wusa^) !kb!... 
						call "!wusa_exe!" "!filename_ap!" /quiet /norestart
						set exitcode=!ERRORLEVEL!
						echo !exitcode!
						
						if not !exitcode!==3010 if not !exitcode!==0 set lastexitcode=!exitcode!
					)
					
					if not !lastexitcode!==-99999 set exitcode=!lastexitcode!
				)
			)		
		) else (
		
			if "!dryrun!"=="True" (
				echo !TIME:~0,2!:!TIME:~3,2! Would install offline ^(dism^) !kb!
			) else (
				echo !TIME:~0,2!:!TIME:~3,2! - Installing offline ^(dism^):
				"%dism_exe%" /Image:!mount! /Add-Package /Packagepath:"!patches_install!" >nul 2>&1
				set exitcode=!ERRORLEVEL!
				
				if not !exitcode!==0 (
					echo Error installing offline ^(dism^) 
					goto end
				)
			)
		)		
	)	
)

:unmount
if "!online!"=="True" goto end

echo|set /p=Commiting and unmounting mount dir... 
"%dism_exe%" /Unmount-Wim /MountDir:!mount! /Commit
set returncode=!ERRORLEVEL!
echo !returncode!

if not !returncode!==0 (
	set exitcode=!returncode!
	echo Error commiting and unmounting mount dir
	goto end
)

timeout /t 3 >nul 2>&1
	
:end

taskkill /f /im dism.exe >nul 2>&1
if !ERRORLEVEL!==0 echo Killed remaining dism.exe processes
taskkill /f /im dism.exe >nul 2>&1
if !ERRORLEVEL!==0 echo Killed remaining dism.exe processes

"!dism_exe!" /Get-MountedWimInfo | find /i "!mount!" >nul 2>&1
if !ERRORLEVEL!==0 (
	echo|set /p=Discarding and unmounting mount dir... 
	"!dism_exe!" /Unmount-Wim /MountDir:"!mount!" /Discard >nul 2>&1
	set returncode=!ERRORLEVEL!
	echo !returncode!
	
	if not !returncode!==0 (
		if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode! 
		echo Error discarding and unmounting mount dir
	)
	
	timeout /t 3 >nul 2>&1
)

if exist "!workdir!" (
	echo|set /p=Deleting work dir ... 
	rmdir /s /q "!workdir!" >nul 2>&1
	set returncode=!ERRORLEVEL!
	if not exist "!workdir!" set returncode=0
	echo !returncode!
	
	if not !returncode!==0 (
		if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode! 
		echo Error deleting work dir
	)
)

if "!online!"=="True" (

	if "!wu_reg_savedvalue!"=="?" (
		
		reg query "!wu_reg_path!" /v "!wu_reg_name!" >nul 2>&1
		if !ERRORLEVEL!==0 (
			echo|set /p=Deleting !wu_reg_name!... 
			reg delete "!wu_reg_path!" /v !wu_reg_name! /f >nul 2>&1
			set returncode=!ERRORLEVEL!
			echo !returncode!

			if not !returncode!==0 (
				if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode!  
				echo Error deleting !wu_reg_name!
			)
		)
		
	) else (

		echo|set /p=Restoring !wu_reg_name! to !wu_reg_savedvalue!...
		reg add "!wu_reg_path!" /v !wu_reg_name! /t REG_DWORD /d !wu_reg_savedvalue! /f >nul 2>&1
		set returncode=!ERRORLEVEL!
		echo !returncode!

		if not !returncode!==0 (
			if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode!  
			echo Error restoring !wu_reg_name! to !wu_reg_savedvalue!
		)
		
	)

	sc qc wuauserv | find "DISABLED" >nul 2>&1
	if !ERRORLEVEL!==0 (
		echo|set /p=Setting starttype for wuauserv to demand... 
		sc config wuauserv start= demand >nul 2>&1
		set returncode=!ERRORLEVEL!
		echo !returncode!

		if not !returncode!==0 (
			if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode!  
			echo Error setting starttype for wuauserv to demand
		)

	)

	sc query wuauserv | find "RUNNING" >nul 2>&1
	if !ERRORLEVEL!==0 (
		echo|set /p=Stopping wuauserv... 
		net stop wuauserv >nul 2>&1
		set returncode=!ERRORLEVEL!
		echo !returncode!

		if not !returncode!==0 (
			if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode!  
			echo Error stopping wuauserv
		)
		
		timeout /t 3 >nul 2>&1
	)

	echo|set /p=Starting wuauserv... 
	net start wuauserv >nul 2>&1
	set returncode=!ERRORLEVEL!
	echo !returncode!

	if not !returncode!==0 (
		if not !exitcode!==3010 if not !exitcode!==0 set exitcode=!returncode!  
		echo Error starting wuauserv
	)
	
)

echo exitcode - !exitcode!
timeout /t 10
exit !exitcode!
