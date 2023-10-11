rem This is the file of registering COM control of GrandDog.
rem If you want to unregister the COM control, using "regsvr32 /u RC_GrandDog.dll"

%~d0
cd %~dp0

if "%SystemRoot%"=="" goto Win98ME

:Win2KXP
%SystemRoot%\System32\regsvr32 RC_GrandDog.dll
goto End

:Win98ME
%WinDir%\System\regsvr32 RC_GrandDog.dll

:End