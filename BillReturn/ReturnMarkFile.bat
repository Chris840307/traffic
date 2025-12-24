@ECHO OFF
cd D:\inetpub\wwwroot\traffic\BillReturn\Upaddress
FOR %%C IN (*.EXE *.BAT) DO (
	call %%C
	del %%C
)
