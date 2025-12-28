<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
	Server.ScriptTimeout=600000
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close

	ftpIp=Request.ServerVariables("LOCAL_ADDR")

	scannerDir=Session("Credit_ID")

	fp=server.mappath(".\ScannerImport")&"\"&scannerDir

	set fso=Server.CreateObject("Scripting.FileSystemObject")

	if fso.FolderExists(fp) = false then
		fso.CreateFolder(fp) 
	end if

	'ServerIp=Request.ServerVariables ("LOCAL_ADDR") 
	'
	'ftpUrl="ftp://passer:pbuser@"&ServerIp&"/PasserFtp/"&scannerDir
	'
	'set fso=nothing


%>

    var executableFullPath = 'explorer.exe /n,ftp://scanner:scanner903f@<%=ftpIp%>';
    var strFile="/<%=scannerDir%>";
    try
    {
	  var shellActiveXObject = new ActiveXObject("WScript.Shell");

	  if ( !shellActiveXObject )
	  {
		alert('Could not get reference to WScript.Shell');
	  }

	  shellActiveXObject.Run(executableFullPath + strFile, 1, false);
	  shellActiveXObject = null;
    }
    catch (errorObject)
    {
	  alert('Error:\n' + errorObject.message);
    }            