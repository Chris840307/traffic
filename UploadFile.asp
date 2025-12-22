<% 
Response.Expires = -1
Server.ScriptTimeout = 60000

%>
<!-- #include file="freeaspupload.asp" -->
<%
' ****************************************************
' Change the value of the variable below to the pathname
' of a directory with write permissions, for example "C:\Inetpub\wwwroot"

  Dim uploadsDirVar
  uploadsDirVar = "d:\Inetpub\wwwroot"&request("type")

        dim fso
        dim directory 
        directory="D:\\Inetpub\\wwwroot"&Replace(request("type"),"\","\\")
response.write uploadsDirVar&"<br>"
response.write directory
  
' ****************************************************
' Note: this file uploadTester.asp is just an example to demonstrate
' the capabilities of the freeASPUpload.asp class. There are no plans
' to add any new features to uploadTester.asp itself. Feel free to add
' your own code. If you are building a content management system, you
' may also want to consider this script: http://www.webfilebrowser.com/

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();

</script>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>現場照片上傳</title>
<body>
<form name="myForm" method="Post" enctype="multipart/form-data">
				
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr><td bgcolor="#FFCC33"><font size="4"><strong></strong></font>現場照片上傳</td></tr>
    		</table>
    		  <p><p>檔案 &nbsp;&nbsp;<input type="file"  name="attach1" size="50"> <br>

  注意：檔名不可包含中文、每檔案容量限制上傳 1.2M<br>

  <p></p>
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr><td bgcolor="#FFCC33" align="center">
						
							<input type=submit name=submit value="上傳">  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
							<input type="button" value="清除" name="btnClear" onclick="location='UploadFile.asp'">							
							
			</td></tr>
		</table>

		　</form>
</html>


							
<%		
'上傳檔案到伺服器
function SaveFiles
    Dim Upload, fileName, fileSize, ks, i, fileKey

    Set Upload = New FreeASPUpload
    Upload.Save(uploadsDirVar)
end function


    SaveFiles()
%>