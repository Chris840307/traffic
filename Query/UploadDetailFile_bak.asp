<% 
Response.Expires = -1
Server.ScriptTimeout = 60000

%>
<!-- #include file="freeASPUploadforDetail.asp" -->
<!--#include virtual="/traffic/Common/db.ini"-->
<%
' ****************************************************
' Change the value of the variable below to the pathname
' of a directory with write permissions, for example "C:\Inetpub\wwwroot"
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing
	conn.close
	Set conn=Nothing 
	' 建立新的目錄存放 每個舉發單的影像資料
		dim Newdirectory 
		If sys_City="高雄市" Then
			Newdirectory="F:\\Image\\BillBaseDetail\\" &request("SN")
		else
			Newdirectory="d:\\F\\Image\\BillBaseDetail\\" &request("SN")
		End if
        
     	set Newfso=Server.CreateObject("Scripting.FileSystemObject")
	    if (Newfso.FolderExists(Newdirectory))=false then
       		Newfso.CreateFolder(Newdirectory)
     	end if
	'設定要上傳的路徑
	Dim uploadsDirVar
	If sys_City="高雄市" Then
		uploadsDirVar = "F:\\Image\\BillBaseDetail\\"&request("SN") 
	else
		uploadsDirVar = "d:\\F\\Image\\BillBaseDetail\\"&request("SN") 
	End if
  

        dim fso
        dim directory 
	If sys_City="高雄市" Then
		directory="F:\\Image\\BillBaseDetail\\" &request("SN")
	else
		directory="d:\\F\\Image\\BillBaseDetail\\" &request("SN")
	End if
        
        
     	set fso=Server.CreateObject("Scripting.FileSystemObject")
	    if (fso.FolderExists(directory))=false then
       		fso.CreateFolder(directory)
     	end if
    	
        set fso=nothing
  
  
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
<title>舉發單影像資料上傳</title>
<body>
<form name="myForm" method="Post" enctype="multipart/form-data">
				
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			
			<tr><td bgcolor="#FFCC33"><font size="4"><strong></strong></font>舉發單影像資料上傳</td></tr>
    		</table>
    		  <p><p><b>請選擇檔案</b> &nbsp;&nbsp;<input type="file"  name="attach1" size="30" style="font-size: 12pt; height:30px;"> <br>


  <p></p>
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr><td bgcolor="#FFCC33" align="center">
						
							<input type=submit name=submit value="上傳"  style="font-size: 12pt; width: 60px; height:30px;">  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
							<input type="button" value="清除" name="btnClear" onclick="location='UploadFile.asp'"  style="font-size: 12pt; width: 60px; height:30px;" >							
							
			</td></tr>
		</table>
 <br><br>
  注意：<b>路徑</b> 與 <b>檔名</b>不可包含 <b>中文、空白</b>、每檔案容量限制上傳 1.2M<br>
  <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*可以上傳<b>多筆</b>影像資料，檔案格式為 JPG , GIF 。
	<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*檔案名稱請以<b>日期+時間，避免重覆</b>。<br> 
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;如:20090715081211.jpg (到秒)。  
		
	<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*後續可於<b>舉發單詳細</b>中查詢到影像資料<br><br>
	<font size="5"> </font>
	<img src="example.jpg" >
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