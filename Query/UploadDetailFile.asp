<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<% 
Server.ScriptTimeout = 6800
%>
<!--#include FILE="../Common/upload_5xsoft.inc"-->
<%
	getServerIP=Request.ServerVariables("LOCAL_ADDR")
	BillSN=trim(request("SN"))

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing

	KLdirectory=""
	KLimageOrder=""
	If sys_City="基隆市" Then
		strchk="select * from BILLILLEGALIMAGE where billsn=" & BillSN
		set rsChk=conn.execute(strchk)
		If Not rsChk.eof Then
			KLdirectory=Replace(Server.MapPath(Trim(rsChk("IISImagePath"))),"\","\\") & "\\"
			If Trim(rsChk("ImageFileNameB") & " ")="" Then
				KLimageOrder="B"
			ElseIf Trim(rsChk("ImageFileNameC") & " ")="" Then
				KLimageOrder="C"
			End If 
		Else
			KLdirectory="D:\\Image\\finish\\BillBaseDetail\\"& BillSN
			KLimageOrder="A"
		End If 
		rsChk.close
		Set rsChk=Nothing 
	End If 

	' 建立新的目錄存放 每個舉發單的影像資料
	dim Newdirectory 
	If sys_City="高雄市" Then
		If getServerIP="10.133.2.161" then
			Newdirectory="\\10.133.2.163\image\BillBaseDetail\" & BillSN
		elseIf CheckServerIPForUploadImage="10.133.2.239" Then
			uploadsDirVar="\\10.133.2.176\image\BillBaseDetail\" & BillSN
		Else 
			Newdirectory="F:\Image\BillBaseDetail\" & BillSN
		End If 
	elseIf sys_City="苗栗縣" Then
		Newdirectory="F:\\Image\\BillBaseDetail\\" & BillSN
	elseIf sys_City="花蓮縣" Then
		Newdirectory="F:\\Image\\BillBaseDetail\\" & BillSN
	elseIf sys_City="基隆市" Then
		Newdirectory = KLdirectory 
	elseIf sys_City="澎湖縣" Then
		Newdirectory="E:\\Image\\Finish\\BillBaseDetail\\" & BillSN
	elseIf sys_City="金門縣" Then
		Newdirectory="F:\\Image\\Finish\\BillBaseDetail\\" & BillSN
	else
		Newdirectory="d:\\F\\Image\\BillBaseDetail\\" & BillSN
	End if

	set Newfso=Server.CreateObject("Scripting.FileSystemObject")
	if (Newfso.FolderExists(Newdirectory))=false then
		Newfso.CreateFolder(Newdirectory)
	end If

	'設定要上傳的路徑
	Dim uploadsDirVar
	If sys_City="高雄市" Then
		If getServerIP="10.133.2.161" then
			uploadsDirVar="\\10.133.2.163\image\BillBaseDetail\" & BillSN
		elseIf CheckServerIPForUploadImage="10.133.2.239" Then
			uploadsDirVar="\\10.133.2.176\image\BillBaseDetail\" & BillSN
		Else 
			uploadsDirVar = "F:\Image\BillBaseDetail\"& BillSN
		End If 
	elseIf sys_City="苗栗縣" Then
		uploadsDirVar = "F:\Image\BillBaseDetail\"& BillSN
	elseIf sys_City="花蓮縣" Then
		uploadsDirVar = "F:\\Image\\BillBaseDetail\\"& BillSN
	elseIf sys_City="基隆市" Then
		uploadsDirVar = KLdirectory 
	elseIf sys_City="澎湖縣" Then
		uploadsDirVar = "E:\\Image\\Finish\\BillBaseDetail\\" & BillSN
	elseIf sys_City="金門縣" Then
		uploadsDirVar = "F:\\Image\\Finish\\BillBaseDetail\\" & BillSN
	else
		uploadsDirVar = "d:\\F\\Image\\BillBaseDetail\\"& BillSN
	End if

'response.write uploadsDirVar

if Request.TotalBytes>1 then
dim upload,file,formName,FileType
set upload=new upload_5xSoft ''建立上傳對象

	if Right(uploadsDirVar, 1) <> "\" then uploadsDirVar = uploadsDirVar & "\"
	
	'response.write UpFilePath&FileNameTmp1

FileType=".jpg"
for each formName in upload.file ''列出所有上傳的檔案
	set file=upload.file(formName)  ''生成一個檔案對象
	If sys_City="澎湖縣" Then
		TypeFlag = 1        '檔案為允許的類型
	Else
		if Instr(FileType,GetExtendName(file.FileName)) then
			TypeFlag = 1        '檔案為允許的類型
		else
			TypeFlag = 0		'檔案為不允許的類型
			Response.write "<script>"
			Response.Write "alert('不支援您所上傳的檔案類型："&GetExtendName(file.FileName)&"');"
			response.write "self.close();"
			Response.write "</script>"
			exit for
		end if
	End If 
	FileNameTmp1="Upload"&year(now)&month(now)&day(now)&hour(now)&minute(now)&Second(now)&"_"&file.FileName
	if TypeFlag = 1 then
		'response.end
		file.SaveAs uploadsDirVar&FileNameTmp1   ''儲存檔案
		Response.write "<script>"
		Response.Write "alert('上傳檔案成功!');"
		response.write "opener.myForm.submit();"
		response.write "self.close();"
		Response.write "</script>"
		'fileStr = fileStr & "<img src='pic/addon.gif'><a href='"& UpFilePath&file.FileName&"' target='_blank'>查看上傳的檔案﹕<font color='red'>" & file.FileName &"</font> ("& file.FileSize &" kb)</a><br>"
		'FileNameStr = UpFilePath&fname
	end if
next
set upload=nothing  ''刪除此對象
end if

If sys_City="基隆市" Then
	if TypeFlag = 1 Then
		If KLimageOrder="A" Then
			strUpd="Insert into BILLILLEGALIMAGE(BILLSN,IMAGEFILENAMEA,IMAGEFILENAMEB,IMAGEFILENAMEC,IISIMAGEPATH)" &_
				" values("&BillSN&",'"&FileNameTmp1&"','','','/Imgfix/BillBaseDetail/"&BillSN&"/')"
			conn.execute strUpd
		ElseIf KLimageOrder="B" Then
			strUpd="Update BILLILLEGALIMAGE set IMAGEFILENAMEB='"&FileNameTmp1&"' where BillSn=" & BillSN
			conn.execute strUpd
		ElseIf KLimageOrder="C" Then
			strUpd="Update BILLILLEGALIMAGE set IMAGEFILENAMEC='"&FileNameTmp1&"' where BillSn=" & BillSN
			conn.execute strUpd
		End if
	End If 
End If 

conn.close
set conn=Nothing 

function GetExtendName(FileName)
dim ExtName
ExtName = LCase(FileName)
ExtName = right(ExtName,3)
ExtName = right(ExtName,3-Instr(ExtName,"."))
GetExtendName = ExtName
end function
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單影像資料上傳</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name="myForm" method="post" enctype="multipart/form-data" >
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>舉發單影像資料上傳</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>上傳格式<font color="red">(*.jpg)</font></td>
					<td>
						檔案1:<input type="file" name="file1" style="width:400" class="btn1" value="" size="50">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input type="submit" name="Submit" value="上傳">
			<input type="button" name="Submit2" value="關閉" onclick="self.close();">
			<img src="space.gif" width="20" height="5">
		</td>
	</tr>
</table>
  <br>
 <%If sys_City="澎湖縣" Then%>
  注意：<b>路徑</b> 與 <b>檔名</b>不可包含 <b>中文、空白</b>、每檔案容量限制上傳 80M<br>
 <%End if%>
</form>
</body>
</html>