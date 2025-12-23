<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<% 

%>
<!--#include FILE="../Common/upload_5xsoft.inc"-->
<%
	BillSN=trim(request("SN"))
	Name2=trim(request("SelectImg"))

if Request.TotalBytes>1 then
dim upload,file,formName,FileType
set upload=new upload_5xSoft ''建立上傳對象

	strL="select * from apconfigure where ID=36"
	set rsL=conn.execute(strL)
	if not rsL.eof then
		ImgLocation=trim(rsL("Value"))
	end if
	rsL.close
	set rsL=nothing
	
	'設定要上傳的路徑
	strFile="select a.DirectoryName from ProsecutionImage a,ProsecutionImageDetail b " &_
		" where a.filename=b.filename and a.operatora=b.operator and b.billsn="&BillSN
	'response.write strFile
	'response.end
	set rsFile=conn.execute(strFile)
	if not rsFile.eof then
		UpFilePath = Replace(replace(ImgLocation & rsFile("DirectoryName") ,"\\","\"),"F:\Image\finish\","\\10.133.2.163\image\finish\")
	end if
	rsFile.close
	set rsFile=Nothing
	if Right(UpFilePath, 1) <> "\" then UpFilePath = UpFilePath & "\"
	if BillSN<>"" then
		FileNameTmp1="Cov"&year(now)&month(now)&day(now)&hour(now)&minute(now)&Second(now)&".JPG"
		if Name2="A" then
			strA=" ImageFileNameA='"&FileNameTmp1&"'"
		elseif Name2="B" then
			strA=" ImageFileNameB='"&FileNameTmp1&"'"
		elseif Name2="C" then
			strA=" ImageFileNameC='"&FileNameTmp1&"'"
		end if
		strUpd1="Update BILLILLEGALIMAGE set "&strA&" where BillSN="&BillSN
		conn.execute strUpd1
		strUpd2="Update ProsecutionImage set "&strA&" where FileName=(select FileName from ProsecutionImageDetail where BillSN="&BillSN&") and OperatorA=(select Operator from ProsecutionImageDetail where BillSN="&BillSN&")"
		conn.execute strUpd2
	end If
	'response.write UpFilePath&FileNameTmp1

FileType=".jpg"
for each formName in upload.file ''列出所有上傳的檔案
	set file=upload.file(formName)  ''生成一個檔案對象
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
	if TypeFlag = 1 then
		fname = "Member_"&trim(Session("Map_MemberID"))&".jpg"
		'response.end
		file.SaveAs UpFilePath&FileNameTmp1   ''儲存檔案
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
  注意：<b>路徑</b> 與 <b>檔名</b>不可包含 <b>中文、空白</b>、每檔案容量限制上傳 1.2M<br>
</form>
</body>
</html>