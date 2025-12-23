<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<% 
''將現在的日期和時間轉為檔案名稱
function makefilename(fname)
  fname = now()
  fname = replace(fname,"-","")
  fname = replace(fname," ","") 
  fname = replace(fname,":","")
  fname = replace(fname,"PM","")
  fname = replace(fname,"AM","")
  fname = replace(fname,"上午","")
  fname = replace(fname,"下午","")
  makefilename=fname
end function
%>
<!--#include FILE="../Common/upload_5xsoft.inc"-->
<%
if Request.TotalBytes>1 then
dim upload,file,formName,FileType
set upload=new upload_5xSoft ''建立上傳對象
UpFilePath="\\10.133.2.163\image\finish\Type5\S220060233\20111221\"
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
		file.SaveAs Server.mappath(UpFilePath&fname)   ''儲存檔案
		Response.write "<script>"
		Response.Write "alert('"&file.FilePath&file.FileName&" ("&file.FileSize&") => 上傳檔案成功!');"
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
<title>簽章圖片上傳</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name="myForm" method="post" enctype="multipart/form-data" >
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>簽章圖片上傳</font></td>
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
			<input type="Hidden" name="SN" value="<%=request("MemberID")%>">
			<img src="space.gif" width="20" height="5">
		</td>
	</tr>
</table>
</form>
</body>
</html>