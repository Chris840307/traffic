<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>影像覆蓋</title>
<%
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
	<table width='100%' border='0' align="center" cellpadding="0" cellspacing="0">
	
	<%
	strI="select * from BILLILLEGALIMAGE where billsn="&trim(request("SN"))
	set rsI=conn.execute(strI)
	If Not rsI.eof Then
	%>
	<%	if trim(rsI("ImageFileNameA"))="" then	%>
	<tr>
		<td>
			<input type="button" name="save" value="新增影像" onclick='window.open("UploadReFile.asp?SN=<%=trim(request("SN"))%>&SelectImg=A","UploadFile","left=0,top=0,location=0,width=700,height=465,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 80px; height:26px;">
			追加新影像,請用 新增影像 功能
			<hr>
		</td>
	</tr>
	<%	end if%>
	<%	if trim(rsI("ImageFileNameA"))<>"" and isnull(rsI("ImageFileNameB")) then	%>
	<tr>
		<td>
			<input type="button" name="save" value="新增影像" onclick='window.open("UploadReFile.asp?SN=<%=trim(request("SN"))%>&SelectImg=B","UploadFile","left=0,top=0,location=0,width=700,height=465,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 80px; height:26px;">
			追加新影像,請用 新增影像 功能
			<hr>
		</td>
	</tr>
	<%	end if%>
	<%	if trim(rsI("ImageFileNameA"))<>"" and trim(rsI("ImageFileNameB"))<>"" and isnull(rsI("ImageFileNameC")) then	%>
	<tr>
		<td>
			<input type="button" name="save" value="新增影像" onclick='window.open("UploadReFile.asp?SN=<%=trim(request("SN"))%>&SelectImg=C","UploadFile","left=0,top=0,location=0,width=700,height=465,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 80px; height:26px;">
			追加新影像,請用 新增影像 功能
			<hr>
		</td>
	</tr>
	<%	end if%>
	<tr>
		<td>
	<%	if trim(rsI("ImageFileNameA"))<>"" then	%>
		<input type="button" name="save" value="選擇檔案 , 覆蓋下圖" onclick='window.open("UploadReFile.asp?SN=<%=trim(request("SN"))%>&SelectImg=A","UploadFile","left=0,top=0,location=0,width=700,height=465,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 160px; height:26px;">
		<br>
		<img src="<%=trim(rsI("IISImagePath")) & trim(rsI("ImageFileNameA"))%>" width="420"><br>
	<%	end if%>
		</td>
		<td>
	<%	if trim(rsI("ImageFileNameB"))<>"" then	%>
		<input type="button" name="save" value="選擇檔案 , 覆蓋下圖" onclick='window.open("UploadReFile.asp?SN=<%=trim(request("SN"))%>&SelectImg=B","UploadFile","left=0,top=0,location=0,width=700,height=465,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 160px; height:26px;">
		<br>
		<img src="<%=trim(rsI("IISImagePath")) & trim(rsI("ImageFileNameB"))%>" width="420"><br>
	<%	end if%>
		</td>
	</tr>
	<tr>
		<td>
	<%	if trim(rsI("ImageFileNameC"))<>"" then	%>
		<input type="button" name="save" value="選擇檔案 , 覆蓋下圖" onclick='window.open("UploadReFile.asp?SN=<%=trim(request("SN"))%>&SelectImg=C","UploadFile","left=0,top=0,location=0,width=700,height=465,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 160px; height:26px;">
		<br>
		<img src="<%=trim(rsI("IISImagePath")) & trim(rsI("ImageFileNameC"))%>" width="400"><br>
	<%	end if%>
		</td>
	</tr>
	<%
	else
		response.write "非影像建檔案件，查無違規影像!"
	end if
	rsI.close
	set rsI=nothing
	%>
	</table>
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function Inert_Data(SCode,SStreet){
	opener.myForm.MemberStation.value=SCode;
	opener.Layer5.innerHTML=SStreet;
	opener.TDStationErrorLog=0;
	window.close();
}
</script>
</html>
