<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<%
Server.ScriptTimeout=6000
if trim(request("DB_Add"))="ADD" then
	if trim(request("OpenGovDate"))<>"" then
		sUpdSQL="update BillMailHistory set OpenGovDate=" & funGetDate(gOutDt(request("OpenGovDate")),0)  & _
			" where billSn in (select Billsn from Dcilog where batchnumber='"&trim(request("sys_Batchnumber"))&"')"
		conn.execute sUpdSQL
	end if
	strSQL="update Dcilog set FileName='',SeqNo='',dcireturnstatusid='' where batchnumber='"&trim(request("sys_Batchnumber"))&"'"
	conn.execute(strSQL)
response.Write "處理完成，請將本網頁關閉!!"
end if 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>資料修改</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
批號 <input type="text" value="" name="sys_Batchnumber" size="12">
<br>
公告日<input class="btn1" type='text' size='7' id='OpenGovDate' name='OpenGovDate'>
<br>
<input type="button" name="save" value="處理" onclick="funAdd();">
<input type="Hidden" name="DB_Add" value="">
</form>
</body>
<script language="javascript">
function funAdd(){
	var err=0;
	if(myForm.sys_Batchnumber.value==''){
		err=1;
		alert("批號不可空白");
	}else{
		myForm.DB_Add.value='ADD';
		myForm.submit();
	}
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		self.close();
	}
}

</script>
</html>