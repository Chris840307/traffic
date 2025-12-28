<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>國稅局移送紀錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%

if trim(request("DB_State"))="Add" then

	strSQL="insert into PASSERETAX(SN,BILLSN,BILLNO,OPENGOVNUMBER,ACCEPTDATE,RECORDDATE,RECORDMEMBERID) values((select nvl(max(SN),0)+1 from PASSERETAX),"&request("BillSN")&",(select billno from passerbase where sn='"&request("BillSN")&"' and recordstateid=0),'"&trim(request("Sys_OpenGovNumber"))&"',"&funGetDate(gOutDT(request("Sys_AcceptDate")),0)&",sysdate,"&Session("User_ID")&")" 

	conn.execute(strSQL)
end If 

if request("DB_State")="Del" then
	strSQL="Delete from PASSERETAX where SN="&request("SN")
	conn.execute strSQL
end If 

strSQL="Select * from PasserEtax where BILLSN="&request("BillSN")
set rsload=conn.execute(strSQL)
%>
<BODY>
<form name="myForm" method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">國稅局收文紀錄</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
						<table width="100%" border="0">
							<tr>
								<td>收文日期</td>
								<td nowrap>
									<input name="Sys_AcceptDate" class="btn1" type="text" value="" size="10" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
									<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_AcceptDate');">
								</td>
								<td nowrap>
									收文文號
								</td>
								<td>
									<input name="Sys_OpenGovNumber" class="btn1" type="text" value="" size="10" maxlength="50">
								</td>
								
								<td>
									<input type="button" name="btnAdd" value="新增" onclick="funAdd();">
									<input name="btnexit" type="button" value=" 關 閉 " onclick="funExit();">
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" class="style3">收文紀錄列表</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th>收文日期</th>
					<th>收文文號</th>
					<th>操作</th>
				</tr><%
				while Not rsload.eof
					response.write "<tr align='center' bgcolor='#FFFFFF'"
					lightbarstyle 0
					response.write ">"
					
'					response.write "<td>"
'					If not ifnull(rsload("Imagefilename")) Then						
'						Response.Write "<a href=""./Picture/"&trim(rsload("Imagefilename"))&""" target=""_blank"">"
'					end if
'
'					Response.Write gInitDT(rsload("ArrivedDate"))
'
'					If not ifnull(rsload("Imagefilename")) Then	Response.Write "</a>"
'
'					Response.Write "</td>"

					response.write "<td>"&gInitDT(rsload("ACCEPTDATE"))&"</td>"
					response.write "<td>"&rsload("OpenGovNumber")&"</td>"

					response.write "<td>"
					response.write "<input type=""button"" name=""Del"" value=""刪除"" onclick=""funDel('"&rsload("SN")&"');"">"
					response.write "</td>"

					response.write "</tr>"
					rsload.movenext
				wend%>
			</table>
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_State" value="">
<input type="Hidden" name="BillSN" value="<%=request("BillSN")%>">
<input type="Hidden" name="SN" value="">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">

function funAdd(){
	var err=0;

	if(myForm.Sys_AcceptDate.value==""){
		err=1;
		alert("收文日必須輸入!!");
	}else if(myForm.Sys_AcceptDate.value!=""){
		if(!dateCheck(myForm.Sys_AcceptDate.value)){
			err=1;
			alert("收文日輸入不正確!!");
		}
	}
	if(err==0){
		if(myForm.Sys_OpenGovNumber.value==""){
			err=1;
			alert("收文文號必須輸入!!");
		}
	}

	if(err==0){
		myForm.DB_State.value='Add';
		myForm.submit();
	}
}

function funDel(SN){
	if(confirm('確定刪除此筆紀錄嗎？')){
		myForm.SN.value=SN;
		myForm.DB_State.value='Del';
		myForm.submit();
	}
}

function funMap(SN){
	UrlStr="SendStyle.asp?SN="+SN;
	newWin(UrlStr,"winMap",700,150,50,10,"yes","yes","yes","no");
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}

function funExit(){
	opener.myForm.submit(); 
	self.close();
}
</script>
<%
conn.close
set conn=nothing
%>