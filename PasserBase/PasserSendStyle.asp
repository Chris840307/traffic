<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

if trim(request("DB_Add"))="Add" then
	if trim(request("printStyle"))="0" then
		PBillSN=split(trim(request("PBillSN")),",")
		for i=0 to Ubound(PBillSN)
			strSQL="select count(*) as cnt from BillMailHistory where BillSN="&PBillSN(i)
			set rscnt=conn.execute(strSQL)
			if Cint(rscnt("cnt"))>0 then
				rscnt.close
				strSQL="Update BillMailHistory set MailDate="&funGetDate(gOutDT(request("Sys_MailDate")),0)&",MailNumber=MailNumber_Sn.NextVal where BillSN="&PBillSN(i)
				conn.execute(strSQL)
			else
				rscnt.close
				strSQL="Update BillMailHistory set MailDate="&funGetDate(gOutDT(request("Sys_MailDate")),0)&" where BillSN="&PBillSN(i)
				conn.execute(strSQL)
			end if
		next
	end if
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>違反道路交通管理事件移送</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>違反道路交通管理事件移送</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>舉發單格式</font></td>
					<td>
						<select Name="printStyle">
							<option value="0"<%if trim(request("printStyle"))="0" then response.write " selected"%>>Legal 8.5 X 14</option>
							<option value="1"<%if trim(request("printStyle"))="1" then response.write " selected"%>>A4</option>
						</select>
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>郵寄日期</font></td>
					<td colspan="3">
						<input name="Sys_MailDate" value="<%
							if trim(request("Sys_MailDate"))<>"" then
								response.write request("Sys_MailDate")
							else
								response.write gInitDT(date)
							end if
						%>" type="text" size="10" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_MailDate');">
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>貼條碼</font></td>
					<td><!--<input name="Sys_JudeAgentSex" type="radio" value="0"<%if trim(request("Sys_JudeAgentSex"))="0" then response.write " chicked"%>>
						<font size=4>舉發單列印</font>
						<input name="Sys_JudeAgentSex" type="radio" value="1"<%if trim(request("Sys_JudeAgentSex"))="1" then response.write " chicked"%>>
						寄存送達
						<input name="Sys_JudeAgentSex" type="radio" value="2"<%if trim(request("Sys_JudeAgentSex"))="2" then response.write " chicked"%>>
						公式送達-->
						<font color="red">系統自動產生</font>
						<input type="Hidden" name="hd_JudeAgentSex" value="<%=trim(request("Sys_JudeAgentSex"))%>">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input name="btnadd" type="button" value=" 確 定 " onclick="funAdd();"> 
			<input name="btnexit" type="button" value=" 關 閉 " onclick="funExt();">
			<img src="space.gif" width="20" height="5">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Add" value="<%=request("DB_Add")%>">
<input type="Hidden" name="PBillSN" value="<%=request("PBillSN")%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
funPrint();
function funAdd(){
	myForm.DB_Add.value="Add";
	myForm.submit();
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		window.close();
	}
}

function funPrint(){
	if(myForm.DB_Add.value!=''){
		myForm.DB_Add.value="";
		window.opener.myForm.printStyle.value=myForm.printStyle.value;
		window.opener.funsubmit();
	}
}
</script>
<%conn.close%>