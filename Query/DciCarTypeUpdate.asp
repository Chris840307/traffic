<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<title>車種修改</title>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)
	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

if trim(request("kinds"))="DB_Update" Then
	BillNotmp=Trim(request("BillNo1"))&Trim(request("BillNo2"))
	CarNotmp=Trim(request("CarNo"))
	CarTypetmp=Trim(request("CarType"))
	ChkFlag="0"
	strChk="select * from Billbasedcireturn where billno='"&BillNotmp&"' and carno='"&CarNotmp&"' and exchangetypeid='W'"
	Set rsChk=conn.execute(strChk)
	If rsChk.eof Then
		ChkFlag="1"
	End If
	rsChk.close
	Set rsChk=nothing
		
	If ChkFlag="1" Then
	%>
<script language="JavaScript">
	alert("查無此筆資料!!!!");
</script>
<%
	Else
		strUpd="Update Billbasedcireturn set DciReturnCarType='"&Trim(CarTypetmp)&"' where billno='"&BillNotmp&"' and carno='"&CarNotmp&"' and exchangetypeid='W'"
		conn.execute strUpd

		strI="insert into Log values(log_sn.nextval+3000,2222,"&Session("User_ID")&",'"&Session("Ch_Name")&"','"&userip&"',sysdate,'單號:"&BillNotmp&",車種:"&Trim(CarTypetmp)&"')"
					Conn.execute strI
%>
<script language="JavaScript">
	alert("修改完成");
</script>
<%	End If 
end if
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4"><strong>車種修改</strong></td>
			</tr>
			<tr >
				<td width="15%" align="center" height="30" bgcolor="#EBFBE3">單號</td>
				<td width="85%" >
					<input type="text" value="" maxlength="9" name="BillNo2" onkeyup="this.value=this.value.toUpperCase()">
				</td>
			</tr>
			<tr >
				<td width="15%" align="center" height="30" bgcolor="#EBFBE3">車號</td>
				<td width="85%" >
					<input type="text" value="" maxlength="8" name="CarNo" onkeyup="this.value=this.value.toUpperCase()">
				</td>
			</tr>
			<tr >
				<td width="15%" align="center" height="30" bgcolor="#EBFBE3">車種</td>
				<td width="85%" >
					<select name="CarType">
<%
	strT="select * from DciCode where TypeID=5 order by ID"
	Set rsT=conn.execute(strT)
	If Not rsT.Bof Then rsT.MoveFirst 
	While Not rsT.Eof
%>
						<option value="<%=Trim(rsT("ID"))%>"><%=Trim(rsT("Content"))%></option>
<%	rsT.MoveNext
	Wend
	rsT.close
	set rsT=nothing
%>
					</select>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="close" value=" 確 定 " onclick="BillDel();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
'						if CheckPermission(237,4)=false then
'							response.write "disabled"
'						end if
						%>>

				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="DBillSN" value="<%=trim(request("DBillSN"))%>">
				</td>
			</tr>
		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function BillDel(){
		myForm.kinds.value="DB_Update";
		myForm.submit();
}
</script>
</html>
