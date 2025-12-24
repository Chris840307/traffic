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
<title>固定桿編號查詢</title>
<%
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post" onsubmit="return select_street();">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4">固定桿編號查詢</td>
			</tr>
			<tr>
				<td colspan="4">&nbsp;&nbsp;&nbsp;
				路段名稱：<input type="text" name="StreetName" value="<%=trim(request("StreetName"))%>">
				<br>
				固定桿編號：<input type="text" name="FixNum" value="<%=trim(request("FixNum"))%>">
				<input type="button" value="查詢" onclick="select_street();">
				<input type="button" name="close" value="關閉視窗" onclick="window.close();">
				<input type="hidden" value="" name="kinds">
				</td>
			</tr>
			<tr bgcolor="#FFCC33">
				<td colspan="4">固定桿編號列表</td>
			</tr>
			<tr bgcolor="#EBFBE3">
				<td width="15%" align="center">固定桿編號</td>
				<td width="45%" align="center">路段</td>
				<td width="15%" align="center">路段代碼</td>
				<td width="25%" align="center">型式</td>
			</tr>
<%
if trim(request("kinds"))="DB_select" then
	'取得縣市名稱
	CityName=""
	strIllAddr="select * from ApConfigure where ID=31"
	set rsIA=conn.execute(strIllAddr)
	if not rsIA.eof then
		CityName=trim(rsIA("Value"))
	end if
	rsIA.close
	set rsIA=nothing
	strSql=""
	if trim(request("StreetName"))<>"" then
		strSql="where Address Like '%"&trim(request("StreetName"))&"%'"
	end if
	if trim(request("FixNum"))<>"" then
		if strSql="" then
			strSql="where EquipMentID Like '%"&trim(request("FixNum"))&"%'"
		else
			strSql=strSql&" and EquipMentID Like '%"&trim(request("FixNum"))&"%'"
		end if
	end if
	strProject="select EquipMentID,TypeID,Address,StreetID from FixEquip "&strSql&" order by EquipMentID"
	set rsProject=conn.execute(strProject)
	If Not rsProject.Bof Then rsProject.MoveFirst 
	While Not rsProject.Eof
%>
			<tr title="請點選.." onclick="Inert_Data('<%=trim(rsProject("EquipMentID"))%>','<%=CityName&trim(rsProject("Address"))%>','<%=trim(rsProject("StreetID"))%>');" <%lightbarstyle 1 %>>
				<td bgcolor="#FFFFCC" align="center"><%=trim(rsProject("EquipMentID"))%></td>
				<td><%=trim(rsProject("Address"))%></td>
				<td><%=trim(rsProject("StreetID"))%></td>
				<td><%
				strType="select Content from Code where TypeID=18 and ID='"&trim(rsProject("TypeID"))&"'"
				set rsType=conn.execute(strType)
				if not rsType.eof then
					response.write trim(rsType("Content"))
				end if
				rsType.close
				set rsType=nothing
				%></td>
			</tr>
<%	rsProject.MoveNext
	Wend
	rsProject.close
	set rsProject=nothing
end if
%>

		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function select_street(){
	myForm.kinds.value="DB_select";
	myForm.submit();
}
function Inert_Data(FID,SStreet,SCode){
	opener.myForm.FixID.value=FID;
	//if (opener.myForm.IllegalAddress.value==""){
		opener.myForm.IllegalAddressID.value=SCode;
		opener.myForm.IllegalAddress.value=SStreet;
	//}
	//opener.myForm.BillMemID2.value=FID;
	window.close();
}
</script>
</html>
