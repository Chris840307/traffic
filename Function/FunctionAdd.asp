<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>權限設定系統-新增</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
if trim(request("kinds"))="UpdateFunc" then
	'先刪除舊的
	strDel="delete from FunctionData where GroupID="&trim(request("GroupID"))
	conn.execute strDel 
	
	strDelDetail="delete from FunctionDataDetail where GroupID="&trim(request("GroupID"))
	conn.execute strDelDetail

	SysIDTemp=split(request("SysID"),",")

	for Fsn=0 to ubound(SysIDTemp)
		if trim(request("FuncUse"&Fsn))="1" then
			strIns="Insert into FunctionData(GroupID,SystemID,Function)" &_
				" values("&trim(request("GroupID"))&","&trim(SysIDTemp(Fsn))&",1)"
			conn.execute strIns

			strInsDetail="Insert into FunctionDataDetail(SN,GroupID,SystemID,InsertFlag" &_
									",DeleteFlag,UpdateFlag,SelectFlag)" &_
									" values(FUNCTIONDATADETAIL_SN.nextval,"&trim(request("GroupID"))&","&trim(SysIDTemp(Fsn))&","&trim("1")&","&trim("1")&","&trim("1")&","&trim("1")&")"
			conn.execute strInsDetail
		end if
	next
'FUNCTIONDATADETAIL_SN
end if

%>
<SCRIPT LANGUAGE=javascript>
<!--
-->
</Script>

<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {font-size: 14px}
.style3 {font-size: 15px}
-->
</style></head>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<!-- #include file="../Common/checkFunc.inc"-->
<body>
<FORM NAME="myForm" METHOD="POST">  	
<table width="100%" height="100%" bgcolor="dddddd" border="0" cellpadding="1">
	<tr>
		<td colspan="6" height="27" bgcolor="#FFCC33">
			<span class="pagetitle style3">權限設定系統-新增</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td height="33" bgcolor="#FFFFCC">
			<div align="right" class="style3">群組</div>
		</td>
		<td colspan="5">
			<select name="GroupID" onchange="selectFunc();">
			<option value="">選擇群組...</option>
		<%
			sqlGroup= "Select ID,Content from Code where TypeID=10 order by ShowOrder"
			set RsGroup=Server.CreateObject("ADODB.RecordSet")

			RsGroup.open sqlGroup,Conn,3,3
			While Not RsGroup.Eof			
		%>
		 <option value="<%=trim(RsGroup("ID"))%>" <%if trim(RsGroup("ID"))=trim(request("GroupID")) then response.write "selected"%>><%=RsGroup("Content")%></option>
		<%
				RsGroup.MoveNext
			Wend
			RsGroup.close
			set RsGroup=nothing
		%>          
			</select>
			<input type="button" value="確定" onclick="selectFunc();">
        </td>
	</tr>
	<tr bgcolor="#FFCC55">
		<td height="27" width="30%" align="center"><strong>系統</strong></td>
		<td width="15%" align="center"><strong>使用權限</strong></td>
		<td width="14%" align="center"><strong>新增</strong></td>
		<td width="14%" align="center"><strong>修改</strong></td>
		<td width="14%" align="center"><strong>刪除</strong></td>
		<td width="14%" align="center"><strong>查詢</strong></td>
	</tr>
<%	SN=0
	strFunc="Select ID,Content from Code where TypeID=11 order by Content"
	set rsFunc=conn.execute(strFunc)
	If Not rsFunc.Bof Then rsFunc.MoveFirst 
	While Not rsFunc.Eof
%>
<%
	UseTemp=0
	SelTemp=0
	InsTemp=0
	UpdTemp=0
	DelTemp=0
if trim(request("kinds"))<>"" then
	strSys="select * from FunctionDataDetail where SystemID="&trim(rsFunc("ID"))&" and GroupID="&trim(request("GroupID"))
	set rsSys=conn.execute(strSys)
	if not rsSys.eof then
		UseTemp=1
		SelTemp=trim(rsSys("SelectFlag"))
		InsTemp=trim(rsSys("InsertFlag"))
		UpdTemp=trim(rsSys("UpdateFlag"))
		DelTemp=trim(rsSys("DeleteFlag"))
	end if
	rsSys.close
	set rsSys=nothing
end if
%>
	<tr height="32" bgcolor="#FFFFFF">
		<td bgcolor="#FFFFCC" align="left">
			<%=rsFunc("Content")%>
			<input type="hidden" name="SysID" value="<%=rsFunc("ID")%>">
		</td>
		<td align="center">
			<input type="radio" name="FuncUse<%=SN%>" value="1" <%if trim(UseTemp)=1 then response.write "checked"%> onclick="FuncChkYes(<%=SN%>);">可
			&nbsp;&nbsp;<input type="radio" name="FuncUse<%=SN%>" value="0" <%if trim(UseTemp)=0 then response.write "checked"%> onclick="FuncChkNo(<%=SN%>);">否
		</td>
		<td align="center">
			<input type="radio" name="FuncIns<%=SN%>" value="1" <%if trim(InsTemp)=1 then response.write "checked"%> disabled>可
			<input type="radio" name="FuncIns<%=SN%>" value="0" <%if trim(InsTemp)=0 then response.write "checked"%> disabled>否
		</td>
		<td align="center">
			<input type="radio" name="FuncUpd<%=SN%>" value="1" <%if trim(UpdTemp)=1 then response.write "checked"%> disabled>可
			<input type="radio" name="FuncUpd<%=SN%>" value="0" <%if trim(UpdTemp)=0 then response.write "checked"%> disabled>否
		</td>
		<td align="center">
			<input type="radio" name="FuncDel<%=SN%>" value="1" <%if trim(DelTemp)=1 then response.write "checked"%> disabled>可
			<input type="radio" name="FuncDel<%=SN%>" value="0" <%if trim(DelTemp)=0 then response.write "checked"%> disabled>否
		</td>
		<td align="center">
			<input type="radio" name="FuncSel<%=SN%>" value="1" <%if trim(SelTemp)=1 then response.write "checked"%> disabled>可
			<input type="radio" name="FuncSel<%=SN%>" value="0" <%if trim(SelTemp)=0 then response.write "checked"%> disabled>否
		</td>
	</tr>
<%		SN=SN+1
	rsFunc.MoveNext
	Wend
	rsFunc.close
	set rsFunc=nothing
%>
	<tr bgcolor="#FFFFFF" height="33">
        <td bgcolor="#FFFFCC"><div align="right" class="style3 style3">修改人員</div></td>
        <td colspan="5"><%=Session("Ch_Name")%></td>
	</tr>
	<tr>
		<td colspan="6" height="35" bgcolor="#FFDD77"><p align="center" class="style1">
			<input type="button" name="Submit423" value="確 定" onclick="UpdateFunc();">
			<span class="style3"><img src="space.gif" width="9" height="8"></span>  
		    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉"></p>	
			<input type="hidden" name="kinds" value="">
		</td>
	</tr>

</table>
</FORM>
</body>
<SCRIPT LANGUAGE=javascript>
<!--
function selectFunc(){
	if (myForm.GroupID.value==""){
		alert("請先選擇群組！");
	}else{
		myForm.kinds.value="selectFunc";
		myForm.submit();
	}
}
function UpdateFunc(){
	if (myForm.GroupID.value==""){
		alert("請先選擇群組！");
	}else{
		myForm.kinds.value="UpdateFunc";
		myForm.submit();
	}
}
function FuncChkYes(Sn){
	eval("myForm.FuncIns"+Sn+"[0]").checked=true;
	eval("myForm.FuncUpd"+Sn+"[0]").checked=true;
	eval("myForm.FuncDel"+Sn+"[0]").checked=true;
	eval("myForm.FuncSel"+Sn+"[0]").checked=true;	
}
function FuncChkNo(Sn){
	eval("myForm.FuncIns"+Sn+"[1]").checked=true;
	eval("myForm.FuncUpd"+Sn+"[1]").checked=true;
	eval("myForm.FuncDel"+Sn+"[1]").checked=true;
	eval("myForm.FuncSel"+Sn+"[1]").checked=true;	}
-->
</Script>
</html>
