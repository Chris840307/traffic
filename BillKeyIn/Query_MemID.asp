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
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>身分證字號查詢</title>
<%
MemOrder=trim(request("MemOrder"))
MemType=trim(request("MemType"))

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	UserUnitTypeID=trim(Session("Unit_ID"))
	'使用者所屬上曾單位
	strUT="select UnitTypeID from UnitInfo where UnitID!='0000' and UnitID='"&trim(Session("Unit_ID"))&"'"
        'strUT="select UnitTypeID from UnitInfo"
	set rsUT=conn.execute(strUT)
	if not rsUT.eof then
		UserUnitTypeID=trim(rsUT("UnitTypeID"))
	end if
	rsUT.close
	set rsUT=nothing
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="4">身分證字號查詢</td>
			</tr>
			<tr>
				<td colspan="4">
				單位：<select name="UnitID">
					<option value="">請選擇</option>
<%
	if trim(request("UnitID"))<>"" then
		UnitPlus=trim(request("UnitID"))
	else
		UnitPlus=trim(session("Unit_ID"))
	end if
        '2025/12/11 增加 UnitID!='0000' 不顯示宏謙科技
	strUnit="select UnitID,UnitName from UnitInfo where UnitID!='0000' order by UnitID"
	set rsUnit=conn.execute(strUnit)
	If Not rsUnit.Bof Then rsUnit.MoveFirst 
	While Not rsUnit.Eof
%>
					<option value="<%=trim(rsUnit("UnitID"))%>" <%
					if UnitPlus=trim(rsUnit("UnitID")) then					
						response.write "selected"
					end if
					%>><%=trim(rsUnit("UnitName"))%></option>
<%	rsUnit.MoveNext
	Wend
	rsUnit.close
	set rsUnit=nothing
%>
				</select>
				<input type="button" name="BB1" value="查詢" onclick="DB_Select();">&nbsp;&nbsp;
				<input type="button" name="close" value="關閉視窗" onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				</td>
			<tr>
			<tr bgcolor="#1BF5FF">
				<td colspan="4">人員列表</td>
			</tr>
			<tr bgcolor="#FAFAF5">
				<td width="15%" align="center">身分證字號</td>
				<td width="35%" align="center">單位</td>
				<td width="25%" align="center">姓名</td>
			</tr>
<%
if trim(request("kinds"))="DB_Select" then
	strSQL=""
	if trim(request("UnitID"))<>"" then
		strSQL=" and UnitID='"&trim(request("UnitID"))&"'"
	end if
	If Not ifnull(request("Mem")) Then
		strSQL=" and chName like '%"&trim(request("Mem"))&"%'"
	End if

	If MemType = "P" Then
		if trim(request("BillMem1"))<>"" then
			if sys_City="高雄縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName then
				strSQL=" and LoginID like '%"&trim(request("BillMem1"))&"%'"
			else
				strSQL=" and ChName like '%"&trim(request("BillMem1"))&"%'"
			end if
		end if
	End if

	UTypeFlag=0
        '2025/12/11 增加 UnitID!='0000' 不顯示宏謙科技
	strProject="select LoginID,UnitID,ChName,CreditID,MemberID from MemberData where UnitID!='0000' and AccountStateID=0 and RecordstateID=0"&strSQL&" order by UnitID,LoginID"
	set rsProject=conn.execute(strProject)
	If Not rsProject.Bof Then rsProject.MoveFirst 
	While Not rsProject.Eof
			UName=""
			strUName="select UnitName,UnitTypeID from UnitInfo where UnitID='"&trim(rsProject("UnitID"))&"'"
			set rsUName=conn.execute(strUName)
			if not rsUName.eof then
				UName=trim(rsUName("UnitName"))
				UTypeID=trim(rsUName("UnitTypeID"))
			end if
			rsUName.close
			set rsUName=nothing
			if sys_City<>"台中市" and (instr(UName,"保安")<1) then
				if UserUnitTypeID=UTypeID then
					UTypeFlag=0
				else
					UTypeFlag=1
				end if
			else
				UTypeFlag=0
			end if%>
			<tr title="請點選.." onclick="Inert_Data('<%=trim(rsProject("MemberID"))%>','<%=trim(rsProject("ChName"))%>','<%=trim(rsProject("LoginID"))%>','<%=trim(rsProject("UnitID"))%>','<%=UName%>','<%=UTypeID%>','<%=UTypeFlag%>');" <%lightbarstyle 1 %>>
				<td bgcolor="#EBE5FF" align="center"><%=trim(rsProject("LoginID"))%></td>
				<td><%=UName%></td>
				<td><%=trim(rsProject("ChName"))%></td>
			</tr>
<%	rsProject.MoveNext
	Wend
	rsProject.close
	set rsProject=nothing
end if
%>

		</table>
		<input type="hidden" value="<%=request("Mem")%>" name="Mem">
		<input type="hidden" value="<%=request("MemOrder")%>" name="MemOrder">
		<input type="hidden" value="<%=request("MemType")%>" name="MemType">
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function DB_Select(){
	myForm.kinds.value="DB_Select";
	myForm.Mem.value='';
	myForm.submit();
}
function RE_Select(){
	myForm.kinds.value="DB_Select";
	myForm.submit();
}
<%
if trim(request("kinds"))="" then
	response.write "RE_Select()"
end if
%>
function Inert_Data(MCode,MName,MID,MUnitID,MUnit,UTypeID,UTypeFlag){
	if (UTypeFlag=='1'){
		alert("舉發人 " +MID+" "+MName+ " 隸屬於其他分局，請至『人員管理系統』，檢查員警資料是否正確後再做建檔!!");
	}
<%
'行人攤販要代應到案處所及舉發單位
if MemType="P" then%>
	<%if MemOrder="1" then%>
		opener.myForm.BillMemID1.value=MCode;
		opener.myForm.BillMemName1.value=MName;
		<%if sys_City="高雄縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
			opener.myForm.BillMem1.value=MID;
			opener.Layer12.innerHTML=MName;
		<%else%>
			opener.myForm.BillMem1.value=MName;
			opener.Layer12.innerHTML=MID;
		<%end if%>
		opener.TDMemErrorLog1=0;
		window.close();
	<%elseif MemOrder="2" then%>
		opener.myForm.BillMemID2.value=MCode;
		opener.myForm.BillMemName2.value=MName;
		<%if sys_City="高雄縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
			opener.myForm.BillMem2.value=MID;
			opener.Layer13.innerHTML=MName;
		<%else%>
			opener.myForm.BillMem2.value=MName;
			opener.Layer13.innerHTML=MID;
		<%end if%>
		opener.TDMemErrorLog2=0;
		window.close();	
	<%elseif MemOrder="3" then%>
		opener.myForm.BillMemID3.value=MCode;
		opener.myForm.BillMemName3.value=MName;
		<%if sys_City="高雄縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
			opener.myForm.BillMem3.value=MID;
			opener.Layer14.innerHTML=MName;
		<%else%>
			opener.myForm.BillMem3.value=MName;
			opener.Layer14.innerHTML=MID;
		<%end if%>
		opener.TDMemErrorLog3=0;
		window.close();
	<%elseif MemOrder="4" then%>
		opener.myForm.BillMemID4.value=MCode;
		opener.myForm.BillMemName4.value=MName;
		<%if sys_City="高雄縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName then%>
			opener.myForm.BillMem4.value=MID;
			opener.Layer17.innerHTML=MName;
		<%else%>
			opener.myForm.BillMem4.value=MName;
			opener.Layer17.innerHTML=MID;
		<%end if%>		
		opener.TDMemErrorLog4=0;
		window.close();
	<%end if%>
	/*
	if (opener.myForm.MemberStation.value==""){
		opener.myForm.MemberStation.value=MUnitID;
		opener.Layer5.innerHTML=MUnit;
		opener.TDStationErrorLog=0;
	}
	*/
	//if (opener.myForm.BillUnitID.value==""){
		opener.myForm.BillUnitID.value=MUnitID;
		opener.Layer6.innerHTML=MUnit;
		opener.TDUnitErrorLog=0;
	//}
<%
'雲林縣的攔停要判斷分局是否不同
elseif MemType="CarS" and sys_City="雲林縣" then
%>
	<%if MemOrder="1" then%>
		opener.myForm.BillMemID1.value=MCode;
		opener.myForm.BillMem1.value=MID;
		opener.myForm.BillMemName1.value=MName;
		opener.myForm.BillUnitTypeID1.value=UTypeID;
		opener.Layer12.innerHTML=MName;
		opener.TDMemErrorLog1=0;
			opener.myForm.BillUnitID.value=MUnitID;
			opener.Layer6.innerHTML=MUnit;
			opener.TDUnitErrorLog=0;
			if (opener.myForm.BillUnitTypeID2.value!="" && opener.myForm.BillUnitTypeID1.value!=opener.myForm.BillUnitTypeID2.value){
				alert("舉發人1與舉發人2屬於不同分局!!");
			}else if (opener.myForm.BillUnitTypeID3.value!="" && opener.myForm.BillUnitTypeID1.value!=opener.myForm.BillUnitTypeID3.value){
				alert("舉發人1與舉發人3屬於不同分局!!");
			}else if (opener.myForm.BillUnitTypeID4.value!="" && opener.myForm.BillUnitTypeID1.value!=opener.myForm.BillUnitTypeID4.value){
				alert("舉發人1與舉發人4屬於不同分局!!");
			}
		window.close();
	<%elseif MemOrder="2" then%>
		opener.myForm.BillMemID2.value=MCode;
		opener.myForm.BillMem2.value=MID;
		opener.myForm.BillMemName2.value=MName;
		opener.myForm.BillUnitTypeID2.value=UTypeID;
		opener.Layer13.innerHTML=MName;
		opener.TDMemErrorLog2=0;
		if (opener.myForm.BillUnitTypeID1.value!="" && opener.myForm.BillUnitTypeID2.value!=opener.myForm.BillUnitTypeID1.value){
			alert("舉發人2與舉發人1屬於不同分局!!");
		}else if (opener.myForm.BillUnitTypeID3.value!="" && opener.myForm.BillUnitTypeID2.value!=opener.myForm.BillUnitTypeID3.value){
			alert("舉發人2與舉發人3屬於不同分局!!");
		}else if (opener.myForm.BillUnitTypeID4.value!="" && opener.myForm.BillUnitTypeID2.value!=opener.myForm.BillUnitTypeID4.value){
			alert("舉發人2與舉發人4屬於不同分局!!");
		}
		if (opener.myForm.BillUnitID.value==""){
			opener.myForm.BillUnitID.value=MUnitID;
			opener.Layer6.innerHTML=MUnit;
			opener.TDUnitErrorLog=0;
		}
		window.close();	
	<%elseif MemOrder="3" then%>
		opener.myForm.BillMemID3.value=MCode;
		opener.myForm.BillMem3.value=MID;
		opener.myForm.BillMemName3.value=MName;
		opener.myForm.BillUnitTypeID3.value=UTypeID;
		opener.Layer14.innerHTML=MName;
		opener.TDMemErrorLog3=0;
		if (opener.myForm.BillUnitTypeID1.value!="" && opener.myForm.BillUnitTypeID3.value!=opener.myForm.BillUnitTypeID1.value){
			alert("舉發人3與舉發人1屬於不同分局!!");
		}else if (opener.myForm.BillUnitTypeID2.value!="" && opener.myForm.BillUnitTypeID3.value!=opener.myForm.BillUnitTypeID2.value){
			alert("舉發人3與舉發人2屬於不同分局!!");
		}else if (opener.myForm.BillUnitTypeID4.value!="" && opener.myForm.BillUnitTypeID3.value!=opener.myForm.BillUnitTypeID4.value){
			alert("舉發人3與舉發人4屬於不同分局!!");
		}
		if (opener.myForm.BillUnitID.value==""){
			opener.myForm.BillUnitID.value=MUnitID;
			opener.Layer6.innerHTML=MUnit;
			opener.TDUnitErrorLog=0;
		}
		window.close();
	<%elseif MemOrder="4" then%>
		opener.myForm.BillMemID4.value=MCode;
		opener.myForm.BillMem4.value=MID;
		opener.myForm.BillMemName4.value=MName;
		opener.myForm.BillUnitTypeID4.value=UTypeID;
		opener.Layer17.innerHTML=MName;
		opener.TDMemErrorLog3=0;
		if (opener.myForm.BillUnitTypeID1.value!="" && opener.myForm.BillUnitTypeID4.value!=opener.myForm.BillUnitTypeID1.value){
			alert("舉發人4與舉發人1屬於不同分局!!");
		}else if (opener.myForm.BillUnitTypeID2.value!="" && opener.myForm.BillUnitTypeID4.value!=opener.myForm.BillUnitTypeID2.value){
			alert("舉發人4與舉發人2屬於不同分局!!");
		}else if (opener.myForm.BillUnitTypeID3.value!="" && opener.myForm.BillUnitTypeID4.value!=opener.myForm.BillUnitTypeID3.value){
			alert("舉發人4與舉發人3屬於不同分局!!");
		}
		if (opener.myForm.BillUnitID.value==""){
			opener.myForm.BillUnitID.value=MUnitID;
			opener.Layer6.innerHTML=MUnit;
			opener.TDUnitErrorLog=0;
		}
		window.close();
	<%end if%>
<%
'車輛只要代舉發單位
else
%>
	<%if MemOrder="1" then%>
		opener.myForm.BillMemID1.value=MCode;
		opener.myForm.BillMem1.value=MID;
		opener.myForm.BillMemName1.value=MName;
		opener.Layer12.innerHTML=MName;
		opener.TDMemErrorLog1=0;
			opener.myForm.BillUnitID.value=MUnitID;
			opener.Layer6.innerHTML=MUnit;
			opener.TDUnitErrorLog=0;
		window.close();
	<%elseif MemOrder="2" then%>
		opener.myForm.BillMemID2.value=MCode;
		opener.myForm.BillMem2.value=MID;
		opener.myForm.BillMemName2.value=MName;
		opener.Layer13.innerHTML=MName;
		opener.TDMemErrorLog2=0;
		if (opener.myForm.BillUnitID.value==""){
			opener.myForm.BillUnitID.value=MUnitID;
			opener.Layer6.innerHTML=MUnit;
			opener.TDUnitErrorLog=0;
		}
		window.close();	
	<%elseif MemOrder="3" then%>
		opener.myForm.BillMemID3.value=MCode;
		opener.myForm.BillMem3.value=MID;
		opener.myForm.BillMemName3.value=MName;
		opener.Layer14.innerHTML=MName;
		opener.TDMemErrorLog3=0;
		if (opener.myForm.BillUnitID.value==""){
			opener.myForm.BillUnitID.value=MUnitID;
			opener.Layer6.innerHTML=MUnit;
			opener.TDUnitErrorLog=0;
		}
		window.close();
	<%elseif MemOrder="4" then%>
		opener.myForm.BillMemID4.value=MCode;
		opener.myForm.BillMem4.value=MID;
		opener.myForm.BillMemName4.value=MName;
		opener.Layer17.innerHTML=MName;
		opener.TDMemErrorLog3=0;
		if (opener.myForm.BillUnitID.value==""){
			opener.myForm.BillUnitID.value=MUnitID;
			opener.Layer6.innerHTML=MUnit;
			opener.TDUnitErrorLog=0;
		}
		window.close();
	<%end if%>
<%
end if
%>
}
</script>
</html>
