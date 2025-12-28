<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!-- #include file="../Common/Bannernodata.asp"-->
<%

DB_Selt=trim(request("DB_Selt"))
if DB_Selt="Selt" then
	strwhere=""

	if request("Sys_year")<>"" then
		ArgueDate1=gOutDT(request("Sys_year")&"0101")&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_year")&"1231")&" 23:59:59"
		strwhere=strwhere&" and IllegalDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end If 

	if request("Sys_type")="1" then
		strwhere=strwhere&" and not Exists(select 'N' from PasserJude where billsn=passerBase.sn)"

	elseif request("Sys_type")="2" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserJude where billsn=passerBase.sn) and not Exists(select 'N' from PasserSend where billsn=passerBase.sn)"

	elseif request("Sys_type")="3" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserCreditor where PetitionDate is not null and billsn=passerBase.sn)"

	elseif request("Sys_type")="4" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserSend where billsn=passerBase.sn) and not Exists(select 'N' from PasserCreditor where PetitionDate is not null and billsn=passerBase.sn)"

	elseif request("Sys_type")="5" then
		strwhere=strwhere&" and TRUNC(SYSDATE-DEALLINEDATE) > 184 and Not Exists(select 'N' from PasserJude where billsn=passerBase.SN)"

	elseif request("Sys_type")="6" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserJude where TRUNC(SYSDATE-JudeDate) > 184 and billsn=passerBase.SN)" & _
		" and Not Exists(select 'N' from PasserSend where billsn=passerBase.SN)"

	end If

	if request("IllegalDate1")<>"" and request("IllegalDate2")<>"" then
		ArgueDate1=gOutDT(request("IllegalDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("IllegalDate2"))&" 23:59:59"
		strwhere=strwhere&" and IllegalDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end If 
	
	if request("Sys_DetailType")="1" then
		strwhere=strwhere&" and not Exists(select 'N' from PasserJude where billsn=passerBase.sn)"

	elseif request("Sys_DetailType")="2" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserJude where billsn=passerBase.sn) and not Exists(select 'N' from PasserSend where billsn=passerBase.sn)"

	elseif request("Sys_DetailType")="3" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserCreditor where PetitionDate is not null and billsn=passerBase.sn)"

	elseif request("Sys_DetailType")="4" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserSend where billsn=passerBase.sn) and not Exists(select 'N' from PasserCreditor where PetitionDate is not null and billsn=passerBase.sn)"

	elseif request("Sys_type")="5" then
		strwhere=strwhere&" and TRUNC(SYSDATE-DEALLINEDATE) > 184 and Not Exists(select 'N' from PasserJude where billsn=passerBase.SN)"

	elseif request("Sys_type")="6" then
		strwhere=strwhere&" and Exists(select 'Y' from PasserJude where TRUNC(SYSDATE-JudeDate) > 184 and billsn=passerBase.SN)" & _
		" and Not Exists(select 'N' from PasserSend where billsn=passerBase.SN)"

	end If 

	strSQL="select SN,BillNo,IllegalDate,(select UnitName from UnitInfo where UnitID=PasserBase.BillUnitID) UnitName," & _
			"(select ChName from MemberData where MemberID=PasserBase.BillMemID1) UnitChName," & _
			"Driver,Rule1,Rule2,Rule3,Forfeit1,Forfeit2,Forfeit3,(select max(JudeDate) from PasserJude where billsn=passerBase.sn) JudeDate," & _
			"(select max(SendDate) from PasserSend where billsn=passerBase.sn) SendDate," & _
			"(select max(petitiondate) from PasserCreditor where PetitionDate is not null and billsn=passerBase.sn) petitiondate " & _
			"from passerbase where MemberStation=(select UnitID from UnitInfo where UnitName='"&trim(Request("Sys_UnitName"))&"')" & _
			" and recordstateid=0 and billstatus <> 9" & strwhere & " order by UnitName,BillNo"

	set rs=conn.execute(strSQL)

end If 
	
	Set cntobj = Server.CreateObject("Scripting.Dictionary")

	strSQL="select to_number(to_char(illegaldate,'YYYY'))-1911 illegal_year," & _
			"sum((case when (select count(1) from PASSERCREDITOR where PetitionDate is not null and billsn=passerBase.sn)>0 then 1 else 0 end)) CreditCnt," & _
			"sum((case when (select count(1) from PASSERJUDE where billsn=passerBase.sn) > 0 " & _
			"and (select count(1) from PASSERSEND where billsn=passerBase.sn)=0 then 1 else 0 end)) SendCnt," & _
			"sum((case when (select count(1) from PASSERJUDE where billsn=passerBase.sn) = 0 then 1 else 0 end)) JudeCnt," & _
			"sum((case when (select count(1) from PASSERSEND where billsn=passerBase.sn) > 0 " & _
			"and (select count(1) from PASSERCREDITOR where PetitionDate is not null and billsn=passerBase.sn)=0 then 1 else 0 end)) OtherCnt," & _
			"sum((select count(1) from passerbase pr where TRUNC(SYSDATE-DEALLINEDATE) > 184" & _
			" and Not Exists(select 'N' from PasserJude where billsn=pr.SN) and SN=PasserBase.SN)) NotJudeCnt," & _
			"sum((select count(1) from passerbase pr where Exists(select 'Y' from PasserJude where TRUNC(SYSDATE-JudeDate) > 184" & _
			" and billsn=pr.SN) and Not Exists(select 'N' from PasserSend where billsn=pr.SN) and SN=PasserBase.SN)) NotPasserSend " & _
			"from passerbase where MemberStation=(select UnitID from UnitInfo where UnitName='"&trim(Request("Sys_UnitName"))&"') " & _
			"and recordstateid=0 and billstatus <> 9 group by (to_number(to_char(illegaldate,'YYYY'))-1911) order by illegal_year DESC"

	set rs1 = Conn.Execute(strSQL)
	maxyear=0:minyear=0
	While Not rs1.Eof

		If maxyear = 0 Then maxyear=cdbl(rs1("illegal_year"))
		minyear=cdbl(rs1("illegal_year"))

		cntobj.Add rs1("illegal_year")&"_1",cdbl(rs1("JudeCnt"))
		cntobj.Add rs1("illegal_year")&"_2",cdbl(rs1("SendCnt"))
		cntobj.Add rs1("illegal_year")&"_3",cdbl(rs1("CreditCnt"))
		cntobj.Add rs1("illegal_year")&"_4",cdbl(rs1("OtherCnt"))
		cntobj.Add rs1("illegal_year")&"_5",cdbl(rs1("NotJudeCnt"))
		cntobj.Add rs1("illegal_year")&"_6",cdbl(rs1("NotPasserSend"))

		rs1.MoveNext
	Wend  
	rs1.close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>單位裁罰狀況明細表</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<form name="myForm" method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33" height="33"><%=Request("Sys_UnitName")%>裁罰狀況明細表</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="1" cellpadding="1" cellspacing="1">
				<tr align="center">
					<th height="30">年度</th>
					<%
					For i = maxyear to minyear step -1
						Response.Write "<th height=""30"">"& i &"</th>"
					Next
					%>
				</tr>
					<%
					response.write "<tr bgcolor='#FFFFFF' align='center'>"
					response.write "<td bgcolor=""#FFCC33"" align='right'>未裁罰件數</td>"
					For i = maxyear to minyear step -1
						
						response.write "<td align='left' onclick=""funSetQry('" & i & "','1');"">"&cntobj.Item(i & "_1")&"</td>"

					Next
					response.write "</tr>"

					response.write "<tr bgcolor='#FFFFFF' align='center'>"
					response.write "<td bgcolor=""#FFCC33"" align='right'>未移送件數</td>"
					For i = maxyear to minyear step -1
						
						response.write "<td align='left' onclick=""funSetQry('" & i & "','2');"">"&cntobj.Item(i & "_2")&"</td>"

					Next
					response.write "</tr>"

					response.write "<tr bgcolor='#FFFFFF' align='center'>"
					response.write "<td bgcolor=""#FFCC33"" align='right'>取得債權憑證</td>"
					For i = maxyear to minyear step -1
						
						response.write "<td align='left' onclick=""funSetQry('" & i & "','3');"">"&cntobj.Item(i & "_3")&"</td>"

					Next
					response.write "</tr>"

					response.write "<tr bgcolor='#FFFFFF' align='center'>"
					response.write "<td bgcolor=""#FFCC33"" align='right'>移送執行中</td>"
					For i = maxyear to minyear step -1
						
						response.write "<td align='left' onclick=""funSetQry('" & i & "','4');"">"&cntobj.Item(i & "_4")&"</td>"

					Next
					response.write "</tr>"

					response.write "<tr bgcolor='#FFFFFF' align='center'>"
					response.write "<td bgcolor=""#FFCC33"" align='right'>逾期未裁決</td>"
					For i = maxyear to minyear step -1
						
						response.write "<td align='left' onclick=""funSetQry('" & i & "','5');"">"&cntobj.Item(i & "_5")&"</td>"

					Next
					response.write "</tr>"

					response.write "<tr bgcolor='#FFFFFF' align='center'>"
					response.write "<td bgcolor=""#FFCC33"" align='right'>逾期未移送</td>"
					For i = maxyear to minyear step -1
						
						response.write "<td align='left' onclick=""funSetQry('" & i & "','6');"">"&cntobj.Item(i & "_6")&"</td>"

					Next
					response.write "</tr>"
					%>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table border="1" bgcolor="#FFFFFF" width="100%">
				<tr>
					<td>違規日期：</td>
					<td>
						<input name="IllegalDate1" class="btn1" type="text" value="<%=request("IllegalDate1")%>" size="5" maxlength="8" onkeyup="chknumber(this);">
						<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('IllegalDate1');">
						~
						<input name="IllegalDate2" class="btn1" type="text" value="<%=request("IllegalDate2")%>" size="5" maxlength="8" onkeyup="chknumber(this);">
						<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('IllegalDate2');">
					</td>
					<td>項目</td>
					<td>
						<select name="Sys_DetailType" class="btn1">
							<option value="">請選擇</option>
							<option value="1"<%If Request("Sys_DetailType") = "1" Then Response.Write " selected"%>>未裁罰件數</option>
							<option value="2"<%If Request("Sys_DetailType") = "2" Then Response.Write " selected"%>>未移送件數</option>
							<option value="3"<%If Request("Sys_DetailType") = "3" Then Response.Write " selected"%>>取得債權憑證</option>
							<option value="4"<%If Request("Sys_DetailType") = "4" Then Response.Write " selected"%>>移送執行中</option>
							<option value="5"<%If Request("Sys_DetailType") = "5" Then Response.Write " selected"%>>逾期未裁決</option>
							<option value="6"<%If Request("Sys_DetailType") = "6" Then Response.Write " selected"%>>逾期未移送</option>
						</select>
					</td>
					<td>
						<input type="submit" name="btnSelt" value="查詢" onClick='funSelt();'>&nbsp;&nbsp;
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" height="33">

		</td>
	</tr>
	<%if DB_Selt="Selt" then%>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="1" cellpadding="1" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="30">序號</th>
					<th height="30">違規日</th>
					<th height="34">舉發單號</th>
					<th height="34">舉發單位</th>
					<th height="34">舉發人</th>
					<th height="34">違規人</th>
					<th height="34">法條</th>
					<th height="34">金額</th>
					<th height="34">裁決日</th>
					<th height="34">移送日</th>
					<th height="34">取得債權日</th>
				</tr><%
					filecnt=0
					while Not rs.eof 
						filecnt=filecnt+1
						response.write "<tr bgcolor='#FFFFFF' align='center' "
						lightbarstyle 0 
						Response.Write " onclick=""funPasserDetail('" & rs("SN") & "');"""
						response.write ">"
						response.write "<td align='left'>"&filecnt&"</td>"
						response.write "<td align='left'>"&gInitDT(trim(rs("IllegalDate")))&"</td>"
						response.write "<td align='left'>"&rs("BillNo")&"</td>"
						response.write "<td align='left'>"&rs("UnitName")&"</td>"
						response.write "<td align='left'>"&rs("UnitChName")&"</td>"
						response.write "<td align='left'>"&rs("Driver")&"</td>"

						response.write "<td align='left'>"
						Response.Write rs("Rule1")
						If not ifnull(rs("Rule2")) Then Response.Write "/"&rs("Rule2")
						If not ifnull(rs("Rule3")) Then Response.Write "/"&rs("Rule3")
						Response.Write "</td>"

						response.write "<td align='left'>"
						Response.Write rs("Forfeit1")
						If not ifnull(rs("Forfeit2")) Then Response.Write "/"&rs("Forfeit2")
						If not ifnull(rs("Forfeit3")) Then Response.Write "/"&rs("Forfeit3")
						Response.Write "</td>"

						response.write "<td align='left'>"&gInitDT(trim(rs("JudeDate")))&"</td>"
						response.write "<td align='left'>"&gInitDT(trim(rs("SendDate")))&"</td>"
						response.write "<td align='left'>"&gInitDT(trim(rs("petitiondate")))&"</td>"

						response.write "</tr>"
						rs.movenext
					wend
					rs.close%>
			</table>
		</td>
	</tr>
	<%end If 
	set cntobj=nothing
%>
</table>
<input type="Hidden" name="Sys_UnitName" value="<%=Request("Sys_UnitName")%>">
<input type="Hidden" name="DB_Selt" value="<%=DB_Selt%>">

<input type="Hidden" name="Sys_year" value="">
<input type="Hidden" name="Sys_type" value="">
<input type="Hidden" name="BillSn" value="">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">

function funSelt(){
	myForm.DB_Selt.value="Selt";
	myForm.submit();
}

function funPasserDetail(BillSN){

	myForm.BillSn.value=BillSN;

	UrlStr="../Query/ViewBillBaseData_people.asp";
	myForm.action=UrlStr;
	myForm.target="DetailPeople";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funSetQry(sysYear,sysType){
	myForm.Sys_year.value=sysYear;
	myForm.Sys_type.value=sysType;

	myForm.DB_Selt.value="Selt";
	myForm.submit();
}
</script>
<%conn.close%>