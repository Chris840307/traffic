<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set rsUnit=conn.Execute(strSQL)
DB_UnitID=trim(rsUnit("UnitID"))
DB_UnitName=trim(rsUnit("UnitName"))
theContactTel=trim(rsUnit("Tel"))
theSubUnitSecBossName=trim(rsUnit("SecondManagerName"))
theBigUnitBossName=trim(rsUnit("ManageMemberName"))
theBankAccount=trim(rsUnit("BankAccount"))
rsUnit.close

if trim(request("DB_Add"))="ADD" then
	strSQL="insert into PasserUrge(BillSN,BillNo,OpenGovNumber,UrgeDate,UrgeTypeID,SendAddress,ForFeit,BigUnitBossName,SubUnitSecBossName,ContactTel,RecordStateID,RecordDate,RecordMemberID) values("&request("BillSN")&",'"&request("BillNo")&"','"&request("Sys_OpenGovNumber")&"',"&funGetDate(gOutDT(request("Sys_UrgeDate")),0)&",'"&request("Sys_UrgeTypeID")&"','"&request("Sys_SendAddress")&"','"&request("Sys_ForFeit")&"','"&request("Sys_BigUnitBossName")&"','"&request("Sys_SubUnitSecBossName")&"','"&request("Sys_ContactTel")&"',0,"&funGetDate(now,1)&","&Session("User_ID")&")"
	conn.execute(strSQL)

	strSQL="update PasserUrge set SendAddress='"&request("Sys_SendAddress")&"' where billsn="&request("BillSN")
	conn.execute(strSQL)

	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
elseif trim(request("DB_Add"))="Update" then
	strSQL="Update PasserUrge set OpenGovNumber='"&request("Sys_OpenGovNumber")&"',UrgeDate="&funGetDate(gOutDT(request("Sys_UrgeDate")),0)&",UrgeTypeID='"&request("Sys_UrgeTypeID")&"',SendAddress='"&request("Sys_SendAddress")&"',ForFeit='"&request("Sys_ForFeit")&"',BigUnitBossName='"&request("Sys_BigUnitBossName")&"',SubUnitSecBossName='"&request("Sys_SubUnitSecBossName")&"',ContactTel='"&request("Sys_ContactTel")&"',RecordStateID=0,RecordDate="&funGetDate(now,1)&",RecordMemberID="&Session("User_ID")&" where BillSN="&request("BillSN")&" and BillNo='"&request("BillNo")&"'"
	conn.execute(strSQL)

	If not ifnull(request("Sys_SendAddress")) Then
		strSQL="update passerbase set DriverAddress='"&request("Sys_SendAddress")&"' where sn="&request("BillSN")

		conn.execute(strSQL)
	End if

	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if

'if trim(request("DB_Add"))<>"" then
'       if trim(request("Sys_ContactTel")) <>"" then 
'  	 strSQL="Update UnitInfo set TEL='"&trim(request("Sys_ContactTel"))&"' where UnitID='"&DB_UnitID&"'"
' 	 conn.execute(strSQL)
'       end if
'       if trim(request("Sys_BigUnitBossName")) <>"" then 
'  	 strSQL="Update UnitInfo set ManageMemberName='"&trim(request("Sys_BigUnitBossName"))&"' where UnitID='"&DB_UnitID&"'"
' 	 conn.execute(strSQL)
'       end if
'       if trim(request("Sys_SubUnitSecBossName")) <>"" then 
'  	 strSQL="Update UnitInfo set SecondManagerName='"&trim(request("Sys_SubUnitSecBossName"))&"' where UnitID='"&DB_UnitID&"'"
' 	 conn.execute(strSQL)
'       end if
'
'       if trim(request("Sys_BankAccount")) <>"" then 
'  	 strSQL="Update UnitInfo set BankAccount='"&trim(request("Sys_BankAccount"))&"' where UnitID='"&DB_UnitID&"'"
' 	 conn.execute(strSQL)
'       end if
'
'end if

strSql="select a.SN as BillSN,a.BillNo,a.ForFeit1,a.RuleVer,a.DoubleCheckStatus,a.RecordMemberID,b.OpenGovNumber as JudeOGN,c.OpenGovNumber as UrgeOGN,c.UrgeDate,c.BigUnitBossName,c.SubUnitSecBossName,c.ContactTel,c.SendAddress,c.UrgeTypeID,c.ForFeit,a.Driver,a.DriverBirth,a.DriverID,a.DriverAddress,a.IllegalDate,a.IllegalAddress,a.DealLineDate,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID from PasserBase a,PasserJude b,PasserUrge c where a.SN="&trim(request("PBillSN"))&" and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+)"
set rsfound=conn.execute(strSql)

if trim(rsfound("UrgeDate"))<>"" then
	UrgeDate=gInitDT(rsfound("UrgeDate"))
else
	UrgeDate=gInitDT(date)
end if


strState="select * from PasserUrge where BillSN="&trim(request("PBillSN"))
set rsState=conn.execute(strState)
BillEof=0
if rsState.eof then BillEof=1
rsState.close

Sys_BillNum=trim(rsfound("BillNo"))
If sys_City="台南縣" Then Sys_BillNum=(year(now)-1911)&right("000000"&trim(rsfound("DoubleCheckStatus")),4)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>違反道路交通管理事件催繳</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>

<body onkeydown="KeyDown()">
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td height="20" bgcolor="#1BF5FF">違反道路交通管理事件催繳</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td bgcolor="#EBE5FF" align="right">舉發單號</td>
					<td><%=rsfound("BillNo")%></td>
					<td align="right" nowrap bgcolor="#EBE5FF">催繳文號</td>
					<td colspan="5">
						<input name="Sys_OpenGovNumber" class="btn1" type="text" size="12" maxlength="12" value="<%
							if Not ifnull(rsfound("UrgeOGN")) then
								response.write rsfound("UrgeOGN")
							else
								response.write Sys_BillNum
							end if
							%>">
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#EBE5FF">舉發單位</td>
					<td><%=DB_UnitName%></td>
					<td align="right" nowrap bgcolor="#EBE5FF">受處分人</td>
					<td><%=rsfound("Driver")%></td>
					<td align="right" nowrap bgcolor="#EBE5FF">催繳日期</td>
					<td colspan="3">
						<input name="Sys_UrgeDate" class="btn1" type="text" size="4" maxlength="8" value="<%=UrgeDate%>" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_UrgeDate');">
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#EBE5FF">出生日期</td>
					<td><%=gInitDT(rsfound("DriverBirth"))%></td>
					<td align="right" nowrap bgcolor="#EBE5FF">承辦人</div></td>
					<td><%=Session("Ch_Name")%></td>
					<td bgcolor="#EBE5FF" align="right">分局長</td>
					<td><input name="Sys_SubUnitSecBossName" class="btn1" type="text" size="12" maxlength="12" value="<%
							if trim(rsfound("SubUnitSecBossName"))<>"" then
								theSubUnitSecBossName=trim(rsfound("SubUnitSecBossName"))
							end if
							if trim(rsfound("BigUnitBossName"))<>"" then
								theBigUnitBossName=trim(rsfound("BigUnitBossName"))
							end if
							if trim(rsfound("ContactTel"))<>"" then
								theContactTel=trim(rsfound("ContactTel"))
							end if
							response.write trim(theSubUnitSecBossName)
							session("Sys_UnitChName")=trim(theSubUnitSecBossName)
							%>">
					</td>
					<td bgcolor="#EBE5FF" nowrap align="right">局長</td>
					<td><input name="Sys_BigUnitBossName" class="btn1" type="text" size="12" maxlength="12" value="<%=trim(theBigUnitBossName)%>"></td>
				</tr>
				<tr>
					<td bgcolor="#EBE5FF" nowrap align="right">身分證號</td>
					<td><%=rsfound("DriverID")%></td>
					<td bgcolor="#EBE5FF" nowrap align="right">聯絡電話</td>
					<td>
						<input name="Sys_ContactTel" class="btn1" type="text" size="12" maxlength="12" value="<%=trim(theContactTel)%>">
					</td>
					<td bgcolor="#EBE5FF" nowrap>劃撥帳號</td>
					<td colspan="3">
						<%if Not ifnull(theBankAccount) then
							response.write theBankAccount
						else
							response.write "<input name=""Sys_BankAccount"" class=""btn1"" type=""text"" size=""31"" maxlength=""30"">"
						end if%>
					</td>
				</tr>
				<tr>
					<td bgcolor="#EBE5FF" nowrap align="right">住址</td>
					<td><%=rsfound("DriverAddress")%></td>
					<td bgcolor="#EBE5FF" nowrap align="right">催繳方式</td>
					<td colspan="3">
						<input name="Sys_UrgeTypeID" class="btn1" type="radio" value="0"<%if trim(rsfound("UrgeTypeID"))="0" then response.write " checked"%>>
						電話
						<input name="Sys_UrgeTypeID" class="btn1" type="radio" value="1"<%if trim(rsfound("UrgeTypeID"))="1" then response.write " checked"%>>
						信函
						<input name="Sys_UrgeTypeID" class="btn1" type="radio" value="2"<%if trim(rsfound("UrgeTypeID"))="2" then response.write " checked"%>>
						催繳書
					</td>
				</tr>
				<tr>
					<td bgcolor="#EBE5FF" nowrap align="right">違規時間</td>
					<td><%=gInitDT(rsfound("IllegalDate"))%></td>
					<td bgcolor="#EBE5FF" nowrap align="right">罰款金額</td>
					<td colspan="5">
						  <input name="Sys_ForFeit" class="btn1" type="text" size="12" maxlength="12" value="<%
						  if trim(rsfound("ForFeit"))<>"" then
								response.write rsfound("ForFeit")
						  else
								response.write rsfound("ForFeit1")
						  end if
						  %>" onkeyup="value=value.replace(/[^\d]/g,'')">
					</td>
				</tr>
				<tr>
					<td bgcolor="#EBE5FF" nowrap align="right">違規地點</td>
					<td><%=rsfound("IllegalAddress")%></td>
					<td bgcolor="#EBE5FF" nowrap align="right">違規人寄送地址</td>
					<td colspan="3"><input name="Sys_SendAddress" class="btn1" type="text" value="<%	if trim(rsfound("SendAddress"))<>"" then
						response.write rsfound("SendAddress")
					else
						response.write rsfound("DriverAddress")
					end if%>" size="48" maxlength="50"></td>
				</tr>
				<tr>
					<td bgcolor="#EBE5FF" nowrap align="right">應到案日期</td>
					<td><%=gInitDT(rsfound("DealLineDate"))%></td>
					<td bgcolor="#EBE5FF">&nbsp;</td>
					<td colspan="5">&nbsp;</td>
				</tr>
				<tr>
					<td bgcolor="#EBE5FF" nowrap align="right">違反法條</td>
					<td><%
						if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
							response.write trim(rsfound("Rule1"))&"，"
							strRule1="select * from Law where ItemID='"&trim(rsfound("Rule1"))&"' and VERSION='"&trim(rsfound("RuleVer"))&"'"
							set rsRule1=conn.execute(strRule1)
							if not rsRule1.eof then
								response.write cint(trim(rsRule1("Level1")))
								if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
									response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
								else
									response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
								end if
								response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
								response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
								response.write "<br>"&trim(rsRule1("IllegalRule"))
							end if
							rsRule1.close
							set rsRule1=nothing
						end if	
						if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
							response.write "<br>"&trim(rsfound("Rule2"))&"，"
							strRule1="select * from Law where ItemID='"&trim(rsfound("Rule2"))&"' and VERSION='"&trim(rsfound("RuleVer"))&"'"
							set rsRule1=conn.execute(strRule1)
							if not rsRule1.eof then
								response.write cint(trim(rsRule1("Level1")))
								if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
									response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
								else
									response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
								end if
								response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
								response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
								response.write "<br>"&trim(rsRule1("IllegalRule"))
							end if
							rsRule1.close
							set rsRule1=nothing
						end if	
						if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
							response.write "<br>"&trim(rsfound("Rule3"))&"，"
							strRule1="select * from Law where ItemID='"&trim(rsfound("Rule3"))&"' and VERSION='"&trim(rsfound("RuleVer"))&"'"
							set rsRule1=conn.execute(strRule1)
							if not rsRule1.eof then
								response.write cint(trim(rsRule1("Level1")))
								if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
									response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
								else
									response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
								end if
								response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
								response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
								response.write "<br>"&trim(rsRule1("IllegalRule"))
							end if
							rsRule1.close
							set rsRule1=nothing
						end if	
						if trim(rsfound("Rule4"))<>"" and not isnull(rsfound("Rule4")) then
							response.write "<br>"&trim(rsfound("Rule4"))&"，"
							strRule1="select * from Law where ItemID='"&trim(rsfound("Rule4"))&"' and VERSION='"&trim(rsfound("RuleVer"))&"'"
							set rsRule1=conn.execute(strRule1)
							if not rsRule1.eof then
								response.write cint(trim(rsRule1("Level1")))
								if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
									response.write "&nbsp; ,"&cint(trim(rsRule1("Level1")))
								else
									response.write "&nbsp; ,"&cint(trim(rsRule1("Level2")))
								end if
								response.write "&nbsp; ,"&cint(trim(rsRule1("Level3")))
								response.write "&nbsp; ,"&cint(trim(rsRule1("Level4")))
								response.write "<br>"&trim(rsRule1("IllegalRule"))
							end if
							rsRule1.close
							set rsRule1=nothing
						end if
						%>
					</td>
					<td nowrap bgcolor="#EBE5FF">&nbsp;</td>
					<td colspan="5">&nbsp;</td>
				</tr>
				<tr>
					<td nowrap bgcolor="#EBE5FF" align="right">代保管物品</td>
					<td><%	FastenerTemp=""
							strFastener="select Confiscate from PasserConfiscate where BillSN="&trim(request("PBillSN"))
							set rsFastener=conn.execute(strFastener)
							If Not rsFastener.Bof Then rsFastener.MoveFirst 
							While Not rsFastener.Eof
								if FastenerTemp="" then
									FastenerTemp=trim(rsFastener("Confiscate"))
								else
									FastenerTemp=FastenerTemp&"，"&trim(rsFastener("Confiscate"))
								end if
								rsFastener.MoveNext
							Wend
							rsFastener.close
							set rsFastener=nothing
							response.write FastenerTemp
						%>
					</td>
					<td bgcolor="#EBE5FF" align="right"></td>
					<td colspan="5"></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#1BF5FF">
			<input name="btnadd" type="button" value=" 確 定 " onclick="funAdd();"> 
			<%if sys_City<>"基隆市" then%>
				<input name="btnprint" type="button" value="列印催繳書"  onclick='PrintReports();'>
				<img src="space.gif" width="20" height="5">
			<%end if%>
			<input name="btnexit" type="button" value=" 關 閉 " onclick="funExt();">
			<img src="space.gif" width="20" height="5">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Add" value="">
<input type="Hidden" name="BillSN" value="<%=rsfound("BillSN")%>">
<input type="Hidden" name="BillNo" value="<%=rsfound("BillNo")%>">
<input type="Hidden" name="PBillSN" value="<%=request("PBillSN")%>">
<input type="hidden" name="BillEof" value="<%=BillEof%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}
}
function funAdd(){
	runServerScript("chkAddNew.asp?BillSN="+myForm.BillSN.value+"&BillNo="+myForm.BillNo.value);
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		opener.myForm.submit();
		self.close();
	}
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}
function PrintReports(){
	if(myForm.BillEof.value=='0'){
		//UrlStr="PasserUrge_Word.asp?PBillSN=<%=request("PBillSN")%>";		
		UrlStr="PasserJudeBatList.asp?BillSN=<%=request("PBillSN")%>&Sys_PasserUrge=1&Session_JudeName=<%=Session("Ch_Name")%>";
		newWin(UrlStr,"inputWin",900,550,50,10,"yes","yes","yes","no");
		//window.parent.frames("mainFrame").DP();
	}else{
		alert("請先進行存檔!!");
	}
}
</script>
<%rsfound.close
conn.close%>
