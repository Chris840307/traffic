<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!-- #include file="../Common/Bannernoimage.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>行人慢車道路障礙裁罰資料修改作業</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
'AuthorityCheck(223)
'==========POST=========
'單號  error
'if trim(request("BillSN"))="" then
'	theBillno=trim(request("BillSN"))
'end if
'new代表新增案件 , update 代表資料庫已有該案件
AuthorityCheck(235)
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

F5str="116"
F5StrName="F5"
F6Str="117"
F6StrName="F6"
if sys_City="高雄市" or sys_City="高港局" then
	F5str="117"
	F5StrName="F6"
	F6Str="116"
	F6StrName="F5"
end if

if trim(request("filetype"))="" then
	thefiletype=""
else
	thefiletype=trim(request("filetype"))
end if
'告發類別
if trim(request("Billtype"))="" then
	theBilltype=""
else
	theBilltype=trim(request("Billtype"))
end if
'==========cookie==========
'填單人代碼
theRecordMemberID=trim(Session("User_ID"))
'=========fucntion=========
function DateFormatChange(changeDate)
	DateFormatChange=funGetDate(gOutDT(changeDate),1)

	'if left(trim(changeDate),1)="0" then
	'	theFormatDate=cint(mid(trim(changeDate),2,2))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'else
	'	theFormatDate=cint(left(trim(changeDate),3))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'end if
	'DateFormatChange=theFormatDate
end function

'==========================
'修改告發單
if trim(request("kinds"))="DB_insert" then
	'違規日期
	theIllegalDate=""
	if trim(request("BillFillDate"))<>"" then
		theIllegalDate=funGetDate(gOutDT(request("IllegalDate") ) &" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2),1)
	else
		theIllegalDate = "null"
	end if	


	'檢查是否有罰款金額
	if trim(request("ForFeit1"))="" then
		theForFeit1="null"
	else
		theForFeit1=trim(request("ForFeit1"))
	end if
	if trim(request("ForFeit2"))="" then
		theForFeit2="null"
	else
		theForFeit2=trim(request("ForFeit2"))
	end if
	if trim(request("ForFeit3"))="" then
		theForFeit3="null"
	else
		theForFeit3=trim(request("ForFeit3"))
	end if
	if trim(request("ForFeit4"))="" then
		theForFeit4="null"
	else
		theForFeit4=trim(request("ForFeit4"))
	end if
	'駕駛人生日
	theDriverBirth=""
	if trim(request("DriverBrith"))<>"" then
		theDriverBirth=DateFormatChange(trim(request("DriverBrith")))
	else 
		theDriverBirth = "null"
	end if
	
	'填單日期
	theBillFillDate=""
	if trim(request("BillFillDate"))<>"" then
		theBillFillDate=DateFormatChange(trim(request("BillFillDate")))
	else
		theBillFillDate = "null"
	end if
	'應到案日期
	theDealLineDate=""
	if trim(request("DealLineDate"))<>"" then
		theDealLineDate=DateFormatChange(trim(request("DealLineDate")))
	else
		theDealLineDate="null"
	end if

	if trim(request("Billtype"))="" then '現在一律變為1 表示攔停
		theBilltype="1"
	else
		theBilltype=trim(request("Billtype"))
	end if

	'建檔日期
	theRecordDate=year(now)&"/"&month(now)&"/"&day(now)&" "&hour(now)&":"&minute(now)&":"&second(now)

	'PasserBase
	zipid=""
	
strSQL = "select ZipID,ZipName from Zip where ZipName like '"&replace(left(trim(request("DriverAddress")),5),"臺","台")&"%'"

	set rszip=conn.execute(strSQL)
	If not rszip.eof Then
		zipid=rszip("zipid")
	else
		rszip.close
		
strSQL = "select ZipID,ZipName from Zip where ZipName like '"&replace(left(trim(request("DriverAddress")),3),"臺","台")&"%'"
		set rszip=conn.execute(strSQL)
		if Not rszip.eof then zipid=rszip("zipid")
	end if
	rszip.close
	If sys_City="高雄市" Then
		UpdateAdd=",IllegalZip='"&trim(request("IllegalZip"))&"'"
	End if	
	strUpd="update PasserBase set BillTypeID='"&trim(theBilltype)&"'" &_
		",BillNo='"&UCase(trim(request("Billno1")))&"',IllegalDate="&theIllegalDate&_
		",IllegalAddressID='"&trim(request("IllegalAddressID"))&"',IllegalAddress='"&trim(request("IllegalAddress"))&"'" &_
		",Rule1='"&trim(request("Rule1"))&"',ForFeit1="&theForFeit1 &_
		",Rule2='"&trim(request("Rule2"))&"',ForFeit2="&theForFeit2&",Rule3='"&trim(request("Rule3"))&"'" &_
		",ForFeit3="&theForFeit3&",Rule4='"&trim(request("Rule4"))&"',ForFeit4="&theForFeit4 &_
		",ProjectID='"&trim(request("ProjectID"))&"',DriverID='"&UCase(trim(request("DriverPID")))&"'" &_
		",DriverBirth="&theDriverBirth&",Driver='"&trim(request("DriverName"))&"'" &_
		",DriverAddress='"&trim(request("DriverAddress"))&"',DriverZip='"&trim(zipid)&"'" &_
		",MemberStation='"&trim(request("MemberStation"))&"',BillUnitID='"&trim(request("BillUnitID"))&"'" &_
		",BillMemID1='"&trim(request("BillMemID1"))&"',BillMem1='"&trim(request("BillMemName1"))&"'" &_
		",BillMemID2='"&trim(request("BillMemID2"))&"',BillMem2='"&trim(request("BillMemName2"))&"'" &_
		",BillMemID3='"&trim(request("BillMemID3"))&"',BillMem3='"&trim(request("BillMemName3"))&"'" &_
		",BillMemID4='"&trim(request("BillMemID4"))&"',BillMem4='"&trim(request("BillMemName4"))&"'" &_
		",BillFillerMemberID='"&trim(request("BillMemID1"))&"',BillFiller='"&trim(request("BillMemName1"))&"'" &_
		",BillFillDate="&theBillFillDate&",DealLineDate="&theDealLineDate&_
		",Note='"&trim(request("Note"))&"',IsLecture='"&trim(request("IsLecture"))&"'" &_
		",DriverSex='"&trim(request("DriverSex"))&"',SignType='"&UCase(trim(request("SignType")))&"'"&UpdateAdd &_
		",DoubleCheckStatus="&trim(request("Sys_DoubleCheckStatus"))&" where SN="&trim(request("BillSN"))

		conn.execute strUpd
		ConnExecute strUpd,353
	'行人攤販行沒入物品 PasserConfiscate
	strDel="delete from PasserConfiscate where BillSN="&trim(request("BillSN"))
	conn.execute strDel
	if trim(request("Fastener1"))<>"" then
		Ftemp=split(trim(request("Fastener1")),",")
		if ubound(Ftemp)>=0 then
			Fvaluetemp=split(Ftemp(0),"_")
			FID=trim(Fvaluetemp(0))
			Fvalue=trim(Fvaluetemp(1))
			strInsFastene1="insert into PasserConfiscate(BillSN,BillNo,Confiscate,ConfiscateID)" &_
					" values("&trim(request("BillSN"))&",'"&UCase(trim(request("Billno1")))&"','"&Fvalue&"','"&FID&"')"
			conn.execute strInsFastene1
			ConnExecute strInsFastene1,353
		end if
		if ubound(Ftemp)>=1 then
			Fvaluetemp=split(Ftemp(1),"_")
			FID=trim(Fvaluetemp(0))
			Fvalue=trim(Fvaluetemp(1))
			strInsFastene2="insert into PasserConfiscate(BillSN,BillNo,Confiscate,ConfiscateID)" &_
						" values("&trim(request("BillSN"))&",'"&UCase(trim(request("Billno1")))&"','"&Fvalue&"','"&FID&"')"
			conn.execute strInsFastene2
			ConnExecute strInsFastene2,353
		end if
		if ubound(Ftemp)>=2 then
			Fvaluetemp=split(Ftemp(2),"_")
			FID=trim(Fvaluetemp(0))
			Fvalue=trim(Fvaluetemp(1))
			strInsFastene3="insert into PasserConfiscate(BillSN,BillNo,Confiscate,ConfiscateID)" &_
						" values("&trim(request("BillSN"))&",'"&UCase(trim(request("Billno1")))&"','"&Fvalue&"','"&FID&"')"
			conn.execute strInsFastene3
			ConnExecute strInsFastene3,353
		end if
	end if
%>
<script language="JavaScript">
	alert("修改完成");
</script>
<%
end if


strSql="select * from PasserBase where SN="&trim(request("BillSN"))
set rs1=conn.execute(strSql)

%>

<style type="text/css">
<!--
.style1 {font-size: 14px}
.style3 {font-size: 15px}
.style4 {font-size: 13px}
.style5 {	color: #FF0000;
	font-size: 12px
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
	<form name="myForm" method="post">  
		<table width='993' border='1' align="left" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="6"><strong>裁罰資料修改作業</strong>&nbsp; &nbsp; 日期格式：951220 &nbsp;時間格式：2300(24小時制)</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" align="right"><span class="style5">＊</span>單號</td>
				<td><input name="Billno1" type="text" value="<%
			if trim(rs1("Billno"))<>"" and not isnull(rs1("Billno")) then
				response.write trim(rs1("Billno"))
			end if
				%>" onkeydown="funTextControl(this);" size="10" maxlength="9"></td>
				<td bgcolor="#EBE5FF" align="right">違規人姓名</td>
				<td><input type="text" size="10" value="<%
			if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
				response.write trim(rs1("Driver"))
			end if
				%>" onkeydown="funTextControl(this);" name="DriverName">
				</td>
				<td bgcolor="#EBE5FF" align="right">違規人出生日期</td>
				<td><input type="text" size="10" maxlength="7" value="<%
			if trim(rs1("DriverBirth"))<>"" and not isnull(rs1("DriverBirth")) then
				response.write ginitdt(trim(rs1("DriverBirth")))
			end if
				%>" name="DriverBrith" onkeydown="funTextControl(this);" onkeyup="focusToDriverPID()">
				</td>

			</tr>
			<tr>
				<td bgcolor="#EBE5FF" align="right"><span class="style5">＊</span>違規人身份證號</td>
				<td><input type="text" size="10" value="<%
			if trim(rs1("Driverid"))<>"" and not isnull(rs1("Driverid")) then
				response.write trim(rs1("Driverid"))
			end if
				%>" name="DriverPID" onkeydown="funTextControl(this);" onkeyup="value=value.toUpperCase()" onBlur="FuncChkPID();">
				</td>
				<td bgcolor="#EBE5FF" align="right">違規人地址</td>
				<td colspan="3">
				
				<input type="text" class="btn5" size="3" value="<%=trim(rs1("DriverZip"))%>" name="DriverZip"  onBlur="getDriverZip(this,'DriverAddress');" onkeydown="funTextControl(this);">
				區號
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick="QryDriverZip();">

				<input type="text" size="40" value="<%
				if trim(rs1("DriverAddress"))<>"" and not isnull(rs1("DriverAddress")) then
					response.write trim(rs1("DriverAddress"))
				end if
				%>" name="DriverAddress" onkeydown="funTextControl(this);">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" align="right"><span class="style5">＊</span>違規日期</td>
				<td>
				<input type="text" size="10" value="<%
				if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
					response.write ginitdt(trim(rs1("IllegalDate")))
				end if
				%>" maxlength="7" name="IllegalDate" onkeydown="funTextControl(this);" onkeyup="getDealLineDate();">
				</td>
				<td bgcolor="#EBE5FF" align="right"><span class="style5">＊</span>違規時間</td>
				<td colspan="3">
				<input type="text" size="10" value="<%
				if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
					if hour(rs1("IllegalDate"))>9 then
						theChangeTime=theChangeTime&hour(rs1("IllegalDate"))
					else
						theChangeTime=theChangeTime&"0"&hour(rs1("IllegalDate"))
					end if
					if minute(rs1("IllegalDate"))>9 then
						theChangeTime=theChangeTime&minute(rs1("IllegalDate"))
					else
						theChangeTime=theChangeTime&"0"&minute(rs1("IllegalDate"))
					end if
					response.write theChangeTime
				end if
				%>" maxlength="4" name="IllegalTime" onkeydown="funTextControl(this);" onkeyup="value=value.replace(/[^\d]/g,'')">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" align="right">違規地點代碼</td>
				<td >
					
					<input type="text" size="8" value="<%
				if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
					response.write trim(rs1("IllegalAddressID"))
				end if
				%>" name="IllegalAddressID" onkeyup="getillStreet();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage_Street_People","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#EBE5FF" align="right"><span class="style5">＊</span>違規地點</td>
				<td colspan="3">
					<%if sys_City="高雄市" then %>
						區號
						<input type="text" class="btn5" size="3" value="<%
				if trim(rs1("IllegalZip"))<>"" and not isnull(rs1("IllegalZip")) then
					bIllZip=trim(rs1("IllegalZip"))
					response.write trim(rs1("IllegalZip"))
				else
					bIllZip=""
				end if 
						%>" name="IllegalZip" onKeyUp="getIllZip();" onkeydown="funTextControl(this);" maxlength="3">
						<Input type="hidden" name="OldIllegalZip" value="<%=Trim(rs1("IllegalZip"))%>">
						
						<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick="QryIllegalZip();">
						<div id="LayerIllZip" style="position:absolute ; width:160px; height:30; z-index:0;  border: 1px none #000000;""><%
					if trim(bIllZip)<>"" then
						strZip1="select ZipName from Zip where ZipNo='"&trim(bIllZip)&"'"
						set rsZip1=conn.execute(strZip1)
						if not rsZip1.eof then
							response.write trim(rsZip1("ZipName"))
						end if
						rsZip1.close
						set rsZip1=nothing
					end if
					%></div><br>
					<%end if%>
					<input type="text" size="40" value="<%
				if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
					response.write trim(rs1("IllegalAddress"))
				end if
				%>" name="IllegalAddress" onkeydown="funTextControl(this);">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" align="right"><span class="style5">＊</span>違規法條一</td>
				<td colspan="5">
					<input type="text" size="10" value="<%
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					response.write trim(rs1("Rule1"))
				end if
				%>" name="Rule1" onKeyUp="getRuleData1();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=trim(rs1("RuleVer"))%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer1" style="position:absolute ; width:570px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"'"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rsR1("IllegalRule"))
					end if
					rsR1.close
					set rsR1=nothing
				end if
				%></div>
					<input type="hidden" name="ForFeit1" value="<%
				if trim(rs1("ForFeit1"))<>"" and not isnull(rs1("ForFeit1")) then
					response.write trim(rs1("ForFeit1"))
				end if
				%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" align="right">違規法條二</td>
				<td colspan="5">
					<input type="text" size="10" value="<%
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					response.write trim(rs1("Rule2"))
				end if
				%>" name="Rule2" onKeyUp="getRuleData2();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=trim(rs1("RuleVer"))%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer2" style="position:absolute ; width:570px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"'"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rsR1("IllegalRule"))
					end if
					rsR1.close
					set rsR1=nothing
				end if
				%></div>
					<input type="hidden" name="ForFeit2" value="<%
				if trim(rs1("ForFeit2"))<>"" and not isnull(rs1("ForFeit2")) then
					response.write trim(rs1("ForFeit2"))
				end if
				%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" align="right"><span class="style5">＊</span>應到案日期</td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
					response.write ginitdt(trim(rs1("DealLineDate")))
				end if
				%>" maxlength="7" name="DealLineDate" onkeydown="funTextControl(this);" onkeyup="value=value.replace(/[^\d]/g,'')">
				</td>
				<td bgcolor="#EBE5FF" align="right" nowrap><span class="style5">＊</span>應到案處所代碼</td>
				<td>
					<input type="text" size="4" value="<%
				if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
					response.write trim(rs1("MemberStation"))
				end if
				%>" name="MemberStation" onKeyup="getStation();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=S","WebPage1","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style4">
					<div id="Layer5" style="position:absolute ; width:250px; height:30px; z-index:0;  border: 1px none #000000;"><%
				if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
					strS="select UnitName from UnitInfo where UnitID='"&trim(rs1("MemberStation"))&"'"
					set rsS=conn.execute(strS)
					if not rsS.eof then
						response.write trim(rsS("UnitName"))
					end if
					rsS.close
					set rsS=nothing
				end if
				%></div>
				</span>
				</td>
				<td bgcolor="#EBE5FF" align="right"><span class="style5">＊</span>舉發人
						<% if sys_City<>"高雄縣" or sys_City<>"高雄市" then 
									response.write "姓名"
							 else
									response.write "代碼"
							 end if
							%>1</td>
		  		<td>
					<input type="text" size="4" value="<%
				if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
					strMem1="select ChName from MemberData where MemberID="&trim(rs1("BillMemID1"))
					set rsMem1=conn.execute(strMem1)
					if not rsMem1.eof then
						response.write trim(rsMem1("ChName"))
					end if
					rsMem1.close
					set rsMem1=nothing
				end if
				%>" name="BillMem1" onblur="getBillMemID1();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemType=P&MemOrder=1","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer12" style="position:absolute ; width:92px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
					strMem1="select LoginID,ChName from MemberData where MemberID="&trim(rs1("BillMemID1"))
					set rsMem1=conn.execute(strMem1)
					if not rsMem1.eof then
						response.write trim(rsMem1("LoginID"))
					end if
					rsMem1.close
					set rsMem1=nothing
				end if
				%></div>
					<input type="hidden" value="<%
				if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
					response.write trim(rs1("BillMemID1"))
				end if
				%>" name="BillMemID1">
					<input type="hidden" value="<%
				if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
					response.write trim(rs1("BillMem1"))
				end if
				%>" name="BillMemName1">
				</td>
			</tr>
			<tr>
				
				<td bgcolor="#EBE5FF" align="right" width="14%">舉發人
						<% if sys_City<>"高雄縣" or sys_City<>"高雄市" then 
									response.write "姓名"
							 else
									response.write "代碼"
							 end if
							%>2</td>
				<td width="20%">
					<input type="text" size="4" value="<%
				if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
					strMem2="select Chname from MemberData where MemberID="&trim(rs1("BillMemID2"))
					set rsMem2=conn.execute(strMem2)
					if not rsMem2.eof then
						response.write trim(rsMem2("Chname"))
					end if
					rsMem2.close
					set rsMem2=nothing
				end if
				%>" name="BillMem2" onblur="getBillMemID2();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemType=P&MemOrder=2","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer13" style="position:absolute ; width:92px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
					response.write trim(rs1("BillMemID2"))
				end if
				%></div>
					<input type="hidden" value="<%
				if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
					response.write trim(rs1("BillMemID2"))
				end if
				%>" name="BillMemID2">
					<input type="hidden" value="<%
				if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
					response.write trim(rs1("BillMem2"))
				end if
				%>" name="BillMemName2">
				</td>
				<td bgcolor="#EBE5FF" align="right" width="13%">舉發人
						<% if sys_City<>"高雄縣" or sys_City<>"高雄市" then 
									response.write "姓名"
							 else
									response.write "代碼"
							 end if
							%>3</td>
				<td width="20%">
					<input type="text" size="4" value="<%
				if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
					strMem3="select Chname from MemberData where MemberID="&trim(rs1("BillMemID3"))
					set rsMem3=conn.execute(strMem3)
					if not rsMem3.eof then
						response.write trim(rsMem3("Chname"))
					end if
					rsMem3.close
					set rsMem3=nothing
				end if
				%>" name="BillMem3" onblur="getBillMemID3();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemType=P&MemOrder=3","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer14" style="position:absolute ; width:92px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
					response.write trim(rs1("BillMemID3"))
				end if
				%></div>
					<input type="hidden" value="<%
				if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
					response.write trim(rs1("BillMemID3"))
				end if
				%>" name="BillMemID3">
					<input type="hidden" value="<%
				if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
					response.write trim(rs1("BillMem3"))
				end if
				%>" name="BillMemName3">
				</td>
				<td bgcolor="#EBE5FF" align="right" width="13%">舉發人
						<% if sys_City<>"高雄縣" or sys_City<>"高雄市" then 
									response.write "姓名"
							 else
									response.write "代碼"
							 end if
							%>4</td>
				<td width="20%">
					<input type="text" size="4" value="<%
				if trim(rs1("BillMemID4"))<>"" and not isnull(rs1("BillMemID4")) then
					strMem4="select Chname from MemberData where MemberID="&trim(rs1("BillMemID4"))
					set rsMem4=conn.execute(strMem4)
					if not rsMem4.eof then
						response.write trim(rsMem4("Chname"))
					end if
					rsMem4.close
					set rsMem4=nothing
				end if
				%>" name="BillMem4" onblur="getBillMemID4();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemType=P&MemOrder=4","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer17" style="position:absolute ; width:92px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
					response.write trim(rs1("BillMem4"))
				end if
				%></div>
					<input type="hidden" value="<%
				if trim(rs1("BillMemID4"))<>"" and not isnull(rs1("BillMemID4")) then
					response.write trim(rs1("BillMemID4"))
				end if
				%>" name="BillMemID4">
					<input type="hidden" value="<%
				if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
					response.write trim(rs1("BillMem4"))
				end if
				%>" name="BillMemName4">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" align="right" height="33">代保管物</td>
				<td nowrap>
<%	strConfiscate=""
	strFas="select ConfiscateID from PasserConfiscate where BillSN="&trim(request("BillSN"))
	set rsFas=conn.execute(strFas)
	If Not rsFas.Bof Then rsFas.MoveFirst 
	While Not rsFas.Eof
		if strConfiscate="" then
			strConfiscate=trim(rsFas("ConfiscateID"))
		else
			strConfiscate=strConfiscate&","&trim(rsFas("ConfiscateID"))
		end if
	rsFas.MoveNext
	Wend
	rsFas.close
	set rsFas=nothing

	strItem="select ID,Content from Code where TypeID=2 and Not(ID<478 or ID=479) order by ID"
	set rsItem=conn.execute(strItem)
	If Not rsItem.Bof Then rsItem.MoveFirst 
	While Not rsItem.Eof
%>
					<input type="checkbox" value="<%=trim(rsItem("ID"))&"_"&trim(rsItem("Content"))%>" name="Fastener1" <%
					if strConfiscate<>"" then
						if instr(strConfiscate,trim(rsItem("ID")))>0 then
							response.write "checked"
						end if
					end if
					%>><%=trim(rsItem("Content"))%>&nbsp;
<%	
	rsItem.MoveNext
	Wend
	rsItem.close
	set rsItem=nothing

%>
				</td>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>簽收狀況</div></td>
				<td colspan="3">
					<input type="text" size="5" value="A" maxlength="1" name="SignType" onBlur="funcSignType();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<font color="#ff000" size="2">
					A簽收/ U拒簽收/ 2拒簽已收/ 3已簽拒收/ 5補開單 
					</font>
				</td>
			</tr>				
			<tr height="6"><td></td></tr>
			<tr>
				<td bgcolor="#EBE5FF" align="right"><span class="style5">＊</span>舉發單位代號</td>
				<td>
					<input type="text" size="4" value="<%
				if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
					response.write trim(rs1("BillUnitID"))
				end if
				%>" name="BillUnitID" onKeyUp="getUnit();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=U","WebPage_Unit_People","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style4">
					<div id="Layer6" style="position:absolute ; width:250px; height:30px; z-index:0;  border: 1px none #000000;"><%
				if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
					strU="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
					set rsU=conn.execute(strU)
					if not rsU.eof then
						response.write trim(rsU("UnitName"))
					end if
					rsU.close
					set rsU=nothing
				end if
					%></div>
					</span>
				</td>
				<td bgcolor="#EBE5FF"><div align="right">專案代碼</div></td>
				<td>
					<input type="text" size="5" value="<%=trim(rs1("ProjectID"))%>" name="ProjectID" onkeyup="ProjectF5()" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onClick='window.open("Query_Project.asp","WebPage_project","left=0,top=0,location=0,width=700,height=455,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer001" style="position:absolute ; width:180px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"><%
						if not ifnull(rs1("ProjectID")) then

							strProject="select Name from Project where ProjectID='"&trim(rs1("ProjectID"))&"'"

							set rsProject=conn.execute(strProject)
							if not rsProject.eof then
								response.write trim(rsProject("Name"))
							end if
							rsProject.close
						end if
					%></div>
					</span>
				</td>
				<!--<td bgcolor="#EBE5FF" align="right">是否講習</td>
				<td>
					<input type="radio" value="1" name="IsLecture" <%
			if trim(rs1("IsLecture"))<>"" and not isnull(rs1("IsLecture")) then
				if trim(rs1("IsLecture"))="1" then
					response.write "checked"
				end if
			end if
				%>>是
					<input type="radio" value="0" name="IsLecture" <%
			if trim(rs1("IsLecture"))<>"" and not isnull(rs1("IsLecture")) then
				if trim(rs1("IsLecture"))="0" then
					response.write "checked"
				end if
			end if
				%>>否
				</td>
				<td bgcolor="#EBE5FF" align="right">告發類別</td>
				<td>
				<input type="text" size="3" maxlength="1" value="<%
			if trim(rs1("BillTypeID"))<>"" and not isnull(rs1("BillTypeID")) then
				response.write trim(rs1("BillTypeID"))
			end if
				%>" name="BillType" onkeyup="value=value.replace(/[^\d]/g,'')">
				<font color="#ff000" size="2">1慢車/2行人/3道路障礙</font>
				</td>-->
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" align="right"> <span class="style5">＊</span>填單日期</td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
					response.write ginitdt(trim(rs1("BillFillDate")))
				end if
				%>" maxlength="7" name="BillFillDate" onkeydown="funTextControl(this);" onkeyup="value=value.replace(/[^\d]/g,'')">
				</td>
				<td bgcolor="#EBE5FF" align="right">備註</td>
				<td colspan="3">
					<input type="text" size="46" value="<%
				if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
					response.write trim(rs1("Note"))
				end if
				%>" name="Note" onkeydown="funTextControl(this);">
				</td>
			</tr>
			<tr>
			  <td bgcolor="#1BF5FF" align="center" colspan="6">
					<font color="red"><B>建檔序號第<span id="Seqfile"><input type="text" value="<%=trim(rs1("DoubleCheckStatus"))%>" class="btn1" size="10" name="Sys_DoubleCheckStatus" onkeyup="value=value.replace(/[^\d]/g,'')"></span>號</B></font>
					&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="button" value="儲 存 F2" onclick="InsertBillVase();" <%
				'1:查詢 ,2:新增 ,3:修改 ,4:刪除
				if trim(rs1("RecordMemberID"))<>trim(session("User_ID")) then
					if CheckPermission(235,3)=false and CheckPermission(224,3)=false then
						response.write "disabled"
					end if
				end if
					%> class="btn1">
					<input type="hidden" value="<%=trim(rs1("RuleVer"))%>" name="RuleVerSion">
					<input type="hidden" value="" name="kinds">
					<input type="hidden" value="<%=trim(request("BillSN"))%>" name="BillSN">
                    <span class="style1"><span class="style3"><img src="/image/space.gif" width="29" height="8"></span>
                    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉" class="btn1">
                </span>
				<!-- 違規人性別 -->
				<input type="hidden" value="<%=trim(rs1("DriverSex"))%>" name="DriverSex">
				<input type="hidden" value="" name="Mem">
				<input type="hidden" value="" name="MemOrder">
				<input type="hidden" value="" name="MemType">
			  </td>
			</tr>
		</table>
	</form>
<%
rs1.close
set rs1=nothing
conn.close
set conn=nothing
%>
</body>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var TDLawNum=0;
var TDLawErrorLog1=0;
var TDLawErrorLog2=0;
var TDLawErrorLog3=0;
var TDLawErrorLog4=0;
var TDStationErrorLog=0;
var TDUnitErrorLog=0;
var TDFastenerErrorLog1=0;
var TDFastenerErrorLog2=0;
var TDFastenerErrorLog3=0;
var TDMemErrorLog1=0;
var TDMemErrorLog2=0;
var TDMemErrorLog3=0;
var TDMemErrorLog4=0;
var TDIllZipErrorLog=0;
var sys_City="<%=sys_City%>";
<%If Trim(sys_City)="台南市" then%>
MoveTextVar("Billno1,DriverName,DriverBrith||DriverPID,DriverZip,DriverAddress||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||SignType||BillUnitID||BillFillDate,Note");
<%elseif Trim(sys_City)="高雄市" then %>
MoveTextVar("Billno1,DriverName,DriverBrith||DriverPID,DriverAddress||IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||Rule2||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||SignType||BillUnitID||BillFillDate,Note");

<%else%>
MoveTextVar("Billno1,DriverName,DriverBrith||DriverPID,DriverZip,DriverAddress||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||SignType||BillUnitID||BillFillDate,Note");
<%end if%>
//新增告發單
function InsertBillVase(){
	var error=0;
	var errorString="";
	var TodayDate=<%=ginitdt(date)%>;
	if (myForm.Billno1.value==""){
		error=error+1;
		errorString=error+"：請輸入單號。";
	}else{
	   if (myForm.Billno1.value != ""){
		  chkResult = chkBillNumber(myForm.Billno1,"[舉發單起始碼] 格式錯誤!!"); 
	     if (chkResult != "Y"){
			  error=error+1;
			  errorString=error+"：舉發單號格式錯誤。";
		 }
	   }
	}
	/*
	if (myForm.BillType.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入告發類別。";
	}
	*/
	if(myForm.DriverBrith.value!=""){
		if(!dateCheck( myForm.DriverBrith.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規人出生日期輸入錯誤。";	
		}
	}
	if(myForm.DriverPID.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入違規人身份證號碼。";	
	}
	if (myForm.IllegalDate.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規日期。";
	}else if(!dateCheck( myForm.IllegalDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
	}/*else if (!ChkIllegalDate(myForm.IllegalDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過三個月期限。";
	}*/
	if (myForm.IllegalTime.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規時間。";
	}else if(myForm.IllegalTime.value.length < 4){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規時間輸入錯誤。";
	}else if(myForm.IllegalTime.value.substr(0,2) > 23 || myForm.IllegalTime.value.substr(0,2) < 0){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規時間輸入錯誤。";
	}else if(myForm.IllegalTime.value.substr(2,2) > 59 || myForm.IllegalTime.value.substr(2,2) < 0){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規時間輸入錯誤。";
	}
<%if sys_City="高雄市" then%>
	if (TDIllZipErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規地點區號 輸入錯誤。";
	}
	//else if(myForm.IllegalZip.value==""){
	//	error=error+1;
	//	errorString=errorString+"\n"+error+"：請輸入違規地點區號。";
	//}
<%end if%>
	if (myForm.IllegalAddress.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規地點。";
	}
	if (myForm.Rule1.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規法條一。";
	}else if (TDLawErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條一輸入錯誤。";
	}
	if (myForm.Rule1.value==myForm.Rule2.value && myForm.Rule1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條一與違規法條二重複。";
	}

	if (TDLawErrorLog2==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條二輸入錯誤。";
	}

	if (myForm.BillFillDate.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入填單日期。";
	}else if (!dateCheck( myForm.BillFillDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
	}else if(TodayDate < myForm.BillFillDate.value){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期不得比今天晚。";
	}

	if (Layer5.innerHTML==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入應到案處所。";
	}else if (TDStationErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案處所輸入錯誤。";
	}
	if (myForm.DealLineDate.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入應到案日期。";
	}else if (!dateCheck( myForm.DealLineDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}
	if (myForm.SignType.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入簽收狀況。";
	}
	if (myForm.BillUnitID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入舉發單位代號。";
	}else if (TDUnitErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發單位代號輸入錯誤。";
	}
	if (myForm.BillMem1.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入舉發人臂章號碼。";
	}else if (TDMemErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人臂章號碼1 輸入錯誤。";
	}
	if (TDMemErrorLog2==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人臂章號碼2 輸入錯誤。";
	}
	if (TDMemErrorLog3==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人臂章號碼3 輸入錯誤。";
	}
	if (TDMemErrorLog4==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人臂章號碼4 輸入錯誤。";
	}
	if (myForm.BillMem1.value==myForm.BillMem2.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人臂章號碼1 與 舉發人臂章號碼2 重複。";
	}else if (myForm.BillMem1.value==myForm.BillMem3.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人臂章號碼1 與 舉發人臂章號碼3 重複。";
	}else if (myForm.BillMem1.value==myForm.BillMem4.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人臂章號碼1 與 舉發人臂章號碼4 重複。";
	}
	if (myForm.BillMem2.value==myForm.BillMem3.value && myForm.BillMem2.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人臂章號碼2 與 舉發人臂章號碼3 重複。";
	}else if (myForm.BillMem2.value==myForm.BillMem4.value && myForm.BillMem2.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人臂章號碼2 與 舉發人臂章號碼4 重複。";
	}
	if (myForm.BillMem3.value==myForm.BillMem4.value && myForm.BillMem3.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人臂章號碼3 與 舉發人臂章號碼4 重複。";
	}
	if (myForm.BillFillDate.value < myForm.IllegalDate.value){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比填單日晚。";
	}else if(TodayDate < myForm.IllegalDate.value){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比今天晚。";
	}
	if (error==0){
		myForm.kinds.value="DB_insert";
		myForm.submit();
	}else{
		alert(errorString);
	}
}


//違規事實1(ajax)
function getRuleData1(){
	if (myForm.Rule1.value.length > 6){
		var Rule1Num=myForm.Rule1.value;
		var VerNo=myForm.RuleVerSion.value;
		runServerScript("getRuleDetail.asp?RuleOrder=1&RuleID="+Rule1Num+"&RuleVer="+VerNo);
	}else if (myForm.Rule1.value.length <= 6 && myForm.Rule1.value.length > 0){
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=1;
	}else{
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=0;
	}
}
//違規事實2(ajax)
function getRuleData2(){
	if (myForm.Rule2.value.length > 6){
		var Rule2Num=myForm.Rule2.value;
		var VerNo=myForm.RuleVerSion.value;
		runServerScript("getRuleDetail.asp?RuleOrder=2&RuleID="+Rule2Num+"&RuleVer="+VerNo);
	}else if (myForm.Rule2.value.length <= 6 && myForm.Rule2.value.length > 0){
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=1;
	}else{
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=0;
	}
}

//到案處所(ajax)
function getStation(){
	if (myForm.MemberStation.value.length > 1){
		var StationNum=myForm.MemberStation.value;
		runServerScript("getMemberStation2.asp?StationID="+StationNum);
	}else{
		Layer5.innerHTML=" ";
		TDStationErrorLog=1;
	}
}
//舉發單位(ajax)
function getUnit(){
	if (myForm.BillUnitID.value.length > 1){
		var BillUnitNum=myForm.BillUnitID.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum);
	}else{
		Layer6.innerHTML=" ";
		TDUnitErrorLog=1;
	}
}
//違規地點代碼(ajax)
function getillStreet(){
	if (event.keyCode==<%=F5str%>){	
		event.keyCode=0;
		event.returnValue=false;
		Ostreet=myForm.IllegalAddressID.value;
		window.open("Query_Street.asp?OStreetID="+Ostreet,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.IllegalAddressID.value.length > 4){
		var illAddrNum=myForm.IllegalAddressID.value;
		runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
	}
}
//舉發人一(ajax)
function getBillMemID1(){
	if (event.keyCode==<%=F5str%>){	
		event.keyCode=0;
		myForm.Mem.value=myForm.BillMem1.value;
		myForm.MemType.value='P';
		myForm.MemOrder.value=1;
		UrlStr="Query_MemID.asp";		
		myForm.action=UrlStr;
		myForm.target="WebPage_Street_People";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		if (myForm.BillMem1.value.length > 1){
			var BillMemNum=myForm.BillMem1.value;
			runServerScript("getBillPeoPleMemID.asp?MType=People&MemOrder=1&MemID="+BillMemNum);
		}else if (myForm.BillMem1.value.length <= 1 && myForm.BillMem1.value.length > 0){
			Layer12.innerHTML=" ";
			myForm.BillMemID1.value="";
			myForm.BillMemName1.value="";
			TDMemErrorLog1=1;
		}else{
			Layer12.innerHTML=" ";
			myForm.BillMemID1.value="";
			myForm.BillMemName1.value="";
			TDMemErrorLog1=0;
		}
	}
}
//舉發人二(ajax)
function getBillMemID2(){
	if (event.keyCode==<%=F5str%>){	
		event.keyCode=0;
		myForm.Mem.value=myForm.BillMem1.value;
		myForm.MemType.value='P';
		myForm.MemOrder.value=2;
		UrlStr="Query_MemID.asp";		
		myForm.action=UrlStr;
		myForm.target="WebPage_Street_People";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		if (myForm.BillMem2.value.length > 1){
			var BillMemNum=myForm.BillMem2.value;
			runServerScript("getBillPeoPleMemID.asp?MType=People&MemOrder=2&MemID="+BillMemNum);
		}else if (myForm.BillMem2.value.length <= 1 && myForm.BillMem2.value.length > 0){
			Layer13.innerHTML=" ";
			myForm.BillMemID2.value="";
			myForm.BillMemName2.value="";
			TDMemErrorLog2=1;
		}else{
			Layer13.innerHTML=" ";
			myForm.BillMemID2.value="";
			myForm.BillMemName2.value="";
			TDMemErrorLog2=0;
		}
	}
}
//舉發人三(ajax)
function getBillMemID3(){
	if (event.keyCode==<%=F5str%>){	
		event.keyCode=0;
		myForm.Mem.value=myForm.BillMem1.value;
		myForm.MemType.value='P';
		myForm.MemOrder.value=3;
		UrlStr="Query_MemID.asp";		
		myForm.action=UrlStr;
		myForm.target="WebPage_Street_People";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		if (myForm.BillMem3.value.length > 1){
			var BillMemNum=myForm.BillMem3.value;
			runServerScript("getBillPeoPleMemID.asp?MType=People&MemOrder=3&MemID="+BillMemNum);
		}else if (myForm.BillMem3.value.length <= 1 && myForm.BillMem3.value.length > 0){
			Layer14.innerHTML=" ";
			myForm.BillMemID3.value="";
			myForm.BillMemName3.value="";
			TDMemErrorLog3=1;
		}else{
			Layer14.innerHTML=" ";
			myForm.BillMemID3.value="";
			myForm.BillMemName3.value="";
			TDMemErrorLog3=0;
		}
	}
}
//舉發人四(ajax)
function getBillMemID4(){
	if (event.keyCode==<%=F5str%>){	
		event.keyCode=0;
		myForm.Mem.value=myForm.BillMem1.value;
		myForm.MemType.value='P';
		myForm.MemOrder.value=4;
		UrlStr="Query_MemID.asp";		
		myForm.action=UrlStr;
		myForm.target="WebPage_Street_People";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		if (myForm.BillMem4.value.length > 1){
			var BillMemNum=myForm.BillMem4.value;
			runServerScript("getBillPeoPleMemID.asp?MType=People&MemOrder=4&MemID="+BillMemNum);
		}else if (myForm.BillMem4.value.length <= 1 && myForm.BillMem4.value.length > 0){
			Layer17.innerHTML=" ";
			myForm.BillMemID4.value="";
			myForm.BillMemName4.value="";
			TDMemErrorLog4=1;
		}else{
			Layer17.innerHTML=" ";
			myForm.BillMemID4.value="";
			myForm.BillMemName4.value="";
			TDMemErrorLog4=0;
		}
	}
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
	return win;
}
//簽收狀況(小轉大寫，限定A or U)
function funcSignType(){
	if (myForm.SignType.value=="a" || myForm.SignType.value=="u"){
		myForm.SignType.value=myForm.SignType.value.toUpperCase();
	}
	if (myForm.SignType.value==""){
		myForm.SignType.focus();
		alert("簽收狀況未填寫!!");
	}
}
//由違規日期帶入應到案日期
function getDealLineDate(){
	if (!ChkIllegalDate(myForm.IllegalDate.value)){
		alert("違規日期已超過三個月期限，請確認是否正確。");
	}
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');
	BFillDateTemp=myForm.IllegalDate.value;
	if (BFillDateTemp.length >= 6){
		Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
		Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
		Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
		var BFillDate=new Date(Byear,Bmonth-1,Bday)
		var DLineDate=new Date()
		DLineDate=DateAdd("d",30,BFillDate);
		Dyear=parseInt(DLineDate.getYear())-1911;
		Dmonth=DLineDate.getMonth()+1;
		Dday=DLineDate.getDate();
		Dyear=Dyear.toString();
		if (Dmonth < 10){
			Dmonth="0"+Dmonth;
		}
		if (Dday < 10){
			Dday="0"+Dday;
		}
		myForm.DealLineDate.value=Dyear+Dmonth+Dday;
	}
}
//檢查單號是否有在GETBILLBASE內
function CheckPeopleBillNoExist(){
	myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
	var BillNum=myForm.Billno1.value;
	if (myForm.Billno1.value.length >= 9){
		runServerScript("getPeopleBillNoExist.asp?BillNo="+BillNum);
	}
}
function setCheckPeopleBillNoExist(GetBillFlag,PasserBaseFlag,BillSN,MLoginID,MMemberID,MMemName,MUnitID,MUnitName,SUnitID,SUnitName){
	if (GetBillFlag==0){
		//alert("此單號不存在於領單紀錄中！");
		//document.myForm.Billno1.value="";
	}else{
		document.myForm.BillMem1.value=MLoginID;
		document.myForm.BillMemID1.value=MMemberID;
		document.myForm.BillMemName1.value=MMemName;
		Layer12.innerHTML=MMemName;
		TDMemErrorLog1=0;
		//if (document.myForm.BillUnitID.value==""){
			document.myForm.BillUnitID.value=MUnitID;
			Layer6.innerHTML=MUnitName;
			TDUnitErrorLog=0;
		//}
		//if (document.myForm.MemberStation.value==""){
			document.myForm.MemberStation.value=SUnitID;
			Layer5.innerHTML=SUnitName;
			TDStationErrorLog=0;
		//}
	}
	if (PasserBaseFlag==1){
		alert("此單號已建檔！");
		document.myForm.Billno1.value="";
	}else if (PasserBaseFlag==0){
		if(confirm('此單號已建檔，是否要修改此筆舉發單？')){
		window.open("BillKeyIn_People_Update.asp?BillSN="+BillSN,"Page_Upd_PassBill1","left=0,top=0,location=0,width=750,height=550,resizable=yes,scrollbars=yes")
			document.myForm.Billno1.value="";
		}else{
			document.myForm.Billno1.value="";
		}
	}
}
function CallChkLaw1(){
}
function CallChkLaw2(){
}
function FuncChkPID(){
	if (myForm.DriverPID.value.length == 10){
		if (!check_tw_id(myForm.DriverPID.value)){
			alert("身分證輸入錯誤！");
			//myForm.DriverPID.focus();
			if (myForm.DriverPID.value.substr(1,1)=="1"){
				document.myForm.DriverSex.value="1";
			}else{
				document.myForm.DriverSex.value="2";
			}
		}else{
			if (myForm.DriverPID.value.substr(1,1)=="1"){
				document.myForm.DriverSex.value="1";
			}else{
				document.myForm.DriverSex.value="2";
			}
		}
	}else if (myForm.DriverPID.value.length > 0 && myForm.DriverPID.value.length < 10){
		alert("身分證輸入錯誤！");
		//myForm.DriverPID.focus();
		if (myForm.DriverPID.value.substr(1,1)=="1"){
			document.myForm.DriverSex.value="1";
		}else{
			document.myForm.DriverSex.value="2";
		}
	}
}
function focusToDriverPID(){
	myForm.DriverBrith.value=myForm.DriverBrith.value.replace(/[^\d]/g,'');
	if (myForm.DriverBrith.value.length==6){
		var x=new Date();
		var thisYear=x.getYear()-1911;
		BFillDateTmp=myForm.DriverBrith.value;
		BirthYear=parseInt(BFillDateTmp.substr(0,2));
		if ((thisYear-BirthYear) < 10){
			alert("違規人年齡低於十歲!!");
		}
	}
}
function funTextControl(obj){
	if (event.keyCode==13){ //Enter換欄
		event.keyCode=0;
		event.returnValue=false;
		CodeEnter(obj.name);
	/*}else if (event.keyCode==37){ //左換欄
		event.keyCode=0;
		event.returnValue=false;
		CodeMoveLeft(obj.name);*/
	}else if (event.keyCode==38){ //上換欄
		event.keyCode=0;
		event.returnValue=false;
		//CodeMoveUp(obj.name);
		CodeMoveLeft(obj.name);
	/*}else if (event.keyCode==39){ //右換欄
		event.keyCode=0;
		event.returnValue=false;
		CodeMoveRight(obj.name);*/
	}else if (event.keyCode==40){ //下換欄
		event.keyCode=0;
		event.returnValue=false;
		//CodeMoveDown(obj.name);
		CodeMoveRight(obj.name);
	}
}
function KeyDown(){ 
	if (event.keyCode==<%=F5str%>){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
		InsertBillVase();
	}
}
function getDriverZip(obj,objName){
	if(obj.value!=''&&obj.value.length>2){
		runServerScript("getZipName.asp?ZipID="+obj.value+"&getZipName="+objName);
	}else if(obj.value!=''&&obj.value.length<3){
		alert("郵遞區號錯誤!!");
	}
}
function QryDriverZip(){
	window.open("Query_Zip.asp?ZipCity=&IllegalZip="+myForm.DriverZip.value+"&ObjName=DriverZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");

}
<%if sys_City="高雄市" then%>
function QryIllegalZip(){
	window.open("Query_Zip.asp?ZipCity="+sys_City+"&IllegalZip="+myForm.IllegalZip.value+"&ObjName=IllegalZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes,status=yes");

}
<%end if %>
<%if sys_City="高雄市" then%>

function getIllZip(){
	runServerScript("getZipNameForCar.asp?ZipID="+myForm.IllegalZip.value);
}
<%end if %>
</script>
</html>
