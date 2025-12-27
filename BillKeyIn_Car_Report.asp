<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>逕舉資料建檔作業</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
AuthorityCheck(223)
'==========POST=========
'單號
if trim(request("billno"))="" then
	theBillno=""
else
	theBillno=trim(request("billno"))
end if
'new代表新增案件 , update 代表資料庫已有該案件
if trim(request("filetype"))="" then
	thefiletype=""
else
	thefiletype=trim(request("filetype"))
end if
' 告發類別
' theBilltype=1  1 攔停  2 逕舉
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
	'smith remark
	'if left(trim(changeDate),1)="0" then
	'	theFormatDate=cint(mid(trim(changeDate),2,2))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'else
	'	theFormatDate=cint(left(trim(changeDate),3))+1911&"/"&mid(trim(changeDate),4,2)&"/"&mid(trim(changeDate),6,2)
	'end if
	'DateFormatChange=theFormatDate
end function
'==========================
	'要到ApConfigure抓法條版本
	strRuleVer="select Value from ApConfigure where ID=3"
	set rsRuleVer=conn.execute(strRuleVer)
	if not rsRuleVer.eof then
		theRuleVer=trim(rsRuleVer("Value"))
	end if
	rsRuleVer.close
	set rsRuleVer=nothing
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

'新增告發單
if trim(request("kinds"))="DB_insert" then
	'chkIsExistBillNumFlag=0
	if trim(request("Billno1"))<>"" then
		strchkno="select BillNo from BillBase where BillNo='"&trim(request("Billno1"))&"' and RecordStateID=0"
		set rschkno=conn.execute(strchkno)
		if not rschkno.eof then
			chkIsExistBillNumFlag=1
		else
			chkIsExistBillNumFlag=0
		end if
		rschkno.close
		set rschkno=nothing
	else
		chkIsExistBillNumFlag=0
	end if
	if chkIsExistBillNumFlag=0 then
		'違規日期
		theIllegalDate=""
		if trim(request("IllegalDate"))<>"" then
			theIllegalDate=funGetDate(gOutDT(request("IllegalDate") ) &" "&left(trim(request("IllegalTime")),2)&":"&right(trim(request("IllegalTime")),2),1)
		else
			theIllegalDate = "null"
		end if	

		
		'檢查是否有罰款金額
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
		'第三責任險處理
		if trim(request("Insurance"))="" then
			theInsurance=0
		else
			theInsurance=cint(trim(request("Insurance")))
		end if
		'採証工具處理
		if trim(request("UseTool"))="" then
			theUseTool=0
		else
			theUseTool=trim(request("UseTool"))
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
		'建檔日期
		'theRecordDate=year(now)&"/"&month(now)&"/"&day(now)&" "&hour(now)&":"&minute(now)&":"&second(now)

		'時速處理
		if trim(request("IllegalSpeed"))="" then
			theIllegalSpeed="null"
		else
			theIllegalSpeed=trim(request("IllegalSpeed"))
		end if
		'限速處理
		if trim(request("RuleSpeed"))="" then
			theRuleSpeed="null"
		else
			theRuleSpeed=trim(request("RuleSpeed"))
		end if
		'輔助車種處理
		if trim(request("CarAddID"))="" then
			theCarAddID="0"
		else
			theCarAddID=trim(request("CarAddID"))
		end if
		'查流水號
		strSN="select BillBase_seq.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theSN=trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing

		'BillBase
		strInsert="insert into BillBase(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
					",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
					",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
					",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
					",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
					",BillMemID2,BillMem2,BillMemID3,BillMem3,BillMemID4,BillMem4" &_
					",BillFillerMemberID,BillFiller" &_
					",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
					",Note,EquipmentID,RuleVer,DriverSex,TrafficAccidentType)" &_
					" values("&theSN&",'"&trim(request("BillType"))&"','"&UCase(trim(request("Billno1")))&"'" &_
					",'"&UCase(trim(request("CarNo")))&"',"&trim(request("CarSimpleID")) &_						          
					","&theCarAddID&","&theIllegalDate&",'"&trim(request("IllegalAddressID"))&"'" &_
					",'"&trim(request("IllegalAddress"))&"','"&trim(request("Rule1"))&"',"&theIllegalSpeed &_
					","&theRuleSpeed&","&trim(request("ForFeit1"))&",'"&trim(request("Rule2"))&"'" &_
					","&theForFeit2&",'"&trim(request("Rule3"))&"',"&theForFeit3&",'"&trim(request("Rule4"))&"'" &_
					","&theForFeit4&","&theInsurance&","&theUseTool&",'"&trim(request("ProjectID"))&"'" &_
					",'"&UCase(trim(request("DriverPID")))&"',"& theDriverBirth &",'"&trim(request("DriverName"))&"'" &_
					",'"&trim(request("DriverAddress"))&"','"&trim(request("DriverZip"))&"','"&trim(request("MemberStation"))&"'" &_
					",'"&trim(request("BillUnitID"))&"','"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
					",'"&trim(request("BillMemID2"))&"','"&trim(request("BillMemName2"))&"'" &_
					",'"&trim(request("BillMemID3"))&"','"&trim(request("BillMemName3"))&"'" &_
					",'"&trim(request("BillMemID4"))&"','"&trim(request("BillMemName4"))&"'" &_
					",'"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
					","&theBillFillDate&","&theDealLineDate&",'0',0,SYSDate,'" & theRecordMemberID &"'" &_
					",'"&trim(request("Note"))&"','"&trim(request("IsMail"))&"','"&theRuleVer&"'" &_
					",'"&trim(request("DriverSex"))&"',''" &_
					")"
					'response.write strInsert
					conn.execute strInsert
					'theDriverBirth , theBillFillDate   


		'舉發單扣件明細檔 BillFastenerDetail
		if trim(request("Fastener1"))<>"" then
			strInsFastene1="insert into BillFastenerDetail(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
						" values(BillFastenerDetail_seq.nextval,"&theSN&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener1"))&"','')"
			conn.execute strInsFastene1
		end if
		if trim(request("Fastener2"))<>"" then
			strInsFastene2="insert into BillFastenerDetail(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
						" values(BillFastenerDetail_seq.nextval,"&theSN&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener2"))&"','')"
			conn.execute strInsFastene2
		end if

	%>
	<script language="JavaScript">
		//alert("新增完成");
	</script>
	<%
	else
	%>
	<script language="JavaScript">
		alert("此單號：<%=UCase(trim(request("Billno1")))%>，已建檔！！");
	</script>
	<%
	end if
end if

Session.Contents.Remove("BillTime_Report")
BillTime_ReportTmp=DateAdd("s" , 1, now)
Session("BillTime_Report")=date&" "&hour(BillTime_ReportTmp)&":"&minute(BillTime_ReportTmp)&":"&second(BillTime_ReportTmp)
'response.write Session("BillTime_Report")

'總共幾筆
Session.Contents.Remove("BillCnt_Report")
Session.Contents.Remove("BillOrder_Report")
strSqlCnt="select count(*) as cnt from BillBase where BillTypeID='2' and BillStatus in ('0') and RecordStateID=0 and RecordMemberID="&theRecordMemberID&" and ImageFileName is null"
set rsCnt1=conn.execute(strSqlCnt)
	Session("BillCnt_Report")=trim(rsCnt1("cnt"))
	Session("BillOrder_Report")=trim(rsCnt1("cnt"))+1
rsCnt1.close
set rsCnt1=nothing

bIllegalDate=""
bIllegalTime=""
bIllegalAddressID=""
bIllegalAddress=""
bRule1=""
bForFeit1=""
bRule2=""
bRule4=""
bForFeit2=""
bLoginID1=""
bBillMem1=""
bBillMemID1=""
bLoginID2=""
bBillMem2=""
bBillMemID2=""
bLoginID3=""
bBillMem3=""
bBillMemID3=""
bLoginID4=""
bBillMem4=""
bBillMemID4=""
bBillUnitID=""
bBillType=""
bDealLineDate=""
bMemberStation=""
bBillFillDate=""
bRuleSpeed=""
bEquipMent=""
bCarAddId=""
'抓上一筆的資料

strSql="select * from (select * from BillBase" &_
	" where BillTypeID='2' and BillStatus in ('0') and RecordStateID=0 and RecordMemberID="&theRecordMemberID &_
	" and RecordDate between TO_DATE('"&date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') " &_
	" and TO_DATE('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and ImageFileName is null order by RecordDate desc)" &_
	" where rownum=1"
set rs1=conn.execute(strSql)
if not rs1.eof then
	if trim(rs1("BillNo"))<>"" and not isnull(rs1("BillNo")) then
		bBillType="1"
	else
		bBillType="2"
	end if
	if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
		bRuleSpeed=trim(rs1("RuleSpeed"))
	end	if
	if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
		bIllegalDate=ginitdt(trim(rs1("IllegalDate")))
	end if
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
		bIllegalTime=theChangeTime
	end if
	if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
		bIllegalAddressID=trim(rs1("IllegalAddressID"))
	end	if
	if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
		bIllegalAddress=trim(rs1("IllegalAddress"))
	end	if
	if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
		bRule1=trim(rs1("Rule1"))
	end	if
	if trim(rs1("ForFeit1"))<>"" and not isnull(rs1("ForFeit1")) then
		bForFeit1=trim(rs1("ForFeit1"))
	end	if
	if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
		bRule2=trim(rs1("Rule2"))
	end	if
	if trim(rs1("ForFeit2"))<>"" and not isnull(rs1("ForFeit2")) then
		bForFeit2=trim(rs1("ForFeit2"))
	end	if
	if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
		bRule4=trim(rs1("Rule4"))
	end	if
	if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
		strMem1="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID1"))
		set rsMem1=conn.execute(strMem1)
		if not rsMem1.eof then
			bLoginID1=trim(rsMem1("LoginID"))
		end if
		rsMem1.close
		set rsMem1=nothing
	end if
	if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
		bBillMem1=trim(rs1("BillMem1"))
	end if
	if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
		bBillMemID1=trim(rs1("BillMemID1"))
	end if
	if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
		strMem2="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID2"))
		set rsMem2=conn.execute(strMem2)
		if not rsMem2.eof then
			bLoginID2=trim(rsMem2("LoginID"))
		end if
		rsMem2.close
		set rsMem2=nothing
	end if
	if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
		bBillMem2=trim(rs1("BillMem2"))
	end if
	if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
		bBillMemID2=trim(rs1("BillMemID2"))
	end if
	if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
		strMem3="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID3"))
		set rsMem3=conn.execute(strMem3)
		if not rsMem3.eof then
			bLoginID3=trim(rsMem3("LoginID"))
		end if
		rsMem3.close
		set rsMem3=nothing
	end if
	if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
		bBillMem3=trim(rs1("BillMem3"))
	end if
	if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
		bBillMemID3=trim(rs1("BillMemID3"))
	end if
	if trim(rs1("BillMemID4"))<>"" and not isnull(rs1("BillMemID4")) then
		strMem4="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID4"))
		set rsMem4=conn.execute(strMem4)
		if not rsMem4.eof then
			bLoginID4=trim(rsMem4("LoginID"))
		end if
		rsMem4.close
		set rsMem4=nothing
	end if
	if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
		bBillMem4=trim(rs1("BillMem4"))
	end if
	if trim(rs1("BillMemID4"))<>"" and not isnull(rs1("BillMemID4")) then
		bBillMemID4=trim(rs1("BillMemID4"))
	end if
	if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
		bBillUnitID=trim(rs1("BillUnitID"))
	end if
	if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
		bDealLineDate=ginitdt(trim(rs1("DealLineDate")))
	end if
	if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
		bMemberStation=trim(rs1("MemberStation"))
	end if
	if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
		bBillFillDate=trim(ginitdt(rs1("BillFillDate")))
	end if
	if trim(rs1("UseTool"))<>"" and not isnull(rs1("UseTool")) then
		bUseTool=trim(rs1("UseTool"))
	end if
	if trim(rs1("EquipMentID"))<>"" and not isnull(rs1("EquipMentID")) then
		bEquipMent=trim(rs1("EquipMentID"))
	end if
	if trim(rs1("CarAddId"))<>"" and not isnull(rs1("CarAddId")) then
		bCarAddId=trim(rs1("CarAddId"))
	end if
end if 
rs1.close
set rs1=nothing

%>

<style type="text/css">
<!--
.style1 {font-size: 14px}
.style3 {font-size: 15px}
.style4 {
	color: #FF0000;
	font-size: 12px
	}
.style5 {
	font-size: 12px
	}
.style6 {font-size: 16px}
.style7 {
	color: #FF0000;
	font-size: 12px;
	line-height:14px
	}
.style8 {
	color: #000000;
	font-size: 12px;
	line-height:14px;
}
.style9 {
	color: #FF0000;
	font-size: 12px;
	line-height:14px;
	font-weight: bold;
}
.style10 {
	color: #FF0000;
	font-size: 12px;
	line-height:14px;
}
.btn2 {font-size: 13px}
#LayerA1 {
	position:absolute;
	width:980px;
	height:182px;
	z-index:1;
	overflow: scroll;
}
.style99 {
	font-size: 15px;
	line-height:18px;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
<!-- #include file="../Common/Bannernoimage.asp"-->
	<form name="myForm" method="post">  

		<table width='985' border='1' align="center" cellpadding="1">
			<tr bgcolor="#1BF5FF">
				<td colspan="6" height="45"><strong>逕舉資料建檔作業</strong>&nbsp; &nbsp; 日期格式：951220 &nbsp;時間格式：2300(24小時制)&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
				建檔日期：<%=ginitdt(now)%>
				<br>
				<input type="checkbox" name="ReportChk" value="1" onclick="funcReportChk();" <%
				if bBillType="1" then
					response.write "checked"
				end if
				%>>逕舉手開單&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
				<input type="checkbox" name="CaseInByMem" value="1" <%
			if sys_City="嘉義縣" or sys_City="嘉義市" then
				if trim(request("CaseInByMem"))="1" then
					response.write "checked"
				end if
			end if
				%>>逾違規日期超過三個月強制建檔
				</td>
			</tr>	
			<tr>
				<td bgcolor="#EBE5FF"><div align="right">單號</div></td>
				<td <%
				if sys_City<>"嘉義縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"嘉義市" and sys_City<>"台南市" and sys_City<>"宜蘭縣" and sys_City<>"雲林縣" then
					response.write "colspan='6'"
				end if
				%>>
				
				<input name="Billno1" type="text" value="<%=theBillno%>" size="10" maxlength="9" onBlur="CheckBillNoExist();" onkeydown="funTextControl(this);" style=ime-mode:disabled <%
				if bBillType<>"1" then
					response.write "disabled"
				end if
				%>>
				</td>
<%if sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" or sys_City="嘉義市" or sys_City="台南市" or sys_City="宜蘭縣" then%>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>填單日期</div></td>
				<td colspan="3">
				<input type="text" size="10" value="<%
				if bBillType<>"1" then
					response.write ginitdt(date)
				else
					if trim(bBillFillDate)="" then
						response.write ginitdt(date)
					else
						response.write bBillFillDate
					end if
				end if
				%>" maxlength="6" name="BillFillDate" onfocus="this.select()" onBlur="getDealLineDate()" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
<%elseif sys_City="雲林縣" then%>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>違規車號</div></td>
				<td colspan="3">
				<input type="text" size="10" name="CarNo" onBlur="getVIPCar();" onkeydown="funTextControl(this);" style=ime-mode:disabled maxlength="8">
			    <div id="Layer7" style="position:absolute; width:140px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold;"></div>
				</td>
<%end if%>
			</tr>
			<tr>
<%if sys_City<>"雲林縣" then%>
			  <td bgcolor="#EBE5FF" width="13%"><div align="right"><span class="style4">＊</span>違規車號</div></td>
				<td width="32%">
				<input type="text" size="10" name="CarNo" onBlur="getVIPCar();" onkeydown="funTextControl(this);" style=ime-mode:disabled maxlength="8">
			    <div id="Layer7" style="position:absolute; width:140px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold;"></div>
				</td>
				<td bgcolor="#EBE5FF" width="13%"><div align="right"><span class="style4">＊</span>簡式車種</div></td>
				<td colspan="3">
				<input type="text" maxlength="1" size="4" value="" name="CarSimpleID" onBlur="getRuleAll();" onkeydown="funTextControl(this);" onfocus="this.select();" style=ime-mode:disabled>
				<font color="#ff000" size="2">1汽車 / 2拖車/ 3重機/ 4輕機/ 6 臨時車牌</font>
				&nbsp;
				<div id="Layer170" style="position:absolute; width:245px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style8">550cc以上重機簡式<br>車種請選擇重機</span>
				</div>
				</td>
<%else%>
				<td bgcolor="#EBE5FF" width="13%"><div align="right"><span class="style4">＊</span>簡式車種</div></td>
				<td width="32%">
				<input type="text" maxlength="1" size="4" value="" name="CarSimpleID" onBlur="getRuleAll();" onkeydown="funTextControl(this);" onfocus="this.select();" style=ime-mode:disabled>
				<div id="Layer170" style="position:absolute; width:245px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style7">1汽車 / 2拖車/ 3重機<br>/ 4輕機/ 6 臨時車牌</span>
				</div>
				&nbsp;<img src="/image/space.gif" width="120" height="8">
				<div id="Layer170" style="position:absolute; width:245px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style8">550cc以上重機簡式<br>車種請選擇重機</span>
				</div>
				</td>
				<td bgcolor="#EBE5FF" align="right">輔助車種</td>
				<td colspan="3">
                 <input type="text" maxlength="2" size="4" value="<%
			if sys_City="宜蘭縣" then
				if trim(request("CarAddID"))="8" then
					response.write trim(request("CarAddID"))
				end if
			end if
				%>" name="CarAddID" onBlur="getAddID();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				<div id="Layer110" style="position:absolute; width:338px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style10">1大貨/ 2大客/ 3砂石/ 4土方/ 5動力/ 6貨櫃/ 7大型重機 /8拖吊<br>/9(550cc)重機 /10計程車/ 11危險物品</span>
				</div>
				</td>
<%end if%>

			</tr>
			<tr>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>違規日期</div></td>
				<td>
				<input type="text" size="10" maxlength="6" name="IllegalDate" onkeydown="funTextControl(this);" onblur="getDealLineDate_Stop()" value="<%=bIllegalDate%>" style=ime-mode:disabled>
				</td>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>違規時間</div></td>
				<td colspan="3">
				<input type="text" size="4" maxlength="4" name="IllegalTime" value="<%=bIllegalTime%>" onBlur="value=value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
			</tr>
<%if sys_City="雲林縣" or sys_City="宜蘭縣" or sys_City="台南縣" or sys_City="台南市" or sys_City="嘉義市"  then%>
			<tr>
				<td bgcolor="#EBE5FF"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" name="RuleSpeed" onBlur="RuleSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled  value="<%=bRuleSpeed%>">
				</td>
				<td bgcolor="#EBE5FF"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled >
				</td>
			</tr>
<%end if%>
<%if sys_City<>"嘉義市" then %>
			<tr>
				<td bgcolor="#EBE5FF" width="13%"><div align="right">違規地點代碼</div></td>
				<td>
					<input type="text" size="10" value="<%=bIllegalAddressID%>" name="IllegalAddressID" onKeyUp="getillStreet();" onblur="funGetSpeedRule()" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<input type="hidden" name="OldIllegalAddressID" value="<%=bIllegalAddressID%>">
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>違規地點</div></td>
				<td colspan="3">
					<input type="text" size="29" value="<%=bIllegalAddress%>" name="IllegalAddress" style=ime-mode:active onkeyup="AutoGetIllStreet();" onblur="funGetSpeedRule()" onkeydown="funTextControl(this);">
					<input type="checkbox" name="chkHighRoad" value="1" <%if trim(request("chkHighRoad"))="1" then response.write "checked"%> onclick="setIllegalRule()"><span class="style1">快速道路</span>
				</td>
			</tr>
<%end if%>
<%if sys_City="彰化縣" or sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then%>

			<tr>
				<td bgcolor="#EBE5FF"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" name="RuleSpeed" onBlur="RuleSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled  value="<%=bRuleSpeed%>">
				</td>
				<td bgcolor="#EBE5FF"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled >
				</td>
			</tr>
<%end if%>
			<tr>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>違規法條一</div></td>
				<td colspan="5">
					<input type="text" maxlength="8" size="10" value="<%=bRule1%>" name="Rule1" onkeyup="getRuleData1();" onfocus="this.select()" onkeydown="funTextControl(this);" onchange="DelSpace1();" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")' alt="查詢法條">
					<img src="../Image/BillLawPlusButton2.png" width="25" height="23" onclick="Add_LawPlus()" alt="附加說明">
					<div id="Layer1" style="position:absolute ; width:560px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
					if trim(bRule1)<>"" then
						strRule1="select IllegalRule from Law where ItemID='"&trim(bRule1)&"' and Version='"&trim(theRuleVer)&"'"
						set rsRule1=conn.execute(strRule1)
						if not rsRule1.eof then
							response.write trim(rsRule1("IllegalRule"))
						end if
						rsRule1.close
						set rsRule1=nothing
						if trim(bRule4)<>"" then
							response.write "("&bRule4&")"
						end if
					end if
					%></div>
					<input type="hidden" name="ForFeit1" value="<%=trim(bForFeit1)%>">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF"><div align="right">違規法條二</div></td>
				<td colspan="5">
					<input type="text" maxlength="8" size="10" value="<%
				if sys_City<>"南投縣" then
					response.write bRule2
				end if
					%>" name="Rule2" onkeyup="getRuleData2();" onkeydown="funTextControl(this);" onchange="DelSpace2();" style=ime-mode:disabled onBlur="TabFocus()">
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer2" style="position:absolute ; width:590px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if sys_City<>"南投縣" then
					if trim(bRule2)<>"" then
						strRule2="select IllegalRule from Law where ItemID='"&trim(bRule2)&"' and Version='"&trim(theRuleVer)&"'"
						set rsRule2=conn.execute(strRule2)
						if not rsRule2.eof then
							response.write trim(rsRule2("IllegalRule"))
						end if
						rsRule2.close
						set rsRule2=nothing
					end if
				end if
					%></div>
					<input type="hidden" name="ForFeit2" value="<%
				if sys_City<>"南投縣" then
					response.write trim(bForFeit2)
				end if
					%>">
					<img src="space.gif" width="590" height="2">
					<img src="../Image/Law3.jpg" width="45" height="25" onclick='InsertLaw()' alt="違規法條三">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" id="TDLaw1" align="right"></td>
				<td colspan="5" id="TDLaw2"></td>
			</tr>
<%if sys_City="嘉義市" then %>
			<tr>
				<td bgcolor="#EBE5FF" width="13%"><div align="right">違規地點代碼</div></td>
				<td>
					<input type="text" size="10" value="<%=bIllegalAddressID%>" name="IllegalAddressID" onKeyUp="getillStreet();" onblur="funGetSpeedRule()" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<input type="hidden" name="OldIllegalAddressID" value="<%=bIllegalAddressID%>">
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>違規地點</div></td>
				<td colspan="3">
					<input type="text" size="29" value="<%=bIllegalAddress%>" name="IllegalAddress" style=ime-mode:active onkeyup="AutoGetIllStreet();" onblur="funGetSpeedRule()" onkeydown="funTextControl(this);">
					<input type="checkbox" name="chkHighRoad" value="1" <%if trim(request("chkHighRoad"))="1" then response.write "checked"%> onclick="setIllegalRule()"><span class="style1">快速道路</span>
				</td>
			</tr>
<%end if%>
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then%>

			<tr>
				<td bgcolor="#EBE5FF"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" name="RuleSpeed" onBlur="RuleSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled  value="<%
						response.write bRuleSpeed
					%>">
				</td>
				<td bgcolor="#EBE5FF"><div align="right">實際車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled >
				</td>
			</tr>
<%end if%>
			<tr>
				<td id="DLDate1" bgcolor="#EBE5FF" align="right"></td>
				<td id="DLDate2">
				<input type="hidden" size="6" value="" maxlength="6" name="DealLineDate" onBlur="DealLineDateReplace()" style=ime-mode:disabled>
				</td>
				<td id="DLDate3" bgcolor="#EBE5FF" align="right"></td>
				<td id="DLDate4" colspan="3"></td>
			</tr>

			<tr>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>舉發人代碼1</div></td>
		  		<td>
					<input type="text" size="10" name="BillMem1" value="<%=trim(bLoginID1)%>" onkeyup="getBillMemID1();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer12" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=trim(bBillMem1)%></div>
					<input type="hidden" value="<%=trim(bBillMemID1)%>" name="BillMemID1">
					<input type="hidden" value="<%=trim(bBillMem1)%>" name="BillMemName1">
				</td>
				<td bgcolor="#EBE5FF"><div align="right">舉發人代碼2</div></td>
		  		<td colspan="3">
					<input type="text" size="10" name="BillMem2" value="<%=trim(bLoginID2)%>" onkeyup="getBillMemID2();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer13" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=trim(bBillMem2)%></div>
					<input type="hidden" value="<%=trim(bBillMemID2)%>" name="BillMemID2">
					<input type="hidden" value="<%=trim(bBillMem2)%>" name="BillMemName2">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF"><div align="right">舉發人代碼3</div></td>
		  		<td>
					<input type="text" size="10" name="BillMem3" value="<%=trim(bLoginID3)%>" onkeyup="getBillMemID3();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=3","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer14" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=trim(bBillMem3)%></div>
					<input type="hidden" value="<%=trim(bBillMemID3)%>" name="BillMemID3">
					<input type="hidden" value="<%=trim(bBillMem3)%>" name="BillMemName3">
				</td>
				<td bgcolor="#EBE5FF"><div align="right">舉發人代碼4</div></td>
		  		<td colspan="3">
					<input type="text" size="10" name="BillMem4" value="<%=trim(bLoginID4)%>" onkeyup="getBillMemID4();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=4","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer17" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=trim(bBillMem4)%></div>
					<input type="hidden" value="<%=trim(bBillMemID4)%>" name="BillMemID4">
					<input type="hidden" value="<%=trim(bBillMem4)%>" name="BillMemName4">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>舉發單位</div></td>
				<td <%
				if sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" or sys_City="嘉義市" or sys_City="台南市" or sys_City="宜蘭縣" or sys_City="雲林縣" then
					response.write "colspan='5'"
				end if
				%>>
					<input type="text" size="10" name="BillUnitID" value="<%=trim(bBillUnitID)%>" onkeyup="getUnit();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<div id="Layer6" style="position:absolute ; width:180px; height:30px; z-index:0;  border: 1px none #000000;"><%
					if trim(bBillUnitID)<>"" then
						strUnitName="select UnitName from UnitInfo where UnitID='"&trim(bBillUnitID)&"'"
						set rsUnitName=conn.execute(strUnitName)
						if not rsUnitName.eof then
							response.write trim(rsUnitName("UnitName"))
						end if
						rsUnitName.close
						set rsUnitName=nothing
					end if
					%></div>
				</td>
<%if sys_City<>"嘉義縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"嘉義市" and sys_City<>"台南市" and sys_City<>"雲林縣" and sys_City<>"宜蘭縣" then%>
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>填單日期</div></td>
				<td colspan="3">
				<input type="text" size="10" value="<%
				if bBillType<>"1" then
					response.write ginitdt(date)
				else
					if trim(bBillFillDate)="" then
						response.write ginitdt(date)
					else
						response.write bBillFillDate
					end if
				end if
				%>" maxlength="6" name="BillFillDate" onBlur="getDealLineDate()" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
<%end if%>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF"><div align="right">專案代碼</div></td>
				<td>
					<input type="text" size="10" value="" name="ProjectID" style=ime-mode:disabled onkeyup="ProjectF5()" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onClick='window.open("Query_Project.asp","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer001" style="position:absolute ; width:180px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
				</td>
<%if sys_City="雲林縣" then%>	
				<td bgcolor="#EBE5FF"><div align="right"><span class="style4">＊</span>填單日期</div></td>
				<td colspan="3">
				<input type="text" size="10" value="<%
				if bBillType<>"1" then
					response.write ginitdt(date)
				else
					if trim(bBillFillDate)="" then
						response.write ginitdt(date)
					else
						response.write bBillFillDate
					end if
				end if
				%>" maxlength="6" name="BillFillDate" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
<%else%>
				<td bgcolor="#EBE5FF" align="right">輔助車種</td>
				<td colspan="3">
                 <input type="text" maxlength="2" size="4" value="<%
			if sys_City="宜蘭縣" then
				if trim(request("CarAddID"))="8" then
					response.write trim(request("CarAddID"))
				end if
			elseif sys_City="台南市" then
				if bUseTool="8" then
					response.write trim(request("CarAddID"))
				end if
			end if
				%>" name="CarAddID" onBlur="getAddID();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
                <div id="Layer110" style="position:absolute; width:338px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style10">1大貨/ 2大客/ 3砂石/ 4土方/ 5動力/ 6貨櫃/ 7大型重機 /8拖吊<br>/9(550cc)重機 /10計程車/ 11危險物品</span>
				</div>
				</td>
<%end if%>
			</tr>
			<tr>
				<td bgcolor="#EBE5FF" width="13%"><div align='right'>採証工具</div></td>
				<td>
					<input maxlength="1" size="4" value="<%
				if sys_City="嘉義縣" or sys_City="台南市" or sys_City="花蓮縣" or sys_City="高雄縣" then
					response.write bUseTool
				end if
					%>" name="UseTool"  onBlur="getFixID();" onkeydown="funTextControl(this);" type='text' style=ime-mode:disabled> 
			        <div id="Layer11" style="position:absolute; width:275px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold; visibility: hidden;"> <font color="#0000FF">&nbsp;&nbsp;<font color="#000000">固定桿編號：</font></font>
                    <input type='text' size='6' name='FixID' value='<%=request("FixID")%>' onBlur="setFixEquip();" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_FixEquip.asp","WebPageFix","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					</div>
					<font color="#ff000" size="2"> 1固定桿/ 2雷達三腳架/ 3相機/ 8逕舉手開單</font>
				</td>
				<td bgcolor="#EBE5FF"><div align="right">備註</div></td>
				<td>
					<input type="text" size="15" value="" name="Note" onkeydown="funTextControl(this);" style=ime-mode:active>
				</td>
				<td bgcolor="#EBE5FF"><div align="right">代保管物</div></td>
				<td>
					1. <input type="text" size="2" value="" name="Fastener1" onkeyup="getFastener1();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=1","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
					<div id="Layer8" style="position:absolute ; width:90px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000; visibility: hidden;"></div>
					<input type="hidden" value="" name="Fastener1Val">

					2. <input type="text" size="2" value="" name="Fastener2" onkeyup="getFastener2();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton2.png" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=2","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
					 <div id="Layer9" style="position:absolute ; width:90px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000; visibility: hidden;"></div>
	                 <input type="hidden" value="" name="Fastener2Val">
				</td>
			</tr>
			<tr>
				<td bgcolor="#1BF5FF" align="center" colspan="6">
					<input type="button" name value="儲 存 <%
					if sys_City="台東縣" or sys_City="高雄縣" then
						response.write "F9"
					else
						response.write "F2"
					end if
					%>" onclick="InsertBillVase();" <%
				'1:查詢 ,2:新增 ,3:修改 ,4:刪除
				if CheckPermission(223,2)=false then
					response.write "disabled"
				end if
					%> class="btn1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit1343" onClick="location='BillKeyIn_Car_Report.asp'" value="清 除 F4" class="btn1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit5322" onClick="funcOpenBillQry()" value="查 詢 F6" class="btn1">
					<input type="hidden" name="kinds" value="">
                    <span class="style1">
                    <span class="style3"><img src="/image/space.gif" width="29" height="8"></span>
                    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" class="btn1">
					<span class="style3"><img src="/image/space.gif" width="29" height="8"></span>
                    <input type="button" name="Submit4232" onClick="funPrintCaseList_Report();" value="建檔清冊 <%
					if sys_City="高雄縣" then
						response.write "F2"
					else
						response.write "F10"
					end if
					%>" class="btn1">
                </span>

				<br>

				<div id="Layer1f69" style="position:absolute; width:250px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style9">(重點工作報表針對特殊車種 需要在建檔時 輔助車種中 輸入   3砂石/ 8拖吊 /10計程車)</span>
				</div>
				<img src="/image/space.gif" width="250" height="8">
				<input type="button" name="SubmitBack2" onClick="location='BillKeyIn_Report_Back.asp?PageType=First'" value="<< 第一筆 Home" class="btn1">
				<img src="/image/space.gif" width="29" height="8">
				<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Report_Back.asp?PageType=Back'" value="< 上一筆 PgUp" class="btn1">
				<div id="Layer1c69" style="position:absolute; width:160px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style7">使用上一筆搜尋功能只能查詢到自己建檔且未入案的舉發單</span>
				</div>
				<img src="/image/space.gif" width="220" height="8">
				<!-- 逕舉類別 -->
				<input type="hidden" size="3" maxlength="1" value="2" name="BillType" readonly>
				<!-- 應到案處所 -->
				<input type="hidden" size="4" value="" name="MemberStation" onkeyup="getStation();">
				<div id="Layer5" style="position:absolute ; width:241px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
				<!-- 附加說明 -->
				<input type="hidden" value="<%=bRule4%>" name="Rule4">
			</tr>
		</table>		
<%if sys_City="XX縣都不開" then%>
		<table width='985' border='1' align="center" cellpadding="1">
			<tr>
				<td>
				<div id="LayerA1">
				
				<table width='960' border='1' align="center" cellpadding="1">
				<tr bgcolor="#33FFFF">
					<td width="9%"><span class="style99">單號</span></td>
					<td width="7%"><span class="style99">車號</span></td>
					<td width="11%"><span class="style99">違規時間</span></td>
					<td width="19%"><span class="style99">違規地點</span></td>
					<td width="16%"><span class="style99">違規法條</span></td>
					<td width="7%"><span class="style99">限速、重</span></td>
					<td width="7%"><span class="style99">車速、重</span></td>
					<td width="16%"><span class="style99">舉發人</span></td>
					<td width="8%"><span class="style99">操作</span></td>
				</tr>
<%
	strBillView="select * from billbase where BillTypeID='2' and BillStatus in ('0') and RecordStateID=0 and RecordMemberID="&theRecordMemberID&" and RecordDate between TO_DATE('"&date&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and ImageFileName is null order by RecordDate desc"
	set rsBillView=conn.execute(strBillView)
	If Not rsBillView.Bof Then rsBillView.MoveFirst 
	While Not rsBillView.Eof
		'for i=0 to 1000

%>
				<tr>
					<td><span class="style99"><%
					if trim(rsBillView("BillNo"))<>"" and not isnull(rsBillView("BillNo")) then
						response.write trim(rsBillView("BillNo"))
					else
						response.write "&nbsp;"
					end if
					%></span></td>
					<td><span class="style99"><%=trim(rsBillView("CarNo"))%></span></td>
					<td><span class="style99"><%=year(rsBillView("IllegalDate"))&"/"&Month(rsBillView("IllegalDate"))&"/"&Day(rsBillView("IllegalDate"))&" "&Hour(rsBillView("IllegalDate"))&":"&Minute(rsBillView("IllegalDate"))%></span></td>
					<td><span class="style99"><%
					if trim(rsBillView("IllegalAddress"))<>"" and not isnull(rsBillView("IllegalAddress")) then
						response.write trim(rsBillView("IllegalAddress"))
					else
						response.write "&nbsp;"
					end if
					%></span></td>
					<td><span class="style99"><%
					if trim(rsBillView("Rule1"))<>"" and not isnull(rsBillView("Rule1")) then
						response.write trim(rsBillView("Rule1"))
					else
						response.write "&nbsp;"
					end if
					if trim(rsBillView("Rule2"))<>"" and not isnull(rsBillView("Rule2")) then
						response.write "/"&trim(rsBillView("Rule2"))
					end if
					if trim(rsBillView("Rule3"))<>"" and not isnull(rsBillView("Rule3")) then
						response.write "/"&trim(rsBillView("Rule3"))
					end if
					if trim(rsBillView("Rule4"))<>"" and not isnull(rsBillView("Rule4")) then
						response.write "/"&trim(rsBillView("Rule4"))
					end if
					%></span></td>
					<td><span class="style99"><%
					if trim(rsBillView("RuleSpeed"))<>"" and not isnull(rsBillView("RuleSpeed")) then
						response.write trim(rsBillView("RuleSpeed"))
					else
						response.write "&nbsp;"
					end if
					%></span></td>
					<td><span class="style99"><%
					if trim(rsBillView("IllegalSpeed"))<>"" and not isnull(rsBillView("IllegalSpeed")) then
						response.write trim(rsBillView("IllegalSpeed"))
					else
						response.write "&nbsp;"
					end if
					%></span></td>
					<td><span class="style99"><%
					if trim(rsBillView("BillMem1"))<>"" and not isnull(rsBillView("BillMem1")) then
						response.write trim(rsBillView("BillMem1"))
					else
						response.write "&nbsp;"
					end if
					if trim(rsBillView("BillMem2"))<>"" and not isnull(rsBillView("BillMem2")) then
						response.write "/"&trim(rsBillView("BillMem2"))
					end if
					if trim(rsBillView("BillMem3"))<>"" and not isnull(rsBillView("BillMem3")) then
						response.write "/"&trim(rsBillView("BillMem3"))
					end if
					if trim(rsBillView("BillMem4"))<>"" and not isnull(rsBillView("BillMem4")) then
						response.write "/"&trim(rsBillView("BillMem4"))
					end if
					%></span></td>
					<td><span class="style99">
						<input type="button" name="b1" value="修改" onclick='window.open("../BillKeyIn/BillKeyIn_Car_Report_Update.asp?BillSN=<%=trim(rsBillView("SN"))%>","WebPage2_Update","left=0,top=0,location=0,width=1000,height=650,resizable=yes,scrollbars=yes")' <%
							'1:查詢 ,2:新增 ,3:修改 ,4:刪除
							if CheckPermission(234,3)=false then
								response.write "disabled"
							end if
							%> style="font-size: 10pt; width: 40px; height:26px;">
					</span></td>
				</tr>
<%		rsBillView.MoveNext
		'next
	Wend
	rsBillView.close
	set rsBillView=nothing
%>
				</table>
				
				</div>
				</td>
			</tr>
		</table>
<%end if%>
	</form>
<%
conn.close
set conn=nothing
%>
</body>
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
var TDProjectIDErrorLog=0;
var TDVipCarErrorLog=0;
var TodayDate=<%=ginitdt(date)%>;
var ButtonSubmit=0;
<%if sys_City="彰化縣" then %>
MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="台南市" or sys_City="宜蘭縣" then %>
MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="嘉義市" then %>
MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||IllegalAddressID,IllegalAddress||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="台南縣" then %>
MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%elseif sys_City="雲林縣" then %>
MoveTextVar("Billno1,CarNo||CarSimpleID,CarAddID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,BillFillDate||UseTool,Note,Fastener1,Fastener2");
<%else%>
MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%end if%>
//新增告發單
function InsertBillVase(){
	var error=0;
	var errorString="";
	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");
	if (myForm.ReportChk.checked==true){
		ReadBillNo=myForm.Billno1.value.replace(' ','');
		if (ReadBillNo==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入單號。";
		}else if(ReadBillNo.length!=9){
			error=error+1;
			errorString=errorString+"\n"+error+"：單號不足九碼。";
		}
	}
	if (myForm.BillType.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入告發類別。";
	}else if (myForm.BillType.value=="2"){
		
		/*smith remark 逕舉不一定要輸入固定桿編號. 可能是員警拍照
		if (myForm.FixID.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入固定桿編號。";
		}
		*/
	}
	if (myForm.CarNo.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規車號。";
	}
	if (myForm.CarSimpleID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入簡式車種。";
	}/*else if(myForm.CarNo.value != "" && chkCarNoFormat(myForm.CarNo.value)!= 0) {
		if (chkCarNoFormat(myForm.CarNo.value) != myForm.CarSimpleID.value){
			error=error+1;
			errorString=errorString+"\n"+error+"：車號格式與簡式車種不符。";
		}
	}*/
	if (myForm.IllegalDate.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規日期。";
	}else if(!dateCheck( myForm.IllegalDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
	}else if (!ChkIllegalDate(myForm.IllegalDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期已超過三個月期限。";
	}
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
	/*if (myForm.ReportChk.checked==false){
		if (myForm.IllegalAddressID.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入違規地點代碼。";
		}
	}*/
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
	}else if (myForm.Rule1.value.substr(0,2)>68){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條一輸入錯誤。";
	}
	if (myForm.Rule1.value==myForm.Rule2.value && myForm.Rule1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條一與違規法條二重複。";
	}
	if (myForm.Rule2.value!=""){
		if (TDLawErrorLog2==1){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二輸入錯誤。";
		}else if (myForm.Rule2.value.substr(0,2)>68){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二輸入錯誤。";
		}
	}
	if (TDLawErrorLog3==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條三輸入錯誤。";
	}
	if (TDLawNum>=1 && myForm.Rule3.value!=""){
		if (myForm.Rule1.value==myForm.Rule3.value){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條一與違規法條三重複。";
		}
		if (myForm.Rule2.value==myForm.Rule3.value){	
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二與違規法條三重複。";
		}
	}
	if (myForm.BillFillDate.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入填單日期。";
	}else if (!dateCheck( myForm.BillFillDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
<%if sys_City<>"宜蘭縣" and sys_City<>"嘉義縣" and sys_City<>"嘉義市" then%>
	}else if(TodayDate < myForm.BillFillDate.value){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期不得比今天晚。";
<%else%>
	}else if(TodayDate < myForm.BillFillDate.value && myForm.ReportChk.checked==true){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期不得比今天晚。";
<%end if%>
	}else if (!ChkIllegalDate(myForm.BillFillDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期已超過三個月期限。";
	}
	if (myForm.MemberStation.value==""){
		if(myForm.BillType.value=="1"){
			//攔停才嗆破輸入
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入應到案處所。";
		}
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
	}else if (!ChkIllegalDate(myForm.DealLineDate.value) && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期已超過三個月期限。";
	}
	if (myForm.BillUnitID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入舉發單位代號。";
		TDUnitErrorLog==0
	}else if (TDUnitErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發單位代號輸入錯誤。";
	}
	if (myForm.BillMem1.value==""){
		//固定桿不需要輸入舉發人
		if (myForm.UseTool.value!="1"){
		    error=error+1;
			errorString=errorString+"\n"+error+"：請輸入舉發人代碼1。";
		}
	}else if (TDMemErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼1 輸入錯誤。";
	}
	if (TDMemErrorLog2==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼2 輸入錯誤。";
	}
	if (TDMemErrorLog3==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼3 輸入錯誤。";
	}
	if (TDMemErrorLog4==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼4 輸入錯誤。";
	}
	/*
	if (myForm.BillMem1.value==myForm.BillMem2.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼1 與 舉發人代碼2 重複。";
	}else if (myForm.BillMem1.value==myForm.BillMem3.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼1 與 舉發人代碼3 重複。";
	}else if (myForm.BillMem1.value==myForm.BillMem4.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼1 與 舉發人代碼4 重複。";
	}
	if (myForm.BillMem2.value==myForm.BillMem3.value && myForm.BillMem2.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼2 與 舉發人代碼3 重複。";
	}else if (myForm.BillMem2.value==myForm.BillMem4.value && myForm.BillMem2.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼2 與 舉發人代碼4 重複。";
	}
	if (myForm.BillMem3.value==myForm.BillMem4.value && myForm.BillMem3.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人代碼3 與 舉發人代碼4 重複。";
	}
	*/
	if (TDFastenerErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物1 輸入錯誤。";
	}
	if (TDFastenerErrorLog2==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物2 輸入錯誤。";
	}
	if (myForm.Fastener1.value==myForm.Fastener2.value && myForm.Fastener1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物1 與代保管物2 重複。";
	}
	if (myForm.BillFillDate.value < myForm.IllegalDate.value){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比填單日晚。";
	}else if(TodayDate < myForm.IllegalDate.value){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比今天晚。";
	}
	if(myForm.DealLineDate.value < myForm.BillFillDate.value){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期不得比填單日期早。";
	}
	if (TDProjectIDErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：專案代碼輸入錯誤。";
	}
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
		if(parseInt(myForm.RuleSpeed.value) > parseInt(myForm.IllegalSpeed.value)){
			error=error+1;
			errorString=errorString+"\n"+error+"：限速、限重大於實際車速、車重。";
		}
	}
	if ((myForm.Rule1.value.substr(0,5))=="33101" || (myForm.Rule1.value.substr(0,2))=="40" || (myForm.Rule1.value.substr(0,5))=="43102"){
	<%if sys_City="台中市" then%>
		IllegalRule=getIllegalRule2(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
	<%else%>
		IllegalRule=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
	<%end if%>
		if (IllegalRule != myForm.Rule1.value){
			error=error+1;
			errorString=errorString+"\n"+error+"：超速法條與車速不符。";
		}
	}else if ((myForm.Rule2.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,5))=="43102"){
	<%if sys_City="台中市" then%>
		IllegalRule=getIllegalRule2(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
	<%else%>
		IllegalRule=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
	<%end if%>
		if (IllegalRule != myForm.Rule2.value){
			error=error+1;
			errorString=errorString+"\n"+error+"：超速法條與車速不符。";
		}
	}
	if ((myForm.Rule1.value.substr(0,2))=="36" && (myForm.CarSimpleID.value=="3" || myForm.CarSimpleID.value=="4")){
			error=error+1;
			errorString=errorString+"\n"+error+"：法條與車種不符。";
	}else if ((myForm.Rule2.value.substr(0,2))=="36" && (myForm.CarSimpleID.value=="3" || myForm.CarSimpleID.value=="4")){
			error=error+1;
			errorString=errorString+"\n"+error+"：法條與車種不符。";
	}
<%if sys_City="雲林縣" then %>
	if (TDVipCarErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：車號 "+myForm.CarNo.value+" 為業管車輛。";
	}
<%end if%>
<%if sys_City="台中市" then %>
	if (((myForm.Rule1.value.substr(0,2))=="55" || (myForm.Rule2.value.substr(0,2))=="55") && (myForm.ReportChk.checked==false)){
		error=error+1;
		errorString=errorString+"\n"+error+"：第55條不可逕行舉發。";
	}
<%end if%>
	if (((myForm.Rule1.value.substr(0,3))=="293" || (myForm.Rule2.value.substr(0,3))=="293") && (myForm.RuleSpeed.value=="" || myForm.IllegalSpeed.value=="")){
		error=error+1;
		errorString=errorString+"\n"+error+"：您選擇超重法條，但是未輸入限重或車重。";
	}
	if (error==0 && ButtonSubmit==0){
			getChkCarIllegalDate();
	}else{
		alert(errorString);
	}
}
//檢查同車號同法條在同一天違規日期及違規時間前後兩小時內
function setChkCarIllegalDate(CarCnt,Illdate,RuleDetail)
{
	var ErrorStr="";
	if (CarCnt=="1"){
		ChkCarIlldateFlag="1";
	}else{
		ChkCarIlldateFlag="0";
	}
<%if sys_City="雲林縣" or sys_City="南投縣" then%>
	if (myForm.ReportChk.checked!=true){
		getDealDateValue=<%=getReportDealDateValue%>;	//要加幾天
		BFillDateTemp=myForm.BillFillDate.value;
		if (BFillDateTemp.length >= 6 && myForm.BillType.value=="2"){
			Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
			Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
			Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
			var BFillDate=new Date(Byear,Bmonth-1,Bday);
			var DLineDate=new Date()
			DLineDate=DateAdd("d",getDealDateValue,BFillDate);
			Dyear=parseInt(DLineDate.getYear())-1911;
			Dmonth=parseInt(DLineDate.getMonth())+1;
			Dday=DLineDate.getDate();
			Dyear=Dyear.toString();
			if (Dmonth < 10){
				Dmonth="0"+Dmonth;
			}
			if (Dday < 10){
				Dday="0"+Dday;
			}
		}
	}else{	//逕舉手開單+攔停天數
		getDealDateValue="30";
		BFillDateTemp=myForm.BillFillDate.value;
		if (BFillDateTemp.length >= 6 && myForm.BillType.value=="2"){
			Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
			Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
			Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
			var BFillDate=new Date(Byear,Bmonth-1,Bday);
			var DLineDate=new Date()
			DLineDate=DateAdd("d",getDealDateValue,BFillDate);
			Dyear=parseInt(DLineDate.getYear())-1911;
			Dmonth=parseInt(DLineDate.getMonth())+1;
			Dday=DLineDate.getDate();
			Dyear=Dyear.toString();
			if (Dmonth < 10){
				Dmonth="0"+Dmonth;
			}
			if (Dday < 10){
				Dday="0"+Dday;
			}
		}
	}
	if (myForm.DealLineDate.value!=Dyear+Dmonth+Dday){
		ErrorStr=ErrorStr+"應到案日不是違規日加"+getDealDateValue+"天，請確認是否正確。";
	}
<%end if%>
	if (RuleDetail==1){
		ErrorStr=ErrorStr+"\n違規事實與簡式車種不符，請確認是否正確。";
	}
	if (ChkCarIlldateFlag=="1"){
		ErrorStr=ErrorStr+"\n此車號於"+Illdate+"，有相同違規舉發，請確認有無連續開單。";
	}
	if (ErrorStr!=""){
		if(confirm(ErrorStr+"\n是否確定要存檔？")){
			myForm.kinds.value="DB_insert";
			ButtonSubmit=1;
			myForm.submit();
		}
	}else{
		myForm.kinds.value="DB_insert";
		ButtonSubmit=1;
		myForm.submit();
	}
}
//檢查同車號同法條在同一天違規日期及違規時間前後兩小時內
function getChkCarIllegalDate(){
	NewIllDate=myForm.IllegalDate.value;
	NewIllTime=myForm.IllegalTime.value;
	NewIllRule1=myForm.Rule1.value;
	NewIllRule2=myForm.Rule2.value;
	NewCarNo=myForm.CarNo.value;
	NewCarSimpleID=myForm.CarSimpleID.value;
	runServerScript("getChkCarIllegalDate.asp?CarID="+NewCarNo+"&IllDate="+NewIllDate+"&IllTime="+NewIllTime+"&IllRule1="+NewIllRule1+"&IllRule2="+NewIllRule2+"&CarSimpleID="+NewCarSimpleID);
}

//是否為特殊用車&檢查是否有同車號在同一天建檔
function getVIPCar(){
	myForm.CarNo.value=myForm.CarNo.value.toUpperCase();
	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");
	if (myForm.CarNo.value.length >= 1){
		var CarNum=myForm.CarNo.value;
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType==0){
			alert("車牌格式錯誤，如該車輛無車牌或舊式車牌則可忽略此訊息！");
<%if sys_City<>"南投縣" and sys_City<>"彰化縣" and sys_City<>"屏東縣" then %>
			//myForm.CarNo.select();
<%end if%>
		}else{
			runServerScript("getVIPCarForKeyIn.asp?CarID="+CarNum+"&BillType=2");
			myForm.CarSimpleID.value=CarType;
			//myForm.CarSimpleID.select();
		}
	}else{
		Layer7.innerHTML=" ";
		myForm.CarSimpleID.value="";
	}
}
//檢查輔助車種
function getAddID(){
	//myForm.CarAddID.value=myForm.CarAddID.value.replace(/[^\d]/g,'');
	if (myForm.CarAddID.value.length>0){
		if (myForm.CarAddID.value != "1" && myForm.CarAddID.value != "2" && myForm.CarAddID.value != "3" && myForm.CarAddID.value != "4" && myForm.CarAddID.value != "5" && myForm.CarAddID.value != "6" && myForm.CarAddID.value != "7" && myForm.CarAddID.value != "8" && myForm.CarAddID.value != "9" && myForm.CarAddID.value != "10" && myForm.CarAddID.value != "11" && myForm.CarAddID.value != "0"){
			alert("輔助車種填寫錯誤!");
			//myForm.CarAddID.value = "";
			myForm.CarAddID.select();
		}
	}
}
//檢查簡式車種
function getRuleAll(){
	//myForm.CarSimpleID.value=myForm.CarSimpleID.value.replace(/[^\d]/g,'');
	if (myForm.CarSimpleID.value.length>0){
		if (myForm.CarSimpleID.value != "1" && myForm.CarSimpleID.value != "2" && myForm.CarSimpleID.value != "3" && myForm.CarSimpleID.value != "4" && myForm.CarSimpleID.value != "6"){
			alert("簡式車種填寫錯誤!");
			myForm.CarSimpleID.select();
			//myForm.CarSimpleID.value = "";
		}
	}
}
//法條刪掉其他符號
function DelSpace1(){
	myForm.Rule1.value=myForm.Rule1.value.replace(/[^\d]/g,'');
	getRuleData1();
}
function DelSpace2(){
	myForm.Rule2.value=myForm.Rule2.value.replace(/[^\d]/g,'');
	getRuleData2();
}
function DelSpace3(){
	myForm.Rule3.value=myForm.Rule3.value.replace(/[^\d]/g,'');
	getRuleData3();
}
//違規事實1(ajax)
function getRuleData1(){
	if (myForm.Rule1.value.length > 6){
		var Rule1Num=myForm.Rule1.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail_forLawPlus.asp?RuleOrder=1&RuleID="+Rule1Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		CallChkLaw1();
		if ((myForm.Rule1.value.substr(0,2))!="33" && (myForm.Rule1.value.substr(0,2))!="40" && (myForm.Rule1.value.substr(0,2))!="43" && (myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="33" && (myForm.Rule2.value.substr(0,2))!="40" && (myForm.Rule2.value.substr(0,2))!="43" && (myForm.Rule2.value.substr(0,2))!="29"){
			myForm.RuleSpeed.value="";
		}
	}else if (myForm.Rule1.value.length <= 6 && myForm.Rule1.value.length > 0){
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=1;
	}else{
		Layer1.innerHTML=" ";
		myForm.ForFeit1.value="";
		TDLawErrorLog1=0;
	}
	AutoGetRuleID(1);
}
//違規事實2(ajax)
function getRuleData2(){
	if (myForm.Rule2.value.length > 6){
		var Rule2Num=myForm.Rule2.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail.asp?RuleOrder=2&RuleID="+Rule2Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		CallChkLaw2();
		if ((myForm.Rule1.value.substr(0,2))!="33" && (myForm.Rule1.value.substr(0,2))!="40" && (myForm.Rule1.value.substr(0,2))!="43" && (myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="33" && (myForm.Rule2.value.substr(0,2))!="40" && (myForm.Rule2.value.substr(0,2))!="43" && (myForm.Rule2.value.substr(0,2))!="29"){
			myForm.RuleSpeed.value="";
		}
	}else if (myForm.Rule2.value.length <= 6 && myForm.Rule2.value.length > 0){
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=1;
	}else{
		Layer2.innerHTML=" ";
		myForm.ForFeit2.value="";
		TDLawErrorLog2=0;
	}
	AutoGetRuleID(2);

}
//違規事實3(ajax)
function getRuleData3(){
	if (myForm.Rule3.value.length > 6){
		var Rule3Num=myForm.Rule3.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail.asp?RuleOrder=3&RuleID="+Rule3Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		//CallChkLaw3();
	}else if (myForm.Rule3.value.length <= 6 && myForm.Rule3.value.length > 0){
		Layer3.innerHTML=" ";
		myForm.ForFeit3.value="";
		TDLawErrorLog3=1;
	}else{
		Layer3.innerHTML=" ";
		myForm.ForFeit3.value="";
		TDLawErrorLog3=0;
	}
}
//增加違規法條
function InsertLaw(){
	TDLawNum=1;
	TDLaw1.innerHTML="違規法條三";
	TDLaw2.innerHTML="<input type='text' size='10' value='' name='Rule3' onKeyUp='getRuleData3();' onchange='DelSpace1();'  onkeydown='funTextControl(this);'> <img src='../Image/BillkeyInButton2.png' width='25' height='23' onclick='OpenQueryLaw3()' alt='查詢法條'> <div id='Layer3' style='position:absolute ; width:589px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;'></div><input type='hidden' name='ForFeit3' value=''>";

	if (myForm.ReportChk.checked==true){
	<%if sys_City="彰化縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" or sys_City="宜蘭縣" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||IllegalAddressID,IllegalAddress||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,CarNo||CarSimpleID,CarAddID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,BillFillDate||UseTool,Note,Fastener1,Fastener2");
	<%else%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>
	}else{
	<%if sys_City="彰化縣" then %>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
		MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" or sys_City="宜蘭縣" then %>
		MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義市" then %>
		MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||IllegalAddressID,IllegalAddress||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南縣" then %>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" then %>
		MoveTextVar("Billno1,CarNo||CarSimpleID,CarAddID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,BillFillDate||UseTool,Note,Fastener1,Fastener2");
	<%else%>
		MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>
	}

	myForm.Rule3.focus();
}
function OpenQueryLaw3(){
	window.open("Query_Law.asp?LawOrder=3&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes");
}
function TabFocus(){
	//建檔時除了超重超速時游標才跳至限速限量欄位，其它法條則游標不跳至超重超速
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value;
<%if sys_City<>"南投縣" and sys_City<>"台中縣" and sys_City<>"雲林縣" and sys_City<>"彰化縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then %>
	if ((Rule1tmp.substr(0,5))!="33101" && (Rule1tmp.substr(0,2))!="40" && (Rule1tmp.substr(0,5))!="43102" && (Rule1tmp.substr(0,3))!="293" && (Rule2tmp.substr(0,5))!="33101" && (Rule2tmp.substr(0,2))!="40" && (Rule2tmp.substr(0,5))!="43102" && (Rule2tmp.substr(0,3))!="293"){
	<%if sys_City="屏東縣" then%>
		myForm.DealLineDate.select();
	<%else%>
		if (myForm.ReportChk.checked==false){
			myForm.BillMem1.select();
		}else{
			myForm.DealLineDate.select();
		}
	<%end if%>
	}
<%end if%>
	//MemberStationLaw="21,35,57,61,62";
	//法條代碼遇到21,35,57,61,62，應到案處所自動帶當地監理所
	//if (MemberStationLaw.indexOf(Rule1tmp.substr(0,2))!=-1 || MemberStationLaw.indexOf(Rule2tmp.substr(0,2))!=-1){
	//	myForm.MemberStation.value=<%=trim(BillLawMemberStation)%>;
	//	getStation();
	//}
}
//到案處所(ajax)
function getStation(){
	if (myForm.MemberStation.value.length > 1){
		var StationNum=myForm.MemberStation.value;
		runServerScript("getMemberStation.asp?StationID="+StationNum);
	}else{
		Layer5.innerHTML=" ";
		TDStationErrorLog=1;
	}
}
//舉發單位(ajax)
function getUnit(){
	myForm.BillUnitID.value=myForm.BillUnitID.value.toUpperCase();
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_Unit.asp?SType=U","WebPage_Station12","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillUnitID.value.length > 0){
		var BillUnitNum=myForm.BillUnitID.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum);
	}else{
		Layer6.innerHTML=" ";
		TDUnitErrorLog=1;
	}
}

function UserInputBillType(){

}
//逕舉不一定要輸入固定桿編號. 除了是下方選擇使用固定桿
function getFixID(){
	if (myForm.UseTool.value.length == "1"){
		if (myForm.UseTool.value != "0" && myForm.UseTool.value != "1" && myForm.UseTool.value != "2" && myForm.UseTool.value != "3" && myForm.UseTool.value != "8"){
			alert("採証工具填寫錯誤!");
			myForm.UseTool.select();
			//myForm.UseTool.value = "";
		}else if (myForm.UseTool.value == "1"){
			//Layer11.style.visibility = "visible"; 
		}else{
			//Layer11.style.visibility = "hidden"; 
		}
	}
}
//違規地點代碼(ajax)
function getillStreet(){
<%if sys_City<>"基隆市" and sys_City<>"彰化縣" then%>
	if (myForm.IllegalAddressID.value!=myForm.OldIllegalAddressID.value){
		myForm.IllegalAddressID.value=myForm.IllegalAddressID.value.toUpperCase();
	}
<%end if%>
	if (event.keyCode!=13){
		if (event.keyCode==116){	
			event.keyCode=0;
			OstreetID=myForm.IllegalAddressID.value;
			window.open("Query_Street.asp?OstreetID="+OstreetID,"WebPage_Street_People2","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
		}else if (myForm.IllegalAddressID.value.length >= 1){
			var illAddrNum=myForm.IllegalAddressID.value;
			runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
		}
	}
}
//舉發人一(ajax)
function getBillMemID1(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=1","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem1.value.length > 1){
		var BillMemNum=myForm.BillMem1.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=1&MemID="+BillMemNum);
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
//舉發人二(ajax)
function getBillMemID2(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=2","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem2.value.length > 1){
		var BillMemNum=myForm.BillMem2.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=2&MemID="+BillMemNum);
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
//舉發人三(ajax)
function getBillMemID3(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=3","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem3.value.length > 1){
		var BillMemNum=myForm.BillMem3.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=3&MemID="+BillMemNum);
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
//舉發人四(ajax)
function getBillMemID4(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=4","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem4.value.length > 1){
		var BillMemNum=myForm.BillMem4.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=4&MemID="+BillMemNum);
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
//保管物品一(ajax)
function getFastener1(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_Fastener.asp?FaOrder=1","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.Fastener1.value.length == 1){
		var FastenerNum=myForm.Fastener1.value;
		runServerScript("getFastener.asp?FastenerOrder=1&FastenerID="+FastenerNum);
	}else if (myForm.Fastener1.value.length == 0){
		Layer8.innerHTML=" ";
		TDFastenerErrorLog1=0;
		myForm.Fastener1Val.value="";
	}else{
		Layer8.innerHTML=" ";
		TDFastenerErrorLog1=1;
		myForm.Fastener1Val.value="";
	}
}
//保管物品二(ajax)
function getFastener2(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_Fastener.asp?FaOrder=2","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.Fastener2.value.length == 1){
		var FastenerNum=myForm.Fastener2.value;
		runServerScript("getFastener.asp?FastenerOrder=2&FastenerID="+FastenerNum);
	}else if (myForm.Fastener2.value.length == 0){
		Layer9.innerHTML=" ";
		TDFastenerErrorLog2=0;
		myForm.Fastener2Val.value="";
	}else{
		Layer9.innerHTML=" ";
		TDFastenerErrorLog2=1;
		myForm.Fastener2Val.value="";
	}
}

function getBillFillDate(){
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');
	if (myForm.IllegalDate.value.length >= 6 ){
		myForm.BillFillDate.value=myForm.IllegalDate.value;
		getDealLineDate();
	}
}
//逕舉由填單日期帶入應到案日期
function getDealLineDate(){
	if (myForm.ReportChk.checked!=true){
	<%if sys_City<>"屏東縣" then%>
		getDealDateValue=<%=getReportDealDateValue%>;	//要加幾天
		myForm.BillFillDate.value=myForm.BillFillDate.value.replace(/[^\d]/g,'');
		BFillDateTemp=myForm.BillFillDate.value;
		if (BFillDateTemp.length >= 6 && myForm.BillType.value=="2"){
			Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
			Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
			Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
			var BFillDate=new Date(Byear,Bmonth-1,Bday);
			var DLineDate=new Date()
			DLineDate=DateAdd("d",getDealDateValue,BFillDate);
			Dyear=parseInt(DLineDate.getYear())-1911;
			Dmonth=parseInt(DLineDate.getMonth())+1;
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
	<%end if%>
	}else{	//逕舉手開單+攔停天數
<%if (trim(Session("UnitLevelID"))<>"2" and sys_City="台中縣") or sys_City<>"台中縣" then%>
	<%if sys_City<>"基隆市" and sys_City<>"南投縣" and sys_City<>"屏東縣" and sys_City<>"台中縣" and sys_City<>"台中市" then%>
	<%if sys_City="台中縣" or sys_City="彰化縣" or sys_City="宜蘭縣" or sys_City="台南縣" or sys_City="台東縣" or sys_City="嘉義市" or sys_City="雲林縣" then%>
		getDealDateValue="30";
	<%elseif sys_City="台南市" then%>
		if (myForm.IsMail[0].checked!=true){
			getDealDateValue=<%=getStopDealDateValue%>;
		}else{
			getDealDateValue="30";
		}
	<%else%>
		getDealDateValue=<%=getStopDealDateValue%>;	//要加幾天
	<%end if%>
		myForm.BillFillDate.value=myForm.BillFillDate.value.replace(/[^\d]/g,'');
		BFillDateTemp=myForm.BillFillDate.value;
		if (BFillDateTemp.length >= 6 && myForm.BillType.value=="2"){
			Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
			Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
			Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
			var BFillDate=new Date(Byear,Bmonth-1,Bday);
			var DLineDate=new Date()
			DLineDate=DateAdd("d",getDealDateValue,BFillDate);
			Dyear=parseInt(DLineDate.getYear())-1911;
			Dmonth=parseInt(DLineDate.getMonth())+1;
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
	<%end if%>
<%end if%>
	}
}
//逕舉手開單由違規日期帶入應到案日期+14
function getDealLineDate_Stop(){
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');

	if(TodayDate < myForm.IllegalDate.value){
		alert("違規日期不得大於今天!!");
		myForm.IllegalDate.select();
	}
<%if sys_City="屏東縣" then%>
	if (myForm.ReportChk.checked!=false){
		getSDealDateValue=<%
			response.write getStopDealDateValue
		%>;
		//要加幾天
		BFillDateTemp=myForm.IllegalDate.value;
		if (BFillDateTemp.length >= 6){
			//myForm.BillFillDate.value=myForm.IllegalDate.value;
			Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
			Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
			Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
			var BFillDate=new Date(Byear,Bmonth-1,Bday)
			var DLineDate=new Date()
			DLineDate=DateAdd("d",getSDealDateValue,BFillDate);
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
<%end if%>
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
	return win;
}
//用固定桿編號抓出違規地點
function setFixEquip(){
	if (myForm.FixID.value.length > 2){
		var FixNum=myForm.FixID.value;
		runServerScript("getFixIDAddress.asp?FixNum="+FixNum);
	}
}
function RuleSpeedforLaw(){
	myForm.RuleSpeed.value=myForm.RuleSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then %>
	CallChkLaw1();
	CallChkLaw2();
<%end if%>
	var IntError=0;
	var StrError="";
	if (myForm.RuleSpeed.value > <%
	if sys_City="雲林縣" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：限速、限重超過<%
	if sys_City="雲林縣" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= <%
	if sys_City="雲林縣" then 
		response.write "100"
	else
		response.write "60"
	end if
	%>){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速<%
	if sys_City="雲林縣" then 
		response.write "100"
	else
		response.write "60"
	end if
	%>公里以上。";
			<%if sys_City="南投縣" then %>
				if (myForm.Rule2.value=="" && myForm.Rule1.value!="4340003"){
					myForm.Rule2.value="4340003";
					getRuleData2();
				}else if(TDLawNum==0 && myForm.Rule1.value!="4340003" && myForm.Rule2.value!="4340003"){
					InsertLaw();
					myForm.Rule3.value="4340003";
					getRuleData3();
				}
			<%else%>
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速60公里以上需加開法條4340003(處車主)!!";
			<%end if%>
			}
		}
	}
	if (IntError!=0){
		alert(StrError+"\n\n請確認是否正確!");
	}
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule();
	<%end if%>
}
function IllegalSpeedforLaw(){
	myForm.IllegalSpeed.value=myForm.IllegalSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then %>
	CallChkLaw1();
	CallChkLaw2();
<%end if%>
	var IntError=0;
	var StrError="";
	if (myForm.IllegalSpeed.value > <%
	if sys_City="雲林縣" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：車速、車重超過<%
	if sys_City="雲林縣" then 
		response.write "150"
	else
		response.write "100"
	end if
	%>。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= <%
	if sys_City="雲林縣" then 
		response.write "100"
	else
		response.write "60"
	end if
	%>){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速<%
	if sys_City="雲林縣" then 
		response.write "100"
	else
		response.write "60"
	end if
	%>公里以上。";
			<%if sys_City="南投縣" then %>
				if (myForm.Rule2.value=="" && myForm.Rule1.value!="4340003"){
					myForm.Rule2.value="4340003";
					getRuleData2();
				}else if(TDLawNum==0 && myForm.Rule1.value!="4340003" && myForm.Rule2.value!="4340003"){
					InsertLaw();
					myForm.Rule3.value="4340003";
					getRuleData3();
				}
			<%else%>
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速60公里以上需加開法條4340003(處車主)!!";
			<%end if%>
			}
		}
	}
	if (IntError!=0){
		alert(StrError+"\n\n請確認是否正確!");
	}

	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule();
	<%end if%>
}
function CallChkLaw1(){
	if (!funcChkLaw(myForm.Rule1.value) && !funcChkLaw(myForm.Rule2.value) && myForm.Rule1.value!="" && myForm.Rule2.value!=""){
		if (!funcChkLaw(myForm.Rule1.value)){
			alert("請確認法條一是否填寫正確");
		}
	}else if (!funcChkLaw(myForm.Rule1.value) && myForm.Rule2.value==""){
		alert("請確認法條一是否填寫正確");
	}<%if sys_City="台中縣" then%>else if ((myForm.Rule1.value.substr(0,2)!="33" && myForm.Rule2.value.substr(0,2)!="33") && (myForm.chkHighRoad.checked==true)){
		alert("快速道路選項為勾選狀態!!");
	}else if ((myForm.Rule1.value.substr(0,2)=="33" || myForm.Rule2.value.substr(0,2)=="33") && (myForm.chkHighRoad.checked==false)){
		alert("快速道路選項未勾選!!");
	}<%end if%>
}
function CallChkLaw2(){
	if (!funcChkLaw(myForm.Rule1.value) && !funcChkLaw(myForm.Rule2.value) && myForm.Rule1.value!="" && myForm.Rule2.value!=""){
		if (!funcChkLaw(myForm.Rule2.value)){
			alert("請確認法條二是否填寫正確");
		}
	}else if (!funcChkLaw(myForm.Rule2.value) && myForm.Rule1.value==""){
		alert("請確認法條二是否填寫正確");
	}
}
//法律條文建檔檢查
function funcChkLaw(thisLaw){
	if (thisLaw.length>=2){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			//當有打速限及車速時 法條一定落在33XXXX,40XXXX,43XXXX
			if ((thisLaw.substr(0,5))!="33101" && (thisLaw.substr(0,2))!="40" && (thisLaw.substr(0,5))!="43102" && (thisLaw.substr(0,3))!="293"){
				return false;
			}else{
				//違規地點含有"快速道路"判斷法條是否選33XXX而非選40XXX
				if ((myForm.IllegalAddress.value.indexOf("快速道路",0)) != -1){
					if ((thisLaw.substr(0,2))=="40"){
						return false;
					}else{
						return true;
					}
				}else{
					return true;
				}
			}
		}else{
			return true;
		}
	}else{
		return true;
	}
}
function CheckBillNoExist(){
	myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
	var BillNum=myForm.Billno1.value;
	if (myForm.Billno1.value!=""){
		ReadBillNo=myForm.Billno1.value.replace(' ','');
		if (ReadBillNo.length==9){
			runServerScript("getCheckBillNoExist.asp?BillNo="+BillNum);
		}else{
			alert("單號不足九碼！");
			myForm.Billno1.select();
		}
	}
}

//勾選後才可以輸入單號
function funcReportChk(){
	<%if trim(bDealLineDate)<>"" then%>
		var bDealLineDate=<%=trim(bDealLineDate)%>;
	<%else%>
		var bDealLineDate="";
	<%end if%>
	if (myForm.ReportChk.checked==true){
		myForm.Billno1.disabled=false;
		myForm.UseTool.value="8";
		//LayerDLDate.style.visibility = "visible"; 
		//LayerMStation.style.visibility = "visible";
		//myForm.MemberStation.disabled=false;
		DLDate1.innerHTML="應到案日期";
		DLDate3.innerHTML="是否郵寄";
		<%if sys_City="台中縣" then%>
			bDealLineDate=""
		<%end if%>
		<%if sys_City="雲林縣" then%>
			myForm.BillFillDate.value="";
			bDealLineDate=""
		<%end if%>
		DLDate2.innerHTML="<input type='text' size='6' value='"+bDealLineDate+"' maxlength='6' name='DealLineDate' onBlur='DealLineDateReplace()' onkeydown='funTextControl(this);' style=ime-mode:disabled>";
		DLDate4.innerHTML="<input type='radio' name='IsMail' value='1' <%
		if bEquipMent<>"-1" or isnull(bEquipMent) then
			response.write "checked"
		end if
		%>>是<input type='radio' name='IsMail' value='-1' <%
		if bEquipMent="-1" then
			response.write "checked"
		end if
		%>>否";
	<%if sys_City="彰化縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" or sys_City="宜蘭縣" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||IllegalAddressID,IllegalAddress||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,CarNo||CarSimpleID,CarAddID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,BillFillDate||UseTool,Note,Fastener1,Fastener2");
	<%else%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>

	}else{
<%if sys_City<>"屏東縣" then %>
		myForm.Billno1.value="";
		myForm.Billno1.disabled=true;
		if (myForm.UseTool.value=="8"){
			myForm.UseTool.value="";
		}
		//LayerDLDate.style.visibility = "hidden"; 
		//LayerMStation.style.visibility = "hidden"; 
		//myForm.MemberStation.disabled=true;
		myForm.MemberStation.Type="Text";
		DLDate1.innerHTML="";
		DLDate2.innerHTML="<input type='hidden' size='6' value='"+bDealLineDate+"' maxlength='6' name='DealLineDate' onBlur='DealLineDateReplace()' style=ime-mode:disabled>";
		DLDate3.innerHTML="";
		DLDate4.innerHTML="<input type='hidden' size='6' value='1' maxlength='6' name='IsMail' style=ime-mode:disabled>";
		getDealLineDate();
	<%if sys_City="彰化縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南市" or sys_City="宜蘭縣" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||IllegalAddressID,IllegalAddress||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="台南縣" then %>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,CarNo||CarSimpleID,CarAddID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||IllegalAddressID,IllegalAddress||Rule1||Rule2||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,BillFillDate||UseTool,Note,Fastener1,Fastener2");
	
	<%else%>
	MoveTextVar("Billno1||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID,BillFillDate||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
	<%end if%>
<%else%>
		myForm.Billno1.value="";
		myForm.Billno1.disabled=true;
		if (myForm.UseTool.value=="8"){
			myForm.UseTool.value="";
		}
		DLDate1.innerHTML="應到案日期";
		DLDate3.innerHTML="是否郵寄";
		DLDate2.innerHTML="<input type='text' size='6' value='"+bDealLineDate+"' maxlength='6' name='DealLineDate' onBlur='DealLineDateReplace()' onkeydown='funTextControl(this);' style=ime-mode:disabled>";
		DLDate4.innerHTML="<input type='radio' name='IsMail' value='1' checked>是";
		getDealLineDate();

		MoveTextVar("Billno1,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID||ProjectID,CarAddID||UseTool,Note,Fastener1,Fastener2");
<%end if%>
	}
}
function DealLineDateReplace(){
	myForm.DealLineDate.value=myForm.DealLineDate.value.replace(/[^\d]/g,'');

}


//逕舉建檔清冊
function funPrintCaseList_Report(){
	UrlStr="../Query/PrintCaseDataList_Report.asp?CallType=1";
	newWin(UrlStr,"CaseListWin2342",980,575,0,0,"yes","yes","yes","no");
}

function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false; 
<%if sys_City="台東縣" or sys_City="高雄縣" then%>
	}else if (event.keyCode==120){ //台東縣F9存檔
		event.keyCode=0;   
		InsertBillVase();
<%else%>
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
		InsertBillVase();
<%end if%>
	}else if (event.keyCode==115){ //F4清除
		event.keyCode=0;   
		location='BillKeyIn_Car_Report.asp'
	}else if (event.keyCode==117){ //F6查詢
		event.keyCode=0;   
		funcOpenBillQry();
	}else if (event.keyCode==119){ //F8關閉
		event.keyCode=0;   
		window.close();
<%if sys_City="高雄縣" then%>
	}else if (event.keyCode==113){ //高雄縣F2查詢未建檔
		event.keyCode=0;   
		funPrintCaseList_Report();
<%else%>
	}else if (event.keyCode==121){ //F10查詢未建檔
		event.keyCode=0;   
		funPrintCaseList_Report();
<%end if%>
	}else if (event.keyCode==33){ //上一筆PageUp
		event.keyCode=0;   
		location='BillKeyIn_Report_Back.asp?PageType=Back'
	}else if (event.keyCode==36){ //第一筆Home
		event.keyCode=0;   
		location='BillKeyIn_Report_Back.asp?PageType=First'
	}
}
function funcOpenBillQry(){
	Sys_RMemberID=<%=session("User_ID")%>;
	Sys_BillTypeID=2;
	window.open("EasyBillQry.asp","WebPage86_Update","left=300,top=200,location=0,width=350,height=160,resizable=yes,scrollbars=yes");
}
function AutoGetIllStreet(){	//按F5可以直接顯示相關路段
	if (event.keyCode==116){	
		event.keyCode=0;
		Ostreet=myForm.IllegalAddress.value;
		window.open("Query_Street.asp?OStreet="+Ostreet,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
}
function AutoGetRuleID(LawOrder){	//按F5可以直接顯示相關法條
	if (event.keyCode==116){	
		event.keyCode=0;
		if (LawOrder==1){
			ORuleID=myForm.Rule1.value;
		}else{
			ORuleID=myForm.Rule2.value;
		}
	window.open("Query_Law.asp?LawOrder="+LawOrder+"&RuleVer=<%=theRuleVer%>&ORuleID="+ORuleID,"WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
}
function ProjectF5(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_Project.asp","WebPage_Street_People","left=0,top=0,location=0,width=800,height=460,resizable=yes,scrollbars=yes");
	}
	if (myForm.ProjectID.value.length > 0){
		var BillProjectID=myForm.ProjectID.value;
		runServerScript("getProjectID.asp?BillProjectID="+BillProjectID);
	}else{
		Layer001.innerHTML="";
		TDProjectIDErrorLog=0;
	}
}
//附加說明
function Add_LawPlus(){
	if (myForm.Rule1.value==""){
		alert("請先輸入違規法條一!!");
	}else{
	RuleID=myForm.Rule1.value;
	window.open("Query_LawPlus.asp?RuleID="+RuleID+"&theRuleVer=<%=theRuleVer%>","WebPage1","left=20,top=10,location=0,width=500,height=455,resizable=yes,scrollbars=yes");
	}
}
function funGetSpeedRule(){
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule();
	<%end if%>
}
//用地點、車速抓違規法條
function setIllegalRule(){
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!="" && myForm.IllegalAddress.value!=""){
		if ((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		<%if sys_City="台中市" then%>
			IllegalRule=getIllegalRule2(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%else%>
			IllegalRule=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
		<%end if%>
			if (IllegalRule!="Null"){
				if ((myForm.Rule2.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,5))=="43102"){
					myForm.Rule2.value=IllegalRule;
					getRuleData2();
				}else{
					myForm.Rule1.value=IllegalRule;
					getRuleData1();
				}
			}
		}
	}
}
	//-----------上下左右-------------
	function funTextControl(obj){
		if (event.keyCode==13){ //Enter換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeEnter(obj.name);
		}	
		/*if (event.keyCode==37){ //左換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveLeft(obj.name);
		}else */else if (event.keyCode==38){ //上換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveLeft(obj.name);
		}/*else if (event.keyCode==39){ //右換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveRight(obj.name);
		}*/else if (event.keyCode==40){ //下換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveRight(obj.name);
		}else if (event.keyCode==9){ 
			if (obj==myForm.IllegalAddress){
				event.keyCode=0;
				event.returnValue=false;
<%if sys_City="彰化縣" or sys_City="嘉義縣" or sys_City="台東縣" or sys_City="高雄縣" then%>
				myForm.RuleSpeed.select();
<%elseif sys_City="嘉義市" then%>
			if (myForm.ReportChk.checked==false){
				myForm.BillMem1.select();
			}else{
				myForm.DealLineDate.select();
			}
<%else%>
				myForm.Rule1.select();
<%end if%>
			}
		}
<%if sys_City="雲林縣" then%>
		if (obj==myForm.BillFillDate){
			getDealLineDate();
		}
<%end if%>
	}
	//------------------------------
if (myForm.ReportChk.checked==false){
<%if sys_City="台南市" or sys_City="宜蘭縣" then%>
	myForm.BillFillDate.focus();
<%else%>
	myForm.CarNo.focus();
<%end if%>
}else{
	myForm.Billno1.focus();
}
funcReportChk();
<%if sys_City<>"台中縣" then%>
getDealLineDate();
<%else%>
if (myForm.ReportChk.checked==false){
	getDealLineDate();
}
<%end if%>
</script>
</html>
