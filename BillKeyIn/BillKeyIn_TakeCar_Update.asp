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
<title>攔停資料修改</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
'AuthorityCheck(223)
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

'修改告發單
if trim(request("kinds"))="DB_insert" then
	'有改單號的話，先檢查有沒有重覆的單號
	if trim(request("Billno1"))<>trim(request("OldBillNo")) then
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
		'交通事故類別
		if trim(request("TrafficAccidentType"))="" then
			theTrafficAccidentType=""
		else
			theTrafficAccidentType=trim(request("TrafficAccidentType"))
		end if

		'簽收狀況 A=A,U 2 3 =U ,5=''
		if UCase(trim(request("SignType")))="A" then
			UserSignType="A"
		elseif UCase(trim(request("SignType")))="U" or UCase(trim(request("SignType")))="2" or UCase(trim(request("SignType")))="3" then
			UserSignType="U"
		else
			UserSignType=""
		end if

		'BillBase
		strUpd="update BillBase set BillTypeID='"&trim(request("BillType"))&"',BillNo='"&UCase(trim(request("Billno1")))&"'" &_
			",CarNo='"&UCase(trim(request("CarNo")))&"',CarSimpleID="&trim(request("CarSimpleID")) &_
			",CarAddID="&theCarAddID&",IllegalDate="&theIllegalDate&_
			",IllegalAddressID='"&trim(request("IllegalAddressID"))&"',IllegalAddress='"&trim(request("IllegalAddress"))&"'" &_
			",Rule1='"&trim(request("Rule1"))&"',IllegalSpeed="&theIllegalSpeed&",RuleSpeed="&theRuleSpeed &_
			",ForFeit1="&trim(request("ForFeit1"))&",Rule2='"&trim(request("Rule2"))&"',ForFeit2="&theForFeit2 &_
			",Rule3='"&trim(request("Rule3"))&"',ForFeit3="&theForFeit3&",Rule4='"&trim(request("Rule4"))&"'" &_
			",ForFeit4="&theForFeit4&",Insurance="&theInsurance&",UseTool="&theUseTool &_
			",ProjectID='"&trim(request("ProjectID"))&"',DriverID='"&UCase(trim(request("DriverPID")))&"'" &_
			",DriverBirth="&theDriverBirth&",Driver='"&trim(request("DriverName"))&"'" &_
			",DriverAddress='"&trim(request("DriverAddress"))&"',DriverZip='"&trim(request("DriverZip"))&"'" &_
			",MemberStation='"&trim(request("MemberStation"))&"',BillUnitID='"&trim(request("BillUnitID"))&"'" &_
			",BillMemID1='"&trim(request("BillMemID1"))&"',BillMem1='"&trim(request("BillMemName1"))&"'" &_
			",BillMemID2='"&trim(request("BillMemID2"))&"',BillMem2='"&trim(request("BillMemName2"))&"'" &_
			",BillMemID3='"&trim(request("BillMemID3"))&"',BillMem3='"&trim(request("BillMemName3"))&"'" &_
			",BillMemID4='"&trim(request("BillMemID4"))&"',BillMem4='"&trim(request("BillMemName4"))&"'" &_
			",BillFillerMemberID='"&trim(request("BillMemID1"))&"',BillFiller='"&trim(request("BillMemName1"))&"'" &_
			",BillFillDate="&theBillFillDate&",DealLineDate="&theDealLineDate&_
			",Note='"&trim(request("Note"))&"',EquipmentID='"&trim(request("IsMail"))&"',DriverSex='"&trim(request("DriverSex"))&"'" &_
			",TrafficAccidentNo='"&trim(request("TrafficAccidentNo"))&"',TrafficAccidentType='"&theTrafficAccidentType&"',SignType='"&UserSignType&"'" &_
			" where SN="&trim(request("BillSN"))
			conn.execute strUpd
			if sys_City="台中市" then
				'ConnExecute strUpd&"~!@"&trim(request("OldBillData")),353
				ConnExecute trim(request("OldBillData")),353
			else
				ConnExecute strUpd,353
			end if
			'========================titan==========================
'			thewhere=""
'			CaseInDate=""
'			CaseINCnt=0
'			strCType="select a.DciCaseInDate from BillBaseDCIReturn a " &_
'			" where a.BillNo='"&trim(request("Billno1"))&"' and a.CarNo='"&trim(request("CarNo"))&"'" &_
'			" and ExchangeTypeID='W'" &_
'			" and Status in ('Y','S','n') order by DciCaseInDate Desc"
'			set rsCType=conn.execute(strCType)
'			if not rsCType.eof then
'				'response.write trim(rsCType("DciCaseInDate"))
'				CaseInDate=gOutDT(trim(rsCType("DciCaseInDate")))
'				CaseINCnt=1
'			else
'				strCType2="select ExchangeDate from DciLog where BillSn='"&trim(request("BillSN"))&"' " &_
'					" and ExchangeTypeID='W' " &_
'					" and (DciReturnStatusID in ('Y','S','n') or DciReturnStatusID is Null)" &_
'					" order by ExchangeDate Desc"
'				set rsCType2=conn.execute(strCType2)
'				if not rsCType2.eof then
'					CaseInDate=trim(rsCType2("ExchangeDate"))
'					CaseINCnt=1
'				end if
'				rsCType2.close
'				set rsCType2=nothing
'			end if
'			rsCType.close
'			set rsCType=nothing
'			'計算入案幾天
'			CountCaseIN=0
'			if CaseInDate<>"" then
'				CountCaseIN=dateDiff("d",CaseInDate,now)
'			end if
'			if Not ifnull(request("OldBillNo")) then
'				thewhere=" where BillSn="&trim(request("BillSN"))
'				'----入案成功超過7天才能修改----
'				'if (CountCaseIN>7 and CaseINCnt=1) or CaseINCnt=0 then
'					strUpd4="Update BillBaseDciReturn  set BillNo='"&trim(request("Billno1"))&"',CarNo='"&trim(request("CarNo"))&"' where BillNo='"&trim(request("OldBillNo"))&"' and CarNo='"&trim(request("OldCarNo"))&"' and ExchangeTypeID<>'A'"
'					conn.execute strUpd4
'					strSQL="update DCILog set BillNo='"&UCase(trim(request("Billno1")))&"',CarNo='"&UCase(trim(request("CarNo")))&"'"&thewhere
'					conn.execute strSQL
'				'end if
'				'-------------------------------
'				strSQL="update BillMailHistory set BillNo='"&UCase(trim(request("Billno1")))&"',CarNo='"&UCase(trim(request("CarNo")))&"'"&thewhere
'				conn.execute strSQL
'			else
'				thewhere1=" where BillSn="&trim(request("BillSN"))
'				'----入案成功超過7天才能修改----
'				if (CountCaseIN>7 and CaseINCnt=1) or CaseINCnt=0 then
'					'thewhere2=" where CarNo='"&trim(request("OldCarNo"))&"'"
'					'strSQL="update BillBaseDCIReturn set CarNo='"&UCase(trim(request("CarNo")))&"'"&thewhere
'					'conn.execute strSQL
'					strSQL="update DCILog set CarNo='"&UCase(trim(request("CarNo")))&"'"&thewhere1
'					conn.execute strSQL
'				end if
'				'-------------------------------
'				strSQL="update BillMailHistory set CarNo='"&UCase(trim(request("CarNo")))&"'"&thewhere1
'				conn.execute strSQL
'			end if
			'=======================================================
		'簽收狀況 BillUserSignDate
		if UCase(trim(request("SignType")))="2" or UCase(trim(request("SignType")))="3" or UCase(trim(request("SignType")))="5" then
			strSelSign="select * from BillUserSignDate where BillSn="&trim(request("BillSN"))
			set rsSelSign=conn.execute(strSelSign)
			if not rsSelSign.eof then
				strUpdSignType="Update BillUserSignDate set SignStateID='"&UCase(trim(request("SignType")))&"' where BillSn="&trim(request("BillSN"))
				conn.execute strUpdSignType
			else
				strInsSignType="insert into BillUserSignDate values("&trim(request("BillSN"))&",'"&UCase(trim(request("SignType")))&"','','')"
				conn.execute strInsSignType
			end if
			rsSelSign.close
			set rsSelSign=nothing
		else
			strDelSignType="delete from BillUserSignDate where BillSn="&trim(request("BillSN"))
			conn.execute strDelSignType
		end if

		'舉發單扣件明細檔 BillFastenerDetail
		strDel="delete from BillFastenerDetail where BillSN="&trim(request("BillSN"))
		conn.execute strDel
		if trim(request("Fastener1"))<>"" then
			strInsFastene1="insert into BillFastenerDetail(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
						" values(BillFastenerDetail_seq.nextval,"&trim(request("BillSN"))&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener1"))&"','')"
			conn.execute strInsFastene1
			ConnExecute strInsFastene1,353
		end if
		if trim(request("Fastener2"))<>"" then
			strInsFastene2="insert into BillFastenerDetail(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
						" values(BillFastenerDetail_seq.nextval,"&trim(request("BillSN"))&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener2"))&"','')"
			conn.execute strInsFastene2
			ConnExecute strInsFastene2,353
		end if
		if trim(request("Fastener3"))<>"" then
			strInsFastene3="insert into BillFastenerDetail(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
						" values(BillFastenerDetail_seq.nextval,"&trim(request("BillSN"))&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener3"))&"','')"
			conn.execute strInsFastene3
			ConnExecute strInsFastene3,353
		end if

%>
<script language="JavaScript">
	alert("修改完成!!");
</script>
<%
	else
%>
<script language="JavaScript">
	alert("此單號：<%=UCase(trim(request("Billno1")))%>，已經建檔!");
</script>
<%
	end if
end if

strSql="select * from BillBase where SN="&trim(request("BillSN"))
set rs1=conn.execute(strSql)
%>

<style type="text/css">
<!--
.style1 {font-size: 14px}
.style3 {font-size: 15px}
.style4 {
	color: #FF0000;
	font-size: 12px
}
.style5 {font-size: 12px}
.style6 {font-size: 16px}
.style7 {
	color: #FF0000;
	font-size: 15px
}
.style9 {
	color: #FF0000;
	font-size: 12px;
	line-height:14px;
}
.style8 {
	color: #000000;
	font-size: 12px;
	line-height:14px;
}
.style10 {
	color: #FF0000;
	font-size: 12px;
	line-height:14px;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
<%if sys_City<>"台中縣" then%>
<!-- #include file="../Common/Bannernoimage.asp"-->
<%end if%>
	<form name="myForm" method="post">  
		<table width='985' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="6"><strong>攔停資料修改</strong>&nbsp; &nbsp; 日期格式：951220 &nbsp;時間格式：2300(24小時制)&nbsp; &nbsp; <input type="checkbox" name="CaseInByMem" value="1" <%if trim(request("CaseInByMem"))="1" then response.write "checked"%>>逾違規日期超過三個月強制建檔</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>單號</div></td>
				<td>
					<input name="Billno1" type="text" value="<%
				if trim(rs1("Billno"))<>"" and not isnull(rs1("Billno")) then
					response.write trim(rs1("Billno"))
				end if
				%>" size="10" maxlength="9" onblur="CheckBillNoExist()" onkeydown="funTextControl(this);">
				<input name="OldBillNo" type="hidden" value="<%
				if trim(rs1("Billno"))<>"" and not isnull(rs1("Billno")) then
					response.write trim(rs1("Billno"))
					OldBillData="Billno="&trim(rs1("Billno"))
				else
					OldBillData="Billno="
				end if
				%>">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">保險證</div></td>
				<td <%
			if sys_City<>"台東縣" then
				response.write "colspan='3'"
			end if
				%>>
				    <input type="text" maxlength="1" size="3" value="<%
				if trim(rs1("Insurance"))<>"" and not isnull(rs1("Insurance")) then
					response.write trim(rs1("Insurance"))
					OldBillData=OldBillData&",Insurance="&trim(rs1("Insurance"))
				else
					OldBillData=OldBillData&",Insurance="
				end if
				%>" name="Insurance" onKeyUp="value=value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);">
					<div id="Layer111" style="position:absolute; width:220px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
					<span class="style9">0有出示/ 1未出示/ 2肇事且未出示/<br>3逾期或未保險/ 4肇事且逾期或未保險</span>
					</div>
				</td>
			<%if sys_City="台東縣" then%>
				<td bgcolor="#FFFFCC"><div align="right"><!-- <span class="style4">＊</span> -->違規人姓名</div></td>
				<td><input type="text" size="13" value="<%
				if trim(rs1("Driver"))<>"" and not isnull(rs1("Driver")) then
					response.write trim(rs1("Driver"))
					OldBillData=OldBillData&",Driver="&trim(rs1("Driver"))
				else
					OldBillData=OldBillData&",Driver="
				end if
				%>" name="DriverName" onkeydown="funTextControl(this);"></td> 
			<%end if%>
			</tr>
	<%if sys_City<>"高雄市" then%>
 			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規人證號</div></td>
				<td ><input type="text" size="10" value="<%
				if trim(rs1("Driverid"))<>"" and not isnull(rs1("Driverid")) then
					response.write trim(rs1("Driverid"))
					OldBillData=OldBillData&",Driverid="&trim(rs1("Driverid"))
				else
					OldBillData=OldBillData&",Driverid="
				end if
				%>" name="DriverPID" <%
		if sys_City="南投縣" then
			response.write "maxlength='10'"
		end if
				%> onBlur="FuncChkPID();" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">違規人出生日</div></td>
				<td <%
		if sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then 
				response.write "colspan=""3"""
		end if
				%>>
				<input type="text" size="10" maxlength="6" value="<%
				if trim(rs1("DriverBirth"))<>"" and not isnull(rs1("DriverBirth")) then
					response.write ginitdt(trim(rs1("DriverBirth")))
					OldBillData=OldBillData&",DriverBirth="&trim(rs1("DriverBirth"))
				else
					OldBillData=OldBillData&",DriverBirth="
				end if
				%>" name="DriverBrith" onkeyup="focusToDriverPID()" onkeydown="funTextControl(this);">
				</td>
		<%if sys_City="嘉義縣" or sys_City="宜蘭縣" or sys_City="台東縣" or sys_City="高雄縣" or sys_City="台南縣" or sys_City="台南市" or sys_City="嘉義市" then%>
				<td bgcolor="#FFFFCC" align="right"><span class="style4">＊</span>填單日期</td>
				<td>
					<input type="text" size="5" value="<%
				if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
					response.write ginitdt(trim(rs1("BillFillDate")))
					OldBillData=OldBillData&",BillFillDate="&trim(rs1("BillFillDate"))
				else
					OldBillData=OldBillData&",BillFillDate="
				end if
				%>" maxlength="7" name="BillFillDate" onBlur="getDealLineDate_Stop()" onkeydown="funTextControl(this);">
				</td>
		<%end if%>
			</tr>
	<%end if%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規車號</div></td>
				<td><input type="text" size="8" maxlength="8" value="<%
				if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
					response.write trim(rs1("CarNo"))
					OldBillData=OldBillData&",CarNo="&trim(rs1("CarNo"))
				else
					OldBillData=OldBillData&",CarNo="
				end if
				%>" name="CarNo" onBlur="getVIPCar();" onkeydown="funTextControl(this);">
				<input name="OldCarNo" type="hidden" value="<%
				if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
					response.write trim(rs1("CarNo"))
				end if
				%>">
			    <div id="Layer7" style="position:absolute; width:115px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold;"></div>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>簡式車種</div></td>
				<td colspan="3">
				<input type="text" maxlength="1" size="3" value="<%
				if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
					response.write trim(rs1("CarSimpleID"))
					OldBillData=OldBillData&",CarSimpleID="&trim(rs1("CarSimpleID"))
				else
					OldBillData=OldBillData&",CarSimpleID="
				end if
				%>" name="CarSimpleID" onBlur="getRuleAll();" onfocus="this.select();" onkeydown="funTextControl(this);">
				<font color="#ff000" size="2"> 1汽車/ 2拖車/ 3重機/ 4輕機/ 6 臨時車牌</font>
				</td>
			</tr>
	<%if sys_City="高雄市" then%>
 			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規人證號</div></td>
				<td ><input type="text" size="10" value="<%
				if trim(rs1("Driverid"))<>"" and not isnull(rs1("Driverid")) then
					response.write trim(rs1("Driverid"))
					OldBillData=OldBillData&",Driverid="&trim(rs1("Driverid"))
				else
					OldBillData=OldBillData&",Driverid="
				end if
				%>" name="DriverPID" onBlur="FuncChkPID();" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">違規人出生日</div></td>
				<td <%
		if sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then 
				response.write "colspan=""3"""
		end if
				%>>
				<input type="text" size="10" maxlength="6" value="<%
				if trim(rs1("DriverBirth"))<>"" and not isnull(rs1("DriverBirth")) then
					response.write ginitdt(trim(rs1("DriverBirth")))
					OldBillData=OldBillData&",DriverBirth="&trim(rs1("DriverBirth"))
				else
					OldBillData=OldBillData&",DriverBirth="
				end if
				%>" name="DriverBrith" onkeyup="focusToDriverPID()" onkeydown="funTextControl(this);">
				</td>
			</tr>
	<%end if%>
	<%if sys_City="高雄市" then%>
			<tr>
			  <td bgcolor="#FFFFCC"><div align="right">代保管物1</div></td>
			  <td><input type="text" size="5" value="<%
				itemTemp=""
				strItem="select * from BillFastenerDetail where BillSN="&trim(request("BillSN"))
				set rsItem=conn.execute(strItem)
				If Not rsItem.Bof Then rsItem.MoveFirst 
				While Not rsItem.Eof
					if itemTemp="" then
						itemTemp=trim(rsItem("FastenerTypeID"))
					else
						itemTemp=itemTemp&"&"&trim(rsItem("FastenerTypeID"))
					end if
				rsItem.MoveNext
				Wend
				rsItem.close
				set rsItem=nothing
				ItemVal=split(itemTemp,"&")
				if ubound(ItemVal)>=0 then
					response.write ItemVal(0)
				end if
				%>" name="Fastener1" onkeyup="getFastener1();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=1","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer8" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;">
                  <%
				if ubound(ItemVal)>=0 then
					FVal1=""
					strFVal="select Content from DCIcode where TypeID=6 and ID='"&ItemVal(0)&"'"
					set rsFVal1=conn.execute(strFVal)
					if not rsFVal1.eof then
						response.write trim(rsFVal1("Content"))
						FVal1=trim(rsFVal1("Content"))
					end if
					rsFVal1.close
					set rsFVal1=nothing
				end if
					%>
                </div>
                <input type="hidden" value="<%
					response.write FVal1
					%>" name="Fastener1Val"></td>
			  <td bgcolor="#FFFFCC"><div align="right">代保管物2</div></td>
			  <td>
                <input type="text" size="5" value="<%
				if ubound(ItemVal)>=1 then
					response.write ItemVal(1)
				end if
				%>" name="Fastener2" onkeyup="getFastener2();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=2","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer9" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;">
                  <%
				if ubound(ItemVal)>=1 then
					FVal2=""
					strFVal="select Content from DCIcode where TypeID=6 and ID='"&ItemVal(1)&"'"
					set rsFVal1=conn.execute(strFVal)
					if not rsFVal1.eof then
						response.write trim(rsFVal1("Content"))
						FVal2=trim(rsFVal1("Content"))
					end if
					rsFVal1.close
					set rsFVal1=nothing
				end if
					%>
                </div>
                <input type="hidden" value="<%=FVal2%>" name="Fastener2Val">
              </td>
			  <td bgcolor="#FFFFCC"><div align="right">代保管物3</div></td>
			  <td>
                <input type="text" size="5" value="<%
				if ubound(ItemVal)>=2 then
					response.write ItemVal(2)
				end if
				%>" name="Fastener3" onkeyup="getFastener3();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=3","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer10" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;">
                  <%
				if ubound(ItemVal)>=2 then
					FVal3=""
					strFVal="select Content from DCIcode where TypeID=6 and ID='"&ItemVal(2)&"'"
					set rsFVal1=conn.execute(strFVal)
					if not rsFVal1.eof then
						response.write trim(rsFVal1("Content"))
						FVal3=trim(rsFVal1("Content"))
					end if
					rsFVal1.close
					set rsFVal1=nothing
				end if
					%>
                </div>
                <input type="hidden" value="<%=FVal3%>" name="Fastener3Val">
              </td>
			</tr>
	<%end if%>
			<tr>
		<%if sys_City="雲林縣" then%>
				<td bgcolor="#FFFFCC" align="right">輔助車種</td>
				<td>
				<input type="text" maxlength="2" size="2" value="<%
				if trim(rs1("CarAddID"))<>"" and not isnull(rs1("CarAddID")) then
					response.write trim(rs1("CarAddID"))
					OldBillData=OldBillData&",CarAddID="&trim(rs1("CarAddID"))
				else
					OldBillData=OldBillData&",CarAddID="
				end if
				%>" name="CarAddID" onBlur="getAddID();" onkeydown="funTextControl(this);">
				<div id="Layer96543" style="position:absolute ; width:220px; height:30px; z-index:0;">
				<span class="style10">1大貨/ 2大客/ 3砂石/ 4土方/ 5動力/ 6貨櫃/<br>7大型重機 /8拖吊 /9(550cc)重機 /10計程車/<br>11危險物品</span>		
				</div>
				</td>
		<%end if%>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規日期</div></td>
				<td>
				<input type="text" size="10" value="<%
				if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
					response.write ginitdt(trim(rs1("IllegalDate")))
				end if
				%>" maxlength="6" name="IllegalDate" onfocus="this.select()" onBlur="getDealLineDate()" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規時間</div></td>
				<td <%if sys_City<>"雲林縣" then response.write "colspan=""3"""%>>
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
					OldBillData=OldBillData&",IllegalDate="&year(rs1("IllegalDate"))&"/"&month(rs1("IllegalDate"))&"/"&Day(rs1("IllegalDate"))&" "&Hour(rs1("IllegalDate"))&":"&Minute(rs1("IllegalDate"))&":0"
				else
					OldBillData=OldBillData&",IllegalDate="
				end if
				%>" maxlength="4" name="IllegalTime" onBlur="value=value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);">
				</td>
			</tr>
<%if sys_City="南投縣" then%>
			<tr>
			  <td bgcolor="#FFFFCC"><div align="right">代保管物1</div></td>
			  <td><input type="text" size="5" value="<%
				itemTemp=""
				strItem="select * from BillFastenerDetail where BillSN="&trim(request("BillSN"))
				set rsItem=conn.execute(strItem)
				If Not rsItem.Bof Then rsItem.MoveFirst 
				While Not rsItem.Eof
					if itemTemp="" then
						itemTemp=trim(rsItem("FastenerTypeID"))
					else
						itemTemp=itemTemp&"&"&trim(rsItem("FastenerTypeID"))
					end if
				rsItem.MoveNext
				Wend
				rsItem.close
				set rsItem=nothing
				ItemVal=split(itemTemp,"&")
				if ubound(ItemVal)>=0 then
					response.write ItemVal(0)
				end if
				%>" name="Fastener1" onkeyup="getFastener1();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=1","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer8" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;">
                  <%
				if ubound(ItemVal)>=0 then
					FVal1=""
					strFVal="select Content from DCIcode where TypeID=6 and ID='"&ItemVal(0)&"'"
					set rsFVal1=conn.execute(strFVal)
					if not rsFVal1.eof then
						response.write trim(rsFVal1("Content"))
						FVal1=trim(rsFVal1("Content"))
					end if
					rsFVal1.close
					set rsFVal1=nothing
				end if
					%>
                </div>
                <input type="hidden" value="<%
					response.write FVal1
					%>" name="Fastener1Val"></td>
			  <td bgcolor="#FFFFCC"><div align="right">代保管物2</div></td>
			  <td>
                <input type="text" size="5" value="<%
				if ubound(ItemVal)>=1 then
					response.write ItemVal(1)
				end if
				%>" name="Fastener2" onkeyup="getFastener2();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=2","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer9" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;">
                  <%
				if ubound(ItemVal)>=1 then
					FVal2=""
					strFVal="select Content from DCIcode where TypeID=6 and ID='"&ItemVal(1)&"'"
					set rsFVal1=conn.execute(strFVal)
					if not rsFVal1.eof then
						response.write trim(rsFVal1("Content"))
						FVal2=trim(rsFVal1("Content"))
					end if
					rsFVal1.close
					set rsFVal1=nothing
				end if
					%>
                </div>
                <input type="hidden" value="<%=FVal2%>" name="Fastener2Val">
              </td>
			  <td bgcolor="#FFFFCC"><div align="right">代保管物3</div></td>
			  <td>
                <input type="text" size="5" value="<%
				if ubound(ItemVal)>=2 then
					response.write ItemVal(2)
				end if
				%>" name="Fastener3" onkeyup="getFastener3();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=3","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer10" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;">
                  <%
				if ubound(ItemVal)>=2 then
					FVal3=""
					strFVal="select Content from DCIcode where TypeID=6 and ID='"&ItemVal(2)&"'"
					set rsFVal1=conn.execute(strFVal)
					if not rsFVal1.eof then
						response.write trim(rsFVal1("Content"))
						FVal3=trim(rsFVal1("Content"))
					end if
					rsFVal1.close
					set rsFVal1=nothing
				end if
					%>
                </div>
                <input type="hidden" value="<%=FVal3%>" name="Fastener3Val">
              </td>
			</tr>
<%end if%>
<%if sys_City="嘉義市" then%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
					response.write trim(rs1("RuleSpeed"))
					OldBillData=OldBillData&",RuleSpeed="&trim(rs1("RuleSpeed"))
				else
					OldBillData=OldBillData&",RuleSpeed="
				end if
				%>" name="RuleSpeed" onkeyup="RuleSpeedforLaw()" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style6">實際車速、車重</span></div></td>
				<td colspan='3'>
					<input type="text" size="10" value="<%
				if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
					response.write trim(rs1("IllegalSpeed"))
					OldBillData=OldBillData&",IllegalSpeed="&trim(rs1("IllegalSpeed"))
				else
					OldBillData=OldBillData&",IllegalSpeed="
				end if
				%>" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);">
				<div id="LayerWeight" style="position:absolute ; width:90px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
				</td>
				
			</tr>
<%end if%>
<%if sys_City<>"嘉義市" then %>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">違規地點代碼</div></td>
				<td >
					<input type="text" size="8" value="<%
				if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
					response.write trim(rs1("IllegalAddressID"))
					OldBillData=OldBillData&",IllegalAddressID="&trim(rs1("IllegalAddressID"))
				else
					OldBillData=OldBillData&",IllegalAddressID="
				end if
				%>" name="IllegalAddressID" onkeyup="getillStreet();" onkeydown="funTextControl(this);" onblur="funGetSpeedRule()">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage_Illaddr","left=0,top=0,location=0,width=700,height=455,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規地點</div></td>
				<td colspan="3">
					<input type="text" size="40" value="<%
				if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
					response.write trim(rs1("IllegalAddress"))
					OldBillData=OldBillData&",IllegalAddress="&trim(rs1("IllegalAddress"))
				else
					OldBillData=OldBillData&",IllegalAddress="
				end if
				%>" name="IllegalAddress" onkeydown="funTextControl(this);" onblur="funGetSpeedRule()">
				<input type="checkbox" name="chkHighRoad" value="1" onclick="setIllegalRule()" <%
					if Left(trim(rs1("Rule1")),2)="33" then
						response.write "checked"
					elseif trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
						if Left(trim(rs1("Rule2")),2)="33" then
							response.write "checked"
						elseif trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
							if Left(trim(rs1("Rule3")),2)="33" then
								response.write "checked"
							end if
						end if
					end if
					%>><span class="style1">快速道路</span>
				</td>
			</tr>
<%end if%>
<%if sys_City="彰化縣" or sys_City="雲林縣" or sys_City="嘉義縣" or sys_City="宜蘭縣" or sys_City="台東縣" or sys_City="高雄縣" or sys_City="台南縣" or sys_City="台南市" then%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
					response.write trim(rs1("RuleSpeed"))
					OldBillData=OldBillData&",RuleSpeed="&trim(rs1("RuleSpeed"))
				else
					OldBillData=OldBillData&",RuleSpeed="
				end if
				%>" name="RuleSpeed" onkeyup="RuleSpeedforLaw()" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style6">實際車速、車重</span></div></td>
				<td colspan='3'>
					<input type="text" size="10" value="<%
				if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
					response.write trim(rs1("IllegalSpeed"))
					OldBillData=OldBillData&",IllegalSpeed="&trim(rs1("IllegalSpeed"))
				else
					OldBillData=OldBillData&",IllegalSpeed="
				end if
				%>" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);">
				<div id="LayerWeight" style="position:absolute ; width:90px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
				</td>
				
			</tr>
<%end if%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規法條一</div></td>
				<td colspan="5">
					<input type="text" maxlength="8" size="10" value="<%
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					response.write trim(rs1("Rule1"))
					OldBillData=OldBillData&",Rule1="&trim(rs1("Rule1"))
				else
					OldBillData=OldBillData&",Rule1="
				end if
				%>" name="Rule1" onKeyUp="getRuleData1();" onchange="DelSpace1();" onblur="AutoKeyCarNo()" onkeydown="funTextControl(this);">
					
					<input type="text" name="ForFeit1" value="<%
				if trim(rs1("ForFeit1"))<>"" and not isnull(rs1("ForFeit1")) then
					response.write trim(rs1("ForFeit1"))
				end if
				%>" size="5">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=trim(rs1("RuleVer"))%>","WebPage_Law","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes")' alt="查詢法條">
					<!-- <img src="../Image/BillLawPlusButton.jpg" width="25" height="23" onclick="Add_LawPlus()" alt="附加說明"> -->
					<div id="Layer1" style="position:absolute ; width:609px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					strCarImple=""
					if left(trim(rs1("Rule1")),4)="2110" or trim(rs1("Rule1"))="4310102" or trim(rs1("Rule1"))="4310103" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple=" and CarSimpleID in ('3','0')"
						else
							strCarImple=""
						end if
					end if
					strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rsR1("IllegalRule"))
					end if
					rsR1.close
					set rsR1=nothing
				end if
'				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
'					response.write "("&trim(rs1("Rule4"))&")"
'				end if
				%></div>
				</td>
				
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">違規法條二</div></td>
				<td colspan="5">
					<input type="text" maxlength="8" size="10" value="<%
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					response.write trim(rs1("Rule2"))
					OldBillData=OldBillData&",Rule2="&trim(rs1("Rule2"))
				else
					OldBillData=OldBillData&",Rule2="
				end if
				%>" name="Rule2" onKeyUp="getRuleData2();" onchange="DelSpace2();" onblur="AutoKeyCarNo()" onkeydown="funTextControl(this);">
					<input type="text" name="ForFeit2" value="<%
				if trim(rs1("ForFeit2"))<>"" and not isnull(rs1("ForFeit2")) then
					response.write trim(rs1("ForFeit2"))
				end if
				%>" size="5">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=trim(rs1("RuleVer"))%>","WebPage_Law","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer2" style="position:absolute ; width:609px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					strCarImple=""
					if left(trim(rs1("Rule2")),4)="2110" or trim(rs1("Rule2"))="4310102" or trim(rs1("Rule2"))="4310103" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple=" and CarSimpleID in ('3','0')"
						else
							strCarImple=""
						end if
					end if
					strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rsR1("IllegalRule"))
					end if
					rsR1.close
					set rsR1=nothing
				end if
				%></div>
					
				<img src="space.gif" width="605" height="2">
<%if trim(rs1("Rule3"))="" or isnull(rs1("Rule3")) then%>
				<img src="../Image/Law3.jpg" width="45" height="25" onclick='InsertLaw()' alt="違規法條三">
<%end if%>
				</td>
				
			</tr>
<%if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then%>
			<tr>
				<td bgcolor="#FFFFCC" id="TDLaw1" align="right"><div align="right">違規法條三</div></td>
				<td colspan="5" id="TDLaw2">
					<input type="text" maxlength="8" size="10" value="<%
				if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
					response.write trim(rs1("Rule3"))
					OldBillData=OldBillData&",Rule3="&trim(rs1("Rule3"))
				else
					OldBillData=OldBillData&",Rule3="
				end if
				%>" name="Rule3" onKeyUp="getRuleData3();" onchange="DelSpace3();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=3&RuleVer=<%=trim(rs1("RuleVer"))%>","WebPage_Law","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer3" style="position:absolute ; width:609px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
					strCarImple=""
					if left(trim(rs1("Rule3")),4)="2110" or trim(rs1("Rule3"))="4310102" or trim(rs1("Rule3"))="4310103" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple=" and CarSimpleID in ('3','0')"
						else
							strCarImple=""
						end if
					end if

					strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule3"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rsR1("IllegalRule"))
					end if
					rsR1.close
					set rsR1=nothing
				end if
				%></div>
					<input type="hidden" name="ForFeit3" value="<%
				if trim(rs1("ForFeit3"))<>"" and not isnull(rs1("ForFeit3")) then
					response.write trim(rs1("ForFeit3"))
				end if
				%>">
				<img src="space.gif" width="605" height="2">
<%if trim(rs1("Rule4"))="" or isnull(rs1("Rule4")) then%>
				<img src="../Image/Law4.jpg" width="45" height="25" onclick='InsertLaw2()' alt="違規法條四">
<%end if%>
				</td>
				
			</tr>
<%else%>
			<tr>
				<td bgcolor="#FFFFCC" id="TDLaw1" align="right"></td>
				<td colspan="5" id="TDLaw2"></td>
			</tr>
<%end if%>
<%if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then%>
			<tr>
				<td bgcolor="#FFFFCC" id="TDLaw3" align="right"><div align="right">違規法條四</div></td>
				<td colspan="5" id="TDLaw4">
					<input type="text" maxlength="8" size="10" value="<%
				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					response.write trim(rs1("Rule4"))
					OldBillData=OldBillData&",Rule4="&trim(rs1("Rule4"))
				else
					OldBillData=OldBillData&",Rule4="
				end if
				%>" name="Rule4" onKeyUp="getRuleData4();" onchange="DelSpace4();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=4&RuleVer=<%=trim(rs1("RuleVer"))%>","WebPage_Law","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer4" style="position:absolute ; width:609px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					strCarImple=""
					if left(trim(rs1("Rule4")),4)="2110" or trim(rs1("Rule4"))="4310102" or trim(rs1("Rule4"))="4310103" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple=" and CarSimpleID in ('3','0')"
						else
							strCarImple=""
						end if
					end if

					strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule4"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rsR1("IllegalRule"))
					end if
					rsR1.close
					set rsR1=nothing
				end if
				%></div>
					<input type="hidden" name="ForFeit4" value="<%
				if trim(rs1("ForFeit4"))<>"" and not isnull(rs1("ForFeit4")) then
					response.write trim(rs1("ForFeit4"))
				end if
				%>">
				</td>
				
			</tr>
<%else%>
			<tr>
				<td bgcolor="#FFFFCC" id="TDLaw3" align="right"></td>
				<td colspan="5" id="TDLaw4"></td>
			</tr>
<%end if%>
<%if sys_City="嘉義市" then %>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">違規地點代碼</div></td>
				<td >
					<input type="text" size="8" value="<%
				if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
					response.write trim(rs1("IllegalAddressID"))
					OldBillData=OldBillData&",IllegalAddressID="&trim(rs1("IllegalAddressID"))
				else
					OldBillData=OldBillData&",IllegalAddressID="
				end if
				%>" name="IllegalAddressID" onkeyup="getillStreet();" onkeydown="funTextControl(this);" onblur="funGetSpeedRule()">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage_Illaddr","left=0,top=0,location=0,width=700,height=455,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規地點</div></td>
				<td colspan="3">
					<input type="text" size="40" value="<%
				if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
					response.write trim(rs1("IllegalAddress"))
					OldBillData=OldBillData&",IllegalAddress="&trim(rs1("IllegalAddress"))
				else
					OldBillData=OldBillData&",IllegalAddress="
				end if
				%>" name="IllegalAddress" onkeydown="funTextControl(this);" onblur="funGetSpeedRule()">
				<input type="checkbox" name="chkHighRoad" value="1" onclick="setIllegalRule()" <%
					if Left(trim(rs1("Rule1")),2)="33" then
						response.write "checked"
					elseif trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
						if Left(trim(rs1("Rule2")),2)="33" then
							response.write "checked"
						elseif trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
							if Left(trim(rs1("Rule3")),2)="33" then
								response.write "checked"
							end if
						end if
					end if
					%>><span class="style1">快速道路</span>
				</td>
			</tr>
<%end if%>
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
					response.write trim(rs1("RuleSpeed"))
					OldBillData=OldBillData&",RuleSpeed="&trim(rs1("RuleSpeed"))
				else
					OldBillData=OldBillData&",RuleSpeed="
				end if
				%>" name="RuleSpeed" onkeyup="RuleSpeedforLaw()" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style6">實際車速、車重</span></div></td>
				<td colspan='3'>
					<input type="text" size="10" value="<%
				if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
					response.write trim(rs1("IllegalSpeed"))
					OldBillData=OldBillData&",IllegalSpeed="&trim(rs1("IllegalSpeed"))
				else
					OldBillData=OldBillData&",IllegalSpeed="
				end if
				%>" name="IllegalSpeed" onkeyup="IllegalSpeedforLaw()" onkeydown="funTextControl(this);">
				<div id="LayerWeight" style="position:absolute ; width:90px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
				</td>
				
			</tr>
<%end if%>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>應到案日期</div></td>
				<td>
					<input type="text" size="10" value="<%
				if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
					response.write ginitdt(trim(rs1("DealLineDate")))
					OldBillData=OldBillData&",DealLineDate="&trim(rs1("DealLineDate"))
				else
					OldBillData=OldBillData&",DealLineDate="
				end if
				%>" maxlength="6" name="DealLineDate" onBlur="value=value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);">
				</td>


				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>應到案處所</div></td>
				<td>
					<input type="text" size="5" value="<%
				if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
					response.write trim(rs1("MemberStation"))
					OldBillData=OldBillData&",MemberStation="&trim(rs1("MemberStation"))
				else
					OldBillData=OldBillData&",MemberStation="
				end if
				%>" name="MemberStation" onKeyup="getStation();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Station.asp","WebPage_MemStation_1","left=0,top=0,location=0,width=760,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer5" style="position:absolute ; width:120px; height:30px; z-index:0;  border: 1px none #000000;"><%
				if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
					strS="select DciStationName from Station where StationID='"&trim(rs1("MemberStation"))&"'"
					set rsS=conn.execute(strS)
					if not rsS.eof then
						response.write trim(rsS("DciStationName"))
					end if
					rsS.close
					set rsS=nothing
				end if
				%></div>
					</span>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發人代碼1</div></td>
		  		<td>
					<input type="text" size="5" value="<%
				if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
					strMem1="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID1"))
					set rsMem1=conn.execute(strMem1)
					if not rsMem1.eof then
						response.write trim(rsMem1("LoginID"))
					end if
					rsMem1.close
					set rsMem1=nothing
					OldBillData=OldBillData&",BillMemID1="&trim(rs1("BillMemID1"))
				else
					OldBillData=OldBillData&",BillMemID1="
				end if
				%>" name="BillMem1" onkeyup="getBillMemID1();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage_Mem","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer12" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"><%
				if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
					response.write trim(rs1("BillMem1"))
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
				
				<td bgcolor="#FFFFCC" width="12%"><div align="right">舉發人代碼2</div></td>
				<td width="24%">
					<input type="text" size="5" value="<%
				if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
					strMem2="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID2"))
					set rsMem2=conn.execute(strMem2)
					if not rsMem2.eof then
						response.write trim(rsMem2("LoginID"))
					end if
					rsMem2.close
					set rsMem2=nothing
					OldBillData=OldBillData&",BillMemID2="&trim(rs1("BillMemID2"))
				else
					OldBillData=OldBillData&",BillMemID2="
				end if
				%>" name="BillMem2" onkeyup="getBillMemID2();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage_Mem","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer13" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"><%
				if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
					response.write trim(rs1("BillMem2"))
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
				<td bgcolor="#FFFFCC" width="12%"><div align="right">舉發人代碼3</div></td>
				<td width="22%">
					<input type="text" size="5" value="<%
				if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
					strMem3="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID3"))
					set rsMem3=conn.execute(strMem3)
					if not rsMem3.eof then
						response.write trim(rsMem3("LoginID"))
					end if
					rsMem3.close
					set rsMem3=nothing
					OldBillData=OldBillData&",BillMemID3="&trim(rs1("BillMemID3"))
				else
					OldBillData=OldBillData&",BillMemID3="
				end if
				%>" name="BillMem3" onkeyup="getBillMemID3();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=3","WebPage_Mem","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer14" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"><%
				if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
					response.write trim(rs1("BillMem3"))
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
				<td bgcolor="#FFFFCC" width="12%"><div align="right">舉發人代碼4</div></td>
				<td width="18%">
					<input type="text" size="5" value="<%
				if trim(rs1("BillMemID4"))<>"" and not isnull(rs1("BillMemID4")) then
					strMem4="select LoginID from MemberData where MemberID="&trim(rs1("BillMemID4"))
					set rsMem4=conn.execute(strMem4)
					if not rsMem4.eof then
						response.write trim(rsMem4("LoginID"))
					end if
					rsMem4.close
					set rsMem4=nothing
					OldBillData=OldBillData&",BillMemID4="&trim(rs1("BillMemID4"))
				else
					OldBillData=OldBillData&",BillMemID4="
				end if
				%>" name="BillMem4" onkeyup="getBillMemID4();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=4","WebPage_Mem","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer17" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"><%
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
<%if sys_City<>"南投縣" and sys_City<>"高雄市" then%>
			<tr>
			  <td bgcolor="#FFFFCC"><div align="right">代保管物1</div></td>
			  <td><input type="text" size="5" value="<%
				itemTemp=""
				strItem="select * from BillFastenerDetail where BillSN="&trim(request("BillSN"))
				set rsItem=conn.execute(strItem)
				If Not rsItem.Bof Then rsItem.MoveFirst 
				While Not rsItem.Eof
					if itemTemp="" then
						itemTemp=trim(rsItem("FastenerTypeID"))
					else
						itemTemp=itemTemp&"&"&trim(rsItem("FastenerTypeID"))
					end if
				rsItem.MoveNext
				Wend
				rsItem.close
				set rsItem=nothing
				ItemVal=split(itemTemp,"&")
				if ubound(ItemVal)>=0 then
					response.write ItemVal(0)
				end if
				%>" name="Fastener1" onkeyup="getFastener1();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=1","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer8" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;">
                  <%
				if ubound(ItemVal)>=0 then
					FVal1=""
					strFVal="select Content from DCIcode where TypeID=6 and ID='"&ItemVal(0)&"'"
					set rsFVal1=conn.execute(strFVal)
					if not rsFVal1.eof then
						response.write trim(rsFVal1("Content"))
						FVal1=trim(rsFVal1("Content"))
					end if
					rsFVal1.close
					set rsFVal1=nothing
				end if
					%>
                </div>
                <input type="hidden" value="<%
					response.write FVal1
					%>" name="Fastener1Val"></td>
			  <td bgcolor="#FFFFCC"><div align="right">代保管物2</div></td>
			  <td>
                <input type="text" size="5" value="<%
				if ubound(ItemVal)>=1 then
					response.write ItemVal(1)
				end if
				%>" name="Fastener2" onkeyup="getFastener2();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=2","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer9" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;">
                  <%
				if ubound(ItemVal)>=1 then
					FVal2=""
					strFVal="select Content from DCIcode where TypeID=6 and ID='"&ItemVal(1)&"'"
					set rsFVal1=conn.execute(strFVal)
					if not rsFVal1.eof then
						response.write trim(rsFVal1("Content"))
						FVal2=trim(rsFVal1("Content"))
					end if
					rsFVal1.close
					set rsFVal1=nothing
				end if
					%>
                </div>
                <input type="hidden" value="<%=FVal2%>" name="Fastener2Val">
              </td>
			  <td bgcolor="#FFFFCC"><div align="right">代保管物3</div></td>
			  <td>
                <input type="text" size="5" value="<%
				if ubound(ItemVal)>=2 then
					response.write ItemVal(2)
				end if
				%>" name="Fastener3" onkeyup="getFastener3();" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=3","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                <div id="Layer10" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;">
                  <%
				if ubound(ItemVal)>=2 then
					FVal3=""
					strFVal="select Content from DCIcode where TypeID=6 and ID='"&ItemVal(2)&"'"
					set rsFVal1=conn.execute(strFVal)
					if not rsFVal1.eof then
						response.write trim(rsFVal1("Content"))
						FVal3=trim(rsFVal1("Content"))
					end if
					rsFVal1.close
					set rsFVal1=nothing
				end if
					%>
                </div>
                <input type="hidden" value="<%=FVal3%>" name="Fastener3Val">
              </td>
			</tr>
<%end if%>
			<tr>

				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發單位</div></td>
				<td>
					<input type="text" size="5" value="<%
				if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
					response.write trim(rs1("BillUnitID"))
					OldBillData=OldBillData&",BillUnitID="&trim(rs1("BillUnitID"))
				else
					OldBillData=OldBillData&",BillUnitID="
				end if
				%>" name="BillUnitID" onKeyUp="getUnit();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=U","WebPage_memUnit","left=0,top=0,location=0,width=700,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer6" style="position:absolute ; width:120px; height:30px; z-index:0;  border: 1px none #000000;"><%
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
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>簽收狀況</div></td>
				<td <%
		if sys_City="嘉義縣" or sys_City="宜蘭縣" or sys_City="台東縣" or sys_City="高雄縣" or sys_City="台南縣" or sys_City="台南市" or sys_City="嘉義市" then 
				response.write "colspan=""3"""
		end if
				%>>
					<input type="text" size="5" value="<%
				if trim(rs1("SignType"))<>"" and not isnull(rs1("SignType")) then
					if trim(rs1("SignType"))="A" then
						response.write trim(rs1("SignType"))
					else
						strUserSign="select * from BillUserSignDate where BillSn="&trim(rs1("SN"))
						set rsUserSign=conn.execute(strUserSign)
						if not rsUserSign.eof then
							response.write trim(rsUserSign("SignStateID"))
						else
							response.write "U"
						end if
						rsUserSign.close
						set rsUserSign=nothing
					end if
					OldBillData=OldBillData&",SignType="&trim(rs1("SignType"))
				else
					response.write "5"
					OldBillData=OldBillData&",SignType="
				end if
				%>" maxlength="1" name="SignType" onBlur="funcSignType();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<div id="Layer96hg543" style="position:absolute; width:145px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
					<span class="style10">
					A簽收/ U拒簽收/ 2拒簽已收/ 3已簽拒收/ 5補開單
					</span>		
					</div>
				</td>	
			<%if sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then%>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>填單日期</div></td>
				<td>
					<input type="text" size="5" value="<%
				if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
					response.write ginitdt(trim(rs1("BillFillDate")))
				end if
				%>" maxlength="7" name="BillFillDate" onkeydown="funTextControl(this);">
				</td>
			<%end if%>
			</tr>
			<tr>
		<%if sys_City<>"雲林縣" then%>
				<td bgcolor="#FFFFCC" align="right">輔助車種</td>
				<td>
				<input type="text" maxlength="2" size="2" value="<%
				if trim(rs1("CarAddID"))<>"" and not isnull(rs1("CarAddID")) then
					response.write trim(rs1("CarAddID"))
					OldBillData=OldBillData&",CarAddID="&trim(rs1("CarAddID"))
				else
					OldBillData=OldBillData&",CarAddID="
				end if
				%>" name="CarAddID" onBlur="getAddID();" onkeydown="funTextControl(this);">
				<div id="Layer96543" style="position:absolute ; width:220px; height:30px; z-index:0;">
				<span class="style10">1大貨/ 2大客/ 3砂石/ 4土方/ 5動力/ 6貨櫃/<br>7大型重機 /8拖吊 /9(550cc)重機 /10計程車/<br>11危險物品</span>		
				</div>
				</td>
		<%end if%>
				<td bgcolor="#FFFFCC"><div align="right">是否郵寄</div></td>
				<td>
					<input type="radio" name="IsMail" value="1" <%
					if trim(rs1("EquipMentID"))<>"" and not isnull(rs1("EquipMentID")) then
						OldBillData=OldBillData&",EquipMentID="&trim(rs1("EquipMentID"))
					else
						OldBillData=OldBillData&",EquipMentID="
					end if
					if trim(rs1("EquipmentID"))<>"" and not isnull(rs1("EquipmentID")) then
						if trim(rs1("EquipmentID"))="1" then
							response.write "checked"
						end if
					end if
					%>>是
					<input type="radio" name="IsMail" value="-1" <%
					if trim(rs1("EquipmentID"))<>"" and not isnull(rs1("EquipmentID")) then
						if trim(rs1("EquipmentID"))="-1" then
							response.write "checked"
						end if
					else
						response.write "checked"
					end if
					%>>否
				</td>
				<td bgcolor="#FFFFCC"><div align="right">專案代碼</div></td>
				<td <%if sys_City="雲林縣" then response.write "colspan=""3"""%>><input type="text" size="5" value="<%
				if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then
					response.write trim(rs1("ProjectID"))
					OldBillData=OldBillData&",ProjectID="&trim(rs1("ProjectID"))
				else
					OldBillData=OldBillData&",ProjectID="
				end if
				%>" name="ProjectID" onkeyup="ProjectF5()" onkeydown="funTextControl(this);">
                  <img src="../Image/BillkeyInButton.jpg" width="25" height="23" onClick='window.open("Query_Project.asp","WebPage_project","left=0,top=0,location=0,width=700,height=455,resizable=yes,scrollbars=yes")'>
				  <span class="style5">
				  <div id="Layer001" style="position:absolute ; width:180px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"><%
				if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then
					strProject="select Name from Project where ProjectID='"&trim(rs1("ProjectID"))&"'"
					set rsProject=conn.execute(strProject)
					if not rsProject.eof then
						response.write trim(rsProject("Name"))
					end if
					rsProject.close
					set rsProject=nothing
				end if
					%></div>
					</span>
				</td>
				
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">備註</td>
				<td>
					<input type="text" size="20" value="<%
				if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
					response.write trim(rs1("Note"))
					OldBillData=OldBillData&",Note="&trim(rs1("Note"))
				else
					OldBillData=OldBillData&",Note="
				end if
				%>" name="Note" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">交通事故案號</div></td>
				<td>
					<input type="text" size="16" name="TrafficAccidentNo" Value="<%
				if trim(rs1("TrafficAccidentNo"))<>"" and not isnull(rs1("TrafficAccidentNo")) then
					response.write trim(rs1("TrafficAccidentNo"))
					OldBillData=OldBillData&",TrafficAccidentNo="&trim(rs1("TrafficAccidentNo"))
				else
					OldBillData=OldBillData&",TrafficAccidentNo="
				end if
				%>" onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">交通事故種類</div></td>
				<td>
					<input type="text" maxlength="1" size="5" name="TrafficAccidentType" Value="<%
				if trim(rs1("TrafficAccidentType"))<>"" and not isnull(rs1("TrafficAccidentType")) then
					response.write trim(rs1("TrafficAccidentType"))
					OldBillData=OldBillData&",TrafficAccidentType="&trim(rs1("TrafficAccidentType"))
				else
					OldBillData=OldBillData&",TrafficAccidentType="
				end if
				%>" onBlur="chkTrafficAccidentType();" onkeydown="funTextControl(this);">
					<font color="#ff000" size="2"> 1 / 2 / 3</font>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" align="center" colspan="6">
					<input type="button" value="儲 存 <%
					if sys_City="台東縣" then
						response.write "F9"
					else
						response.write "F2"
					end if
					%>" onclick="InsertBillVase();" <%
				'1:查詢 ,2:新增 ,3:修改 ,4:刪除
				if trim(rs1("RecordMemberID"))<>trim(session("User_ID")) then
					if CheckPermission(236,3)=false and CheckPermission(234,3)=false then
						response.write "disabled"
					end if
				end if
					%> class="btn1">
					
                    <span class="style3"><img src="/image/space.gif" height="8"></span>
					<input type="button" value="修改監理站回傳資料" onclick="UpdateDciData();" class="btn1">
                    <span class="style3"><img src="/image/space.gif" height="8"></span>

                    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉" class="btn1">
					<input type="hidden" value="<%=trim(rs1("RuleVer"))%>" name="RuleVerSion">
					<input type="hidden" value="" name="kinds">
					<input type="hidden" value="<%=OldBillData%>" name="OldBillData">
					<input type="hidden" value="<%=trim(request("BillSN"))%>" name="BillSN">
					<br>
		
				<!-- 告發類別 -->
				<input type="hidden" size="3" maxlength="1" value="<%
				if trim(rs1("BillTypeID"))<>"" and not isnull(rs1("BillTypeID")) then
					response.write trim(rs1("BillTypeID"))
				end if
				%>" name="BillType">
				<!-- 違規人性別 -->
				<input type="hidden" name="DriverSex" value="<%=trim(rs1("DriverSex"))%>">
				<!-- 附加說明 -->
				<!-- <input type="hidden" name="Rule4" value=" --><%'=trim(rs1("Rule4"))%><!-- "> -->
				</td>
			</tr>
		</table>		
	</form>

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
var TodayDate=<%=ginitdt(date)%>;
<%if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then%>
	<%if sys_City="嘉義縣" or sys_City="宜蘭縣" or sys_City="高雄縣" or sys_City="台南縣" or sys_City="台南市" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||Rule4||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||Rule4||IllegalAddressID,IllegalAddress||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="台東縣" then %>
	MoveTextVar("Billno1,Insurance,DriverName||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||Rule4||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="彰化縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||Rule4||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||CarAddID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||Rule4||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="南投縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||Fastener1,Fastener2,Fastener3||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||Rule4||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="高雄市" then %>
	MoveTextVar("Billno1,Insurance||CarNo,CarSimpleID||DriverPID,DriverBrith||Fastener1,Fastener2,Fastener3||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||Rule4||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%else%>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||Rule4||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%end if%>
<%elseif trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then%>
	<%if sys_City="嘉義縣" or sys_City="宜蘭縣" or sys_City="高雄縣" or sys_City="台南縣" or sys_City="台南市" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||IllegalAddressID,IllegalAddress||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="台東縣" then %>
	MoveTextVar("Billno1,Insurance,DriverName||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="彰化縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||CarAddID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="南投縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||Fastener1,Fastener2,Fastener3||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="高雄市" then %>
	MoveTextVar("Billno1,Insurance||CarNo,CarSimpleID||DriverPID,DriverBrith||Fastener1,Fastener2,Fastener3||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%else%>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%end if%>
<%else%>
	<%if sys_City="嘉義縣" or sys_City="宜蘭縣" or sys_City="高雄縣" or sys_City="台南縣" or sys_City="台南市" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||IllegalAddressID,IllegalAddress||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="台東縣" then %>
	MoveTextVar("Billno1,Insurance,DriverName||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="彰化縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||CarAddID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="南投縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||Fastener1,Fastener2,Fastener3||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="高雄市" then %>
	MoveTextVar("Billno1,Insurance||CarNo,CarSimpleID||DriverPID,DriverBrith||Fastener1,Fastener2,Fastener3||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%else%>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%end if%>
<%end if%>
//修改告發單
function InsertBillVase(){
	var error=0;
	var errorString="";
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value;
	ReadBillNo=myForm.Billno1.value.replace(/[\s　]+/g, "");
	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");
	myForm.DriverPID.value=myForm.DriverPID.value.replace(/[\s　]+/g, "");
	if (myForm.Billno1.value=="" && myForm.BillType.value!="2"){
		error=error+1;
		errorString=error+"：請輸入單號。";
	}else if(ReadBillNo.length!=9){     
		error=error+1;
		errorString=error+"：單號不足九碼。";
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
	if (myForm.DriverBrith.value!=""){
		if(!dateCheck( myForm.DriverBrith.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規人出生日期輸入錯誤。";	
		}	
	}
	/*
	if (myForm.DriverName.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規人姓名。";
	}
	
	if (myForm.DriverPID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規人身份證號。";
	}
	*/
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
	}else if (!ChkIllegalDate(myForm.IllegalDate.value) && myForm.TrafficAccidentNo.value=="" && myForm.CaseInByMem.checked==false){
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
	if (TDLawNum==2 && myForm.Rule4.value!=""){
		if(myForm.Rule1.value==myForm.Rule4.value){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條一與違規法條四重複。";
		}
		if (myForm.Rule2.value==myForm.Rule4.value){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二與違規法條四重複。";
		}
		if (myForm.Rule3.value==myForm.Rule4.value){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條三與違規法條四重複。";
		}
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
	if (TDLawErrorLog4==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條四輸入錯誤。";
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
	}else if (!ChkIllegalDate(myForm.BillFillDate.value) && myForm.TrafficAccidentNo.value=="" && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期已超過三個月。";
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
		if(myForm.BillType.value=="1"){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入應到案日期。";
		}
	}else if (!dateCheck( myForm.DealLineDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if (!ChkIllegalDate(myForm.DealLineDate.value) && myForm.TrafficAccidentNo.value=="" && myForm.CaseInByMem.checked==false){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期已超過三個月。";
	}
	if (myForm.BillUnitID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入舉發單位代號。";
	}else if (TDUnitErrorLog==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發單位代號輸入錯誤。";
	}
	if (myForm.SignType.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入簽收狀況。";
	}
	if (myForm.BillMem1.value==""){
		    error=error+1;
			errorString=errorString+"\n"+error+"：請輸入舉發人代碼1。";
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
//	if (myForm.Insurance.value==""){
//		error=error+1;
//		errorString=errorString+"\n"+error+"：請輸入第三責任險。";
//	}
	if (TDFastenerErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物1 輸入錯誤。";
	}
	if (TDFastenerErrorLog2==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物2 輸入錯誤。";
	}
	if (TDFastenerErrorLog3==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物3 輸入錯誤。";
	}
	if (myForm.Fastener1.value==myForm.Fastener2.value && myForm.Fastener1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物1 與代保管物2 重複。";
	}
	if (myForm.Fastener1.value==myForm.Fastener3.value && myForm.Fastener1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物1 與代保管物3 重複。";
	}
	if (myForm.Fastener2.value==myForm.Fastener3.value && myForm.Fastener2.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：代保管物2 與代保管物3 重複。";
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
		if(parseInt(myForm.RuleSpeed.value) > 300){
			error=error+1;
			errorString=errorString+"\n"+error+"：限速、限重大於 300Km/h。";
		}
		if(parseInt(myForm.IllegalSpeed.value) > 300){
			error=error+1;
			errorString=errorString+"\n"+error+"：實際車速、車重大於 300Km/h。";
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
	if (((myForm.Rule1.value.substr(0,3))=="293" || (myForm.Rule2.value.substr(0,3))=="293") && (myForm.RuleSpeed.value=="" || myForm.IllegalSpeed.value=="")){
		error=error+1;
		errorString=errorString+"\n"+error+"：您選擇超重法條，但是未輸入限重或車重。";
	}
	if (error==0){
			getChkCarSimpleIDandRule();
	}else{
		alert(errorString);
	}
}

//檢查車種跟法條內容相不相符
function getChkCarSimpleIDandRule(){
	NewIllDate=myForm.IllegalDate.value;
	NewIllTime=myForm.IllegalTime.value;
	NewIllRule1=myForm.Rule1.value;
	NewIllRule2=myForm.Rule2.value;
	NewCarNo=myForm.CarNo.value;
	NewCarSimpleID=myForm.CarSimpleID.value;
	runServerScript("getChkCarSimpleIDandRule.asp?CarID="+NewCarNo+"&IllDate="+NewIllDate+"&IllTime="+NewIllTime+"&IllRule1="+NewIllRule1+"&IllRule2="+NewIllRule2+"&CarSimpleID="+NewCarSimpleID);
}
function setChkCarSimpleIDandRule(RuleDetail){
	ErrorStr="";
	if (RuleDetail==1){
		ErrorStr="違規事實與簡式車種不符，請確認是否正確。";
	}
	<%if sys_City="南投縣" then%>
	//檢查到案日有沒有違規日+15天
	if (myForm.IsMail(0).checked==true){
		getDealDateValue=<%=getReportDealDateValue%>;
	}else{
		getDealDateValue=<%=getStopDealDateValue%>;	
	}
	BFillDateTemp=myForm.IllegalDate.value;
	if (BFillDateTemp.length >= 6){
		Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
		Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
		Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
		var BFillDate=new Date(Byear,Bmonth-1,Bday)
		var DLineDate=new Date()
		DLineDate=DateAdd("d",getDealDateValue,BFillDate);
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
		if (myForm.DealLineDate.value<Dyear+Dmonth+Dday && myForm.TrafficAccidentType.value=="" && myForm.CaseInByMem.checked==false){
			ErrorStr=ErrorStr+"應到案日小於違規日加"+getDealDateValue+"天，請確認是否正確。";
		}
	}
	<%end if%>
	if (ErrorStr!=""){
		if(confirm(ErrorStr+"\n是否確定要存檔？")){
			myForm.kinds.value="DB_insert";
			myForm.submit();
		}
	}else{
		document.myForm.kinds.value="DB_insert";
		document.myForm.submit();
	}
}
//是否為特殊用車
function getVIPCar(){
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value;
	myForm.CarNo.value=myForm.CarNo.value.toUpperCase();
	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");
	if (myForm.CarNo.value.length >= 3 && ((Rule1tmp.substr(0,2))!="32" && (Rule2tmp.substr(0,2))!="32" && (Rule1tmp.substr(0,5))!="12102" && (Rule2tmp.substr(0,5))!="12102" && (Rule1tmp.substr(0,3))!="334" && (Rule2tmp.substr(0,3))!="334")){
		var CarNum=myForm.CarNo.value;
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType==0){
			alert("車牌格式錯誤，如該車輛無車牌則可忽略此訊息！");
			//myForm.CarNo.select();
		}else{
			//runServerScript("getVIPCar.asp?CarID="+CarNum);
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
function DelSpace4(){
	myForm.Rule4.value=myForm.Rule4.value.replace(/[^\d]/g,'');
	getRuleData4();
}
//違規事實1(ajax)
function getRuleData1(){
	if (myForm.Rule1.value.length > 6){
		var Rule1Num=myForm.Rule1.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=myForm.RuleVerSion.value;
		runServerScript("getRuleDetail.asp?RuleOrder=1&RuleID="+Rule1Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		CallChkLaw1();
		ShowLayerWeight();
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
		var VerNo=myForm.RuleVerSion.value;
		runServerScript("getRuleDetail.asp?RuleOrder=2&RuleID="+Rule2Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		CallChkLaw2();
		ShowLayerWeight();
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
function AutoKeyCarNo(){
	//車牌號碼遇到32 與DCI 傳輸固定用身分證號前6碼
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value
<%if sys_City<>"南投縣" and sys_City<>"花蓮縣" and sys_City<>"台中市" and sys_City<>"台東縣" and sys_City<>"宜蘭市" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"高雄市" and sys_City<>"嘉義市" and sys_City<>"屏東縣" then%>
	if ((Rule1tmp.substr(0,2))=="32" || (Rule2tmp.substr(0,2))=="32" || (Rule1tmp.substr(0,5))=="12102" || (Rule2tmp.substr(0,5))=="12102" || (Rule1tmp.substr(0,3))=="334" || (Rule2tmp.substr(0,3))=="334"){
		myForm.CarNo.value=myForm.DriverPID.value.substr(0,6);
	}
<%end if%>
}
//違規事實3(ajax)
function getRuleData3(){
	if (myForm.Rule3.value.length > 6){
		var Rule3Num=myForm.Rule3.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=myForm.RuleVerSion.value;
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
////違規事實4(ajax)
function getRuleData4(){
	if (myForm.Rule4.value.length > 6){
		var Rule4Num=myForm.Rule4.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=myForm.RuleVerSion.value;
		runServerScript("getRuleDetail.asp?RuleOrder=4&RuleID="+Rule4Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		//CallChkLaw4();
	}else if (myForm.Rule4.value.length <= 6 && myForm.Rule4.value.length > 0){
		Layer4.innerHTML=" ";
		myForm.ForFeit4.value="";
		TDLawErrorLog4=1;
	}else{
		Layer4.innerHTML=" ";
		myForm.ForFeit4.value="";
		TDLawErrorLog4=0;
	}
}
////到案處所(ajax)
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
	if (myForm.BillUnitID.value.length > 0){
		var BillUnitNum=myForm.BillUnitID.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum);
	}else{
		Layer6.innerHTML=" ";
		TDUnitErrorLog=1;
	}
}
//保管物品一(ajax)
function getFastener1(){
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
//保管物品三(ajax)
function getFastener3(){
	if (myForm.Fastener3.value.length == 1){
		var FastenerNum=myForm.Fastener3.value;
		runServerScript("getFastener.asp?FastenerOrder=3&FastenerID="+FastenerNum);
	}else if (myForm.Fastener3.value.length == 0){
		Layer10.innerHTML=" ";
		TDFastenerErrorLog3=0;
		myForm.Fastener3Val.value="";
	}else{
		Layer10.innerHTML=" ";
		TDFastenerErrorLog3=1;
		myForm.Fastener3Val.value="";
	}
}
//違規地點代碼(ajax)
function getillStreet(){
<%if sys_City<>"基隆市" and sys_City<>"彰化縣" then%>
	myForm.IllegalAddressID.value=myForm.IllegalAddressID.value.toUpperCase();
<%end if%>
	if (event.keyCode==116){	
		event.keyCode=0;
		OstreetID=myForm.IllegalAddressID.value;

		window.open("Query_Street.asp?OstreetID="+OstreetID,"WebPage_Street_People2","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}else if (myForm.IllegalAddressID.value.length >= 1){
		var illAddrNum=myForm.IllegalAddressID.value;
		runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
	}
	
}
//舉發人一(ajax)
function getBillMemID1(){
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
//攔停由違規日期帶入應到案日期
function getDealLineDate(){
	if(TodayDate < myForm.IllegalDate.value){
		alert("違規日期不得大於今天!!");
		myForm.IllegalDate.select();
	}
	getDealDateValue=<%=getStopDealDateValue%>;	//要加幾天
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');
<%if sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then %>
	BFillDateTemp=myForm.IllegalDate.value;
	if (BFillDateTemp.length >= 6){
		Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
		Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
		Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
		var BFillDate=new Date(Byear,Bmonth-1,Bday)
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
		myForm.BillFillDate.value=myForm.IllegalDate.value;
	}
<%end if%>
}
//嘉義縣用填單日+15
function getDealLineDate_Stop(){
	getSDealDateValue=15;	//要加幾天
	myForm.BillFillDate.value=myForm.BillFillDate.value.replace(/[^\d]/g,'');
	BFillDateTemp=myForm.BillFillDate.value;
<%if sys_City="宜蘭縣" then %>
	myForm.IllegalDate.value=myForm.BillFillDate.value;
<%end if%>
	if (BFillDateTemp.length >= 6){
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
function CheckBillNoExist(){
	myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
	var BillNum=myForm.Billno1.value;
	if (myForm.Billno1.value!=""){
		ReadBillNo=myForm.Billno1.value.replace(' ','');
		if (ReadBillNo.length != 9 ){
			alert("單號不足九碼！");
			myForm.Billno1.select();
		}else if(myForm.Billno1.value != myForm.OldBillNo.value){
			runServerScript("getCheckBillNoExist.asp?BillNo="+BillNum);
		}
	}
}

function chkTrafficAccidentType(){
	myForm.TrafficAccidentType.value=myForm.TrafficAccidentType.value.toUpperCase();
	if (myForm.TrafficAccidentType.value.length >= 1){
		if (myForm.TrafficAccidentType.value!="1" && myForm.TrafficAccidentType.value!="2" && myForm.TrafficAccidentType.value!="3" && myForm.TrafficAccidentType.value!=" "){
			alert("交通事故種類填寫錯誤!");
			myForm.TrafficAccidentType.select();
			//myForm.TrafficAccidentType.value="";
		}
	}
}
function RuleSpeedforLaw(){
	myForm.RuleSpeed.value=myForm.RuleSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then %>
	CallChkLaw1();
	CallChkLaw2();
<%end if%>
	//CallChkLaw3();
	//CallChkLaw4();
	var IntError=0;
	var StrError="";
	if (myForm.RuleSpeed.value > 100){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：限速、限重超過100。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) > 60){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速60公里以上。";
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
	ShowLayerWeight();
}
function IllegalSpeedforLaw(){
	myForm.IllegalSpeed.value=myForm.IllegalSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
<%if sys_City<>"彰化縣" and sys_City<>"雲林縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" then %>
	CallChkLaw1();
	CallChkLaw2();
<%end if%>
	//CallChkLaw3();
	//CallChkLaw4();
	var IntError=0;
	var StrError="";
	if (myForm.IllegalSpeed.value > 100){
		IntError=IntError+1;
		StrError=StrError+"\n"+IntError+"：車速、車重超過100。";
	}
	if((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) > 60){
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：車速超過限速60公里以上。";
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
	ShowLayerWeight();
}
function funGetSpeedRule(){
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
	}
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
/*
function CallChkLaw3(){
	if (TDLawNum>=1){
		if (!funcChkLaw(myForm.Rule3.value)){
			alert("請確認法條三是否填寫正確");
		}	
	}
}
function CallChkLaw4(){
	if (TDLawNum==2){
		if (!funcChkLaw(myForm.Rule4.value)){
			alert("請確認法條四是否填寫正確");
		}	
	}
}
*/
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
//顯示超重的法條
function ShowLayerWeight(){
	if (((myForm.Rule1.value.substr(0,3))=="293" || (myForm.Rule2.value.substr(0,3))=="293") && myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!="" && ((myForm.IllegalSpeed.value-myForm.RuleSpeed.value) > 0)){
		OverWeight=Math.ceil(myForm.IllegalSpeed.value-myForm.RuleSpeed.value);
		if (OverWeight < 10){
			OverWeight="0"+OverWeight;
		}
		if ((myForm.Rule1.value.substr(0,3))=="293"){
			LawWeight=myForm.Rule1.value;
		}else if ((myForm.Rule1.value.substr(0,3))=="293"){
			LawWeight=myForm.Rule2.value;
		}
		LayerWeight.innerHTML="293"+OverWeight+LawWeight.substring(5,LawWeight.length);
	}else{
		LayerWeight.innerHTML=" ";
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
				if (((myForm.IllegalAddress.value.indexOf("快速道路",0)) != -1) && ((myForm.IllegalAddress.value.indexOf("快速公路",0)) != -1)){
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
//增加違規法條
function InsertLaw(){
	TDLawNum=1;
	TDLaw1.innerHTML="違規法條三";
	TDLaw2.innerHTML="<input type='text' size='10' value='' name='Rule3' onKeyUp='getRuleData3();' onchange='DelSpace3();'  onkeydown='funTextControl(this);'> <img src='../Image/BillkeyInButton.jpg' width='25' height='23' onclick='OpenQueryLaw3()' alt='查詢法條'> <div id='Layer3' style='position:absolute ; width:610px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;'></div><input type='hidden' name='ForFeit3' value=''><img src='space.gif' width='613' height='2'><img src='../Image/Law4.jpg' width='45' height='25' onclick='InsertLaw2()' alt='違規法條四'>";

	<%if sys_City="嘉義縣" or sys_City="宜蘭縣" or sys_City="高雄縣" or sys_City="台南縣" or sys_City="台南市" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||IllegalAddressID,IllegalAddress||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="台東縣" then %>
	MoveTextVar("Billno1,Insurance,DriverName||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="彰化縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||CarAddID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="南投縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||Fastener1,Fastener2,Fastener3||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="高雄市" then %>
	MoveTextVar("Billno1,Insurance||CarNo,CarSimpleID||DriverPID,DriverBrith||Fastener1,Fastener2,Fastener3||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%else%>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	
	<%end if%>
	myForm.Rule3.focus();
}
function OpenQueryLaw3(){
	window.open("Query_Law.asp?LawOrder=3&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes");
}
//增加違規法條
function InsertLaw2(){
	TDLawNum=2;
	TDLaw3.innerHTML="違規法條四";
	TDLaw4.innerHTML="<input type='text' size='10' value='' name='Rule4' onKeyUp='getRuleData4();' onchange='DelSpace4();'  onkeydown='funTextControl(this);'> <img src='../Image/BillkeyInButton.jpg' width='25' height='23' onclick='OpenQueryLaw4()' alt='查詢法條'> <div id='Layer4' style='position:absolute ; width:610px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;'></div><input type='hidden' name='ForFeit4' value=''>";

	<%if sys_City="嘉義縣" or sys_City="宜蘭縣" or sys_City="高雄縣" or sys_City="台南縣" or sys_City="台南市" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||Rule4||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="嘉義市" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||Rule4||IllegalAddressID,IllegalAddress||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="台東縣" then %>
	MoveTextVar("Billno1,Insurance,DriverName||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||Rule4||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="彰化縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||Rule4||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="雲林縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||CarAddID,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||RuleSpeed,IllegalSpeed||Rule1||Rule2||Rule3||Rule4||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="南投縣" then %>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||Fastener1,Fastener2,Fastener3||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||Rule4||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%elseif sys_City="高雄市" then %>
	MoveTextVar("Billno1,Insurance||CarNo,CarSimpleID||DriverPID,DriverBrith||Fastener1,Fastener2,Fastener3||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||Rule4||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	<%else%>
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith||CarNo,CarSimpleID||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||Rule4||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||Fastener1,Fastener2,Fastener3||BillUnitID,SignType,BillFillDate||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
	
	<%end if%>
	myForm.Rule4.focus();
}
function OpenQueryLaw4(){
	window.open("Query_Law.asp?LawOrder=4&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes");
}
//刪除違規法條
//function DeleteLaw(){
//	if (TDLawNum==1){
//		TDLaw1.innerHTML=" ";
//		TDLaw2.innerHTML=" ";
//		TDLawNum=TDLawNum-1;
//		myForm.Lawdel.disabled=true;
//	}else if (TDLawNum==2){
//		TDLaw3.innerHTML=" ";
//		TDLaw4.innerHTML=" ";
//		TDLawNum=TDLawNum-1;
//		myForm.Lawadd.disabled=false;
//	}
//}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
	return win;
}
function LawOpen3(){
	UrlStr="Query_Law.asp?LawOrder=3&RuleVer=<%=trim(rs1("RuleVer"))%>";
	newWin(UrlStr,"WebPage1",550,355,0,0,"yes","no","yes","no");
}
function LawOpen4(){
	UrlStr="Query_Law.asp?LawOrder=4&RuleVer=<%=trim(rs1("RuleVer"))%>";
	newWin(UrlStr,"WebPage1",550,355,0,0,"yes","no","yes","no");
}
function FuncChkPID(){
	myForm.DriverPID.value=myForm.DriverPID.value.toUpperCase();
	myForm.DriverPID.value=myForm.DriverPID.value.replace(/[\s　]+/g, "");
	if (myForm.DriverPID.value.length == 10){
		if (!check_tw_id(myForm.DriverPID.value)){
			alert("身分證輸入錯誤！");
			//myForm.DriverPID.focus();
		}else{
			if (myForm.DriverPID.value.substr(1,1)=="1"){
				document.myForm.DriverSex.value="1";
			}else{
				document.myForm.DriverSex.value="2";
			}
		}
	}else {
		alert("身分證輸入錯誤！");
		//myForm.DriverPID.focus();
	}
}
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
<%if sys_City="台東縣" then%>
	}else if (event.keyCode==120){ //台東縣F9存檔
		event.keyCode=0;   
		InsertBillVase();
	}
<%else%>
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
		InsertBillVase();
	}
<%end if%>
}
function focusToDriverPID(){
	myForm.DriverBrith.value=myForm.DriverBrith.value.replace(/[^\d]/g,'');
	if (myForm.DriverBrith.value.length==6){
		BBrithTmp=myForm.DriverBrith.value;

		BByear=parseInt(BBrithTmp.substr(0,BBrithTmp.length-4))+1911;
		BBmonth=BBrithTmp.substr(BBrithTmp.length-4,2);
		BBday=BBrithTmp.substr(BBrithTmp.length-2,2);
		var BrithDate=new Date(BByear,BBmonth-1,BBday)
		var DLineDate=new Date()

		BirthYear=DateAdd("y",-18,DLineDate);

		if (BirthYear < BrithDate){
			alert("違規人年齡低於18歲!!");
		}
	}
}
function ProjectF5(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_Project.asp","WebPage_Street_People","left=0,top=0,location=0,width=800,height=460,resizable=yes,scrollbars=yes");
	}/*else if (event.keyCode==13){	
		event.keyCode=0;
		OnKeyUpNextTag("Note");
	}*/
	if (myForm.ProjectID.value.length > 0){
		var BillProjectID=myForm.ProjectID.value;
		runServerScript("getProjectID.asp?BillProjectID="+BillProjectID);
	}else{
		Layer001.innerHTML="";
		TDProjectIDErrorLog=0;
	}
}
//簽收狀況(小轉大寫，限定A or U)
	function funcSignType(){
		myForm.SignType.value=myForm.SignType.value.toUpperCase();
		if (myForm.SignType.value==""){
			myForm.SignType.focus();
			alert("簽收狀況未填寫!!");
		}else if (myForm.SignType.value!="A" && myForm.SignType.value!="U" && myForm.SignType.value!="2" && myForm.SignType.value!="3" && myForm.SignType.value!="5"){
			myForm.SignType.select();
			alert("簽收狀況填寫錯誤!!");
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
//修改監理站回傳資料
function UpdateDciData(){
	window.open("BillKeyIn_Dci_Update.asp?BillSN="+myForm.BillSN.value+"&theUpdVer=1","WebPage1","left=20,top=10,location=0,width=900,height=600,resizable=yes,scrollbars=yes");
}
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
<%if sys_City="彰化縣" or sys_City="雲林縣" or sys_City="嘉義縣" or sys_City="宜蘭縣" or sys_City="台東縣" or sys_City="高雄縣" or sys_City="台南縣" or sys_City="台南市" then%>
				myForm.RuleSpeed.select();
<%elseif sys_City="嘉義市" then%>
				myForm.DealLineDate.select();
<%else%>
				myForm.Rule1.select();
<%end if%>
			}
		}
	}
//--------------------------
	ShowLayerWeight();
</script>
<%
rs1.close
set rs1=nothing
conn.close
set conn=nothing
%>
</html>
