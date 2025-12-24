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
<title>拖吊資料建檔作業</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
AuthorityCheck(236)
'==========POST=========
'單號
if trim(request("billno1"))="" then
	theBillno=""
else
	theBillno=Mid(trim(request("billno1")),1,3)
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
	theBilltype="1"
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
	set rsCity=Nothing
	sys_city="高雄市"
	getStopDealDateValue=30
	'******測試**********
	%>
<script language="JavaScript">
		
		
 </script>
<%
	'********************
'新增告發單
if trim(request("kinds"))="DB_insert" then
	strBillChk="select * from BillBase where BillNo='"&UCase(trim(request("Billno1")))&UCase(trim(request("Billno2")))&"' and RecordStateID=0"
	set rsBillChk=conn.execute(strBillChk)
	if rsBillChk.eof then
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
			theInsurance = "null"
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
		end If
		'拖吊已結
		theCarAddID="8"
		'交通事故類別
		if trim(request("TrafficAccidentType"))="" then
			theTrafficAccidentType=""
		else
			theTrafficAccidentType=trim(request("TrafficAccidentType"))
		end if

		'查流水號
		strSN="select BillBase_seq.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theSN=trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing

		'簽收狀況 A=A,U 2 3 =U ,5=''
		if UCase(trim(request("SignType")))="A" then
			UserSignType="A"
		elseif UCase(trim(request("SignType")))="U" or UCase(trim(request("SignType")))="2" or UCase(trim(request("SignType")))="3" then
			UserSignType="U"
		else
			UserSignType=""
		end if

		'BillBase
		strInsert="insert into BillBase(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
					",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
					",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
					",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
					",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
					",BillMemID2,BillMem2,BillMemID3,BillMem3,BillMemID4,BillMem4" &_
					",BillFillerMemberID,BillFiller" &_
					",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
					",Note,EquipmentID,RuleVer,DriverSex,TrafficAccidentNo,TrafficAccidentType,SignType,CaseInDate)" &_
					" values("&theSN&",'"&trim(request("BillType"))&"','"&UCase(trim(request("Billno1")))&UCase(trim(request("Billno2")))&"'" &_
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
					","&theBillFillDate&","&theDealLineDate&",'9',0,SYSDate,'" & theRecordMemberID &"'" &_
					",'"&trim(request("Note"))&"','"&trim(request("IsMail"))&"','"&theRuleVer&"'" &_
					",'"&trim(request("DriverSex"))&"','"&trim(request("TrafficAccidentNo"))&"','"&theTrafficAccidentType&"','"&UserSignType&"'" &_
					",sysdate)"
					conn.execute strInsert

				strInsCar="insert into DCILog(SN,BillSN,BillNo,BillTypeID,CarNo,BillUnitID,RecordDate" &_
					",RecordMemberID,ExchangeDate,ExchangeTypeID,DCIwindowName,BatchNumber)"&_
					"values(DCILOG_SEQ.nextval,BillBase_seq.Currval,'',2,'"&UCase(trim(request("CarNo")))&"'" &_
					",'"&trim(request("BillUnitID"))&"',sysdate,"&Session("User_ID")&",sysdate,'A','Z','"&"A"&Right("0"&Year(date)-1911,3)&Right("0"&Month(date),2)&Right("0"&day(date),2)&"'" &_
					")" 
					conn.execute strInsCar
					'theDriverBirth , theBillFillDate   
		'簽收狀況 BillUserSignDate
		if UCase(trim(request("SignType")))="2" or UCase(trim(request("SignType")))="3" or UCase(trim(request("SignType")))="5" then
			strInsSignType="insert into BillUserSignDate values("&theSN&",'"&UCase(trim(request("SignType")))&"','','')"
			conn.execute strInsSignType
		end if
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
		if trim(request("Fastener3"))<>"" then
			strInsFastene3="insert into BillFastenerDetail(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
						" values(BillFastenerDetail_seq.nextval,"&theSN&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener3"))&"','')"
			conn.execute strInsFastene3
		end if

		if sys_City="高雄縣" then
			BillNotKeyInStr=chkBillNoIsAllKeyIn(UCase(trim(request("Billno1")))&UCase(trim(request("Billno2"))))
			if BillNotKeyInStr<>"" then
	%>
	<script language="JavaScript">
		//alert("新增完成！\n下列單號尚未開單：<%=BillNotKeyInStr%>");
	</script>
<%			end if
		end if
	else
	%>
	<script language="JavaScript">
		alert("此單號：<%=UCase(trim(request("Billno1")))&UCase(trim(request("Billno2")))%>，已建檔！！");
	</script>
<%	
	end if
end if

Session.Contents.Remove("BillTime_Stop")
BillTime_StopTmp=DateAdd("s" , 1, now)
Session("BillTime_Stop")=date&" "&hour(BillTime_StopTmp)&":"&minute(BillTime_StopTmp)&":"&second(BillTime_StopTmp)
'response.write Session("BillTime_Stop")
'總共幾筆
Session.Contents.Remove("BillCnt_Stop")
Session.Contents.Remove("BillOrder_Stop")
strSqlCnt="select count(*) as cnt from BillBase where BillTypeID='1' and BillStatus in ('9') and RecordDate between TO_DATE('"&date&"/0/0/0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&date&"/23/59/59','YYYY/MM/DD/HH24/MI/SS') and RecordStateID=0 and RecordMemberID="&theRecordMemberID

set rsCnt1=conn.execute(strSqlCnt)
	Session("BillCnt_Stop")=trim(rsCnt1("cnt"))
	Session("BillOrder_Stop")=trim(rsCnt1("cnt"))+1
rsCnt1.close
set rsCnt1=nothing
%>

<style type="text/css">
<!--
.style1 {font-size: 14px;}
.style3 {font-size: 15px;}
.style4 {
	color: #FF0000;
	font-size: 12px;
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
.style5 {font-size: 12px;}
.style6 {font-size: 16px;}
.style7 {
	color: #FF0000;
	font-size: 12px;
	line-height:14px;
}
.style8 {
	color: #000000;
	font-size: 12px;
	line-height:14px;
}
.btn2 {font-size: 13px;}

-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
<br>

	<form name="myForm" method="post">  
		<table width='1000' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="6"><strong>拖吊資料建檔作業</strong>&nbsp; &nbsp; 日期格式：951220 &nbsp;時間格式：2300(24小時制)&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 建檔日期：<%=ginitdt(now)%>&nbsp; &nbsp; <input type="checkbox" name="CaseInByMem" value="1">逾違規日期超過三個月強制建檔</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" ><div align="right"><span class="style4">＊</span>單號</div></td>
				<td colspan=5>
				<input name="Billno1" type="text" value="<%=theBillno%>" size="3" maxlength="9" onkeyup="GetBillNo();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				<input name="Billno2" type="text" value="" size="10" maxlength="9" onkeydown="funTextControl(this);" onkeyup="GetBillNo2();" style=ime-mode:disabled>
				</td>
			</tr>
			
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規車號</div></td>
				<td>
				<input type="text" size="10" name="CarNo" onBlur="getVIPCar();" onkeydown="funTextControl(this);" style=ime-mode:disabled maxlength="8">
			    <div id="Layer7" style="position:absolute; width:100px; height:24px; z-index:0; border: 1px none #000000; color: #FF0000;"><span class="style8"></span>
				</div>
				<div id="Layer137" style="position:absolute; width:100px; height:24px; z-index:0; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000;">
				</div>
				
				</td>
				<td bgcolor="#FFFFCC" ><div align="right"><span class="style4">＊</span>簡式車種</div></td>
				<td colspan="3">
				<input type="text" maxlength="1" size="3" value="<%=trim(request("CarSimpleID"))%>" name="CarSimpleID" onBlur="getRuleAll();" onfocus="this.select();" onkeydown="funTextControl(this);" onkeyup="GetCarSimpleID();" style=ime-mode:disabled>
				<div id="Layer111" style="position:absolute; width:275px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<font color="#ff000" size="2"> 1汽車/ 2拖車/ 3重機/ 4輕機/ 6 臨時車牌</font>&nbsp;&nbsp;
				<div id="Layer170" style="position:absolute; width:245px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style8">
					無車牌之拼裝車車號請輸入身份證前六碼
					<br>
					550cc以上重機簡式車種請選擇重機
				</span>
				  </div>
				</div>
				
				</td>
		    </tr>
<tr>
				<td bgcolor="#FFFFCC"><div align="right">違規人證號</div></td>
				<td>
				<input type="text" size="10" <%
			response.write "maxlength='10'"
				%> name="DriverPID" onBlur="FuncChkPID();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				<div id="Layer1127" style="position:absolute; width:100px; height:24px; z-index:0; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000;">
				</div>
				</td>
				<td bgcolor="#FFFFCC" align="right">違規人出生日</td>
				<td <%
				response.write "colspan=""3"""
				%>><input type="text" size="10" maxlength="6" name="DriverBrith" onBlur="focusToDriverPID()" onkeydown="funTextControl(this);" style=ime-mode:disabled></td>
			</tr>
			<tr>

				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規日期</div></td>
				<td>
				<input type="text" size="10" maxlength="6" value="<%
				response.write trim(request("IllegalDate"))
				%>" name="IllegalDate" onfocus="this.select()" onBlur="getDealLineDate()" onkeydown="funTextControl(this);" onkeyup="GetIllegalDate();" style=ime-mode:disabled>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規時間</div></td>
				<td colspan="3">
				<input type="text" size="10" maxlength="4" value="<%
				response.write trim(request("IllegalTime"))
				%>" name="IllegalTime" onkeyup="GetIllegalTime();" onBlur="value=value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">違規地點代碼</div></td>
				<td>
					<input type="text" size="10" value="<%=request("IllegalAddressID")%>" name="IllegalAddressID" onKeyUp="getillStreet();" onkeydown="funTextControl(this);" style=ime-mode:disabled onblur="getillStreet2();" onfocus="this.select()">
					<input type="hidden" name="OldIllegalAddressID" value="<%=request("IllegalAddressID")%>">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=455,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規地點</div></td>
				<td colspan="3">
					<input type="text" size="44" value="<%=trim(request("IllegalAddress"))%>" name="IllegalAddress" style=ime-mode:active onkeyup="AutoGetIllStreet();" onkeydown="funTextControl(this);" onblur="funGetSpeedRule()" onfocus="this.select()" >
					<input type="hidden" name="chkHighRoad" value="1">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" ><div align="right"><span class="style4">＊</span>違規法條一</div></td>
				<td>
					<input type="text" maxlength="8" size="10" value="<%
				response.write trim(request("Rule1"))
					%>" name="Rule1" onkeyup="getRuleData1();" onchange="DelSpace1();" onblur="AutoKeyCarNo();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%
						response.write theRuleVer
					%>","WebPage1","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes")'alt="查詢法條">
					</td><td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規金額</div></td><td colspan=3>
					<input type="text" name="ForFeit1" size="10" maxlength="8" value="<%
						response.write request("ForFeit1")
					%>"  onkeydown="funTextControl(this);" onBlur="value=value.replace(/[^\d]/g,'')" style=ime-mode:disabled>
					<!-- <img src="../Image/BillLawPlusButton.jpg" width="25" height="23" onclick="Add_LawPlus()" alt="附加說明"> -->
					<div id="Layer1" style="position:absolute ; width:400px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
					if trim(request("Rule1"))<>"" then
						strRule1="select IllegalRule from Law where ItemID='"&trim(request("Rule1"))&"' and Version='"&trim(theRuleVer)&"'"
						set rsRule1=conn.execute(strRule1)
						if not rsRule1.eof then
							response.write trim(rsRule1("IllegalRule"))
						end if
						rsRule1.close
						set rsRule1=nothing
					end if
					%></div>

				</td>
			</tr>
<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發人代碼1</div></td>
		  		<td>
					<input type="text" size="5" name="BillMem1" onkeyup="getBillMemID1();" onkeydown="funTextControl(this);" style=ime-mode:disabled value="<%
					response.write trim(request("BillMem1"))
					%>">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemType=CarS&MemOrder=1","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer12" style="position:absolute ; width:90px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
					if trim(request("BillMemID1"))<>"" then
						strMemName1="select ChName from MemberData where MemberID="&trim(request("BillMemID1"))
						set rsMemName1=conn.execute(strMemName1)
						if not rsMemName1.eof then 
							response.write rsMemName1("ChName")
						end if
						rsMemName1.close
						set rsMemName1=nothing
					end if
					%></div>
					<input type="hidden" value="<%
					response.write trim(request("BillMemID1"))
					%>" name="BillMemID1">
					<input type="hidden" value="<%
					response.write trim(request("BillMemName1"))
					%>" name="BillMemName1">
					<input type="hidden" value="<%
					response.write trim(request("BillUnitTypeID1"))
					%>" name="BillUnitTypeID1">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>應到案日期</div></td>
				<td>
					<input type="text" size="10" maxlength="6" name="DealLineDate" value="<%=trim(request("DealLineDate"))%>" onBlur="value=value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>應到案處所</div></td>
				<td>
					<input type="text" size="5" value="32" name="MemberStation" onkeyup="getStation();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Station.asp","WebPage1","left=0,top=0,location=0,width=760,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer5" style="position:absolute ; width:241px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"><%
						if trim(request("MemberStation"))<>"" then
							strStation="select DciStationName from Station where StationID='"&trim(request("MemberStation"))&"'"
							set rsStation=conn.execute(strStation)
							if not rsStation.eof then
								response.write trim(rsStation("DciStationName"))
							end if
							rsStation.close
							set rsStation=Nothing
						End if
					%></div>
					</span>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發單位</div></td>
				<td colspan=5>
					<input type="text" size="5" name="BillUnitID" onkeyup="getUnit();" onkeydown="funTextControl(this);" style=ime-mode:disabled value="<%=trim(request("BillUnitID"))%>">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=700,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer6" style="position:absolute ; width:227px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
						if trim(request("BillUnitID"))<>"" then
							strBillName="select UnitName from UnitInfo where UnitID='"&trim(request("BillUnitID"))&"'"
							set rsBillName=conn.execute(strBillName)
							if not rsBillName.eof then
								response.write trim(rsBillName("UnitName"))
							end if
							rsBillName.close
							set rsBillName=nothing
						end if
					%></div>
					</span>
			</tr>
<!----------------------------------------------------------------------------------------------------------------------->
					<input type="hidden" maxlength="8" size="10" value="" name="Rule2">
					<input type="hidden" size="5" value="A" maxlength="1" name="SignType">
				    <input type="hidden" maxlength="1" size="3" value="<%=trim(request("Insurance"))%>" name="Insurance">
					<input type="hidden" size="5" value="<%=trim(request("BillFillDate"))%>" maxlength="6" name="BillFillDate">
					<input type="hidden" name="ForFeit2" value="">
					<input type="hidden" size="10" name="RuleSpeed">
					<input type="hidden" size="10" name="IllegalSpeed">
					<input type="hidden" size="5" name="BillMem2" value="">
					<input type="hidden" value="" name="BillMemID2">
					<input type="hidden" value="" name="BillMemName2">
					<input type="hidden" value="" name="BillUnitTypeID2">
					<input type="hidden" size="5" name="BillMem3" value="">
					<input type="hidden" value="" name="BillMemID3">
					<input type="hidden" value="" name="BillMemName3">
					<input type="hidden" value="" name="BillUnitTypeID3">
					<input type="hidden" size="5" name="BillMem4" value="">
					<input type="hidden" value="" name="BillMemID4">
					<input type="hidden" value="" name="BillMemName4">
					<input type="hidden" value="" name="BillUnitTypeID4">
				    <input type="hidden" size="5" value="" name="Fastener1">
	                <input type="hidden" value="" name="Fastener1Val">
	                <input type="hidden" size="5" value="" name="Fastener2">
                    <input type="hidden" value="" name="Fastener2Val">
               	    <input type="hidden" size="5" value="" name="Fastener3">
                    <input type="hidden" value="" name="Fastener3Val">
					<input type="hidden" maxlength="2" size="3" value="" name="CarAddID">
					<input type="hidden" name="IsMail" value="-1">
					<input type="hidden" size="5" value="" name="ProjectID">
					<input type="hidden" size="20" value="" name="Note">
					<input type="hidden" size="10" name="TrafficAccidentNo" Value="">
					<input type="hidden" maxlength="1" size="5" name="TrafficAccidentType" Value="">
<!----------------------------------------------------------------------------------------------------------------------->
			<tr>
				<td bgcolor="#FFDD77" align="center" colspan="6">
					<input type="button" value="儲 存 <%
						response.write "F2"
					%>" onclick="InsertBillVase();" <%
				'1:查詢 ,2:新增 ,3:修改 ,4:刪除
				if CheckPermission(236,2)=false then
					response.write "disabled"
				end if
					%> class="btn1">
					<input type="hidden" name="kinds" value="">
					
                    <span class="style1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit1343" onClick="location='BillKeyIn_TakeCar.asp'" value="清 除 F4" class="btn1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit5322" onClick="funcOpenBillQry();" value="查 詢 F5" class="btn1">
					<img src="/image/space.gif" width="29" height="8">
                  </span>
                    <span class="style3">
                    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" class="btn1">
					<img src="/image/space.gif" width="29" height="8">
<!--					<input type="button" name="Submit4232" onClick="funPrintCaseList_Stop();" value="建檔清冊 F10" class="btn1">-->
					<!-- <img src="/image/space.gif" width="29" height="8">
			        <input type="button" name="SubmitNext" onClick="location='BillKeyIn_Car.asp'" value="下一筆"> -->
                </span>
				<!-- 告發類別 -->
				<input type="hidden" size="3" maxlength="1" value="<%=theBilltype%>" name="BillType">
				<!-- 違規人性別 -->
				<input type="hidden" value="" name="DriverSex">
				<!-- 附加說明 -->
				<!-- <input type="hidden" value="" name="Rule4"> -->
				<br>

				<img src="/image/space.gif" width="250" height="8">
				<input type="button" name="SubmitBack2" onClick="location='BillKeyIn_TakeCar_Back.asp?PageType=First'" value="<< 第一筆 Home" class="btn1">
				<img src="/image/space.gif" width="29" height="8">
				<input type="button" name="SubmitBack" onClick="location='BillKeyIn_TakeCar_Back.asp?PageType=Back'" value="< 上一筆 PgUp" class="btn1">
	<%
					response.write "建檔"&Session("BillCnt_Stop")&"筆"
					
				%>
				<div id="Layer1c69" style="position:absolute; width:160px; height:24px; z-index:0; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<span class="style7">使用上一筆搜尋功能只能查詢到自己建檔</span>
				</div>
				<img src="/image/space.gif" width="220" height="8">
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
var ChkCarIlldateFlag=0;
var TDProjectIDErrorLog=0;
var TDVipCarErrorLog=0;
var TodayDate=<%=ginitdt(date)%>;


MoveTextVar("Billno1,Billno2,CarNo,CarSimpleID,DriverPID||DriverBrith,IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1,ForFeit1||BillMem1,DealLineDate,MemberStation||BillUnitID");
//新增告發單
function InsertBillVase(){
	var error=0;
	var errorString="";
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value;
	ReadBillNo=myForm.Billno1.value.replace(/[\s　]+/g, "");
	ReadBillNo2=myForm.Billno2.value.replace(/[\s　]+/g, "");
	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");
	myForm.DriverPID.value=myForm.DriverPID.value.replace(/[\s　]+/g, "");
	if (myForm.Billno1.value==""){
		error=error+1;
		errorString=error+"：請輸入單號。";
		myForm.Billno1.select();		
	}else if(ReadBillNo.length!=3){     
		error=error+1;
		errorString=error+"：單號不足九碼。";
		myForm.Billno1.select();
	}
	if (myForm.Billno2.value==""){
		error=error+1;
		errorString=error+"：請輸入單號。";
		myForm.Billno2.select();		
	}else if(ReadBillNo2.length!=6){     
		error=error+1;
		errorString=error+"：單號不足九碼。";
		myForm.Billno2.select();
	}
	if (myForm.BillType.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入告發類別。";
	}
	if (myForm.DriverBrith.value!=""){
		if(!dateCheck( myForm.DriverBrith.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規人出生日期輸入錯誤。";	
		}
	}
	/*if (myForm.DriverName.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規人姓名。";
	}*/

	if (myForm.CarNo.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規車號。";
	}
	if (myForm.ForFeit1.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規金額。";
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
		TDUnitErrorLog==0
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
	if (((myForm.Rule1.value.substr(0,3))=="293" || (myForm.Rule2.value.substr(0,3))=="293") && (myForm.RuleSpeed.value=="" || myForm.IllegalSpeed.value=="")){
		error=error+1;
		errorString=errorString+"\n"+error+"：您選擇超重法條，但是未輸入限重或車重。";
	}
	if ((myForm.Rule1.value.substr(0,2))=="36" && (myForm.CarSimpleID.value=="3" || myForm.CarSimpleID.value=="4")){
			error=error+1;
			errorString=errorString+"\n"+error+"：法條與車種不符。";
	}else if ((myForm.Rule2.value.substr(0,2))=="36" && (myForm.CarSimpleID.value=="3" || myForm.CarSimpleID.value=="4")){
			error=error+1;
			errorString=errorString+"\n"+error+"：法條與車種不符。";
	}
	if (error==0){
			getChkCarIllegalDate();
	}else{
		alert(errorString);
	}
}

//檢查違規日期是否超過45天(高雄縣)
function ChkIllegalDateKS(IllDate){
	Iyear=parseInt(IllDate.substr(0,IllDate.length-4))+1911;
	Imonth=IllDate.substr(IllDate.length-4,2);
	Iday=IllDate.substr(IllDate.length-2,2);
	var IFillDate=new Date(Iyear,Imonth-1,Iday);
	var thisDay=new Date((new Date()).getYear(),(new Date()).getMonth(),(new Date()).getDate());
	var OverDate=new Date();
	OverDate=DateAdd("d",-45,thisDay);
	if (OverDate > IFillDate){
		return false;
	}else{
		return true;
	}
}

//增加違規法條
function InsertLaw(){
	TDLawNum=1;
	TDLaw1.innerHTML="違規法條三";
	TDLaw2.innerHTML="<input type='text' size='10' value='' name='Rule3' onKeyUp='getRuleData3();' onchange='DelSpace3();'  onkeydown='funTextControl(this);'> <img src='../Image/BillkeyInButton.jpg' width='25' height='23' onclick='OpenQueryLaw3()' alt='查詢法條'> <div id='Layer3' style='position:absolute ; width:609px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;'></div><input type='hidden' name='ForFeit3' value=''><img src='space.gif' width='613' height='2'><img src='../Image/Law4.jpg' width='45' height='25' onclick='InsertLaw2()' alt='違規法條四'>";

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
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||Fastener1,Fastener2,Fastener3||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
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
function InsertLaw2(){
	TDLawNum=2;
	TDLaw3.innerHTML="違規法條四";
	TDLaw4.innerHTML="<input type='text' size='10' value='' name='Rule4' onKeyUp='getRuleData4();' onchange='DelSpace4();'  onkeydown='funTextControl(this);'> <img src='../Image/BillkeyInButton.jpg' width='25' height='23' onclick='OpenQueryLaw4()' alt='查詢法條'> <div id='Layer4' style='position:absolute ; width:609px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;'></div><input type='hidden' name='ForFeit4' value=''>";

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
	MoveTextVar("Billno1,Insurance||DriverPID,DriverBrith,BillFillDate||CarNo,CarSimpleID||IllegalDate,IllegalTime||Fastener1,Fastener2,Fastener3||IllegalAddressID,IllegalAddress||Rule1||Rule2||Rule3||Rule4||RuleSpeed,IllegalSpeed||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||BillUnitID,SignType||CarAddID,ProjectID||Note,TrafficAccidentNo,TrafficAccidentType");
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
//檢查同車號同法條在同一天違規日期及違規時間前後兩小時內
function setChkCarIllegalDate(CarCnt,Illdate,RuleDetail)
{
	var ErrorStr="";
	if (CarCnt=="1"){
		ChkCarIlldateFlag="1";
	}else{
		ChkCarIlldateFlag="0";
	}
<%if sys_City<>"雲林縣" then%>
	ErrorStr="";
	if (RuleDetail==1){
		ErrorStr="違規事實與簡式車種不符，請確認是否正確。";
	}
	if (ChkCarIlldateFlag=="1"){
		ErrorStr=ErrorStr+"\n此車號於"+Illdate+"，有相同違規舉發，請確認有無連續開單。";
	}
	<%if sys_City="高雄縣" then%>
	if (!ChkIllegalDateKS(myForm.IllegalDate.value)){
		ErrorStr=ErrorStr+"\n違規日期已超過45天。";
	}
	<%end if%>
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
		myForm.kinds.value="DB_insert";
		myForm.submit();
	}
<%else%>
	//雲林的欄停不用檢查同一天違規日建檔
	ErrorStr="";
	//檢查到案日有沒有違規日+15天
	getDealDateValue=<%=getStopDealDateValue%>;	
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
		if (myForm.DealLineDate.value!=Dyear+Dmonth+Dday && myForm.TrafficAccidentType.value=="" && myForm.CaseInByMem.checked==false){
			ErrorStr=ErrorStr+"應到案日不是違規日加"+getDealDateValue+"天，請確認是否正確。";
		}
	}

	if (RuleDetail==1){
		ErrorStr=ErrorStr+"\n違規事實與簡式車種不符，請確認是否正確。";
	}
	if (ErrorStr!=""){
		if(confirm(ErrorStr+"\n是否確定要存檔？")){
			myForm.kinds.value="DB_insert";
			myForm.submit();
		}
	}else{
		myForm.kinds.value="DB_insert";
		myForm.submit();
	}
<%end if%>
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
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value;
	myForm.CarNo.value=myForm.CarNo.value.toUpperCase();
	myForm.CarNo.value=myForm.CarNo.value.replace(/[\s　]+/g, "");
	
	
		if (myForm.CarNo.value.length >= 4 && ((Rule1tmp.substr(0,2))!="32" && (Rule2tmp.substr(0,2))!="32" && (Rule1tmp.substr(0,5))!="12102" && (Rule2tmp.substr(0,5))!="12102" && (Rule1tmp.substr(0,3))!="334" && (Rule2tmp.substr(0,3))!="334")){
			var CarNum=myForm.CarNo.value;
			CarType=chkCarNoFormat(myForm.CarNo.value);
			if (CarType==0){
				alert("車牌格式錯誤，如該車輛無車牌則可忽略此訊息！");
				//myForm.CarNo.select();
			}else{
				runServerScript("getVIPCarForKeyIn.asp?CarID="+CarNum+"&BillType=1");
				myForm.CarSimpleID.value=CarType;
				myForm.DriverPID.select();
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
		var VerNo=<%=theRuleVer%>;
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91))
	{
						
		runServerScript("getRuleDetail2.asp?RuleOrder=1&RuleID="+Rule1Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		CallChkLaw1();
	}
	
		
	

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
function TabFocus(){
	//建檔時除了超重超速時游標才跳至限速限量欄位，其它法條則游標不跳至超重超速
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value;
<%if sys_City<>"南投縣" and sys_City<>"雲林縣" and sys_City<>"彰化縣" and sys_City<>"嘉義縣" and sys_City<>"宜蘭縣" and sys_City<>"台東縣" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"嘉義市" and sys_City<>"台中縣" then %>
		if ((Rule1tmp.substr(0,5))!="33101" && (Rule1tmp.substr(0,2))!="40" && (Rule1tmp.substr(0,5))!="43102" && (Rule1tmp.substr(0,3))!="293" && (Rule2tmp.substr(0,5))!="33101" && (Rule2tmp.substr(0,2))!="40" && (Rule2tmp.substr(0,5))!="43102" && (Rule2tmp.substr(0,3))!="293"){
			myForm.DealLineDate.select();
		}
<%end if%>
	//法條代碼遇到32 與DCI 傳輸固定用身分證號前六碼
	AutoKeyCarNo();
}
function AutoKeyCarNo(){
	//法條代碼遇到32 與DCI 傳輸固定用身分證號前六碼
	Rule1tmp=myForm.Rule1.value.substr(0,3);
	Rule2tmp=myForm.Rule2.value.substr(0,3);
	Rule1tmpb=myForm.Rule1.value.substr(0,2);
	Rule2tmpb=myForm.Rule2.value.substr(0,2);
	Rule1tmpc=myForm.Rule1.value.substr(0,5);
	Rule2tmpc=myForm.Rule2.value.substr(0,5);
<%if sys_City<>"南投縣" and sys_City<>"花蓮縣" and sys_City<>"台中市" and sys_City<>"台東縣" and sys_City<>"宜蘭市" and sys_City<>"高雄縣" and sys_City<>"台南縣" and sys_City<>"台南市" and sys_City<>"高雄市" and sys_City<>"嘉義市" and sys_City<>"屏東縣" then %>
	if (Rule1tmp=="320" || Rule2tmp=="320" || Rule1tmpc=="12102" || Rule2tmpc=="12102" || Rule1tmp=="321" || Rule2tmp=="321" || Rule1tmp=="322" || Rule2tmp=="322" || Rule1tmp=="334" || Rule2tmp=="334"){
		myForm.CarNo.value=myForm.DriverPID.value.substr(0,6);
	}
<%end if%>
<%if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" then %>
	MemberStationLaw="21,35,57,61,62";
	//法條代碼遇到21,35,57,61,62，應到案處所自動帶當地監理所
	if (((MemberStationLaw.indexOf(Rule1tmpb)!=-1 || MemberStationLaw.indexOf(Rule2tmpb)!=-1) && Rule1tmpb !="" && Rule2tmpb !="") || (MemberStationLaw.indexOf(Rule1tmpb)!=-1 && Rule2tmpb =="" && Rule1tmpb !="") || (MemberStationLaw.indexOf(Rule2tmpb)!=-1 && Rule1tmpb =="" && Rule2tmpb !="")){
		myForm.MemberStation.value=<%=trim(BillLawMemberStation)%>;
		getStation();
	}
<%end if%>
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
//違規事實4(ajax)
function getRuleData4(){
	if (myForm.Rule4.value.length > 6){
		var Rule4Num=myForm.Rule4.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
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
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_Station.asp","WebPage_Station12","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}else if (myForm.MemberStation.value.length > 1){
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
//保管物品三(ajax)
function getFastener3(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_Fastener.asp?FaOrder=3","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
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
	if (myForm.IllegalAddressID.value!=myForm.OldIllegalAddressID.value){
		//myForm.IllegalAddressID.value=myForm.IllegalAddressID.value.toUpperCase();

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
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91))
	{
	
		if (myForm.IllegalAddressID.value.length == 5 )
		{
			myForm.IllegalAddress.select();
		}			
	
	}			
			
		}
	}
			
}
//違規地點代碼OnBlur
function getillStreet2(){
	if (myForm.IllegalAddress.value==""){
		if (myForm.IllegalAddressID.value.length > 1){
			var illAddrNum=myForm.IllegalAddressID.value;
			runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
		}
	}
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule();
	<%end if%>
}
//舉發人一(ajax)
function getBillMemID1(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemType=CarS&MemOrder=1","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem1.value.length > 1){
		var BillMemNum=myForm.BillMem1.value;
		runServerScript("getBillMemID.asp?MType=CarS&MemOrder=1&MemID="+BillMemNum);
	}else if (myForm.BillMem1.value.length <= 1 && myForm.BillMem1.value.length > 0){
		Layer12.innerHTML=" ";
		myForm.BillMemID1.value="";
		myForm.BillMemName1.value="";
		myForm.BillUnitTypeID1.value="";
		TDMemErrorLog1=1;
	}else{
		Layer12.innerHTML=" ";
		myForm.BillMemID1.value="";
		myForm.BillMemName1.value="";
		myForm.BillUnitTypeID1.value="";
		TDMemErrorLog1=0;
	}
}
//舉發人二(ajax)
function getBillMemID2(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemType=CarS&MemOrder=2","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem2.value.length > 1){
		var BillMemNum=myForm.BillMem2.value;
		runServerScript("getBillMemID.asp?MType=CarS&MemOrder=2&MemID="+BillMemNum);
	}else if (myForm.BillMem2.value.length <= 1 && myForm.BillMem2.value.length > 0){
		Layer13.innerHTML=" ";
		myForm.BillMemID2.value="";
		myForm.BillMemName2.value="";
		myForm.BillUnitTypeID2.value="";
		TDMemErrorLog2=1;
	}else{
		Layer13.innerHTML=" ";
		myForm.BillMemID2.value="";
		myForm.BillMemName2.value="";
		myForm.BillUnitTypeID2.value="";
		TDMemErrorLog2=0;
	}
}
//舉發人三(ajax)
function getBillMemID3(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemType=CarS&MemOrder=3","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem3.value.length > 1){
		var BillMemNum=myForm.BillMem3.value;
		runServerScript("getBillMemID.asp?MType=CarS&MemOrder=3&MemID="+BillMemNum);
	}else if (myForm.BillMem3.value.length <= 1 && myForm.BillMem3.value.length > 0){
		Layer14.innerHTML=" ";
		myForm.BillMemID3.value="";
		myForm.BillMemName3.value="";
		myForm.BillUnitTypeID3.value="";
		TDMemErrorLog3=1;
	}else{
		Layer14.innerHTML=" ";
		myForm.BillMemID3.value="";
		myForm.BillMemName3.value="";
		myForm.BillUnitTypeID3.value="";
		TDMemErrorLog3=0;
	}
}
//舉發人四(ajax)
function getBillMemID4(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemType=CarS&MemOrder=4","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem4.value.length > 1){
		var BillMemNum=myForm.BillMem4.value;
		runServerScript("getBillMemID.asp?MType=CarS&MemOrder=4&MemID="+BillMemNum);
	}else if (myForm.BillMem4.value.length <= 1 && myForm.BillMem4.value.length > 0){
		Layer17.innerHTML=" ";
		myForm.BillMemID4.value="";
		myForm.BillMemName4.value="";
		myForm.BillUnitTypeID4.value="";
		TDMemErrorLog4=1;
	}else{
		Layer17.innerHTML=" ";
		myForm.BillMemID4.value="";
		myForm.BillMemName4.value="";
		myForm.BillUnitTypeID4.value="";
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
	BFillDateTemp=myForm.IllegalDate.value;
	if (BFillDateTemp.length >= 6){
		myForm.BillFillDate.value=myForm.IllegalDate.value;
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
		myForm.DealLineDate.value=Dyear+Dmonth+Dday;
	}
}
//嘉義縣用填單日+15
function getDealLineDate_Stop(){
	getSDealDateValue=15;	//要加幾天
	myForm.BillFillDate.value=myForm.BillFillDate.value.replace(/[^\d]/g,'');
	BFillDateTemp=myForm.BillFillDate.value;
<%if sys_City="宜蘭縣" or sys_City="台東縣" then %>
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
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
	return win;
}
function LawOpen3(){
	UrlStr="Query_Law.asp?LawOrder=3&RuleVer=<%=theRuleVer%>";
	newWin(UrlStr,"WebPage1",550,355,0,0,"yes","no","yes","no");
}
function LawOpen4(){
	UrlStr="Query_Law.asp?LawOrder=4&RuleVer=<%=theRuleVer%>";
	newWin(UrlStr,"WebPage1",550,355,0,0,"yes","no","yes","no");
}
function CheckBillNoExist(){
	myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
	myForm.Billno2.value=myForm.Billno2.value.toUpperCase();
	var BillNum=myForm.Billno1.value+myForm.Billno2.value;
	if (myForm.Billno1.value!="" || myForm.Billno2.value!=""){
		ReadBillNo=myForm.Billno1.value.replace(' ','')+myForm.Billno2.value.replace(' ','');
		if (ReadBillNo.length==9){
			runServerScript("getCheckBillNoExist.asp?BillNo="+BillNum);
		}else{
			alert("單號不足九碼！");
			myForm.Billno1.select();
		}
	}
}
function chkTrafficAccidentType(){
	//myForm.TrafficAccidentType.value=myForm.TrafficAccidentType.value.toUpperCase();
	if (myForm.TrafficAccidentType.value.length >= 1){
		if (myForm.TrafficAccidentType.value!="1" && myForm.TrafficAccidentType.value!="2" && myForm.TrafficAccidentType.value!="3" && myForm.TrafficAccidentType.value!=" "){
			alert("交通事故種類填寫錯誤!");
			myForm.TrafficAccidentType.select();
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

function GetBillNo(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91))
	{
		myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
		if (myForm.Billno1.value.length == 3 )
		{
			myForm.Billno2.select();
		}			
	if (myForm.Billno1.value.length == 9 )
	{
		myForm.Billno2.value=myForm.Billno1.value.substr(3,9);
		myForm.Billno1.value=myForm.Billno1.value.substr(0,3);
    	myForm.Billno2.focus();
	}		
	}

}

function GetIllegalDate(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91))
	{
	
		if (myForm.IllegalDate.value.length == 6 )
		{
			myForm.IllegalTime.select();
		}			
	
	}

}

function GetIllegalTime(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91))
	{
	
		if (myForm.IllegalTime.value.length == 4 )
		{
			myForm.IllegalAddressID.select();
		}			
	
	}

}

function GetBillNo2(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91))
	{
		myForm.Billno2.value=myForm.Billno2.value.toUpperCase();
		if (myForm.Billno2.value.length == 6 && myForm.Billno1.value!="")
		{
			
			myForm.CarNo.select();
		
		}		
		if (myForm.Billno2.value.length == 9 && myForm.Billno1.value=="")
		{
		
		myForm.Billno1.value=myForm.Billno2.value.substr(0,3);
		myForm.Billno2.value=myForm.Billno2.value.substr(3,9);		
		
		}			
	
	}

}

function GetCarSimpleID(){
	if ((event.keyCode>47 && event.keyCode<58) || (event.keyCode>95 && event.keyCode<106) || (event.keyCode>64 && event.keyCode<91))
	{
		if (myForm.CarSimpleID.value.length == 1 )
		{
			myForm.DriverPID.select();
		}			
	
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
*/
/*
function CallChkLaw4(){
	if (TDLawNum==2){
		if (!funcChkLaw(myForm.Rule4.value)){
			alert("請確認法條四是否填寫正確");
		}	
	}
}
*/
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
function FuncChkPID(){
	myForm.DriverPID.value=myForm.DriverPID.value.toUpperCase();
	myForm.DriverPID.value=myForm.DriverPID.value.replace(/[\s　]+/g, "");
	if (myForm.DriverPID.value!=""){
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
	}else{
		alert("身分證輸入錯誤！");
		//myForm.DriverPID.focus();
	}}
}

function focusToCarNo(){
	//myForm.Insurance.value=myForm.Insurance.value.replace(/[^\d]/g,'');
	if (myForm.Insurance.value.length=="1"){
		if 	(myForm.Insurance.value != "0" && myForm.Insurance.value != "1" && myForm.Insurance.value != "2" && myForm.Insurance.value != "3" && myForm.Insurance.value != "4"){
			alert("保險證輸入錯誤！");
			myForm.Insurance.select();
		}
	}
}
function KeyDown(){ 
	if (event.keyCode==117){	//F6鎖死
		event.keyCode=0;   
		event.returnValue=false;   
<%if sys_City="台東縣" then%>
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
		location='BillKeyIn_TakeCar.asp'
	}else if (event.keyCode==116){ //F5查詢
		event.keyCode=0;   
		funcOpenBillQry();
		event.returnValue=false;   
	}else if (event.keyCode==119){ //F8關閉
		event.keyCode=0;   
		window.close();
	}else if (event.keyCode==121){ //F10查詢未建檔
		event.keyCode=0;   
		funPrintCaseList_Stop();
	}else if (event.keyCode==33){ //上一筆PageUp
		event.keyCode=0;   
		location='BillKeyIn_TakeCar_Back.asp?PageType=Back'
	}else if (event.keyCode==36){ //第一筆Home
		event.keyCode=0;   
		location='BillKeyIn_TakeCar_Back.asp?PageType=First'
	}
}
function funcOpenBillQry(){
	Sys_RMemberID=<%=session("User_ID")%>;
	Sys_BillTypeID=1;
	window.open("EasyTakeCarBillQry.asp","WebPage86_Update","left=300,top=200,location=0,width=350,height=160,resizable=yes,scrollbars=yes");
}
function AutoGetIllStreet(){	//按F6可以直接顯示相關路段
	if (event.keyCode==117){	
		event.keyCode=0;
		Ostreet=myForm.IllegalAddress.value;
		window.open("Query_Street.asp?OStreet="+Ostreet,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
}
function AutoGetRuleID(LawOrder){	//按F6可以直接顯示相關法條
	if (event.keyCode==117){	
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
function focusToDriverPID(){
	myForm.DriverBrith.value=myForm.DriverBrith.value.replace(/[^\d]/g,'');
	if (myForm.DriverBrith.value.length==6){
		//var x=new Date();
		//var thisYear=x.getYear()-1911;
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
	//攔停建檔清冊
	function funPrintCaseList_Stop(){
		UrlStr="../Query/PrintCaseDataList_Stop.asp?CallType=1";
		newWin(UrlStr,"CaseListWin",980,575,0,0,"yes","yes","yes","no");
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
	//-------按Enter到下一欄--------
	function OnBlurNextTag(tag1){
		if (event.keyCode==13){	
			event.keyCode=0;
			eval("myForm."+tag1).select();
		}
	}
	function OnKeyUpNextTag(tag1){
		eval("myForm."+tag1).select();
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
		}*/else if (event.keyCode==38){ //上換欄
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
	//------------------------------
myForm.Billno2.focus();

if (myForm.MemberStation.value.length > 1){
		var StationNum=myForm.MemberStation.value;
		runServerScript("getMemberStation.asp?StationID="+StationNum);
	}

</script>
</html>

