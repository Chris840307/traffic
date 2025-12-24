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
<title>攔停資料打驗校對作業</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
'AuthorityCheck(236)
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

'新增告發單
if trim(request("kinds"))="DB_insert" then
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
		theCarAddID="null"
	else
		theCarAddID=trim(request("CarAddID"))
	end if

	'查流水號
	strSN="select BillBaseTemp_seq.nextval as SN from Dual"
	set rsSN=conn.execute(strSN)
	if not rsSN.eof then
		theSN=trim(rsSN("SN"))
	end if
	rsSN.close
	set rsSN=nothing

	'BillBase
	strInsert="insert into BillBaseTmp(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
				",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
				",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
				",BillMemID2,BillMem2,BillMemID3,BillMem3,BillMemID4,BillMem4" &_
				",BillFillerMemberID,BillFiller" &_
				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
				",Note,EquipmentID,RuleVer,DriverSex,TrafficAccidentNo,TrafficAccidentType,SignType)" &_
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
				",'"&trim(request("Note"))&"','"&trim(request("FixID"))&"','"&theRuleVer&"'" &_
				",'"&trim(request("DriverSex"))&"','"&trim(request("TrafficAccidentNo"))&"','"&trim(request("TrafficAccidentType"))&"','"&UCase(trim(request("SignType")))&"'" &_
				")"
				conn.execute strInsert
				'theDriverBirth , theBillFillDate   

	'舉發單扣件明細檔 BillFastenerDetail
	if trim(request("Fastener1"))<>"" then
		strInsFastene1="insert into BillFastenerDetailTemp(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
					" values(BillFastenerDetailTemp_seq.nextval,"&theSN&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener1"))&"','')"
		conn.execute strInsFastene1
	end if
	if trim(request("Fastener2"))<>"" then
		strInsFastene2="insert into BillFastenerDetailTemp(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
					" values(BillFastenerDetailTemp_seq.nextval,"&theSN&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener2"))&"','')"
		conn.execute strInsFastene2
	end if
	if trim(request("Fastener3"))<>"" then
		strInsFastene3="insert into BillFastenerDetailTemp(SN,BillSN,CarNo,FastenerTypeID,Fastener)" &_
					" values(BillFastenerDetailTemp_seq.nextval,"&theSN&",'"&UCase(trim(request("CarNo")))&"','"&trim(request("Fastener3"))&"','')"
		conn.execute strInsFastene3
	end if
%>
<script language="JavaScript">
	location='BillBaseDoubleCheck_Car.asp?BillNo=<%=UCase(trim(request("Billno1")))%>';
</script>
<%
end if
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
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
	<form name="myForm" method="post">  
		<table width='985' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="6"><strong>攔停資料打驗校對作業</strong>&nbsp; &nbsp; 日期格式：951220 &nbsp;時間格式：2300(24小時制)&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 建檔日期：<%=ginitdt(now)%></td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>單號</div></td>
				<td>
					<input name="Billno1" type="text" value="<%=theBillno%>" size="10" maxlength="9" onBlur="CheckBillNoExist();"></td>
				<td bgcolor="#FFFFCC"><div align="right">保險證</div></td>
				<td align="left" colspan="3">
				    <input type="text" maxlength="1" size="3" value="" name="Insurance" onKeyUp="focusToCarNo();">
					<div id="Layer111" style="position:absolute; width:470px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
					<font color="#ff000" size="2">0有出示/ 1未出示/ 2肇事且未出示/ 3逾期或未保險/ 4肇事且逾期或未保險</font>
					</div>
				</td>
			  </td>
			</tr>
			<tr>
				<!-- <td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規人姓名</div></td>
				<td>
			    <input type="text" size="13" name="DriverName">
				</td> -->
				<!-- <td bgcolor="#FFFFCC"><div align="right">違規人地址</div></td>
				<td colspan="3"> -->
				<!-- <input type="text" size="5" name="DriverZip"> -->
				<!-- <input type="text" size="40" value="" name="DriverAddress">
				</td> -->
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規人證號</div></td>
				<td>
				<input type="text" size="10" name="DriverPID" onkeyup="value=value.toUpperCase()" onBlur="FuncChkPID();">
				</td>
				<td bgcolor="#FFFFCC" align="right">違規人出生日</td>
				<td colspan="3">
				<input type="text" size="10" maxlength="6" name="DriverBrith" onkeyup="focusToDriverPID()">
				</td>
				
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規車號</div></td>
				<td> 
				<input type="text" size="10" name="CarNo" maxlength="8" onkeyup="value=value.toUpperCase()" onBlur="getVIPCar();">
			    <div id="Layer7" style="position:absolute; width:115px; height:24px; z-index:0;  border: 1px none #000000; color: #FF0000; font-weight: bold;">
				</div>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>簡式車種</div></td>
				<td colspan="3">
				<input type="text" maxlength="1" size="3" value="" name="CarSimpleID" onkeyup="getRuleAll();">
				<div id="Layer111" style="position:absolute; width:155px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;">
				<font color="#ff000" size="2"> 1汽車/ 2拖車/ 3重機/ 4輕機</font>
				</div>
				</td>
			</tr>
			<tr>
				
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規日期</div></td>
				<td>
				<input type="text" size="10" maxlength="6" name="IllegalDate" onkeyup="getDealLineDate()">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規時間</div></td>
				<td colspan="3">
				<input type="text" size="10" maxlength="4" name="IllegalTime" onkeyup="value=value.replace(/[^\d]/g,'')">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">違規地點代碼</div></td>
				<td>
					<input type="text" size="8" value="<%=request("IllegalAddressID")%>" name="IllegalAddressID" onkeyup="getillStreet();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規地點</div></td>
				<td colspan="3">
					<input type="text" size="40" value="<%=trim(request("IllegalAddress"))%>" name="IllegalAddress">
				</td>

			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規法條一</div></td>
				<td colspan="5">
					<input type="text" maxlength="8" size="10" value="<%=request("Rule1")%>" name="Rule1" onKeyUp="getRuleData1();" onblur="AutoKeyCarNo()">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes")' alt="查詢法條">
					<img src="../Image/BillLawPlusButton.jpg" width="25" height="23" onclick="Add_LawPlus()" alt="附加說明">
					<div id="Layer1" style="position:absolute ; width:580px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
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
					<input type="hidden" name="ForFeit1" value="<%=request("ForFeit1")%>">
				</td>
				
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">違規法條二</div></td>
				<td colspan="5">
					<input type="text" maxlength="8" size="10" value="<%=request("Rule2")%>" name="Rule2" onKeyUp="getRuleData2();" onBlur="TabFocus();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=850,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer2" style="position:absolute ; width:609px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
					if trim(request("Rule2"))<>"" then
						strRule2="select IllegalRule from Law where ItemID='"&trim(request("Rule2"))&"' and Version='"&trim(theRuleVer)&"'"
						set rsRule2=conn.execute(strRule2)
						if not rsRule2.eof then
							response.write trim(rsRule2("IllegalRule"))
						end if
						rsRule2.close
						set rsRule2=nothing
					end if
					%></div>
					<input type="hidden" name="ForFeit2" value="<%=request("ForFeit2")%>">
				</td>
				
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">限速、限重</div></td>
				<td>
					<input type="text" size="10" name="RuleSpeed" onBlur="RuleSpeedforLaw()" >
				</td>
				<td bgcolor="#FFFFCC"><div align="right">車速、車重</div></td>
				<td colspan="3">
					<input type="text" size="10" name="IllegalSpeed" onBlur="IllegalSpeedforLaw()" >
				</td>
				
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>應到案日期</div></td>
				<td>
					<input type="text" size="10" maxlength="6" name="DealLineDate" onkeyup="value=value.replace(/[^\d]/g,'')">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>應到案處所</div></td>
				<td>
					<input type="text" size="5" value="" name="MemberStation" onKeyup="getStation();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Station.asp","WebPage1","left=0,top=0,location=0,width=760,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer5" style="position:absolute ; width:120px; height:30px; z-index:0;  border: 1px none #000000;"></div>
					</span>
				</td>

				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發人代碼1</div></td>
		  		<td>
					<input type="text" size="5" name="BillMem1" onkeyup="getBillMemID1();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer12" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID1">
					<input type="hidden" value="" name="BillMemName1">
				</td>
			</tr>
			<tr>
				
				<td bgcolor="#FFFFCC" width="12%"><div align="right">舉發人代碼2</div></td>
				<td width="22%">
					<input type="text" size="5" value="" name="BillMem2" onkeyup="getBillMemID2();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer13" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID2">
					<input type="hidden" value="" name="BillMemName2">
				</td>
				<td bgcolor="#FFFFCC" width="12%"><div align="right">舉發人代碼3</div></td>
				<td width="22%">
					<input type="text" size="5" value="" name="BillMem3" onkeyup="getBillMemID3();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=3","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer14" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID3">
					<input type="hidden" value="" name="BillMemName3">
				</td>
				<td bgcolor="#FFFFCC" width="12%"><div align="right">舉發人代碼4</div></td>
		  		<td width="20%">
					<input type="text" size="5" name="BillMem4" onkeyup="getBillMemID4();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=4","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer17" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID4">
					<input type="hidden" value="" name="BillMemName4">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">代保管物1</div></td>
				<td><input type="text" size="5" value="" name="Fastener1" onkeyup="getFastener1();">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=1","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                  <div id="Layer8" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"></div>
                  <input type="hidden" value="" name="Fastener1Val">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">代保管物2</div></td>
				<td>
                  <input type="text" size="5" value="" name="Fastener2" onkeyup="getFastener2();">
				  <img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=2","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                  <div id="Layer9" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"></div>
                  <input type="hidden" value="" name="Fastener2Val">
                </td>
				<td bgcolor="#FFFFCC"><div align="right">代保管物3</div></td>
				<td>
				  <input type="text" size="5" value="" name="Fastener3" onKeyUp="getFastener3();">
				  <img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Fastener.asp?FaOrder=3","FastPage","left=0,top=0,location=0,width=500,height=355,resizable=yes,scrollbars=yes")'>
                  <div id="Layer10" style="position:absolute ; width:90px; height:30px; z-index:0;  border: 1px none #000000;"></div>
                  <input type="hidden" value="" name="Fastener3Val">
				</td>
			</tr>
			<tr>
				<td></td>
			</tr>
			<tr>
				<td></td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">舉發單位</div></td>
				<td >
					<input type="text" size="5" name="BillUnitID" onKeyUp="getUnit();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=700,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer6" style="position:absolute ; width:120px; height:30px; z-index:0;  border: 1px none #000000;"></div>
					</span>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>簽收狀況</div></td>
				<td >
					<input type="text" size="5" value="A" maxlength="1" name="SignType" onBlur="funcSignType();" style=ime-mode:disabled>
					<font color="#ff000" size="2">
					A:簽收 / U:拒收
					</font>
				</td>	
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>填單日期</div></td>
				<td>
					<input type="text" size="10" value="" maxlength="6" name="BillFillDate" onkeyup="value=value.replace(/[^\d]/g,'')">
				</td>
				
			</tr>
			<tr>
			    <td bgcolor="#FFFFCC" align="right">輔助車種</td>
				<td colspan="3">
				<input type="text" maxlength="1" size="3" value="" name="CarAddID" onKeyUp="getAddID();">
				<font color="#ff000" size="2">1大貨/ 2大客/ 3砂石/ 4土方/ 5動力/ 6貨櫃/ 7大型重機</font>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">專案代碼</div></td>
				<td ><input type="text" size="10" value="" name="ProjectID">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onClick='window.open("Query_Project.asp","WebPage1","left=0,top=0,location=0,width=700,height=455,resizable=yes,scrollbars=yes")'>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">備註</div></td>
				<td>
					<input type="text" size="20" value="" name="Note">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">交通事故案號</div></td>
				<td>
					<input type="text" size="16" name="TrafficAccidentNo" Value="">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">交通事故種類</div></td>
				<td>
					<input type="text" maxlength="1" size="5" name="TrafficAccidentType" Value="" onKeyUp="chkTrafficAccidentType();">
					<font color="#ff000" size="2"> 1 / 2 / 3</font>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" align="center" colspan="6">
					<input type="button" value="儲 存 F2" onclick="InsertBillVase();" <%
				'1:查詢 ,2:新增 ,3:修改 ,4:刪除
				if CheckPermission(236,2)=false then
					response.write "disabled"
				end if
					%> class="btn1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit1343" onClick="location='BillKeyIn_Car2.asp'" value="清 除 F4" class="btn1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit5322" onClick="funcOpenBillQry();" value="查 詢 F6" class="btn1">
					<input type="hidden" name="kinds" value="">
                    <span class="style1">
                    <span class="style3"><img src="/image/space.gif" width="29" height="8"></span>
                    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" class="btn1">
                </span>
				<!-- 告發類別 -->
				<input type="hidden" size="3" maxlength="1" value="<%=theBilltype%>" name="BillType">		
				<!-- 違規人性別 -->
				<input type="hidden" name="DriverSex" value="">
				<!-- 附加說明 -->
				<input type="hidden" value="" name="Rule4">
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
//新增告發單
function InsertBillVase(){
	var error=0;
	var errorString="";
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value;
	var TodayDate=<%=ginitdt(date)%>;
	if (myForm.Billno1.value=="" && myForm.BillType.value!="2"){
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
	/*
	if (myForm.DriverName.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規人姓名。";
	}
	*/
	if (myForm.DriverPID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規人身份證號。";
	}
	if (myForm.DriverPID.value!=""){
		if(myForm.DriverPID.value.length > 10){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規人身份證號輸入錯誤。";	
		}else if (myForm.DriverPID.value.length < 10){
			myForm.DriverBrith.value="";
		}
	}
	if (myForm.CarNo.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規車號。";
	}else if (chkCarNoFormat(myForm.CarNo.value)==0 && ((Rule1tmp.substr(0,2))!="32" && (Rule2tmp.substr(0,2))!="32" && (Rule1tmp.substr(0,3))!="334" && (Rule2tmp.substr(0,3))!="334")){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規車號格式錯誤。";
	}
	if (myForm.CarSimpleID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入簡式車種。";
	}else if(myForm.CarNo.value != "" && chkCarNoFormat(myForm.CarNo.value)!= 0) {
		if (chkCarNoFormat(myForm.CarNo.value) != myForm.CarSimpleID.value){
			error=error+1;
			errorString=errorString+"\n"+error+"：車號格式與簡式車種不符。";
		}
	}
	if (myForm.IllegalDate.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規日期。";
	}else if(!dateCheck( myForm.IllegalDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
	}else if (!ChkIllegalDate(myForm.IllegalDate.value) && myForm.TrafficAccidentNo.value==""){
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
//	if (TDLawNum==2 && myForm.Rule4.value!=""){
//		if(myForm.Rule1.value==myForm.Rule4.value){
//			error=error+1;
//			errorString=errorString+"\n"+error+"：違規法條一與違規法條四重複。";
//		}
//		if (myForm.Rule2.value==myForm.Rule4.value){
//			error=error+1;
//			errorString=errorString+"\n"+error+"：違規法條二與違規法條四重複。";
//		}
//		if (myForm.Rule3.value==myForm.Rule4.value){
//			error=error+1;
//			errorString=errorString+"\n"+error+"：違規法條三與違規法條四重複。";
//		}
//	}
	if (myForm.Rule2.value!=""){
		if (TDLawErrorLog2==1){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二輸入錯誤。";
		}else if (myForm.Rule2.value.substr(0,2)>68){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條一輸入錯誤。";
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
	}else if (!ChkIllegalDate(myForm.BillFillDate.value) && myForm.TrafficAccidentNo.value==""){
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
	}else if (!ChkIllegalDate(myForm.DealLineDate.value) && myForm.TrafficAccidentNo.value==""){
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
	if (error==0){
		myForm.kinds.value="DB_insert";
		myForm.submit();
	}else{
		alert(errorString);
	}
}
//增加違規法條
//function InsertLaw(){
//	if (TDLawNum==0){
//		TDLaw1.innerHTML="違規法條三";
//		TDLaw2.innerHTML="<input type='text' size='10' value='' name='Rule3' onKeyUp='getRuleData3();'> <input type='button' value='？' name='LawSelect' onclick='LawOpen3();'> <div id='Layer3' style='position:absolute ; width:469px; height:24px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;'></div><input type='hidden' name='ForFeit3' value=''>";
//		TDLawNum=TDLawNum+1;
//		myForm.Lawdel.disabled=false;
//	}else if (TDLawNum==1){
//		TDLaw3.innerHTML="違規法條四";
//		TDLaw4.innerHTML="<input type='text' size='10' value='' name='Rule4' onKeyUp='getRuleData4();'> <input type='button' value='？' name='LawSelect' onclick='LawOpen4();'> <div id='Layer4' style='position:absolute ; width:469px; height:24px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;'></div><input type='hidden' name='ForFeit4' value=''>";
//		TDLawNum=TDLawNum+1;
//		myForm.Lawadd.disabled=true;
//	}
//}
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
//是否為特殊用車
function getVIPCar(){
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value;
	myForm.CarNo.value=myForm.CarNo.value.replace(" ", "");
	if (myForm.CarNo.value.length >= 4  && ((Rule1tmp.substr(0,2))!="32" && (Rule2tmp.substr(0,2))!="32" && (Rule1tmp.substr(0,3))!="334" && (Rule2tmp.substr(0,3))!="334")){
		var CarNum=myForm.CarNo.value;
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType==0){
			alert("車牌格式錯誤");
			myForm.CarNo.focus();
		}else{
			//runServerScript("getVIPCar.asp?CarID="+CarNum);
			myForm.CarSimpleID.value=CarType;
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
		if (myForm.CarAddID.value != "1" && myForm.CarAddID.value != "2" && myForm.CarAddID.value != "3" && myForm.CarAddID.value != "4" && myForm.CarAddID.value != "5" && myForm.CarAddID.value != "6" && myForm.CarAddID.value != "7"){
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
		if (myForm.CarSimpleID.value != "1" && myForm.CarSimpleID.value != "2" && myForm.CarSimpleID.value != "3" && myForm.CarSimpleID.value != "4"){
			alert("簡式車種填寫錯誤!");
			myForm.CarSimpleID.select();
			//myForm.CarSimpleID.value = "";
		}
	}
}
//違規事實1(ajax)
function getRuleData1(){
	if (myForm.Rule1.value.length > 6){
		var Rule1Num=myForm.Rule1.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail.asp?RuleOrder=1&RuleID="+Rule1Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		CallChkLaw1();
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
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail.asp?RuleOrder=2&RuleID="+Rule2Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
		CallChkLaw2();
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
function TabFocus(){
	//建檔時除了超重超速時游標才跳至限速限量欄位，其它法條則游標不跳至超重超速
	Rule1tmp=myForm.Rule1.value;
	Rule2tmp=myForm.Rule2.value
		if ((Rule1tmp.substr(0,2))!="33" && (Rule1tmp.substr(0,2))!="40" && (Rule1tmp.substr(0,2))!="43" && (Rule1tmp.substr(0,2))!="29" && (Rule2tmp.substr(0,2))!="33" && (Rule2tmp.substr(0,2))!="40" && (Rule2tmp.substr(0,2))!="43" && (Rule2tmp.substr(0,2))!="29"){
			myForm.DealLineDate.focus();
		}
	//法條遇到32 與DCI 傳輸固定用身分證號前八碼
	AutoKeyCarNo();
}
function AutoKeyCarNo(){
	//法條遇到32 與DCI 傳輸固定用身分證號前八碼
	Rule1tmp=myForm.Rule1.value.substr(0,3);
	Rule2tmp=myForm.Rule2.value.substr(0,3);
	if (Rule1tmp=="320" || Rule2tmp=="320" || Rule1tmp=="321" || Rule2tmp=="321" || Rule1tmp=="322" || Rule2tmp=="322" || Rule1tmp=="334" || Rule2tmp=="334"){
		myForm.CarNo.value=myForm.DriverPID.value.substr(0,6);
	}
	MemberStationLaw="21,35,57,61,62";
	//法條代碼遇到21,35,57,61,62，應到案處所自動帶當地監理所
	if (((MemberStationLaw.indexOf(Rule1tmp)!=-1 || MemberStationLaw.indexOf(Rule2tmp)!=-1) && Rule1tmp !="" && Rule2tmp !="") || (MemberStationLaw.indexOf(Rule1tmp)!=-1 && Rule2tmp =="" && Rule1tmp !="") || (MemberStationLaw.indexOf(Rule2tmp)!=-1 && Rule1tmp =="" && Rule2tmp !="")){
		myForm.MemberStation.value=<%=trim(BillLawMemberStation)%>;
		getStation();
	}
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
//function getRuleData4(){
//	if (myForm.Rule4.value.length > 6){
//		var Rule4Num=myForm.Rule4.value;
//		var CarSimpleID=myForm.CarSimpleID.value;
//		var VerNo=<%=theRuleVer%>;
//		runServerScript("getRuleDetail.asp?RuleOrder=4&RuleID="+Rule4Num+"&CarSimpleID="+CarSimpleID+"&RuleVer="+VerNo);
//		//CallChkLaw4();
//	}else if (myForm.Rule4.value.length <= 6 && myForm.Rule4.value.length > 0){
//		Layer4.innerHTML=" ";
//		myForm.ForFeit4.value="";
//		TDLawErrorLog4=1;
//	}else{
//		Layer4.innerHTML=" ";
//		myForm.ForFeit4.value="";
//		TDLawErrorLog4=0;
//	}
//}
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
	if (myForm.BillUnitID.value.length > 1){
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
function UserInputBillType(){

}
//違規地點代碼(ajax)
function getillStreet(){
	if (myForm.IllegalAddressID.value.length > 2){
		var illAddrNum=myForm.IllegalAddressID.value;
		runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
	}
}
//舉發人一(ajax)
function getBillMemID1(){
	if (myForm.BillMem1.value.length > 2){
		var BillMemNum=myForm.BillMem1.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=1&MemID="+BillMemNum);
	}else if (myForm.BillMem1.value.length <= 2 && myForm.BillMem1.value.length > 0){
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
	if (myForm.BillMem2.value.length > 2){
		var BillMemNum=myForm.BillMem2.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=2&MemID="+BillMemNum);
	}else if (myForm.BillMem2.value.length <= 2 && myForm.BillMem2.value.length > 0){
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
	if (myForm.BillMem3.value.length > 2){
		var BillMemNum=myForm.BillMem3.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=3&MemID="+BillMemNum);
	}else if (myForm.BillMem3.value.length <= 2 && myForm.BillMem3.value.length > 0){
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
	if (myForm.BillMem4.value.length > 2){
		var BillMemNum=myForm.BillMem4.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=4&MemID="+BillMemNum);
	}else if (myForm.BillMem4.value.length <= 2 && myForm.BillMem4.value.length > 0){
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
//攔停由違規日期帶入應到案日期+14
function getDealLineDate(){
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');
	BFillDateTemp=myForm.IllegalDate.value;
	if (BFillDateTemp.length >= 6){
		myForm.BillFillDate.value=myForm.IllegalDate.value;
		Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
		Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
		Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
		var BFillDate=new Date(Byear,Bmonth-1,Bday)
		var DLineDate=new Date()
		DLineDate=DateAdd("d",14,BFillDate);
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
	var BillNum=myForm.Billno1.value;
	if (myForm.Billno1.value!=""){
		if (myForm.Billno1.value.length < 10 && myForm.Billno1.value.length > 8 ){
			runServerScript("getDoubleCheckBillNoExist.asp?BillNo="+BillNum);
		}else{
			alert("單號不足九碼！");
			myForm.Billno1.select();
		}
	}
}
function chkTrafficAccidentType(){
	//myForm.TrafficAccidentType.value=myForm.TrafficAccidentType.value.toUpperCase();
	if (myForm.TrafficAccidentType.value.length >= 1){
		if (myForm.TrafficAccidentType.value!="1" && myForm.TrafficAccidentType.value!="2" && myForm.TrafficAccidentType.value!="3"){
			alert("交通事故種類填寫錯誤!");
			//myForm.TrafficAccidentType.value = "";
			myForm.TrafficAccidentType.select();
		}
	}
}
//單號驗證
function setCheckDoubleBillNoExist(GetBillFlag,BillBaseFlag,BillBaseTmpFlag,MLoginID,MMemberID,MMemName,MUnitID,MUnitName)
{
	if (GetBillFlag==0){
		//alert("此單號不存在於領單紀錄中！");
		//document.myForm.Billno1.value="";
	}else{
		if (document.myForm.BillMem1.value==""){
			document.myForm.BillMem1.value=MLoginID;
			document.myForm.BillMemID1.value=MMemberID;
			document.myForm.BillMemName1.value=MMemName;
			Layer12.innerHTML=MMemName;
			TDMemErrorLog1=0;
		}
		if (document.myForm.BillUnitID.value==""){
			document.myForm.BillUnitID.value=MUnitID;
			Layer6.innerHTML=MUnitName;
			TDUnitErrorLog=0;
		}
		if (BillBaseFlag==0){
			alert("此單號尚未建檔！");
			document.myForm.Billno1.value="";
		}else if (BillBaseFlag==1){
			if (BillBaseTmpFlag==0){
				alert("此單號已打驗！");
				document.myForm.Insurance.value="";
			}
		}
	}
}
function RuleSpeedforLaw(){
	myForm.RuleSpeed.value=myForm.RuleSpeed.value.replace(/[^\d]/g,'');
	CallChkLaw1();
	CallChkLaw2();
	//CallChkLaw3();
	//CallChkLaw4();
	if (myForm.RuleSpeed.value > 100){
		alert("限速、限重超過100，請確認是否正確!");
	}
}
function IllegalSpeedforLaw(){
	myForm.IllegalSpeed.value=myForm.IllegalSpeed.value.replace(/[^\d]/g,'');
	CallChkLaw1();
	CallChkLaw2();
	//CallChkLaw3();
	//CallChkLaw4();
	if (myForm.IllegalSpeed.value > 100){
		alert("車速超過100，請確認是否正確!");
	}
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
			if ((thisLaw.substr(0,2))!="33" && (thisLaw.substr(0,2))!="40" && (thisLaw.substr(0,2))!="43" && (thisLaw.substr(0,2))!="29"){
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
function FuncChkPID(){
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
	}else if (myForm.DriverPID.value.length > 0 && myForm.DriverPID.value.length < 10){
		alert("身分證輸入錯誤！");
		//myForm.DriverPID.focus();
	}
}
function focusToCarNo(){
	//myForm.Insurance.value=myForm.Insurance.value.replace(/[^\d]/g,'');
	if (myForm.Insurance.value!=""){
		if 	(myForm.Insurance.value != "0" && myForm.Insurance.value != "1" && myForm.Insurance.value != "2" && myForm.Insurance.value != "3" && myForm.Insurance.value != "4"){
			alert("保險證輸入錯誤！");
			myForm.Insurance.select();
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
//用地點、車速抓違規法條
//function setIllegalRule(){
//	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!="" && myForm.IllegalAddress.value!=""){
//		if ((myForm.Rule1.value.substr(0,2))!="29"){
//			IllegalRule=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value);
//			if (IllegalRule!="Null"){
//				myForm.Rule1.value=IllegalRule;
//				getRuleData1()
//			}
//		}
//	}
//}
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
		InsertBillVase();
	}else if (event.keyCode==115){ //F4清除
		event.keyCode=0;   
		location='BillKeyIn_Car2.asp'
	}else if (event.keyCode==117){ //F6查詢
		event.keyCode=0;   
		funcOpenBillQry();
	}else if (event.keyCode==119){ //F8關閉
		event.keyCode=0;   
		window.close();
	}
}
	//簽收狀況(小轉大寫，限定A or U)
	function funcSignType(){
		myForm.SignType.value=myForm.SignType.value.toUpperCase();
		if (myForm.SignType.value==""){
			myForm.SignType.focus();
			alert("簽收狀況未填寫!!");
		}else if (myForm.SignType.value!="A" && myForm.SignType.value!="U"){
			myForm.SignType.select();
			alert("簽收狀況填寫錯誤!!");
		}
	}
function funcOpenBillQry(){
	Sys_RMemberID=<%=session("User_ID")%>;
	Sys_BillTypeID=1;
	window.open("../Query/BillBaseQry.asp?QryType=1&Sys_RecordMemberID="+Sys_RMemberID+"&Sys_BTypeID="+Sys_BillTypeID,"WebPage4_Update","left=0,top=0,location=0,width=1000,height=660,resizable=yes,scrollbars=yes");
}
	//附加說明
	function Add_LawPlus(){
		if (myForm.Rule1.value==""){
			alert("請先輸入違規法條一!!");
		}else{
		window.open("Query_LawPlus.asp","WebPage1","left=20,top=10,location=0,width=500,height=455,resizable=yes,scrollbars=yes");
		}
	}
myForm.Billno1.focus();

</script>
</html>
