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
<title>逕舉資料打驗校對作業</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
'AuthorityCheck(223)
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
'==========cookie==========
'填單人代碼
theRecordMemberID=trim(Session("User_ID"))
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
	'先檢查是否有這筆車號
	BillCarNoCheck1="0"
	strBill1="select CarNo from BillBase where BillTypeID='2' and CarNo='"&UCase(trim(request("CarNo")))&"'"
	set rsBill1=conn.execute(strBill1)
	if not rsBill1.eof then
		BillCarNoCheck1="1"
	else
		BillCarNoCheck1="0"
	end if
	rsBill1.close
	set rsBill1=nothing

	if BillCarNoCheck1="1" then
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
			theCarAddID="null"
		else
			theCarAddID=trim(request("CarAddID"))
		end if
		'BillBaseTmp
		strInsert="insert into BillBaseTmp(SN,BillTypeID,BillNo,CarNo,CarSimpleID,CarAddID,IllegalDate" & _
					",IllegalAddressID,IllegalAddress,Rule1,IllegalSpeed,RuleSpeed,ForFeit1" &_
					",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,Insurance,UseTool,ProjectID" &_
					",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
					",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
					",BillMemID2,BillMem2,BillMemID3,BillMem3,BillMemID4,BillMem4" &_
					",BillFillerMemberID,BillFiller" &_
					",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
					",Note,EquipmentID,RuleVer,DriverSex)" &_
					" values(BillBase_seq.nextval,'"&trim(request("BillType"))&"','"&UCase(trim(request("Billno1")))&"'" &_
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
					",'"&trim(request("DriverSex"))&"'" &_
					")"
		
					conn.execute strInsert
					'theDriverBirth , theBillFillDate   
%>
<script language="JavaScript">
	location='BillBaseDoubleCheck_Report.asp?CarNo=<%=UCase(trim(request("CarNo")))%>';
</script>
<%
	else
%>
<script language="JavaScript">
	alert("此車牌號碼尚未建檔!");
	//location='BillKeyIn_Car_Report2.asp';
</script>
<%		
	end if
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
.style5 {
	font-size: 12px
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
	<form name="myForm" method="post">  
		<table width='985' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="6"><strong>逕舉資料打驗校對作業</strong>&nbsp; &nbsp; 日期格式：951220 &nbsp;時間格式：2300(24小時制)&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 建檔日期：<%=ginitdt(now)%>
				<br>
				<input type="checkbox" name="ReportChk" value="1" onclick="funcReportChk();" <%
				if bBillType="1" then
					response.write "checked"
				end if
				%>>攔停手開單&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
				<input type="checkbox" name="CaseInByMem" value="1">人工入案(不檢查違規日期)
				</td>
			</tr>
			<tr>
			  <td bgcolor="#FFFFCC"><div align="right">單號</div></td>
				<td colspan="5">
				<input name="Billno1" type="text" value="<%=theBillno%>" size="12" maxlength="9" OnBlur="CheckBillNoExist();" disabled>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" width="13%"><div align="right"><span class="style4">＊</span>違規車號</div></td>
				<td width="35%">
				<input type="text" size="12" name="CarNo" onkeyup="value=value.toUpperCase()" onBlur="getVIPCar();" maxlength="8">
			    <div id="Layer7" style="position:absolute; width:140px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold;"></div>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>簡式車種</div></td>
				<td colspan="3">
				<input type="text" maxlength="1" size="4" value="" name="CarSimpleID" onkeyup="getRuleAll();">
				<font color="#ff000" size="2">1汽車 / 2拖車/ 3重機/ 4輕機</font>
				</td>

			</tr>

			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規日期</div></td>
				<td>
				<input type="text" size="10" maxlength="6" name="IllegalDate">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規時間</div></td>
				<td colspan="2">
				<input type="text" size="4" maxlength="4" name="IllegalTime" onkeyup="value=value.replace(/[^\d]/g,'')">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" width="13%"><div align="right"><span class="style4">＊</span>違規地點代碼</div></td>
				<td>
					<input type="text" size="8" value="<%=request("IllegalAddressID")%>" name="IllegalAddressID" onkeyup="getillStreet();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規地點</div></td>
				<td colspan="2">
					<input type="text" size="40" value="<%=trim(request("IllegalAddress"))%>" name="IllegalAddress">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>違規法條一</div></td>
				<td colspan="4">
					<input type="text" maxlength="8" size="10" value="<%=request("Rule1")%>" name="Rule1" onKeyUp="getRuleData1();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=1&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer1" style="position:absolute ; width:560px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
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
					<input type="text" maxlength="8" size="10" value="<%=request("Rule2")%>" name="Rule2" onKeyUp="getRuleData2();" onBlur="TabFocus()">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=theRuleVer%>","WebPage1","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer2" style="position:absolute ; width:560px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
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
				<td colspan="2">
					<input type="text" size="10" name="IllegalSpeed" onBlur="IllegalSpeedforLaw()" >
				</td>
			</tr>
			<tr>
				<td id="DLDate1" bgcolor="#FFFFCC" align="right"></td>
				<td id="DLDate2" colspan="4">
				<input type="hidden" size="6" value="" maxlength="6" name="DealLineDate" onBlur="DealLineDateReplace()" style=ime-mode:disabled>
								
				</td>

		</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發人代碼1</div></td>
		  		<td>
					<input type="text" size="10" name="BillMem1" onkeyup="getBillMemID1();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer12" style="position:absolute ; width:130px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID1">
					<input type="hidden" value="" name="BillMemName1">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">舉發人代碼2</div></td>
		  		<td>
					<input type="text" size="10" name="BillMem2" onkeyup="getBillMemID2();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer13" style="position:absolute ; width:130px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID2">
					<input type="hidden" value="" name="BillMemName2">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC"><div align="right">舉發人代碼3</div></td>
		  		<td>
					<input type="text" size="10" name="BillMem3" onkeyup="getBillMemID3();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=3","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer14" style="position:absolute ; width:130px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID3">
					<input type="hidden" value="" name="BillMemName3">
				</td>
				<td bgcolor="#FFFFCC"><div align="right">舉發人代碼4</div></td>
		  		<td>
					<input type="text" size="10" name="BillMem4" onkeyup="getBillMemID4();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=4","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer17" style="position:absolute ; width:130px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID4">
					<input type="hidden" value="" name="BillMemName4">
				</td>
			</tr>
			</tr>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>舉發單位</div></td>
				<td>
					<input type="text" size="10" name="BillUnitID" onKeyUp="getUnit();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<div id="Layer6" style="position:absolute ; width:190px; height:30px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>填單日期</div></td>
				<td>
				<input type="text" size="10" value="<%=ginitdt(date)%>" maxlength="6" name="BillFillDate" onkeyup="getDealLineDate()">
				</td>
				
			</tr>


			<tr>
				<td bgcolor="#FFFFCC"><div align="right">專案代碼</div></td>
				<td>
					<input type="text" size="10" value="" name="ProjectID">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onClick='window.open("Query_Project.asp","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
				</td>
				
				<td bgcolor="#FFFFCC" align="right">輔助車種</td>
				<td colspan="2">
                 <input type="text" maxlength="1" size="3" value="" name="CarAddID" onKeyUp="getAddID();">
				<!-- <div id="Layer110" style="position:absolute; width:300px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #11FFFF; border: 1px none #000000; color: #FF0000;"> -->
				<font color="#ff000" size="2">1大貨/ 2大客/ 3砂石/ 4土方/ 5動力/ 6貨櫃/ 7大型重機</font>
				<!-- </div> -->
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" width="13%"><div align='right'>採証工具</div></td>
				<td>
					<input maxlength="1" size="3" value="<%=request("UseTool")%>" name="UseTool"  onkeyup="getFixID();" type='text'> 
			         <div id="Layer11" style="position:absolute; width:285px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold; visibility: hidden;"> <font color="#0000FF">&nbsp;&nbsp;<font color="#000000">固定桿編號：</font></font>
                    <input type='text' size='6' name='FixID' value='<%=request("FixID")%>' onkeyup="setFixEquip();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_FixEquip.asp","WebPageFix","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
				  </div>
				  <font color="#ff000" size="2"> 1固定桿/ 2雷達三腳架/ 3相機/ 8逕舉手開單</font>
				  </td>
				<td bgcolor="#FFFFCC"><div align="right">備註</div></td>
				<td colspan="2">
					<input type="text" size="30" value="" name="Note">
					
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" align="center" colspan="5">
					<input type="button" value="儲 存 F2" onclick="getCheckCarNoExist();" <%
				'1:查詢 ,2:新增 ,3:修改 ,4:刪除
				if CheckPermission(223,2)=false then
					response.write "disabled"
				end if
					%> class="btn1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit1343" onClick="location='BillKeyIn_Car_Report2.asp'" value="清 除 F4" class="btn1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit5322" onClick="funcOpenBillQry()" value="查 詢 F6" class="btn1">
					<input type="hidden" name="kinds" value="">
                    <span class="style1">
                    <span class="style3"><img src="/image/space.gif" width="29" height="8"></span>
                    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" class="btn1">
                </span>
				<!-- 告發類別 -->
				<input type="hidden" size="3" maxlength="1" value="2" name="BillType" readonly>
				<!-- 應到案處所 -->
				<input type="hidden" size="4" value="" name="MemberStation" onkeyup="getStation();">
				<div id="Layer5" style="position:absolute ; width:241px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
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
function getCheckCarNoExist(){
	myForm.CarNo.value=myForm.CarNo.value.toUpperCase();
	var CarNum=myForm.CarNo.value;
	runServerScript("getCheckCarNoExist.asp?CarID="+CarNum);
}
function setCheckCarNoExist(CarNoFlag){
	if (CarNoFlag=="0"){
		alert("此車牌號碼尚未建檔！");
	}else{
		InsertBillVase();
	}
}
function InsertBillVase(){
	var error=0;
	var errorString="";
	var TodayDate=<%=ginitdt(date)%>;
	if (myForm.BillType.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入告發類別。";
	}
	if (myForm.CarNo.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規車號。";
	}else if (chkCarNoFormat(myForm.CarNo.value)==0){
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
	if (myForm.ReportChk.checked==false){
		if (myForm.IllegalAddressID.value==""){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入違規地點代碼。";
		}
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
	if (myForm.Rule2.value!=""){
		if (TDLawErrorLog2==1){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二輸入錯誤。";
		}else if (myForm.Rule2.value.substr(0,2)>68){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二輸入錯誤。";
		}
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
		if(myForm.BillType.value=="1"){
			error=error+1;
			errorString=errorString+"\n"+error+"：請輸入應到案日期。";
		}
	}else if (!dateCheck( myForm.DealLineDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if (!ChkIllegalDate(myForm.BillFillDate.value) && myForm.CaseInByMem.checked==false){
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
//是否為特殊用車
function getVIPCar(){
	strSpecUser=<%=trim(Session("SpecUser"))%>;
	myForm.CarNo.value=myForm.CarNo.value.replace(" ", "");
	if (myForm.CarNo.value.length >= 1){
		var CarNum=myForm.CarNo.value;
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType==0){
			alert("車牌格式錯誤");
			myForm.CarNo.select();
		}else{
			if (strSpecUser=="1"){
				runServerScript("getVIPCar.asp?CarID="+CarNum);
			}
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
		if (myForm.ReportChk.checked==false){
			myForm.BillMem1.select();
		}else{
			myForm.DealLineDate.select();
		}
	}
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
	if (myForm.BillUnitID.value.length > 1){
		var BillUnitNum=myForm.BillUnitID.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum);
	}else{
		Layer6.innerHTML=" ";
		TDUnitErrorLog=1;
	}
}


//逕舉不一定要輸入固定桿編號. 除了是下方選擇使用固定桿
function getFixID(){
	if (myForm.UseTool.value.length == "1"){
		if (myForm.UseTool.value != "0" && myForm.UseTool.value != "1" && myForm.UseTool.value != "2" && myForm.UseTool.value != "3" && myForm.UseTool.value != "8"){
			alert("採証工具填寫錯誤!");
			myForm.UseTool.select();
		}else if (myForm.UseTool.value == "1"){
			//Layer11.style.visibility = "visible"; 
		}else{
			//Layer11.style.visibility = "hidden"; 
		}
	}
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
//逕舉由填單日期帶入應到案日期+29
function getDealLineDate(){
	if (myForm.ReportChk.checked!=true){
		myForm.BillFillDate.value=myForm.BillFillDate.value.replace(/[^\d]/g,'');
		BFillDateTemp=myForm.BillFillDate.value;
		if (BFillDateTemp.length >= 6 && myForm.BillType.value=="2"){
			Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
			Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
			Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
			var BFillDate=new Date(Byear,Bmonth-1,Bday)
			var DLineDate=new Date()
			DLineDate=DateAdd("d",29,BFillDate);
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
	myForm.RuleSpeed.value=myForm.RuleSpeed.value.replace(/[^\d]/g,'');
	CallChkLaw1();
	CallChkLaw2();
	if (myForm.RuleSpeed.value > 100){
		alert("限速、限重超過100，請確認是否正確!");
	}
}
function IllegalSpeedforLaw(){
	myForm.IllegalSpeed.value=myForm.IllegalSpeed.value.replace(/[^\d]/g,'');
	CallChkLaw1();
	CallChkLaw2();
	if (myForm.IllegalSpeed.value > 100){
		alert("車速超過100，請確認是否正確!");
	}
}
function IllegalSpeedforLaw(){
	myForm.IllegalSpeed.value=myForm.IllegalSpeed.value.replace(/[^\d]/g,'');
	CallChkLaw1();
	CallChkLaw2();
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
//勾選後才可以輸入單號
function funcReportChk(){
	if (myForm.ReportChk.checked==true){
		myForm.Billno1.disabled=false;
		myForm.UseTool.value="8";
		//LayerDLDate.style.visibility = "visible"; 
		//LayerMStation.style.visibility = "visible";
		//myForm.MemberStation.disabled=false;
		DLDate1.innerHTML="應到案日期";
		DLDate2.innerHTML="<input type='text' size='6' value='' maxlength='6' name='DealLineDate' onBlur='DealLineDateReplace()' style=ime-mode:disabled>";

	}else{
		myForm.Billno1.value="";
		myForm.Billno1.disabled=true;
		myForm.UseTool.value="";
		//LayerDLDate.style.visibility = "hidden"; 
		//LayerMStation.style.visibility = "hidden"; 
		//myForm.MemberStation.disabled=true;
		myForm.MemberStation.Type="Text";
		DLDate1.innerHTML="";
		DLDate2.innerHTML="<input type='hidden' size='6' value='' maxlength='6' name='DealLineDate' onBlur='DealLineDateReplace()' style=ime-mode:disabled>";

	}
}
function DealLineDateReplace(){
	myForm.DealLineDate.value=myForm.DealLineDate.value.replace(/[^\d]/g,'');

}
function CheckBillNoExist(){
	myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
	if (myForm.Billno1.value!=""){
		if (myForm.Billno1.value.length != 9 ){
			alert("單號不足九碼！");
			myForm.Billno1.select();
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
		location='BillKeyIn_Car_Report2.asp'
	}else if (event.keyCode==117){ //F6查詢
		event.keyCode=0;   
		funcOpenBillQry();
	}else if (event.keyCode==119){ //F8關閉
		event.keyCode=0;   
		window.close();
	}
}
function funcOpenBillQry(){
	Sys_RMemberID=<%=session("User_ID")%>;
	Sys_BillTypeID=2;
	window.open("../Query/BillBaseQry.asp?QryType=1&Sys_RecordMemberID="+Sys_RMemberID+"&Sys_BTypeID="+Sys_BillTypeID,"WebPage4_Update","left=0,top=0,location=0,width=1000,height=660,resizable=yes,scrollbars=yes");
}
funcReportChk();
myForm.CarNo.focus();
getDealLineDate();

</script>
</html>
