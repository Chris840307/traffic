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
<!--#include virtual="/traffic/Common/css.txt"-->
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'on error resume next
'檢查是否可進入本系統
'AuthorityCheck(223)
'==========POST=========
'單號
if trim(request("billno"))="" then
	theBillno=""
else
	theBillno=trim(request("billno"))
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
gCh_Name = session("CH_Name")
gUnit_ID = Session("Unit_ID")
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
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

'===========================================
'是否要放大鏡功能(Y/N)
if sys_City="XX縣" then
	isBig="N" 
else
	isBig="Y" 
end if
'要抓本機照片或是server上的照片(0:本機 1:Server)
if sys_City="高雄市" then
	HowCatchPicture="0" 
else
	HowCatchPicture="1" 
end if
'本機路逕
if trim(Session("ImageSaveLocation"))="" then
	UserPicturePath="C:/Image/"
else
	UserPicturePath=trim(Session("ImageSaveLocation"))
end if
PicturePath="file:///"&UserPicturePath
'============================================
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
	'建檔影像
		theImageFileName=trim(request("gImageFileNameA"))
		theImagePathName=trim(request("gImagePathNameA"))

	'BillBase
	strUpdate="Update BillBase set" &_
		" BillNo='"&UCase(trim(request("Billno1")))&"'" &_
		",CarNo='"&UCase(trim(request("CarNo")))&"',CarSimpleID="&trim(request("CarSimpleID")) &_
		",CarAddID="&theCarAddID&",IllegalDate="&theIllegalDate&_
		",IllegalAddressID='"&trim(request("IllegalAddressID"))&"'" &_
		",IllegalAddress='"&trim(request("IllegalAddress"))&"'" &_
		",Rule1='"&trim(request("Rule1"))&"',IllegalSpeed="&theIllegalSpeed &_
		",RuleSpeed="&theRuleSpeed&",Rule2='"&trim(request("Rule2"))&"'" &_
		",ForFeit1="&trim(request("ForFeit1"))&",ForFeit2="&theForFeit2 &_
		",Insurance="&theInsurance&",UseTool="&theUseTool &_
		",ProjectID='"&trim(request("ProjectID"))&"'" &_
		",MemberStation='"&trim(request("MemberStation"))&"',BillUnitID='"&trim(request("BillUnitID"))&"'" &_
		",BillMemID1='"&trim(request("BillMemID1"))&"',BillMem1='"&trim(request("BillMemName1"))&"'" &_
		",BillMemID2='"&trim(request("BillMemID2"))&"',BillMem2='"&trim(request("BillMemName2"))&"'" &_
		",BillMemID3='"&trim(request("BillMemID3"))&"',BillMem3='"&trim(request("BillMemName3"))&"'" &_
		",BillMemID4='"&trim(request("BillMemID4"))&"',BillMem4='"&trim(request("BillMemName4"))&"'" &_
		",BillFillerMemberID='"&trim(request("BillMemID1"))&"',BillFiller='"&trim(request("BillMemName1"))&"'" &_
		",BillFillDate="&theBillFillDate&",DealLineDate="&theDealLineDate&_
		",Note='"&trim(request("Note"))&"',EquipmentID='1'" &_
		",BillStatus='0',RECORDSTATEID=0" &_
		" where SN="&trim(request("BillSN"))
	
		conn.execute strUpdate
				'response.write strUpdate
				'response.end
				'theDriverBirth , theBillFillDate 

%>
<script language="JavaScript">
	alert("修改完成");
</script>
<%
end if

if trim(request("kinds"))="VerifyResultNull" then
	'該筆紀錄的打驗資料表的 BILLSTATUS 更新為 6
	'strUpdDelTemp="Update BillBaseTmp set billstatus='6',RecordStateID=-1,DelMemberID="&Session("User_ID")&" where CarNo='"&trim(request("CarNo"))&"'"
	'conn.execute strUpdDelTemp

	'更新該筆紀錄的 BILLSTATUS 更新為 6
	strDelBill="Update BillBase set billstatus='6',RecordStateID=-1,DelMemberID='"&Session("User_ID")&"'" &_
		" where SN="&trim(request("BillSN"))
	conn.execute strDelBill

	ConnExecute "舉發單刪除 單號:"&trim(request("Billno1"))&" 車號:"&trim(request("CarNo"))&" 原因:"&DeleteReason&","&trim(NoteTmp)&","&CaseInStatus,352


	'============================

	
end if

	if trim(request("kinds"))="DB_insert" then
		sqlPage=" and RecordDate = TO_DATE('"&trim(Session("BillTime_Image"))&"','YYYY/MM/DD/HH24/MI/SS') order by RecordDate"
	elseif trim(request("kinds"))="VerifyResultNull" then
		sqlPage=" and RecordDate > TO_DATE('"&trim(Session("BillTime_Image"))&"','YYYY/MM/DD/HH24/MI/SS') order by RecordDate"
	elseif trim(request("PageType"))="Back" then
		sqlPage=" and RecordDate < TO_DATE('"&trim(Session("BillTime_Image"))&"','YYYY/MM/DD/HH24/MI/SS') order by RecordDate desc"
		Session("BillOrder_Image")=Session("BillOrder_Image")-1
	elseif trim(request("PageType"))="Next" then
		sqlPage=" and RecordDate > TO_DATE('"&trim(Session("BillTime_Image"))&"','YYYY/MM/DD/HH24/MI/SS') order by RecordDate"
		Session("BillOrder_Image")=Session("BillOrder_Image")+1
	elseif trim(request("PageType"))="First" Or trim(request("PageType"))="" then
		sqlPage=" order by RecordDate"
		Session("BillOrder_Image")=1
	elseif trim(request("PageType"))="Last" then
		sqlPage=" order by RecordDate Desc"
		Session("BillOrder_Image")=Session("BillCnt_Image")
	end If
	
	strSql="select * from (select a.* from BillBase a where Rule1='5620001' and BillStatus='0'" &_
		" and RecordStateid=0 " &_
		" and ImageFileName is not null " & sqlPage&") where rownum<=1"
	set rs1=conn.execute(strSql)

	if rs1.eof then
		if trim(request("PageType"))="Next" then
			Response.Redirect "BillKeyIn_Image_Fix_Back_TD.asp?SessionFlag=1"
		elseif trim(request("PageType"))="Back" then
			Response.Redirect "BillKeyIn_Image_Fix_Back_TD.asp?SessionFlag=1"
		elseif trim(request("PageType"))="First" then
			Response.Redirect "BillKeyIn_Image_Fix_Back_TD.asp?SessionFlag=1"
		elseif trim(request("PageType"))="Last" then
			Response.Redirect "BillKeyIn_Image_Fix_Back_TD.asp?SessionFlag=1"
		end if
	end if


'response.write strSql
	'總共幾筆
	Session.Contents.Remove("BillCnt_Image")
	strSqlCnt="select count(*) as cnt from BillBase where Rule1='5620001' and BillStatus='0' and RecordStateid=0 " &_
		" and ImageFileName is not null"
	set rsCnt1=conn.execute(strSqlCnt) 
		Session("BillCnt_Image")=trim(rsCnt1("cnt"))
	rsCnt1.close
	set rsCnt1=Nothing
	
	If Not rs1.Bof Then rs1.MoveFirst 
	If not rs1.eof then

	Session.Contents.Remove("BillTime_Image")
	Session("BillTime_Image")=year(rs1("RecordDate"))&"/"&month(rs1("RecordDate"))&"/"&day(rs1("RecordDate"))&" "&hour(rs1("RecordDate"))&":"&minute(rs1("RecordDate"))&":"&second(rs1("RecordDate"))
	End If 
%>
<title>數位舉發建檔</title>
<style type="text/css">
<!--
.style2 {font-size: 12px}
.style3 {
font-size: 12px ;
color: #FF0000}
.style4 {
font-size: 12px ;
}
.style5 {
color: #0000FF;
font-size: 13px ;
}
.style6 {
color: #FF0000;
font-size: 13px ;
}
.btn2 {font-size: 13px}
.Text1{
font-weight:bold;
}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">

<form name="myForm" method="post">  
<table width='1000' border='1' align="left" cellpadding="0">
	<tr>
		<td width="24%" height="280" align="center" valign="bottom">
	<%if not rs1.eof then%>
		<%
		theImageFileNameA=""
		theImageFileNameB=""
		theImageFileNameC=""
		theIISImagePath="\traffic\StopCarPicture\"
		if trim(rs1("ImageFileName"))<>"" and not isnull(rs1("ImageFileName")) then
			theImageFileNameA=trim(rs1("ImageFileName"))
		end if


		bPicWebPath = ""
		if trim(theImageFileNameA)<>"" then
			bPicWebPath=theIISImagePath & theImageFileNameA
		end If
		
		'response.write bPicWebPath
		%><br><br><br>&nbsp;&nbsp;
		<div id="div1" style="position:absolute; overflow:hidden;left:20px;top:20px ; width:212px; height:90px <%
		'放大鏡功能
		if isBig="N" then
		%> ;visibility: hidden;<%
		end if
		%>">
			<img id=BigImg style='position:relative' src="<%=bPicWebPath%>">
		</div>
	<%end if%>
	<%
	if sys_City="高雄市" then
		if not rs1.eof then%>
		<!-- 影像小圖C -->
		<%
		sPicWebPath2=""
		if trim(theImageFileNameC)<>"" then
			sPicWebPath2=theIISImagePath&theImageFileNameC
		end if
		%>
		<%if sPicWebPath2<>"" then%>
		<img src="<%=sPicWebPath2%>" border=1 width="210" id="SmallImg2" ondblclick="ChangeImg2()">
		<%end if%>
	<%
		end if
	end if
	%>
		</td>
		<td rowspan="2">
		<!-- 影像大圖 -->
	<%if not rs1.eof then%>
		<%if bPicWebPath<>"" then%>
		<img src="<%=bPicWebPath%>" border=1 height="490" <%
		'放大鏡功能
		if isBig="Y"  then
		%>onmousemove="show(this, '<%=bPicWebPath%>')" onmousedown="show(this, '<%=bPicWebPath%>')"<%
		end if
		%> id="imgSource" src="<%=bPicWebPath%>" >
		<%end if%>
	<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
	<%end if%>
		</td>
	</tr>
	<tr>
		<td height="210" align="center">
	<%if not rs1.eof then%>
		<!-- 影像小圖B -->
		<%
		sPicWebPath=""
		if trim(theImageFileNameB)<>"" then
			sPicWebPath=theIISImagePath&theImageFileNameB
		end if
		%>
		<%if sPicWebPath<>"" then%>
		<img src="<%=sPicWebPath%>" border=1 width="210" id="SmallImg" ondblclick="ChangeImg()">
		<%end if%>
		<br>
		<input type="button" onClick="OpenPic('<%=replace(bPicWebPath,"\","/")%>')" value="大圖一" class="style4">
		<%if trim(theImageFileNameB)<>"" then%>
			<input type="button" onClick="OpenPic('<%=replace(sPicWebPath,"\","/")%>')" value="大圖二" class="style4">
		<%end if%>
		<%if trim(theImageFileNameC)<>"" then%>
			<input type="button" onClick="OpenPic('<%=replace(sPicWebPath2,"\","/")%>')" value="大圖三" class="style4">
		<%end if%>

			
	<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
	<%end if%>
		</td>
	</tr>

	<tr>
		<td height="100" colspan="2" valign="top">
		<%if not rs1.eof then%>
		<table width='100%' border='1' align="left" cellpadding="0">
			<tr>
				<td bgcolor="#FFFFCC"><div align="right"> <span class="style3">＊</span>違規車號</div></td>
				<td>
				<input type="text" size="8" class="Text1" name="CarNo" onBlur="getVIPCar();" value="<%
				if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
					response.write trim(rs1("CarNo"))
				end if
				%>" style=ime-mode:disabled maxlength="8" onkeydown="funTextControl(this);">
				<span class="style6">
			    <div id="Layer7" style="position:absolute; width:170px; height:24px; z-index:0; color: #FF0000; font-weight: bold;"><%
			if trim(Session("SpecUser"))="1" then
				strSC="select count(*) as cnt from SpecCar where CarNo='"&trim(rs1("CarNo"))&"' and RecordStateID<>-1"
				set rsSC=conn.execute(strSC)
				if not rsSC.eof then
					if trim(rsSC("cnt"))<>"0" then
						response.write "＊業管車輛"
					end if
				end if
				rsSC.close
				set rsSC=nothing
			end if
				%></div>
				</span>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span>違規時間&nbsp;</div></td>
				<td>
				<!-- 違規日期 -->&nbsp;
				<input type="text" size="6" maxlength="7" class="Text1" name="IllegalDate" value="<%
				if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then 
					response.write gInitDT(rs1("IllegalDate"))
				end if
				%>" onBlur="getBillFillDate()" style=ime-mode:disabled onkeydown="funTextControl(this);">
				
				<!-- 違規時間 -->
				<input type="text" size="3" maxlength="4" class="Text1" name="IllegalTime" value="<%
				if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then 
					response.write Right("00"&hour(rs1("IllegalDate")),2)&Right("00"&minute(rs1("IllegalDate")),2)
				end if
				%>" onBlur="value=value.replace(/[^\d]/g,'')" style=ime-mode:disabled onkeydown="funTextControl(this);">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span>時速&nbsp;</div></td>
				<td colspan="3">
					限速<input type="text" size="4" maxlength="3" class="Text1" name="RuleSpeed" value="<%
					if trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed")) then
						response.write trim(rs1("RuleSpeed"))
					end if
					%>" onBlur="RuleSpeedforLaw()" style=ime-mode:disabled onkeyup="setIllegalRule();" onkeydown="funTextControl(this);">
					車速<input type="text" size="4" maxlength="3" class="Text1" name="IllegalSpeed" value="<%
					if trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed")) then
						response.write trim(rs1("IllegalSpeed"))
					end if
					%>" onkeyup="IllegalSpeedforLaw()" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<div id="Layer012" style="position:absolute; width:125px; height:27px; z-index:1; background-color: #FFCCCC; visibility: hidden;"><font color="#0000FF" size="2">&nbsp;1汽車 / 2拖車 / 3重機<br>&nbsp;/ 4輕機 / 6臨時車牌</font></div>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span>車種&nbsp;</div></td>
				<td >
				<input type="text" maxlength="1" size="3" value="<%
				if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
					response.write trim(rs1("CarSimpleID"))
				end if
					%>" name="CarSimpleID" onBlur="getRuleAll();" style=ime-mode:disabled onfocus="Layer012.style.visibility='visible';" onkeydown="funTextControl(this);">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" ><div align="right"><span class="style3">＊</span>違規法條一</div></td>
				<td colspan="3">
					<input type="text" maxlength="9" size="8" value="<%
					if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
						response.write trim(rs1("Rule1"))
					end if
					%>" name="Rule1" style=ime-mode:disabled onkeydown="funTextControl(this);" readonly>
					
					<span class="style5">
					<div id="Layer1" style="position:absolute ; width:260px; height:28px; z-index:0; border: 1px none #000000;"><%
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
						strRule1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
						set rsRule1=conn.execute(strRule1)
						if not rsRule1.eof then
							response.write trim(rsRule1("IllegalRule"))
							gLevel1=trim(rsRule1("Level1"))
						end if
						rsRule1.close
						set rsRule1=nothing
					end if
					%></div>
					</span>
					<input type="hidden" name="ForFeit1" value="<%
						response.write gLevel1
					%>">
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span>違規地點</div></td>
				<td colspan="5">
					<input type="text" size="4" value="<%
				if trim(rs1("IllegalAddressID"))<>"" and not isnull(rs1("IllegalAddressID")) then
					response.write trim(rs1("IllegalAddressID"))
				end if
				%>" name="IllegalAddressID" onblur="funGetSpeedRule()" onKeyUp="getillStreet();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="15" height="15" onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>

					<input type="text" size="28" value="<%
					if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
						response.write trim(rs1("IllegalAddress"))
					end if
					%>" name="IllegalAddress" style=ime-mode:active onblur="funGetSpeedRule()" onkeyup="AutoGetIllStreet();" onkeydown="funTextControl(this);">
					<input type="checkbox" name="chkHighRoad" value="1" <%
					if Left(trim(rs1("Rule1")),2)="33" then
						response.write "checked"
					elseif trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
						if Left(trim(rs1("Rule2")),2)="33" then
							response.write "checked"
						end if
					end if
					
					%> onclick="setIllegalRule()"><span class="style1">快速道路</span>
				</td>
		    </tr>
			<tr>
				<td bgcolor="#FFFFCC" ><div align="right">違規法條二</div></td>
				<td colspan="3" >
					<input type="text" maxlength="9" size="8" value="<%
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					response.write trim(rs1("Rule2"))
				end if
				%>" name="Rule2" onKeyUp="getRuleData2();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="15" height="15" onclick='window.open("Query_Law.asp?LawOrder=2&RuleVer=<%=trim(rs1("RuleVer"))%>","WebPage_Law","left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer2" style="position:absolute ; width:260px; height:28px; z-index:0; border: 1px none #000000;"><%
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
						gLevel2=trim(rsR1("Level1"))
					end if
					rsR1.close
					set rsR1=nothing
				end if
				%></div>
					</span>
					<input type="hidden" name="ForFeit2" value="<%
				if trim(rs1("ForFeit2"))<>"" and not isnull(rs1("ForFeit2")) then
					response.write trim(rs1("ForFeit2"))
				else
					if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
						response.write gLevel2
					end if
				end if
				%>">

				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style3">＊</span>舉發人</div></td>
				<td colspan="5">
					1.
					<input type="text" size="4" name="BillMem1" value="<%
				if trim(rs1("BillMemID1"))<>"" and not isnull(rs1("BillMemID1")) then
					strMem1="select LoginID,ChName from MemberData where MemberID="&trim(rs1("BillMemID1"))
					set rsMem1=conn.execute(strMem1)
					if not rsMem1.eof then
						response.write trim(rsMem1("LoginID"))
						MemChName=trim(rsMem1("ChName"))
					end if
					rsMem1.close
					set rsMem1=nothing
				end if
				%>" onKeyUp="getBillMemID1();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="15" height="15" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer12" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=MemChName%></div>
					</span>
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
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
					2.
					<input type="text" size="4" name="BillMem2" value="<%
				if trim(rs1("BillMemID2"))<>"" and not isnull(rs1("BillMemID2")) then
					strMem1="select LoginID,ChName from MemberData where MemberID="&trim(rs1("BillMemID2"))
					set rsMem1=conn.execute(strMem1)
					if not rsMem1.eof then
						response.write trim(rsMem1("LoginID"))
						MemChName2=trim(rsMem1("ChName"))
					end if
					rsMem1.close
					set rsMem1=nothing
				end if
				%>" onKeyUp="getBillMemID2();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="15" height="15" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer13" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=MemChName2%></div>
					</span>
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
					&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
					3.
					<input type="text" size="4" name="BillMem3" value="<%
				if trim(rs1("BillMemID3"))<>"" and not isnull(rs1("BillMemID3")) then
					strMem1="select LoginID,ChName from MemberData where MemberID="&trim(rs1("BillMemID3"))
					set rsMem1=conn.execute(strMem1)
					if not rsMem1.eof then
						response.write trim(rsMem1("LoginID"))
						MemChName3=trim(rsMem1("ChName"))
					end if
					rsMem1.close
					set rsMem1=nothing
				end if
				%>" onKeyUp="getBillMemID3();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="15" height="15" onclick='window.open("Query_MemID.asp?MemOrder=3","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer14" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=MemChName3%></div>
					</span>
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
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" width="10%"><div align="right">舉發人4</div></td>
		  		<td width="12%">
					<input type="text" size="4" name="BillMem4" value="<%
				if trim(rs1("BillMemID4"))<>"" and not isnull(rs1("BillMemID4")) then
					strMem1="select LoginID,ChName from MemberData where MemberID="&trim(rs1("BillMemID4"))
					set rsMem1=conn.execute(strMem1)
					if not rsMem1.eof then
						response.write trim(rsMem1("LoginID"))
						MemChName4=trim(rsMem1("ChName"))
					end if
					rsMem1.close
					set rsMem1=nothing
				end if
				%>" onKeyUp="getBillMemID4();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="15" height="15" onclick='window.open("Query_MemID.asp?MemOrder=4","WebPage1","left=0,top=0,location=0,width=800,height=555,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer17" style="position:absolute ; width:130px; height:30; z-index:0;  border: 1px none #000000;"><%=MemChName4%></div>
					</span>
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
				<td bgcolor="#FFFFCC" width="10%"><div align="right"><span class="style3">＊</span>舉發單位</div></td>
				<td width="16%">
					<input type="text" size="4" name="BillUnitID" value="<%
				if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
					response.write trim(rs1("BillUnitID"))
				end if
				%>" onKeyUp="getUnit();" style=ime-mode:disabled onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="15" height="15" onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer6" style="position:absolute ; width:100px; height:30px; z-index:0;  border: 1px none #000000;"><%
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
				<td bgcolor="#FFFFCC" width="10%"><div id="Layer110" style="position:absolute; width:265px; height:27px; z-index:1; background-color: #FFCCCC; visibility: hidden;">
				<font color="#0000FF" size="2">1大貨/2大客/3砂石/4土方/5動力/6貨櫃/7大型重機/8拖吊/9(550cc)重機 /10計程車/ 11危險物品 </font>
				</div>
				<div align="right"><span class="style3">＊</span>填單日期</div></td>
				<td width="9%">
				<input type="text" size="8" value="<%
				if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
					response.write ginitdt(trim(rs1("BillFillDate")))
				end if
				%>" maxlength="7" name="BillFillDate" onBlur="getDealLineDate()" style=ime-mode:disabled onkeydown="funTextControl(this);">

				</td>
				<td bgcolor="#FFFFCC" align="right" width="8%">輔助車種</td>
				<td width="6%">
                 <input type="text" maxlength="2" size="4" value="<%
				if trim(rs1("CarAddID"))<>"" and not isnull(rs1("CarAddID")) then
					response.write trim(rs1("CarAddID"))
				end if
				%>" name="CarAddID" onBlur="getAddID();" style=ime-mode:disabled  onfocus="Layer110.style.visibility='visible';" onkeydown="funTextControl(this);">
                
				</td>
				<td bgcolor="#FFFFCC" width="8%"><div align="right">專案代碼</div></td>
				<td width="9%">
					<input type="text" size="5" value="<%
				if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then
					response.write trim(rs1("ProjectID"))
				end if
				%>" name="ProjectID" style=ime-mode:disabled onkeyup="ProjectF5()" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="15" height="15" onClick='window.open("Query_Project.asp","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer001" style="position:absolute ; width:180px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000; visibility: hidden;"></div>

					
				<!-- 備註 -->
					<input type="hidden" size="29" value="<%
					if trim(rs1("Note"))<>"" and not isnull(rs1("Note")) then
						response.write trim(rs1("Note"))
					end if
					%>" name="Note" style=ime-mode:active>
				<!-- 採証工具 -->
					<input maxlength="1" size="4" value="<%
				if trim(rs1("UseTool"))<>"" and not isnull(rs1("UseTool")) then
					response.write trim(rs1("UseTool"))
				end if
				%>" name="UseTool"  onkeyup="getFixID();" type='hidden' style=ime-mode:disabled> 
			        <div id="Layer11" style="position:absolute; width:275px; height:24px; z-index:0; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; color: #FF0000; font-weight: bold; visibility: hidden;"> <font color="#0000FF">&nbsp;&nbsp;<font color="#000000">固定桿編號：</font></font>
                    <input type='text' size='6' name='FixID' value='<%
				if trim(rs1("EQUIPMENTID"))<>"" and not isnull(rs1("EQUIPMENTID")) then
					response.write trim(rs1("EQUIPMENTID"))
				end if
				%>' onBlur="setFixEquip();" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_FixEquip.asp","WebPageFix","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					</div>
					<!-- <font color="#ff000" size="2"> 1固定桿/ 2雷達三腳架/ 3相機</font> -->
				</td>
			</tr>
		</table>
		<%else%>
		<font color="#ff000"><strong>無未建檔案件..</strong></font>
		<%end if%>
		</td>
	</tr>
	<tr bgcolor="#FFCC33">
		<td height="28" colspan="2" align="center">
					<input type="button" value="儲 存 F2" onclick="InsertBillVase();" >

					<input type="hidden" name="kinds" value="">
                    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8">
                    <input type="button" name="Submit2932" onClick="funVerifyResult();" value="無 效 F9">
					
					<img src="/image/space.gif" width="20" height="8">
					<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_Fix_Back_TD.asp?PageType=First'" value="<< 第一筆 Home" style="font-size: 9pt; width: 100px; height: 27px" <%
				If Session("BillOrder_Image")=0 Or Session("BillOrder_Image")=1 Then
					response.write "disabled"
				End if
					%>>
					<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_Fix_Back_TD.asp?PageType=Back'" value="< 上一筆 PgUp" style="font-size: 9pt; width: 100px; height: 27px" <%
				If Session("BillOrder_Image")=0 Or Session("BillOrder_Image")=1 Then
					response.write "disabled"
				End if
					%>>
					
					<!-- <img src="/image/space.gif" width="29" height="8"> -->
					<%
						response.write Session("BillCnt_Image")&" / "&Session("BillOrder_Image")
						
					%>
					
					<input type="button" name="SubmitNext" onClick="location='BillKeyIn_Image_Fix_Back_TD.asp?PageType=Next'" value="下一筆 PgDn >" style="font-size: 9pt; width: 100px; height: 27px" <%
				If Trim(Session("BillOrder_Image"))=trim(Session("BillCnt_Image")) Then
					response.write "disabled"
				End if
					%>>
					<input type="button" name="SubmitBack" onClick="location='BillKeyIn_Image_Fix_Back_TD.asp?PageType=Last'" value="最後一筆 End >>" style="font-size: 9pt; width: 100px; height: 27px" <%
				If Trim(Session("BillOrder_Image"))=trim(Session("BillCnt_Image")) Then
					response.write "disabled"
				End if
					%>>
				<!-- 逕舉類別 -->
				<input type="hidden" size="3" maxlength="1" value="2" name="BillType" readonly>
				<!-- 應到案日期 -->
				<input type="hidden" size="12" maxlength="7" value="<%
				if Not rs1.eof then

					if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
						response.write ginitdt(trim(rs1("DealLineDate")))
					end If
				End if
					%>" name="DealLineDate">
				<!-- 應到案處所 -->
				<input type="hidden" size="10" value="<%
				if Not rs1.eof then
					if trim(rs1("MemberStation"))<>"" and not isnull(rs1("MemberStation")) then
						response.write trim(rs1("MemberStation"))
					end If
				End if
					%>" name="MemberStation">
				<input type="hidden" value="<%
				if Not rs1.eof Then
					response.write trim(rs1("SN"))
				End if
				%>" name="BillSN">
				<!-- <input type="button" value="？" name="StationSelect" onclick='window.open("Query_Station.asp","WebPage1","left=0,top=0,location=0,width=660,height=375,resizable=yes,scrollbars=yes")'> -->
				<div id="Layer5" style="position:absolute ; width:221px; height:24px; z-index:0; background-color: #CCFFFF; layer-background-color: #CCFFFF; border: 1px none #000000; visibility :hidden;"></div>

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
var TDVipCarErrorLog=0;
var SpeedError=0;
var TodayDate=<%=ginitdt(date)%>;

MoveTextVar("CarNo,IllegalDate,IllegalTime,RuleSpeed,IllegalSpeed,CarSimpleID||Rule1,IllegalAddressID,IllegalAddress||Rule2,BillMem1,BillMem2,BillMem3||BillMem4,BillUnitID,BillFillDate,CarAddID,ProjectID");
//新增告發單
function InsertBillVase(){
	var error=0;
	var errorString="";
	var TodayDate=<%=ginitdt(date)%>;
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
	}//else if(myForm.CarNo.value != "" && chkCarNoFormat(myForm.CarNo.value)!= 0) {
	//	if (chkCarNoFormat(myForm.CarNo.value) != myForm.CarSimpleID.value){
	//		error=error+1;
	//		errorString=errorString+"\n"+error+"：車號格式與簡式車種不符。";
	//	}
	//}
	if (myForm.IllegalDate.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規日期。";
	}else if(!dateCheck( myForm.IllegalDate.value )){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
	}else if( myForm.IllegalDate.value.substr(0,1)=="9" && myForm.IllegalDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
	}else if( myForm.IllegalDate.value.substr(0,1)=="1" && myForm.IllegalDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
	}else if (!ChkIllegalDate(myForm.IllegalDate.value)){
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
	}else if( myForm.BillFillDate.value.substr(0,1)=="9" && myForm.BillFillDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
	}else if( myForm.BillFillDate.value.substr(0,1)=="1" && myForm.BillFillDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
	}else if(eval(TodayDate) < eval(myForm.BillFillDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：填單日期不得比今天晚。";
	}else if (!ChkIllegalDate(myForm.BillFillDate.value)){
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
	}else if( myForm.DealLineDate.value.substr(0,1)=="9" && myForm.DealLineDate.value.length==7 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if( myForm.DealLineDate.value.substr(0,1)=="1" && myForm.DealLineDate.value.length==6 ){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
	}else if (!ChkIllegalDate(myForm.DealLineDate.value)){
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
	if (myForm.BillMem1.value==""){
		//固定桿不需要輸入舉發人
		//if (myForm.UseTool.value!="1"){
		    error=error+1;
			errorString=errorString+"\n"+error+"：請輸入舉發人1 代碼。";
		//}
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
	if (eval(myForm.BillFillDate.value) < eval(myForm.IllegalDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比填單日晚。";
	}else if(eval(TodayDate) < eval(myForm.IllegalDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比今天晚。";
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
		if ((myForm.Rule1.value.substr(0,3))!="293" && (myForm.Rule2.value.substr(0,3))!="293")	{
			if(parseInt(myForm.RuleSpeed.value) < 30){
				error=error+1;
				errorString=errorString+"\n"+error+"：限速、限重小於 30Km/h。";
			}
		}		
		if(parseInt(myForm.RuleSpeed.value) > 300){
			error=error+1;
			errorString=errorString+"\n"+error+"：限速、限重大於 300Km/h。";
		}
		if(parseInt(myForm.IllegalSpeed.value) > 300){
			error=error+1;
			errorString=errorString+"\n"+error+"：實際車速、車重大於 300Km/h。";
		}
		if((parseInt(myForm.IllegalSpeed.value)-parseInt(myForm.RuleSpeed.value) ) > 150){
			error=error+1;
			errorString=errorString+"\n"+error+"：超速大於 150Km/h。";
		}
	}
<%if sys_City="高雄市" then%>
	if (SpeedError==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：超速 100~150Km/h ，請輸入密碼後才可建檔。";
	}
<%end if%>
<%if sys_City="台東縣" then %>
	if ((myForm.Rule1.value.substr(0,5))=="33101" || (myForm.Rule1.value.substr(0,2))=="40" || (myForm.Rule1.value.substr(0,5))=="43102"){
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule1.value,"3");
		if (IllegalRule == false){
			error=error+1;
			errorString=errorString+"\n"+error+"：超速法條與車速不符。";
		}
	}else if ((myForm.Rule2.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,5))=="43102"){
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule2.value,"3");
		if (IllegalRule == false){
			error=error+1;
			errorString=errorString+"\n"+error+"：超速法條與車速不符。";
		}
	}
<%else%>
	if ((myForm.Rule1.value.substr(0,5))=="33101" || (myForm.Rule1.value.substr(0,2))=="40" || (myForm.Rule1.value.substr(0,5))=="43102"){
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule1.value,"1");
		if (IllegalRule == false){
			error=error+1;
			errorString=errorString+"\n"+error+"：超速法條與車速不符。";
		}
	}else if ((myForm.Rule2.value.substr(0,5))=="33101" || (myForm.Rule2.value.substr(0,2))=="40" || (myForm.Rule2.value.substr(0,5))=="43102"){
		IllegalRule=chkSpeedRuleIsRight(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked,myForm.Rule2.value,"1");
		if (IllegalRule == false){
			error=error+1;
			errorString=errorString+"\n"+error+"：超速法條與車速不符。";
		}
	}
<%end if%>
	if ((myForm.Rule1.value.substr(0,2))=="36" && (myForm.CarSimpleID.value=="3" || myForm.CarSimpleID.value=="4")){
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
	if ((myForm.Rule1.value.substr(0,2))=="55"){
		error=error+1;
		errorString=errorString+"\n"+error+"：第55條不可逕行舉發。";
	}
<%end if%>
	if ((myForm.Rule1.value.substr(0,3))=="293" && (myForm.RuleSpeed.value=="" || myForm.IllegalSpeed.value=="")){
		error=error+1;
		errorString=errorString+"\n"+error+"：您選擇超重法條，但是未輸入限重或車重。";
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
	myForm.CarNo.value=myForm.CarNo.value.toUpperCase();
	myForm.CarNo.value=myForm.CarNo.value.replace(" ", "");
	if (myForm.CarNo.value.length >= 1){
		var CarNum=myForm.CarNo.value;
		CarType=chkCarNoFormat(myForm.CarNo.value);
		if (CarType==0){
			//alert("車牌格式錯誤");
			//myForm.CarNo.focus();
			runServerScript("getVIPCar.asp?CarID="+CarNum);
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
	Layer110.style.visibility='hidden';
	if (myForm.CarAddID.value.length>0){
		if (myForm.CarAddID.value != "1" && myForm.CarAddID.value != "2" && myForm.CarAddID.value != "3" && myForm.CarAddID.value != "4" && myForm.CarAddID.value != "5" && myForm.CarAddID.value != "6" && myForm.CarAddID.value != "7" && myForm.CarAddID.value != "8" && myForm.CarAddID.value != "9" && myForm.CarAddID.value != "10" && myForm.CarAddID.value != "11"){
			alert("輔助車種填寫錯誤!");
			//myForm.CarAddID.value = "";
			myForm.CarAddID.select();
		}
	}
}
//檢查簡式車種
function getRuleAll(){
	//myForm.CarSimpleID.value=myForm.CarSimpleID.value.replace(/[^\d]/g,'');
	Layer012.style.visibility='hidden';
	if (myForm.CarSimpleID.value.length>0){
		if (myForm.CarSimpleID.value != "1" && myForm.CarSimpleID.value != "2" && myForm.CarSimpleID.value != "3" && myForm.CarSimpleID.value != "4" && myForm.CarSimpleID.value != "6"){
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
		var VerNo=<%
		if Not rs1.eof Then
			response.write trim(rs1("RuleVer"))
		end if	
		%>;
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
	//AutoGetRuleID(1);
}
//違規事實2(ajax)
function getRuleData2(){
	if (myForm.Rule2.value.length > 6){
		var Rule2Num=myForm.Rule2.value;
		var CarSimpleID=myForm.CarSimpleID.value;
		var VerNo=<%
		if Not rs1.eof Then
			response.write trim(rs1("RuleVer"))
		end if	
		%>;
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

	//AutoGetRuleID(2);
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
	if (myForm.BillUnitID.value.length > 1){
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
		if (myForm.UseTool.value != "1" && myForm.UseTool.value != "2" && myForm.UseTool.value != "3"){
			alert("採証工具填寫錯誤!");
			myForm.UseTool.focus();
			myForm.UseTool.value = "";
		}else if (myForm.UseTool.value == "1"){
			//Layer11.style.visibility = "visible"; 
		}else{
			//Layer11.style.visibility = "hidden"; 
		}
	}
}
//違規地點代碼(ajax)
function getillStreet(){
	myForm.IllegalAddressID.value=myForm.IllegalAddressID.value.toUpperCase();
	if (event.keyCode==116){	
		event.keyCode=0;
		OstreetID=myForm.IllegalAddressID.value;
		window.open("Query_Street.asp?OstreetID="+OstreetID,"WebPage_Street_People2","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.IllegalAddressID.value.length > 2){
		var illAddrNum=myForm.IllegalAddressID.value;
		runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
	}
}
//舉發人一(ajax)
function getBillMemID1(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=1","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
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
//舉發人2(ajax)
function getBillMemID2(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=2","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
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
//舉發人3(ajax)
function getBillMemID3(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=3","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
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
//舉發人4(ajax)
function getBillMemID4(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=4","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
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

function getBillFillDate(){
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');
		if(eval(TodayDate) < eval(myForm.IllegalDate.value)){
		alert("違規日期不得大於今天!!");
		myForm.IllegalDate.select();
	}

//	if (myForm.IllegalDate.value.length >= 6 ){
//		myForm.BillFillDate.value=myForm.IllegalDate.value;
//		getDealLineDate();
//	}
}
//逕舉由填單日期帶入應到案日期
function getDealLineDate(){
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
	//myForm.RuleSpeed.value=myForm.RuleSpeed.value.replace(/[^\d]/g,'');
	CallChkLaw1();
	CallChkLaw2();
	if (myForm.RuleSpeed.value > 100){
		alert("限速、限重超過100，請確認是否正確!");
	}
}
function IllegalSpeedforLaw(){
	myForm.IllegalSpeed.value=myForm.IllegalSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
	CallChkLaw1();
	CallChkLaw2();
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
	if((myForm.Rule1.value.substr(0,2))!="29"){
		if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
			if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) >= <%
	if sys_City="雲林縣" then 
		response.write "100"
	'elseif sys_City="台東縣" then 
	'	response.write "61"
	else
		response.write "61"
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
				IntError=IntError+1;
				StrError=StrError+"\n"+IntError+"：超過最高限速60公里以上需另單舉發法條4340003(處車主)!!";
			}
		}
	}
	if (IntError!=0){
		alert(StrError+"\n\n請確認是否正確!");
	}
<%if sys_City="高雄市" then%>
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!=""){
		if ((myForm.IllegalSpeed.value - myForm.RuleSpeed.value) > 100 && (myForm.IllegalSpeed.value - myForm.RuleSpeed.value) < 150)
		{
			SpeedError=1;
			window.open("../BillKeyIn/ChkSpeedPW.asp","ChkSpeedPW","left=300,top=100,width=350,height=200,resizable=yes,scrollbars=no");
		}else{
			SpeedError=0;
		}
	}
<%end if%>
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
function CheckBillNoExist(){
	myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
	var BillNum=myForm.Billno1.value;
	if (myForm.Billno1.value.length >= 9){
		runServerScript("getCheckBillNoExist.asp?BillNo="+BillNum);
	}
}

//檢查單號是否有在GETBILLBASE內
function setCheckBillNoExist(GetBillFlag,BillBaseFlag,BillSN,BillType,MLoginID,MMemberID,MMemName,MUnitID,MUnitName)
{
	if (GetBillFlag==0){
		alert("此單號不存在於領單紀錄中！");
		document.myForm.Billno1.value="";
	}else{
		document.myForm.BillMem1.value=MLoginID;
		document.myForm.BillMemID1.value=MMemberID;
		document.myForm.BillMemName1.value=MMemName;
		Layer12.innerHTML=MMemName;
		TDMemErrorLog1=0;
		if (document.myForm.BillUnitID.value==""){
			document.myForm.BillUnitID.value=MUnitID;
			Layer6.innerHTML=MUnitName;
			TDUnitErrorLog=0;
		}
		if (BillBaseFlag==1){
			alert("此單號已建檔！");
			document.myForm.Billno1.value="";
		}else if (BillBaseFlag==0){
			alert('此單號已建檔！');
			document.myForm.Billno1.value="";
		}
	}
}

//逕舉建檔清冊
function funPrintCaseList_Report(){
	UrlStr="../Query/PrintCaseDataList_Report.asp?CallType=1";
	newWin(UrlStr,"CaseListWin2342",980,575,0,0,"yes","yes","yes","no");
}

//審核無效
function funVerifyResult(){
	if(confirm('確定要將此筆違規影像設為無效？')){
		myForm.kinds.value="VerifyResultNull";
		myForm.submit();
	}
}
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
		InsertBillVase();
	/*
	}else if (event.keyCode==115){ //F4清除
		event.keyCode=0;   
		event.returnValue=false;  
		location='BillKeyIn_Image.asp'
	*/
	}else if (event.keyCode==117){ //F6查詢
		event.keyCode=0;   
		event.returnValue=false;  
		funcOpenBillQry();
	}else if (event.keyCode==119){ //F8關閉
		event.keyCode=0;   
		event.returnValue=false;  
		window.close();
	}else if (event.keyCode==120){ //F9審核無效
		event.keyCode=0;   
		event.returnValue=false;  
		funVerifyResult();
	}else if (event.keyCode==121){ //F10查詢未建檔
		event.keyCode=0;   
		event.returnValue=false;  
		funPrintCaseList_Report();
<%If Session("BillOrder_Image")>1 Then%>
	}else if (event.keyCode==33){ //上一筆PageUp
		event.keyCode=0;   
		location='BillKeyIn_Image_Fix_Back_TD.asp?PageType=Back'
	}else if (event.keyCode==36){ //第一筆Home
		event.keyCode=0;   
		location='BillKeyIn_Image_Fix_Back_TD.asp?PageType=First'
<%end if 
If Trim(Session("BillOrder_Image"))<>trim(Session("BillCnt_Image")) Then
%>
	}else if (event.keyCode==34){ //下一筆PageDn
		event.keyCode=0;   
		event.returnValue=false; 
		location='BillKeyIn_Image_Fix_Back_TD.asp?PageType=Next'
	}else if (event.keyCode==35){ //最後一筆End
		event.keyCode=0;   
		event.returnValue=false; 
		location='BillKeyIn_Image_Fix_Back_TD.asp?PageType=Last'
<%end if%>
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
	//if (event.keyCode==116){	
	//	event.keyCode=0;
		if (LawOrder==1){
			ORuleID=myForm.Rule1.value;
		}else{
			ORuleID=myForm.Rule2.value;
		}
		window.open("Query_Law.asp?LawOrder="+LawOrder+"&RuleVer=<%
		if Not rs1.eof Then
			response.write trim(rs1("RuleVer"))
		end if	
		%>&ORuleID="+ORuleID,"WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	//}
}
function funGetSpeedRule(){
	<%if UpdateIllegalRuleFlag=1 then		'是否用車速判斷超速法條
	%>
	setIllegalRule();
	<%end if%>
}
//用地點、車速抓違規法條
function setIllegalRule(){
<%if sys_City="台東縣" then %>
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!="" && myForm.IllegalAddress.value!=""){
		if ((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
			IllegalRule=getIllegalRule3(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
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
<%else%>
	if (myForm.RuleSpeed.value!="" && myForm.IllegalSpeed.value!="" && myForm.IllegalAddress.value!=""){
		if ((myForm.Rule1.value.substr(0,2))!="29" && (myForm.Rule2.value.substr(0,2))!="29"){
			IllegalRule=getIllegalRule(myForm.IllegalAddress.value,myForm.RuleSpeed.value,myForm.IllegalSpeed.value,"",myForm.chkHighRoad.checked);
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
<%end if%>
}
function changeStreet(){
	//if (myForm.getStreetName.value!=""){
		myForm.kinds.value="getStreet";
		myForm.submit();
	//}
}
function NewWindow(Width, Height, URL, WinName){
	var nWidth = Width;
	var nHeight = Height;
	var sURL = URL;
	var nTop = 0;
	var nLeft = 0;
	var sWinSize = "left=" + nLeft + ",top=" + nTop + ",width=" + nWidth.toString(10) + ",height=" + nHeight.toString(10);
	var sWinStatus = "menubar=0,toolbar=0,scrollbars=1,resizable=1,status=0";
	var sWinName = WinName;
	OldObj = window.open(sURL,sWinName,sWinSize + "," + sWinStatus);
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
	//-----------上下左右-------------
	function funTextControl(obj){
		if (event.keyCode==13){ //Enter換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeEnter(obj.name);
		}else if (event.keyCode==38){ //上換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveLeft(obj.name);
		}else if (event.keyCode==40){ //下換欄
			event.keyCode=0;
			event.returnValue=false;
			CodeMoveRight(obj.name);
		}else if (event.keyCode==116){ 
			if (obj==myForm.Rule1){
				AutoGetRuleID(1);
			}else if (obj==myForm.Rule2){
				AutoGetRuleID(2);
			}
		}
	}
	//------------------------------


//開啟檢視圖
function OpenPic(FileName){
//alert(FileName);
	NewWindow(1000, 700, '../ProsecutionImage/ShowMap.asp?PicName=' + FileName.replace(/\+/g, '@2@'), 'MyPic');
}
//開啟詳細資料
function OpenDetail(FileName, SN){
	//+ URL 傳送時會不見,所以置換,到Server Side 再換回來
	NewWindow(1000, 600, '../ProsecutionImage/ProsecutionImageDetail.asp?FileName=' + FileName.replace(/\+/g, '@2@') + '&SN='+SN, 'MyDetail');
}
//開啟檢視圖
function OpenPic2(FileName){
	NewWindow(1000, 700, FileName, 'MyPic');
}


function ChangeImg2(){
<%if sPicWebPath2<>"" then%>
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImg2.src;

	myForm.SmallImg2.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;
<%end if%>
}

//=======================================================
var iDivHeight = 90; //放大?示?域?度
var iDivWidth = 212;//放大?示?域高度
var iMultiple = 3; //放大倍?

//?示放大?，鼠?移?事件和鼠???事件都??用本事件
//??：src代表?略?，sFileName放大?片名?
//原理：依据鼠????略?左上角（0，0）上的位置控制放大?左上角???示?域左上角（0，0）的位置
function show(src, sFileName)
{
//判?鼠?事件?生?是否同?按下了
if ((event.button == 1) && (event.ctrlKey == true))
  iMultiple -= 1;
else
  if (event.button == 1)
  iMultiple += 1;
if (iMultiple < 2) iMultiple = 2;

if (iMultiple > 14) iMultiple = 14;
  
var iPosX, iPosY; //放大????示?域左上角的坐?
var iMouseX = event.offsetX; //鼠????略?左上角的?坐?
var iMouseY = event.offsetY; //鼠????略?左上角的?坐?
var iBigImgWidth = src.clientWidth * iMultiple;  //放大??度，是?略?的?度乘以放大倍?
var iBigImgHeight = src.clientHeight * iMultiple; //放大?高度，是?略?的高度乘以放大倍?

if (iBigImgWidth <= iDivWidth)
{
  iPosX = (iDivWidth - iBigImgWidth) / 2;
}
else
{
  if ((iMouseX * iMultiple) <= (iDivWidth / 2))
  {
  iPosX = 0;
  }
  else
  {
  if (((src.clientWidth - iMouseX) * iMultiple) <= (iDivWidth / 2))
  {
    iPosX = -(iBigImgWidth - iDivWidth);
  }
  else
  {
    iPosX = -(iMouseX * iMultiple - iDivWidth / 2);
  }
  }
}

if (iBigImgHeight <= iDivHeight)
{
  iPosY = (iDivHeight - iBigImgHeight) / 2;
}
else
{
  if ((iMouseY * iMultiple) <= (iDivHeight / 2))
  {
	iPosY = 0;
  }
  else
  {
	  if (((src.clientHeight - iMouseY) * iMultiple) <= (iDivHeight / 2))
	  {
		iPosY = -(iBigImgHeight - iDivHeight);
	  }
	  else
	  {
		iPosY = -(iMouseY * iMultiple - iDivHeight / 2);
	  }
  }
}
div1.style.height = iDivHeight;
div1.style.width = iDivWidth;

myForm.BigImg.width = iBigImgWidth;
myForm.BigImg.height = iBigImgHeight;
myForm.BigImg.style.top = iPosY;
myForm.BigImg.style.left = iPosX;
}
//============================================================
function ChangeImg(){
	oBigImg=myForm.imgSource.src;
	oSmallImg=myForm.SmallImg.src;

	myForm.SmallImg.src=oBigImg;
	myForm.imgSource.src=oSmallImg;
	myForm.BigImg.src=oSmallImg;
}

myForm.CarNo.select();

<%
rs1.close
set rs1=nothing
conn.close
set conn=nothing
%>
</script>
</html>
