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
<title>裁罰資料建檔作業</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
'on error resume next
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
chkUnit=""
SeqUnit=0
If Trim(sys_City)="台南縣" and Year(now)="2007" Then
	SenUnit=split("新營分局,歸仁分局,新化分局,麻豆分局,善化分局,玉井分局,永康分局",",")
	SenNum=split("145,157,43,14,6,78,202",",")
	SqlUnit="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
	set rsUnit=conn.Execute(SqlUnit)
	If Not rsUnit.eof Then chkUnit=trim(rsUnit("UnitName"))
	rsUnit.close
	For i=0 to Ubound(SenUnit)
		If trim(chkUnit)=trim(SenUnit(i)) Then
			SeqUnit=cdbl(SenNum(i))
			exit for
		End if	
	Next
end if

'==========POST=========
'單號
if trim(request("billno"))="" then
	theBillno=""
else
	theBillno=left(trim(request("billno")),3)
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


'========================================================
if isempty(request("kinds")) then
	strSQL="select * from PasserBase where rownum=1"
	set rs=conn.execute(strSQL)
	If Not rs.eof Then
		For i=0 to rs.Fields.count-1
			If trim(rs.Fields.item(i).Name)="DRIVERSEX" Then Exit For
		Next
		If i>rs.Fields.count-1 Then
			strSQL="Alter Table PasserBase ADD (DRIVERSEX VarChar2(4))"
			conn.execute(strSQL)
		End if
	End If 
end if
'========================================================


'新增告發單
if trim(request("kinds"))="DB_Update" then
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
	'建檔日期
	theRecordDate=year(now)&"/"&month(now)&"/"&day(now)&" "&hour(now)&":"&minute(now)&":"&second(now)
	if trim(request("Billtype"))="" then '現在一律變為1 表示攔停
		theBilltype="1"
	else
		theBilltype=trim(request("Billtype"))
	end if
	'PasserBase
	zipid=trim(Request("DriverZip"))

	Sys_DriverAddress=request("DriverAddress")

	If ifnull(Request("DriverZip")) Then
	
		
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

		If ifnull(zipid) and isnumeric(left(request("DriverAddress"),1)) Then
			zipid=left(request("DriverAddress"),3)
			Sys_DriverAddress=mid(request("DriverAddress"),4)
		End if
	end if

	strUpd="update PasserBase set BillTypeID='"&trim(theBilltype)&"'" &_
		",BillNo='"&UCase(trim(request("Billno1")))&"',IllegalDate="&theIllegalDate&_
		",IllegalAddressID='"&trim(request("IllegalAddressID"))&"',IllegalAddress='"&trim(request("IllegalAddress"))&"'" &_
		",Rule1='"&trim(request("Rule1"))&"',ForFeit1="&theForFeit1 &_
		",Rule2='"&trim(request("Rule2"))&"',ForFeit2="&theForFeit2&",Rule3='"&trim(request("Rule3"))&"'" &_
		",ForFeit3="&theForFeit3&",Rule4='"&trim(request("Rule4"))&"',ForFeit4="&theForFeit4 &_
		",ProjectID='"&trim(request("ProjectID"))&"',DriverID='"&UCase(trim(request("DriverPID")))&"'" &_
		",DriverBirth="&theDriverBirth&",Driver='"&trim(request("DriverName"))&"'" &_
		",DriverAddress='"&trim(Sys_DriverAddress)&"',DriverZip='"&trim(zipid)&"'" &_
		",MemberStation='"&trim(request("MemberStation"))&"',BillUnitID='"&trim(request("BillUnitID"))&"'" &_
		",BillMemID1='"&trim(request("BillMemID1"))&"',BillMem1='"&trim(request("BillMemName1"))&"'" &_
		",BillMemID2='"&trim(request("BillMemID2"))&"',BillMem2='"&trim(request("BillMemName2"))&"'" &_
		",BillMemID3='"&trim(request("BillMemID3"))&"',BillMem3='"&trim(request("BillMemName3"))&"'" &_
		",BillMemID4='"&trim(request("BillMemID4"))&"',BillMem4='"&trim(request("BillMemName4"))&"'" &_
		",BillFillerMemberID='"&trim(request("BillMemID1"))&"',BillFiller='"&trim(request("BillMemName1"))&"'" &_
		",BillFillDate="&theBillFillDate&",DealLineDate="&theDealLineDate&_
		",Note='"&trim(request("Note"))&"',IsLecture='"&trim(request("IsLecture"))&"'" &_
		",DriverSex='"&trim(request("DriverSex"))&"',SignType='"&UCase(trim(request("SignType")))&"'" &_
		",DoubleCheckStatus="&trim(request("Sys_DoubleCheckStatus")) &_
		" where SN="&trim(request("PBillSN"))
		conn.execute strUpd
		ConnExecute strUpd,353
	'行人攤販行沒入物品 PasserConfiscate
	strDel="delete from PasserConfiscate where BillSN="&trim(request("PBillSN"))
	conn.execute strDel
	if trim(request("Fastener1"))<>"" then
		Ftemp=split(trim(request("Fastener1")),",")
		if ubound(Ftemp)>=0 then
			Fvaluetemp=split(Ftemp(0),"_")
			FID=trim(Fvaluetemp(0))
			Fvalue=trim(Fvaluetemp(1))
			strInsFastene1="insert into PasserConfiscate(BillSN,BillNo,Confiscate,ConfiscateID)" &_
					" values("&trim(request("PBillSN"))&",'"&UCase(trim(request("Billno1")))&"','"&Fvalue&"','"&FID&"')"
			conn.execute strInsFastene1
			ConnExecute strInsFastene1,353
		end if
		if ubound(Ftemp)>=1 then
			Fvaluetemp=split(Ftemp(1),"_")
			FID=trim(Fvaluetemp(0))
			Fvalue=trim(Fvaluetemp(1))
			strInsFastene2="insert into PasserConfiscate(BillSN,BillNo,Confiscate,ConfiscateID)" &_
						" values("&trim(request("PBillSN"))&",'"&UCase(trim(request("Billno1")))&"','"&Fvalue&"','"&FID&"')"
			conn.execute strInsFastene2
			ConnExecute strInsFastene2,353
		end if
		if ubound(Ftemp)>=2 then
			Fvaluetemp=split(Ftemp(2),"_")
			FID=trim(Fvaluetemp(0))
			Fvalue=trim(Fvaluetemp(1))
			strInsFastene3="insert into PasserConfiscate(BillSN,BillNo,Confiscate,ConfiscateID)" &_
						" values("&trim(request("PBillSN"))&",'"&UCase(trim(request("Billno1")))&"','"&Fvalue&"','"&FID&"')"
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

'	if trim(request("Billno1"))="" then
'		theBillno=""
'	else
'		theBillno=left(trim(request("Billno1")),3)
'	end if
	'PasserBase
	Sys_DoubleCheckStatus=request("Sys_DoubleCheckStatus")
	if ifnull(request("Sys_DoubleCheckStatus")) then
		strSQL="select NVL(Max(DoubleCheckStatus),0) as DoubleCheckStatus from (select DoubleCheckStatus from passerBase where recordMemberID in (select MemberID from MemberData where UnitID='"&trim(Session("Unit_ID"))&"') and TO_CHAR(RecordDate,'YYYY')=TO_CHAR(sysdate,'YYYY') order by RecordDate DESC) where rownum=1"
		set rssum=conn.execute(strSQL)
		Sys_DoubleCheckStatus=CDBL(rssum("DoubleCheckStatus"))+1
		rssum.close
	end if
	strSQL="select BillNo from PasserBase where BillNo='"&UCase(trim(request("Billno1")))&"' and RecordStateId <> -1"
	set rsbill=conn.execute(strSQL)
	If rsbill.eof Then
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
		
		Sys_DriverAddress=request("DriverAddress")
		If ifnull(zipid) and isnumeric(left(request("DriverAddress"),1)) Then
			zipid=left(request("DriverAddress"),3)
			Sys_DriverAddress=mid(request("DriverAddress"),4)
		End if
		

		strInsert="insert into PasserBase(SN,BillTypeID,BillNo,IllegalDate" & _
				",IllegalAddressID,IllegalAddress,Rule1,ForFeit1" &_
				",Rule2,ForFeit2,Rule3,ForFeit3,Rule4,ForFeit4,ProjectID" &_
				",DriverID,DriverBirth,Driver,DriverAddress,DriverZip" &_
				",MemberStation,BillUnitID,BillMemID1,BillMem1" &_
				",BillMemID2,BillMem2,BillMemID3,BillMem3,BillMemID4,BillMem4" &_
				",BillFillerMemberID,BillFiller" &_
				",BillFillDate,DealLineDate,BillStatus,RecordStateID,RecordDate,RecordMemberID" &_
				",Note,RuleVer,IsLecture,DriverSex,SignType,DoubleCheckStatus)" &_
				" values(passerbase_seq.nextval, '"&trim(theBilltype)&"','"&UCase(trim(request("Billno1")))&"'" &_
				","&theIllegalDate&",'"&trim(request("IllegalAddressID"))&"'" &_
				",'"&trim(request("IllegalAddress"))&"','"&trim(request("Rule1"))&"',"&theForFeit1&",'"&trim(request("Rule2"))&"'" &_
				","&theForFeit2&",'"&trim(request("Rule3"))&"',"&theForFeit3&",'"&trim(request("Rule4"))&"'" &_
				","&theForFeit4&",'"&trim(request("ProjectID"))&"'" &_
				",'"&UCase(trim(request("DriverPID")))&"',"&theDriverBirth&",'"&trim(request("DriverName"))&"'" &_
				",'"&trim(Sys_DriverAddress)&"','"&trim(zipid)&"','"&trim(request("MemberStation"))&"'" &_
				",'"&trim(request("BillUnitID"))&"','"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
				",'"&trim(request("BillMemID2"))&"','"&trim(request("BillMemName2"))&"'" &_
				",'"&trim(request("BillMemID3"))&"','"&trim(request("BillMemName3"))&"'" &_
				",'"&trim(request("BillMemID4"))&"','"&trim(request("BillMemName4"))&"'" &_
				",'"&trim(request("BillMemID1"))&"','"&trim(request("BillMemName1"))&"'" &_
				","&theBillFillDate&","&theDealLineDate&",'0','0',SYSDate,'"&theRecordMemberID&"'" &_
				",'"&trim(request("Note"))&"','"&theRuleVer&"','"&trim(request("IsLecture"))&"'" &_
				",'"&trim(request("DriverSex"))&"','"&UCase(trim(request("SignType")))&"',"&trim(Sys_DoubleCheckStatus)&")"

				ConnExecute strInsert,354
				conn.execute strInsert
				'response.write strInsert
				'response.end
	'查流水號
		strSN="select SN from PasserBase where BillNo='"&UCase(trim(request("Billno1")))&"'"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theSN=trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing
		'行人攤販行沒入物品 PasserConfiscate
		if trim(request("Fastener1"))<>"" then
			Ftemp=split(trim(request("Fastener1")),",")
			if ubound(Ftemp)>=0 then
				Fvaluetemp=split(Ftemp(0),"_")
				FID=trim(Fvaluetemp(0))
				Fvalue=trim(Fvaluetemp(1))
				strInsFastene1="insert into PasserConfiscate(BillSN,BillNo,Confiscate,ConfiscateID)" &_
						" values("&theSN&",'"&UCase(trim(request("Billno1")))&"','"&Fvalue&"','"&FID&"')"
				conn.execute strInsFastene1
			end if
			if ubound(Ftemp)>=1 then
				Fvaluetemp=split(Ftemp(1),"_")
				FID=trim(Fvaluetemp(0))
				Fvalue=trim(Fvaluetemp(1))
				strInsFastene2="insert into PasserConfiscate(BillSN,BillNo,Confiscate,ConfiscateID)" &_
							" values("&theSN&",'"&UCase(trim(request("Billno1")))&"','"&Fvalue&"','"&FID&"')"
				conn.execute strInsFastene2
			end if
			if ubound(Ftemp)>=2 then
				Fvaluetemp=split(Ftemp(2),"_")
				FID=trim(Fvaluetemp(0))
				Fvalue=trim(Fvaluetemp(1))
				strInsFastene3="insert into PasserConfiscate(BillSN,BillNo,Confiscate,ConfiscateID)" &_
							" values("&theSN&",'"&UCase(trim(request("Billno1")))&"','"&Fvalue&"','"&FID&"')"
				conn.execute strInsFastene3
			end if
		end if
	end if
	rsbill.close
end if
'If year(now)="2007" Then
'	strSQL="Update PasserBase a set DoubleCheckStatus=(select cnt+157 Tcnt from(select rownum cnt,passerBase.* from (select * from passerBase order by recorddate desc) passerBase where recordMemberID in(select MemberID from MemberData where UnitID='"&Session("Unit_ID")&"'))where sn=a.sn) where recordMemberID in(select MemberID from MemberData where UnitID='"&Session("Unit_ID")&"')"
'	conn.execute(strSQL)
'End if

If ifnull(request("DBCunt") ) Then
	strSQL="select count(*) as cnt from PasserBase where RecorDStateID<>-1 and RecordMemberID in(select MemberID from MemberData where UnitID='"&trim(Session("Unit_ID"))&"' and RecorDStateID<>-1)"
	set rssum=conn.execute(strSQL)
	DBsume=rssum("cnt")
	rssum.close
else
	DBsume=request("DBCunt")+1
End if

strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
rsUnit.close

if Cint(Session("UnitLevelID"))=2 and Instr(strUnitName,"組") >0 and strCity<>"南投縣" then
	strSQLUnit="select UnitID from UnitInfo where UnitID=(select UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"') or UnitTypeID=(select UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"')"
elseif Cint(Session("UnitLevelID"))=2 then
	strSQLUnit="select UnitID from UnitInfo where UnitTypeID='"&trim(Session("Unit_ID"))&"'"
elseif Cint(Session("UnitLevelID"))>2 then
	strSQLUnit="select UnitID from UnitInfo where UnitID=(select UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"') or UnitTypeID=(select UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"')"
else
	strSQLUnit="select UnitID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
end if



strSQL="select NVL(Max(To_number(DoubleCheckStatus)),0) as DoubleCheckStatus from (select DoubleCheckStatus from passerBase where recordMemberID in (select MemberID from MemberData where UnitID='"&trim(Session("Unit_ID"))&"') and TO_CHAR(RecordDate,'YYYY')=TO_CHAR(sysdate,'YYYY')  order by RecordDate DESC,DoubleCheckStatus DESC) where rownum=1"
set rssum=conn.execute(strSQL)
FileSeq=CDBL(rssum("DoubleCheckStatus"))+1
rssum.close

chkdate=gInitDT(dateadd("d",14,now))
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
	font-size: 13px
}
.btn5 {BORDER-RIGHT: #cccccc 1px solid; BORDER-TOP: #cccccc 1px solid; FONT-SIZE: 11pt;
       BORDER-LEFT: #cccccc 1px solid; COLOR: #000000; LINE-HEIGHT: 15px; BORDER-BOTTOM: #cccccc 1px solid;
       FONT-FAMILY: Arial; BACKGROUND-COLOR: #FFFFF0}
-->
</style>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onkeydown="KeyDown()">
<!-- #include file="../Common/bannernoimagepasser.asp"-->
	<form name="myForm" method="post">  
		<table width='993' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="6" nowrap><strong>慢車行人攤販建檔作業</strong>&nbsp; &nbsp; 日期格式：951220 &nbsp;時間格式：2300(24小時制)&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 建檔日期：<span id="CreatDate"><%=ginitdt(now)%></span>&nbsp;&nbsp;&nbsp;第<span id="Sumefile"><%=cint(DBsume)+1%></span>筆</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right"><span class="style4">＊</span>單號</td>
				<td ><input name="Billno1" type="text" class="btn5" value="<%=theBillno%>" size="10" maxlength="9"  onblur="CheckPeopleBillNoExist();" onKeyUp="funAutoTextControl(this);" onkeydown="funTextControl(this);" style=ime-mode:disabled></td>
				<td bgcolor="#FFFFCC" align="right">違規人姓名</td>
				<td><input type="text" class="btn5" size="10" value="" name="DriverName" onkeydown="funTextControl(this);" onBlur="NameLoadDate();" style=ime-mode:active></td>
				</td>
				<td bgcolor="#FFFFCC" align="right">違規人出生日期</td>
				<td>
				<input type="text" class="btn5" size="10" maxlength="7" value="" name="DriverBrith" onBlur="focusToDriverPID()" onkeydown="funTextControl(this);"  onKeyUp="funAutoTextControl(this);" style=ime-mode:disabled>
				</td>				
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right" nowrap>
					<span class="style4" >＊</span>違規人身份證號
				</td>
				<td>
					<input type="text" class="btn5" size="10" maxlength="19" value="" name="DriverPID" onBlur="FuncChkPID();" onkeydown="funTextControl(this);" onKeyUp="funAutoTextControl(this);" style="ime-mode:disabled">
					<table border=0>
						<td bgcolor="#FFFFCC" align="right" nowrap>
							<span class="style4" >＊</span>性別
						</td>
						<td>
							<input type="text" class="btn5" size="10" value="" name="DriverSEX" onFocus="chkDriverSEX();" onkeydown="funTextControl(this);" style="ime-mode:active"></td>
						</td>
					</table>
				</td>
				<td bgcolor="#FFFFCC" align="right">違規人地址</td>

				<td colspan="3">
				<input type="text" class="btn5" size="1" value="" name="DriverZip"  onBlur="getDriverZip(this,'DriverAddress');" onkeydown="funTextControl(this);">
				區號
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick="QryDriverZip();">

				<input type="text" class="btn5" size="40" value="" name="DriverAddress" onkeydown="funTextControl(this);" style=ime-mode:active>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right"><span class="style4">＊</span>違規日期</td>
				<td>
				<input type="text" class="btn5" size="10" value="<%=request("IllegalDate")%>" maxlength="7" name="IllegalDate" onBlur="getDealLineDate();" onkeydown="funTextControl(this);"  onKeyUp="funAutoTextControl(this);" style=ime-mode:disabled>
				</td>
				<td bgcolor="#FFFFCC" align="right"><span class="style4">＊</span>違規時間</td>
				<td colspan="3">
				<input type="text" class="btn5" size="10" value="" maxlength="4" name="IllegalTime" onBlur="value=value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);"  onKeyUp="funAutoTextControl(this);" style=ime-mode:disabled>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">違規地點代碼</td>
				<td >
					<input type="text" class="btn5" size="10" value="<%=trim(request("IllegalAddressID"))%>" name="IllegalAddressID" maxlength="6" onKeyUp="getillStreet();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage_Street_People","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#FFFFCC" align="right"><span class="style4">＊</span>違規地點</td>
				<td colspan="3">
					<%if sys_City="台南市" then %>
						<input type="text" class="btn5" size="1" value="" name="IllegalZip" onBlur="getDriverZip(this,'IllegalAddress');" onkeydown="funTextControl(this);">
						區號
						<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick="QryIllegalZip();">
					<%end if%>
					<input type="text" class="btn5" size="46" value="<%
						response.write request("IllegalAddress")
					%>" name="IllegalAddress" style=ime-mode:active onkeyup="AutoGetIllStreet();" onkeydown="funTextControl(this);"<%
					if sys_City="台南市" Then Response.Write " onfocus=""autoKeyEnd();"""
					%>>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right"><span class="style4">＊</span>違規法條一</td>
				<td colspan="5">
					<input type="text" class="btn5" size="10" value="<%=request("Rule1")%>" name="Rule1" onKeyUP="getRuleData1();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
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
				<td bgcolor="#FFFFCC" align="right">違規法條二</td>
				<td colspan="5">
					<input type="text" class="btn5" size="10" value="<%=request("Rule2")%>" name="Rule2" onKeyUP="getRuleData2();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
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
				<td bgcolor="#FFFFCC" align="right"><span class="style4">＊</span>應到案日期</td>
				<td>
					<input type="text" class="btn5" size="10" value="" maxlength="7" name="DealLineDate" onBlur="value=value.replace(/[^\d]/g,'')" onKeyUp="funAutoTextControl(this);" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
				<td bgcolor="#FFFFCC" align="right" nowrap><span class="style4">＊</span>應到案處所代碼</td>
				<td colspan="1">
					<input type="text" class="btn5" size="4" value="<%=request("MemberStation")%>" name="MemberStation" onKeyUP="getStation();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=S","WebPage1","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer5" style="position:absolute ; width:250px; height:30px; z-index:0;layer-background-color: #CCFFFF; border: 1px none #000000;"><%
						strStation="select UnitName from UnitInfo where UnitID='"&trim(request("MemberStation"))&"'"
						set rsStation=conn.execute(strStation)
						if not rsStation.eof then
							response.write trim(rsStation("UnitName"))
						end if
						rsStation.close
						set rsStation=nothing
					%>
					</div>
					</span>
				</td>				
				<td bgcolor="#FFFFCC" align="right"><span class="style4">＊</span>舉發人<%
						if sys_City<>"高雄縣" and sys_City<>"高雄市" then 
							response.write "姓名"

						else
							response.write "代碼"

						end if%>1
				</td>
		  		<td>
					<input type="text" class="btn5" size="6" value="<%=request("BillMem1")%>" name="BillMem1" onblur="chkBillMemID1();" onFocusIn="fun_chkInput(this);" onKeyUP="getBillMemID1();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemType=P&MemOrder=1","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer12" style="position:absolute ; width:92px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"><%
					if trim(request("BillMem1"))<>"" then
						strMem="select a.ChName,a.CreditID,a.MemberID,a.LoginID,a.UnitID,b.UnitName from MemberData a,UnitInfo b where a.UnitID=b.UnitID and a.ChName like '%"&trim(request("BillMem1"))&"%' and a.AccountStateID=0 and a.RecordstateID=0"
						set rsMem=conn.execute(strMem)
						if not rsMem.eof then response.write trim(rsMem("LoginID"))
						rsMem.close
					end if
					%></div>
					<input type="hidden" value="<%=request("BillMemID1")%>" name="BillMemID1">
					<input type="hidden" value="<%=request("BillMemName1")%>" name="BillMemName1">
				</td>
			</tr>			
			<tr>
				
				<td bgcolor="#FFFFCC" align="right" width="14%">舉發人<%
						if sys_City<>"高雄縣" and sys_City<>"高雄市" then 
							response.write "姓名"

						else
							response.write "代碼"

						end if%>2
				</td>
				<td width="20%">
					<input type="text" class="btn5" size="6" value="<%=request("BillMem2")%>" name="BillMem2" onblur="chkBillMemID2();" onFocusIn="fun_chkInput(this);" onKeyUP="getBillMemID2();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemType=P&MemOrder=2","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer13" style="position:absolute ; width:92px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"><%
					if trim(request("BillMem2"))<>"" then
						strMem="select a.ChName,a.CreditID,a.MemberID,a.LoginID,a.UnitID,b.UnitName from MemberData a,UnitInfo b where a.UnitID=b.UnitID and a.ChName like '%"&trim(request("BillMem2"))&"%' and a.AccountStateID=0 and a.RecordstateID=0"
						set rsMem=conn.execute(strMem)
						if not rsMem.eof then response.write trim(rsMem("LoginID"))
						rsMem.close
					end if
					%></div>
					<input type="hidden" value="<%=request("BillMemID2")%>" name="BillMemID2">
					<input type="hidden" value="<%=request("BillMemName2")%>" name="BillMemName2">
				</td>
				<td bgcolor="#FFFFCC" align="right" width="13%">舉發人<%
						if sys_City<>"高雄縣" and sys_City<>"高雄市" then 
							response.write "姓名"

						else
							response.write "代碼"

						end if%>3
				</td>
				<td width="20%">
					<input type="text" class="btn5" size="6" value="<%=request("BillMem3")%>" name="BillMem3" onblur="chkBillMemID3();" onFocusIn="fun_chkInput(this);" onKeyUP="getBillMemID3();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemType=P&MemOrder=3","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer14" style="position:absolute ; width:92px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
					if trim(request("BillMem3"))<>"" then
						strMem="select a.ChName,a.CreditID,a.MemberID,a.LoginID,a.UnitID,b.UnitName from MemberData a,UnitInfo b where a.UnitID=b.UnitID and a.ChName like '%"&trim(request("BillMem3"))&"%' and a.AccountStateID=0 and a.RecordstateID=0"
						set rsMem=conn.execute(strMem)
						if not rsMem.eof then response.write trim(rsMem("LoginID"))
						rsMem.close
					end if
					%></div>
					<input type="hidden" value="<%=request("BillMemID3")%>" name="BillMemID3">
					<input type="hidden" value="<%=request("BillMemName3")%>" name="BillMemName3">
				</td>
				<td bgcolor="#FFFFCC" align="right" width="13%">舉發人<%
						if sys_City<>"高雄縣" and sys_City<>"高雄市" then 
							response.write "姓名"

						else
							response.write "代碼"

						end if%>4
				</td>
				<td width="20%">
					<input type="text" class="btn5" size="6" value="<%=request("BillMem4")%>" name="BillMem4" onblur="chkBillMemID4();" onFocusIn="fun_chkInput(this);" onKeyUP="getBillMemID4();" onkeydown="funTextControl(this);">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemType=P&MemOrder=4","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer17" style="position:absolute ; width:92px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"><%
					if trim(request("BillMem4"))<>"" then
						strMem="select a.ChName,a.CreditID,a.MemberID,a.LoginID,a.UnitID,b.UnitName from MemberData a,UnitInfo b where a.UnitID=b.UnitID and a.ChName like '%"&trim(request("BillMem4"))&"%' and a.AccountStateID=0 and a.RecordstateID=0"
						set rsMem=conn.execute(strMem)
						if not rsMem.eof then response.write trim(rsMem("LoginID"))
						rsMem.close
					end if
					%></div>
					<input type="hidden" value="<%=request("BillMemID4")%>" name="BillMemID4">
					<input type="hidden" value="<%=request("BillMemName4")%>" name="BillMemName4">
				</td>
			</tr>

			<tr>
				<td height="33" bgcolor="#FFFFCC" align="right">代保管物</td>
				<td nowrap>
<%
	strItem="select * from Code where TypeID=2 and Not(ID<478 or ID=479) order by ID"
	set rsItem=conn.execute(strItem)
	If Not rsItem.Bof Then rsItem.MoveFirst 
	While Not rsItem.Eof
%>
					<input type="checkbox" class="btn5" value="<%=trim(rsItem("ID"))&"_"&trim(rsItem("Content"))%>" name="Fastener1"><%=trim(rsItem("Content"))%>&nbsp;
<%	
	rsItem.MoveNext
	Wend
	rsItem.close
	set rsItem=nothing

%>
				</td>
				<td bgcolor="#FFFFCC"><div align="right"><span class="style4">＊</span>簽收狀況</div></td>
				<td colspan="3">
					<input type="text" class="btn5" size="5" value="A" maxlength="1" name="SignType" onBlur="funcSignType();" onKeyUp="funAutoTextControl(this);" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<font color="#ff000" size="2">
					A簽收/ U拒簽收/ 2拒簽已收/ 3已簽拒收/ 5補開單
					</font>
				</td>
			</tr>				
	<tr height="6"><td></td></tr>		
			<tr>
		  	<td bgcolor="#FFFFCC" align="right"><span class="style4">＊</span>舉發單位</td>
				<td>
					<input type="text" class="btn5" size="4" value="<%=request("BillUnitID")%>" name="BillUnitID" onKeyUP="getUnit();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=800,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer6" style="position:absolute ; width:250px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"><%
						strStation="select UnitName from UnitInfo where UnitID='"&trim(request("BillUnitID"))&"'"
						set rsStation=conn.execute(strStation)
						if not rsStation.eof then
							response.write trim(rsStation("UnitName"))
						end if
						rsStation.close
						set rsStation=nothing
					%></div>
					</span>
				</td>
				<td bgcolor="#FFFFCC"><div align="right">專案代碼</div></td>
				<td colspan="3">
					<input type="text" class="btn5" size="5" value="" name="ProjectID" style=ime-mode:disabled onkeyup="ProjectF5()" onkeydown="funTextControl(this);">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onClick='window.open("Query_Project.asp","WebPage1","left=0,top=0,location=0,width=700,height=455,resizable=yes,scrollbars=yes")'>
				<span class="style5">
				<div id="Layer001" style="position:absolute ; width:180px; height:30px; z-index:0;  layer-background-color: #CCFFFF; border: 1px none #000000;"></div>	
				</span>
				</td>

				<!--<td bgcolor="#FFFFCC" align="right">是否講習</td>
				<td>
					<input type="radio" value="1" name="IsLecture">是
					<input type="radio" value="0" name="IsLecture" checked>否
				</td>
				<td bgcolor="#FFFFCC" align="right">告發類別</td>
				<td colspan="1">
				<input type="text" size="4" maxlength="1" value="<%=theBilltype%>" name="BillType" onBlur="chkBillType()" style=ime-mode:disabled>
				<font color="#ff000" size="2">1慢車/2行人/3道路障礙</font>&nbsp;&nbsp;
				</td>-->
			</tr>	
				
			<tr>
				<td bgcolor="#FFFFCC" align="right"><span class="style4">＊</span>填單日期</td>
				<td>
					<input type="text" class="btn5" size="10" value="" maxlength="7" name="BillFillDate" onBlur="value=value.replace(/[^\d]/g,'')" onKeyUp="funAutoTextControl(this);" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>

					<td bgcolor="#FFFFCC" align="right">備註</td>
					<td colspan="3">
						<input type="text" class="btn5" size="46" value="" name="Note" onkeydown="funTextControl(this);" style=ime-mode:active>
					</td>	
			</tr>			

			<tr>
			  <td bgcolor="#FFDD77" align="center" colspan="6">
					<font color="red"><B>建檔序號第<span id="Seqfile"><input type="text" value="<%=FileSeq%>" class="btn1" size="10" name="Sys_DoubleCheckStatus" onkeyup="value=value.replace(/[^\d]/g,'')"></span>號</B></font>
					&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="button" value="儲 存 F2" onclick="InsertBillVase();" <%
				'1:查詢 ,2:新增 ,3:修改 ,4:刪除
				if CheckPermission(235,2)=false then
					response.write "disabled"
				end if
					%> class="btn1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit1343" onClick="location='BillKeyIn_People.asp'" value="清 除 F4" class="btn1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit5322" onClick="funcOpenBillQry();" value="查 詢 <%=F6StrName%>" class="btn1">
					<img src="/image/space.gif" width="29" height="8">
					<input type="button" name="Submit4232" onClick="funPrintCaseList_Stop();" value="建檔清冊 F10" class="btn1">
					<input type="hidden" value="" name="kinds">
                    <span class="style1"><span class="style3"><img src="/image/space.gif" width="29" height="8"></span>
                    <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉 F8" class="btn1">
					<br>
					<input type="button" name="Submit5322" onClick="funDBfisrt();" value="第一筆" class="btn1">
					<input type="button" name="Submit5322" onClick="funDBupmove();" value="上一筆 PgUp" class="btn1">
					<input type="button" name="Submit5322" onClick="funDBdownmove();" value="下一筆 PgDn" class="btn1">
					<input type="button" name="Submit5322" onClick="funDBlast();" value="最後一筆" class="btn1">
					<input type="button" name="Submit1343" onClick="location='BillKeyIn_People.asp'" value="新增一筆" class="btn1">
                </span>
				<!-- 違規人性別 -->
				<input type="hidden" value="" name="DriverSex">
				<input type="hidden" value="" name="DBFile">
				<input type="hidden" value="<%=DBsume%>" name="DBCunt">
				<input type="hidden" value="" name="PBillSN">
				<input type="hidden" value="" name="Mem">
				<input type="hidden" value="" name="MemOrder">
				<input type="hidden" value="" name="MemType">
			  </td>
			</tr>
		</table>		
	</form>
<%
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
var sys_City="<%=sys_City%>";

if(sys_City="台南市"){

MoveTextVar("Billno1,DriverName,DriverBrith||DriverPID,DriverSEX,DriverZip,DriverAddress||IllegalDate,IllegalTime||IllegalAddressID,IllegalZip,IllegalAddress||Rule1||Rule2||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||SignType||BillUnitID||BillFillDate,Note");

}else{

MoveTextVar("Billno1,DriverName,DriverBrith||DriverPID,DriverSEX,DriverZip,DriverAddress||IllegalDate,IllegalTime||IllegalAddressID,IllegalAddress||Rule1||Rule2||DealLineDate,MemberStation,BillMem1||BillMem2,BillMem3,BillMem4||SignType||BillUnitID||BillFillDate,Note");

}

function funDBupmove(){
	if(myForm.DBFile.value==''){
		myForm.DBFile.value=myForm.DBCunt.value;
		Sumefile.innerHTML=myForm.DBFile.value;
		runServerScript("BillDBmove.asp?DBFile="+myForm.DBFile.value);
	}else if(parseInt(myForm.DBFile.value)>1){
		myForm.DBFile.value=parseInt(myForm.DBFile.value)-1;
		Sumefile.innerHTML=myForm.DBFile.value;
		runServerScript("BillDBmove.asp?DBFile="+myForm.DBFile.value);
	}
}
function funDBdownmove(){
	if(myForm.DBFile.value!=''){
		if(parseInt(myForm.DBFile.value)<parseInt(myForm.DBCunt.value)){
			myForm.DBFile.value=parseInt(myForm.DBFile.value)+1;
		}
		Sumefile.innerHTML=myForm.DBFile.value;
	}
	runServerScript("BillDBmove.asp?DBFile="+myForm.DBFile.value);
}
function funDBfisrt(){
	myForm.DBFile.value=1;
	Sumefile.innerHTML=myForm.DBFile.value;
	runServerScript("BillDBmove.asp?DBFile="+myForm.DBFile.value);
}
function funDBlast(){
	myForm.DBFile.value=myForm.DBCunt.value;
	Sumefile.innerHTML=myForm.DBFile.value;
	runServerScript("BillDBmove.asp?DBFile="+myForm.DBFile.value);
}
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
	if (myForm.DriverPID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規人身份證號碼。";
	}

	if (myForm.,DriverSEX.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請輸入違規人性別。";
	}

	/*if (myForm.BillType.value==""){
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
	}else if (myForm.Rule1.value.substr(0,2)<69){
		 if (myForm.Rule1.value.substr(0,2)!=36&&myForm.Rule1.value.substr(0,2)!=37){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條一輸入錯誤。";
		}
	}
	if (myForm.Rule1.value==myForm.Rule2.value && myForm.Rule1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規法條一與違規法條二重複。";
	}
	if (myForm.Rule2.value!=""){
		if (TDLawErrorLog2==1){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規法條二輸入錯誤。";
		}else if (myForm.Rule2.value.substr(0,2)<69){
			if (myForm.Rule2.value.substr(0,2)!=36&&myForm.Rule2.value.substr(0,2)!=37){
				error=error+1;
				errorString=errorString+"\n"+error+"：違規法條二輸入錯誤。";
			}
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
	}else if (eval(myForm.DealLineDate.value)<=eval(myForm.BillFillDate.value)){
		error=error+1;
		errorString=errorString+"\n"+error+"：應到案日期要大於填單日期。";
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
		errorString=errorString+"\n"+error+"：請輸入舉發人姓名。";
	}else if (TDMemErrorLog1==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人姓名1 輸入錯誤。";
	}
	if (TDMemErrorLog2==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人姓名2 輸入錯誤。";
	}
	if (TDMemErrorLog3==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人姓名3 輸入錯誤。";
	}
	if (TDMemErrorLog4==1){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人姓名4 輸入錯誤。";
	}
	if (myForm.BillMem1.value==myForm.BillMem2.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人姓名1 與 舉發人姓名2 重複。";
	}else if (myForm.BillMem1.value==myForm.BillMem3.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人姓名1 與 舉發人姓名3 重複。";
	}else if (myForm.BillMem1.value==myForm.BillMem4.value && myForm.BillMem1.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人姓名1 與 舉發人姓名4 重複。";
	}
	if (myForm.BillMem2.value==myForm.BillMem3.value && myForm.BillMem2.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人姓名2 與 舉發人姓名3 重複。";
	}else if (myForm.BillMem2.value==myForm.BillMem4.value && myForm.BillMem2.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人姓名2 與 舉發人姓名4 重複。";
	}
	if (myForm.BillMem3.value==myForm.BillMem4.value && myForm.BillMem3.value!=""){
		error=error+1;
		errorString=errorString+"\n"+error+"：舉發人姓名3 與 舉發人姓名4 重複。";
	}
	if (myForm.BillFillDate.value < myForm.IllegalDate.value){
		if(!confirm('違規日期比填單日晚，請確定是否要存檔!!')){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規日期不得比填單日晚。";
		}
	}else if(TodayDate < myForm.IllegalDate.value){
		error=error+1;
		errorString=errorString+"\n"+error+"：違規日期不得比今天晚。";
	}
	if (error==0){
		if(myForm.PBillSN.value!=''){
			myForm.kinds.value="DB_Update";
			myForm.submit();
		}else{
			myForm.PBillSN.value='';
			myForm.kinds.value="DB_insert";
			myForm.submit();
		}
	}else{
		alert(errorString);
	}
}

function getDriverZip(obj,objName){
	if(obj.value!=''&&obj.value.length>2){
		runServerScript("getZipName.asp?ZipID="+obj.value+"&getZipName="+objName);
	}else if(obj.value!=''&&obj.value.length<3){
		alert("郵遞區號錯誤!!");
	}
}

function QryIllegalZip(){
	window.open("Query_Zip.asp?ZipCity="+sys_City+"&IllegalZip="+myForm.IllegalZip.value+"&ObjName=IllegalZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");

}

function QryDriverZip(){
	window.open("Query_Zip.asp?ZipCity=&IllegalZip="+myForm.DriverZip.value+"&ObjName=DriverZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");

}

//違規事實1(ajax)
function getRuleData1(){
	if (myForm.Rule1.value.length > 6){
		var Rule1Num=myForm.Rule1.value;
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail.asp?RuleOrder=1&RuleID="+Rule1Num+"&RuleVer="+VerNo);

		if (myForm.ForFeit1.value!=''&myForm.Rule1.value!=''){
			funAutoCodeEnter(myForm.Rule1);
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
		var VerNo=<%=theRuleVer%>;
		runServerScript("getRuleDetail.asp?RuleOrder=2&RuleID="+Rule2Num+"&RuleVer="+VerNo);
		
		if (myForm.ForFeit2.value!=''&myForm.Rule2.value!=''){
			funAutoCodeEnter(myForm.Rule2);
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
//到案處所(ajax)
function getStation(){
	if (myForm.MemberStation.value.length > 1){
		var StationNum=myForm.MemberStation.value;
		runServerScript("getMemberStation2.asp?StationID="+StationNum);
		funAutoTextControl(myForm.MemberStation);
	}else{
		Layer5.innerHTML=" ";
		TDStationErrorLog=1;
	}
}
//舉發單位(ajax)
function getUnit(){
	if (myForm.BillUnitID.value.length > 1){
		var BillUnitNum=myForm.BillUnitID.value;
		var billmem=myForm.BillMemID1.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum+"&BillMem="+billmem);

		if (Layer6.innerHTML!=''&myForm.BillUnitID.value!=''){
			funAutoCodeEnter(myForm.BillUnitID);
		}

	}else{
		Layer6.innerHTML=" ";
		TDUnitErrorLog=1;
	}
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
//違規地點代碼(ajax)
function getillStreet(){
	if (event.keyCode==<%=F5str%>){	
		event.keyCode=0;
		event.returnValue=false;
		Ostreet=myForm.IllegalAddressID.value;
		window.open("Query_Street.asp?OStreetID="+Ostreet,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.IllegalAddressID.value.length!=''){
		var illAddrNum=myForm.IllegalAddressID.value;
		runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
		funAutoTextControl(myForm.IllegalAddressID);
	}
}
//舉發人一(ajax)
function fun_chkInput(obj){
	if(sys_City=='高雄市'){obj.style.imeMode="disabled";}
}

function chkBillMemID1(){
	if (myForm.BillMem1.value!=''&&myForm.BillMemID1.value==''){
		alert('無此員警資料，請確認人員管理是否有該資料紀錄!!');
	}
}
function chkBillMemID2(){
	if (myForm.BillMem2.value!=''&&myForm.BillMemID2.value==''){
		alert('無此員警資料，請確認人員管理是否有該資料紀錄!!');
	}
}
function chkBillMemID3(){
	if (myForm.BillMem3.value!=''&&myForm.BillMemID3.value==''){
		alert('無此員警資料，請確認人員管理是否有該資料紀錄!!');
	}
}
function chkBillMemID4(){
	if (myForm.BillMem4.value!=''&&myForm.BillMemID4.value==''){
		alert('無此員警資料，請確認人員管理是否有該資料紀錄!!');
	}
}
function getBillMemID1(){
	if (event.keyCode==<%=F5str%>){	
		event.keyCode=0;
		myForm.Mem.value=myForm.BillMem1.value;
		myForm.MemType.value='P';
		myForm.MemOrder.value=1;
		myForm.kinds.value='DB_Select';
		UrlStr="Query_MemID.asp";		
		myForm.action=UrlStr;
		myForm.target="WebPage_Street_People";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.kinds.value='';
	}else{
		if (myForm.BillMem1.value.length > 1){
			var BillMemNum=myForm.BillMem1.value;
			runServerScript("getBillPeoPleMemID.asp?MType=People&MemOrder=1&MemID="+BillMemNum);
			
			if (myForm.BillMem1.value!=''&&myForm.BillMemID1.value!=''){
				funAutoCodeEnter(myForm.BillMem1);
			}		
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
		myForm.Mem.value=myForm.BillMem2.value;
		myForm.MemType.value='P';
		myForm.MemOrder.value=2;
		UrlStr="Query_MemID.asp";
		myForm.kinds.value='DB_Select';
		myForm.action=UrlStr;
		myForm.target="WebPage_Street_People";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.kinds.value='';
	}else{
		if (myForm.BillMem2.value.length > 1){
			var BillMemNum=myForm.BillMem2.value;
			runServerScript("getBillPeoPleMemID.asp?MType=People&MemOrder=2&MemID="+BillMemNum);

			if (myForm.BillMem2.value!=''&&myForm.BillMemID2.value!=''){
				funAutoCodeEnter(myForm.BillMem1);
			}
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
		myForm.Mem.value=myForm.BillMem3.value;
		myForm.MemType.value='P';
		myForm.MemOrder.value=3;
		UrlStr="Query_MemID.asp";
		myForm.kinds.value='DB_Select';
		myForm.action=UrlStr;
		myForm.target="WebPage_Street_People";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.kinds.value='';
	}else{
		if (myForm.BillMem3.value.length > 1){
			var BillMemNum=myForm.BillMem3.value;
			runServerScript("getBillPeoPleMemID.asp?MType=People&MemOrder=3&MemID="+BillMemNum);

			if (myForm.BillMem3.value!=''&&myForm.BillMemID3.value!=''){
				funAutoCodeEnter(myForm.BillMem3);
			}
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
		myForm.Mem.value=myForm.BillMem4.value;
		myForm.MemType.value='P';
		myForm.MemOrder.value=4;
		UrlStr="Query_MemID.asp";
		myForm.kinds.value='DB_Select';
		myForm.action=UrlStr;
		myForm.target="WebPage_Street_People";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.kinds.value='';
	}else{
		if (myForm.BillMem4.value.length > 1){
			var BillMemNum=myForm.BillMem4.value;
			runServerScript("getBillPeoPleMemID.asp?MType=People&MemOrder=4&MemID="+BillMemNum);

			if (myForm.BillMem4.value!=''&&myForm.BillMemID4.value!=''){
				funAutoCodeEnter(myForm.BillMem4);
			}
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
function LawOpen3(){
	UrlStr="Query_Law.asp?LawOrder=3";
	newWin(UrlStr,"WebPage1",550,355,0,0,"yes","no","yes","no");
}
function LawOpen4(){
	UrlStr="Query_Law.asp?LawOrder=4";
	newWin(UrlStr,"WebPage1",550,355,0,0,"yes","no","yes","no");
}
//由違規日期帶入應到案日期
function getDealLineDate(){
	if (!ChkIllegalDate(myForm.IllegalDate.value)){
		alert("違規日期已超過三個月期限，請確認是否正確!!。");
	}
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');
	BFillDateTemp=myForm.IllegalDate.value;
	if (BFillDateTemp.length >= 6){
		myForm.BillFillDate.value=myForm.IllegalDate.value;
		Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
		Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
		Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
		var BFillDate=new Date(Byear,Bmonth-1,Bday)
		var DLineDate=new Date()
		DLineDate=DateAdd("d",15,BFillDate);
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
	if (myForm.PBillSN.value==''){
		myForm.Billno1.value=myForm.Billno1.value.toUpperCase();
		var BillNum=myForm.Billno1.value;
		if (myForm.Billno1.value.length >= 9){
			runServerScript("getPeopleBillNoExist.asp?BillNo="+BillNum);
		}
	}
}
function setCheckPeopleBillNoExist(GetBillFlag,PasserBaseFlag,ChkUnitID,BillSN,MLoginID,MMemberID,MMemName,MUnitID,MUnitName,SUnitID,SUnitName){
	if(ChkUnitID==1){alert("建檔單位非領單單位!!");}
	if (GetBillFlag==0){
		//alert("此單號不存在於領單紀錄中！");
		//document.myForm.Billno1.value="";
	}else{
		//if (document.myForm.BillMem1.value==""){
			document.myForm.BillMemID1.value=MMemberID;
			document.myForm.BillMemName1.value=MMemName;
			TDMemErrorLog1=0;
			<%if sys_City="高雄縣" or sys_City="高雄市" or sys_City="高港局" then%>
				document.myForm.BillMem1.value=MLoginID;
				Layer12.innerHTML=MMemName;
			<%else%>
				document.myForm.BillMem1.value=MMemName;
				Layer12.innerHTML=MLoginID;
			<%end if%>
		//}
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
		document.myForm.Billno1.focus();
	}else if (PasserBaseFlag==0){
		alert('此單號已建檔！');
		document.myForm.Billno1.value="";
		document.myForm.Billno1.focus();
	}
}
function CallChkLaw1(){
}
function CallChkLaw2(){
}
function FuncChkPID(){
	myForm.DriverPID.value=myForm.DriverPID.value.toUpperCase();
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
			runServerScript("DriverIDLoadData.asp?DriverPID="+myForm.DriverPID.value);
		}
	}else if (myForm.DriverPID.value.length > 0 && myForm.DriverPID.value.length < 10){
		alert("身分證輸入錯誤！");
		if (myForm.DriverPID.value.substr(1,1)=="1"){
			document.myForm.DriverSex.value="1";
		}else{
			document.myForm.DriverSex.value="2";
		}
		//myForm.DriverPID.focus();
	}
}

function funAutoCodeEnter(obj){
	<%if sys_City="高雄市" or sys_City="高港局" then%>
	if (event.keyCode>47&&event.keyCode<58){
		CodeEnter(obj.name);
	}else if (event.keyCode>95&&event.keyCode<106){
		CodeEnter(obj.name);
	}else if (event.keyCode>64&&event.keyCode<97){
		CodeEnter(obj.name);
	}
	<%end if%>
}

function funAutoTextControl(obj){
	var objLength=obj.maxLength;

	if(obj.name=='DriverBrith'||obj.name=='IllegalDate'){
		if(obj.value.substr(0,1)>1){
			objLength=6;
			obj.maxLength=6;
		}else{
			objLength=7;
			obj.maxLength=7;
		}
	}
	if(obj.name=='DealLineDate'||obj.name=='BillFillDate'){
		if(obj.value.substr(0,1)>1){
			objLength=6;
			obj.maxLength=6;
		}else{
			objLength=7;
			obj.maxLength=7;
		}
	}
	<%if sys_City="高雄市" or sys_City="高港局" then%>
	if(obj.name=='DriverPID'){
		if(obj.value.length==10){
			CodeEnter(obj.name);
		}
	}

	if (event.keyCode>47&&event.keyCode<58){
		if(obj.value.length==objLength){
			CodeEnter(obj.name);
		}
	}else if (event.keyCode>95&&event.keyCode<106){
		if(obj.value.length==objLength){
			CodeEnter(obj.name);
		}
	}else if (event.keyCode>64&&event.keyCode<97){
		if(obj.value.length==objLength){
			CodeEnter(obj.name);
		}
	}
	<%end if%>
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
	if (obj.name=="IllegalZip"&&event.keyCode==<%=F5str%>){	
		window.open("Query_Zip.asp?ZipCity="+sys_City+"&IllegalZip="+myForm.IllegalZip.value+"&ObjName=IllegalZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}

	if (obj.name=="DriverZip"&&event.keyCode==<%=F5str%>){	
		window.open("Query_Zip.asp?ZipCity=&IllegalZip="+myForm.DriverZip.value+"&ObjName=DriverZip","WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
}

function KeyDown(){
	if (event.keyCode==<%=F5str%>){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}else if (event.keyCode==113){ //F2存檔
		event.keyCode=0;   
		InsertBillVase();
	}else if (event.keyCode==115){ //F4清除
		event.keyCode=0;   
		location='BillKeyIn_People.asp'
	}else if (event.keyCode==<%=F6str%>){ //F6查詢
		event.keyCode=0;   
		funcOpenBillQry();
	}else if (event.keyCode==119){ //F8關閉
		event.keyCode=0;   
		window.close();
	}else if (event.keyCode==121){ //F10查詢未建檔
		event.keyCode=0;
		funPrintCaseList_Stop();
	}else if (event.keyCode==33){ //上一筆PageUp
		event.keyCode=0;   
		funDBupmove();
	}else if (event.keyCode==34){ //下一筆PageDn
		event.keyCode=0;   
		event.returnValue=false; 
		funDBdownmove();
	}/*else if (event.keyCode==36){ //第一筆Home
		event.keyCode=0;   
		funDBfisrt();
	}else if (event.keyCode==35){ //最後一筆End
		event.keyCode=0;   
		event.returnValue=false; 
		funDBlast();
	}*/
}
function funPrintCaseList_Stop(){
	UrlStr="../Query/PrintPeopleDataList.asp?CallType=1";
	newWin(UrlStr,"CaseListWin",300,200,0,0,"yes","yes","yes","no");
}
function funcOpenBillQry(){
	UrlStr="EasyPasserBaseQry.asp";
	newWin(UrlStr,"CaseListWin",250,100,0,0,"yes","yes","yes","no");
}
function funBillNoQuery_Stop(BillNo){
	runServerScript("BillNoDBmove.asp?Sys_Billno="+BillNo);
}
function AutoGetIllStreet(){	//按F5可以直接顯示相關路段
	if (event.keyCode==<%=F5str%>){	
		event.keyCode=0;
		Ostreet=myForm.IllegalAddress.value;
		window.open("Query_Street.asp?OStreet="+Ostreet,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
}
function AutoGetRuleID(LawOrder){	//按F5可以直接顯示相關法條
	if (event.keyCode==<%=F5str%>){	
		event.keyCode=0;
		if (LawOrder==1){
			ORuleID=myForm.Rule1.value;
		}else{
			ORuleID=myForm.Rule2.value;
		}
		window.open("Query_Law.asp?LawOrder="+LawOrder+"&RuleVer=<%=theRuleVer%>&ORuleID="+ORuleID,"WebPage1","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
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
function chkBillType(){
	myForm.BillType.value=myForm.BillType.value.replace(/[^\d]/g,'');
	if (myForm.BillType.value.length=="1"){
		if 	(myForm.BillType.value != "1" && myForm.BillType.value != "2" && myForm.BillType.value != "2" && myForm.BillType.value != "3"){
			alert("告發類別輸入錯誤！");
			myForm.BillType.focus();
			myForm.BillType.value="";
		}
	}
}
function NameLoadDate(){
	if (myForm.DriverName.value.length>2){
		runServerScript("NameLoadData.asp?ChName="+myForm.DriverName.value);
	}
}
function ProjectF5(){
	if (event.keyCode==<%=F5str%>){	
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

function chkDriverSEX();(){
	if (myForm.DriverPID.value.length=10){
		if(myForm.DriverPID.value.substr(2,1))=='1'{
			DriverSEX='男';
		}else{
			DriverSEX='女';
		}
	}
}

myForm.Billno1.focus();
myForm.Billno1.value='<%=theBillno%>';
</script>
</html>
