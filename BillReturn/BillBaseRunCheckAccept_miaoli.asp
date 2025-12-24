<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>逕舉登記簿系統</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 10px; color:#ff0000; }
.btn3{
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
}
-->
</style>
</HEAD>
<BODY>
<%
Server.ScriptTimeout=6000

Function ChkNum(strValue)
	if ISNull(strValue) or trim(strValue)="" or IsEmpty(strValue) then
		ChkNum="null"
	else
		ChkNum=strValue
	end if
End Function

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strCity="select value from Apconfigure where id=3"
set rsCity=conn.execute(strCity)
sys_RuleVer=trim(rsCity("value"))
rsCity.close

if trim(request("DB_Selt"))="PrintOver" then

	strSQL="update BillRunCarAccept set COMPANYACCEPTDATE="&funGetDate(date,0)&",COMPANYMEMBERID=777 where AcceptDate < "&funGetDate(gOutDT(Request("DB_AcceptDate")),0) &" and COMPANYMEMBERID is null and RecordMemberID3 is not null"

	conn.execute(strSQL)

	updstr="RecordMemberID2="&Session("User_ID")&",RecordDate2=sysdate"

	If session("UnitLevelID") = "1" Then updstr=updstr&",RecordMemberID3="&Session("User_ID")&",RecordDate3=sysdate"

	strSQL="Update BillRunCarAccept set "&updstr&" where billunitid='"&trim(Request("DB_BillUnitID"))&"' and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)&" and RecordMemberID2 is null and RecordMemberID1="&trim(request("DB_RecordMemberID1"))
	conn.execute(strSQL)

	Response.write "<script>"
	Response.Write "alert('列印完成！');"
	Response.write "</script>"
	
End If 

if trim(request("DB_Selt"))="SaveCheck" then

	updstr="RecordMemberID3="&Session("User_ID")&",RecordDate3=sysdate"

	If not ifnull(request("DB_RecordMemberID2")) Then
		chkwhere=" and to_char(RecordDate2,'YYYYMMDDHH24')='"&Request("DB_AcceptDate")&"' and RecordMemberID2="&trim(request("DB_RecordMemberID2"))

	else
		chkwhere=" and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)&" and RecordMemberID2 is null and RecordMemberID1="&trim(request("DB_RecordMemberID1"))
	
	End If 

	strSQL="Update BillRunCarAccept set "&updstr&" where billunitid='"&trim(Request("DB_BillUnitID"))&"'"&chkwhere
	conn.execute(strSQL)

	Response.write "<script>"
	Response.Write "alert('簽收送件完成！');"
	Response.write "</script>"
	
End If 

if trim(request("DB_Selt"))="PrintBatOver" then

	updstr="RecordMemberID2="&Session("User_ID")&",RecordDate2=sysdate"

	If session("UnitLevelID") = "3" Then
		chkwhere=" RecordMemberID2 is null and billunitid in('"&trim(Session("Unit_ID"))&"')"	
	else
		chkwhere=" RecordMemberID2 is null and billunitid in(select unitid from unitinfo where unittypeid=(select unittypeid from unitinfo where unitid='"&trim(Session("Unit_ID"))&"'))"
	end if

	strSQL="update BillRunCarAccept set "&updstr&"  where "&chkwhere

	conn.execute(strSQL)
	
	Response.write "<script>"
	Response.Write "alert('設定完成！');"
	Response.write "</script>"

End If  

if trim(request("DB_Selt"))="SaveBat" then

	updstr="RecordMemberID3="&Session("User_ID")&",RecordDate3=sysdate"
	chkwhere="RecordMemberID3"

	strSQL="update BillRunCarAccept set "&updstr&"  where billunitid in(select unitid from unitinfo where unittypeid=(select unittypeid from unitinfo where unitid='"&trim(Session("Unit_ID"))&"')) and "&chkwhere&" is null and RecordMemberID2 is not null"

	conn.execute(strSQL)
	
	Response.write "<script>"
	Response.Write "alert('簽收送件完成！');"
	Response.write "</script>"

End if

if trim(request("DB_Selt"))="Selt" then
	Sys_CarNo=Split(Ucase(trim(request("CarNo"))),",")
	Sys_illegalDate=Split(trim(request("illegalDate")),",")
	Sys_illegalTime=Split(trim(request("illegalTime")),",")
	Sys_CarSimpleID=Split(trim(request("CarSimpleID")),",")
	Sys_CarAddID=Split(trim(request("CarAddID")),",")
	Sys_PeoPleMark=Split(trim(request("PeoPleMark")),",")
	Sys_PeoPleDate=Split(trim(request("PeoPleDate")),",")
	Sys_IllegalAddressID=Split(trim(request("IllegalAddressID")),",")
	Sys_IllegalAddress=Split(trim(request("IllegalAddress")),",")
	Rule1_1=Split(trim(request("Rule1_1")),",")
	Rule1_2=Split(trim(request("Rule1_2")),",")
	Rule1_3=Split(trim(request("Rule1_3")),",")
	Sys_IllegalSpeed=Split(trim(request("IllegalSpeed")),",")
	Sys_RuleSpeed=Split(trim(request("RuleSpeed")),",")
	Sys_BillMemID1=Split(trim(request("BillMemID1")),",")
	Sys_BillMemID2=Split(trim(request("BillMemID2")),",")
	Sys_BillMemID3=Split(trim(request("BillMemID3")),",")
	Sys_BillMemID4=Split(trim(request("BillMemID4")),",")
	Sys_ImageFile=Split(trim(request("ImageFile")),",")
	Sys_PictureFile=Split(trim(request("PictureFile")),",")
	Sys_InformationData=Split(trim(request("InformationData")),",")
	Rule2_1=Split(trim(request("Rule2_1")),",")
	Rule2_2=Split(trim(request("Rule2_2")),",")
	Rule2_3=Split(trim(request("Rule2_3")),",")
	Sys_FastenerTypeID1=Split(trim(request("FastenerTypeID1")),",")
	Sys_FastenerTypeID2=Split(trim(request("FastenerTypeID2")),",")
	Sys_chkBackBillBase=Split(trim(request("Sys_BackBillBase")),",")
	Sys_Note=Split(trim(request("Note")),",")
	Sys_AcceptDate=Trim(Request("AcceptDate"))
	
	Sys_chkBackBillBase=Split(trim(request("Sys_BackBillBase")),",")
	
	old_CarNo=Split(trim(request("old_CarNo")),",")
	old_illegalDate=Split(trim(request("old_illegalDate")),",")
	old_Rule1=Split(trim(request("old_Rule1")),",")

	Sys_Now=DateAdd("n", -5, now)

	DB_RecordMemberID1=trim(request("DB_RecordMemberID1"))

	If ifnull(DB_RecordMemberID1) Then DB_RecordMemberID1=Session("User_ID")
	

	If trim(Request("DB_chkType")) = "" Then
		accwhere=" and AcceptDate="&funGetDate(gOutDT(Request("AcceptDate")),0)&" and RecordMemberID2 is null and RecordMemberID1="&trim(DB_RecordMemberID1)

	elseIf trim(Request("DB_chkType")) <> "0" Then
		accwhere=" and to_char(RecordDate2,'YYYYMMDDHH24')='"&Request("DB_AcceptDate")&"' and RecordMemberID3 is null and RecordMemberID2="&trim(DB_RecordMemberID1)

	End if
	
	For i = 0 to Ubound(Sys_CarNo)
		If (not ifnull(Sys_CarNo(i))) and (not ifnull(Sys_illegalDate(i))) and (not ifnull(Sys_illegalTime(i))) and (not ifnull(Rule1_1(i))) Then

			Sys_Now=DateAdd("s",1,Sys_Now)

			DB_UnitID="":DB_BillMemID1="":DB_BillMemID2="":DB_BillMemID3="":DB_BillMemID4="":DB_illegalDate=""

			SysRule1="":SysRule2="":SysRule3=""

			SysRule1=trim(Rule1_1(i))&trim(Rule1_2(i))&right("000"&trim(Rule1_3(i)),2)&"01"

			If not ifnull(Rule2_1(i)) Then SysRule2=trim(Rule2_1(i))&trim(Rule2_2(i))&right("000"&trim(Rule2_3(i)),2)&"01"

			strSQL="select a.memberid,a.chname,b.UnitID,b.UnitName from (select chname,memberid,unitid from memberdata where Loginid='"&trim(Sys_BillMemID1(i))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo) b where a.unitid=b.unitid"

			set rsmen=conn.execute(strSQL)

			If not rsmen.eof Then
				DB_BillMemID1=trim(rsmen("memberid"))
				DB_UnitID=trim(rsmen("UnitID"))
			end if

			rsmen.close

			strSQL="select a.memberid,a.chname,b.UnitID,b.UnitName from (select chname,memberid,unitid from memberdata where Loginid='"&trim(Sys_BillMemID2(i))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo) b where a.unitid=b.unitid"

			set rsmen=conn.execute(strSQL)

			If not rsmen.eof Then
				DB_BillMemID2=trim(rsmen("memberid"))
			end if

			rsmen.close

			strSQL="select a.memberid,a.chname,b.UnitID,b.UnitName from (select chname,memberid,unitid from memberdata where Loginid='"&trim(Sys_BillMemID3(i))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo) b where a.unitid=b.unitid"

			set rsmen=conn.execute(strSQL)

			If not rsmen.eof Then
				DB_BillMemID3=trim(rsmen("memberid"))
			end if

			rsmen.close	

			strSQL="select a.memberid,a.chname,b.UnitID,b.UnitName from (select chname,memberid,unitid from memberdata where Loginid='"&trim(Sys_BillMemID4(i))&"' and AccountStateID=0 and RecordStateID=0) a,(select UnitID,UnitName from UnitInfo) b where a.unitid=b.unitid"

			set rsmen=conn.execute(strSQL)

			If not rsmen.eof Then
				DB_BillMemID4=trim(rsmen("memberid"))
			end if

			rsmen.close	

			DB_illegalDate=gOutDT(Sys_illegalDate(i))&" "&left(trim(Sys_illegalTime(i)),2)&":"&right(trim(Sys_illegalTime(i)),2)
			
			strWhere="":filedCnt=0
			if not ifnull(old_CarNo(i)) then
				strSQL="select count(1) cmt from BillRunCarAccept where CarNo='"&trim(old_CarNo(i))&"' and IllegalDate="&funGetDate(old_illegalDate(i),1)&" and Rule1='"&trim(old_Rule1(i))&"' and recordstateid=0"
				
				strWhere="CarNo='"&trim(old_CarNo(i))&"' and IllegalDate="&funGetDate(old_illegalDate(i),1)&" and Rule1='"&trim(old_Rule1(i))&"' and recordstateid=0"
			else
				strSQL="select count(1) cmt from BillRunCarAccept where CarNo='"&trim(Sys_CarNo(i))&"' and IllegalDate="&funGetDate(DB_illegalDate,1)&" and Rule1='"&trim(SysRule1)&"' and recordstateid=0"
				
				strWhere="CarNo='"&trim(Sys_CarNo(i))&"' and IllegalDate="&funGetDate(DB_illegalDate,1)&" and Rule1='"&trim(SysRule1)&"' and recordstateid=0"
			end if

			set rsnt=conn.execute(strSQL)

			filedCnt=cdbl(rsnt("cmt"))

			If filedCnt > 0 and ifnull(old_CarNo(i)) Then
				strSQL="delete BillRunCarAccept where "&strWhere

				conn.execute(strSQL)

				filedCnt=0
			End if

			If filedCnt=0 Then

				'strSQL="insert into BillRunCarAccept(CARNO,CARSIMPLEID,BILLUNITID,ILLEGALDATE,ACCEPTDATE,RULE1,RULE2,ILLEGALADDRESS,ILLEGALSPEED,RULESPEED,BILLMEMID1,BILLMEMID2,BILLMEMID3,BILLMEMID4,IMAGEFILE,PICTUREFILE,INFORMATIONDATA,RULEVER,RECORDSTATEID,RECORDMEMBERID1,RECORDDATE) values('"&trim(Sys_CarNo(i))&"',"&ChkNum(Sys_CarSimpleID(i))&",'"&trim(DB_UnitID)&"',"&funGetDate(DB_illegalDate,1)&","&funGetDate(gOutDT(Sys_AcceptDate),0)&",'"&trim(SysRule1)&"','"&trim(SysRule2)&"','"&trim(Sys_IllegalAddress(i))&"',"&ChkNum(Sys_IllegalSpeed(i))&","&ChkNum(Sys_RuleSpeed(i))&","&ChkNum(DB_BillMemID1)&","&ChkNum(DB_BillMemID2)&","&ChkNum(DB_BillMemID3)&","&ChkNum(DB_BillMemID4)&","&ChkNum(Sys_ImageFile(i))&","&ChkNum(Sys_PictureFile(i))&","&ChkNum(Sys_InformationData(i))&",'"&trim(sys_RuleVer)&"',0,"&Session("User_ID")&",sysdate)"

				strSQL="insert into BillRunCarAccept(CARNO,CARSIMPLEID,CARADDID,PEOPLEMARK,BILLUNITID,ILLEGALDATE,ACCEPTDATE,RULE1,RULE2,ILLEGALADDRESSID,ILLEGALADDRESS,ILLEGALSPEED,RULESPEED,BILLMEMID1,BILLMEMID2,BILLMEMID3,BILLMEMID4,IMAGEFILE,PICTUREFILE,INFORMATIONDATA,RULEVER,RECORDSTATEID,RECORDMEMBERID1,RECORDDATE) values('"&trim(Sys_CarNo(i))&"',"&ChkNum(Sys_CarSimpleID(i))&","&ChkNum(Sys_CarAddID(i))&","&ChkNum(Sys_PeoPleMark(i))&",'"&trim(DB_UnitID)&"',"&funGetDate(DB_illegalDate,1)&","&funGetDate(gOutDT(Sys_AcceptDate),0)&",'"&trim(SysRule1)&"','"&trim(SysRule2)&"','"&trim(Sys_IllegalAddressID(i))&"','"&trim(Sys_IllegalAddress(i))&"',"&ChkNum(Sys_IllegalSpeed(i))&","&ChkNum(Sys_RuleSpeed(i))&","&ChkNum(DB_BillMemID1)&","&ChkNum(DB_BillMemID2)&","&ChkNum(DB_BillMemID3)&","&ChkNum(DB_BillMemID4)&","&ChkNum(Sys_ImageFile(i))&","&ChkNum(Sys_PictureFile(i))&","&ChkNum(Sys_InformationData(i))&",'"&trim(sys_RuleVer)&"',0,"&Session("User_ID")&","&funGetDate((Sys_Now),1)&")"

				conn.execute(strSQL)
			else
				'strSQL="Update BillRunCarAccept set CARNO='"&trim(Sys_CarNo(i))&"',CARSIMPLEID="&ChkNum(Sys_CarSimpleID(i))&",BILLUNITID='"&trim(DB_UnitID)&"',ILLEGALDATE="&funGetDate(DB_illegalDate,1)&",ACCEPTDATE="&funGetDate(gOutDT(Sys_AcceptDate),0)&",RULE1='"&trim(SysRule1)&"',RULE2='"&trim(SysRule2)&"',ILLEGALADDRESS='"&trim(Sys_IllegalAddress(i))&"',ILLEGALSPEED="&ChkNum(Sys_IllegalSpeed(i))&",RULESPEED="&ChkNum(Sys_RuleSpeed(i))&",FASTENERTYPEID1='"&trim(Sys_FastenerTypeID1(i))&"',FASTENERTYPEID2='"&trim(Sys_FastenerTypeID2(i))&"',BILLMEMID1="&ChkNum(DB_BillMemID1)&",BILLMEMID2="&ChkNum(DB_BillMemID2)&",BILLMEMID3="&ChkNum(DB_BillMemID3)&",BILLMEMID4="&ChkNum(DB_BillMemID4)&",IMAGEFILE="&ChkNum(Sys_ImageFile(i))&",PICTUREFILE="&ChkNum(Sys_PictureFile(i))&",InformationData="&ChkNum(Sys_InformationData(i))&",RECORDDATE=sysdate where CarNo='"&trim(Sys_CarNo(i))&"' and IllegalDate="&funGetDate(DB_illegalDate,1)&" and Rule1='"&trim(SysRule1)&"' and recordstateid=0"

				strSQL="Update BillRunCarAccept set CARNO='"&trim(Sys_CarNo(i))&"',CARSIMPLEID="&ChkNum(Sys_CarSimpleID(i))&",CarAddID="&ChkNum(Sys_CarAddID(i))&",PEOPLEMARK="&ChkNum(Sys_PeoPleMark(i))&",BILLUNITID='"&trim(DB_UnitID)&"',ILLEGALDATE="&funGetDate(DB_illegalDate,1)&",RULE1='"&trim(SysRule1)&"',RULE2='"&trim(SysRule2)&"',ILLEGALADDRESSID='"&trim(Sys_IllegalAddressID(i))&"',ILLEGALADDRESS='"&trim(Sys_IllegalAddress(i))&"',ILLEGALSPEED="&ChkNum(Sys_IllegalSpeed(i))&",RULESPEED="&ChkNum(Sys_RuleSpeed(i))&",BILLMEMID1="&ChkNum(DB_BillMemID1)&",BILLMEMID2="&ChkNum(DB_BillMemID2)&",BILLMEMID3="&ChkNum(DB_BillMemID3)&",BILLMEMID4="&ChkNum(DB_BillMemID4)&",IMAGEFILE="&ChkNum(Sys_ImageFile(i))&",PICTUREFILE="&ChkNum(Sys_PictureFile(i))&",InformationData="&ChkNum(Sys_InformationData(i))&" where "&strWhere&accwhere

				conn.execute(strSQL)
			End If 
			

			updstr=""
			
			If not ifnull(Sys_chkBackBillBase(i)) Then
				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"recordstateid=-1"
			End if

			If not ifnull(Sys_Note(i)) Then
				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"Note='"&Sys_Note(i)&"'"
			End If 
			
			DB_PeoPleDate=""

			If trim(Sys_PeoPleMark(i))= "1" and (not ifnull(Sys_PeoPleDate(i))) Then
				DB_PeoPleDate=gOutDT(Sys_PeoPleDate(i))

				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"PeoPleDate="&funGetDate(DB_PeoPleDate,0)

			else
				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"PeoPleDate=null"

			End if

			
			If not ifnull(updstr) Then
				strSQL="Update BillRunCarAccept set "&updstr&" where CarNo='"&trim(Sys_CarNo(i))&"' and IllegalDate="&funGetDate(DB_illegalDate,1)&" and Rule1='"&trim(SysRule1)&"' "&accwhere

				conn.execute(strSQL)
			End if

			rsnt.close
		end if
	Next
	Response.write "<script>"
	Response.Write "alert('簽收送件完成！');"
	Response.write "</script>"
end if
%>
<form name="myForm" method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="37" bgcolor="#FFCC33" class="pagetitle">
			<strong>逕舉登記簿系統</strong>
			<a href="./Upaddress/CheckAccept.doc"><font size="3" color="blue"><u>使用說明</u></font> </a> 
			<B><font size="5" color="red">列印完請立即點選『列印完成』。</font></B>
				<%
					Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<a href=""./Upaddress/SystemRunAccept.xls"">"										
					Response.Write "<font size=""3"" color=""blue""><u>逕舉登記簿匯入檔案 下載</u></font></a>"
				%>
				<input type="button" name="btnOK"class="btn3" style="width:200px;height:25px;font-size:15px;" value="匯入逕舉資料(委外廠商使用)" onclick="funChkSelt();">
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						<table border="0">
							<tr>
								<td>
									待審核記錄<span id='BillBaseOrder'></span>
									&nbsp;&nbsp;									
									<input type="button" name="btnExp" class="btn3" style="width:240px;height:25px;font-size:15px;" value="匯入逕舉登記簿檔案(員警使用)" onclick="funChkInput();">
									請於105年8月15日以前的檔案需重新下載
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp;&nbsp
									<%If session("UnitLevelID") <> "3" Then%>
										<input type="button" name="btnOK" class="btn3" style="width:120px;height:25px;font-size:16px;" value="全部審核通過" onclick="funSaveBat();">
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<!--										<input type="button" name="cancel" value="合併列印"  onclick="funAcceptAllLoad();">
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; 
										<input type="button" name="cancel" value="合併列印完成"  onclick="funPrintBatOver();">-->
									<%End if%>
									<input type="button" name="cancel" value="清除" onClick="location='BillBaseRunCheckAccept_miaoli.asp'">
									<br>
									<font size="3" color="red"><B>匯入的檔案內容裡法條必須確實按照條項款格式，例如40.0.0，不可只輸入40或40.0</B></font>
									<br>
									<table width="900" border="1" cellpadding="0" cellspacing="0">
										<tr bgcolor="#EBFBE3" align="center">
											<th width="125">點收日</th>
											<th width="200">送件單位</th>
											<th width="150">建檔人</th>
											<th width="100">件數</th>
											<th width="100">退件數</th>
											<th width="400">點收</th>
										</tr>
										<tr>
											<td colspan="6">
												<Div style="overflow:auto;width:100%;height:200px;background:#FFFFFF">
												<table width="100%" border="0" cellpadding="1" cellspacing="0"><%
												CarCode=",1,2,3,4,5,6,":chkExp=1:chkTime=1

												strSQL="select count(1) chkCnt from BillRunCarAccept where to_Number(to_char(RecordDate2,'YYYYMMDDHH24'))=to_Number(to_char(sysdate,'YYYYMMDDHH24')) and RecordMemberID1="&trim(Session("User_ID"))

												set rstime=conn.execute(strSQL)

												If cdbl(rstime("chkCnt")) > 0 and session("UnitLevelID") = "1" Then chkTime=0

												rstime.close

												If session("UnitLevelID") = "3" Then
													chkwhere="RecordMemberID2 is null and RecordMemberID3 is null"
													showUnit=" and billunitid in('"&trim(Session("Unit_ID"))&"')"	

												elseIf session("UnitLevelID") = "2" Then
													chkwhere="(RecordMemberID2 is not null or RecordMemberID1="&trim(Session("User_ID"))&") and RecordMemberID3 is null"
													showUnit=" and billunitid in(select unitid from unitinfo where unittypeid=(select unittypeid from unitinfo where unitid='"&trim(Session("Unit_ID"))&"'))"

												else
													chkwhere="RecordMemberID3 is null"
													showUnit=" and billunitid in(select unitid from unitinfo where unittypeid=(select unittypeid from unitinfo where unitid='"&trim(Session("Unit_ID"))&"'))"

												end if

												strSQL="select a.AcceptDate,a.BillUnitID,chkType,a.RecordMemberID1,(select chName from MemberData where Memberid=a.RecordMemberID1) ChName,a.suess,a.delss,c.UnitName from (select DeCode(RecordMemberID2,null,to_char(AcceptDate,'YYYY/MM/DD'),to_char(RecordDate2,'YYYYMMDDHH24')) AcceptDate,BillUnitID,decode(RecordMemberID2,null,'0','1') chkType,Decode(RecordMemberID2,null,RecordMemberID1,RecordMemberID2) RecordMemberID1,sum(decode(recordstateid,0,1,0)) suess, sum(decode(recordstateid,-1,1,0)) delss from BillRunCarAccept where "&chkwhere&showUnit&" group by DeCode(RecordMemberID2,null,to_char(AcceptDate,'YYYY/MM/DD'),to_char(RecordDate2,'YYYYMMDDHH24')),BillUnitID,decode(RecordMemberID2,null,'0','1'),Decode(RecordMemberID2,null,RecordMemberID1,RecordMemberID2)) a,UnitInfo c where a.BillUnitID=c.UnitID order by a.AcceptDate,a.BillUnitID"

												set rs=conn.execute(strSQL)
												While not rs.eof
													Response.Write "<tr align=""center"">"

													strAcceptDate=""

													If trim(rs("chkType"))="1" Then
														Response.Write "<td width=""95"">"&(left(rs("AcceptDate"),4)-1911)&Mid(rs("AcceptDate"),5,4)&"</th>"
														strAcceptDate=rs("AcceptDate")

													else
														Response.Write "<td width=""95"">"&gInitDT(rs("AcceptDate"))&"</th>"
														strAcceptDate=gInitDT(rs("AcceptDate"))

													End If  

													If session("UnitLevelID") = "1" and trim(rs("chkType"))="0" Then
														If trim(Session("Ch_Name"))= trim(rs("ChName")) and gInitDT(rs("AcceptDate"))<> gInitDT(date) Then
															chkExp=0
														End if 
													end If 

													Response.Write "<td width=""155"">"&rs("UnitName")&"</th>"
													Response.Write "<td width=""155"">"&rs("ChName")&"</th>"
													Response.Write "<td width=""75"">"&rs("suess")&"</th>"
													Response.Write "<td width=""75"">"&rs("delss")&"</th>"
													Response.Write "<td width=""""><input type=""button"" name=""btnAcc"" class=""btn3"" style=""width:40px;height:25px;font-size:16px;"" value=""詳細"" onclick=""funAcceptLoad('"&strAcceptDate&"','"&rs("BillUnitID")&"','"&rs("UnitName")&"','"&cdbl(rs("suess"))+cdbl(rs("delss"))&"','"&rs("RecordMemberID1")&"','"&rs("chkType")&"');"">"
													

													Response.Write "&nbsp;<input type=""button"" name=""Update"" class=""btn3"" style=""width:40px;height:25px;font-size:16px;"" value=""列印"" class=""btn3"" style=""width:40px;height:25px;font-size:12px;"" onclick=""funAcceptRunList('"&strAcceptDate&"','"&rs("BillUnitID")&"','"&rs("RecordMemberID1")&"','"&rs("chkType")&"');"">"
													
													If session("UnitLevelID") = "2" and trim(rs("chkType"))="1" Then
														Response.Write "&nbsp;&nbsp;<input type=""button"" name=""Update"" class=""btn3"" value=""審核"" class=""btn3"" style=""width:40px;height:25px;font-size:16px;"" onclick=""funSaveCheck('"&strAcceptDate&"','"&rs("BillUnitID")&"','"&rs("RecordMemberID1")&"','"&rs("chkType")&"');"">"

													else
													
														Response.Write "&nbsp;<input type=""button"" name=""Update1"" class=""btn3"" style=""width:80px;height:25px;font-size:16px;"" value=""列印完成"" onclick=""funPrintOver('"&strAcceptDate&"','"&rs("BillUnitID")&"','"&rs("RecordMemberID1")&"');"">"
													
													end if

													Response.Write "</td></tr>"

													rs.movenext
												Wend
												rs.close
												%>
												</table>
												</div>
											</td>
										</tr>
									</table>
								</td>
								<td>
									<%
										If session("UnitLevelID") = "2" Then
											Response.Write "整批審核說明<input type=text name='BatNote' size=32 class='btn1'>"
											Response.Write "<br>"
											Response.Write "<input type=""button"" class=""btn3"" style=""width:120px;height:25px;font-size:16px;"" name=""btn_BatBack"" value=""整批退件(取消)"" onClick=""funBatBack();"">"
										end if

									%>
									<br><br>
									<input type="Hidden" name='AcceptDate' size="5" class='btn1' maxlength='7' value="<%=gInitDT(now)%>">
									
									
								<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
									<br>
									<br><br>
									
								</td>
							</tr>

						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">逕舉登記簿紀錄列表 ( 輸入完成按Enter可自動跳到下一格 )
		<img src="space.gif" width="29" height="8">
		<input type="button" name="insert" class="btn3" style="width:80px;height:25px;font-size:16px;" value="再多50筆" onClick="insertRow(fmyTable)">
		<br>
			<%
			Response.Write "<BR><B>"
			Response.Write "車種代碼：『1汽車、2拖車、3重機/550cc以上、4輕機、5動力機械、6臨時車牌』；"
			
			Response.Write "砂石註記：『1是、0不是』；"

			Response.Write "民眾檢舉：『1是、空白不是』；"

			Response.Write "</B>"
			%>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<Div style="overflow:auto;width:100%;height:330px;background:#FFFFFF">
				<table id='fmyTable' width='100%' border='0' bgcolor='#FFFFFF'>
					<tr bgcolor="#ffffff">
						<td align='center' bgcolor="#ffffff" nowrap></td>
					</tr>
				</table>
			</div>
		</td>
	</tr>
	<tr align="center">
		<td height="35" bgcolor="#FFDD77">
			<input type="button" name="btnOK1" class="btn3" style="width:80px;height:25px;font-size:16px;" value="確定存檔" onclick="funSelt();">
			<input type="button" name="insert2" class="btn3" style="width:80px;height:25px;font-size:16px;" value="再多50筆" onClick="insertRow(fmyTable)">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="">
<input type="Hidden" name="chkcnt" value="">
<input type="Hidden" name="DB_AcceptDate" value="">
<input type="Hidden" name="DB_BillUnitID" value="">
<input type="Hidden" name="DB_RecordMemberID1" value="">
<input type="Hidden" name="DB_RecordMemberID2" value="">
<input type="Hidden" name="DB_chkType" value="">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var cunt=0;
var chkExp=<%=chkExp%>;
var chkTime=<%=chkTime%>;
if(!chkExp){
	alert("請先處理之前案件!!");
	myForm.btnExp.disabled=true;
	myForm.btnOK1.disabled=true;
}

if(!chkTime){
	alert("系統處理中，需等待1小時，請先關閉此功能!!");
	myForm.btnExp.disabled=true;
	myForm.btnOK1.disabled=true;
}

function insertRow(isTable){
	for(i=0;i<=49;i++){
		Rindex = isTable.rows.length;
		if(isTable.rows.length>0){
		    Cindex = isTable.rows[Rindex-1].cells.length;
		}else{
		    Cindex=0;
		}
		if(Rindex==0||Cindex==1){
		    nextRow = isTable.insertRow(Rindex);
		    txtArea = nextRow.insertCell(0);
		}else{
		    if(cunt==0){
		        Cindex=0;
		        isTable.rows[Rindex-1].deleteCell();
		    }
		    txtArea =isTable.rows[Rindex-1].insertCell(Cindex);
		}
		cunt++;
		//txt_nameStr = "item"+cunt;
		var cnt_num=("0000"+cunt).substr(("0000"+cunt).length-3,3);

		if(cnt_num%2==0){txtArea.style.backgroundColor ="#EEEEEE";}

		txtArea.innerHTML =
		"<b>" + cnt_num + "</b>&nbsp;&nbsp;"+
		"<span style='color:#ff0000;'>*</span>" +
		"車號&nbsp;&nbsp;<input type=text name='CarNo' size=5 class='btn1' onkeyup='UpperCase(this);' onkeydown='keyCarNo("+cunt+");'>" +
		"&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"違規日期<input type=text style='ime-mode:disabled;' name='illegalDate' size=5 class='btn1' onkeyup='chknumber(this);' onkeydown='KeyillegalDate("+cunt+");' maxlength='7'>" +
		"&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"時間<input type=text style='ime-mode:disabled;' name='illegalTime' size=1 class='btn1' onkeyup='chknumber(this);' onkeydown='KeyillegalTime("+cunt+");' maxlength='4'>" +
		"&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"車種<input type=text style='ime-mode:disabled;' name='CarSimpleID' size=1 class='btn1' onkeyup='chknumber(this);' onkeydown='KeyCarSimpleID("+cunt+");' maxlength='1'>" +
		"&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"砂石註記<input type=text style='ime-mode:disabled;' name='CarAddID' size=1 class='btn1' value='0' maxlength='3' onkeyup='chknumber(this);' onkeydown='KeyCarAddID("+cunt+");' maxlength='1'>" +
		"&nbsp;&nbsp;" +
		"民眾檢舉<input type=text style='ime-mode:disabled;' name='PeoPleMark' size=1 class='btn1' value='' maxlength='3' onkeyup='chknumber(this);' onkeydown='KeyPeoPleMark("+cunt+");' maxlength='1'>" +

		"民眾檢舉日期<input type=text style='ime-mode:disabled;' name='PeoPleDate' size=5 class='btn1' onkeyup='chknumber(this);' onkeydown='KeyPeoPleDate("+cunt+");' maxlength='7'>" +

		"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
		"固定桿代碼<input type=text style='ime-mode:disabled;' name='IllegalAddressID' size=2 class='btn1' onkeydown='KeyIllegalAddressID("+cunt+");' onkeyup='getIllegalData("+cunt+");'>"+
		"<span style='color:#ff0000;'>*</span>" +
		"違規地點<input type=text name='IllegalAddress' size=30 class='btn1' onkeydown='KeyIllegalAddress("+cunt+");'>" +
		
		"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + 
		"<span style='color:#ff0000;'>*</span>" +
		"法條1&nbsp;<input type=text style='ime-mode:disabled;' name='Rule1_1' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule1_1("+cunt+");'>條"  +		
		"<input type=text style='ime-mode:disabled;' name='Rule1_2' size=1 maxlength='1' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule1_2("+cunt+");'>項"  +		
		"<input type=text style='ime-mode:disabled;' name='Rule1_3' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule1_3("+cunt+");'>款"  +		
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
		"限速<input type=text style='ime-mode:disabled;' name='RuleSpeed' maxlength='3' size=1 class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRuleSpeed("+cunt+");'>" +
		"&nbsp;&nbsp;" +
		"車速<input type=text style='ime-mode:disabled;' name='IllegalSpeed' maxlength='3'  size=1 class='btn1' onkeyup='chknumber(this);' onkeydown='KeyIllegalSpeed("+cunt+");'>" +

		"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"員警1&nbsp;<input type=text style='ime-mode:disabled;' name='BillMemID1' size=2 class='btn1' onkeydown='KeyBillMemID1("+cunt+");' onkeyup='getBillMem1("+cunt+");'>" +
		"<span class='style1' id='BillMemName1'></span>" +
		"&nbsp;&nbsp;" +
		"員警2<input type=text style='ime-mode:disabled;' name='BillMemID2' size=2 class='btn1' onkeydown='KeyBillMemID2("+cunt+");' onkeyup='getBillMem2("+cunt+");'>" +
		"<span class='style1' id='BillMemName2'></span>" +
		"&nbsp;&nbsp;" +
		"員警3<input type=text style='ime-mode:disabled;' name='BillMemID3' size=2 class='btn1' onkeydown='KeyBillMemID3("+cunt+");' onkeyup='getBillMem3("+cunt+");'>" +
		"<span class='style1' id='BillMemName3'></span>" +
		"&nbsp;&nbsp;" +
		"員警4<input type=text style='ime-mode:disabled;' name='BillMemID4' size=2 class='btn1' onkeydown='KeyBillMemID4("+cunt+");' onkeyup='getBillMem4("+cunt+");'>" +
		"<span class='style1' id='BillMemName4'></span>" +
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
		"圖<input type=text style='ime-mode:disabled;' name='ImageFile' size=1 class='btn1' onkeyup='chknumber(this);' onkeydown='KeyImageFile("+cunt+");' maxlength='1'>" +
		"&nbsp;&nbsp;" +
		"相<input type=text style='ime-mode:disabled;' name='PictureFile' size=1 class='btn1' value='1' onkeyup='chknumber(this);' onkeydown='KeyPictureFile("+cunt+");' maxlength='1'>" +
		"&nbsp;&nbsp;" +
		"資<input type=text style='ime-mode:disabled;' name='InformationData' size=1 class='btn1' onkeyup='chknumber(this);' onkeydown='KeyInformationData("+cunt+");' maxlength='1'>" +
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
		"法條2<input type=text style='ime-mode:disabled;' name='Rule2_1' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule2_1("+cunt+");'>條" +
		"<input type=text style='ime-mode:disabled;' name='Rule2_2' size=1 maxlength='1' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule2_2("+cunt+");'>項" +
		"<input type=text style='ime-mode:disabled;' name='Rule2_3' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule2_3("+cunt+");'>款" +
		"<input type='hidden' name='old_CarNo' value=''><input type='hidden' name='old_illegalDate' value=''><input type='hidden' name='old_Rule1' value=''>"+
		"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" + <%
		If session("UnitLevelID") <> "3" Then%>
			"審核說明<input type=text name='Note' size=32 class='btn1' onkeydown='KeyNote("+cunt+");'>" +
			"退件<input class='btn1' type='checkbox' name='chkBackBillBase' value='-1' onclick='funChkBackBillBase("+cunt+");'>" +
			"<input type='hidden' name='Sys_BackBillBase'><hr>";<%
		else%>
			"<input type='hidden' name='Sys_BackBillBase'><input type='hidden' name='Note' size=32 class='btn1'><hr>";<%
		end if%>
	}
}
function funChkSelt(){
	UrlStr="BillBaseCheckRunAcceptSendStyle_miaoli.asp";
	myForm.action=UrlStr;
	myForm.target="ChkSelt";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funChkInput(){
	UrlStr="BillBaseSystemAcceptSendStyle_miaoli.asp";
	myForm.action=UrlStr;
	myForm.target="ChkSelt";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funAcceptRunList(AcceptDate,BillUnitID,RecordMemberID1,chkType){
	var UnitLevelID='<%=session("UnitLevelID")%>';

	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordMemberID1.value="";
	myForm.DB_RecordMemberID2.value="";

	if(chkType=='0'){
		myForm.DB_RecordMemberID1.value=RecordMemberID1;
	}else{
		myForm.DB_RecordMemberID2.value=RecordMemberID1;
	}

	if(UnitLevelID=='1'||UnitLevelID=='3'){
		alert("請立即點選『列印完成』!!");
	}

	UrlStr="AcceptRunList.asp";
	
	myForm.action=UrlStr;
	myForm.target="PrintAccept";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funAcceptAllLoad(){
	myForm.DB_AcceptDate.value="";
	myForm.DB_BillUnitID.value="";
	myForm.DB_RecordMemberID1.value="";

	UrlStr="AcceptRunList.asp";
	
	myForm.action=UrlStr;
	myForm.target="PrintAccept";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPrintBatOver(){
	myForm.DB_AcceptDate.value="";
	myForm.DB_BillUnitID.value="";
	myForm.DB_RecordMemberID1.value="";
	myForm.DB_RecordMemberID2.value="";

	myForm.DB_Selt.value="PrintBatOver";
	myForm.submit();
}

function funPrintOver(AcceptDate,BillUnitID,RecordMemberID1){
	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordMemberID1.value=RecordMemberID1;

	myForm.DB_Selt.value="PrintOver";
	myForm.submit();
}

function funSaveCheck(AcceptDate,BillUnitID,RecordMemberID1,chkType){
	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordMemberID1.value="";
	myForm.DB_RecordMemberID2.value="";

	if(chkType=='0'){
		myForm.DB_RecordMemberID1.value=RecordMemberID1;
	}else{
		myForm.DB_RecordMemberID2.value=RecordMemberID1;
	}

	myForm.DB_Selt.value="SaveCheck";
	myForm.submit();
}

function funSaveBat(){

	myForm.DB_Selt.value="SaveBat";
	myForm.submit();
}

function funAcceptLoad(AcceptDate,UnitID,UnitName,Cmt,RecordMemberID1,chkType){
	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_RecordMemberID1.value=RecordMemberID1;
	myForm.DB_chkType.value=chkType;

	BillBaseOrder.innerHTML="<font size=3 color='Red'>『目前查閱"+AcceptDate+":"+UnitName+":"+Cmt+"件』</font>";

	runServerScript("getRunCarAcceptData_miaoli.asp?AcceptDate="+AcceptDate+"&UnitID="+UnitID+"&objcnt="+myForm.CarNo.length+"&RecordMemberID1="+RecordMemberID1+"&chkType="+chkType);

	for(i=Cmt;i<myForm.CarNo.length;i++){

		myForm.CarNo[i].value='';

		myForm.illegalDate[i].value='';

		myForm.illegalTime[i].value='';

		myForm.CarSimpleID[i].value='';
		
		myForm.CarAddID[i].value='';

		myForm.PeoPleMark[i].value='';

		myForm.PeoPleDate[i].value='';

		myForm.IllegalAddressID[i].value='';

		myForm.IllegalAddress[i].value='';

		myForm.Rule1_1[i].value='';

		myForm.Rule1_2[i].value='';

		myForm.Rule1_3[i].value='';

		myForm.IllegalSpeed[i].value='';

		myForm.RuleSpeed[i].value='';

		myForm.BillMemID1[i].value='';

		BillMemName1[i].innerHTML='';

		myForm.BillMemID2[i].value='';

		BillMemName2[i].innerHTML='';

		myForm.BillMemID3[i].value='';

		BillMemName3[i].innerHTML='';

		myForm.BillMemID4[i].value='';

		BillMemName4[i].innerHTML='';

		myForm.ImageFile[i].value='';

		myForm.PictureFile[i].value='';

		myForm.InformationData[i].value='';

		myForm.Rule2_1[i].value='';

		myForm.Rule2_2[i].value='';

		myForm.Rule2_3[i].value='';
		
		myForm.Sys_BackBillBase[i].value='';
		
		myForm.chkBackBillBase[i].checked=false;

		myForm.Note[i].value='';
		
		myForm.old_CarNo[i].value='';

		myForm.old_illegalDate[i].value='';

		myForm.old_Rule1[i].value='';

	}
}
function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}

function keyCarNo(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		if (myForm.CarNo[itemcnt-1].value.length>9){
			alert("車號不可超過9碼!!");
		}else{
			if(myForm.IllegalAddress[itemcnt-2]){
				myForm.illegalDate[itemcnt-1].value=myForm.illegalDate[itemcnt-2].value;

				myForm.IllegalAddressID[itemcnt-1].value=myForm.IllegalAddressID[itemcnt-2].value;

				myForm.IllegalAddress[itemcnt-1].value=myForm.IllegalAddress[itemcnt-2].value;

				myForm.RuleSpeed[itemcnt-1].value=myForm.RuleSpeed[itemcnt-2].value;

				BillMemName1[itemcnt-1].innerHTML=BillMemName1[itemcnt-2].innerHTML;
				myForm.BillMemID1[itemcnt-1].value=myForm.BillMemID1[itemcnt-2].value;

				BillMemName2[itemcnt-1].innerHTML=BillMemName2[itemcnt-2].innerHTML;
				myForm.BillMemID2[itemcnt-1].value=myForm.BillMemID2[itemcnt-2].value;

				BillMemName3[itemcnt-1].innerHTML=BillMemName3[itemcnt-2].innerHTML;
				myForm.BillMemID3[itemcnt-1].value=myForm.BillMemID3[itemcnt-2].value;

				BillMemName4[itemcnt-1].innerHTML=BillMemName4[itemcnt-2].innerHTML;
				myForm.BillMemID4[itemcnt-1].value=myForm.BillMemID4[itemcnt-2].value;

				myForm.Rule1_1[itemcnt-1].value=myForm.Rule1_1[itemcnt-2].value;
				myForm.Rule1_2[itemcnt-1].value=myForm.Rule1_2[itemcnt-2].value;
				myForm.Rule1_3[itemcnt-1].value=myForm.Rule1_3[itemcnt-2].value;

				myForm.PeoPleMark[itemcnt-1].value=myForm.PeoPleMark[itemcnt-2].value;
				myForm.PeoPleDate[itemcnt-1].value=myForm.PeoPleDate[itemcnt-2].value;
			}
			myForm.illegalDate[itemcnt-1].focus();
		}			
	}
}
function KeyillegalDate(itemcnt){

	if (event.keyCode==13||event.keyCode==9){
		myForm.illegalTime[itemcnt-1].focus();
	}
}

function KeyillegalTime(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.CarSimpleID[itemcnt-1].focus();
	}
}

function KeyCarSimpleID(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.CarAddID[itemcnt-1].focus();
	}
}

function KeyCarAddID(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.PeoPleMark[itemcnt-1].focus();
	}
}

function KeyPeoPleMark(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.PeoPleDate[itemcnt-1].focus();
	}
}

function KeyPeoPleDate(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.IllegalAddressID[itemcnt-1].focus();
	}
}

function KeyIllegalAddressID(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.IllegalAddress[itemcnt-1].focus();
	}
}

function getIllegalData(itemcnt){
	if(myForm.IllegalAddressID[itemcnt-1].value!=''){
		UpperCase(myForm.IllegalAddressID[itemcnt-1]);
		myForm.IllegalAddress[itemcnt-1].value="";
		myForm.RuleSpeed[itemcnt-1].value="";
		runServerScript("getIllegalData_miaoli.asp?VL_PLCE_CD="+myForm.IllegalAddressID[itemcnt-1].value+"&itemcnt="+(itemcnt-1));
	}
}

function KeyIllegalAddress(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Rule1_1[itemcnt-1].focus();
	}
}


function KeyRule1_1(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Rule1_2[itemcnt-1].focus();
	}
}
function KeyRule1_2(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Rule1_3[itemcnt-1].focus();
	}
}
function KeyRule1_3(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.RuleSpeed[itemcnt-1].focus();
	}
}

function KeyIllegalSpeed(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.BillMemID1[itemcnt-1].focus();
	}
}

function KeyRuleSpeed(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.IllegalSpeed[itemcnt-1].focus();
	}
}

function KeyBillMemID1(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.ImageFile[itemcnt-1].focus();
	}
}

function getBillMem1(itemcnt){
	UpperCase(myForm.BillMemID1[itemcnt-1]);
	BillMemName1[itemcnt-1].innerHTML="";
	runServerScript("CheckStopCarAcceptMemID_miaoli.asp?LoginID="+myForm.BillMemID1[itemcnt-1].value+"&innerObj=BillMemName1&itemcnt="+(itemcnt-1));
}

function KeyBillMemID2(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.ImageFile[itemcnt-1].focus();
	}
}

function getBillMem2(itemcnt){
	UpperCase(myForm.BillMemID2[itemcnt-1]);
	BillMemName2[itemcnt-1].innerHTML="";
	runServerScript("CheckStopCarAcceptMemID_miaoli.asp?LoginID="+myForm.BillMemID2[itemcnt-1].value+"&innerObj=BillMemName2&itemcnt="+(itemcnt-1));
}

function KeyBillMemID3(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.ImageFile[itemcnt-1].focus();
	}
}

function getBillMem3(itemcnt){
	UpperCase(myForm.BillMemID3[itemcnt-1]);
	BillMemName3[itemcnt-1].innerHTML="";
	runServerScript("CheckStopCarAcceptMemID_miaoli.asp?LoginID="+myForm.BillMemID3[itemcnt-1].value+"&innerObj=BillMemName3&itemcnt="+(itemcnt-1));
}

function KeyBillMemID4(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.ImageFile[itemcnt-1].focus();
	}
}

function getBillMem4(itemcnt){
	UpperCase(myForm.BillMemID4[itemcnt-1]);
	BillMemName4[itemcnt-1].innerHTML="";
	runServerScript("CheckStopCarAcceptMemID_miaoli.asp?LoginID="+myForm.BillMemID4[itemcnt-1].value+"&innerObj=BillMemName4&itemcnt="+(itemcnt-1));
}

function KeyImageFile(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.PictureFile[itemcnt-1].focus();
	}
}

function KeyPictureFile(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.InformationData[itemcnt-1].focus();
	}
}

function KeyInformationData(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.CarNo[itemcnt].focus();
	}
}

function KeyFastenerTypeID1(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.CarNo[itemcnt].focus();
	}
}

function KeyFastenerTypeID2(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.CarNo[itemcnt].focus();
	}
}

function KeyRule2_1(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Rule2_2[itemcnt-1].focus();
	}
}
function KeyRule2_2(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Rule2_3[itemcnt-1].focus();
	}
}
function KeyRule2_3(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.CarNo[itemcnt].focus();
	}
}

function KeyNote(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.CarNo[itemcnt].focus();
	}
}

function funChkBackBillBase(itemcnt){
	if(myForm.chkBackBillBase[itemcnt-1].checked){
		myForm.Note[itemcnt-1].disabled=false;
		myForm.Sys_BackBillBase[itemcnt-1].value="-1";
	}else{
		myForm.Note[itemcnt-1].disabled=true;
		myForm.Sys_BackBillBase[itemcnt-1].value="";
	}
}


function DeleteRow(isTable){
	if(isTable.rows.length>0){
		Rindex = isTable.rows.length;
		Cindex = isTable.rows(Rindex-1).cells.length;
		if(Cindex==1){
			cunt--;
			isTable.rows(Rindex-1).deleteCell();
			isTable.deleteRow();
		}else{
			cunt--;
			isTable.rows(Rindex-1).deleteCell();
		}
	}
}

function funBatBack(){
	for(i=0;i<myForm.CarNo.length;i++){
		if(myForm.illegalDate[i].value!=''){
			myForm.chkBackBillBase[i].click();
			funChkBackBillBase(i+1);

			if(myForm.chkBackBillBase[i].checked){
				myForm.Note[i].value=myForm.BatNote.value;
			}else{
				myForm.Note[i].value="";
			}
		}
	}
}

function funSelt(){
	var err=0;
	var errmsg="";
	var chkDate=<%=gInitDT(now)%>;
	
	if(myForm.AcceptDate.value==''){
		err=1;
		errmsg=errmsg+"點收日期不可空白!!\n";

	}

	for(i=0;i<myForm.CarNo.length;i++){
		if(myForm.CarNo[i].value!=''){
			if(myForm.illegalDate[i].value==''||myForm.illegalTime[i].value==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規日期時間不可空白!!\n";
			}

			if(myForm.illegalDate[i].value > chkDate ){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規日期不可的超過今天!!\n";
			}

			if(myForm.illegalDate[i].value.length!=7){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規日期錯誤!!\n";
			}

			if(myForm.illegalTime[i].value > 2359){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規時間不正確!!\n";
			}

			if(myForm.illegalTime[i].value.length!=4){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規時間錯誤!!\n";
			}

			if(myForm.CarSimpleID[i].value==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行車種不可空白!!\n";
			}

			if(myForm.CarSimpleID[i].value!=''&& "<%=CarCode%>".indexOf(myForm.CarSimpleID[i].value,0)<0){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行車種錯誤!!\n";
			}
			
			if(myForm.CarAddID[i].value=='' || (myForm.CarAddID[i].value!=1 && myForm.CarAddID[i].value!=0)){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行砂石錯誤!!\n";
			}

			if(myForm.PeoPleMark[i].value!='' && myForm.PeoPleMark[i].value!=1){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行民眾檢舉註記錯誤!!\n";
			}

			if(myForm.PeoPleMark[i].value==1 && myForm.PeoPleDate[i].value.length!=7){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行檢舉時間錯誤!!\n";
			}

			if(myForm.IllegalAddress[i].value==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規地點不可空白!!\n";
			}

			if(myForm.Rule1_1[i].value==''||myForm.Rule1_2[i].value==''||myForm.Rule1_3[i].value==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行法條1不可空白!!\n";
			}

			if(myForm.Rule2_1[i].value!=''||myForm.Rule2_2[i].value!=''||myForm.Rule2_3[i].value!=''){
				if(myForm.Rule2_1[i].value==''||myForm.Rule2_2[i].value==''||myForm.Rule2_3[i].value==''){
					err=1;
					errmsg=errmsg+"第 "+(i+1)+" 行法條2不可空白!!\n";
				}
			}

			if(myForm.BillMemID1[i].value==''||BillMemName1[i].innerHTML==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行舉發員警1錯誤!!\n";
			}

			if(myForm.BillMemID2[i].value!=''&&BillMemName2[i].innerHTML==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行舉發員警2錯誤!!\n";
			}

			if(myForm.BillMemID3[i].value!=''&&BillMemName3[i].innerHTML==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行舉發員警3錯誤!!\n";
			}

			if(myForm.BillMemID4[i].value!=''&&BillMemName4[i].innerHTML==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行舉發員警4錯誤!!\n";
			}

			if(err!=0){
				break;
			}else{
				myForm.Note[i].disabled=false;
			}
		}
	}

	if(err==0){
		myForm.DB_Selt.value="Selt";
		myForm.submit();
	}else{
		alert(errmsg);
	}
}

function UpperCase(obj){
	if(obj.value!=obj.value.toUpperCase()){
		obj.value=obj.value.toUpperCase();
	}
}


insertRow(fmyTable);

</script>