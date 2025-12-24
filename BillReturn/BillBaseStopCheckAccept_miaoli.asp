<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>攔停登記簿系統</TITLE>
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

Server.ScriptTimeout=6000
strCity="select value from Apconfigure where id=3"
set rsCity=conn.execute(strCity)
sys_RuleVer=trim(rsCity("value"))
rsCity.close

if trim(request("DB_Selt"))="PrintOver" then

	strSQL="update BillStopCarAccept set COMPANYACCEPTDATE="&funGetDate(date,0)&",COMPANYMEMBERID=777 where AcceptDate < "&funGetDate(gOutDT(Request("DB_AcceptDate")),0) &" and COMPANYMEMBERID is null and RecordMemberID3 is not null"

	conn.execute(strSQL)

	updstr="RecordMemberID2="&Session("User_ID")&",RecordDate2=sysdate"

	If session("UnitLevelID") = "1" Then updstr=updstr&",RecordMemberID3="&Session("User_ID")&",RecordDate3=sysdate"

	strSQL="Update BillStopCarAccept set "&updstr&" where billunitid='"&trim(Request("DB_BillUnitID"))&"' and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)&" and RecordMemberID2 is null and RecordMemberID1="&trim(request("DB_RecordMemberID1"))
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

	strSQL="Update BillStopCarAccept set "&updstr&" where billunitid='"&trim(Request("DB_BillUnitID"))&"'"&chkwhere

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

	strSQL="update BillStopCarAccept set "&updstr&"  where "&chkwhere

	conn.execute(strSQL)
	
	Response.write "<script>"
	Response.Write "alert('設定完成！');"
	Response.write "</script>"

End if

if trim(request("DB_Selt"))="SaveBat" then

	updstr="RecordMemberID3="&Session("User_ID")&",RecordDate3=sysdate"
	chkwhere="RecordMemberID3"

	strSQL="update BillStopCarAccept set "&updstr&"  where billunitid in(select unitid from unitinfo where unittypeid=(select unittypeid from unitinfo where unitid='"&trim(Session("Unit_ID"))&"')) and "&chkwhere&" is null and RecordMemberID2 is not null"

	conn.execute(strSQL)
	
	Response.write "<script>"
	Response.Write "alert('簽收送件完成！');"
	Response.write "</script>"

End if

if trim(request("DB_Selt"))="Selt" then

	Sys_BillNo=Split(Ucase(trim(request("item"))),",")
	Sys_CarNo=Split(Ucase(trim(request("CarNo"))),",")
	Sys_CarSimpleID=Split(trim(request("CarSimpleID")),",")
	Sys_CarAddID=Split(trim(request("CarAddID")),",")
 	Sys_illegalDate=Split(trim(request("illegalDate")),",")
	Sys_illegalTime=Split(trim(request("illegalTime")),",")
	Sys_IllegalAddress=Split(trim(request("IllegalAddress")),",")
	Rule1_1=Split(trim(request("Rule1_1")),",")
	Rule1_2=Split(trim(request("Rule1_2")),",")
	Rule1_3=Split(trim(request("Rule1_3")),",")
	Rule1_4=Split(trim(request("Rule1_4")),",")
	Sys_Fastener1=Split(trim(request("FastenerTypeID1")),",")
	Sys_BillMemID1=Split(trim(request("BillMemID1")),",")
	Sys_BillMemID2=Split(trim(request("BillMemID2")),",")
	Sys_BillMemID3=Split(trim(request("BillMemID3")),",")
	Sys_BillMemID4=Split(trim(request("BillMemID4")),",")
	Sys_BillStatus=Split(trim(request("BillStatus")),",")
	Rule2_1=Split(trim(request("Rule2_1")),",")
	Rule2_2=Split(trim(request("Rule2_2")),",")
	Rule2_3=Split(trim(request("Rule2_3")),",")
	Rule2_4=Split(trim(request("Rule2_4")),",")
	Sys_Fastener2=Split(trim(request("FastenerTypeID2")),",")
	Sys_Rule3=Split(trim(request("Rule3")),",")
	Rule3_1=Split(trim(request("Rule3_1")),",")
	Rule3_2=Split(trim(request("Rule3_2")),",")
	Rule3_3=Split(trim(request("Rule3_3")),",")
	Rule3_4=Split(trim(request("Rule3_4")),",")
	Sys_chkBackBillBase=Split(trim(request("Sys_BackBillBase")),",")
	'Sys_DriverName=Split(trim(request("DriverName")),",")

	Sys_Note=Split(trim(request("Note")),",")

	Sys_AcceptDate=Trim(Request("AcceptDate"))

	old_BillNo=Split(trim(request("old_BillNo")),",")

	Sys_Now=now

	DB_RecordMemberID1=trim(request("DB_RecordMemberID1"))

	If ifnull(DB_RecordMemberID1) Then DB_RecordMemberID1=Session("User_ID")

	If trim(Request("DB_chkType")) = "" Then
		accwhere=" and AcceptDate="&funGetDate(gOutDT(Request("AcceptDate")),0)&" and RecordMemberID2 is null and RecordMemberID1="&trim(DB_RecordMemberID1)

	elseIf trim(Request("DB_chkType")) <> "0" Then
		accwhere=" and to_char(RecordDate2,'YYYYMMDDHH24')='"&Request("DB_AcceptDate")&"' and RecordMemberID3 is null and RecordMemberID2="&trim(DB_RecordMemberID1)

	End if
	
	For i = 0 to Ubound(Sys_BillNo)
		If not ifnull(Sys_illegalDate(i)) Then
			DB_UnitID="":DB_BillMemID1="":DB_BillMemID2="":DB_BillMemID3="":DB_BillMemID4="":DB_illegalDate=""

			Sys_Now=DateAdd("s",1,Sys_Now)

			SysRule1="":SysRule2="":SysRule3=""

			SysRule1=trim(Rule1_1(i))&trim(Rule1_2(i))&right("000"&trim(Rule1_3(i)),2)&"01"&trim(Rule1_4(i))

			If not ifnull(Rule2_1(i)) Then SysRule2=trim(Rule2_1(i))&trim(Rule2_2(i))&right("000"&trim(Rule2_3(i)),2)&"01"&trim(Rule2_4(i))

			If not ifnull(Rule3_1(i)) Then SysRule3=trim(Rule3_1(i))&trim(Rule3_2(i))&right("000"&trim(Rule3_3(i)),2)&"01"&trim(Rule3_4(i))

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

			If not ifnull(old_BillNo(i)) Then
				strSQL="select count(1) cmt from BillStopCarAccept where BillNo='"&trim(old_BillNo(i))&"' and recordstateid=0"

				strWhere="BillNo='"&trim(old_BillNo(i))&"' and recordstateid=0"
			else
				strSQL="select count(1) cmt from BillStopCarAccept where BillNo='"&trim(Sys_BillNo(i))&"' and recordstateid=0"

				strWhere="BillNo='"&trim(Sys_BillNo(i))&"' and recordstateid=0"
			End If 
			
			set rsnt=conn.execute(strSQL)

			filedCnt=cdbl(rsnt("cmt"))

			If filedCnt > 0 and ifnull(old_BillNo(i)) Then
				strSQL="delete BillStopCarAccept where "&strWhere

				conn.execute(strSQL)

				filedCnt=0
			End if 

			If filedCnt=0 Then
				strSQL="insert into BillStopCarAccept(Billno,Carno,IllegalDate,CarSimpleID,CarAddID,IllegalAddress,Rule1,Rule2,Rule3,BillUnitID,BillMemID1,BillMemID2,BillMemID3,BillMemID4,FastenerTypeID1,FastenerTypeID2,BillStatus,Acceptdate,RecordStateID,RecordMemberID1,RecordDate) values('"&trim(Sys_BillNo(i))&"','"&trim(Sys_CarNo(i))&"',"&funGetDate(DB_illegalDate,1)&","&ChkNum(Sys_CarSimpleID(i))&","&ChkNum(Sys_CarAddID(i))&",'"&trim(Sys_IllegalAddress(i))&"','"&trim(SysRule1)&"','"&trim(SysRule2)&"','"&trim(SysRule3)&"','"&DB_UnitID&"',"&ChkNum(DB_BillMemID1)&","&ChkNum(DB_BillMemID2)&","&ChkNum(DB_BillMemID3)&","&ChkNum(DB_BillMemID4)&",'"&trim(Sys_Fastener1(i))&"','"&trim(Sys_Fastener2(i))&"','"&trim(Sys_BillStatus(i))&"',"&funGetDate(gOutDT(Sys_AcceptDate),0)&",0,"&Session("User_ID")&","&funGetDate((Sys_Now),1)&")"

				conn.execute(strSQL)
			else
				strSQL="Update BillStopCarAccept set Billno='"&trim(Sys_BillNo(i))&"',Carno='"&trim(Sys_CarNo(i))&"',IllegalDate="&funGetDate(DB_illegalDate,1)&",CarSimpleID="&ChkNum(Sys_CarSimpleID(i))&",CarAddID="&ChkNum(Sys_CarAddID(i))&",IllegalAddress='"&trim(Sys_IllegalAddress(i))&"',Rule1='"&trim(SysRule1)&"',Rule2='"&trim(SysRule2)&"',Rule3='"&trim(SysRule3)&"',BillUnitID='"&DB_UnitID&"',BillMemID1="&ChkNum(DB_BillMemID1)&",BillMemID2="&ChkNum(DB_BillMemID2)&",BillMemID3="&ChkNum(DB_BillMemID3)&",BillMemID4="&ChkNum(DB_BillMemID4)&",FastenerTypeID1='"&trim(Sys_Fastener1(i))&"',FastenerTypeID2='"&trim(Sys_Fastener2(i))&"',BillStatus='"&trim(Sys_BillStatus(i))&"' where "&strWhere&accwhere

				conn.execute(strSQL)
			End if
			updstr=""
			
			If not ifnull(Sys_chkBackBillBase(i)) Then
				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"recordstateid=-1"
			End if

			If not ifnull(Sys_Note(i)) Then
				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"Note='"&Sys_Note(i)&"'"
			End if
			
			If not ifnull(updstr) Then
				strSQL="Update BillStopCarAccept set "&updstr&" where billno='"&trim(Sys_BillNo(i))&"' "&accwhere

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
			<strong>攔停登記簿系統</strong>
			<a href="./Upaddress/CheckAccept.doc"><font size="3" color="blue"><u>使用說明</u></font></a>
			<B><font size="5" color="red">列印完請立即點選『列印完成』。</font></B>
			<input type="button" name="btnOK" class="btn3" style="width:250px;height:25px;font-size:15px;" value="匯入攔停登記簿檔案(委外廠商使用)" onclick="funChkSelt();">			
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						<table border="0">
							<tr>
								<td>
									<strong>待審核記錄</strong><span id='BillBaseOrder'></span><br>
									&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
									<%If session("UnitLevelID") <> "3" Then%>
										<input type="button" name="btnOK" class="btn3" style="width:120px;height:25px;font-size:16px;" value="全部審核通過" onclick="funSaveBat();">
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<!--										<input type="button" name="cancel" value="合併列印"  onclick="funAcceptAllLoad();">
										&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; 
										<input type="button" name="cancel" value="合併列印完成"  onclick="funPrintBatOver();">-->
									<%End if%>
									
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp; 
									<input type="button" name="cancel" value="清除" onClick="location='BillBaseStopCheckAccept_miaoli.asp'">
									<br>
									<table width="750" border="1" cellpadding="0" cellspacing="0">
										<tr bgcolor="#EBFBE3" align="center">
											<th width="60">建檔日</th>
											<th width="105">送件單位</th>
											<th width="60">建檔人</th>
											<th width="60">件數</th>
											<th width="60">退件數</th>
											<th width="220">操作</th>
										</tr>
										<tr>
											<td colspan="6">
												<Div style="overflow:auto;width:100%;height:200px;background:#FFFFFF">
												<table width="100%" border="0" cellpadding="1" cellspacing="0"><%
												chkExp=1

												strSQL="select zipname from zip where zipname like '苗栗%'"
												set rszip=conn.execute(strSQL)
												strZip="<option value=''>請選擇</option>"
												while not rszip.eof
													strZip=strZip&"<option value='"&trim(rszip("zipname"))&"'>"&trim(rszip("zipname"))&"</option>"
													rszip.movenext
												wend
												rszip.close
												

												If session("UnitLevelID") = "3" Then
													chkwhere="RecordMemberID2 is null and RecordMemberID3 is null"
													showUnit=" and billunitid in('"&trim(Session("Unit_ID"))&"')"	

												elseIf session("UnitLevelID") = "2" Then
													chkwhere="(RecordMemberID2 is not null or RecordMemberID1="&trim(Session("User_ID"))&") and RecordMemberID3 is null"
													showUnit=" and billunitid in(select unitid from unitinfo where unittypeid in(select unittypeid from unitinfo where unitid='"&trim(Session("Unit_ID"))&"'))"

												else
													chkwhere="RecordMemberID3 is null"
													showUnit=" and billunitid in(select unitid from unitinfo where unittypeid in(select unittypeid from unitinfo where unitid='"&trim(Session("Unit_ID"))&"'))"

												end if

												strSQL="select a.AcceptDate,a.BillUnitID,a.chkType,a.RecordMemberID1,(select chName from MemberData where Memberid=a.RecordMemberID1) ChName,a.suess,a.delss,c.UnitName from (select DeCode(RecordMemberID2,null,to_char(AcceptDate,'YYYY/MM/DD'),to_char(RecordDate2,'YYYYMMDDHH24')) AcceptDate,BillUnitID,decode(RecordMemberID2,null,'0','1') chkType,Decode(RecordMemberID2,null,RecordMemberID1,RecordMemberID2) RecordMemberID1,sum(decode(recordstateid,0,1,0)) suess, sum(decode(recordstateid,-1,1,0)) delss from BillStopCarAccept where "&chkwhere&showUnit&" group by DeCode(RecordMemberID2,null,to_char(AcceptDate,'YYYY/MM/DD'),to_char(RecordDate2,'YYYYMMDDHH24')),BillUnitID,decode(RecordMemberID2,null,'0','1'),Decode(RecordMemberID2,null,RecordMemberID1,RecordMemberID2)) a,UnitInfo c where a.BillUnitID=c.UnitID order by a.AcceptDate,a.BillUnitID"

												set rs=conn.execute(strSQL)
												While not rs.eof
													Response.Write "<tr align=""center"">"

													strAcceptDate=""

													If trim(rs("chkType"))="1" Then
														Response.Write "<td width=""80"">"&(left(rs("AcceptDate"),4)-1911)&Mid(rs("AcceptDate"),5,4)&"</th>"
														strAcceptDate=rs("AcceptDate")

													else
														Response.Write "<td width=""80"">"&gInitDT(rs("AcceptDate"))&"</th>"
														strAcceptDate=gInitDT(rs("AcceptDate"))

													End If  

													If session("UnitLevelID") = "1" and trim(rs("chkType"))="0" Then
														If trim(Session("Ch_Name"))= trim(rs("ChName")) and gInitDT(rs("AcceptDate"))<> gInitDT(date) Then
															chkExp=0
														End if 
													end If

													Response.Write "<td width=""130"">"&rs("UnitName")&"</th>"
													Response.Write "<td width=""80"">"&rs("ChName")&"</th>"
													Response.Write "<td width=""75"">"&rs("suess")&"</th>"
													Response.Write "<td width=""75"">"&rs("delss")&"</th>"
													Response.Write "<td width=""""><input type=""button"" name=""btnAcc"" class=""btn3"" style=""width:40px;height:25px;font-size:16px;"" value=""詳細"" onclick=""funAcceptLoad('"&strAcceptDate&"','"&rs("BillUnitID")&"','"&rs("UnitName")&"','"&cdbl(rs("suess"))+cdbl(rs("delss"))&"','"&rs("RecordMemberID1")&"','"&rs("chkType")&"');"">"

													Response.Write "&nbsp;<input type=""button"" name=""Update"" class=""btn3"" value=""列印"" class=""btn3"" style=""width:40px;height:25px;font-size:16px;"" onclick=""funAcceptRunList('"&strAcceptDate&"','"&rs("BillUnitID")&"','"&rs("RecordMemberID1")&"','"&rs("chkType")&"');"">"

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
							
							CarCode=",1,2,3,4,5,6,"
							AcceptCode=",A,U,2,3,5,"
							FastCode=""
							Response.Write "<br><br><font size=3 ><B>"
							Response.Write "車種代碼"
							Response.Write "</B></font><br>"

							Response.Write "<font size=3 color=""Red"">"

							Response.Write "1汽車 &nbsp; 2拖車 &nbsp; 3重機/550cc以上&nbsp; 4輕機&nbsp; 5動力機械&nbsp 6臨時車牌"						
							Response.Write "</font><br>"

							Response.Write "<font size=3 ><B>"
							Response.Write "砂石註記"
							Response.Write "</B></font><br>"

							Response.Write "<font size=3 color=""Red"">"

							Response.Write "1是&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  0不是"						
							Response.Write "</font><br>"

							Response.Write "<font size=3 ><B>"
							Response.Write "簽收代碼"
							Response.Write "</B></font><br>"

							Response.Write "<font size=3 color=""Red"">"

							Response.Write "A簽收 &nbsp;  U拒簽收 &nbsp; 2拒簽已收 &nbsp; 3已簽拒收 &nbsp; 5補開單"
							
							Response.Write "</font><BR><BR>"

							If session("UnitLevelID") = "2" Then
								Response.Write "整批審核說明<input type=text name='BatNote' size=32 class='btn1'>"
								Response.Write "<br>"
								Response.Write "<input type=""button"" class=""btn3"" style=""width:120px;height:25px;font-size:16px;"" name=""btn_BatBack"" value=""整批退件(取消)"" onClick=""funBatBack();"">"
							end if

'							Response.Write "<br><font size=3 ><B>"
'							Response.Write "扣件代碼"
'							Response.Write "</B></font><br>"

'							Response.Write "<font size=3 color=""Red"">"
'							strSQL="select ID,Content from DCICode where typeid=6 order by ID"
'							set rscode=conn.execute(strSQL)
'							While not rsCode.eof
								
'								Response.Write rscode("ID")&""&rscode("Content")

'								FastCode=FastCode&rscode("ID")&"&nbsp;"

'								If trim(rscode("ID"))="6" Then
'									Response.Write "<BR>"
'								elseIf trim(rscode("ID"))="E" Then
'									Response.Write "<BR>"
'								else
'									Response.Write "&nbsp;"
'								end if
											
'								rsCode.movenext
'							Wend

'							FastCode=","&FastCode

'							Response.Write "</font>"
'							rsCode.close
						%>								
								
									<input type="Hidden" name='AcceptDate' size="5" class='btn1' maxlength='7' value="<%=gInitDT(now)%>">
							
									<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
		
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">攔停點收紀錄列表 ( 輸入完成按Enter可自動跳到下一格 ) <img src="space.gif" width="9" height="8">
		<B>違警案件『車號』、『車種』、『砂石註記』等攔位無需填寫。</B><input type="button" name="insert" class="btn3" style="width:80px;height:25px;font-size:16px;" value="再多50筆" onClick="insertRow(fmyTable)"></td>
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
var chklaw=",18,21,29,31,32,63,67,81,85,90,";
if(!chkExp){
	alert("請先處理之前案件!!");
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

		txtArea.innerHTML = "<b>" + cnt_num + "</b>&nbsp;&nbsp;"+
		"<span style='color:#ff0000;'>*</span>" +
		"單號<input type=text style='ime-mode:disabled;' name='item' size=10 class='btn1' maxlength='9' onfocus='funAddBillNo("+cunt+");' onkeyup='UpperCase(this);' onkeydown='KeyFunction("+cunt+");' maxlength='11'>" +
		"&nbsp;&nbsp;" +
		"車號<input type=text style='ime-mode:disabled;' name='CarNo' size=8 class='btn1' maxlength='10' onkeyup='UpperCase(this);' onkeydown='KeyCarNo("+cunt+");'>" +
		"&nbsp;&nbsp;" +
		"車種<input type=text style='ime-mode:disabled;' name='CarSimpleID' size=2 class='btn1' maxlength='3' onkeyup='chknumber(this);' onkeydown='KeyCarSimpleID("+cunt+");' maxlength='1'>" +
		"&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"砂石註記<input type=text style='ime-mode:disabled;' name='CarAddID' size=1 class='btn1' value='0' maxlength='3' onkeyup='chknumber(this);' onkeydown='KeyCarAddID("+cunt+");' maxlength='1'>" +
		"&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"違規日<input type=text style='ime-mode:disabled;' name='illegalDate' size=5 class='btn1' maxlength='7' onkeyup='chknumber(this);' onkeydown='KeyillegalDate("+cunt+");' maxlength='7'>" +
		"&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"時間<input type=text style='ime-mode:disabled;' name='illegalTime' size=1 class='btn1' maxlength='5' onkeyup='chknumber(this);' onkeydown='KeyillegalTime("+cunt+");' maxlength='4'>" +
		"&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"鄉鎮市<select name='IllegalAddress'><%=strZip%></select>" +
		"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"法條1<input type=text style='ime-mode:disabled;' name='Rule1_1' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule1_1("+cunt+");'>條"  +		
		"<input type=text style='ime-mode:disabled;' name='Rule1_2' size=1 maxlength='1' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule1_2("+cunt+");'>項"  +		
		"<input type=text style='ime-mode:disabled;' name='Rule1_3' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule1_3("+cunt+");'>款"  +		
		"之<input type=text style='ime-mode:disabled;' name='Rule1_4' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule1_4("+cunt+");'>"  +
		"<img src=\u0022../Image/BillkeyInButton.jpg\u0022 width=\u002225\u0022 height=\u002223\u0022 onclick='window.open(\u0022Query_Law.asp?LawOrder=1&RuleVer=<%=sys_RuleVer%>&objCnt="+cunt+"\u0022,\u0022WebPage1\u0022,\u0022left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes\u0022);'>" +
		"&nbsp;&nbsp;" +
		"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"簽收&nbsp;&nbsp;<input type=text style='ime-mode:disabled;' name='BillStatus' size=1 value='A' class='btn1' onkeyup='UpperCase(this);' onkeydown='KeyBillStatus("+cunt+");' maxlength='1'>" +
		"&nbsp;&nbsp;" +
		"<span style='color:#ff0000;'>*</span>" +
		"員警1<input type=text style='ime-mode:disabled;' name='BillMemID1' size=2 class='btn1' onkeydown='KeyBillMemID1("+cunt+");' onkeyup='getBillMem1("+cunt+");'>" +
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
		"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +		
		"法條2<input type=text style='ime-mode:disabled;' name='Rule2_1' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule2_1("+cunt+");'>條" +
		"<input type=text style='ime-mode:disabled;' name='Rule2_2' size=1 maxlength='1' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule2_2("+cunt+");'>項" +
		"<input type=text style='ime-mode:disabled;' name='Rule2_3' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule2_3("+cunt+");'>款" +
		"之<input type=text style='ime-mode:disabled;' name='Rule2_4' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule2_4("+cunt+");'>"  +
		"<img src=\u0022../Image/BillkeyInButton.jpg\u0022 width=\u002225\u0022 height=\u002223\u0022 onclick='window.open(\u0022Query_Law.asp?LawOrder=2&RuleVer=<%=sys_RuleVer%>&objCnt="+cunt+"\u0022,\u0022WebPage1\u0022,\u0022left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes\u0022);'>" +
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
		"法條3<input type=text style='ime-mode:disabled;' name='Rule3_1' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule3_1("+cunt+");'>條" +
		"<input type=text style='ime-mode:disabled;' name='Rule3_2' size=1 maxlength='1' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule3_2("+cunt+");'>項" +
		"<input type=text style='ime-mode:disabled;' name='Rule3_3' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule3_3("+cunt+");'>款" +
		"之<input type=text style='ime-mode:disabled;' name='Rule3_4' size=1 maxlength='2' class='btn1' onkeyup='chknumber(this);' onkeydown='KeyRule3_4("+cunt+");'>"  +
		"<img src=\u0022../Image/BillkeyInButton.jpg\u0022 width=\u002225\u0022 height=\u002223\u0022 onclick='window.open(\u0022Query_Law.asp?LawOrder=3&RuleVer=<%=sys_RuleVer%>&objCnt="+cunt+"\u0022,\u0022WebPage1\u0022,\u0022left=0,top=0,location=0,width=900,height=555,resizable=yes,scrollbars=yes\u0022);'>" +
		"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
		"扣件1<select name='FastenerTypeID1'><option value=''>請選擇</option><option value='1'>行照</option><option value='2'>駕照</option><option value='3'>告發單</option>" +
		"<option value='4'>牌一面</option><option value='5'>牌二面</option><option value='6'>空車</option><option value='7'>其它</option><option value='8'>錄影</option>" +
		"<option value='9'>照片</option><option value='A'>汽駕</option><option value='B'>機駕</option><option value='C'>本牌乙</option><option value='D'>本牌二</option>" +
		"<option value='E'>移牌乙</option><option value='F'>移牌二</option><option value='G'>臨時牌</option><option value='H'>試車牌</option><option value='I'>執登證</option>" +
		"<option value='J'>測速器</option><option value='K'>喇叭</option><option value='L'>車鑰匙</option></select>" +
		"&nbsp&nbsp" +
		"扣件2<select name='FastenerTypeID2'><option value=''>請選擇</option><option value='1'>行照</option><option value='2'>駕照</option><option value='3'>告發單</option>" +
		"<option value='4'>牌一面</option><option value='5'>牌二面</option><option value='6'>空車</option><option value='7'>其它</option><option value='8'>錄影</option>" +
		"<option value='9'>照片</option><option value='A'>汽駕</option><option value='B'>機駕</option><option value='C'>本牌乙</option><option value='D'>本牌二</option>" +
		"<option value='E'>移牌乙</option><option value='F'>移牌二</option><option value='G'>臨時牌</option><option value='H'>試車牌</option><option value='I'>執登證</option>" +
		"<option value='J'>測速器</option><option value='K'>喇叭</option><option value='L'>車鑰匙</option></select>" +
		"<input type='hidden' name='old_BillNo' value=''>"+
		"&nbsp;&nbsp;" +
		//"違規人姓名<input type=hidden name='DriverName' size=2 class='btn1' onkeydown='KeyDriverName("+cunt+");' maxlength='4'>" +
		//"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" +
		"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"+	<%
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
	UrlStr="BillBaseCheckAcceptSendStyle_miaoli.asp";
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

	UrlStr="AcceptStopList.asp";
	
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
	myForm.DB_RecordMemberID2.value="";

	UrlStr="AcceptStopList.asp";
	
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

function funPrintOver(AcceptDate,BillUnitID,RecordMemberID1){
	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordMemberID1.value=RecordMemberID1;
	myForm.DB_RecordMemberID2.value="";

	myForm.DB_Selt.value="PrintOver";
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

	runServerScript("getStopCarAcceptData_miaoli.asp?AcceptDate="+AcceptDate+"&UnitID="+UnitID+"&objcnt="+myForm.item.length+"&RecordMemberID1="+RecordMemberID1+"&chkType="+chkType);

	for(i=Cmt;i<myForm.item.length;i++){

		myForm.item[i].value='';
				
		myForm.CarNo[i].value='';

		myForm.CarSimpleID[i].value='';

		myForm.CarAddID[i].value='';

		myForm.illegalDate[i].value='';

		myForm.illegalTime[i].value='';

		myForm.IllegalAddress[i].value='';

		myForm.Rule1_1[i].value='';

		myForm.Rule1_2[i].value='';

		myForm.Rule1_3[i].value='';

		myForm.Rule1_4[i].value='';

		myForm.FastenerTypeID1[i].value='';

		myForm.BillMemID1[i].value='';

		BillMemName1[i].innerHTML='';

		myForm.BillMemID2[i].value='';

		BillMemName2[i].innerHTML='';

		myForm.BillMemID3[i].value='';

		BillMemName3[i].innerHTML='';

		myForm.BillMemID4[i].value='';

		BillMemName4[i].innerHTML='';

		myForm.BillStatus[i].value='';

		myForm.Rule2_1[i].value='';

		myForm.Rule2_2[i].value='';

		myForm.Rule2_3[i].value='';

		myForm.Rule2_4[i].value='';

		myForm.FastenerTypeID2[i].value='';

		myForm.Rule3_1[i].value='';

		myForm.Rule3_2[i].value='';

		myForm.Rule3_3[i].value='';

		myForm.Rule3_4[i].value='';

		myForm.Sys_BackBillBase[i].value='';
		
		myForm.chkBackBillBase[i].checked=false;

		myForm.Note[i].value='';

		myForm.old_BillNo[i].value='';

	}
}

function KeyFunction(itemcnt) {

	if (event.keyCode==13||event.keyCode==9) {
		if (!chkBillNo(itemcnt)){
			alert("單號長度必須為9碼!!");
		}else{

			if(myForm.IllegalAddress[itemcnt-2]){

				myForm.illegalDate[itemcnt-1].value=myForm.illegalDate[itemcnt-2].value;
				
				myForm.IllegalAddress[itemcnt-1].value=myForm.IllegalAddress[itemcnt-2].value;

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

			}

			myForm.CarNo[itemcnt-1].focus();
		}
	}
}
function funAddBillNo(itemcnt) {
	var BillNoStart="";
	var BillNoEnd=""
	if(itemcnt!=1){
		if(myForm.item[itemcnt-2]){
			BillNoStart=myForm.item[itemcnt-2].value.substr(0,2);
			BillNoEnd=myForm.item[itemcnt-2].value.substr(2,7);

			if(myForm.item[itemcnt-1]){
				myForm.item[itemcnt-1].value=BillNoStart+("000000"+(Number(BillNoEnd)+1)).substr(("000000"+(Number(BillNoEnd)+1)).length-7,7);
			}
		}
	}
}
function KeyCarNo(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		if (myForm.CarNo[itemcnt-1].value.length>9){
			alert("車號不可超過9碼!!");
		}else{
			myForm.CarSimpleID[itemcnt-1].focus();
		}		
	}
}

function KeyCarSimpleID(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.CarAddID[itemcnt-1].focus();
	}
}

function KeyCarAddID(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.illegalDate[itemcnt-1].focus();
	}
}

function KeyillegalDate(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.illegalTime[itemcnt-1].focus();
	}
}

function KeyillegalTime(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Rule1_1[itemcnt-1].focus();
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
		if(chklaw.search(myForm.Rule1_1[itemcnt-1].value)<=0){
			myForm.BillStatus[itemcnt-1].focus();
		}else{
			myForm.Rule1_4[itemcnt-1].focus();
		}
	}
}
function KeyRule1_4(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.BillStatus[itemcnt-1].focus();
	}
}
function KeyFastenerTypeID1(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}
function KeyBillMemID1(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.BillMemID2[itemcnt-1].focus();
	}
}
function getBillMem1(itemcnt){
	UpperCase(myForm.BillMemID1[itemcnt-1]);
	BillMemName1[itemcnt-1].innerHTML="";
	runServerScript("CheckStopCarAcceptMemID_miaoli.asp?LoginID="+myForm.BillMemID1[itemcnt-1].value+"&innerObj=BillMemName1&itemcnt="+(itemcnt-1));
}
function KeyBillMemID2(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}
function getBillMem2(itemcnt){
	UpperCase(myForm.BillMemID2[itemcnt-1]);
	BillMemName2[itemcnt-1].innerHTML="";
	runServerScript("CheckStopCarAcceptMemID_miaoli.asp?LoginID="+myForm.BillMemID2[itemcnt-1].value+"&innerObj=BillMemName2&itemcnt="+(itemcnt-1));
}

function KeyBillMemID3(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}
function getBillMem3(itemcnt){
	UpperCase(myForm.BillMemID3[itemcnt-1]);
	BillMemName3[itemcnt-1].innerHTML="";
	runServerScript("CheckStopCarAcceptMemID_miaoli.asp?LoginID="+myForm.BillMemID3[itemcnt-1].value+"&innerObj=BillMemName3&itemcnt="+(itemcnt-1));
}

function KeyBillMemID4(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}
function getBillMem4(itemcnt){
	UpperCase(myForm.BillMemID4[itemcnt-1]);
	BillMemName4[itemcnt-1].innerHTML="";
	runServerScript("CheckStopCarAcceptMemID_miaoli.asp?LoginID="+myForm.BillMemID4[itemcnt-1].value+"&innerObj=BillMemName4&itemcnt="+(itemcnt-1));
}

function KeyBillStatus(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.BillMemID1[itemcnt-1].focus();
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
		if(chklaw.search(myForm.Rule2_1[itemcnt-1].value)<=0){
			myForm.item[itemcnt].focus();
		}else{
			myForm.Rule2_4[itemcnt-1].focus();
		}
	}
}
function KeyRule2_4(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}
function KeyFastenerTypeID2(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}
function KeyRule3_1(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Rule3_2[itemcnt-1].focus();
	}
}
function KeyRule3_2(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Rule3_3[itemcnt-1].focus();
	}
}
function KeyRule3_3(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		if(chklaw.search(myForm.Rule3_1[itemcnt-1].value)<=0){
			myForm.item[itemcnt].focus();
		}else{
			myForm.Rule3_4[itemcnt-1].focus();
		}
	}
}
function KeyRule3_4(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}
function KeyDriverName(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}
function KeyNote(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}

function funChkBackBillBase(itemcnt){
	if(myForm.chkBackBillBase[itemcnt-1].checked){
		myForm.Sys_BackBillBase[itemcnt-1].value="-1";
	}else{
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
	for(i=0;i<myForm.item.length;i++){
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

	for(i=0;i<myForm.item.length;i++){
		if(myForm.illegalDate[i].value!=''){
			
			if(myForm.item[i].value!='' && myForm.item[i].value.substr(0,1)!='F'){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行單號錯誤!!\n";
			}

			if(Number(myForm.Rule1_1[i].value)>68){

				myForm.CarNo[i].value="違警";
			}

			if(myForm.CarNo[i].value==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行車號不可空白!!\n";
			}
			
			if(myForm.illegalDate[i].value==''||myForm.illegalTime[i].value==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規日期時間不可空白!!\n";
			}

			if(myForm.illegalDate[i].value.length!=7){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規日期錯誤!!\n";
			}

			if(myForm.illegalDate[i].value > chkDate ){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規日期不可超過今天!!\n";
			}

			if(myForm.illegalTime[i].value > 2359){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規時間不正確!!\n";
			}

			if(myForm.illegalTime[i].value.length!=4){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行違規時間錯誤!!\n";
			}
			
			if(Number(myForm.Rule1_1[i].value)<69){
				if(myForm.CarSimpleID[i].value==''){
					err=1;
					errmsg=errmsg+"第 "+(i+1)+" 行車種不可空白!!\n";
				}
			}

			if(myForm.CarSimpleID[i].value!=''&& "<%=CarCode%>".indexOf(myForm.CarSimpleID[i].value,0)<0){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行車種錯誤!!\n";
			}

			if(Number(myForm.Rule1_1[i].value)<69){

				if(myForm.CarAddID[i].value=='' || (myForm.CarAddID[i].value!=1 && myForm.CarAddID[i].value!=0)){
					err=1;
					errmsg=errmsg+"第 "+(i+1)+" 行砂石錯誤!!\n";
				}
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

			if(myForm.Rule3_1[i].value!=''||myForm.Rule3_2[i].value!=''||myForm.Rule3_3[i].value!=''){
				if(myForm.Rule3_1[i].value==''||myForm.Rule3_2[i].value==''||myForm.Rule3_3[i].value==''){
					err=1;
					errmsg=errmsg+"第 "+(i+1)+" 行法條3不可空白!!\n";
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

			if(myForm.BillStatus[i].value==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行簽收狀態不可空白!!\n";
			}

			if(myForm.BillStatus[i].value!=''&& "<%=AcceptCode%>".indexOf(myForm.BillStatus[i].value,0)<0){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行簽收狀態錯誤!!\n";
			}

/*			if(myForm.FastenerTypeID1[i].value!=''&& "<%=FastCode%>".indexOf(myForm.FastenerTypeID1[i].value,0)<0){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行扣件1錯誤!!\n";
			}

			if(myForm.FastenerTypeID2[i].value!=''&& "<%=FastCode%>".indexOf(myForm.FastenerTypeID2[i].value,0)<0){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行扣件2錯誤!!\n";
			}
*/
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