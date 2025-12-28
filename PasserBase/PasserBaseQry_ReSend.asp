<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
If trim(Request("chk_City"))="台南市" and (not ifnull(Request("Ch_Name"))) Then

	strSQL="select * from memberdata where RecordstateID=0 and chName='"&trim(Request("Ch_Name"))&"' and UnitID in(select UnitID from Unitinfo where UnitName like '"&trim(Request("Unit_Name"))&"%') order by MODIFYTIME DESC"

	set rsmen=conn.execute(strSQL)

	If not rsmen.eof Then
		Session("FuncID")="233,1,1,1,1&&254,1,1,1,1&&252,1,1,1,1&&253,1,1,1,1&&250,1,1,1,1&&297,1,1,1,1&&249,1,1,1,1&&255,1,1,1,1&&245,1,1,1,1&&228,1,1,1,1&&226,1,1,1,1&&290,1,1,1,1&&230,1,1,1,1&&281,1,1,1,1&&280,1,1,1,1&&1601,1,1,1,1&&285,1,1,1,1&&1500,1,1,1,1&&235,1,1,1,1&&224,1,1,1,1&&299,1,1,1,1&&236,1,1,1,1&&293,1,1,1,1&&270,1,1,1,1&&229,1,1,1,1&&298,1,1,1,1&&263,1,1,1,1&&1729,1,1,1,1&&220,1,1,1,1&&260,1,1,1,1&&262,1,1,1,1&&261,1,1,1,1&&227,1,1,1,1&&901,1,1,1,1&&234,1,1,1,1&&223,1,1,1,1&&221,1,1,1,1"

		Session("Unit_ID")=rsmen("UnitID")
		Session("User_ID")=rsmen("MemberID")
		Session("Ch_Name")=rsmen("ChName")
		Session("Credit_ID")=rsmen("CreditID")
		Session("Group_ID")=rsmen("GroupRoleID")
		Session("ManagerPower")=rsmen("ManagerPower")
		'單位等級
		strUnit="select UnitLevelID,DCIwindowName from UnitInfo where UnitID='"&trim(rsmen("UnitID"))&"'"
		set rsUnit=conn.execute(strUnit)
		if not rsUnit.eof then
			UnitLevel=trim(rsUnit("UnitLevelID"))
			DCIwindow=trim(rsUnit("DCIwindowName"))
		end if
		rsUnit.close
		set rsUnit=nothing
		Session("UnitLevelID")=UnitLevel
		Session("DCIwindowName")=DCIwindow
	else
		Response.Redirect "/traffic/Traffic_Login.asp?Error=無使用之權限,請洽工程師。"
	End if
	rsmen.close
End if

%>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>慢車行人道路障礙裁罰</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
.btn3{
   font-size:16px;
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
   font-weight:900;
}
.font10{
   font-size:16px;
   font-family:新細明體;
}
</style>

</head>
<%
'檢查是否可進入本系統
AuthorityCheck(224)
DB_Orderby=split("Driver,IllegaLDate",",")
DB_KindSelt=trim(request("DB_KindSelt"))
DB_Selt=trim(request("DB_Selt"))
DB_Display=request("DB_Display")

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

sys_City=replace(sys_City,"台中縣","台中市")
sys_City=replace(sys_City,"台南縣","台南市")

showCreditor=false
if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then
	showCreditor=true

	If isempty(request("DB_Selt")) Then

		strSQL="delete PasserCreditor where PetitionDate is null"

		conn.execute(strSQL)

		strSQL="delete TRAFFIC.PASSERSENDDETAIL where (select count(1) cnt from PASSERSENDDETAIL dt where billsn=PASSERSENDDETAIL.billsn and SENDDATE=PASSERSENDDETAIL.SENDDATE)>1 and (select count(1) cnt from passersend sd where billsn=PASSERSENDDETAIL.billsn and SENDDATE=PASSERSENDDETAIL.SENDDATE)=0 and not exists(select 'N' from PasserCreditor where SendDetailSN=PasserSendDetail.SN and PetitionDate is not null)"
		conn.execute(strSQL)

		strSQL="delete TRAFFIC.PASSERSENDDETAIL where (select count(1) cnt from PASSERSENDDETAIL dt where billsn=PASSERSENDDETAIL.billsn)=1 and (select count(1) cnt from passersend sd where billsn=PASSERSENDDETAIL.billsn and SENDDATE=PASSERSENDDETAIL.SENDDATE)=1 and not exists(select 'N' from PasserCreditor where SendDetailSN=PasserSendDetail.SN and PetitionDate is not null)"
		conn.execute(strSQL)
	
	End if 
	
end If

strRul="select Value from Apconfigure where ID=3"
set rsRul=conn.execute(strRul)
RuleVer=trim(rsRul("Value"))
rsRul.Close

'==========================================================================================
strSQL="select * from UnitInfo where rownum=1"
set rs=conn.execute(strSQL)
If Not rs.eof Then
	For i=0 to rs.Fields.count-1
		If trim(rs.Fields.item(i).Name)="PASSERSENDBANKNAME" Then Exit For
	Next
	If i>rs.Fields.count-1 Then
		strSQL="Alter Table UnitInfo ADD (PASSERSENDBANKNAME VarChar2(80))"
		conn.execute(strSQL)
	End if
	For i=0 to rs.Fields.count-1
		If trim(rs.Fields.item(i).Name)="PASSERSENDBANKACCOUNT" Then Exit For
	Next
	If i>rs.Fields.count-1 Then
		strSQL="Alter Table UnitInfo ADD (PASSERSENDBANKACCOUNT VarChar2(30))"
		conn.execute(strSQL)
	End if
End if

'==========================================================================================

if DB_Selt="Selt" then
	strwhere="":strDate=""
	'日期
	if trim(request("Sys_DoubleCheckStatus"))<>"" then
		strwhere=strwhere&" and DoubleCheckStatus="&trim(request("Sys_DoubleCheckStatus"))
	end if

	if request("RecordDate1")<>"" and request("RecordDate2")<>""then
		ArgueDate1=gOutDT(request("RecordDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("RecordDate2"))&" 23:59:59"
		strwhere=strwhere&" and RecordDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	if request("BillFillDate1")<>"" and request("BillFillDate2")<>""then
		ArgueDate1=gOutDT(request("BillFillDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("BillFillDate2"))&" 23:59:59"
		strwhere=strwhere&" and BillFillDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	if request("IllegalDate1")<>"" and request("IllegalDate2")<>""then
		ArgueDate1=gOutDT(request("IllegalDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("IllegalDate2"))&" 23:59:59"
		strwhere=strwhere&" and IllegalDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end If 
	
	if request("Sys_SendDetailDate1")<>"" and request("Sys_SendDetailDate2")<>""then
		ArgueDate1=gOutDT(request("Sys_SendDetailDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("Sys_SendDetailDate2"))&" 23:59:59"

		strwhere=strwhere&" and Exists(select 'Y' from PasserSendDetail where SendDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and billsn=a.sn)"
	end If 
	
	if trim(request("sys_CreditorTypeID"))<>"" Or (trim(request("Sys_PetitionDate1"))<>"" And trim(request("Sys_PetitionDate2"))<>"") Then
		strPasserCreditorAdd=""
		If trim(request("sys_CreditorTypeID")) <> "-1" Then
			If trim(request("sys_CreditorTypeID"))<>"" Then
				strPasserCreditorAdd=" CreditorTypeID in('"&trim(request("sys_CreditorTypeID"))&"')"
			End If 
			If trim(request("Sys_PetitionDate1"))<>"" And trim(request("Sys_PetitionDate2"))<>"" Then
				PetitionDate1=gOutDT(request("Sys_PetitionDate1"))&" 0:0:0"
				PetitionDate2=gOutDT(request("Sys_PetitionDate2"))&" 23:59:59"
				If strPasserCreditorAdd="" Then
					strPasserCreditorAdd=" PetitionDate between TO_DATE('"&PetitionDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&PetitionDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				Else
					strPasserCreditorAdd=strPasserCreditorAdd&" and PetitionDate between TO_DATE('"&PetitionDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&PetitionDate2&"','YYYY/MM/DD/HH24/MI/SS')"
				End If 
			End If
		end if

		If trim(request("sys_CreditorTypeID")) = "-1" Then
			strwhere=strwhere&" and not Exists(select 'N' from PasserSendDetail where Exists(select 'Y' from PasserCreditor where CreditorTypeID in('0','1') and SendDetailSN=PasserSendDetail.SN) and billsn=a.sn)"
		else
			strwhere=strwhere&" and Exists(select 'Y' from PasserSendDetail where Exists(select 'Y' from PasserCreditor where "&strPasserCreditorAdd&" and SendDetailSN=PasserSendDetail.SN) and BillSn=a.SN)"
		End if		
		
	end if  

	if request("DeallIneDate1")<>"" and request("DeallIneDate2")<>""then
		ArgueDate1=gOutDT(request("DeallIneDate1"))&" 0:0:0"
		ArgueDate2=gOutDT(request("DeallIneDate2"))&" 23:59:59"
		strwhere=strwhere&" and DeallIneDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if

	'單位
	if trim(request("Sys_BillUnitID"))<>"" then
		strwhere=strwhere&" and BillUnitID in ('"&request("Sys_BillUnitID")&"')"
	end if 

	'單號
	if request("Sys_BillNo")<>"" then
		strwhere=strwhere&" and BillNo ='" & Ucase(request("Sys_BillNo")) &  "'"
	end If 

	'違規人姓名
	if request("Sys_Driver")<>"" then
		strwhere=strwhere&" and Driver='"&request("Sys_Driver")&"'"
	end If 
 
	'違規人身分証號
	if request("Sys_DriverID")<>"" then
		strwhere=strwhere&" and DriverID='"&Ucase(request("Sys_DriverID"))&"'"
	end if

	if request("Sys_Rule")<>"" then
		strwhere=strwhere&" and (Rule1 like '"&trim(request("Sys_Rule"))&"%' or Rule2 like '"&trim(request("Sys_Rule"))&"%' or Rule3 like '"&trim(request("Sys_Rule"))&"%' or Rule4 like '"&trim(request("Sys_Rule"))&"%')"
	end If 
	
	if request("Sys_MemberStation")<>"" then
		strwhere=strwhere&" and MemberStation in('"&request("Sys_MemberStation")&"')"
	end If 

	if request("Sys_SendKind")="A" then

		strwhere=strwhere&" and not exists(select 'N' from PasserSendDetail where Not Exists(select 'N' from PasserCreditor where SendDetailSN=PasserSendDetail.SN and PetitionDate is not null) and billsn=a.sn) and BillStatus<>9"

	end If 

	if trim(request("Sys_Order"))<>"" then
		orderstr=" order by "&request("Sys_Order")
	end if
end if

if DB_Selt="Selt" then
	if trim(strwhere)="" then
		DB_Selt="":DB_Display=""
		Response.write "<script>"
		Response.Write "alert('必須有查詢條件！');"
		Response.write "</script>"
	end if
end If 

Sys_SendBillSN=trim(request("Sys_SendBillSN"))

if strwhere="" then

	strwhere=strwhere&" and not exists(select 'N' from PasserSendDetail where Not Exists(select 'N' from PasserCreditor where SendDetailSN=PasserSendDetail.SN and PetitionDate is not null) and billsn=a.sn) and BillStatus<>9"

	if trim(Session("UnitLevelID")) > "1" then

		 strwhere=strwhere&" and MemberStation=(select UnitTypeID from Unitinfo where UnitID='"&Session("Unit_ID")&"')"
	end If 

	orderstr=" order by DoubleCheckStatus,BillNo"

	DB_Display="show"
	'DB_Selt="Selt"
end If 

if DB_Display="show" then

	showFiled=",(select max(PetitionDate) PetitionDate from PasserCreditor where billsn=a.SN) PetitionDate"
	
	strSQLTemp="select a.SN,a.IllegalDate,a.BillNo,a.Driver,a.IllegalAddress,a.Rule1," &_
	"a.RuleVer,a.FORFEIT1,a.FORFEIT2,a.FORFEIT3,a.FORFEIT4,a.BILLSTATUS," &_
	"a.BillMem1,a.DoubleCheckStatus," &_
	"(Select SendDate from PasserSend where billsn=a.sn) SENDDATE," &_
	"(Select MakeSureDate from PasserSend where billsn=a.sn) MakeSureDate," &_
	"(Select UrgeDate from PasserUrge where billsn=a.sn) URGEDATE," &_
	"(Select MakeSureDate from PasserSend where billsn=a.sn) MakeSureDate," &_
	"(select Max(ArrivedDate) ArrivedDate from PassersEndArrived where PasserSN=a.sn) ArrivedDate," &_
	"(select MAX(PayDate) PayDate from PasserPay where billsn=a.sn) PayDate"&showFiled &_
	" from PasserBase a where RecordStateID=0 "&strwhere&orderstr
	'set rsfound=conn.execute(strSQL)
	tmpSQL=strwhere

	strSQL="select a.SN,a.BillNo,a.DoubleCheckStatus,a.Driver,a.IllegalDate,a.Rule1"&showFiled&" from PasserBase a where RecordStateID=0 "&strwhere&orderstr
	
	cntUbound=0
	set rssn=conn.execute(strSQL)
	BillSN="":BillNo="":DoubleCheckStatus=""
	while Not rssn.eof
		cntUbound=cntUbound+1
		if trim(BillSN)<>"" then
			BillSN=BillSN&","
			BillNo=BillNo&","
		end if
		BillSN=BillSN&rssn("SN")	
		BillNo=BillNo&rssn("BillNo")
		rssn.movenext
	wend
	rssn.close

	strCnt="select count(1) as cnt from PasserBase a where RecordStateID=0 "&strwhere
	set Dbrs=conn.execute(strCnt)
	DBsum=0
	If Not Dbrs.eof Then
		DBsum=cint(Dbrs("cnt"))
	end if
	Dbrs.close

	if sys_City="宜蘭縣" or sys_City="澎湖縣" then

		strSQL="select 'Y' from PasserBase a where RecordStateID=0 "&strwhere

		upSQL="Update PasserBase set ForFeit1=(select max(Level2) from law where version="&RuleVer&" and ItemID=PasserBase.Rule1),ForFeit2=(select max(Level2) from law where version="&RuleVer&" and ItemID=PasserBase.Rule2) where Exists("&strSQL&"  and a.SN=PasserBase.Sn) and DeallineDate < to_Date('"&date&"','YYYY/MM/DD') and BillStatus=0 and RecordStateid=0 and Exists(select 'Y' from law where itemid=PasserBase.Rule1 and Level1=PasserBase.ForFeit1 and VerSion="&RuleVer&")"

		conn.execute(upSQL)
	end If 	
	'response.write strSQLTemp
	'response.end
end If 

If showCreditor Then
	if (sys_City="基隆市" or sys_City="屏東縣") and isempty(request("DB_Selt")) then
		
		errmsg="":PasserCnt=0:PasserSendCnt=0:MeberStation=""

		If Session("UnitLevelID") > 1 then MemberStation=" and MemberStation=(select UnitTypeID from Unitinfo where UnitID='"&Session("Unit_ID")&"')"

		strSQL="select (select UnitName from Unitinfo where UnitID=tmp.MemberStation) MebUnitName,cnt from ( select MemberStation,count(1) cnt from PasserBase where billstatus<>9 and recordstateid=0 and TRUNC(sysdate-DeallineDate)>60 and BillFillDate > to_date('"&(year(now)-6)&"/12/31','YYYY/MM/DD') and not exists(select 'Y' from PasserJude where billsn=PasserBase.SN) and Not exists(select 'Y' from PasserSend where billsn=PasserBase.SN)"&MemberStation&" group by MemberStation ) tmp where cnt > 0 order by MebUnitName"

		set rssn=conn.execute(strSQL)

		while Not rssn.eof
			PasserCnt=cdbl(rssn("cnt"))

			if trim(PasserCnt)>0 then errmsg=errmsg&rssn("MebUnitName")&"共有" & PasserCnt & "筆已逾60天尚未裁決\n"

			rssn.movenext
		wend
		
		rssn.close

		strSQL="select (select UnitName from Unitinfo where UnitID=tmp.MemberStation) MebUnitName,cnt from ( select MemberStation,count(1) cnt from PasserBase where billstatus<>9 and recordstateid=0 and BillFillDate > to_date('"&(year(now)-6)&"/12/31','YYYY/MM/DD') and exists(select 'Y' from PasserJude where billsn=PasserBase.SN and TRUNC(sysdate-JudeDate)>365) and Not exists(select 'Y' from PasserSend where billsn=PasserBase.SN)"&MemberStation&" group by MemberStation ) tmp where cnt > 0 order by MebUnitName"

		set rssn=conn.execute(strSQL)

		while Not rssn.eof
			PasserSendCnt=cdbl(rssn("cnt"))

			if trim(PasserSendCnt)>0 then errmsg=errmsg&rssn("MebUnitName")&"共有" & PasserSendCnt & "筆已裁決逾一年未移送\n"

			rssn.movenext
		wend
		
		rssn.close
		
		If not ifnull(errmsg) Then
			Response.write "<script>"
			Response.Write "alert('" & errmsg & "！');"
			Response.write "</script>"
		end if
		
	end If 
End if 
%>
<body onLoad="funLoadSend();">
<form name="myForm" method="post">
<table width="100%" border="0">
	<tr height="30">
		<td height="30" bgcolor="#FFCC33">
			<font size="4"><b>慢車行人道路障礙舉發單紀錄</b></font>
		</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
						<tr>
							<td>
							建檔序號
							</td><td>
							<input name="Sys_DoubleCheckStatus" class="btn1" type="text" value="<%=request("Sys_DoubleCheckStatus")%>" size="8" onkeyup="value=value.replace(/[^\d]/g,'')">
							</td><td>
							舉發單位
							</td><td>
							<select name="Sys_BillUnitID" class="btn1"><%
								strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
								set rsUnit=conn.execute(strSQL)
								If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
								rsUnit.close
								strUnitID=""
								if trim(Session("UnitLevelID"))="1" then
									strSQL="select UnitID,UnitName from UnitInfo order by UnitID,UnitName"
									strtmp=strtmp+"<option value="""">所有單位</option>"
								elseif trim(Session("UnitLevelID"))="2" then
									strSQL="select UnitID,UnitName,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"' or UnitTypeID=(select UnitTypeid from UnitInfo where UnitID='"&Session("Unit_ID")&"') or UnitLevelID=1 order by UnitTypeID,UnitName"

									set rs1=conn.execute(strSQL)
									while Not rs1.eof
										if trim(strUnitID)="" and rs1("UnitLevelID")>1 then
											strUnitID=trim(rs1("UnitID"))
										elseif rs1("UnitLevelID")>1 then 
											strUnitID=strUnitID&"','"&trim(rs1("UnitID"))
										end if
										rs1.movenext
									wend
									rs1.close
									strtmp=strtmp+"<option value="""">所有單位</option>"
									'strtmp=strtmp+"<option value="""&strUnitID&""">管轄單位</option>"
								elseif trim(Session("UnitLevelID"))="3" then
									strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"' order by UnitTypeID,UnitName"
								end if
								set rs1=conn.execute(strSQL)
								while Not rs1.eof
									strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
									if trim(rs1("UnitID"))=trim(request("Sys_BillUnitID")) then
										strtmp=strtmp+" selected"
									end if
									strtmp=strtmp+">"&rs1("UnitName")&"</option>"
									rs1.movenext
								wend
								rs1.close
								strtmp=strtmp+"</select>"
								response.write strtmp%>	
							</td>
						</tr>
						<tr>
							<td>
								<strong><font color="red">舉發單號</font></strong>
							</td>
							<td>
								<input name="Sys_BillNo" maxlength="9" size="8" class="btn1" type="text" value="<%=Ucase(request("Sys_BillNo"))%>" size="8" maxlength="20">
							</td>
							<td nowrap>
								應到案處
							</td>
							<td>
								<select name="Sys_MemberStation" class="btn1"><%
									strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
									set rsUnit=conn.execute(strSQL)
									If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
									rsUnit.close
									strUnitID="":strtmp=""
									if trim(Session("UnitLevelID"))="1" then
										strSQL="select UnitID,UnitName from UnitInfo order by UnitID,UnitName"
										strtmp=strtmp+"<option value="""">所有單位</option>"
									else
										strSQL="select UnitID,UnitName,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"' or UnitTypeID=(select UnitTypeid from UnitInfo where UnitID='"&Session("Unit_ID")&"') or UnitLevelID=1 order by UnitTypeID,UnitName"

										set rs1=conn.execute(strSQL)
										while Not rs1.eof
											if trim(strUnitID)="" and rs1("UnitLevelID")>1 then
												strUnitID=trim(rs1("UnitID"))
											elseif rs1("UnitLevelID")>1 then 
												strUnitID=strUnitID&"','"&trim(rs1("UnitID"))
											end if
											rs1.movenext
										wend
										rs1.close
										strtmp=strtmp+"<option value="""&strUnitID&""">所有單位</option>"
									end if
									set rs1=conn.execute(strSQL)
									while Not rs1.eof
										strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
										if trim(rs1("UnitID"))=trim(request("Sys_MemberStation")) then
											strtmp=strtmp+" selected"
										end if
										strtmp=strtmp+">"&rs1("UnitName")&"</option>"
										rs1.movenext
									wend
									rs1.close
									strtmp=strtmp+"</select>"
									response.write strtmp%>
							</td>
							<td nowrap>
								再移送狀況
							</td>
							<td>
								<select Name="Sys_SendKind" class="btn1">
									<option value="A"<%if trim(request("Sys_SendKind"))="A" then response.write " selected"%>>已取得債權</option>
									<option value="C"<%if trim(request("Sys_SendKind"))="B" then response.write " selected"%>>全部</option>
									
								</select>
							</td>
						</tr>
						<tr>
							<td>
								違規人名
							</td>
							<td>
								<input name="Sys_Driver" class="btn1" type="text" value="<%=request("Sys_Driver")%>" size="7" maxlength="8">
							</td>
							<td>
								身分證號
							</td>
							<td>
								<input name="Sys_DriverID" class="btn1" type="text" value="<%=Ucase(request("Sys_DriverID"))%>" size="10" maxlength="12">
							</td>
							<td>
								列表排序
							</td>
							<td>
								<select Name="Sys_Order" class="btn1">
									<option value="DoubleCheckStatus,BillNo"<%if trim(request("Sys_Order"))="DoubleCheckStatus,BillNo" then response.write " selected"%>>建檔序號</option>
									<option value="Driver,BillNo"<%if trim(request("Sys_Order"))="Driver,BillNo" then response.write " selected"%>>違規人</option>
									<option value="IllegaLDate,BillNo"<%if trim(request("Sys_Order"))="IllegaLDate,BillNo" then response.write " selected"%>>違規日期</option>
									<option value="Rule1,Driver,IllegaLDate,BillNo"<%if trim(request("Sys_Order"))="Rule1,Driver,IllegaLDate,BillNo" then response.write " selected"%>>法條,違規人,違規日期</option>
									<option value="Driver,IllegaLDate,Rule1,BillNo"<%if trim(request("Sys_Order"))="Driver,IllegaLDate,Rule1,BillNo" then response.write " selected"%>>違規人,違規日期,法條</option>
									<%if showCreditor then%>
										<option value="PetitionDate,Driver,BillNo"<%if trim(request("Sys_Order"))="PetitionDate,Driver,BillNo" then response.write " selected"%>>取得債權日,違規人,單號</option>
									<%end if%>
									
								</select>
							</td>
						</tr>
						<tr>
							<td>違規日期</td>
							<td nowrap>
								<input name="IllegalDate1" class="btn1" type="text" value="<%=request("IllegalDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('IllegalDate1');">
							~
								<input name="IllegalDate2" class="btn1" type="text" value="<%=request("IllegalDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('IllegalDate2');">
							</td>
							<td>應到案日期</td>
							<td nowrap>
								<input name="DeallIneDate1" class="btn1" type="text" value="<%=request("DeallIneDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('DeallIneDate1');">
							~
								<input name="DeallIneDate2" class="btn1" type="text" value="<%=request("DeallIneDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('DeallIneDate2');">
							</td>
						</tr>
						<tr>
							<td>填單日期</td>
							<td nowrap>
								<input name="BillFillDate1" class="btn1" type="text" value="<%=request("BillFillDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('BillFillDate1');">
							~
								<input name="BillFillDate2" class="btn1" type="text" value="<%=request("BillFillDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('BillFillDate2');">
							</td>
							<td>再次移送日</td>
							<td nowrap>
								<input name="Sys_SendDetailDate1" class="btn1" type="text" value="<%=request("Sys_SendDetailDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('Sys_SendDetailDate1');">
							~
								<input name="Sys_SendDetailDate2" class="btn1" type="text" value="<%=request("Sys_SendDetailDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('Sys_SendDetailDate2');">
							</td>						
							<td>債權狀況</td>
							<td>
								<select name="sys_CreditorTypeID">

									<option value="-1"<%if trim(Request("sys_CreditorTypeID"))="-1" then response.write " Selected"%>>未申請債權</option>

									<option value=""<%if trim(Request("sys_CreditorTypeID"))="" then response.write " Selected"%>>全部</option>

									<option value="0','1"<%if trim(Request("sys_CreditorTypeID"))="0','1" then response.write " Selected"%>>已申請債權</option>

									<option value="0"<%if trim(Request("sys_CreditorTypeID"))="0" then response.write " Selected"%>>清償中</option>

									<option value="1"<%if trim(Request("sys_CreditorTypeID"))="1" then response.write " Selected"%>>無個人財產</option>

								</select>
							</td>
						</tr>
						<tr>
							<td>建檔日期</td>
							<td nowrap>
								<input name="RecordDate1" class="btn1" type="text" value="<%=request("RecordDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('RecordDate1');">
							~
								<input name="RecordDate2" class="btn1" type="text" value="<%=request("RecordDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('RecordDate2');">
							</td>
							<td>債權取得日</td>
							<td nowrap>
								<input name="Sys_PetitionDate1" class="btn1" type="text" value="<%=request("Sys_PetitionDate1")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('Sys_PetitionDate1');">
							~
								<input name="Sys_PetitionDate2" class="btn1" type="text" value="<%=request("Sys_PetitionDate2")%>" size="5" maxlength="8" onkeyup="value=value.replace(/[^\d]/g,'')">
								<input type="button" name="datestr" value="..." class="btn3" style="width:20px;height:20px;" onclick="OpenWindow('Sys_PetitionDate2');">
							</td>
							
							<td colspan=2>								
								<input type="submit" name="btnSelt" value="查詢" class="btn3" style="width:60px;height:25px;" onclick="funSelt();">
								<input type="button" name="btnCls" value="清除" class="btn3" style="width:60px;height:25px;" onClick="location='PasserBaseQry_ReSend.asp'">
							
								<input type="button" name="btnSelt" class="btn3" style="width:85px;height:25px;" value="資料回復" onclick="funRecallData();" <%
								'1:查詢 ,2:新增 ,3:修改 ,4:刪除
								if CheckPermission(224,3)=false then
									response.write "disabled"
								end if
								%>>
								&nbsp; &nbsp;
								<input type="button" name="btnSelt" value="整批戶籍地址更正" onclick="funUpdAddress();" >
							</td>
							<!--
							<td>法條代碼</td>
							<td nowrap>
								<input name="Sys_Rule" class="btn1" type="text" value="<%=request("Sys_Rule")%>" size="9">
								
								&nbsp;&nbsp;

								<input type="submit" name="btnSelt" value="查詢" class="btn3" style="width:60px;height:25px;" onclick="funSelt();">
								<input type="button" name="btnCls" value="清除" class="btn3" style="width:60px;height:25px;" onClick="location='PasserBaseQry_Jude.asp'">
							
								<input type="button" name="btnSelt" class="btn3" style="width:80px;height:20px;" value="資料回復" onclick="funRecallData();" <%
								'1:查詢 ,2:新增 ,3:修改 ,4:刪除
								if CheckPermission(224,3)=false then
									response.write "disabled"
								end if
								%>>
								&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
								<input type="button" name="btnSelt" value="整批戶籍地址更正" onclick="funUpdAddress();" >
							</td>
							-->
						</tr>
					</td>
				</tr>
			</table>
		</td> 
	</tr>
	<tr>
		<td bgcolor="#FFCC33" class="style3">
			舉發單紀錄列表
			<img src="space.gif" width="56" height="8">
			每頁
			<select name="sys_MoveCnt" onchange="repage();">
				<option value="0"<%if trim(request("sys_MoveCnt"))="0" then response.write " Selected"%>>10</option>
				<option value="10"<%if trim(request("sys_MoveCnt"))="10" then response.write " Selected"%>>20</option>
				<option value="20"<%if trim(request("sys_MoveCnt"))="20" then response.write " Selected"%>>30</option>
				<option value="30"<%if trim(request("sys_MoveCnt"))="30" then response.write " Selected"%>>40</option>
				<option value="40"<%if trim(request("sys_MoveCnt"))="40" then response.write " Selected"%>>50</option>
				<option value="50"<%if trim(request("sys_MoveCnt"))="50" then response.write " Selected"%>>60</option>
				<option value="60"<%if trim(request("sys_MoveCnt"))="60" then response.write " Selected"%>>70</option>
				<option value="70"<%if trim(request("sys_MoveCnt"))="70" then response.write " Selected"%>>80</option>
				<option value="80"<%if trim(request("sys_MoveCnt"))="80" then response.write " Selected"%>>90</option>
				<option value="90"<%if trim(request("sys_MoveCnt"))="90" then response.write " Selected"%>>100</option>
			</select>
			筆<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 )</strong>
			<font color="red"><B>操作項目若不勾選，即為全部處理。</B></font>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th class="font10" nowrap>操作<br>項目</th>
					<th class="font10" nowrap>序號</th>
					<th class="font10" nowrap>違規日</th>
					<th class="font10" nowrap>舉發單號</th>
					<th class="font10" nowrap>舉發人</th>
					<th class="font10" nowrap>違規人</th>
					<th class="font10" nowrap>法條</th>
					<th class="font10" nowrap>移送日</th>
					<th class="font10" nowrap>確定日</th>
					<th class="font10" nowrap>再移送日</th>
					<th class="font10" nowrap>債權日</th>
					<th class="font10" nowrap>繳費日</th>
					<th class="font10" nowrap>已繳<br>金額</th>
					<th class="font10" nowrap>結案<br>狀態</th>
					<th class="font10">操作</th>
				</tr>
				<%
			if DB_Display="show" then
				set rsfound=conn.execute(strSQLTemp)
				if Trim(request("DB_Move"))="" then
					DBcnt=0
				else
					DBcnt=request("DB_Move")
				end if
				if Not rsfound.eof then rsfound.move DBcnt
				for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
					if rsfound.eof then exit for
					response.write "<tr align='center' bgcolor='#FFFFFF'"
					lightbarstyle 0
					response.write ">"
					response.write "<td class=""font10""><input class=""btn1"" type=""checkbox"" name=""chkSend"" value="""&trim(rsfound("Sn"))&""" onclick=funChkSend();></td>"
					response.write "<td class=""font10"">"&trim(rsfound("DoubleCheckStatus"))&"</td>"
					response.write "<td class=""font10"">"&gInitDT(trim(rsfound("IllegalDate")))&"</td>"
					response.write "<td class=""font10"">"&trim(rsfound("BillNo"))&"</td>"
					response.write "<td class=""font10"">"&trim(rsfound("BillMem1"))&"</td>"
					response.write "<td class=""font10"">"&trim(rsfound("Driver"))&"</td>"
					response.write "<td class=""font10"">"&trim(rsfound("Rule1"))&"</td>"
					
					SENDDATEtmp=""
					if showCreditor then
						strSD="select min(SendDate) as SendDate from PasserSendDetail where BillSn="&Trim(rsfound("sn"))&" and SendDate is not null"
						Set rsSD=conn.execute(strSD)
						If Not rsSD.eof then	
							SENDDATEtmp=Trim(rsSD("SendDate"))
						End If
						rsSD.close
						Set rsSD=Nothing 
					end if

					If SENDDATEtmp="" Or IsNull(SENDDATEtmp) Then
						strSD2="select * from PasserSend where BillSn="&Trim(rsfound("sn"))
						Set rsSD2=conn.execute(strSD2)
						If Not rsSD2.eof Then
							SENDDATEtmp=Trim(rsSD2("SendDate"))
						End If
						rsSD2.close
						Set rsSD2=Nothing 
					End If 
					response.write "<td class=""font10"">"&trim(gInitDT(SENDDATEtmp))&"</td>"

					response.write "<td class=""font10"">"&gInitDT(trim(rsfound("MakeSureDate")))&"</td>"

					strSD="select max(SendDate) as SendDate from PasserSendDetail where BillSn="&Trim(rsfound("sn"))&" and SendDate is not null"
					Set rsSD=conn.execute(strSD)
					If Not rsSD.eof then	

						If SENDDATEtmp=Trim(rsSD("SendDate")) Then

							SENDDATEtmp=""
						else

							SENDDATEtmp=Trim(rsSD("SendDate"))
						End if 
					else

						SENDDATEtmp=""
					End If
					rsSD.close
					Set rsSD=Nothing 
					response.write "<td class=""font10"">"&trim(gInitDT(SENDDATEtmp))&"</td>"

					response.write "<td class=""font10"">"&trim(gInitDT(rsfound("PetitionDate")))&"</td>"

					'------------------------------------------------------------------------------------------
					response.write "<td class=""font10"">"&trim(gInitDT(rsfound("PayDate")))&"</td>"

					Sys_Payamount=0
					strSQL="select sum(Payamount) as Sys_Payamount from PasserPay where BillSN="&rsfound("SN")&" and BillNo='"&rsfound("BillNo")&"'"
					set rspay=conn.execute(strSQL)
					if not rspay.eof then Sys_Payamount=rspay("Sys_Payamount")
					rspay.close

					response.write "<td class=""font10"">"&Sys_Payamount&"</td>"
					response.write "<td class=""font10"">"
					if trim(rsfound("BILLSTATUS"))="9" then
						response.write "結案."
					end if
					response.write "</td>"

%>					<td class="font10" nowrap align="left">
							
						<input type='button' value='債權' class="btn3" style="width:60px;height:25px;" onclick='window.open("PasserCreditor.asp?BillSN=<%=trim(rsfound("SN"))%>","WebPage4","left=0,top=0,location=0,width=900,height=575,resizable=yes,scrollbars=yes")'<%If Ifnull(rsfound("SENDDATE")) Then Response.Write " disabled"%>>
						
						<input type="button" name="Update" value="詳細" class="btn3" style="width:60px;height:25px;" onclick='window.open("../Query/ViewBillBaseData_people.asp?BillSn=<%=trim(rsfound("SN"))%>","WebPage1","left=0,top=0,location=0,width=850,height=700,status=yes,resizable=yes,scrollbars=yes")'>
					</td>
<%
					response.write "</tr>"
				rsfound.MoveNext
				next
				rsfound.close
				set rsfound=nothing
			end if
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#FFDD77" align="center" nowrap>
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);" class="btn3" style="width:70px;height:30px;font-size:16px;">
			<span class="style2"><%=fix(Cint(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(Cint(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);" class="btn3" style="width:70px;height:30px;font-size:16px;">
			<img src="space.gif" width="18" height="8">

			<input type="button" name="btnExecel" class="btn3" style="width:120px;height:30px;font-size:14px;" value="批次債權移送" onclick="funSendBatTwo_chromat();">

			<input type="button" name="btnExecel" value="轉換成Excel" class="btn3" style="width:120px;height:30px;font-size:16px;" onclick="funchgExecel();">

			
		</td>
	</tr>

</table>
<input type="Hidden" name="chkSend" value="">
<input type="Hidden" name="DB_Selt" value="<%=trim(DB_Selt)%>">
<input type="Hidden" name="DB_KindSelt" value="<%=trim(DB_KindSelt)%>">
<input type="Hidden" name="DB_Display" value="<%=trim(DB_Display)%>">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="FromILLEGALDATE" value="<%=trim(request("ILLEGALDATE"))%>">
<input type="Hidden" name="TOILLEGALDATE" value="<%=trim(request("ILLEGALDATE1"))%>">
<input type="Hidden" name="orderstr" value="<%=orderstr%>">
<input type="Hidden" name="printStyle" value="">
<input type="Hidden" name="Sys_PasserNotify" value="">
<input type="Hidden" name="Sys_PasserSign" value="">
<input type="Hidden" name="Sys_PasserJude" value="">
<input type="Hidden" name="Sys_PasserJude_Label" value="">
<input type="Hidden" name="Sys_PasserLabel_miaoli" value="">
<input type="Hidden" name="Sys_PasserDeliver" value="">
<input type="Hidden" name="Sys_PasserSend" value="">
<input type="Hidden" name="Sys_PasserJudeSend" value="">
<input type="Hidden" name="Sys_PasserUrge" value="">

<input type="Hidden" name="MailMoneyValue" value="">
<input type="Hidden" name="Session_JudeName" value="">
<input type="Hidden" name="BillUrge" value="">
<input type="Hidden" name="Sys_SendBillSN" value="<%=Sys_SendBillSN%>">
<input type="Hidden" name="hd_BillSN" value="<%=BillSN%>">
<input type="Hidden" name="hd_BillNo" value="<%=BillNo%>">

</form>
</body>
</html>

<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
var PasserWin;
var strBillSN=myForm.Sys_SendBillSN.value;
var ck_Sn=strBillSN.split(',');
<%'response.write "UnitMan('Sys_BillUnitID','Sys_BillMem','"&request("Sys_BillMem")&"');"%>
var Sys_City="<%=sys_City%>";

function funSelt(){
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.DB_Display.value='show';
	myForm.Sys_SendBillSN.value='';
	myForm.DB_KindSelt.value='';
	myForm.submit();
}
function funLoadSend(){
	if(document.getElementsByName("chkSend").length>0){
		for(i=0;i<=ck_Sn.length-1;i++){
			for(j=0;j<=myForm.chkSend.length-1;j++){
				if(ck_Sn[i]==myForm.chkSend[j].value){
					myForm.chkSend[j].checked=true;
				}
			}
		}
	}
}
function funChkSend(){
	var tempSend='';
	var chked=false;
	for(j=0;j<=myForm.chkSend.length-2;j++){
		if(myForm.chkSend[j].checked){
			chked=false
			for(i=0;i<=ck_Sn.length-1;i++){
				if(myForm.chkSend[j].value==ck_Sn[i]){
					chked=true;
				}
			}
			if(!chked){
				if(strBillSN!=''){
					strBillSN=strBillSN+',';
				}
				strBillSN=strBillSN+myForm.chkSend[j].value;
			}
		}
	}

	myForm.Sys_SendBillSN.value=strBillSN;
	ck_Sn=strBillSN.split(',');
	for(i=0;i<=ck_Sn.length-1;i++){
		chked=false
		for(j=0;j<=myForm.chkSend.length-1;j++){
			if(!myForm.chkSend[j].checked){
				if(myForm.chkSend[j].value==ck_Sn[i]){
					chked=true;
				}
			}
		}
		
		if(!chked){
			if(tempSend!=''){
				tempSend=tempSend+',';
			}
			tempSend=tempSend+ck_Sn[i];
		}
	}
	myForm.Sys_SendBillSN.value=tempSend;
	strBillSN=myForm.Sys_SendBillSN.value;
	ck_Sn=strBillSN.split(',');
}

function funPrintPeopleList_Stop(){
	UrlStr="../Query/PrintPeopleDataList_Stop.asp";
	myForm.action=UrlStr;
	myForm.target="PrintPeople";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPasserPay_Stop(){
	UrlStr="PasePayDataList_Stop.asp";
	myForm.action=UrlStr;
	myForm.target="PrintPeople";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funPasserSend_Stop(){
	UrlStr="PasserSendDetail_Execel.asp";
	myForm.action=UrlStr;
	myForm.target="PrintPeople";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPasserNoPay_Stop(){
	UrlStr="PaseNoPayDataList_Stop.asp";
	myForm.action=UrlStr;
	myForm.target="PrintPeople";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPayReportMonth(){
	var error=0;
	if(myForm.PayDate1.value!=""){
		if(!dateCheck(myForm.PayDate1.value)){
			error=1;
			alert("付費日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入付費日!!");
	}
	if (error==0){
		if(myForm.PayDate2.value!=""){
			if(!dateCheck(myForm.PayDate2.value)){
				error=1;
				alert("付費日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入付費日!!");
		}
	}
	if (error==0){
		myForm.action="ReportMonthPay.asp";
		myForm.target="PayReportMonth";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
		
}

function funPayReportMonth_Unit(){
	var error=0;
	if(myForm.PayDate1.value!=""){
		if(!dateCheck(myForm.PayDate1.value)){
			error=1;
			alert("付費日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入付費日!!");
	}
	if (error==0){
		if(myForm.PayDate2.value!=""){
			if(!dateCheck(myForm.PayDate2.value)){
				error=1;
				alert("付費日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入付費日!!");
		}
	}
	if (error==0){
		myForm.action="ReportMonthPay_Unit.asp";
		myForm.target="PayReportMonth";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}
		
}

function funPayReport(){

	myForm.action="../Report/Report0110.asp";
	myForm.target="funPayReport";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPayReport2(){

	myForm.action="../Report/Report0113.asp";
	myForm.target="funPayReport2";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPayReport3(){

	myForm.action="../Report/Report0114.asp";
	myForm.target="funPayReport3";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funSltUrgeDate(){
	var error=0;
	/*if(myForm.Sys_sltUrgeDate1.value!=""){
		if(!dateCheck(myForm.Sys_sltUrgeDate1.value)){
			error=1;
			alert("催告日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入催告日!!");
	}
	if (error==0){
		if(myForm.Sys_sltUrgeDate2.value!=""){
			if(!dateCheck(myForm.Sys_sltUrgeDate2.value)){
				error=1;
				alert("催告日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入催告日!!");
		}
	}*/
	if (error==0){
		myForm.DB_Move.value=0;
		myForm.DB_Selt.value="SltUrgeDate";
		myForm.DB_Display.value='show';
		myForm.submit();
	}
		
}

function funSltJudeDate(){
	var error=0;
	/*if(myForm.Sys_sltJudeDate1.value!=""){
		if(!dateCheck(myForm.Sys_sltJudeDate1.value)){
			error=1;
			alert("裁決日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入裁決日!!");
	}
	if (error==0){
		if(myForm.Sys_sltJudeDate2.value!=""){
			if(!dateCheck(myForm.Sys_sltJudeDate2.value)){
				error=1;
				alert("裁決日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入裁決日!!");
		}
	}*/
	if (error==0){
		myForm.DB_Move.value=0;
		myForm.DB_Selt.value="SltJudeDate";
		myForm.DB_Display.value='show';
		myForm.submit();
	}
		
}

function funSltSendDate(){
	var error=0;
	/*if(myForm.Sys_sltSendDate1.value!=""){
		if(!dateCheck(myForm.Sys_sltSendDate1.value)){
			error=1;
			alert("移送日輸入不正確!!");
		}
	}else{
		error=1;
		alert("請填入移送日!!");
	}
	if (error==0){
		if(myForm.Sys_sltSendDate2.value!=""){
			if(!dateCheck(myForm.Sys_sltSendDate2.value)){
				error=1;
				alert("移送日輸入不正確!!");
			}
		}else{
			error=1;
			alert("請填入移送日!!");
		}
	}*/
	if (error==0){
		myForm.DB_Move.value=0;
		myForm.DB_Selt.value="SltSendDate";
		myForm.DB_Display.value='show';
		myForm.submit();
	}
		
}

function funUrgeDateSelt(){
	var error=0;
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.DB_KindSelt.value="UrgeDateSelt";
	myForm.DB_Display.value='show';
	myForm.submit();
}
function funJudeDateSelt(){
	var error=0;
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.DB_KindSelt.value="JudeDateSelt";
	myForm.DB_Display.value='show';
	myForm.submit();
}
function funSendDateSelt(){
	var error=0;
	myForm.hd_BillSN.value='';
	myForm.hd_BillNo.value='';
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.DB_KindSelt.value="SendDateSelt";
	myForm.DB_Display.value='show';
	myForm.submit();
}
function funPrintStyle(){

		UrlStr="PasserSendStyle.asp";
		newWin(UrlStr,"inputWin",500,500,50,10,"yes","no","yes","no");
		myForm.action="PasserSendStyle.asp";
		myForm.target="inputWin";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funsubmit(){
	PasserWin.close();
	if(myForm.printStyle.value=='0'){
		UrlStr="BillPrints.asp";
	}else{
		UrlStr="BillPrints_a4.asp";
	}
	newWin(UrlStr,"JudeBat",920,600,50,10,"yes","yes","yes","no");
	myForm.action=UrlStr;
	myForm.target="JudeBat";
	myForm.submit();
	myForm.action="";
	myForm.target="";
	setTimeout('funchgprint()',2000);
	
}
function funchgprint(){
	PasserWin.DP();
}

function funUrgeList(){
	UrlStr="BillPrints_legal.asp";
	newWin(UrlStr,"UrgeBat",920,600,50,10,"yes","yes","yes","no");
	myForm.action=UrlStr;
	myForm.target="UrgeBat";
	myForm.submit();
	myForm.action="";
	myForm.target="";
	setTimeout('funchgprint()',2000);
	
}

function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}

function UnitMan(UnitName,UnitMem,tmpMemID){
	if(document.all[UnitName].value!=''){
		runServerScript("UnitAdd.asp?UnitID="+document.all[UnitName].value+"&UnitMem="+UnitMem+"&MemberID="+tmpMemID);
	}else{
		document.all[UnitMem].options[1]=new Option('請選取','');
		document.all[UnitMem].length=1;
	}
}

function funDbMove(MoveCnt){
	if (eval(MoveCnt)>0){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			
			myForm.hd_BillSN.value='';
			myForm.hd_BillNo.value='';
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else{
		if (eval(myForm.DB_Move.value)>0){

			myForm.hd_BillSN.value='';
			myForm.hd_BillNo.value='';
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	PasserWin=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	PasserWin.focus();
	return win;
}
function funchgExecel(){
	if(myForm.DB_Display.value=="show"){
		//newWin("","PasserBaseQry_Execel",980,550,0,0,"yes","yes","yes","no");
		UrlStr="PasserBaseQry_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="PasserBaseQry_Execel";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funJudeExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseJudeList_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funSendExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseSendList_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funJudeListExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="JudeList_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funSendCloseExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseSendCloseList_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funPasserMailMoney(){
	if(myForm.DB_Display.value=="show"){
		newWin("PasserMailMoneyList_Select.asp","PasserMailMoneyList",400,200,50,10,"yes","yes","yes","no");
	}else{
		alert("請先進行查詢");
	}
}
function funPasserReportList(){
	UrlStr="PasserReportList.asp";
	myForm.action=UrlStr;
	myForm.target="PasserReportList";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funUcloseExecel(){
	if(myForm.DB_Display.value=="show"){
		myForm.action="PasserBaseSendUCloseList_Execel.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funPasser_Reserve(){
	if(myForm.DB_Display.value=="show"){
		if(Sys_City=='台南市'){
			myForm.action="PasserBaseReserveList_TaiNaNCity.asp";
		}else{
			myForm.action="PasserBaseReserveList.asp";
		}

		//myForm.action="PasserBaseReserveList.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funUrgeJudeExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseUrgeJudeList_Execel.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funUnitListExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseUnitList_Execel.asp"
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funNoPayUnitListExecel(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserBaseNotPayUnitList_Execel.asp"
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funCountryBat(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserCountryList.asp";
		myForm.action=UrlStr;
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funCreditorStstusList(){
    if(myForm.DB_Display.value=="show"){
		
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserCreditorStstusList.asp"
		myForm.action=UrlStr;
		myForm.target="CreditorStstusList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
    }else{
        alert("請先進行查詢");
    }
}
function funPasserSendCreditorList(){
		
	//UrlStr="PasserArriveList.asp";
	UrlStr="PasserSendCreditorList.asp"
	myForm.action=UrlStr;
	myForm.target="PasserSendCreditorList";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funPasserSendCreditorTwoList(){
		
	//UrlStr="PasserArriveList.asp";
	UrlStr="PasserSendCreditorTwoList.asp"
	myForm.action=UrlStr;
	myForm.target="PasserSendCreditorTwoList";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funCreditorList(){
    if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserCreditorNotPayList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserCreditorNotPayList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
    }else{
        alert("請先進行查詢");
    }
}

function funPassersEndArrivedList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="NotArrivedList.asp"
		myForm.action=UrlStr;
		myForm.target="NotArrivedList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funNoSendList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="NoSendList.asp"
		myForm.action=UrlStr;
		myForm.target="NoSendList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funPasserSendnotCredit(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserSendnotCredit.asp"
		myForm.action=UrlStr;
		myForm.target="PasserSendnotCredit";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funPasserSendtwoList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserSendtwoList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserSendtwoList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funPasserInventoryList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserInventoryList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserInventoryList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funPasserPayDetail_YiLanList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserPayDetail_YiLanList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserPayDetail_YiLanList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funPasserNoPayDetail_YiLanList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserNoPayDetail_YalinList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserNoPayDetail_YalinList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funNoEffectsList(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="PasserNoEffectsList.asp"
		myForm.action=UrlStr;
		myForm.target="PasserArriveList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funGovArriveBat(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="funGovSendList_Exce.asp"
		myForm.action=UrlStr;
		myForm.target="PasserArriveList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funSendArriveBat(){
	if(myForm.DB_Display.value=="show"){
		//UrlStr="PasserArriveList.asp";
		UrlStr="funSendSendList_Exce.asp"
		myForm.action=UrlStr;
		myForm.target="PasserArriveList";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}
function funJudeBat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserJudeBat.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}
function funJudeBat_chromat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserJudeBat_chromat.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}
function funSendBat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserSendDetailBat.asp";
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funSendBat_chromat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserSendDetailBat_chromat.asp";
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funSendBatTwo_chromat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;
		
		myForm.action="PasserSendDetailBatTwo_chromat.asp";
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funUrgeBat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserUrgeBat.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funUrgeBat_chromat(){
		myForm.Sys_SendBillSN.value=myForm.Sys_SendBillSN.value;
		myForm.hd_BillSN.value=myForm.hd_BillSN.value;
		myForm.hd_BillNo.value=myForm.hd_BillNo.value;

		myForm.action="PasserUrgeBat_chromat.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}

function funJudeList(){
		//UrlStr="PasserJudeBatList.asp";
		//newWin(UrlStr,"JudeBat",920,600,50,10,"yes","no","yes","no");
		myForm.action="PasserJudeBatList.asp";
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.Sys_PasserNotify.value="";
		myForm.Sys_PasserSign.value="";
    	myForm.Sys_PasserJude.value="";
		myForm.Sys_PasserDeliver.value="";
		myForm.Sys_PasserSend.value="";
		myForm.Sys_PasserJudeSend.value="";
		myForm.Sys_PasserUrge.value="";
		//myForm.submit();
}

function funJudeList_chromat(){
		//UrlStr="PasserJudeBatList.asp";
		//newWin(UrlStr,"JudeBat",920,600,50,10,"yes","no","yes","no");
		<%if sys_City="台中市" then
			Response.Write "myForm.action=""PasserJudeBatList_chromat.asp"";"
		elseif sys_City="台中縣" then
			Response.Write "myForm.action=""PasserJudeBatListTaiChung_chromat.asp"";"
		end if%>
		myForm.target="JudeBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.Sys_PasserNotify.value="";
		myForm.Sys_PasserSign.value="";
    	myForm.Sys_PasserJude.value="";
		myForm.Sys_PasserDeliver.value="";
		myForm.Sys_PasserSend.value="";
		myForm.Sys_PasserJudeSend.value="";
		myForm.Sys_PasserUrge.value="";
		//myForm.submit();
}

function funSendList(){
		//UrlStr="PasserJudeBatList.asp";
		//newWin(UrlStr,"JudeBat",920,600,50,10,"yes","no","yes","no");
		myForm.action="PasserSendBatList.asp";
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.Sys_PasserNotify.value="";
		myForm.Sys_PasserSign.value="";
    	myForm.Sys_PasserJude.value="";
		myForm.Sys_PasserDeliver.value="";
		myForm.Sys_PasserSend.value="";
		myForm.Sys_PasserJudeSend.value="";
		myForm.Sys_PasserUrge.value="";
		//myForm.submit();
}

function funSendList_chromat(){
		//UrlStr="PasserJudeBatList.asp";
		//newWin(UrlStr,"JudeBat",920,600,50,10,"yes","no","yes","no");
		<%if sys_City="台中市" then
			Response.Write "myForm.action=""PasserSendBatList_chromat.asp"";"
		elseif sys_City="台中縣" then
			Response.Write "myForm.action=""PasserSendBatListTaiChung_chromat.asp"";"
		end if%>
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.Sys_PasserNotify.value="";
		myForm.Sys_PasserSign.value="";
    	myForm.Sys_PasserJude.value="";
		myForm.Sys_PasserDeliver.value="";
		myForm.Sys_PasserSend.value="";
		myForm.Sys_PasserJudeSend.value="";
		myForm.Sys_PasserUrge.value="";
		//myForm.submit();
}

function funSendListTwo_chromat(SendType){
		myForm.Sys_PasserNotify.value=SendType;
		myForm.action="PasserSendBatListTwo_chromat.asp";
		myForm.target="SendBat";
		myForm.submit();
		myForm.action="";
		myForm.target="";
		myForm.Sys_PasserNotify.value="";
		myForm.Sys_PasserSign.value="";
    	myForm.Sys_PasserJude.value="";
		myForm.Sys_PasserDeliver.value="";
		myForm.Sys_PasserSend.value="";
		myForm.Sys_PasserJudeSend.value="";
		myForm.Sys_PasserUrge.value="";
		//myForm.submit();
}

function funRecallData(){
	if(myForm.DB_Display.value=="show"){
		UrlStr="PasserRecallData.asp";
		myForm.action=UrlStr;
		myForm.target="ReCall";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢");
	}
}

function funUpdAddress(){
		UrlStr="PasserUpdAddress.asp";
		myForm.action=UrlStr;
		myForm.target="PasserUpdAddress";
		myForm.submit();
		myForm.action="";
		myForm.target="";

}
</script>
<%
conn.close
set conn=nothing
%>