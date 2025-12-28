<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<%
strRul="select Value from Apconfigure where ID=3"
set rsRul=conn.execute(strRul)
RuleVer=trim(rsRul("Value"))
rsRul.Close

thenPasserCity=""
strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	if trim(rsUInfo("value"))<>"" and not isnull(rsUInfo("value")) then
		thenPasserCity=replace(trim(rsUInfo("value")),"台","臺")
	end if 
end if 
rsUInfo.close
set rsUInfo=nothing

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"

	If sys_City="台南市" and trim(Sys_UnitID)="07A7" Then
		strSQL="select * from UnitInfo where UnitID='0707'"
	End if
	
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if

set rsUnit=conn.Execute(strSQL)
DB_UnitID=trim(rsUnit("UnitID"))
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
DB_ManageMemberName=trim(rsUnit("ManageMemberName"))
rsUnit.close

If Not ifnull(request("Sys_SendBillSN")) Then
	Sys_SendBillSN=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then
	Sys_SendBillSN=request("hd_BillSN")
else
	Sys_SendBillSN=request("BillSN")
End if

if request("DB_Selt")="Save" then
	theJudeDate=gOutDT(request("Sys_JudeDate"))
	strSQL="Update UnitInfo set WordNum='"&trim(request("Sys_WordNum"))&"' where UnitTypeid in(select UnitTypeid from UnitInfo where Unitid='"&Session("Unit_ID")&"') and UnitLevelid=2"
	conn.execute(strSQL)

	'strSQL="Update UnitInfo set SecondManagerName='"&trim(request("Sys_UnitChName"))&"' where UnitID='"&DB_UnitID&"'"
	'conn.execute(strSQL)

	session("Sys_UnitChName")=request("Sys_UnitChName")


	BillSN=Split(Sys_SendBillSN,",")
	tmp_ForFeit=Split(request("ForFeit"),",")
	tmp_ForFeit2=Split(request("ForFeit2"),",")
	theJudeDate=gOutDT(request("Sys_JudeDate"))
	for i=0 to Ubound(BillSN)
		strSQL="select Sn,BillNo,DriverAddress from PasserBase where SN="&BillSN(i)
		set rs=conn.execute(strSQL)
		strSQL="Select * from PasserJude where BillNo='"&rs("BillNo")&"' and BillSN="&rs("Sn")
		set rsJude=conn.execute(strSQL)

		sumForFeit=tmp_ForFeit(i)
		If trim(tmp_ForFeit2(i)) <>"" Then sumForFeit=sumForFeit+cdbl(tmp_ForFeit2(i))

		strPay="select sum(PayAmount) as PaySum from PasserPay where BillSN="&trim(rs("Sn"))
		set rsPay=conn.execute(strPay)
		if trim(rsPay("PaySum"))<>"" and not isnull(rsPay("PaySum")) then

			sumForFeit=sumForFeit-cdbl(rsPay("PaySum"))
		end if
		rsPay.close

		if rsJude.eof then
			strIns="insert into PasserJude(BillSN,BillNO,OpenGovNumber,JudeDate,PunishmentMainBody" &_
				",SimpleReson,ForFeit,DutyUnit,SendAddress,SubUnitSecBossName,RecordStateID,RecordDate,RecordMemberID,Note)" &_
				" values("&trim(rs("Sn"))&",'"&trim(rs("BillNo"))&"'"&_
				",'"&trim(request("Sys_OPENGOVNUMBER_"&i))&"',TO_DATE('"&theJudeDate&"','YYYY/MM/DD')"&_
				",'"&request("PunishmentMainBody_"&i)&"','"&trim(request("SimpleReson_"&i))&"'"&_
				","&trim(sumForFeit)&",'"&trim(request("Sys_DutyUnit"))&"','"&trim(rs("DriverAddress"))&"','"&trim(request("Sys_UnitChName"))&"',0,sysdate,'"&Session("User_ID")&"'" &_
				",'"&trim(request("Note_"&i))&"')" 
			conn.execute(strIns)
'		else
'			strUpd="update PasserJude set OpenGovNumber='"&trim(request("Sys_OPENGOVNUMBER_"&i))&"',PunishmentMainBody='"&request("PunishmentMainBody_"&i)&"',SubUnitSecBossName='"&trim(request("Sys_UnitChName"))&"',SimpleReson='"&trim(request("SimpleReson_"&i))&"',DutyUnit='"&trim(request("Sys_DutyUnit"))&"',Note='"&trim(request("Note_"&i))&"' where BillSN="&trim(rs("Sn"))&" and BillNo='"&trim(rs("BillNo"))&"'"
'			conn.execute(strUpd)
'
'			strSQL="Update PasserJude set JudeDate=TO_DATE('"&theJudeDate&"','YYYY/MM/DD') where BillSN="&trim(rs("Sn"))&" and BillNo='"&trim(rs("BillNo"))&"' and JudeDate is null"
'			conn.execute(strSQL)
		end if
		rsJude.close

		strSQL="Update PasserBase set ForFeit1="&trim(tmp_ForFeit(i))&" where BillNo='"&rs("BillNo")&"' and SN="&rs("Sn")
		conn.execute(strSQL)
		
		If trim(tmp_ForFeit2(i)) <>"" Then
			strSQL="Update PasserBase set ForFeit2="&trim(tmp_ForFeit2(i))&" where BillNo='"&rs("BillNo")&"' and SN="&rs("Sn")
			conn.execute(strSQL)
		end if

		rs.close
	next
	response.write "<script language=""JavaScript"">"
	response.write "window.opener.funJudeList();"
	response.write "</script>"
	Response.End
end if


If Not ifnull(request("Sys_SendBillSN")) Then

	sys_billsn=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then

	sys_billsn=request("hd_BillSN")
else

	sys_billsn=request("BillSN")
End If 

tmp_billsn=split(sys_billsn,",")

sys_billsn=""

For i = 0 to Ubound(tmp_billsn)

	If i >0 then

		If i mod 100 = 0 Then

			sys_billsn=sys_billsn&"@"
		elseif sys_billsn<>"" then

			sys_billsn=sys_billsn&","
		end If 
	end if

	sys_billsn=sys_billsn&tmp_billsn(i)

Next

tmpSQL=""

If Ubound(tmp_billsn) >= 100 Then

	sys_billsn=split(sys_billsn,"@")
	
	For i = 0 to Ubound(sys_billsn)
		
		If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
		
		tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
	Next

else

	tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

End if 

BasSQL="("&tmpSQL&") tmpPasser"

SysWordNum=""
strSQL="select WordNum from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
If Not rs.eof Then SysWordNum=trim(rs("WordNum"))
rs.close
%>
<TITLE> 裁決批次套印 </TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<form name=myForm method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#1BF5FF" class="pagetitle">裁決批次套印</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						<font color="Red"><B>交字號：</B></font>
						<input name="Sys_WordNum" type="text" class="btn1" size="10" maxlength="15" value="<%=SysWordNum%>">交字第
						<br>
						文號產生規則<input class="btn1" type="radio" name="Sys_JudeNo" value="1"<%if trim(request("Sys_JudeNo"))="1" then response.write " checked"%>>
						文號開頭：
						<input name="Sys_JudeSN1" type="text" class="btn1" size="10" maxlength="12" value="<%
							if trim(request("Sys_JudeSN1"))<>"" then
								response.write trim(request("Sys_JudeSN1"))
							end if
						%>">
						流水號
						<input name="Sys_JudeSN2" type="text" class="btn1" size="2" maxlength="5" value="<%
							if trim(request("Sys_JudeSN2"))<>"" then
								response.write trim(request("Sys_JudeSN2"))
							end if
						%>" onkeyup="value=value.replace(/[^\d]/g,'')">　
						<input class="btn1" type="radio" name="Sys_JudeNo" value="3"<%if trim(request("Sys_JudeNo"))="3" then response.write " checked"%>>
						年度 + 建檔序號　
						<input class="btn1" type="radio" name="Sys_JudeNo" value="2"<%if trim(request("Sys_JudeNo"))="2" or trim(request("Sys_JudeNo"))="" then response.write " checked"%>>
						舉發單號
						<input class="btn1" type="radio" name="Sys_JudeNo" value="4"<%if trim(request("Sys_JudeNo"))="4" then response.write " checked"%>>
						建檔序號
						<input type="button" name="btnSelt" value="確定" onclick="funJudeSN();">
						<br>
						承辦人&nbsp;
						<input name="Sys_ChName" type="text" class="btn1" size="10" maxlength="12" value="<%
							if trim(request("Sys_Chmem"))<>"" then
								response.write trim(request("Sys_ChName"))
							else
								response.write trim(Session("Ch_Name"))
							end if
						%>">
						單位主管&nbsp;
						<input name="Sys_UnitChName" type="text" class="btn1" size="10" maxlength="12" value="<%
							strSQL="Select max(SubUnitSecBossName) SubUnitSecBossName from PasserJude where Exists(select 'Y' from "&BasSQL&" where sn=PasserJude.BillSN)"
							set rsda=conn.execute(strSQL)
							if Not ifnull(trim(rsda("SubUnitSecBossName"))) then
								response.write trim(rsda("SubUnitSecBossName"))
							else
								strSQL="select ManageMemberName,SecondManagerName,UnitName from UnitInfo where UnitID='"&DB_UnitID&"'"
								set rsUnit=conn.execute(strSQL)
								sHelpUnitName=rsUnit("UnitName")
								if Not rsUnit.eof then response.write rsUnit("SecondManagerName")
								rsUnit.close
							end if
							rsda.close
						%>">
						裁決日期&nbsp;<input name="Sys_JudeDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%
							strSQL="Select max(JudeDate) JudeDate from PasserJude where Exists(select 'Y' from "&BasSQL&" where sn=PasserJude.BillSN)"
							set rsda=conn.execute(strSQL)
							if Not ifnull(trim(rsda("JudeDate"))) then
								response.write gInitDT(trim(rsda("JudeDate")))
							else
								response.write gInitDT(date)
							end if
							rsda.close
						%>">
						應到案處所&nbsp;
						<select name="Sys_DutyUnit" class="btn1">
							<option value="">請選取</option>
							<%strSQL="select UnitID,UnitName from UnitInfo"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("UnitID")&""""
								if trim(request("Sys_DutyUnit"))<>"" then
									if trim(request("Sys_DutyUnit"))=trim(rs1("UnitID")) then response.write " selected"
								else
									if trim(Session("Unit_ID"))=trim(rs1("UnitID")) then response.write " selected"
								end if
								response.write ">"&rs1("UnitName")&"</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
					</td>										
				</tr>
				<tr>
				<td>
				<font color="gray" size="2">您可以於單位管理中選擇<b> <%=sHelpUnitName%></b> 設定單位主管等基本資料，後續系統會自動帶出。</font>
				</td>
				</tr>
				<tr>
					<td>
						請勾選要產生的文件類型
						<input class="btn1" type="checkbox" name="Sys_PasserNotify" value="1">
						交辦單
						<input class="btn1" type="checkbox" name="Sys_PasserSign" value="1">
						簽辦單
						<input class="btn1" type="checkbox" name="Sys_PasserJude" value="1">
						裁決通知書
						<% If sys_City="台中市" then %>
							<input class="btn1" type="checkbox" name="Sys_Execution" value="1">
							執行單
						<% else %>
							<input type="Hidden" name="Sys_Execution" value="">
						<% end If %>
						<input class="btn1" type="checkbox" name="Sys_PasserJude_Label" value="1">
						裁決通知書(保防版)
						<input class="btn1" type="checkbox" name="Sys_PasserDeliver" value="1">
						送達證書
						<input class="btn1" type="checkbox" name="Sys_PasserJudeSend" value="1">
						寄存通知
						<input class="btn1" type="checkbox" name="Sys_PasserLabel_miaoli" value="1">
						保防標籤
						<input type="button" name="btnSelt" value="確定" onclick="funSelt();">
						<input name="Submit433222" type="button" class="style3" value=" 關 閉 " onclick="self.close();">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#1BF5FF">裁決列表</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<table id='fmyTable' width='100%' border='0' bgcolor='#FFFFFF'><%
				BillSN=Split(Sys_SendBillSN,",")
				BillNo=""
				for i=0 to Ubound(BillSN)
					strSQL="select Sn,BillNo,DoubleCheckStatus,DealLineDate from PasserBase where SN="&BillSN(i)
					set rs=conn.execute(strSQL)

					response.write "<input type=""Hidden"" name=""DoubleCheckStatus"" value="""&(year(date)-1911)&right("0000"&rs("DoubleCheckStatus"),4)&""">"

					response.write "<input type=""Hidden"" name=""Array_DoubleCheckStatus"" value="""&rs("DoubleCheckStatus")&""">"

					strSQL="Select * from PasserJude where BillNo='"&rs("BillNo")&"' and BillSN="&rs("Sn")
					set rsJude=conn.execute(strSQL)
					rsJudeNo="":rsPunishmentMainBody="":rsSimpleReson="":rsNote=""
					if Not rsJude.eof then
						rsJudeNo=trim(rsJude("OPENGOVNUMBER"))
						rsPunishmentMainBody=rsJude("PunishmentMainBody")
						rsSimpleReson=trim(rsJude("SimpleReson"))
						rsNote=trim(rsJude("Note"))
					end if
					rsJude.close

					Sys_ArrivedDate="--"
					strSQL="select ArrivedDate from PassersEndArrived where PasserSN="&trim(rs("Sn"))&" and rownum=1"
					set rsArr=conn.execute(strSQL)
					If not rsArr.eof Then
						Sys_ArrivedDate=gArrDT(DateAdd("d",20,rsArr("ArrivedDate")))
					end if
					Sys_ArrivedDate=split(Sys_ArrivedDate,"-")
					rsArr.close
	
					response.write "<tr><td>"
					response.write "單號："&rs("BillNo")
					response.write "</td><td>"
					response.write "文號"
					response.write "</td><td>"
					response.write "<input name=""Sys_OPENGOVNUMBER_"&i&""" class=""btn1"" type=""text"" size=""31"" maxlength=""30"" value="""&rsJudeNo&""">"
					response.write "</td><td nowrap>"
					response.write "備註"
					response.write "</td><td>"
					response.write "<input name=""Note_"&i&""" type=""text"" class=""btn1"" size=""40"" value="""
					if rsNote <>"" then
						response.write rsNote
					end if
					response.write """></td></tr>"
					response.write "<tr><td>&nbsp;</td>"
					response.write "<td nowrap>處罰主文</td>"
					response.write "<td>"
					response.write "<textarea name=""PunishmentMainBody_"&i&""" class=""btn1"" cols=""41"" rows=""6"">"
'					if rsPunishmentMainBody<>"" then
'						response.write trim(rsPunishmentMainBody)
'					else
'						strSql="select b.Content from BILLFASTENERDETAIL a,DCICode b where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BIllSN="&BillSN(i)
'						set rsfast=conn.execute(strsql)
'						fastring=""
'						while Not rsfast.eof
'							if trim(fastring)<>"" then fastring=fastring&","
'							fastring=fastring&rsfast("Content")
'							rsfast.movenext
'						wend
'						rsfast.close
'						strSQL="Select Rule1,RuleVer,IllegalDate,DealLineDate from PasserBase where SN="&BillSN(i)&" and BillNo='"&BillNo(i)&"'"
'						set rsSql=conn.execute(strSQL)
'						ForFeit=0
'						if trim(rsSql("Rule1"))<>"" then
'							strRule1="select * from Law where ItemID='"&trim(rsSql("Rule1"))&"' and VERSION='"&trim(rsSql("RuleVer"))&"'"
'							set rsRule1=conn.execute(strRule1)
'							if not rsRule1.eof then
'								L1ForFeit=cint(trim(rsRule1("Level1")))
'								if trim(rsRule1("Level2")="" or isnull(rsRule1("Level2"))) then
'									L2ForFeit=cint(trim(rsRule1("Level1")))
'								else
'									L2ForFeit=cint(trim(rsRule1("Level2")))
'								end if
'								L3ForFeit=cint(trim(rsRule1("Level3")))
'								L4ForFeit=cint(trim(rsRule1("Level4")))
'							end if
'							rsRule1.close
'							set rsRule1=nothing
'							if  trim(rsSql("IllegalDate")) > "2007/1/1" then
'								if trim(rsSql("DealLineDate")) > now then
'									ForFeit= L1ForFeit
'								else
'									ForFeit=L4ForFeit
'								end if
'							else
'								if datediff("d",trim(rsSql("DealLineDate")),now)=<0 then
'									ForFeit=L1ForFeit
'								elseif datediff("d",trim(rsSql("DealLineDate")),now)>0 and datediff("d",trim(rsSql("DealLineDate")),now)<=15 then
'									ForFeit=L2ForFeit
'								elseif datediff("d",trim(rsSql("DealLineDate")),now)>15 and datediff("d",trim(rsSql("DealLineDate")),now)<=30 then
'									ForFeit=L3ForFeit
'								elseif datediff("d",trim(rsSql("DealLineDate")),now)>30 then
'									ForFeit=L4ForFeit
'								end if
'							end if
'						end if
'						response.write "一、罰鍰新台灣"&ForFeit&"元整。<br>『限文到十五日內繳納』。<br>二、沒入物："&fastring
					If not ifnull(rsPunishmentMainBody) Then
						Response.Write rsPunishmentMainBody

						strState="select a.DealLineDate,a.rule1,a.rule2,b.Level1,b.Level2,b.Level3,b.Level4,b.IllegalRule from Passerbase a,law b where a.rule1=b.itemid and b.version="&RuleVer&" and a.SN="&rs("Sn")&" and a.BillNo='"&rs("BillNo")&"'"

						set rsSql=conn.execute(strState)

						Sys_ForFeit1=0:Sys_ForFeit2=0
						if not rsSql.eof Then
							if datediff("d",trim(rsSql("DealLineDate")),now)=<0 then
								Sys_ForFeit1=trim(rsSql("Level1"))
							elseif datediff("d",trim(rsSql("DealLineDate")),now)>0 and datediff("d",trim(rsSql("DealLineDate")),now)<=15 then
								Sys_ForFeit1=trim(rsSql("Level2"))
							elseif datediff("d",trim(rsSql("DealLineDate")),now)>15 and datediff("d",trim(rsSql("DealLineDate")),now)<=30 then
								Sys_ForFeit1=trim(rsSql("Level3"))
							elseif datediff("d",trim(rsSql("DealLineDate")),now)>30 then
								Sys_ForFeit1=trim(rsSql("Level4"))
							end if
						end if 
						If sys_City = "基隆市" Then Sys_ForFeit1=trim(rsSql("Level1"))

						If not ifnull(trim(rsSql("rule2"))) Then
							strSQL="select ItemID,Level1,Level2,Level3,Level4 from law where version="&RuleVer&" and itemid='"&trim(rsSql("rule2"))&"'"
							set rslaw=conn.execute(strSQL)
							If not rslaw.eof Then
								if datediff("d",trim(rsSql("DealLineDate")),now)=<0 then
									Sys_ForFeit2=trim(rslaw("Level1"))

								elseif datediff("d",trim(rsSql("DealLineDate")),now)>0 and datediff("d",trim(rsSql("DealLineDate")),now)<=15 then
									Sys_ForFeit2=trim(rslaw("Level2"))

								elseif datediff("d",trim(rsSql("DealLineDate")),now)>15 and datediff("d",trim(rsSql("DealLineDate")),now)<=30 then
									Sys_ForFeit2=trim(rslaw("Level3"))

								elseif datediff("d",trim(rsSql("DealLineDate")),now)>30 then
									Sys_ForFeit2=trim(rslaw("Level4"))

								end If 

								If sys_City = "基隆市" Then Sys_ForFeit2=trim(rsSql("Level1"))
							End if 
							rslaw.close
						End if 
						rsSql.close
					else						
						strState="select a.DealLineDate,a.rule1,a.rule2,b.Level1,b.Level2,b.Level3,b.Level4,b.IllegalRule from Passerbase a,law b where a.rule1=b.itemid and b.version="&RuleVer&" and a.SN="&rs("Sn")&" and a.BillNo='"&rs("BillNo")&"'"

						
						set rsSql=conn.execute(strState)

						Sys_ForFeit1=0:Sys_ForFeit2=0
						if not rsSql.eof Then

							if datediff("d",trim(rsSql("DealLineDate")),now)=<0 then
								Sys_ForFeit1=trim(rsSql("Level1"))
							elseif datediff("d",trim(rsSql("DealLineDate")),now)>0 and datediff("d",trim(rsSql("DealLineDate")),now)<=15 then
								Sys_ForFeit1=trim(rsSql("Level2"))
							elseif datediff("d",trim(rsSql("DealLineDate")),now)>15 and datediff("d",trim(rsSql("DealLineDate")),now)<=30 then
								Sys_ForFeit1=trim(rsSql("Level3"))
							elseif datediff("d",trim(rsSql("DealLineDate")),now)>30 then
								Sys_ForFeit1=trim(rsSql("Level4"))
							end If 

							If sys_City = "基隆市" Then Sys_ForFeit1=trim(rsSql("Level1"))

							If not ifnull(trim(rsSql("rule2"))) Then
								strSQL="select ItemID,Level1,Level2,Level3,Level4 from law where version="&RuleVer&" and itemid='"&trim(rsSql("rule2"))&"'"
								set rslaw=conn.execute(strSQL)
								If not rslaw.eof Then
									if datediff("d",trim(rsSql("DealLineDate")),now)=<0 then
										Sys_ForFeit2=trim(rslaw("Level1"))

									elseif datediff("d",trim(rsSql("DealLineDate")),now)>0 and datediff("d",trim(rsSql("DealLineDate")),now)<=15 then
										Sys_ForFeit2=trim(rslaw("Level2"))

									elseif datediff("d",trim(rsSql("DealLineDate")),now)>15 and datediff("d",trim(rsSql("DealLineDate")),now)<=30 then
										Sys_ForFeit2=trim(rslaw("Level3"))

									elseif datediff("d",trim(rsSql("DealLineDate")),now)>30 then
										Sys_ForFeit2=trim(rslaw("Level4"))

									end If 

									If sys_City = "基隆市" Then Sys_ForFeit2=trim(rsSql("Level1"))
								End if 
								rslaw.close

							End if 
							sum_ForFeit=cdbl(Sys_ForFeit1)+cdbl(Sys_ForFeit2)
							'response.write trim(rsSql("IllegalRule"))

							strPay="select sum(PayAmount) as PaySum from PasserPay where BillSN="&rs("Sn")
							set rsPay=conn.execute(strPay)
							if trim(rsPay("PaySum"))<>"" and not isnull(rsPay("PaySum")) then

								sum_ForFeit=sum_ForFeit-cdbl(rsPay("PaySum"))
							end if
							rsPay.close

							

							response.write "一、罰鍰新臺幣"&to_Money(sum_ForFeit)&"元整。"

							If sys_City = "花蓮縣" then

								response.Write chr(13)&"<br>二、限於接到裁決書之翌日起至@系統帶入裁決日期@前繳納，逾期未繳納者，依行<br>　　政執行法移送管轄地方行政執行分署強制執行。"
							else

								response.Write chr(13)&"<br>二、限於接到裁決書之翌日起30日內限期繳納，逾期未繳納者，依行政執行法<br>　　移送管轄地方行政執行分署強制執行。"
							End if 

							
							
'							If sys_City="台南市" or sys_City="台南縣" Then 
'								response.Write("罰鍰新臺幣"&to_Money(sum_ForFeit)&"元整.(限於接到裁決書之翌日起30日內繳納。逾期不繳納者，依法移送強制執行)。")	
'							else
'								response.write "一、罰鍰新臺幣"&to_Money(sum_ForFeit)&"元整。"
'							end if
'
'							If sys_City="澎湖縣" and sys_City="彰化縣" Then response.Write("(限於接到裁決書之翌日起30日內繳納。)")
'
'							If sys_City="屏東縣" Then response.Write "限於"&Sys_ArrivedDate(0)&"年"&Sys_ArrivedDate(1)&"月"&Sys_ArrivedDate(2)&"日前繳納。"
'
'							If sys_City="花蓮縣" Then Response.Write "<br>二、沒入物：無。"
'
'							If sys_City <> "台中市" and sys_City <> "花蓮縣" and sys_City<>"台南市" and sys_City<>"台南縣" and sys_City<>"苗栗縣" Then
'								response.Write chr(13)&"<br>二、限於接到裁決書之翌日起30日內限期繳納，逾期未繳納者，依行政執行法<br>　　移送管轄地方行政執行分署強制執行。"
'							end if

							
						end if
						rsSql.close
						set rsSql=Nothing
					end if
					response.write "</textarea>"
					Response.Write "<input type=""Hidden"" name=""ForFeit"" value="""&Sys_ForFeit1&""">"
					Response.Write "<input type=""Hidden"" name=""ForFeit2"" value="""&Sys_ForFeit2&""">"
					Response.Write "</td>"
					response.write "<td nowrap>簡要理由</td>"
					response.write "<td>"
					response.write "<textarea name=""SimpleReson_"&i&""" class=""btn1"" cols=""41"" rows=""6"">"
					if rsSimpleReson <>"" then
						response.write rsSimpleReson
					else
						strSQL="Select UnitName,UnitLevelID from UnitInfo where UnitID in(select BillUnitID from PasserBase where SN="&rs("Sn")&" and BillNo='"&rs("BillNo")&"')"
						set rsbill=conn.execute(strSQL)

						tmp_thenPasserCity=thenPasserCity

						If instr(rsbill("UnitName"),"鐵路")>0 Then

							tmp_thenPasserCity=""
						elseIf trim(rsbill("UnitLevelID")) <> "1" Then

							tmp_thenPasserCity="本分局"
						End if
						response.write "受處分人於上開違規時間、地點，因違反道路交通管理處罰條例，經"&tmp_thenPasserCity&replace(trim(rsbill("UnitName")),DB_UnitName,"")&"製單舉發，未依通知日期到案，依「道路交通管理處罰條例」裁決，如主文。"
						rsbill.close
					end if
					response.write "</textarea></td></tr><tr><td colspan=5><hr></td></tr>"
					If BillNo<>"" Then BillNo=BillNo&","
					BillNo=BillNo&rs("BillNo")
					rs.close
				next
			%>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td height="20" bgcolor="#1BF5FF">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="">
<input type="Hidden" name="DoubleCheckStatus" value="00">
<input type="Hidden" name="Array_DoubleCheckStatus" value="00">
<input type="Hidden" name="BillSN" value="<%=Sys_SendBillSN%>">
<input type="Hidden" name="BillNo" value="<%=BillNo%>">
<input type="Hidden" name="JudeCnt" value="<%=i%>">
<input type="Hidden" name="FromILLEGALDATE" value="<%=trim(request("ILLEGALDATE"))%>">
<input type="Hidden" name="TOILLEGALDATE" value="<%=trim(request("ILLEGALDATE1"))%>">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var sys_City="<%=sys_City%>";

function funJudeSN(){
	var strSN="";
	var strCnt="";
	var space=",";
	var strzeo="";

	var strTypeID=myForm.BillNo.value;
	var strJudeSN=strTypeID.split(space);

	if(myForm.Sys_JudeNo[0].checked&&myForm.Sys_JudeSN2.value!=''){
		strCnt=myForm.Sys_JudeSN2.value.length;
		for(i=0;i<=myForm.JudeCnt.value-1;i++){
			strzeo="";
			strSN=eval(eval(myForm.Sys_JudeSN2.value)+eval(i+1));
			for(j=strSN.toString().length;j<strCnt;j++){
				strzeo=strzeo.toString()+'0';
			}
			eval("myForm.Sys_OPENGOVNUMBER_"+i).value=myForm.Sys_JudeSN1.value+strzeo.toString()+strSN.toString();
		}
	}else if(myForm.Sys_JudeNo[1].checked){
		for(i=0;i<=myForm.JudeCnt.value-1;i++){
			eval("myForm.Sys_OPENGOVNUMBER_"+i).value=myForm.DoubleCheckStatus[i].value;
		}
	}else if(myForm.Sys_JudeNo[2].checked){
		for(i=0;i<=myForm.JudeCnt.value-1;i++){
			eval("myForm.Sys_OPENGOVNUMBER_"+i).value=strJudeSN[i];
		}
	}else if(myForm.Sys_JudeNo[3].checked){
		for(i=0;i<=myForm.JudeCnt.value-1;i++){
			eval("myForm.Sys_OPENGOVNUMBER_"+i).value=myForm.Array_DoubleCheckStatus[i].value;
		}
	}
}
function funSelt(){
	if(myForm.BillSN.value!=''){
		if(myForm.Sys_PasserNotify.checked){
			opener.myForm.Sys_PasserNotify.value="1";
		}else{
			opener.myForm.Sys_PasserNotify.value="";
		}

		if(myForm.Sys_PasserJude.checked){
			opener.myForm.Sys_PasserJude.value="1";
		}else{
			opener.myForm.Sys_PasserJude.value="";
		}		

		if(myForm.Sys_Execution.checked){
			opener.myForm.Sys_Execution.value="1";
		}else{
			opener.myForm.Sys_Execution.value="";
		}

		if(myForm.Sys_PasserJude_Label.checked){
			opener.myForm.Sys_PasserJude_Label.value="1";
		}else{
			opener.myForm.Sys_PasserJude_Label.value="";
		}

		if(myForm.Sys_PasserDeliver.checked){
			opener.myForm.Sys_PasserDeliver.value="1";
		}else{
			opener.myForm.Sys_PasserDeliver.value="";
		}

		if(myForm.Sys_PasserJudeSend.checked){
			opener.myForm.Sys_PasserJudeSend.value="1";
		}else{
			opener.myForm.Sys_PasserJudeSend.value="";
		}

		if(myForm.Sys_PasserSign.checked){
			opener.myForm.Sys_PasserSign.value="1";
		}else{
			opener.myForm.Sys_PasserSign.value="";
		}

		if(myForm.Sys_PasserLabel_miaoli.checked){
			opener.myForm.Sys_PasserLabel_miaoli.value="1";
		}else{
			opener.myForm.Sys_PasserLabel_miaoli.value="";
		}

		opener.myForm.Session_JudeName.value=myForm.Sys_ChName.value;
		if(myForm.Sys_OPENGOVNUMBER_0.value==''){funJudeSN();}
		myForm.DB_Selt.value="Save";
		myForm.submit();
	}
}
</script>