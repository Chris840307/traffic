<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strRul="select Value from Apconfigure where ID=3"
set rsRul=conn.execute(strRul)
RuleVer=trim(rsRul("Value"))
rsRul.Close

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

If Not ifnull(request("Sys_SendBillSN")) Then
	Sys_SendBillSN=request("Sys_SendBillSN")
elseif Not ifnull(request("hd_BillSN")) Then
	Sys_SendBillSN=request("hd_BillSN")
else
	Sys_SendBillSN=request("BillSN")
End if

strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
set rsUnit=conn.execute(strSQL)
Sys_UnitID=trim(rsUnit("UnitID"))
Sys_UnitTypeID=trim(rsUnit("UnitTypeID"))
Sys_UnitLevelID=trim(rsUnit("UnitLevelID"))
rsUnit.close

If Sys_UnitLevelID=1 Then
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitID&"'"
else
	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
end if
set rsUnit=conn.Execute(strSQL)
DB_UnitID=trim(rsUnit("UnitID"))
DB_UnitName=trim(rsUnit("UnitName"))
DB_UnitTel=trim(rsUnit("Tel"))
DB_ManageMemberName=trim(rsUnit("ManageMemberName"))
rsUnit.close

if request("DB_Selt")="Save" then
	Session("Ch_Name")=trim(request("Sys_ChName"))
	strSQL="Update UnitInfo set WordNum='"&trim(request("Sys_WordNum"))&"' where UnitID='"&Session("Unit_ID")&"'"
	conn.execute(strSQL)

	'strSQL="Update UnitInfo set ManageMemberName='"&trim(request("Sys_UnitChName"))&"' where UnitID='"&DB_UnitID&"'"
	'conn.execute(strSQL)

	BillSN=Split(Sys_SendBillSN,",")
	theJudeDate=gOutDT(request("Sys_JudeDate"))
	for i=0 to Ubound(BillSN)
		strSQL="select Sn,BillNo,DriverAddress from PasserBase where SN="&BillSN(i)
		set rs=conn.execute(strSQL)
		strSQL="Select * from PasserJude where BillNo='"&rs("BillNo")&"' and BillSN="&rs("Sn")
		set rsJude=conn.execute(strSQL)
		if rsJude.eof then
			strIns="insert into PasserJude(BillSN,BillNO,OpenGovNumber,JudeDate,PunishmentMainBody" &_
				",SimpleReson,ForFeit,DutyUnit,SendAddress,BIGUNITBOSSNAME,RecordStateID,RecordDate,RecordMemberID,Note)" &_
				" values("&trim(rs("Sn"))&",'"&trim(rs("BillNo"))&"'"&_
				",'"&trim(request("Sys_OPENGOVNUMBER_"&i))&"',TO_DATE('"&theJudeDate&"','YYYY/MM/DD')"&_
				",'"&trim(request("PunishmentMainBody_"&i))&"','"&trim(request("SimpleReson_"&i))&"'"&_
				","&trim(request("ForFeit"))&",'"&trim(request("Sys_DutyUnit"))&"','"&trim(rs("DriverAddress"))&"','"&trim(request("Sys_UnitChName"))&"',0,sysdate,'"&Session("User_ID")&"'" &_
				",'"&trim(request("Note_"&i))&"')" 
			conn.execute(strIns)
		else
			strUpd="update PasserJude set OpenGovNumber='"&trim(request("Sys_OPENGOVNUMBER_"&i))&"',JudeDate=TO_DATE('"&theJudeDate&"','YYYY/MM/DD'),PunishmentMainBody='"&trim(request("PunishmentMainBody_"&i))&"',SimpleReson='"&trim(request("SimpleReson_"&i))&"',DutyUnit='"&trim(request("Sys_DutyUnit"))&"',Note='"&trim(request("Note_"&i))&"' where BillSN="&trim(rs("Sn"))&" and BillNo='"&trim(rs("BillNo"))&"'"
			conn.execute(strUpd)
		end if
		rsJude.close
		rs.close
	next
	response.write "<script language=""JavaScript"">"
	response.write "window.opener.funJudeList_chromat();"
	response.write "</script>"
end if
SysWordNum=""
strSQL="select WordNum from UnitInfo where UnitID='"&Session("Unit_ID")&"'"

set rs=conn.execute(strSQL)
If Not rs.eof Then SysWordNum=trim(rs("WordNum"))
rs.close
%>
<TITLE> 裁決批次套印 </TITLE>
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
<META NAME="Author" CONTENT="">
<META NAME="Keywords" CONTENT="">
<META NAME="Description" CONTENT="">
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<form name=myForm method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle">裁決批次套印</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						<font color="Red"><B>交裁字號：</B></font>
						<input name="Sys_WordNum" type="text" class="btn1" size="10" maxlength="15" value="<%=SysWordNum%>">交裁字第
						<br>
						文號產生規則<input class="btn1" type="radio" name="Sys_JudeNo" value="1"<%if trim(request("Sys_JudeNo"))="1" then response.write " checked"%>>
						文號開頭：
						<input name="Sys_JudeSN1" type="text" class="btn1" size="10" maxlength="12" value="<%
							if trim(request("Sys_JudeSN1"))<>"" then
								response.write trim(request("Sys_JudeSN1"))
							end if
						%>">
						流水號位數
						<input name="Sys_JudeSN2" type="text" class="btn1" size="2" maxlength="5" value="<%
							if trim(request("Sys_JudeSN2"))<>"" then
								response.write trim(request("Sys_JudeSN2"))
							end if
						%>" onkeyup="value=value.replace(/[^\d]/g,'')">　
						<input class="btn1" type="radio" name="Sys_JudeNo" value="3"<%if trim(request("Sys_JudeNo"))="3" then response.write " checked"%>>
						年度 + 建檔序號　
						<input class="btn1" type="radio" name="Sys_JudeNo" value="2"<%if trim(request("Sys_JudeNo"))="2" or trim(request("Sys_JudeNo"))="" then response.write " checked"%>>
						舉發單號
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
							if trim(request("Sys_UnitChName"))<>"" then
								response.write trim(request("Sys_UnitChName"))
							else
								strSQL="select MANAGEMEMBERNAME from UnitInfo where UnitID='"&DB_UnitID&"'"
								set rsUnit=conn.execute(strSQL)
								if Not rsUnit.eof then response.write rsUnit("MANAGEMEMBERNAME")
								rsUnit.close
							end if
						%>">
						裁決日期&nbsp;<input name="Sys_JudeDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')" value="<%
							strSQL="Select max(JudeDate) JudeDate from PasserJude where BillSN in("&Sys_SendBillSN&")"
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
						<input type="button" name="btnSelt" value="裁決書套印" onclick="funPasserJudeSelt();">
						<input type="button" name="btnSelt" value="送達證書套印(A4)" onclick="funPasserDeliverSelt();">
						<input type="button" name="btnSelt" value="送達證書套印(lattle)" onclick="funTaiChungCity_DeliverSelt();">
						<input name="Submit433222" type="button" class="style3" value=" 關 閉 " onclick="self.close();">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="JudePrint.jpg"><font size=5 color="blue">裁決書套印格式說明</font></a>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">裁決列表</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<table id='fmyTable' width='100%' border='0' bgcolor='#FFFFFF'><%
				BillSN=Split(Sys_SendBillSN,",")
				BillNo=""
				for i=0 to Ubound(BillSN)
					strSQL="select Sn,BillNo,DoubleCheckStatus from PasserBase where SN="&BillSN(i)
					set rs=conn.execute(strSQL)
					response.write "<input type=""Hidden"" name=""DoubleCheckStatus"" value="""&(year(date)-1911)&right("0000"&rs("DoubleCheckStatus"),4)&""">"
					strSQL="Select * from PasserJude where BillNo='"&rs("BillNo")&"' and BillSN="&rs("Sn")
					set rsJude=conn.execute(strSQL)
					rsJudeNo="":rsPunishmentMainBody="":rsSimpleReson="":rsNote=""
					if Not rsJude.eof then
						rsJudeNo=trim(rsJude("OPENGOVNUMBER"))
						rsPunishmentMainBody=trim(rsJude("PunishmentMainBody"))
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
					if rsPunishmentMainBody<>"" then
						response.write trim(rsPunishmentMainBody)
					else
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
'						response.write "一、罰鍰新台灣"&ForFeit&"元整。『限文到十五日內繳納』。<br>二、沒入物："&fastring
						strState="select a.DealLineDate,a.rule1,b.Level1,b.Level2,b.IllegalRule from Passerbase a,law b where a.rule1=b.itemid and b.version="&RuleVer&" and a.SN="&rs("Sn")&" and a.BillNo='"&rs("BillNo")&"'"
						set rsSql=conn.execute(strState)

						if not rsSql.eof Then
							If DateDiff("d",CDate(date),trim(rsSql("DealLineDate")))>-1 Then 
							  Sys_ForFeit1=trim(rsSql("Level1"))
							Else
							  Sys_ForFeit1=trim(rsSql("Level2"))
							End if
							response.write "一、"&trim(rsSql("IllegalRule"))
							response.write ".罰鍰新臺幣"&to_Money(Sys_ForFeit1)&"元整."
							If sys_City="台南市" or sys_City="台南縣" Then response.Write("(限於接到裁決書之翌日起35日內繳納。)")
							If sys_City="屏東縣" Then response.Write "限於"&Sys_ArrivedDate(0)&"年"&Sys_ArrivedDate(1)&"月"&Sys_ArrivedDate(2)&"日前繳納。"

							strSQL="Select * from PasserConfiscate where BillSN="&rs("Sn")
							set Conf=conn.execute(strSQL)
							strConf=""
							while not conf.eof
								If not ifnull(strConf) Then strConf=strConf&"、"
								strConf=strConf&conf("Confiscate")
								conf.movenext
							wend
							conf.close
							If not ifnull(strConf) Then Response.Write "<br>二、沒入物："&strConf
						end if
						rsSql.close
						set rsSql=Nothing
					end if
					response.write "</textarea></td>"
					response.write "<td nowrap>簡要理由</td>"
					response.write "<td>"
					response.write "<textarea name=""SimpleReson_"&i&""" class=""btn1"" cols=""41"" rows=""6"">"
					if rsSimpleReson <>"" then
						response.write rsSimpleReson
					else
						strSQL="Select UnitName from UnitInfo where UnitID in(select BillUnitID from PasserBase where SN="&rs("Sn")&" and BillNo='"&rs("BillNo")&"')"
						set rsbill=conn.execute(strSQL)
						response.write "受處分人於上開時間、地點，因違反道路交通管理事件被"&trim(rsbill("UnitName"))&"所查獲移送本分局處理，依『道路交通管理處罰條例』裁決如處罰主文。"
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
		<td height="20" bgcolor="#FFDD77">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="">
<input type="Hidden" name="DoubleCheckStatus" value="00">
<input type="Hidden" name="BillSN" value="<%=Sys_SendBillSN%>">
<input type="Hidden" name="BillNo" value="<%=BillNo%>">
<input type="Hidden" name="ForFeit" value="<%=Sys_ForFeit1%>">
<input type="Hidden" name="JudeCnt" value="<%=i%>">
<input type="Hidden" name="FromILLEGALDATE" value="<%=trim(request("ILLEGALDATE"))%>">
<input type="Hidden" name="TOILLEGALDATE" value="<%=trim(request("ILLEGALDATE1"))%>">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
function funJudeSN(){
	var strSN="";
	var strCnt="";
	var space=",";

	var strTypeID=myForm.BillNo.value;
	var strJudeSN=strTypeID.split(space);

	if(myForm.Sys_JudeNo[0].checked&&myForm.Sys_JudeSN2.value!=''){
		for(i=0;i<=myForm.JudeCnt.value-1;i++){
			strSN='';
			strCnt=i+"";
			for(j=strCnt.length;j<=myForm.Sys_JudeSN2.value-1;j++){
				strSN=strSN+'0'
			}
			eval("myForm.Sys_OPENGOVNUMBER_"+i).value=myForm.Sys_JudeSN1.value+strSN+eval(i+1);
		}
	}else if(myForm.Sys_JudeNo[1].checked){
		for(i=0;i<=myForm.JudeCnt.value-1;i++){
			eval("myForm.Sys_OPENGOVNUMBER_"+i).value=myForm.DoubleCheckStatus[i].value;
		}
	}else if(myForm.Sys_JudeNo[2].checked){
		for(i=0;i<=myForm.JudeCnt.value-1;i++){
			eval("myForm.Sys_OPENGOVNUMBER_"+i).value=strJudeSN[i];
		}
	}
}
function funPasserJudeSelt(){
	if(myForm.BillSN.value!=''){
		opener.myForm.Sys_PasserJude.value="1";
		funJudeSN();
		myForm.DB_Selt.value="Save";
		myForm.submit();
	}
}
function funPasserDeliverSelt(){
	if(myForm.BillSN.value!=''){
		opener.myForm.Sys_PasserDeliver.value="2";
		funJudeSN();
		myForm.DB_Selt.value="Save";
		myForm.submit();
	}
}
function funTaiChungCity_DeliverSelt(){
	if(myForm.BillSN.value!=''){
		opener.myForm.Sys_PasserDeliver.value="3";
		funJudeSN();
		myForm.DB_Selt.value="Save";
		myForm.submit();
	}
}
</script>