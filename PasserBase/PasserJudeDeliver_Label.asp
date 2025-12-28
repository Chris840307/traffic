<%
strState="select * from PasserJude where BillSN="&BillSN(i)
set rsState=conn.execute(strState)
if not rsState.eof then
	UOpenGovNumber=trim(rsState("OpenGovNumber"))
	UJudeDate=split(gArrDT(rsState("JudeDate")),"-")
	UAgentName=trim(rsState("AgentName"))
	UAgentBirth=trim(rsState("AgentBirth"))
	UAgentID=trim(rsState("AgentID"))
	UAgentAddress=trim(rsState("AgentAddress"))
	UBigUnitBossName=trim(rsState("BigUnitBossName"))
	USubUnitSecBossName=trim(rsState("SubUnitSecBossName"))
	UContactTel=trim(rsState("ContactTel"))
	UForFeit=trim(rsState("ForFeit"))
	UDutyUnit=trim(rsState("DutyUnit"))
	USendAddress=trim(rsState("SendAddress"))
	UPunishmentMainBody=trim(rsState("PunishmentMainBody"))
	USimpleReson=trim(rsState("SimpleReson"))
	UNote=trim(rsState("Note"))
'	strUInfo="select * from UnitInfo where UnitID='"&trim(rsState("DutyUnit"))&"'"
'	set rsUInfo=conn.execute(strUInfo)
'	if not rsUInfo.eof then
'		DutyUnitName=trim(rsUInfo("UnitName"))
'		DutyAddress=trim(rsUInfo("Address"))
'	end if
'	rsUInfo.close
'	set rsUInfo=nothing
end if
rsState.close
set rsState=nothing
PrintDate=split(gArrDT(date),"-")

UAgentSex=""

strSql="select * from PasserBase where SN="&BillSN(i)
set rsSql=conn.execute(strSql)
if rsSql.eof then response.end

if Not rsSql.eof then

	If Trim(rsSql("DRIVERSEX")) = "1" Then

		UAgentSex="男"
	elseIf Trim(rsSql("DRIVERSEX")) = "2" Then

		UAgentSex="女"
	End if 
	
end if

strUInfo="select * from UnitInfo where UnitID='"&trim(rsSql("BillUnitID"))&"'"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	thenBillUnitName=trim(rsUInfo("UnitName"))
end if
rsUInfo.close
set rsUInfo=nothing

strSql="select confiscate from PasserConfiscate where BIllSN="&trim(rsSql("SN"))
set rsfast=conn.execute(strsql)
fastring=""
while Not rsfast.eof
	if trim(fastring)<>"" then fastring=fastring&","
	fastring=fastring&rsfast("confiscate")
	rsfast.movenext
wend
rsfast.close


chk_UnitLevelID=""
strSQL="select UnitLevelID from UnitInfo where UnitID in(select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"')"
set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then
	chk_UnitLevelID=trim(rsunit("UnitLevelID"))
End if
rsunit.close

strMem="select MANAGEMEMBERNAME,secondmanagername from UnitInfo where UnitID='"&Sys_UnitID&"'"
set rsUnit=conn.execute(strMem)
if Not rsUnit.eof then 
	If ifnull(USubUnitSecBossName) Then USubUnitSecBossName=trim(rsUnit("secondmanagername"))

	If sys_City="高雄市" or sys_City="宜蘭縣" or sys_City="台南市" or sys_City="嘉義市" then
		If not ifnull(rsUnit("secondmanagername")) Then
			MemUnitName="分局長 "&USubUnitSecBossName
		End If 
		
	elseif sys_City="台中市" Then
		If chk_UnitLevelID = "1" Then
			MemUnitName="局長 "&rsUnit("MANAGEMEMBERNAME")
		else
			MemUnitName="分局長 "&USubUnitSecBossName
		End If 

	elseif sys_City="台南市" Then
		MemUnitName=""

	Else
		MemUnitName=rsUnit("MANAGEMEMBERNAME")
	End If 
End If 
rsUnit.close

%>
<br><br>
<table width="635" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="0" colspan="2" nowrap><div align="center" class="style25"><%=thenPasserCity%><%=theUnitName%></div></td>
  </tr>
  <tr valign="bottom">
    <td height="0" colspan="2"><div align="center" class="style25">違反道路交通管理事件裁決書</div></td>
  </tr>
  <tr>
    <td width="110" height="0"><span class="style22"></span></td>
    <td><div align="right" class="style22"><%
		If sys_City="澎湖縣" Then
			Response.Write BillPageUnit&"交裁字第"&UOpenGovNumber&"號"
		else
			Response.Write BillPageUnit&"裁字第"&UOpenGovNumber&"號"
		End if		
	%></div></td>
  </tr>
</table>
<table width="665" border="1" cellspacing=0 cellpadding=0>
  <tr>
    <td width="98" height="36" nowrap><span class="style22">受處分人姓名</span></td>
    <td height="36" colspan="3"><span class="style22"><%=trim(rsSql("DRIVER"))%></span></td>
    <td width="152" height="36" nowrap><span class="style22">原舉發單位通知單</span></td>
    <td width="152" height="36"><span class="style22">第<%=trim(rsSql("BillNo"))%>號</span></td>
  </tr>
  <tr>
    <td  height="36" nowrap><span class="style22">出生年月日</span></td>
    <td width="58" height="36"><span class="style22"><%=gInitDT(trim(rsSql("DriverBirth")))%>&nbsp;</span></td>
    <td width="47"  height="36" nowrap><span class="style22">性別</span></td>
    <td width="84" height="36"><span class="style22">
      <%=UAgentSex%>
    &nbsp;</span></td>
    <td  height="36" nowrap><span class="style22">身分證統一編號</span></td>
    <td><span class="style22"><%=trim(rsSql("DriverID"))%></span></td>
  </tr>
  <tr>
    <td height="36" nowrap><span class="style22">住址</span></td>
    <td colspan="3"><span class="style22"><%=trim(rsSql("DriverZip"))&trim(rsSql("DriverAddress"))%></span>&nbsp;</td>
    <td height="36" nowrap><span class="style22">代保管物件</span></td>
    <td><span class="style22"><%=fastring%>&nbsp;</span></td>
  </tr>
  <tr>
    <td height="36" nowrap><span class="style22">違規時間</span></td>
    <td colspan="3"><span class="style22">
      <%
		if trim(rsSql("IllegalDate"))<>"" then
			IllegalDate=split(gArrDT(rsSql("IllegalDate")),"-")
			response.write IllegalDate(0)&"年"&IllegalDate(1)&"月"&IllegalDate(2)&"日"&hour(rsSql("IllegalDate"))&"時"&minute(rsSql("IllegalDate"))&"分"
		end if%>
	</span></td>
    <td height="36" nowrap><span class="style22">違規地點</span></td>
    <td><span class="style22"><%=trim(rsSql("IllegalAddress"))%></span></td>
  </tr>
  <tr>
    <td height="36"><span class="style22">原舉發通知單<br>
    應到案日期</span></td>
    <td colspan="3" height="36"><span class="style22">
      <%
		if trim(rsSql("DealLineDate"))<>"" then
			DealLineDate=split(gArrDT(rsSql("DealLineDate")),"-")
			response.write DealLineDate(0)&"年"&DealLineDate(1)&"月"&DealLineDate(2)&"日前"
		end if%>
    &nbsp;</span></td>
    <td height="25" nowrap><span class="style22">舉發單位</span></td>
    <td> <span class="style22"><%=thenBillUnitName%> </span></td>
  </tr>
  <tr>
    <td height="36" nowrap><span class="style22">舉發違規事實</span></td>
    <td colspan="5"><span class="style22">
      <%
		if trim(rsSql("Rule1"))<>"" and not isnull(rsSql("Rule1")) then
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule1"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				response.write trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if	
		if trim(rsSql("Rule2"))<>"" and not isnull(rsSql("Rule2")) then
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule2"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				response.write "<br>"&trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if	
		if trim(rsSql("Rule3"))<>"" and not isnull(rsSql("Rule3")) then
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule3"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				response.write "<br>"&trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if	
		if trim(rsSql("Rule4"))<>"" and not isnull(rsSql("Rule4")) then
			strRule1="select * from Law where ItemID='"&trim(rsSql("Rule4"))&"' and VerSion='"&trim(rsSql("RuleVer"))&"'"
			set rsRule1=conn.execute(strRule1)
			if not rsRule1.eof then
				response.write "<br>"&trim(rsRule1("IllegalRule"))
			end if
			rsRule1.close
			set rsRule1=nothing
		end if%>
	</span></td>
  </tr>
  <tr>
    <td height="38" nowrap><span class="style22">違反法條</span></td>
    <td colspan="5"><span class="style22">道路交通管理處罰條例<br><%
		Sys_Rule1=trim(rsSql("Rule1"))
		response.write "第"&left(trim(Sys_Rule1),2)&"條"
		if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
		response.write "第"&Mid(trim(Sys_Rule1),3,1)&"項"

			If cdbl(Mid(trim(Sys_Rule1),4,2)) > 0 Then

				Response.Write "第"&Mid(trim(Sys_Rule1),4,2)&"款"
			End if 

			Response.Write "規定"
		%></span>

		<%if trim(rsSql("Rule2"))<>"" and not isnull(rsSql("Rule2")) then
			Sys_Rule2=trim(rsSql("Rule2"))
			Response.Write "<br><span class=""style22"">"
			response.write "與第"&left(trim(Sys_Rule2),2)&"條"
			if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
			response.write "第"&Mid(trim(Sys_Rule2),3,1)&"項"

			If cdbl(Mid(trim(Sys_Rule2),4,2)) > 0 Then

				Response.Write "第"&Mid(trim(Sys_Rule2),4,2)&"款"
			End if 

			Response.Write "規定"
			Response.Write "</span>"
		end if%>

		</td>
  </tr>
  <tr>
    <td height="74" nowrap><span class="style22">處罰主文</span></td>
    <td colspan="5"><span class="style22"><%=UPunishmentMainBody%></span></td>
  </tr>
  <tr>
    <td nowrap><span class="style22">簡要理由</span></td>
    <td colspan="5"><span class="style22"><%=USimpleReson%></span></td>
  </tr>
  <tr>
    <td height="36" nowrap><span class="style22">裁決日期</span></td>
    <td colspan="5"><span class="style22">中華民國<%=UJudeDate(0)%>年<%=UJudeDate(1)%>月<%=UJudeDate(2)%>日</span></td>
  </tr>
  <tr>
    <td height="41" nowrap><span class="style22">應到案處所</span></td>
    <td colspan="5"><span class="style22"><%=thenPasserCity&theUnitName%>　<br>地址：<%=theUnitAddress%>　<br>電話：<%=theContactTel%></span></td>
  </tr>
  <tr>
    <td height="44" nowrap><span class="style22">機關首長&nbsp;</span></td>
	<td colspan="5" height="44" nowrap><span class="style22"><%=MemUnitName%>&nbsp;</span></td>
  </tr>
  <tr>
    <td><div align="center"><span class="style22">附<br>
    記</span></div></td>
    <td colspan="5" valign="top"><span class="style22">
	一、受處分人不服本裁決者，應以原處分機關（<%=thenPasserCity&"<br>"%>
	&nbsp;&nbsp;&nbsp;&nbsp;<%=theUnitName%>）為被告，向原告住所地、居所地、所在地、違規行為<br>
	&nbsp;&nbsp;&nbsp;&nbsp;地、或原處分機關所在地之地方法院行政訴訟庭提起訴訟；其中撤銷訴訟之<br>
	&nbsp;&nbsp;&nbsp;&nbsp;提起，應於裁決書送達後30日之不變期間內為之。<br>

    二、請依處罰主文所定期限前持本裁決書至應到案處所、郵局繳納罰鍰。<br>

    三、上開罰鍰逾30日仍不繳納者，本分局將依『行政執行法』<br>

	&nbsp;&nbsp;&nbsp;&nbsp;第二章第十一條移送行政執行分署強制執行。<br>
	<%
	if sys_City="宜蘭縣" then
		Response.Write "四、罰鍰繳納方式：1.可親自或委託他人至本分局臨櫃窗口繳納。<br>"
		Response.Write "　　　　　　　　　2.購買郵政匯票（抬頭請註明本分局全銜，<br>"	
		Response.Write "　　　　　　　　　例："&thenPasserCity&theUnitName&"）郵寄至本分局繳納。"
	
	elseif sys_City="台東縣" then
		Response.Write "四、『郵局劃撥』戶名："&theBankName&"　<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;劃撥帳號："&theBankAccount&"<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;如採郵局劃撥，應加收手續費7元。"

	else
		Response.Write "四、『郵局劃撥』戶名："&theBankName&"　<br>"
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;劃撥帳號："&theBankAccount
	end if
	%>
	</span></td>
  </tr>
</table>
<br><br>
<br><br>
<table width="635" border="0" cellpadding="0" cellspacing="0">
  <tr valign="">
    <td colspan="3"><div align="center" class="style25">違反道路交通管理事件裁決書　行政文書</div></td>
  </tr>
  <tr>
    <td colspan="3"><div class="style22"><br>　　　　寄件人：<%=theUnitAddress%>(<%=theUnitName%>)</div></td>
  </tr>
  <tr>
	<td width="100"></td>
    <td><div align="right" class="style17">地　　址：</div></td>
	<td><div align="left" class="style17"><%=trim(rsSql("DriverZip"))&"&nbsp;&nbsp;"&trim(rsSql("DriverAddress"))%></div></td>
  </tr>
  <tr>
	<td width="100"></td>
    <td><div align="right" class="style17">收&nbsp;件&nbsp;人：</div></td>
	
	<td><div align="left" class="style17"><%=trim(rsSql("Driver"))%></div></td>
  </tr>
</table>
<%rsSql.close%>