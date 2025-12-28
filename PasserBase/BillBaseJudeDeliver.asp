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
	strUInfo="select * from UnitInfo where UnitID='"&trim(rsState("DutyUnit"))&"'"
	set rsUInfo=conn.execute(strUInfo)
	if not rsUInfo.eof then
		DutyUnitName=trim(rsUInfo("UnitName"))
		DutyAddress=trim(rsUInfo("Address"))
	end if
	rsUInfo.close
	set rsUInfo=nothing
end if
rsState.close
set rsState=nothing
PrintDate=split(gArrDT(date),"-")
strSql="select * from PasserBase where SN="&BillSN(i)
set rsSql=conn.execute(strSql)
if Not rsSql.eof then
	If not ifnull(Trim(rsSql("DriverID"))) Then
		If Mid(Trim(rsSql("DriverID")),2,1)="1" Then
			Sys_Sex="男"
		elseif Mid(Trim(rsSql("DriverID")),2,1)="2" Then
			Sys_Sex="女"
		End if
	End if
end if
if rsSql.eof then response.end
strUInfo="select * from UnitInfo where UnitID='"&trim(rsSql("BillUnitID"))&"'"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	theSubUnitSecBossName=trim(rsUInfo("SecondManagerName"))
	theBigUnitBossName=trim(rsUInfo("ManageMemberName"))
	theContactTel=trim(rsUInfo("Tel"))
	theBankAccount=trim(rsUInfo("BankAccount"))
	thenBillUnitName=trim(rsUInfo("UnitName"))
end if
rsUInfo.close
set rsUInfo=nothing

strSql="select b.Content from BILLFASTENERDETAIL a,DCICode b where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BIllSN="&trim(rsSql("SN"))
set rsfast=conn.execute(strsql)
fastring=""
while Not rsfast.eof
	if trim(fastring)<>"" then fastring=fastring&","
	fastring=fastring&rsfast("Content")
	rsfast.movenext
wend
rsfast.close

%>
<table width="645" height="100%" border="0">
  <tr>
    <td height="34" colspan="2"><div align="center" class="style7"><%=thenPasserCity%>（第四分局）</div></td>
  </tr>
  <tr valign="bottom">
    <td height="29" colspan="2"><div align="center" class="style2">違反道路交通管理事件裁決書</div></td>
  </tr>
  <tr>
    <td width="312" height="25"><span class="style2">交受處份人</span></td>
    <td width="323"><div align="right" class="style2"><%=thenBillUnitName&"交裁字第"&UOpenGovNumber&"號"%></div></td>
  </tr>
</table>
<table width="645" border="1" cellpadding="4" cellspacing="1" cellspacing=0 cellpadding=0>
  <tr>
    <td width="84" height="25" nowrap><span class="style2">受處分姓名</span></td>
    <td colspan="3"><span class="style2"><%=trim(rsSql("DRIVER"))%></span></td>
    <td width="126" nowrap><span class="style2">原舉發單位通知單</span></td>
    <td width="152"><span class="style2"><%=trim(rsSql("BillNo"))%></span></td>
  </tr>
  <tr>
    <td height="25" nowrap><span class="style2">出生年月日</span></td>
    <td width="58"><span class="style2"><%=gInitDT(trim(rsSql("DriverBirth")))%>&nbsp;</span></td>
    <td width="47" height="25" nowrap><span class="style2">姓別</span></td>
    <td width="84"><span class="style2">
      <%=Sys_Sex%>
    &nbsp;</span></td>
    <td height="25" nowrap><span class="style2">身份證統一編號</span></td>
    <td><span class="style2"><%=trim(rsSql("DriverID"))%></span></td>
  </tr>
  <tr>
    <td height="25" nowrap><span class="style2">住址</span></td>
    <td colspan="3"><span class="style2"><%=trim(rsSql("DriverZip"))&trim(rsSql("DriverAddress"))%></span></td>
    <td height="25" nowrap><span class="style2">代保管物件</span></td>
    <td><span class="style2"><%=fastring%></span></td>
  </tr>
  <tr>
    <td height="25" nowrap><span class="style2">違規時間</span></td>
    <td colspan="3"><span class="style2">
      <%
		if trim(rsSql("IllegalDate"))<>"" then
			IllegalDate=split(gArrDT(rsSql("IllegalDate")),"-")
			response.write IllegalDate(0)&"年"&IllegalDate(1)&"月"&IllegalDate(2)&"日"&hour(rsSql("IllegalDate"))&"時"&minute(rsSql("IllegalDate"))&"分"
		end if%>
	</span></td>
    <td height="25" nowrap><span class="style2">違規地點</span></td>
    <td><span class="style2"><%=trim(rsSql("IllegalAddress"))%></span></td>
  </tr>
  <tr>
    <td height="25"><span class="style2">原舉發通知單<br>
    應到案日期</span></td>
    <td colspan="3"><span class="style2">
      <%
		if trim(rsSql("DealLineDate"))<>"" then
			DealLineDate=split(gArrDT(rsSql("DealLineDate")),"-")
			response.write DealLineDate(0)&"年"&DealLineDate(1)&"月"&DealLineDate(2)&"日前"
		end if%>
    &nbsp;</span></td>
    <td height="25" nowrap><span class="style2">舉發單位</span></td>
    <td> <span class="style2"><%=thenBillUnitName%> </span></td>
  </tr>
  <tr>
    <td height="83" nowrap><span class="style2">舉發違規事實</span></td>
    <td colspan="5"><span class="style2">
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
    <td height="38" nowrap><span class="style2">違反法條</span></td>
    <td colspan="5"><span class="style2">道路交通管理處罰條例<%=trim(rsSql("Rule1"))%>規定</span></td>
  </tr>
  <tr>
    <td height="74" nowrap><span class="style2">處罰主文</span></td>
    <td colspan="5"><span class="style2"><%=UPunishmentMainBody%></span></td>
  </tr>
  <tr>
    <td height="145" nowrap><span class="style2">簡要理由</span></td>
    <td colspan="5"><span class="style2"><%=USimpleReson%></span></td>
  </tr>
  <tr>
    <td height="36" nowrap><span class="style2">裁決日期</span></td>
    <td colspan="5"><span class="style2">中華民國<%=UJudeDate(0)%>年<%=UJudeDate(1)%>月<%=UJudeDate(2)%>日</span></td>
  </tr>
  <tr>
    <td height="41" nowrap><span class="style2">應到案處所</span></td>
    <td colspan="5"><span class="style2"><%=DutyUnitName%>　　　　地址：<%=DutyAddress%></span></td>
  </tr>
  <tr>
    <td height="44" nowrap><span class="style2">承辦人</span></td>
    <td colspan="3"><span class="style2"><%
	if trim(Sys_MemUnitFileName)<>"" then
		response.write "<img src=""../Member/Picture/"&Sys_MemUnitFileName&""" width=""90"" height=""30"">"
	else
		response.write MemUnitName
	end if%>&nbsp;</span></td>
    <td nowrap><span class="style2">單位主管</span></td>
    <td><span class="style2"><%
%>&nbsp;</span></td>
  </tr>
  <tr>
    <td height="168"><div align="center"><span class="style2">附<br>
    記</span></div></td>
    <td colspan="5" valign="top"><span class="style2">一、受處分人不服本裁決聲明異議者，得於收受裁決書之翌日起二十日內，以司法狀紙提交本分局，轉送管轄地方法院。<br>
      二、聲明異議之書狀應載明本裁決書之日、字號及理由。<br>
    三、上開罰鍰逾期仍不繳納者，本分局將『道路交通管理處罰條例』第九十條之二移送行政執行處強制執行。</span></td>
  </tr>
</table>
