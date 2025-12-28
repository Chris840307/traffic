<%
showCreditor=false
if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then
	showCreditor=true
end If

strSql="select a.SN as BillSN,a.BillNo,a.Driver,a.DriverBirth,a.DriverZip,a.DriverID,a.DriverZip,a.DriverAddress,a.IllegalDate,a.IllegalAddress,a.DealLineDate,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,nvl(a.forfeit1,0)+nvl(a.forfeit2,0) ForFeit,b.OpenGovNumber as JudeOGN,b.AgentName as JudeAgentName,b.AgentSex as JudeAgentSex,b.AgentBirth as JudeAgentBirth,b.AgentID as JudeAgentID,b.AgentAddress as JudeAgentAddress,c.OpenGovNumber as UrgeOGN,c.UrgeTypeID,d.OpenGovNumber,d.BigUnitBossName,d.SubUnitSecBossName,d.SendNumber,d.SendDate,d.Agent,d.AgentBirthDate,d.AgentID,d.AgentAddress,d.MakeSureDate,d.LimitDate,d.AttatchJude,d.AttatchUrge,d.AttatchFortune,d.AttatchGround,d.AttatchRegister,d.AttatchFileList,d.AttatchTable,d.ATTATPOSTAGE,d.SafeToExit,d.SAFEACTION,d.SAFEASSURE,d.SAFEDETAIN,d.SAFESHUTSHOP,e.ReturnResonID,e.ArrivedDate,f.ArrivedDate UrgeArrivedDate from PasserBase a,PasserJude b,PasserUrge c,PasserSend d,(select * from PasserSendArrived where ArriveType=0) e,(select PasserSN,ArrivedDate from PasserSendArrived where ArriveType=1)f where a.SN="&trim(BillSN(i))&" and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+) and a.SN=e.PasserSN(+) and a.SN=f.PasserSN(+)"
PrintDate=split(gArrDT(date),"-")
set rsfound=conn.execute(strSql)
MakeSureDate="":LimitDate=""
If Not ifnull(rsfound("MakeSureDate")) Then
	MakeSureDate=split(gArrDT(rsfound("MakeSureDate")),"-")
	LimitDate=split(gArrDT(rsfound("LimitDate")),"-")
	
elseIf Not ifnull(rsfound("ArrivedDate")) Then
	If rsfound("ReturnResonID") = "1" Then

		MakeSureDate=split(gArrDT(DateAdd("d",50,rsfound("ArrivedDate"))),"-")
		LimitDate=split(gArrDT(DateAdd("d",50,rsfound("ArrivedDate"))),"-")

	else

		MakeSureDate=split(gArrDT(DateAdd("d",30,rsfound("ArrivedDate"))),"-")
		LimitDate=split(gArrDT(DateAdd("d",30,rsfound("ArrivedDate"))),"-")
	End if 

	If not ifnull(rsfound("UrgeArrivedDate")) Then LimitDate=split(gArrDT(DateAdd("d",15,rsfound("UrgeArrivedDate"))),"-")
else
	MakeSureDate=split(gArrDT(""),"-")
	LimitDate=split(gArrDT(""),"-")
End if

If Not ifnull(rsfound("SendDate")) Then
	SendDate=split(gArrDT(rsfound("SendDate")),"-")
else
	SendDate=split(gArrDT(""),"-")
End if

Sys_Address=rsfound("DriverAddress")
Sys_Zip=trim(rsfound("DriverZip"))

Sys_OpenGovNumber="　　　　　　　"
PasserDetail_SN=""
strSQL = "select sn,OpenGovNumber from PasserSendDetail where BillSN="&trim(BillSN(i))&" and sn=(select max(sn) from PasserSendDetail where BillSN="&trim(BillSN(i))&")"

set rsp=conn.execute(strSQL)
If not rsp.eof Then
	Sys_OpenGovNumber=rsp("OpenGovNumber")
	PasserDetail_SN=rsp("sn")
End if
rsp.close

Sys_CreditorGovNumber=""
strSQL="select OpenGovNumber from PasserCreditor where BillSN="&trim(BillSN(i))&" and PetitionDate is not null order by PetitionDate DESC"

set rsCre=conn.execute(strSQL)
If not rsCre.eof Then
	Sys_CreditorGovNumber=trim(rsCre("OpenGovNumber"))
End If 
rsCre.close

paySum=0
strSQL="select nvl(sum(PayAmount),0) as PaySum from PasserPay where BillSN="&trim(BillSN(i))
set rspay=conn.execute(strSQL)
If not rspay.eof Then paySum=cdbl(rspay("PaySum"))
rspay.close

Sys_SendNumber="　　　　　　　"
strSQL = "select SendNumber from PasserSendDetail where BillSN="&trim(BillSN(i))&" and sn=(select max(sn) from PasserSendDetail where BillSN="&trim(BillSN(i))&")"

set rsp=conn.execute(strSQL)
If not ifnull(rsp("SendNumber")) Then Sys_SendNumber=rsp("SendNumber")
rsp.close

Sys_Address=Sys_Zip&Sys_Address
%>
<table width="90%" height="1%" border="0" cellspacing=0 cellpadding=0>
<tr><td align="right">
		<table width="200" border="1" cellspacing=0 cellpadding=0>
		  <tr>
			<td width="60" class="style1">移送案號</td>
			<td width="134" align="left" class="style1"><%=Sys_SendNumber%></td>
		  </tr>
		</table>
	</td>
</tr>
</table>
		<table width="95%" height="90%" border="1" cellspacing=0 cellpadding=0>
		  <tr>
			<td colspan="5" align="center" class="style2"><%=thenPasserCity&replace(DB_UnitName,trim(thenPasserCity),"")%>行政執行案件移送書<br>
			<table border="0" width="100%" height="100%" cellspacing="0" cellpadding="0">
			<tr>
				<td align="left">
					<span class="style1">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					</span>
				</td>
				<td align="left">
					<span class="style1">
						發文日期：<%=SendDate(0)&"年"&SendDate(1)&"月"&SendDate(2)&"日"%>
					</span>
				</td>
			</tr>
			<tr>
				<td align="left">
					<span class="style1">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					</span>
				</td>
				<td align="left">
					<span class="style1">
						發文字號：<%=BillPageUnit&"交字第"&Sys_OpenGovNumber&"號"%>
					</span>
				</td>
			</tr>
			</table>
			</td>
		  </tr>
		  <tr>
			<td colspan="3" align="center" class="style3" width="55%">義　　　　務　　　　人</td>
			<td colspan="2" align="center" class="style3" width="45%">法定代理人或代表人</td>
		  </tr>
		  <tr>
			<td width="120" class="style3">姓名或名稱</td>
			<td colspan="2" class="style3"><%=rsfound("Driver")%></td>
			<td colspan="2" class="style3"><%
				if trim(rsfound("Agent"))<>"" then
					response.write rsfound("Agent")
				else
					response.write rsfound("JudeAgentName")
				end if%>&nbsp;
			</td>
		  </tr>
		  <tr>
			<td class="style3">出生年月日</td>
			<td colspan="2" class="style3"><%
				if trim(rsfound("DriverBirth"))<>"" then
					DriverBirth=split(gArrDT(rsfound("DriverBirth")),"-")
					response.write DriverBirth(0)&"年"&DriverBirth(1)&"月"&DriverBirth(2)&"日"
				end if%>&nbsp;</td>
			<td colspan="2" class="style3"><%
				if trim(rsfound("AgentBirthDate"))<>"" then
					AgentBirthDate=split(gArrDT(rsfound("AgentBirthDate")),"-")
				else
					AgentBirthDate=split(gArrDT(rsfound("JudeAgentBirth")),"-")
				end if
				if trim(AgentBirthDate(0))<>"" then
					response.write "　"&AgentBirthDate(0)&"年"&AgentBirthDate(1)&"月"&AgentBirthDate(2)&"日"
				end if%>&nbsp;
			</td>
		  </tr>
		  <tr>
			<td class="style3">性　　　　別</td>
			<td colspan="2" class="style3"><%
			if Not rsfound.eof then
				If not ifnull(Trim(rsfound("DriverID"))) Then
					If Mid(Trim(rsfound("DriverID")),2,1)="1" Then
						Response.write "男"
					elseif Mid(Trim(rsfound("DriverID")),2,1)="2" Then
						Response.write "女"
					End if
				End if
			end if%>&nbsp;</td>
			<td colspan="2">&nbsp;</td>
		  </tr>
		  <tr>
			<td class="style3">職　　　　業</td>
			<td colspan="2">&nbsp;</td>
			<td colspan="2">&nbsp;</td>
		  </tr>
		  <tr>
			<td class="style3" nowrap>身分證統一號<br>碼或營利事業<br>統 一 編 號</td>
			<td colspan="2" class="style3"><%=rsfound("DriverID")%></td>
			<td colspan="2" class="style3"><%
				if trim(rsfound("AgentID"))<>"" then
					response.write rsfound("AgentID")
				else
					response.write rsfound("JudeAgentID")
				end if%>&nbsp;</td>
		  </tr>
		  <tr>
			<td class="style3" nowrap>住 居 所 或<br>事 務 所 、<br>營 業 所 地<br>址 及 郵 遞<br>區　　　號</td>
			<td colspan="2" class="style3">住：<%=replace(Sys_Address&"","台","臺")%>&nbsp;<br>
				居：
			</td class="style3">
			<td colspan="2" class="style3">住：<%
					response.write rsfound("JudeAgentAddress")%>&nbsp;<br>
				居：
			</td>
		  </tr>
		  <tr>
			<td class="style3">執行標的物<br>所　在　地</td>
			<td colspan="2" class="style3">如附件財產目錄所載</td>
			<td width="124" class="style3">分&nbsp;&nbsp;&nbsp;&nbsp;署<br>收案日期</td>
			<td width="200" class="style3"><%
				'if trim(rsfound("SendDate"))<>"" then
					'SendDate=split(gArrDT(rsfound("SendDate")),"-")
					'response.write SendDate(0)&"年"&SendDate(1)&"月"&SendDate(2)&"日"
				'end if%>&nbsp;</td>
		  </tr>

		  <tr>
			<td class="style3">義務發生之原因或<br>裁處罰鍰之法令依據</td>
			<td colspan="2" class="style6"><%
				response.write "違反道路交通管理處罰條例<br>第"&left(trim(rsfound("Rule1")),2)&"條"
				response.write Mid(trim(rsfound("Rule1")),3,1)&"項"&Mid(trim(rsfound("Rule1")),4,2)&"款"&Mid(trim(rsfound("Rule1")),6,2)
				if len(trim(rsfound("Rule1")))>7 then response.write "之"&right(trim(rsfound("Rule1")),1)
				Response.Write "規定。"

				If not ifnull(trim(rsfound("Rule2"))) Then
					response.write "<br>違反道路交通管理處罰條例<br>第"&left(trim(rsfound("Rule2")),2)&"條"
					response.write Mid(trim(rsfound("Rule2")),3,1)&"項"&Mid(trim(rsfound("Rule2")),4,2)&"款"&Mid(trim(rsfound("Rule2")),6,2)
					if len(trim(rsfound("Rule2")))>7 then response.write "之"&right(trim(rsfound("Rule2")),1)
					Response.Write "規定。"
				End if 
				response.write "<br>"
				'if trim(rsfound("SendDate"))<>"" then
					IllegalDate=split(gArrDT(rsfound("IllegalDate")),"-")
					response.write "違規日："&IllegalDate(0)&"年"&IllegalDate(1)&"月"&IllegalDate(2)&"日"
					'response.write BillPageUnit&"交字第"&rsfound("JudeOGN")&"號"
				'end if%>&nbsp;</td>
			<td width="124" class="style3">行政處分或<br>裁定確定日</td>
			<td width="200" class="style3">
				■　<%=MakeSureDate(0)%>年<%=MakeSureDate(1)%>月<%=MakeSureDate(2)%>日<br>
				□尚未確定
			</td>
		  </tr>
		  <tr>
			<td class="style3">移送法條</td>
			<td colspan="2" class="style3">
				■依據行政執行法第11條<br>
				■依據道路交通管理處罰條例第65條1項3款</td>
			<td class="style3">繳&nbsp;納&nbsp;期&nbsp;間<br>&nbsp;屆　滿　日</td>
			<td class="style3">　<%=LimitDate(0)%>年<%=LimitDate(1)%>月<%=LimitDate(2)%>日</td>
		  </tr>
		  <tr>
			<td rowspan="2" class="style3">執行必要費<br>用核銷機關<br>(單位)名稱<br>及統一編號</td>
			<td width="100" class="style3">核&nbsp;銷&nbsp;機&nbsp;關<br>(單位)名稱</td>
			<td class="style6">
				<%=thenPasserCity&replace(DB_UnitName,trim(thenPasserCity),"")%>
			</td>
			<td class="style3">徵&nbsp;收&nbsp;期&nbsp;間<br>&nbsp;屆　滿　日</td>
			<td class="style3">　<%
'				if sys_City<>"台東縣" then
'					Response.Write LimitDate(0)&"年"&LimitDate(1)&"月"&LimitDate(2)&"日"
'				end if
			%></td>
		</tr>
		<tr>
			<td width="100" class="style3">統&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;一<br>編&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;號</td>
			<td class="style3">
				<%=thePasserVATnumber%>&nbsp;
			</td>
			<td class="style3">應納金額</td>
			<td class="style3">新臺幣<%=cdbl(rsfound("ForFeit"))-paySum%>元<br>（細目詳如附件）</td>
		</tr>
		<tr>
			<td rowspan="4" class="style3">
				承&nbsp;辦&nbsp;移&nbsp;送&nbsp;業<br>務機關（單位）<br>
				名&nbsp;稱&nbsp;與&nbsp;受&nbsp;款<br>金&nbsp;融&nbsp;機&nbsp;構&nbsp;帳
				<br>戶&nbsp;及&nbsp;帳&nbsp;號
			</td>
			<td class="style3">承&nbsp;辦&nbsp;機&nbsp;關<br>(單位)名稱</td>
			<td class="style6">
				<%=thenPasserCity&replace(DB_UnitName,trim(thenPasserCity),"")%>
			</td>
			<td class="style3" colspan="2" nowrap>
				■執行（債權）憑證再移送<br>
				■執行憑證編號：<%=Sys_CreditorGovNumber%>
			</td>			
		</tr>		
		<tr>
			<td class="style3">立&nbsp;帳&nbsp;金&nbsp;融<br>機&nbsp;構&nbsp;名&nbsp;稱</td>
			<td class="style3">
				<%=thePasserSendBankAccountName%>&nbsp;
			</td>
			<td class="style3">催繳情形</td>
			<td class="style3">
				<%
'				strchk="select count(*) as cnt from PasserUrge where BillSN="&rsfound("BillSN")&" and BillNo='"&rsfound("BillNo")&"'"
'				Jodestr="1"
'				set rschk=conn.execute(strchk)
'				if trim(rschk("cnt"))="0" then Jodestr=Cint(rschk("cnt"))
'				rschk.close
'				if trim(Jodestr)<>"0" then
'					response.write "■"
'				else
					response.write "□"
'				end if
				Response.Write "業經催繳<br>"

'				if trim(Jodestr)="0" then
					response.write "■"
'				else
'					response.write "□"
'				end if
				Response.Write "未經催繳"
				%>
			</td>						
		</tr>
		<tr>
			<td class="style3">受&nbsp;款&nbsp;帳&nbsp;戶</td>
			<td class="style3">
				<%=thePasserSendBankName%>&nbsp;
			</td>
			<td rowspan="2" class="style3">催繳方式</td>
			<td rowspan="2" class="style1">□電話催繳<br>
				□明信片或信函方式催繳<br>
				□其他方式（方式為　）</td>
		</tr>
		<tr>
			<td class="style3">帳&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;號</td>
			<td class="style3">
				<%=thePasserSendBankAccount%>&nbsp;
			</td>
		</tr>
		<tr>
			<td class="style3">附件</td>
			<td colspan="4">
				<table border="0" width="100%">
				  <tr>
					<td width="278" class="style3" nowrap>
						<%
						if trim(rsfound("AttatchTable"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>附表<br>
						<%
						if trim(rsfound("AttatchJude"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>處分文書、裁定書或義務人依法令負<br>　有義務之證明文件及送達證明文件<br>
						<%
						if trim(rsfound("AttatchUrge"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>義務人經限期履行而逾期仍不履行<br>　之證明文件及送達證明文件<br>
					</td>
					<td width="209" class="style3">
						<%
						if trim(rsfound("AttatchFortune"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>義務人之財產目錄<br>
						<%
						if trim(rsfound("AttatchGround"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>土地登記簿謄本<br>
						<%
						if trim(rsfound("AttatchRegister"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>義務人之戶藉資料<br>
						<%
						if trim(rsfound("AttatchFileList"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>保全措施之資料<br>
						<%
						if trim(rsfound("ATTATPOSTAGE"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>執行（債權）憑證<br>
						□
					</td>
				  </tr>
			  </table>
			</td>
		  </tr>
		  <tr>
			<td class="style3">保全措施</td>
			<td colspan="4" class="style3"><%
						if trim(rsfound("SAFETOEXIT"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>已限制出境（日期：&nbsp;&nbsp;年&nbsp;&nbsp;月&nbsp;&nbsp;日）<%
						if trim(rsfound("SAFEACTION"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>已禁止處分<%
						if trim(rsfound("SAFEASSURE"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>已提供擔保<br><%
						if trim(rsfound("SAFEDETAIN"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>已假扣押<%
						if trim(rsfound("SAFESHUTSHOP"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>已勒令停業</td>
		  </tr>
		  <tr>
			<td colspan="5">
				<table border="0" width="100%">
					<tr>
						<td class="style2">　　　此　　　致</td>
					</tr>
					<tr>
						<td class="style2">法務部行政執行署　<%
						If showCreditor Then
							If trim(rsfound("AgentAddress")) <> "" Then

								Response.Write trim(rsfound("AgentAddress"))
							else

								If not ifnull(rsfound("DriverZip")) Then
									strSQL="select Administrative from zip where zipid='"&trim(rsfound("DriverZip"))&"'"
									set rszip=conn.execute(strSQL)
									If not rszip.eof Then
										Response.Write replace(rszip("Administrative"),"分署","")
									end if
									rszip.close
								else
									tmpzip=getzip(rsfound("DriverAddress"))
									If tmpzip<>"null" Then
										strSQL="select Administrative from zip where zipid='"&trim(tmpzip)&"'"

										set rszip=conn.execute(strSQL)
										Response.Write replace(rszip("Administrative"),"分署","")
										rszip.close
									End if
								End if
							
							End if 
							
						else
							Response.Write trim(rsfound("AgentAddress"))
						End if
						%>　分署</td>
					</tr>
					<tr>
						<td colspan="4" class="style2" align="center"><%
							if sys_City<>"台南市" and sys_City<>"彰化縣" and sys_City<>"彰化縣" and sys_City<>"宜蘭縣" then
								'Response.Write thenPasserCity&replace(DB_UnitName,trim(thenPasserCity),"")&"　分局長"&theSubUnitSecBossName&"決行"

								Response.Write "　　　　　　　　　　　　分局長"&theSubUnitSecBossName
							end if%>　
						</td>
					</tr>
				</table>
			</td>
		  </tr>
		</table>


