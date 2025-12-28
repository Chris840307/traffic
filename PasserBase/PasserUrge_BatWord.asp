<%
strSQL="select OpenGovNumber,UrgeDate from PasserUrge where BillSN="&trim(BillSN(i))
set rsjude=conn.execute(strSQL)
If not rsjude.eof Then
	Sys_OpenGovNumber=trim(rsjude("OpenGovNumber"))
	Sys_UrgeDate=split(gArrDT(rsjude("UrgeDate")),"-")
End if
rsjude.close

strSQL="select OpenGovNumber,JudeDate from PasserJude where BillSN="&trim(BillSN(i))
set rsjude=conn.execute(strSQL)
If not rsjude.eof Then
	Sys_JudeGovNumber=trim(rsjude("OpenGovNumber"))
	Sys_JudeDate=split(gArrDT(rsjude("JudeDate")),"-")
End if
rsjude.close


strPay="select nvl(sum(PayAmount),0) as PaySum from PasserPay where BillSN="&trim(BillSN(i))
set rsPay=conn.execute(strPay)
If not rsPay.eof Then

	if trim(rsPay("PaySum"))<>"" then

		Sys_FORFEIT1=Sys_FORFEIT1-cdbl(rsPay("PaySum"))
	end If 
end If 
rsPay.close

PrintDate=split(gArrDT(date),"-")
'strUInfo="select * from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
'set rsUInfo=conn.execute(strUInfo)
'if not rsUInfo.eof then
'	theSubUnitSecBossName=trim(rsUInfo("SecondManagerName"))
'	theBigUnitBossName=trim(rsUInfo("ManageMemberName"))
'	theContactTel=trim(rsUInfo("Tel"))
'	theBankAccount=trim(rsUInfo("BankAccount"))
'	theBankName=trim(rsUInfo("BankName"))
'	theUnitName=trim(rsUInfo("UnitName"))
'end if
'rsUInfo.close%>
<table width="645" height="90%" border="1" cellspacing=0 cellpadding=0>
  <tr>
    <th height="86" colspan="4"><div align="center" class="style20">違反道路交通管理事件催繳通知書</div>
	</th>
  </tr>
  <tr>
    <td align="center" height="60"><div class="style21">事　　由</div></td>
    <td colspan="3"><span class="style21">違反道路交通管理事件處罰案。　<%=BillPageUnit%>催字第<%=left(Sys_OpenGovNumber&"　　　　　　",9)%>號</span></td>
  </tr>
  <tr>
    <td height="57" align="center"><span class="style21">送達文件</span></td>
    <td><span class="style21">催繳通知書</span></td>
	<td height="57" align="center"><span class="style21">發文日期</span></td>
    <td><span class="style21"><%="民國"&Sys_UrgeDate(0)&"年"&Sys_UrgeDate(1)&"月"&Sys_UrgeDate(2)&"日"%></span></td>
  </tr>
  <tr>
    <td height="61" align="center"><span class="style21">受送達人<br>
    姓　　名</span></td>
    <td colspan="3"><span class="style21">被通知人：<%=Sys_Driver%>、性別：<%=Sys_DriverSex%>、身分證統一<%
	If sys_City="台南市" Then
		response.write "編號"
	Else
		response.write "號碼"
	End If 
	%>：<%=Sys_DriverID%></span></td>
  </tr>
  <tr>
    <td height="64" align="center"><span class="style21">送達處所</span></td>
    <td colspan="3"><span class="style21">戶籍地：<%=Sys_DriverZip&Sys_DriverAddress%></span></td>
  </tr>
  <tr valign="top">
    <td height="224" colspan="4">
		<table border=0 width="100%">
			<tr><td height="81" valign="top"><span class="style21">一、</span></td>
	<%If sys_City="嘉義縣" Then%>
			<td valign="top"><span class="style21">臺端<%=Sys_IllegalDate(0)%>年度違反道路交通管理事件<%=Sys_affair%>件，應繳納新臺幣<%=to_Money(Sys_FORFEIT1)%>元正，已逾期未繳〈本分局<%=BillPageUnit%>裁字第<%=Sys_JudeGovNumber%>號裁決書〉。</span></td></tr>
	<%else%>
			<td valign="top"><span class="style21">臺端<%=Sys_IllegalDate(0)%>年度違反道路交通管理事件<%=Sys_affair%>件，應繳納新臺幣&lt;&lt;金額<%=to_Money(Sys_FORFEIT1)%>元正&gt;&gt;，已逾期未繳。<%=BillPageUnit%>裁字第<%=Sys_JudeGovNumber%>號</span></td></tr>
	<%End If %>
			<tr><td height="110" valign="top"><span class="style21">二、</span></td>
	<%If sys_City="台南市" Then%>
			<td valign="top"><span class="style21">請於收受本通知書後，三十日內至本分局繳納或用匯票匯入本分局帳號，上開繳納罰款如未能依時限繳納，已違反行政執行法第四條，金錢給付義務逾期不履行者，將移送行政執行署所屬分署，依法強制執行。</span></td></tr>
	<%else%>
			<td valign="top"><span class="style21">請於收受本通知書後，三十日內至本分局臨櫃繳納或郵政劃撥入本分局帳號，上開繳納罰款如未能依時限繳納，已違反行政執行法第四條，金錢給付義務逾期不履行者，每一違規案件，將移送行政執行署所屬分署，依法強制執行。</span></td></tr>
	<%End If %>
			<tr><td height="81" valign="top"><span class="style21">三、</span></td>
			<td valign="top"><span class="style21">為顧及臺端之權益及本於便民措施，特再通知。（<%=theContactTel%>）</span></td></tr>
			<tr><td height="88" valign="top"><span class="style21">四、</span></td>
			<td valign="top"><span class="style21">劃撥戶名：<%=theBankName%>。<br>劃撥帳號：<%=theBankAccount%>。</span></td></tr>
	  </table>
  </tr>
  <tr>
    <td height="121" align="center"><span class="style21">發　文<br>
    單　位</span></td>
    <td><%=thenPasserCity&"<br>"&theUnitName%></td>
    <td colspan="2" rowspan>
	
	承辦人：<%
	If sys_City<>"澎湖縣" and sys_City<>"高雄縣" and sys_City<>"嘉義市" and sys_City <> "台東縣" Then
		if trim(Sys_MemUnitFileName)<>"" then
			response.write "<img src=""../Member/Picture/"&Sys_MemUnitFileName&""" width=""90"" height=""30"">"
		else
			'南投竹山分局
			'南投竹山分局  從這邊可以設定 操作者與承辦人不同人
			if Session("Unit_ID") = "05FG" then 
				response.write left(request("Session_JudeName")&"　　　　　　　　",5)
			elseIf sys_City<>"台南市" and Session("Unit_ID") <> "F000" Then
				response.write left(request("Session_JudeName")&"　　　　　　　　",5)
			else
				response.write "　　　　　　　　"
			end if
		end if
	else
		response.write "　　　　　　　　"
	end If 
	Response.Write "<br>"

	If trim(session("Sys_UnitChName"))<>"" Then

		if sys_City="台東縣" or sys_City = "高雄縣" or sys_City="嘉義市" then
			response.write "組長："
			response.write session("Sys_UnitChName")&"　　"

			response.Write "<br>"

		elseif sys_City="嘉義縣" then
			response.write "單位主管："
			response.write session("Sys_UnitChName")&"　　"

			response.Write "<br>"

		elseIf sys_City<>"基隆市" and sys_City<>"高雄市" and sys_City<>"台中市" and sys_City<>"屏東縣" Then
			response.write "單位主官："
			response.write session("Sys_UnitChName")&"　　"

			response.Write "<br>"

		end If 
	End if 
		


		
		
'		If sys_City<>"嘉義市" and sys_City<>"澎湖縣" and sys_City<>"台東縣" and sys_City<>"屏東縣" Then
'			if Session("Unit_ID") <> "05FG" and Session("Unit_ID") <> "F000" then
'				If sys_City<>"宜蘭縣"  and sys_City<>"台南市" and sys_City<>"台南縣" and sys_City<>"花蓮縣" then
'					If sys_City<>"嘉義縣" and sys_City<>"高雄縣" and sys_City<>"彰化縣" and sys_City<>"基隆市" and sys_City<>"高雄市" and sys_City<>"台中市" and sys_City<>"台中縣" and sys_City<>"金門縣" then
'						response.write "局長："&left(theBigUnitBossName&"　　　　　　　　",10)
'					elseif sys_City<>"彰化縣" then
'						response.write "分局長："&left(theSubUnitSecBossName&"　　　　　　　　",10)
'					end if
'				end if
'			end if

		if sys_City="台東縣" or sys_City="嘉義市" then

			response.write "分局長："&left("　　　　　　　　",10)

		elseif sys_City="雲林縣" then

			response.write left("　　　　　　　　",10)

		elseIf sys_City="基隆市" or sys_City="高雄市" or sys_City="台中市" or sys_City="屏東縣" Then
			response.write "分局長："&left(theSubUnitSecBossName&"　　　　　　　　",10)
		end If 
	%>&nbsp;</span></td>
  </tr>
</table>
