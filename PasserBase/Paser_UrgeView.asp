<%
Sys_Driver=trim(rssum("Driver"))
Sys_DriverID=trim(rssum("DriverID"))
Sys_DriverSex="女"
if trim(rssum("DriverSex"))="1" then Sys_DriverSex="男"
Sys_DriverAddress=trim(rssum("DriverAddress"))
Sys_affair=trim(rssum("affair"))
Sys_FORFEIT1=trim(rssum("FORFEIT1"))
Sys_affair=trim(rssum("affair"))
Sys_DriverAddress=trim(rssum("DriverAddress"))
thenPasserUnit="":thenBillUnitName=""
'strUInfo="select * from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
strUInfo="select * from UnitInfo where UnitID='"&trim(rssum("billunitid"))&"'"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	for j=1 to len(trim(rsUInfo("UnitName")))
		if j<>1 then thenPasserUnit=thenPasserUnit&"　"
		thenPasserUnit=thenPasserUnit&Mid(replace(rsUInfo("UnitName"),"交通組",""),j,1)
	next
	thenBillUnitName=replace(trim(rsUInfo("UnitName")),"交通組","")
end if
rsUInfo.close
set rsUInfo=nothing

'strUInfo="select * from UnitInfo where UnitID='"&trim(rsSql("BillUnitID"))&"'"
'set rsUInfo=conn.execute(strUInfo)
'if not rsUInfo.eof then
'	theSubUnitSecBossName=trim(rsUInfo("SecondManagerName"))
'	theBigUnitBossName=trim(rsUInfo("ManageMemberName"))
'	theContactTel=trim(rsUInfo("Tel"))
'	theBankAccount=trim(rsUInfo("BankAccount"))
'	thenBillUnitName=trim(rsUInfo("UnitName"))
'end if
'rsUInfo.close
'set rsUInfo=nothing
PrintDate=split(gArrDT(date),"-")
RePrintDate=split(gArrDT(DateAdd("d",3,date)),"-")
'Sys_Driver=trim(rsSql("Driver"))
'rsSql.close
%>
<table width="645" height="90%" border="1" cellspacing=0 cellpadding=0>
  <tr valign="middle">
    <th height="72" colspan="6"><div align="center"><span class="style2"><%=thenPasserCity&"　"&thenPasserUnit%>　交　辦　單</span></div></th>
  </tr>
  <tr>
    <td rowspan="2" width="105" height="73" nowrap><div align="center" class="style6">受文者</div></td>
    <td rowspan="2" colspan="2" width="250"><div align="center" class="style6"><%=thenBillUnitName%></div></td>
    <td><p align="center" class="style6">交辦日期</p></td>
	<td nowrap colspan="2"><p align="center" class="style6">回覆期限</p></td>
  </tr>
  <tr>
    <td width="169"><div align="center" class="style6"><%=PrintDate(0)%>.<%=PrintDate(1)%>.<%=PrintDate(2)%></div></td>
	<td width="169" colspan="2"><div align="center" class="style6"><%=RePrintDate(0)%>.<%=rePrintDate(1)%>.<%=rePrintDate(2)%></div></td>
  </tr>
  <tr>
    <td height="293"><p align="center" class="style6">交辦事項</p></td>
    <td colspan="5">
		<table border=0 width="100%" height="95%">
			<tr valign="top"><td height="87">
				<span class="style6">一、</span></td>
				<td>
				<span class="style6">
				檢送貴所轄內居民&lt;&lt;被通知人<%=Sys_Driver%>&gt;&gt;違反交通管理處罰條例<%=Papertype%>、送達證書、寄存通知書，請派員送達（<%=Papertype%>交當事人）後將送達證書具證回覆。</span></td></tr>
			<tr valign="top"><td height="84">
				<span class="style6">二、</span></td>
				<td><span class="style6">若違規人（同居人或受雇人）無正當理由拒絕受領，請予留置送達及拍照存證並敘明事實 。</span></td>
		  </tr>
			<tr valign="top"><td height="104">
				<span class="style6">三、</div></td>
				<td>
				<span class="style6">
				屢查未遇應受送達人，請採寄存送達方式，並將寄存送達通知書一份黏貼於應受送達人門首，將催繳通知書交由鄰居轉交或置於應受送達處所信箱內以為送達，並請拍照存證。 </span></td>
		  </tr>
		  <tr valign="top"><td height="104">
				<span class="style6">四、</div></td>
				<td>
				<span class="style6">
				送達地址：<%=Sys_DriverAddress%>。 </span></td>
		  </tr>
	  </table>
	</td>
  </tr>
    <tr>
    <td width="105" height="76" nowrap><div align="center" class="style6">辦理期限</div></td>
    <td align="center" class="style11" nowrap>日期辦畢連同原件具報</td>
	<td nowrap><div align="center" class="style6">組長</div></td>
    <td width="150" nowrap><div align="center" class="style6"><%
	if trim(sys_City)="台南縣" and trim(Session("Unit_ID"))="J01" then
		response.write "林惠民"

	elseif trim(sys_City)="台中市" and left(Session("Unit_ID"),3)="045" then
		response.write "劉彥亨"

	elseif trim(sys_City)="嘉義縣" then

		response.write theBigUnitBossName

	else
'		strSQL="select chName from MemberData where UnitID='"&Session("Unit_ID")&"' and JobID=318 and AccountStateID=0 and RecordStateID=0"
'		set rsmen=conn.execute(strSQL)
'		If not rsmen.eof Then
'			Response.Write rsmen("chName")
'		else
'			response.write theSubUnitSecBossName
'		End if
'		rsmen.close

		Response.Write session("Sys_UnitChName")
	end if
	%></div></td>
    <td nowrap><div align="center" class="style6">承辦人</div></td>
    <td width="300"><div align="center" class="style6"><%=request("Session_JudeName")%></div></td>
  </tr>
  <tr>
    <td height="200"><div align="center" class="style6">呈覆內容</div></td>
    <td colspan="5" valign="bottom"><p class="style6">承辦員警職章：　　　　　　　　　主管職章：</p></td>
  </tr>
</table>

