<%
Sys_Driver=trim(rssum("Driver"))
Sys_DriverID=trim(rssum("DriverID"))
Sys_DriverSex="女"
if trim(rssum("DriverSex"))="1" then Sys_DriverSex="男"
Sys_DriverAddress=trim(rssum("DriverAddress"))
Sys_affair=trim(rssum("affair"))
Sys_FORFEIT1=trim(rssum("FORFEIT1"))
Sys_affair=trim(rssum("affair"))

thenPasserUnit=""
strUInfo="select * from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	for j=1 to len(trim(rsUInfo("UnitName")))
		if j<>1 then thenPasserUnit=thenPasserUnit&"　"
		thenPasserUnit=thenPasserUnit&Mid(trim(rsUInfo("UnitName")),j,1)
	next
	thenBillUnitName=trim(rsUInfo("UnitName"))
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
'Sys_Driver=trim(rsSql("Driver"))
'rsSql.close
%>
<table width="645" height="90%" border="1" cellspacing=0 cellpadding=0>
  <tr valign="middle">
    <th height="72" colspan="6"><div align="center"><span class="style2"><%=thenPasserCity&"　"&thenPasserUnit%>　交　辦　單</span></div></th>
  </tr>
  <tr>
    <td width="105" height="73" nowrap><div align="center" class="style6">受文者</div></td>
    <td colspan="2" width="250"><div align="center" class="style6"><%=thenBillUnitName%></div></td>
    <td nowrap><p align="center" class="style6">交辦日期</p></td>
    <td colspan="2" width="169" colspan="2"><div align="center" class="style6"><%=PrintDate(0)%>.<%=PrintDate(1)%>.<%=PrintDate(2)%></div></td>
  </tr>
  <tr>
    <td height="293"><p align="center" class="style6">交辦事項</p></td>
    <td colspan="5">
		<table border=0 width="100%" height="95%">
			<tr valign="top"><td height="87">
				<span class="style6">一、</span></td>
				<td>
				<span class="style6">
				檢送貴所轄內居民&lt;&lt;被通知人<%=Sys_Driver%>&gt;&gt;違反交通管理處罰條例催繳通知書、送達證書、寄存通知書，請派員送達（催繳通知書交當事人）後將送達證書具證回覆。</span></td></tr>
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
	  </table>
	</td>
  </tr>
    <tr>
    <td width="105" height="76" nowrap><div align="center" class="style6">辦理期限</div></td>
    <td nowrap><div align="center" class="style6">日期辦畢連同原件具報</div></td>
	<td nowrap><div align="center" class="style6">組長</div></td>
    <td><div align="center" class="style6"><%
	if trim(sys_City)="台南縣" and trim(Session("Unit_ID"))="J01" then
		response.write "林惠民"
	else
		response.write theSubUnitSecBossName
	end if
	%></div></td>
    <td nowrap><div align="center" class="style6">承辦人</div></td>
    <td nowrap><div align="center" class="style6"><%=request("Session_JudeName")%></div></td>
  </tr>
  <tr>
    <td height="200"><div align="center" class="style6">呈覆內容</div></td>
    <td colspan="5" valign="bottom"><p class="style6">承辦員警職章：　　　　　　　　　主管職章：</p></td>
  </tr>
</table>

