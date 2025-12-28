<%

strSql="select * from PasserBase where SN="&trim(BillSN(i))
set rs=conn.execute(strSql)
if Not rs.eof then
	Sys_BillTypeID=trim(rs("BillTypeID"))
	Sys_BillNo=trim(rs("BillNo"))
	Sys_DOUBLECHECKSTATUS=trim(rs("DOUBLECHECKSTATUS"))
	Sys_ILLEGALADDRESS=trim(rs("ILLEGALADDRESS"))
	Sys_RuleVer=trim(rs("RuleVer"))
	Sys_Note=trim(rs("Note"))
	sys_Date=split(gArrDT(trim(rs("BillFillDate"))),"-")
	Sys_Driver=trim(rs("Driver"))
	Sys_DriverID=trim(rs("DriverID"))
	Sys_DriverHomeAddress=trim(rs("DriverAddress"))
	Sys_DriverHomeZip=trim(rs("DriverZip"))
	Sys_Rule1=trim(rs("Rule1"))
	Sys_Rule2=trim(rs("Rule2"))
	Sys_Sex=""
	if Not rs.eof then
		If not ifnull(Trim(rs("DriverID"))) Then
			If Mid(Trim(rs("DriverID")),2,1)="1" Then
				Sys_Sex="男"
			elseif Mid(Trim(rs("DriverID")),2,1)="2" Then
				Sys_Sex="女"
			End if
		End if
	end if
	Sys_RecordMemberID=trim(rs("RECORDMEMBERID"))
	Sys_IllegalDate=split(gArrDT(trim(rs("IllegalDate"))),"-")
	Sys_IllegalDate_h=hour(trim(rs("IllegalDate")))
	Sys_IllegalDate_m=minute(trim(rs("IllegalDate")))
	Sys_DealLineDate=split(gArrDT(trim(rs("DealLineDate"))),"-")
	DealLineDate=trim(rs("DealLineDate"))
	Sys_DriverBirth=split(gArrDT(trim(rs("DriverBirth"))),"-")
	Sys_BillFillerMemberID=0
	Sys_Billmem1ID=trim(rs("BILLMEMID1"))
	Sys_STATIONNAME=trim(rs("MemberStation"))
	Sys_billUnitid=trim(rs("BillUnitID"))
end if
rs.close

Sys_UrgeDate="":Sys_OpenGovNumber=""
If not ifnull(request("BillUrge")) Then
	strSQL="select OpenGovNumber,UrgeDate from PasserUrge where BillSN="&trim(BillSN(i))
	set rsjude=conn.execute(strSQL)
	If not rsjude.eof Then
		Sys_OpenGovNumber=trim(rsjude("OpenGovNumber"))
		Sys_UrgeDate=split(gArrDT(trim(rsjude("UrgeDate"))),"-")
	End if
	rsjude.close

elseif trim(UrgeNo)<>"" then

	strSQL="select OpenGovNumber,JudeDate from PasserJude where BillSN="&trim(BillSN(i))
	set rsjude=conn.execute(strSQL)
	If not rsjude.eof Then
		Sys_OpenGovNumber=trim(rsjude("OpenGovNumber"))
		Sys_UrgeDate=split(gArrDT(trim(rsjude("JudeDate"))),"-")
	End if
	rsjude.close
End if

If ifnull(Sys_OpenGovNumber) Then
	Sys_OpenGovNumber=trim(Sys_BillNo)
	Sys_UrgeDate=split(gArrDT(date),"-")
End If 

Sys_UnitLevelID="":Sys_UnitName=""
strSQL="select UnitLevelID,UnitName from unitinfo where UnitID='"&trim(Session("Unit_ID"))&"'"
set rs=conn.execute(strSql)
If not rs.eof Then Sys_UnitLevelID=trim(rs("UnitLevelID"))
If not rs.eof Then Sys_UnitName=trim(rs("UnitName"))
rs.close

If instr(Sys_UnitName,"分局") >0 Then
	strUnit="select * from UnitInfo where UnitID='"&Sys_STATIONNAME&"'"
else
	strUnit="select * from UnitInfo where UnitID='"&Sys_billUnitid&"'"
end if

set rsUnit=conn.execute(strUnit)
If Not rsUnit.eof Then
	If Not rsUnit.eof Then sysunit=replace(rsUnit("UnitName"),"交通組","")
	if Not rsUnit.eof then Sys_UnitAddress=trim(rsUnit("Address"))
	if Not rsUnit.eof then Sys_UnitTel=trim(rsUnit("Tel"))
End if
rsUnit.close

strUnit="select * from UnitInfo where UnitID='"&Sys_STATIONNAME&"'"

set rsUnit=conn.execute(strUnit)
If Not rsUnit.eof Then
	Sys_STATIONNAME=trim(rsUnit("UnitName"))
End if
rsUnit.close

Sys_Level1=0:Sys_Level2=0
strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VERSION=(select value from apconfigure where ID=3)"
set rsRule1=conn.execute(strRule1)
if not rsRule1.eof then
	If DateDiff("d",CDate(date),trim(DealLineDate))>-1 Then 
	  Sys_Level1=trim(rsRule1("Level1"))
	Else
	  Sys_Level1=trim(rsRule1("Level2"))
	End if
end if
rsRule1.close
set rsRule1=nothing

If Not ifnull(Sys_Rule2) Then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VERSION=(select value from apconfigure where ID=3)"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		If DateDiff("d",CDate(date),trim(DealLineDate))>-1 Then 
		  Sys_Level2=trim(rsRule1("Level1"))
		Else
		  Sys_Level2=trim(rsRule1("Level2"))
		End if
	end if
	rsRule1.close
	set rsRule1=nothing
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_DriverHomeZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close

Sum_Level=cdbl(Sys_Level1)+cdbl(Sys_Level2)

strSql="select a.LoginID,a.ChName,a.ImageFilename as MemberFileName,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&trim(Session("User_ID"))
set mem=conn.execute(strsql)
if Not mem.eof then Sys_BillFillerMemberID=trim(mem("LoginID"))
if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
if Not mem.eof then Sys_ChName=trim(mem("ChName"))
if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
mem.close


'If not ifnull(Sys_Billmem1ID) Then
	'strSql="select a.LoginID,a.ChName,b.UnitName,a.ImageFilename as MemberFileName,b.UnitName,b.UnitID,b.UnitLevelID,b.UnitTypeID,b.ImageFilename,b.Tel,b.Address from MemberData a,UnitInfo b where a.UnitID=b.UnitID and MemberID="&Sys_Billmem1ID
'	set mem=conn.execute(strsql)
'	if Not mem.eof then Sys_BillUnitName=trim(mem("UnitName"))
'	if Not mem.eof then Sys_BillUnitTel=trim(mem("Tel"))
'	if Not mem.eof then Sys_BillUnitAddress=trim(mem("Address"))
'	if Not mem.eof then Sys_UnitTypeID=trim(mem("UnitTypeID"))
'	if Not mem.eof then Sys_UnitLevelID=trim(mem("UnitLevelID"))
'	if Not mem.eof then Sys_UnitID=trim(mem("UnitID"))
	'if Not mem.eof then Sys_UnitFilename=trim(mem("ImageFilename"))
	'if Not mem.eof then Sys_ChName=trim(mem("ChName"))
	'if Not mem.eof then Sys_MemberFileName=trim(mem("MemberFileName"))
'	mem.close
'End if

'If Sys_UnitLevelID=1 Then
'	strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
'else
	'strSQL="select * from UnitInfo where UnitID='"&Sys_UnitTypeID&"'"
'end if
'set unit=conn.Execute(strSQL)
'If Not unit.eof Then sysunit=replace(unit("UnitName"),"交通組","")
'if Not unit.eof then Sys_UnitAddress=trim(unit("Address"))
'if Not unit.eof then Sys_UnitTel=trim(unit("Tel"))
'unit.close

if trim(Sys_Rule1)<>"0" and not isnull(Sys_Rule1) then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule1)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_IllegalRule1=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end if

Sys_IllegalRule2=""
if trim(Sys_Rule2)<>"0" and not isnull(Sys_Rule2) then
	strRule1="select * from Law where ItemID='"&trim(Sys_Rule2)&"' and VerSion='"&Sys_RuleVer&"'"
	set rsRule1=conn.execute(strRule1)
	if not rsRule1.eof then
		Sys_IllegalRule2=trim(rsRule1("IllegalRule"))
	end if
	rsRule1.close
	set rsRule1=nothing
end if

strSql="select DCISTATIONID,DCISTATIONNAME,StationTel,StationID from Station where DCIStationID='"&Sys_DCIReturnStation&"'"
set rs=conn.execute(strSql)
if Not rs.eof then Sys_DCISTATIONID=trim(rs("DCISTATIONID"))
if Not rs.eof then Sys_StationTel=trim(rs("StationTel"))
if Not rs.eof then Sys_StationID=trim(rs("StationID"))
rs.close

strSql="select MailNumber,StoreAndSendMailNumber,MailTypeID,MailDate from BillMailHistory where BILLSN="&trim(BillSN(i))

set rs=conn.execute(strSql)
if Not rs.eof then Sys_MailNumber=trim(rs("MailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))

rs.close
if isnull(Sys_DriverHomeZip) or trim(Sys_DriverHomeZip)="" then Sys_DriverHomeZip="001"
if isnull(Sys_OwnerZip) or trim(Sys_OwnerZip)="" then Sys_OwnerZip="001"

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
Sys_MailNumber=0

Sys_BillNo_BarCode=Sys_OpenGovNumber

If sys_City="高雄市" or sys_City="苗栗縣" or sys_City="台中縣" Then

	Sys_BillNo_BarCode=Sys_BillNo

elseIf sys_City="彰化縣" Then

	Sys_BillNo_BarCode="D0"&Sys_BillNo

else

	Sys_BillNo_BarCode=Sys_OpenGovNumber
end if

If sys_City="金門縣" or sys_City="台南市" or sys_City="屏東縣" or sys_City="嘉義市" or sys_City="彰化縣" or sys_City="保二總隊第二大隊第一中隊" or sys_City="保二總隊三大隊二中隊" or sys_City="保二總隊四大隊二中隊" Then

	DelphiASPObj.GenSendStoreBillno Sys_BillNo_BarCode,0,57,160,1

else

	DelphiASPObj.GenSendStoreBillno Sys_BillNo_BarCode,0,57,160
end if

strSql="select MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&trim(BillSN(i))

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close
%>
<br><br>
<table width="645" height="40" border="0" cellspacing=0 cellpadding=0>
	<tr><th rowspan=2 valign="bottom" align="center" width="80%">
			<strong><span class="style25">　　<%=thenPasserCity&replace(sysunit,trim(thenPasserCity),"")%>送達證書</span></strong>
		</th>
		<th valign="bottom" align="left" class="style22" nowrap>
			<%If not ifnull(UrgeDate) Then			
				If sys_City<>"高雄市" Then
					Response.Write UrgeDate&"："

					If sys_City<>"嘉義縣" Then
						response.Write right("00"&Sys_UrgeDate(0),3)&Sys_UrgeDate(1)&Sys_UrgeDate(2)
					end if
				end if
			End if%>
		</th>
	</tr>
	<tr><th valign="bottom" align="left" class="style22" nowrap>
			<%If sys_City<>"高雄市" Then
				Response.Write "序　　號："

				If sys_City="澎湖縣" Then
					response.Write Sys_BillNo
				elseif sys_City<>"嘉義縣" then
					response.Write Sys_OpenGovNumber
				end if
			end if%>
		</th>
	</tr>
</table>
<table width="645" border="1" height="85%" cellspacing=0 cellpadding=0>
  <tr>
    <td colspan="2" align="center"><span class="style22">受送達人名稱姓名地址</span></td>
    <td colspan="3" Valign="top"><span class="style22"><%
		Sys_DriverHomeZip=replace(Sys_DriverHomeZip,"001","")
		response.write funcCheckFont(trim(Sys_Driver),20,1)&"　"
		response.write "<br>"&Sys_DriverHomeZip&Sys_DriverHomeAddress
	%>&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><span class="style23">文　　　　　　　　　號</span></td>
    <td colspan="3" nowrap><span class="style23"><%=BillPageUnit%><%=UrgeNo%><img src=<%="""../BarCodeImage/"&Sys_BillNo_BarCode&".jpg"""%>>號</span></td>
  </tr>
  <tr>
    <td colspan="2" align="center" nowrap><span class="style23">送　達　文　書</span></td>
    <td colspan="3"><span class="style24">違反道路交通管理事件<%=Papertype%><br>
	<%
		If sys_City<>"高雄市" Then
			response.write "道路交通管理處罰條例第"&left(trim(Sys_Rule1),2)&"條"
			if len(trim(Sys_Rule1))>7 then response.write "之"&right(trim(Sys_Rule1),1)
				response.write "第"&Mid(trim(Sys_Rule1),3,1)&"項第"&Mid(trim(Sys_Rule1),4,2)&"款"&Mid(trim(Sys_Rule1),6,2)&"規定"
				'response.write "(期限內自動繳納處新台幣"&Sys_Level1&"元)"
			if trim(Sys_Rule2)<>"" then
				response.write "<br>第"&left(trim(Sys_Rule2),2)&"條"
				if len(trim(Sys_Rule2))>7 then response.write "之"&right(trim(Sys_Rule2),1)
				response.write "第"&Mid(trim(Sys_Rule2),3,1)&"項第"&Mid(trim(Sys_Rule2),4,2)&"款"&Mid(trim(Sys_Rule2),6,2)&"規定"
				'response.write "(期限內自動繳納處新台幣"&Sys_Level2&"元)"
			end if
		end if
			%></span></td>
  </tr>
  <tr>
    <td rowspan="2" align="center"><span class="style23">原寄郵局日戳</span></td>
    <td rowspan="2" align="center"><span class="style23">送達郵局日戳</span></td>
    <td colspan="2" align="center"><span class="style24">送達處所（由送達人填記）</span></td>
    <td rowspan="2" width="20%" align="center"><span class="style23">送達人簽章</span></td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0>
			<tr><td>□</td><td class="style23"><span class="style23">同上記載地址</span></td>
			</tr>
			<tr><td>□</td><td class="style23"><span class="style23">改送：</span></td>
			</tr>
		</table>
	</td>
  </tr>
  <tr>
    <td rowspan="2">&nbsp;</td>
    <td rowspan="2">&nbsp;</td>
    <td colspan="2" align="center"><span class="style24">送達時間（由送達人填記）</span></td>
    <td rowspan="2"><span class="style23">&nbsp;</span></td>
  </tr>
  <tr>
    <td colspan="2">
		<span class="style232">
			<table width="100%" border=0>
				<tr><td><span class="style24">中華民國</span></td>
				<td><span class="style24">　　　　年　　　　月　　　　日</span></td>
				</tr>
				<tr><td class="style24">&nbsp;</td>
				<td><span class="style24">　　　　午　　　　時　　　　分</span></td>
				</tr>
			</table>
		</span>
	</td>
  </tr>
  <tr>
    <td colspan="5" align="center">
		<span class="style22">送　　　　　　　　達　　　　　　　　方　　　　　　　　式</span>
	</td>
  </tr>
  <tr>
    <td colspan="5" align="center">
		<span class="style22">由　　送　　達　　人　　在　　□　　上　　劃　　v　　選　　記</span>
	</td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0><tr><td>□</td><td><span class="style23">已將文書交與應受送達人</span></td></tr></table>		
	</td>
    <td colspan="3">
		<table border=0>
			<tr><td>□</td><td><span class="style23">本人　　　　　　　　　　　　　　　（簽名或蓋章）</span></td></tr>
		</table>
	</td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0>
			<tr><td valign="top">□</td>
				<td>
					<span class="style23">未獲會晤本人，已將文書交與有辨別事理能力之同居人、
					受雇人或應送達處所之接收郵件人員</span>
				</td>
			</tr>
		</table>
	</td>
    <td colspan="3">
		<table border=0>
			<tr><td valign="top">□</td><td class="style24">同居人</span></td>
			</tr>
			<tr><td valign="top">□</td><td class="style24">受雇人　　　　　　　　　　　　　　　　　（簽名或蓋章）</span></td>
			</tr>
			<tr><td valign="top">□</td><td class="style24">應送達處所接收郵件人員</span></td>
			</tr>
		</table>
	</td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0>
			<tr><td valign="top">□</td>
			<td>
				<span class="style24">應受送達之本人、同居人或受雇人收領後，拒絕或不能簽名或蓋章者，
				由送達人記明其事由</span>
			</td>
			</tr>
		</table>
	</td>
    <td colspan="3"><span class="style23">送達人填記：</span></td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0>
			<tr><td valign="top">□</td>
			<td><span class="style24">應受送達之本人、同居人、受雇人或應受送達處所接收郵件人員無正當理由
				拒絕領經送達人將文書留置於送達處所，以為送達</span></td>
			</tr>
		</table>
	</td>
    <td colspan="3">
		<table border=0>
			<tr><td valign="top">□</td><td><span class="style24">本人</span></td>
			</tr>
			<tr><td valign="top">□</td><td><span class="style24">同居人　　　　　　　　　拒絕收領</span></td>
			</tr>
			<tr><td valign="top">□</td><td><span class="style24">受雇人</span></td>
			</tr>
			<tr><td valign="top">□</td><td><span class="style24">應受送達處所接收郵件人員</span></td>
			</tr>
		</table>
	</td>
  </tr>
  <tr>
    <td colspan="2">
		<table border=0>
			<tr><td valign="top">□</td><td><span class="style24">未獲會晤本人亦無受領文書之同居人、受雇人或應受送達處所接收郵件人員，
				已將該送達文書：</span></td>
			</tr>
			<tr><td valign="top">□</td><td><span class="style24">應受送達之本人、同居人、受雇人或應受送達處所接收郵件人員無正當理由
				拒絕收領，並有難達留置情事，已將該送達文書：</span></td>
			</tr>
		</table>
	</td>
    <td colspan="2">
		<table border=0>
			<tr><td valign="top">□</td><td nowrap><span class="style24">寄存於　　　　　　　　　派出所</span></td>
			</tr>
			<tr><td valign="top">□</td><td nowrap><span class="style24">寄存於　　　　　　　　　鄉（鎮、市、區）
			<br>　　　　　　　　　　　　公所</span></td>
			</tr>
			<tr><td valign="top">□</td><td nowrap><span class="style24">寄存於　　　　　　　　　鄉（鎮、市、區）
			<br>　　　　　　　　　　　　公所
			<br>　　　　　　　　　　　　村（里）辦公處</span></td>
			</tr>
			<tr><td valign="top">□</td><td nowrap><span class="style24">寄存於　　　　　　　　　郵局</span></td>
			</tr>
		</table>
	</td>
    <td><span class="style24">並作送達通知書二份，一份黏貼於應受送達人住居所、事務所、營業所或其就業處所門首，一份□交由鄰居轉交或□置於該受送達處所信箱或其他適當位置，以為送達。</span></td>
  </tr>
  <tr>
    <td colspan="2" nowrap><span class="style23">送　達　人　注　意　事　項</span></td>
    <td colspan="3">
		<table border=0>
			<tr><td valign="top"><span class="style24">一、</span></td>
			<td>
			<span class="style24">依上述送達方法送達者，送達人應即將本送達證書，
			提出於交送達之行政機關附卷。</span></td>
			</tr>
			<tr><td valign="top"><span class="style24">二、</span></td>
			<td>
			<span class="style24">不能依上述送達方法送達者，送達人應製作記載該事由之報告書，
			提出於交送達之行政機關附卷，並繳回應送達之文書。</span></td>
			</tr>
		</table>
	</td>
  </tr>
</table>
<span class="style32"><strong><%

If sys_City = "苗栗縣" Then
	Response.Write "※請繳回苗栗中苗郵局第260號信箱　"
	If sys_City<>"高雄市" Then
		Response.Write "應到案處所："&Sys_STATIONNAME&"　操作人員："&Sys_BillFillerMemberID
	End if
	Response.Write "<br>本證書送回地址：苗栗中苗郵局第260號信箱"

else
	Response.Write "※請繳回"&thenPasserCity&replace(sysunit,trim(thenPasserCity),"")&"　"
	If sys_City<>"高雄市" Then
		Response.Write "應到案處所："&Sys_STATIONNAME&"　操作人員："&Sys_BillFillerMemberID
	End if
	Response.Write "<br>本證書送回地址："&Sys_UnitAddress

end if
	%></strong>
<%
if sys_City<>"台南縣" then
	response.write "<br>寄存送達之文書，應保存3個月，如未經領取，請退還交送達機關。"
end if
%>
</span>