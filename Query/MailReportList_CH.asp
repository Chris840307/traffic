<%@LANGUAGE="VBSCRIPT" CODEPAGE="950"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://localhost/traffic/smsx.cab#Version=6,1,432,1">
</object>
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>交寄大宗函件</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
Server.ScriptTimeout = 800
Response.flush
'權限
'AuthorityCheck(234)
%>
<style type="text/css">
<!--
.style33 {
	font-size: 9pt;
	font-family: "標楷體";
}
.style5 {
	font-size: 10pt;
	font-family: "標楷體";}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
-->
</style>
</head>

<body>

<%
'郵資
theMailMoney=trim(request("MailMoneyValue"))
'使用者單位資料
UnitName=""
UnitAddress=""
UnitTel=""
strUnitName="select Value from ApConfigure where ID=40"
set rsUnitName=conn.execute(strUnitName)
if not rsUnitName.eof then
	TitleUnitName=trim(rsUnitName("value"))
end if
rsUnitName.close
set rsUnitName=nothing

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

strSendMailUnit="select UnitName,Address,Tel from MemberData a,UnitInfo b " &_
		" where a.MemberID="&trim(Session("User_ID"))&" and a.UnitID=b.UnitID"
set rsSendMailUnit=conn.execute(strSendMailUnit)
if not rsSendMailUnit.eof then
	
	UnitName=TitleUnitName&trim(rsSendMailUnit("UnitName"))
	UnitAddress=trim(rsSendMailUnit("Address"))
	UnitTel=trim(rsSendMailUnit("Tel"))
end if
rsSendMailUnit.close
set rsSendMailUnit=nothing

strwhere=request("SQLstr")
	ExchangeTypeFlag="W"
	strExchangeType="select a.ExchangeTypeID from DciLog a,BillBase f where a.BillSN=f.SN "&_
		" and f.RecordStateID=0 "&strwhere
	set rsEType=conn.execute(strExchangeType)
	if not rsEType.eof then
		if trim(rsEType("ExchangeTypeID"))="N" then
			ExchangeTypeFlag="N"
		else
			ExchangeTypeFlag="W"
		end if
	else
		ExchangeTypeFlag="W"
	end if
	rsEType.close
	set rsEType=nothing

	if ExchangeTypeFlag="N" then
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.RecordDate" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillSn=g.BillSn" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and e.ExchangeTypeID='N' and e.Status in ('S','N')" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by g.UserMarkDate"
	else
		strSQL="select a.BillSN,a.BillNo,a.BillTypeID,a.CarNo,a.ExchangeTypeID,f.RecordDate" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','5','9','a','j','A','H','K','L','T','V')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N' and e.Status='S'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by f.RecordDate"
	end if

set rs1=conn.execute(strSQL)
	if ExchangeTypeFlag="N" then
		strCnt="select count(*) as cnt" &_
		" from DCILog a" &_
		",MemberData b,DCIReturnStatus d,BillBaseDciReturn e,BillBase f,BillMailHistory g where a.BillSN=f.SN" &_
		" and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillSn=g.BillSn" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and e.ExchangeTypeID='N' and e.Status in ('S','N')" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	else
		strCnt="select count(*) as cnt from DCILog a,MemberData b,DCIReturnStatus d" &_
		",BillBaseDciReturn e,BillBase f where a.BillSN=f.SN and f.RecordStateID=0" &_
		" and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+)" &_
		" and a.BillNO=e.BillNo and a.CarNo=e.CarNO and a.ExchangeTypeID=e.ExchangeTypeID" &_
		" and a.DciReturnStatusID=e.Status and (((((d.DCIreturnStatus=1 and (e.DciErrorCarData not in ('1','3','5','9','a','j','A','H','K','L','T','V')) and f.UseTool<>8) or (d.DCIreturnStatus=1 and f.UseTool=8 and (f.EquipmentID<>'-1' or f.EquipmentID is null))) and f.BillTypeID='2')" &_
		" or (f.BillTypeID='1' and a.DCIReturnStatusID in ('Y','S','n','L') and (f.EquipmentID<>'-1' or f.EquipmentID is null)) and a.ExchangeTypeID='W') or (e.ExchangeTypeID='N' and e.Status='S'))" &_
		" and a.RecordMemberID=b.MemberID(+) "&strwhere
	end if
set rsCnt=conn.execute(strCnt)
if not rsCnt.eof then
	if trim(rsCnt("cnt"))="0" then
		pagecnt=1
	else
		pagecnt=fix(Cint(rsCnt("cnt"))/20+0.9999999)
	end if
end if
rsCnt.close
set rsCnt=nothing

MDate=""
if ExchangeTypeFlag="N" then
	strMailDate="select g.StoreAndSendMailDate as MDate from DciLog a,BillBase f,BillMailHistory g " &_
		" where f.Sn=g.BillSn and f.Sn=a.BillSn "&strwhere
else
	strMailDate="select g.MailDate as MDate from DciLog a,BillBase f,BillMailHistory g " &_
		" where f.Sn=g.BillSn and f.Sn=a.BillSn "&strwhere
end if
	'response.write strMailDate
	set rsMailDate=conn.execute(strMailDate)
	if not rsMailDate.eof then
		MDate=trim(rsMailDate("MDate"))
	end if
	rsMailDate.close
	set rsMailDate=nothing
	if MDate="" or isnull(MDate) then
		MDate=now
	end if

CaseSN=0
If Not rs1.Bof Then
	rs1.MoveFirst 
else
	response.write "查無可列印大宗函件之舉發單！！"
end if
While Not rs1.Eof
if mailSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"

	strList=""
	mailSN=0
	pageNum=fix(CaseSN/20)+1
	for i=1 to 20
		if rs1.eof then exit for
		mailSN=mailSN+1
		CaseSN=CaseSN+1
		strList=strList&"<tr>"
		'順序號碼
		strList=strList&"<td align=""center"">"&mailSN&"</td>"
		'掛號號碼
		theMailNumber=""
		'移送監理站日期
		theSendDocDate=""
		strSqlH="select MailNumber,StoreAndSendMailNumber,SendOpenGovDocToStationDate from BillMailHistory where BillSN="&trim(rs1("BillSN"))
		set rsH=conn.execute(strSqlH)
		if not rsH.eof then
			if trim(rsH("SendOpenGovDocToStationDate"))<>"" and not isnull(rsH("SendOpenGovDocToStationDate")) then
				theSendDocDate=trim(rsH("SendOpenGovDocToStationDate"))
			end if
			if trim(rs1("ExchangeTypeID"))="W" then
				if trim(rsH("MailNumber"))<>"" and not isnull(rsH("MailNumber")) then
					theMailNumber=trim(rsH("MailNumber"))&"&nbsp;"
				else
					theMailNumber="&nbsp;"
				end if
			elseif trim(rs1("ExchangeTypeID"))="N" then
				if trim(rsH("StoreAndSendMailNumber"))<>"" and not isnull(rsH("StoreAndSendMailNumber")) then
					theMailNumber=trim(rsH("StoreAndSendMailNumber"))&"&nbsp;"
				else
					theMailNumber="&nbsp;"
				end if
			else
				theMailNumber="&nbsp;"
			end if
		else
			theMailNumber="&nbsp;"
		end if
		rsH.close
		set rsH=nothing
		strList=strList&"<td align=""center"">"&theMailNumber&"</td>"

		if trim(rs1("BillTypeID"))="2" then	'逕舉要抓Owner
			if ExchangeTypeFlag="N" then
				strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','n','S','L')"
				set rsD2=conn.execute(strSqlD2)
				if not rsD2.eof then
					if ExchangeTypeFlag="N" then	'單退先抓W看有沒有做戶籍補正，沒有的話再抓A,再沒有就抓owner
						if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress"))  then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
							if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
								GetMailMem="&nbsp;"
							else
								GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
						else
							strSqlD="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='A' and Status='S'"
							set rsD=conn.execute(strSqlD)
							if not rsD.eof then
								if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
									if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
										GetMailMem="&nbsp;"
									else
										GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
									end if
									GetMailAddress=trim(rsD("DriverHomeZip"))&trim(rsD("DriverHomeAddress"))
								else
									if isnull(rsD("Owner")) or trim(rsD("Owner"))="" then
										GetMailMem="&nbsp;"
									else
										GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
									end if
									GetMailAddress="(車)"&trim(rsD("OwnerZip"))&trim(rsD("OwnerAddress"))
								end if
							else
								strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
								set rsZip=conn.execute(strZip)
								if not rsZip.eof then
									ZipName=trim(rsZip("ZipName"))
								end if
								rsZip.close
								set rsZip=nothing
				
								if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
									GetMailMem="&nbsp;"
								else
									GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
								end if
								GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
							end if
							rsD.close
							set rsD=nothing
						end if
					else
						'入案直接抓owner
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing
							if isnull(rsD2("Owner")) or trim(rsD2("Owner"))="" then
								GetMailMem="&nbsp;"
							else
								GetMailMem=trim(replace(rsD2("Owner")," "," &nbsp;"))
							end if
							GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
					end if
				end if
				rsD2.close
				set rsD2=nothing
			else
					strSqlD2="select DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','S','n','L')"
					set rsD2=conn.execute(strSqlD2)
					if not rsD2.eof then
						if trim(rsD2("DriverHomeAddress"))<>"" and not isnull(rsD2("DriverHomeAddress")) and ExchangeTypeFlag="N" then
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							GetMailAddress=trim(rsD2("DriverHomeZip"))&ZipName&trim(rsD2("DriverHomeAddress"))
						else
							strZip="select ZipName from Zip where ZipID='"&trim(rsD2("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing

							GetMailMem=trim(rsD2("Owner"))
							if ExchangeTypeFlag="N" then
								GetMailAddress="(車)"&trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
							else
								GetMailAddress=trim(rsD2("OwnerZip"))&ZipName&trim(rsD2("OwnerAddress"))
							end if
						end if
					end if
					rsD2.close
					set rsD2=nothing
			end if
		else	'攔停抓Driver
			strSqlD="select Driver,DriverHomeZip,DriverHomeAddress,Owner,OwnerZip,OwnerAddress from BIllBaseDCIReturn where ((BillNo='"&trim(rs1("BillNo"))&"' and CarNo='"&trim(rs1("CarNo"))&"') or (BillNo is null and CarNo='"&trim(rs1("CarNo"))&"')) and ExchangeTypeID='W' and Status in ('Y','S','n','L')"
			set rsD=conn.execute(strSqlD)
			if not rsD.eof then
				
				if sys_City="彰化縣" then
					if not isnull(rsD("Driver")) and trim(rsD("Driver"))<>"" then
						GetMailMem=trim(replace(rsD("Driver")," "," &nbsp;"))
					elseif not isnull(rsD("Owner")) and trim(rsD("Owner"))<>"" then
						GetMailMem=trim(replace(rsD("Owner")," "," &nbsp;"))
					else
						GetMailMem="&nbsp;"
					end if
					if trim(rsD("DriverHomeAddress"))<>"" and not isnull(rsD("DriverHomeAddress")) then
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
						GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&trim(rsD("DriverHomeAddress"))
					else
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("OwnerZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing

						GetMailAddress="(車)"&trim(rsD("OwnerZip"))&ZipName&trim(rsD("OwnerAddress"))
					end if
				else
					if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣"  then
						ZipName=""
					else
						strZip="select ZipName from Zip where ZipID='"&trim(rsD("DriverHomeZip"))&"'"
						set rsZip=conn.execute(strZip)
						if not rsZip.eof then
							ZipName=trim(rsZip("ZipName"))
						end if
						rsZip.close
						set rsZip=nothing
					end if
					if isnull(rsD("Driver")) or trim(rsD("Driver"))="" then
						GetMailMem="&nbsp;"
					else
						GetMailMem=trim(replace(rsD("Driver")," "," &nbsp;"))
					end if
					GetMailAddress=trim(rsD("DriverHomeZip"))&ZipName&trim(rsD("DriverHomeAddress"))
				end if

				
			end if
			rsD.close
			set rsD=nothing
		end if
		'收件人姓名
		strList=strList&"<td align=""left"" class=""style33"">"&funcCheckFont(GetMailMem,14,1)&"</td>"
		'收件地址
		strList=strList&"<td align=""left"" class=""style33"">"&funcCheckFont(GetMailAddress,14,1)&"</td>"
		strList=strList&"<td align=""center"">&nbsp;</td>"
		'郵資
		if theMailMoney<>"" then
			theMailMoneyTmp=theMailMoney
		else
			theMailMoneyTmp="&nbsp;"
		end if
		strList=strList&"<td align=""center"">"&theMailMoneyTmp&"</td>"
		'備考=單號
		strList=strList&"<td align=""center"">"&trim(rs1("BillNO"))&"</td>"
		strList=strList&"</tr>"
		rs1.MoveNext
	next
	if mailSN<20 then
		mailSNTmp=mailSN
		for Sp=1 to 20-mailSN
			mailSNTmp=mailSNTmp+1
			strList=strList&"<tr>"
			'順序號碼
			strList=strList&"<td align=""center"">"&mailSNTmp&"</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"<td align=""center"">&nbsp;</td>"
			strList=strList&"</tr>"
		next
	end if


%>
<table width="100%" align="center">
<tr>
<td>
	<table width="100%" align="center" cellpadding="3" border="0">
	<tr>
		<td height="45"></td>
	</tr>
	<tr>
		<td width="34%"><span class="style7">
		頁&nbsp;&nbsp;次 &nbsp;<%=pageNum%> of <%=pagecnt%>
		</span></td>
		<td rowspan="3" width="39%" align="center"><span class="style7">
		<table width="100%">
		<tr>
			<td colspan="3" height="30"><div align="center"><u><span class="style6">中 華 郵 政</span></u></div></td> 
		</tr>
		<tr>
			<td width="37%" rowspan="3" align="right" class="style7">交寄大宗</td>
			<td width="26%" class="style7"><u>限時掛號</u></td>
			<td width="37%" rowspan="3" align="left" class="style7">函件執據</td>
		</tr>
		<tr>
			<td class="style7"><u>掛 &nbsp; &nbsp;號</u></td>
		</tr>
		<tr>
			<td class="style7"><u>快捷郵件</u></td>
		</tr>
		</table>
		
		</span></td>
		<td rowspan="3" width="27%"><div align="right"><img src="../Image/MailPic.JPG" width="100" height="70" /></div></td>
	</tr>
	<tr>
		<td height="40" valign="top"><span class="style7">
		中華民國 <%
		response.write year(MDate)-1911
		%>年 <%
		response.write right("00"&month(MDate),2)
		%>月 <%
		response.write right("00"&day(MDate),2)
		%>日
		<br>
		移送監理站日期 <%
			if theSendDocDate<>"" then
				response.write left(theSendDocDate,3)
			end if
		%>年 <%
			if theSendDocDate<>"" then
				response.write mid(theSendDocDate,4,2)
			end if
		%>月 <%
			if theSendDocDate<>"" then
				response.write mid(theSendDocDate,6,2)
			end if
		%>日
		</span></td>
	</tr>
	<tr>
		<td><span class="style7">
		寄件人 <%
		response.write UnitName
		%>
		</span></td>
	</tr>
	<tr>
		<td><span class="style7">
		寄件人代表 ___________
		</span></td>
		<td><span class="style7">
		詳細地址：<u><%=UnitAddress%></u>
		</span></td>
		<td><span class="style7">
		電話號碼：<u><%=UnitTel%></u>
		</span></td>
	</tr>
	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="1" cellspacing="0" cellpadding="3">
	<tr>
	<td width="6%" rowspan="2"><div align="center"><span class="style5">順序<br>
	  號碼</span></div></td>
	<td width="10%" rowspan="2"><div align="center"><span class="style5">掛號號碼</span></div></td>
	<td colspan="2"><div align="center"><span class="style5">收件人</span></div></td>
	<td width="5%" rowspan="2"><div align="center"><span class="style22">是否<br>
	  回執<br>[V]</span></div></td>
	

	<td width="6%" rowspan="2"><div align="center"><span class="style5">郵資</span></div></td>
	<td width="9%" rowspan="2"><div align="center"><span class="style5">備考</span></div></td>
	</tr>
	<tr>
	<td width="23%" class="style5"><div align="center">姓名</div></td>
	<td width="41%" class="style5"><div align="center">送達地名(或地址)</div></td>
	</tr>
	<%=strList%>
	</table>
</td>
</tr>
<tr>
<td>
	<table align="center" width="100%" border="0">
	<tr>
	<td width="66%" valign="top">
	  <p><span class="style11">(1) 限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br>
	    (2) 函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br>
	    (3) 將本埠與外埠函件分別列單交寄。
	    <br>
	    (4)如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局</span><span class="style11">經辦員逐件核對。<br>
	    (5) 日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br>
	    (6) 錢鈔或有價證券請利用報值或保價交寄。</span><br>
	    
	      </p>
	  </td>
	<td width="34%" class="style5" valign="Top">
	  <p>限時掛號<br>
	    掛號函件/共 
	    <%=mailSN%> 
	    件照收無誤<br>
	    快捷郵件<br>
	    
	    <br>
	    郵資共計  
	    <%
		if theMailMoney<>"" then
			response.write theMailMoney*mailSN
		else
			response.write "&nbsp;"
		end if
		%> 
	    元	  </p>
	  <p align="right">______________<br>經辦員簽署&nbsp; </p>
	  </td>
	</tr>
	</table>
</td>
</tr>
</table>

<%		
	
Wend
rs1.close
set rs1=nothing
%>			
</body>
<script language="javascript">
printWindow(true,3.08,3.08,3.08,3.08);
</script>
</html>
