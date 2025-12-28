<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5" />
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {
	font-size: 10pt;
	font-family: "標楷體";}
.style2 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style3 {
	font-size: 9pt;
	font-family: "標楷體";}
-->
</style>
<title>交寄大宗函件</title>
</head>
<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<%
	Sys_MailMoneyValue=trim(request("MailMoneyValue"))
	thenPasserCity=""
	strUInfo="select * from Apconfigure where ID=40"
	set rsUInfo=conn.execute(strUInfo)
	if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
	rsUInfo.close
	set rsUInfo=nothing

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close

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

	if trim(request("Sys_Order"))<>"" then
		orderstr=" order by "&request("Sys_Order")
	end if	

	strSQL="select sn from PasserBase where Exists(select 'Y' from "&BasSQL&" where sn=PasserBase.sn)"&orderstr

	set rs=conn.execute(strSQL)
	BillSN=""
	While Not rs.eof
		If Not ifnull(BillSN) Then BillSN=BillSN&","
		BillSN=BillSN&rs("sn")
		rs.movenext
	Wend
	rs.close
	PageCnt=0
	BillSN=Split(BillSN,",")
	
	Sys_Station=Session("Unit_ID")

	If not ifnull(Request("Sys_MemberStation")) Then
		strSQL="select MemberStation from PasserBase where SN="&BillSN(0)
		set rs=conn.execute(strSQL)
		If not rs.eof Then
			Sys_Station=trim(rs("MemberStation"))
		End if 
		rs.close		
	End if 

	strSQL="select UnitID,UnitTypeID,UnitLevelID from UnitInfo where UnitID='"&Sys_Station&"'"
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
	set unit=conn.Execute(strSQL)
	DB_UnitID=trim(unit("UnitID"))
	DB_UnitName=trim(unit("UnitName"))
	DB_UnitTel=trim(unit("Tel"))
	DB_UnitAddress=trim(unit("Address"))
	DB_BankName=trim(unit("BankName"))
	DB_BankAccount=trim(unit("BankAccount"))
	DB_ManageMemberName=trim(unit("ManageMemberName"))
	unit.close

	PageCnt=fix((Ubound(BillSN)+1)/20+0.99)
	For y = 1 to 2
		If y > 1 Then response.write "<div class=""PageNext"">&nbsp;</div>"

		For i=0 to PageCnt-1
			if i>0 then response.write "<div class=""PageNext"">&nbsp;</div>"%>
			<table width="100%" align="center">
				<tr>
					<td>
						<table width="100%" align="center" cellpadding="3" border="0">
							<tr>
								<td height="45"></td>
							</tr>
							<tr>
								<td width="34%" class="style1">頁&nbsp;&nbsp;次 &nbsp;<%=i+1%> of <%=PageCnt%></td>
								<td rowspan="3" width="39%" align="center">
									<table width="100%">
										<tr>
											<td colspan="3" height="30"><div align="center"><u><span class="style2">中 華 郵 政</span></u></div></td> 
										</tr>
										<tr>
											<td width="37%" rowspan="3" align="right" class="style1">交寄大宗</td>
											<td width="26%" class="style1"><u>限時掛號</u></td>
											<td width="37%" rowspan="3" align="left" class="style1">函件執據</td>
										</tr>
										<tr>
											<td class="style1"><u>掛 &nbsp; &nbsp;號</u></td>
										</tr>
										<tr>
											<td class="style1"><u>快捷郵件</u></td>
										</tr>
									</table>
								</td>
								<td rowspan="3" width="27%" align="right"><img src="../Image/MailPic.JPG" width="100" height="70" /></td>
							</tr>
							<tr>
								<td height="40" valign="top" class="style1">
								中華民國 <%
								response.write year(date)-1911
								%>年 <%
								response.write right("00"&month(date),2)
								%>月 <%
								response.write right("00"&day(date),2)
								%>日
								<br>
								移送監理站日期　年　月　日
								</td>
							</tr>
							<tr>
								<td class="style1">
									寄件人 <%response.write thenPasserCity&replace(DB_UnitName,thenPasserCity,"")%>
								</td>
							</tr>
							<tr>
								<td class="style1">
									寄件人代表 ___________
								</td>
								<td class="style1">
									詳細地址：<u><%=DB_UnitAddress%></u>
								</td>
								<td class="style1">
									電話號碼：<u><%=DB_UnitTel%></u>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<table align="center" width="100%" border="1" cellspacing="0" cellpadding="3">
						<tr>
							<td width="6%" rowspan="2" align="center" class="style1">順序<br>號碼</td>
							<td width="10%" rowspan="2" align="center" class="style1">掛號號碼</td>
							<td colspan="2" align="center" class="style1">收件人</td>
							<td width="5%" rowspan="2" align="center" class="style1">是否<br>回執<br>[V]</td>
							<td width="6%" rowspan="2" align="center" class="style1">郵資</td>
							<td width="9%" rowspan="2" align="center" class="style1">備考</td>
						</tr>
						<tr>
							<td width="23%" align="center" class="style1">姓名</td>
							<td width="41%" align="center" class="style1">送達地名(或地址)</td>
						</tr><%filecnt=0
							For j=1 to 20
								If Ubound(BillSN)>=filecnt Then
									filecnt=i*20+j

									strSQL="select sn,Driver,DriverAddress,BillNo from PasserBase where SN="&BillSN(filecnt-1)				
									set rspbill=conn.execute(strSQL)

									if Not rspbill.eof then
										Sys_ZipID=""

										strSQL="select ZipID from Zip where ZipName like '"&left(trim(rspbill("DriverAddress")),6)&"%'"
										set rszip=conn.execute(strSQL)
										If Not rszip.eof Then Sys_ZipID=trim(rszip("ZipID"))
										rszip.close

										Sys_MailNumber=""

										strSQL="select min(SendMailStation) SendMailStation from PassersEndArrived where ArriveType=2 and PasserSN="&trim(rspbill("sn"))

										set rsSend=conn.execute(strSQL)

										if Not rsSend.eof then
											Sys_MailNumber=rsSend("SendMailStation")

										end If 
										response.write "<tr>"
										response.write "<td align=""left"" class=""style3"">"&filecnt&"</td>"
										response.write "<td align=""left"" class=""style3"">"&Sys_MailNumber&"&nbsp;</td>"
										response.write "<td align=""left"" class=""style3"">"&trim(rspbill("Driver"))&"</td>"
										response.write "<td align=""left"" class=""style3"">"&Sys_ZipID&trim(rspbill("DriverAddress"))&"</td>"
										response.write "<td align=""left"" class=""style3"">&nbsp;&nbsp;</td>"
										response.write "<td align=""left"" class=""style3"">"&Sys_MailMoneyValue&"&nbsp;</td>"
										response.write "<td align=""left"" class=""style3"">"&trim(rspbill("BillNo"))&"</td>"
										response.write "</tr>"
									end if
									rspbill.close
								else
									response.write "<tr>"
									response.write "<td align=""left"" class=""style3"">&nbsp;&nbsp;</td>"
									response.write "<td align=""left"" class=""style3"">&nbsp;&nbsp;</td>"
									response.write "<td align=""left"" class=""style3"">&nbsp;&nbsp;</td>"
									response.write "<td align=""left"" class=""style3"">&nbsp;&nbsp;</td>"
									response.write "<td align=""left"" class=""style3"">&nbsp;&nbsp;</td>"
									response.write "<td align=""left"" class=""style3"">&nbsp;&nbsp;</td>"
									response.write "<td align=""left"" class=""style3"">&nbsp;&nbsp;</td>"
									response.write "</tr>"
								end if
							Next
						response.write "</table>"%>
					</td>
				</tr>
				<tr>
					<td>
						<table align="center" width="100%" border="0">
							<tr>
								<td width="66%" valign="top">
								  <p><span class="style1">(1) 限時掛號、掛號函件與快捷郵件不得同列一單，請將標題塗去其二。<br>
									(2) 函件背面應註明順序號碼，並按號碼次序排齊滿二十件為一組分組交寄。<br>
									(3) 將本埠與外埠函件分別列單交寄。
									<br>
									(4)如有證明郵資、重量必要者，應由寄件人自行在聯單相關欄內分別註明，並結填總郵資，交郵局</span><span class="style1">經辦員逐件核對。<br>
									(5) 日後如須查詢，應於交寄日起六個月內檢同原件封面式樣向原寄局為之，並將本執據送驗。<br>
									(6) 錢鈔或有價證券請利用報值或保價交寄。</span><br>
									
									  </p>
								  </td>
								<td width="34%" class="style1" valign="Top">
									<p>限時掛號<br>
									掛號函件/共 
									<%=(filecnt)%> 
									件照收無誤<br>
									快捷郵件<br>
									<br>
									郵資共計  
									<%
									if Sys_MailMoneyValue<>"" then
										response.write Sys_MailMoneyValue*(filecnt)
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
		<%Next
	next
%>
</body>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
printWindow(true,3.08,3.08,3.08,3.08);
</script>
</html>