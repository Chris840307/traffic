<%
'檢查是否可進入本系統
	strSQLTemp="select SN,IllegalDate,BillNo,Driver," &_
	"BillFillDate,DeallineDate,FORFEIT1,DriverID,Rule1," &_
	"(Select JudeDate from PasserJude where billsn=PasserBase.sn) JUDEDATE," &_
	"(Select UrgeDate from PasserUrge where billsn=PasserBase.sn) URGEDATE," &_
	"(Select OpenGovNumber from PasserUrge where billsn=PasserBase.sn) OpenGovNumber," &_
	"(Select UnitName from UnitInfo where UnitID=PasserBase.BillUnitID) BillUnitName" &_
	" from PasserBase where RecordStateID=0 and Exists(select 'Y' from PasserUrge where BillSN=PasserBase.SN) and Exists(select 'Y' from "&BasSQL&" where SN=PasserBase.SN) order by DriverID"
%>
	<table width="645" border="0" cellspacing=0 cellpadding=0>
		<tr>
			<td align="center" class="style7"><strong>催告催繳裁決清冊</strong></td>
		</tr>
		<tr>
			<td align="left" class="style11">列印日期：<%=gInitDt(date)%></td>
		</tr>
		<tr>
			<td align="left" class="style11">處理時間：<%=request("FromILLEGALDATE")%><%if trim(request("FromILLEGALDATE"))<>"" and trim(request("TOILLEGALDATE"))<>"" then response.write "～"%><%=request("TOILLEGALDATE")%></td>
		</tr>
		<tr>
			<td align="left" class="style11">登入者：<%=Session("Ch_Name")%></td>
		</tr>
	</table>
	<table width="645" border="1" cellspacing=0 cellpadding=0>
		<tr>
			<td class="style11" nowrap>違規人ID</td>
			<td class="style11" nowrap>違規人</td>
			<td class="style11" nowrap>單號</td>
			<td class="style11" nowrap>違規日期</td>
			<td class="style11" nowrap>填單日期</td>
			<td class="style11" nowrap>應到案日</td>
			<td class="style11" nowrap>法條一</td>
			<td class="style11" nowrap>舉發單位</td>
			<td class="style11" nowrap>裁決日期</td>
			<td class="style11" nowrap>催告日期/催告案號</td>
			<td class="style11" nowrap>金額</td>
		</tr><%
			set rsfound=conn.execute(strSQLTemp)
			filecnt=0:sumFile=0:tmpDriverID="":showDriver="":showDriverID="":sumNT=0:cntNt=0:Sys_Payamount=0

'				strSQL="select sum(Payamount) as Sys_Payamount from PasserPay where BillSN="&rsfound("SN")&" and BillNo='"&rsfound("BillNo")&"'"
'				set rspay=conn.execute(strSQL)
'				if not rspay.eof then Sys_Payamount=rspay("Sys_Payamount")
'				rspay.close

				'if Not rsfound.eof then
			tmpDriverID=trim(rsfound("DriverID"))
			showDriverID=trim(rsfound("DriverID"))
			showDriver=trim(rsfound("Driver"))
					'if Not isnull(Sys_Payamount) then sumNT=Cint(Sys_Payamount)
				'end if
				'response.write "<tr><td>"
				'response.write "<table width=""100%"" border=""0"" cellpadding=""4"" cellspacing=""1"">"
				while Not rsfound.eof
					sumFile=sumFile+1
'					Sys_Payamount=0
'					strSQL="select sum(Payamount) as Sys_Payamount from PasserPay where BillSN="&rsfound("SN")&" and BillNo='"&rsfound("BillNo")&"'"
'					set rspay=conn.execute(strSQL)
'					if not rspay.eof then
						tmpDriverID=trim(rsfound("DriverID"))
						Sys_Payamount=trim(rsfound("FORFEIT1"))
'						if Not isnull(rspay("Sys_Payamount")) then Sys_Payamount=rspay("Sys_Payamount")
'					end if
'					rspay.close

					response.write "<tr><td align=""left"" class=""style11"">"
					response.write trim(showDriverID)
					response.write "&nbsp;</td>"
					response.write "<td align=""left"" class=""style11"">"&trim(showDriver)&"&nbsp;</td>"
					response.write "<td align=""left"" class=""style11"">"&trim(rsfound("BillNo"))&"&nbsp;</td>"
					response.write "<td align=""left"" class=""style11"">"&trim(gInitDT(DateValue(rsfound("IllegalDate"))))&"&nbsp;</td>"
					response.write "<td align=""left"" class=""style11"">"&trim(gInitDT(rsfound("BillFillDate")))&"&nbsp;</td>"
					response.write "<td align=""left"" class=""style11"">"&trim(gInitDT(rsfound("DeallineDate")))&"&nbsp;</td>"
					response.write "<td align=""left"" class=""style11"">"&trim(rsfound("Rule1"))&"&nbsp;</td>"
					response.write "<td align=""left"" class=""style11"">"&trim(rsfound("BillUnitName"))&"&nbsp;</td>"
					response.write "<td align=""left"" class=""style11"">"&trim(gInitDT(rsfound("JudeDate")))&"&nbsp;</td>"
					response.write "<td align=""left"" class=""style11"">"&trim(gInitDT(rsfound("UrgeDate")))
					if trim(gInitDT(rsfound("UrgeDate")))<>"" and trim(rsfound("OpenGovNumber"))<>"" then
						response.write "／"
					end if
					response.write trim(rsfound("OpenGovNumber"))&"&nbsp;</td>"
					response.write "<td class=""style11"">"&Sys_Payamount&"&nbsp;</td>"
					response.write "</tr>"

					if Not isnull(Sys_Payamount) then sumNT=sumNT+Cint(Sys_Payamount)
					filecnt=filecnt+1
					rsfound.MoveNext
					showDriver="":showDriverID=""
					if Not rsfound.eof then
						if trim(rsfound("DriverID"))<>trim(tmpDriverID) then
							showDriverID=trim(rsfound("DriverID"))
							showDriver=trim(rsfound("Driver"))
							response.write "<tr><td align=""right"" colspan=""9"" class=""style11"">小計：</td>"
							response.write "<td align=""right"" class=""style11"">"&filecnt&"筆"&"</td>"
							response.write "<td align=""right"" class=""style11"">"&sumNT&"&nbsp;</td></tr>"
							cntNt=cntNt+sumNT
							tmpDriverID=trim(rsfound("DriverID"))
							sumNT=0
							filecnt=0
						end if
					end if
				wend
				response.write "<tr><td align=""right"" colspan=""9"" class=""style11"">小計：</td>"
				response.write "<td align=""right"" class=""style11"">"&filecnt&"筆"&"</td>"
				response.write "<td align=""right"" class=""style11"">"&sumNT&"&nbsp;</td></tr>"
				cntNt=cntNt+sumNT
				response.write "<tr><td align=""right"" colspan=""9"" class=""style11"">共計：</td>"
				response.write "<td align=""right"" class=""style11"">"&sumFile&"筆"&"</td>"
				response.write "<td align=""right"" class=""style11"">"&cntNt&"&nbsp;</td></tr></table>"
				rsfound.close
				set rsfound=nothing
				%>