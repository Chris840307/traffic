<%
Server.ScriptTimeout=60000
'檢查是否可進入本系統
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

	strwhere=""
	if request("FromILLEGALDATE")<>"" and request("TOILLEGALDATE")<>""then
		ArgueDate1=gOutDT(request("FromILLEGALDATE"))&" 0:0:0"
		ArgueDate2=gOutDT(request("TOILLEGALDATE"))&" 23:59:59"
		strwhere=" and a.IllegalDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS')"
	end if
	strSQLTemp="select a.driveraddress,a.SN,a.IllegalDate,a.BillNo,a.Driver,a.DriverID," &_
	"a.Rule1,a.FORFEIT1,a.BillFillDate,a.DeallineDate," &_
	"(Select SendDate from PasserSend where billsn=a.sn) SENDDATE," &_
	"(Select UnitName from UnitInfo where UnitID=a.BillUnitID) BillUnitName," &_
	"(Select sum(Payamount) from PasserPay where billsn=a.sn) Payamount," &_
	"(Select JudeDate from PasserJude where billsn=a.sn) JUDEDATE," &_
	"(Select OpenGovNumBer from PasserJude where billsn=a.sn) JudeNo" &_
	" from PasserBase a where Exists(select 'Y' from "&BasSQL&" where sn=a.sn) order by DriverID"

%><table width="645" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td align="center" class="style7"><strong>強制執行移送清冊</strong></td>
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
			<%  'smith 20091007 新增 基隆市清冊的格式			
			    noshowcity="基隆市"  
				if sys_City=noshowcity then %>
				<td class="style27" nowrap>流水號</td>
			<%end if%>
			<td class="style27" nowrap>違規人ID</td>
			<td class="style27" nowrap>違規人</td>
			
			<% if sys_City<>noshowcity then %>
				<td class="style27" nowrap>單號</td>
				<td class="style27" nowrap>違規日</td>
				<td class="style27" nowrap>填單日</td>
				<td class="style27" nowrap>應到案日</td>
				<td class="style27" nowrap>法條一</td>
				<td class="style27" nowrap>舉發單位</td>
				<td class="style27" nowrap>裁決日/裁決案號</td>
			<% else %>
				<td class="style27" nowrap>繳納期間屆滿日</td>
				<td class="style27" nowrap>行政處分或<br>裁定確定日</td>
				<td width="200" class="style27" nowrap>戶籍地址</td>			
				<td class="style27" nowrap>裁決案號</td>
			<%end if%>
			
			<% if sys_City<>noshowcity then %>			
				<td class="style27" nowrap>金額</td>
			<%end if%>				
		</tr><%
			set rsfound=conn.execute(strSQLTemp)
			filecnt=0:sumFile=0:showDriver="":showDriverID="":tmpDriverID="":sumNT=0:cntNt=0:Sys_Payamount=0
			if Not rsfound.eof then
				tmpDriverID=trim(rsfound("DriverID"))
				showDriverID=trim(rsfound("DriverID"))
				showDriver=trim(rsfound("Driver"))
			end if
			if Not isnull(rsfound("PayaMount")) then sumNT=Cint(rsfound("PayaMount"))
				
				while Not rsfound.eof
					sumFile=sumFile+1
					tmpDriverID=trim(rsfound("DriverID"))
					Sys_Payamount=trim(rsfound("FORFEIT1"))
					
					
					if sys_City=noshowcity then 
						response.write "<td align=""left"" class=""style27"" nowrap>"&sumFile&"</td>"
					end if
					response.write "<td align=""left"" class=""style27"" nowrap>"&trim(showDriverID)&"&nbsp;</td>"
					response.write "<td align=""left"" class=""style27"" nowrap>"&trim(showDriver)&"&nbsp;</td>"
					
					if sys_City<>noshowcity then
						response.write "<td align=""left"" class=""style27"" nowrap>"&trim(rsfound("BillNo"))&"&nbsp;</td>"
						response.write "<td align=""left"" class=""style27"">"&trim(gInitDT(DateValue(rsfound("IllegalDate"))))&"&nbsp;</td>"
						response.write "<td align=""left"" class=""style27"" nowrap>"&trim(gInitDT(rsfound("BillFillDate")))&"&nbsp;</td>"
						response.write "<td align=""left"" class=""style27"" nowrap>"&trim(gInitDT(rsfound("DeallineDate")))&"&nbsp;</td>"
						response.write "<td align=""left"" class=""style27"" nowrap>"&trim(rsfound("Rule1"))&"&nbsp;</td>"
						response.write "<td align=""left"" class=""style27"" nowrap>"&trim(rsfound("BillUnitName"))&"&nbsp;</td>"
						response.write "<td align=""left"" class=""style27"" nowrap>"&trim(gInitDT(rsfound("JUDEDATE")))
											IF TRIM(GINITDT(RSFOUND("JUDEDATE")))<>"" AND TRIM(RSFOUND("JUDENO"))<>"" THEN
												RESPONSE.WRITE "／"
											END IF						
                        response.write trim(rsfound("JudeNo"))&"&nbsp;</td>"													
					else 
						response.write "<td align=""left"" class=""style27"">"&" "& "&nbsp;</td>"
						response.write "<td align=""left"" class=""style27"" nowrap>"&" "& "&nbsp;</td>"
						response.write "<td align=""left"" class=""style27"" nowrap>"&trim(rsfound("driveraddress"))& "&nbsp;</td>"					
						response.write "<td align=""left"" class=""style27"" nowrap>"& trim(rsfound("JudeNo"))& "&nbsp;</td>"
						
					end if

					
					
					if sys_City<>noshowcity then
						response.write "<td class=""style27"" nowrap>"&Sys_Payamount&"&nbsp;</td>"
					end if
					response.write "</tr>"
					
						if Not isnull(Sys_Payamount) then sumNT=sumNT+Cint(Sys_Payamount)
						filecnt=filecnt+1
						rsfound.MoveNext
						showDriver="":showDriverID=""
						if Not rsfound.eof then
							if trim(rsfound("DriverID"))<>trim(tmpDriverID) then
								showDriverID=trim(rsfound("DriverID"))
								showDriver=trim(rsfound("Driver"))
								if sys_City<>noshowcity then
									response.write "<tr><td align=""right"" colspan=""8"" class=""style27"">小計：</td>"
									response.write "<td align=""right"" class=""style27"">"&filecnt&"筆"&"</td>"
									response.write "<td align=""right"" class=""style27"">"&sumNT&"&nbsp;</td></tr>"
								end if
								cntNt=cntNt+sumNT
								tmpDriverID=trim(rsfound("DriverID"))
								sumNT=0
								filecnt=0
							end if
						end if
					
				wend
				if sys_City<>noshowcity then
					response.write "<tr><td align=""right"" colspan=""8"" class=""style27"">小計：</td>"
					response.write "<td align=""right"" class=""style27"">"&filecnt&"筆"&"</td>"
					response.write "<td align=""right"" class=""style27"">"&sumNT&"&nbsp;</td></tr>"
					cntNt=cntNt+sumNT
					response.write "<tr><td align=""right"" colspan=""8"" class=""style27"">共計：</td>"
					response.write "<td align=""right"" class=""style27"">"&sumFile&"筆"&"</td>"
					response.write "<td align=""right"" class=""style27"">"&cntNt&"&nbsp;</td></tr></table>"
				end if
				rsfound.close
				set rsfound=nothing
				%>
