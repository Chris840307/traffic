<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>車籍資料列表</title>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
'權限
'AuthorityCheck(234)

fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_車籍資料清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<%
Server.ScriptTimeout = 60800
Response.flush

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

RecordDate=split(gInitDT(date),"-")
	dcitype=trim(request("dcitype"))
	
	OrderAdd=""
	If  sys_City="花蓮縣" Then
		If Trim(Session("Ch_Name"))="停管入案" Then
			OrderAdd="e.CarNo,"
		End If 
	End If 
	strSQL="select distinct c.IllegalDate,c.SN,c.CarSimpleID,c.Rule1,c.Rule2,c.Rule3,c.Rule4,c.RuleVer,c.IllegalAddress,c.RuleSpeed,c.IllegalSpeed,c.RecordStateID,c.RecordDate,e.BillNo,e.CarNo,e.A_Name,e.DCIReturnCarType,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.Nwner,e.NwnerID,e.NwnerAddress,e.NwnerZip,e.DCIReturnCarStatus,e.DciCounterID from (select * from DCILog "&Request ("strDCISQL")&") a,MemberData b,BillBase c,DCIReturnStatus d,BillBaseDCIReturn e where a.BillSN=c.SN and e.ExchangeTypeID='A' and e.Status='S' and c.CarNo=e.CarNo (+) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and c.RecordStateID=0 "&request("SQLstr")&" order by "&OrderAdd&"c.RecordDate"
	set rsfound=conn.execute(strSQL)

	strCnt="select count(*) as cnt from (select distinct c.IllegalDate,c.SN,c.CarSimpleID,c.Rule1,c.Rule2,c.Rule3,c.Rule4,c.RuleVer,c.IllegalAddress,c.RuleSpeed,c.IllegalSpeed,c.RecordStateID,c.RecordDate,e.BillNo,e.CarNo,e.DCIReturnCarType,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerAddress,e.OwnerZip,e.DCIReturnCarStatus,e.DciCounterID from (select * from DCILog "&Request ("strDCISQL")&") a,MemberData b,BillBase c,DCIReturnStatus d,BillBaseDCIReturn e where a.BillSN=c.SN and e.ExchangeTypeID='A' and e.Status='S' and c.CarNo=e.CarNo (+) and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and c.RecordStateID=0 "&request("SQLstr")&")"
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	Dbrs.close
	tmpSQL=strwhere
%>

</head>
<body>
<form name=myForm method="post">
	<table width="100%" border="1" cellpadding="4" cellspacing="1">
		<tr>
			<td colspan="<%
					if sys_City<>"花蓮縣" then
						if trim(Session("SpecUser"))="1" then
							colcount="15"
						else
							colcount="14"
						end if
					else
						if trim(Session("SpecUser"))="1" then
							colcount="12"
						else
							colcount="11"
						end if
					end if
					response.write colcount
					%>" align="center">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr>
					<td colspan="<%
						response.write colcount
					%>" align="center">
						<font size="3"><strong>車籍資料清冊</strong></font>
						(共 <%=DBsum%> 筆)
					</td>
				</tr>
				<tr>
					<td colspan="<%
						response.write colcount
					%>" align="right">
						印表單位：<%
						UnitID=Session("Unit_ID")
						strUnit="select UnitName from UnitInfo where UnitID='"&UnitID&"'"
						set rsUnit=conn.execute(strUnit)
						if not rsUnit.eof then
							response.write trim(rsUnit("UnitName"))
						end if
						rsUnit.close
						set rsUnit=nothing
						%>
					</td>
				</tr>
				<tr>
					<td colspan="<%
						response.write colcount
					%>" align="right">
						印表時間：<%=year(now)-1911%> - <%=month(now)%> - <%=day(now)%> - <%=hour(now)%> : <%=minute(now)%>
					</td>
				</tr>
			</table>
			</td>
		</tr>

		<tr>
			<td width="38"></td>
			<td width="60" height="38">車號</td>
		<%if sys_City<>"花蓮縣" then %>
			<td width="38">牌類</td>
		<%end if%>
			<td width="80" class="style3">違規日期</td>
			<td width="50" class="style3">時間</td>
		
			<td width="50">車別</td>
			<td width="50">廠牌</td>
			<td width="50">顏色</td>
			<td width="65">車主姓名</td>
			<td width="150">車主地址</td>
			<!-- <td width="65">原車主姓名</th>
			<td width="250">原車主地址</th>
			<td width="250">駕駛人戶籍地址</td> -->
			<td width="200">違規地點</td>
		<%if sys_City<>"花蓮縣" then %>
			<td width="50">限速、重</td>
			<td width="50">車速、重</td>
		<%end if%>
		<%if trim(Session("SpecUser"))="1" then%>
			<td width="45">業管車</td>
		<%end if%>
			<td width="61">車籍狀態</td>
		<%if sys_City<>"台中市" and sys_City<>"嘉義縣" then %>
			<td width="80">行駕照狀態</td>
		<%end if%>
		<%if sys_City<>"花蓮縣" then %>
			<td width="105">處理狀態</td>
		<%end if%>
		<%if sys_City="花蓮縣" Or sys_City="嘉義縣" then %>
			<td>違規事實</td>
		<%end if%>
		</tr>
		<%	ListSN=0
			if Not rsfound.eof then rsfound.move DBcnt
			While Not rsfound.Eof
			ListSN=ListSN+1
%>				<tr bgcolor="#ffffff">
					<td height="38" align="left"><%=ListSN%></td>
					<td><%="&nbsp;"&rsfound("CarNo")%></td>
				<%if sys_City<>"花蓮縣" then %>
					<td><%
					'簡式車種
					if trim(rsfound("CarSimpleID"))="1" then
						response.write "汽車"
					elseif trim(rsfound("CarSimpleID"))="2" then
						response.write "拖車"
					elseif trim(rsfound("CarSimpleID"))="3" then
						response.write "重機"
					elseif trim(rsfound("CarSimpleID"))="4" then
						response.write "輕機"
					end if								
					%></td>
				<%end if%>
			<%'---------------------------------------------%>
			<td class="style3"><%
			'違規日期
			if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
				'smith > don't delet the point , if delete , user will auto change to 19xx format
				response.write "&nbsp;"& year(rsfound("IllegalDate"))-1911&"/ "&month(rsfound("IllegalDate"))&"/"&day(rsfound("IllegalDate"))
			end if
			%></td>
			<td class="style3"><%
			'時間
			
			if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
				if len(hour(rsfound("IllegalDate"))) < 2 then 
					sHour = "0" & hour(rsfound("IllegalDate"))
				else
					sHour = hour(rsfound("IllegalDate"))	
				end if
				if len(minute(rsfound("IllegalDate"))) < 2 then 
					sMinute = "0" & minute(rsfound("IllegalDate"))
				else
					sMinute = minute(rsfound("IllegalDate"))	
				end if
				response.write sHour&":"&sMinute
			end if
			%></td>
		<%'---------------------------------------------			%>				
					<td><%
					'詳細車種
					if trim(rsfound("DCIReturnCarType"))<>"" and not isnull(rsfound("DCIReturnCarType")) then
						strCType="select * from DCIcode where TypeID=5 and ID='"&trim(rsfound("DCIReturnCarType"))&"'"
						set rsCType=conn.execute(strCType)
						if not rsCType.eof then
							response.write trim(rsCType("Content"))
						end if
						rsCType.close
						set rsCType=nothing
					end if								
					%></td>
					<td width="6%"><%
					'車輛廠牌
						if (trim(rsfound("A_Name"))<>"" and not isnull(rsfound("A_Name"))) then
							response.write funcCheckFont(rsfound("A_Name"),19,0)
						end if
					%></td>
					<td><%
					'車輛顏色
					if trim(rsfound("DCIReturnCarColor"))<>"" and not isnull(rsfound("DCIReturnCarColor")) then
						ColorLen=cint(Len(rsfound("DCIReturnCarColor")))
						for Clen=1 to ColorLen
							colorID=mid(rsfound("DCIReturnCarColor"),Clen,1)
							strColor="select * from DCIcode where TypeID=4 and ID='"&trim(colorID)&"'"
							set rsColor=conn.execute(strColor)
							if not rsColor.eof then
								response.write trim(rsColor("Content"))
							end if
							rsColor.close
							set rsColor=nothing
						next
					end if
					%></td>
					<td><%
					if trim(rsfound("Owner"))<>"" then
						response.write funcCheckFont(rsfound("Owner"),20,0)
					end if
					%></td>
					<td><%
					'車主地址
					if (trim(rsfound("OwnerAddress"))<>"" and not isnull(rsfound("OwnerAddress"))) then
						response.write trim(rsfound("OwnerZip"))&funcCheckFont(rsfound("OwnerAddress"),19,0)
					end if
					%></td>
					<!-- <td> --><%'=rsfound("Nwner")%><!-- </td> -->
					<!-- <td> --><%
					'原車主地址
					'if (trim(rsfound("NwnerAddress"))<>"" and not isnull(rsfound("NwnerAddress"))) then
					'	response.write trim(rsfound("NwnerZip"))&trim(rsfound("NwnerAddress"))
					'end if
					%><!-- </td> -->
					<!-- <td> --><%
					'駕駛人地址
					'if (trim(rsfound("DriverHomeAddress"))<>"" and not isnull(rsfound("DriverHomeAddress"))) then
					'	response.write trim(rsfound("DriverHomeZip"))&trim(rsfound("DriverHomeAddress"))
					'end if
					%><!-- </td> -->
					
					<td width="10%"><%
					'違規地點
					if trim(rsfound("IllegalAddress"))<>"" and not isnull(rsfound("IllegalAddress")) Then
						if sys_City="嘉義縣" Then
							response.write Replace(trim(rsfound("IllegalAddress")),"嘉義縣","")
						else
							response.write trim(rsfound("IllegalAddress"))
						End If 
					else
						response.write "&nbsp;"
					end if
					%></td>
				<%if sys_City<>"花蓮縣" then %>
					<td width="6%"><%
					'限速
					if trim(rsfound("RuleSpeed"))<>"" and not isnull(rsfound("RuleSpeed")) then
						response.write trim(rsfound("RuleSpeed"))
					else
						response.write "&nbsp;"
					end if
					%></td><td width="6%"><%
					'車速
					if trim(rsfound("IllegalSpeed"))<>"" and not isnull(rsfound("IllegalSpeed")) then
						response.write trim(rsfound("IllegalSpeed"))
					else
						response.write "&nbsp;"
					end if
					%></td>
				<%end if%>
				<%if trim(Session("SpecUser"))="1" then%>
					<td align="center"><%
					'業管車
					if sys_City="花蓮縣" then 
						strVip="select * from SpecCar where RecordStateID=0"
						set rsVip=conn.execute(strVip)
						If Not rsVip.Bof Then rsVip.MoveFirst 
						While Not rsVip.Eof
							if instr(trim(rsfound("Owner")),trim(rsVip("CarNo")))>0 then
								response.write "＊"
							end if
						rsVip.MoveNext
						Wend
						rsVip.close
						set rsVip=nothing
					else
						strVip="select Count(*) as cnt from SpecCar where CarNo='"&trim(rsfound("CarNo"))&"' and RecordStateID=0"
						set rsVip=conn.execute(strVip)
						if cint(trim(rsVip("cnt"))) > 0 then
							response.write "＊"
						end if
						rsVip.close
						set rsVip=nothing
					end if
					%></td>
				<%end if%>
					<td><%
					'車籍狀態
						if trim(rsfound("DCIReturnCarStatus"))<>"" and not isnull(rsfound("DCIReturnCarStatus")) then
							strCstatus="select Content from DCIcode where TypeID=10 and ID='"&trim(rsfound("DCIReturnCarStatus"))&"'"
							set rsCS=conn.execute(strCstatus)
							if not rsCS.eof then
								response.write trim(rsCS("Content"))
							end if 
							rsCS.close
							set rsCS=nothing
						end if
					%></td>
				<%if sys_City<>"台中市" And sys_City<>"嘉義縣" then %>
					<td><%
					'行駕照狀態
						if trim(rsfound("DciCounterID"))<>"" and not isnull(rsfound("DciCounterID")) then
							If trim(rsfound("DciCounterID"))="x" Then
								 response.write "<strong>駕照過期</strong>"
							ElseIf trim(rsfound("DciCounterID"))="y" Then
								response.write "<strong>行照過期</strong>"
							ElseIf trim(rsfound("DciCounterID"))="v" Then
								response.write "<strong>行駕照過期</strong>"
							End If 
						end if
					%></td>
				<%end if%>
				<%if sys_City<>"花蓮縣" then %>
					<td><%
					'處理狀態
					strStatus="select ExchangeTypeID,DCIReturnStatusID from DCILog where BillSN="&trim(rsfound("SN"))&" order by ExchangeDate Desc"
					set rsStatus=conn.execute(strStatus)
					if not rsStatus.eof then
						strSID="select StatusContent from DCIReturnStatus where DCIactionId='"&trim(rsStatus("ExchangeTypeID"))&"' and DCIreturn='"&trim(rsStatus("DCIReturnStatusID"))&"'"
						set rsSID=conn.execute(strSID)
						if not rsSID.eof then
							response.write trim(rsSID("StatusContent"))
						end if
						rsSID.close
						set rsSID=nothing
					end if
					rsStatus.close
					set rsStatus=nothing
					%></td>
				<%end if%>
				<%if sys_City="花蓮縣" Or sys_City="嘉義縣" then %>
					<td align="left" width="22%"><%
					if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
						response.write trim(rsfound("Rule1"))
'						strCarImple=""
'						if left(trim(rsfound("Rule1")),4)="2110" then
'							if trim(rsfound("CarSimpleID"))=1 or trim(rsfound("CarSimpleID"))=2 then
'								strCarImple=" and CarSimpleID in ('5','0')"
'							elseif trim(rsfound("CarSimpleID"))=3 or trim(rsfound("CarSimpleID"))=4 then
'								strCarImple=" and CarSimpleID in ('3','0')"
'							else
'								strCarImple=""
'							end if
'						end if
'						strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rsfound("Rule1"))&"' and Version='"&trim(rsfound("RuleVer"))&"'"&strCarImple
'						set rsR1=conn.execute(strR1)
'						if not rsR1.eof then 
'							response.write " "&trim(rsR1("IllegalRule"))
'						end if
'						rsR1.close
'						set rsR1=nothing
					end if
					if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
						response.write "<br>"&trim(rsfound("Rule2"))
						strCarImple=""
'						if left(trim(rsfound("Rule2")),4)="2110" then
'							if trim(rsfound("CarSimpleID"))=1 or trim(rsfound("CarSimpleID"))=2 then
'								strCarImple=" and CarSimpleID in ('5','0')"
'							elseif trim(rsfound("CarSimpleID"))=3 or trim(rsfound("CarSimpleID"))=4 then
'								strCarImple=" and CarSimpleID in ('3','0')"
'							else
'								strCarImple=""
'							end if
'						end if
'
'						strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rsfound("Rule2"))&"' and Version='"&trim(rsfound("RuleVer"))&"'"&strCarImple
'						set rsR1=conn.execute(strR1)
'						if not rsR1.eof then 
'							response.write " "&trim(rsR1("IllegalRule"))
'						end if
'						rsR1.close
'						set rsR1=nothing
					end if
					if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
						response.write "<br>"&trim(rsfound("Rule3"))
						strCarImple=""
					end if
					if (trim(rsfound("RuleSpeed"))<>"" and not isnull(rsfound("RuleSpeed"))) and (trim(rsfound("IllegalSpeed"))<>"" and not isnull(rsfound("IllegalSpeed"))) then
						response.write "<br>速限"&trim(rsfound("RuleSpeed"))&"公里時速"&trim(rsfound("IllegalSpeed"))&"公里，超速"&trim(rsfound("IllegalSpeed"))-trim(rsfound("RuleSpeed"))&"公里"
					end if
					%></td>
				<%end if%>
				</tr>
<%
			rsfound.MoveNext
		Wend
		rsfound.close
		set rsfound=nothing
		%>
		</tr>
	</table>
</form>
</body>
</html>
<%conn.close%>