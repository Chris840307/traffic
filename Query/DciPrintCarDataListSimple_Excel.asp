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
Server.ScriptTimeout = 800
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
	tmpSQL=request("SQLstr")
%>

</head>
<body>
<form name=myForm method="post">
	<table width="100%" border="1" cellpadding="4" cellspacing="1">
		

		<tr>
			<td width="22"></td>
			<td width="60" height="38">車號</td>
		
			<td width="60" class="style3">違規日</td>
			<td width="40" class="style3">時間</td>
		
			<td width="80">廠牌</td>
			<td width="35">顏色</td>
			<td width="80">車主姓名</td>
		<%if sys_City="花蓮縣" then %>
			<td width="145">車主地址</td>		
		<%end if%>
			<td width="125">違規地點</td>
		<%if sys_City<>"花蓮縣" then %>
			<td width="80">限制 / 實際</td>
		<%end if%>
		<%if sys_City<>"台中市" then %>
			<td width="80">行駕照狀態</td>
		<%end if%>
		</tr>
		<%	ListSN=0
			if Not rsfound.eof then rsfound.move DBcnt
			While Not rsfound.Eof
			ListSN=ListSN+1
%>				<tr bgcolor="#ffffff">
					<td height="55" align="left"><%=ListSN%></td>
					<td><%="&nbsp;"&rsfound("CarNo")%></td>
				
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
					'車輛廠牌
						if (trim(rsfound("A_Name"))<>"" and not isnull(rsfound("A_Name"))) then
							response.write funcCheckFont(rsfound("A_Name"),15,0)
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
						response.write funcCheckFont(rsfound("Owner"),15,0)
					end if
					%></td>
			<%if sys_City="花蓮縣" then %>
					<td><%
					'車主地址
					if (trim(rsfound("OwnerAddress"))<>"" and not isnull(rsfound("OwnerAddress"))) then
						response.write trim(rsfound("OwnerZip"))&funcCheckFont(rsfound("OwnerAddress"),15,0)
					end if
					%></td>
			<%end if%>
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
					
					<td><%
					'違規地點
					if trim(rsfound("IllegalAddress"))<>"" and not isnull(rsfound("IllegalAddress")) then
						response.write trim(rsfound("IllegalAddress"))
					else
						response.write "&nbsp;"
					end if
					%></td>
				<%if sys_City<>"花蓮縣" then %>
					<td><%
					'限速
					if trim(rsfound("RuleSpeed"))<>"" and not isnull(rsfound("RuleSpeed")) then
						response.write trim(rsfound("RuleSpeed")) & " / "  & trim(rsfound("IllegalSpeed"))
					else
						response.write "&nbsp;"
					end if
					%></td>
				<%end if%>
				<%if sys_City<>"台中市" then %>
					<td><%
					'行駕照狀態
					if trim(rsfound("DciCounterID"))<>"" and not isnull(rsfound("DciCounterID")) Then
						If trim(rsfound("DciCounterID"))="x" then
							response.write "<strong>駕照過期</strong>"
						ElseIf trim(rsfound("DciCounterID"))="y" Then 
							response.write "<strong>行照過期</strong>"
						ElseIf trim(rsfound("DciCounterID"))="v" Then 
							response.write "<strong>行駕照過期</strong>"
						Else
							response.write "&nbsp;"
						End If 
					else
						response.write "&nbsp;"
					end if
					%></td>
				<%end if%>
				
			
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