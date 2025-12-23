<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_舉發單核對清冊.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

	strSQL="select h.Illegaldate,h.IllegalAddress,h.Rule1,h.Rule2,h.Rule3,h.IllegalSpeed,h.RuleSpeed,a.SN,a.BillSN,a.RecordDate,a.ReturnMarkType,a.FileName,a.DCIReturnStatusID,a.ExchangeTypeID,a.DciErrorCarData,a.DCIErrorIDdata,b.ChName,a.BillNo,a.CarNo,a.BillTypeID,a.EXCHANGEDATE,a.RecordMemberID,a.seqNo,a.BatchNumber,c.Content as BillTypeName,d.DCIReturn,d.StatusContent,d.DCIRETURNSTATUS,e.DCIActionName,f.DCIreturn as CarErrorSN,f.StatusContent as CarErrorContent,g.DCIreturn as DCIErrorSN,g.StatusContent as DCIErrorContent,i.UnitName from (select * from DCILog"&request("strDCISQL")&") a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h,UnitInfo i where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN and h.billUnitID=i.UnitID "&request("TempSQL")&" order by a.ExchangeDate,a.BillNo"
	
	set rsfound=conn.execute(strSQL)

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing

If  sys_City="台南市" Then
	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

	strI="insert into Log values(log_sn.nextval+3000,360,"&Trim(Session("User_ID"))&",'"&Trim(Session("Ch_Name"))&"','"&userip&"',sysdate,'上傳下載資料查詢(匯出EXCEL):"&Replace(strSQL,"'","""")&"')"
	'response.write strI
	Conn.execute strI
End If 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單核對清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<body>

<table width="100%" border="1" cellpadding="4" cellspacing="1">
	<tr>
		<td align="center" colspan="12"><strong>舉發單核對清冊</strong></td>
	</tr>
	<tr>
		<td></td>
		<td width="70">舉發單號</td>
		<td width="70">車號</td>
		<td>車種</td>
		<td>廠牌</td>
		<td>顏色</td>
		<td>違規日期</td>
		<td>違規時間</td>
		<td>車主姓名</td>
		<td>違規地點</td>					
		<td>違規事實</td>
		<td>限制 / 實際</td>
	</tr>
	<%	i=0
		ReturnMarkType=split("3,4,5,Y",",")
		ReturnMarkName=Split("單退,寄存,公示,撤消",",")
		while Not rsfound.eof
			i=i+1
			
			response.write "<tr align='center'>"
			StrBass="select a.Owner,a.A_Name,a.DciReturnCarColor,c.ID as CarStatusID,c.Content as CarStatusName,d.ID as Rule4,d.Content as Rule4Name,e.DCIStationName,a.DciReturnCarType from (select * from BillBaseDCIReturn where EXCHANGETYPEID='A'  and CarNo='"&rsfound("CarNo")&"') a,(select ID,Content from DCICode where TypeID=10) c,(select ID,Content from DCICode where TypeID=10) d,Station e where a.DCIReturnCarStatus=c.ID(+) and a.Rule4=d.ID(+) and a.DCIReturnStation=e.DCIStationID(+)"
			set rsCarType=conn.execute(strBass)
			Sys_DciReturnCarColor="":Sys_DCIStationName="":Sys_A_Name="":Sys_CarStatusID="":Sys_CarStatusName="":Sys_Rule4="":Sys_Rule4Name="":Sys_CarColorID="":Sys_CarColorName="":Sys_CarTypeID=""
			if not rsCarType.eof then
				Sys_DciReturnCarColor=trim(rsCarType("DciReturnCarColor"))
				Sys_DCIStationName=trim(rsCarType("DCIStationName"))
				Sys_A_Name=trim(rsCarType("A_Name"))
				Sys_CarStatusID=trim(rsCarType("CarStatusID"))
				Sys_CarStatusName=trim(rsCarType("CarStatusName"))
				Sys_Rule4=trim(rsCarType("Rule4"))
				Sys_Rule4Name=trim(rsCarType("Rule4Name"))
				Sys_Owner=trim(rsCarType("Owner"))
				
				strCar="select * from DciCode where TypeID=5 and ID='"&Trim(rsCarType("DciReturnCarType"))&"'"
				Set rsCar=conn.execute(strCar)
				If Not rsCar.eof Then
					Sys_CarTypeID=trim(rsCar("Content"))
				End If
				rsCar.close
				Set rsCar=Nothing 
			end if
			rsCarType.close

			StrBass="select a.DciReturnCarColor,b.DCIStationName from (select * from BillBaseDCIReturn where EXCHANGETYPEID='W' and CarNo='"&trim(rsfound("CarNo"))&"' and BillNo='"&trim(rsfound("BillNo"))&"') a,Station b where a.DCIReturnStation=b.DCIStationID(+)"

			set rsCarType=conn.execute(strBass)
			if not rsCarType.eof then
				Sys_DciReturnCarColor=trim(rsCarType("DciReturnCarColor"))
				If trim(rsfound("ExchangeTypeID"))<>"A" then Sys_DCIStationName=trim(rsCarType("DCIStationName"))
			end if
			rsCarType.close

			if len(Sys_DciReturnCarColor)>1 then Sys_DciReturnCarColor=left(Sys_DciReturnCarColor,1)&","&right(Sys_DciReturnCarColor,1)
			if ifnull(Sys_DciReturnCarColor) then Sys_DciReturnCarColor=""
			Sys_CarColorID=split(Sys_DciReturnCarColor,",")
			for y=0 to ubound(Sys_CarColorID)
				strColor="select Content from DCICode where TypeID=4 and ID='"&Sys_CarColorID(y)&"'"
				set rscolor=conn.execute(strColor)
				if not rscolor.eof then
					if trim(Sys_CarColorName)<>"" then Sys_CarColorName=Sys_CarColorName&","
					Sys_CarColorName=Sys_CarColorName&trim(rscolor("Content"))
				end if
				rscolor.close
			Next
			response.write "<td>"&i&"</td>"
			response.write "<td>"&rsfound("BillNo")&"</td>"
			response.write "<td>"&rsfound("CarNo")&"</td>"
			response.write "<td>"&Sys_CarTypeID&"</td>"
			response.write "<td> "& Sys_A_Name &"</td>"
			response.write "<td>"& Sys_CarColorName &"</td>"


			response.write "<td>"&year(rsfound("illegaldate"))-1911&"/"&month(rsfound("illegaldate"))&"/"&day(rsfound("illegaldate"))&"&nbsp;</td>"
			response.write "<td>"&hour(rsfound("illegaldate"))&":"&minute(rsfound("illegaldate"))&"&nbsp;</td>"

			response.write "<td>"&Sys_Owner&"</td>"		
			response.write "<td align=""left""> "&trim(rsfound("IllegalAddress")) &"</td>"
			response.write "<td align=""left""> "
			response.write trim(rsfound("Rule1"))
'			strRule1="select * from Law where ItemID='"&trim(rsfound("Rule1"))&"' and Version=2"
'			Set rsR1=conn.execute(strRule1)
'			If Not rsR1.eof Then
'				response.write rsR1("IllegalRule")
'			End If
'			rsR1.close
'			Set rsR1=Nothing 
			If trim(rsfound("Rule2"))<>"" then
				response.write "<br>"&trim(rsfound("Rule2"))
'				strRule1="select * from Law where ItemID='"&trim(rsfound("Rule2"))&"' and Version=2"
'				Set rsR1=conn.execute(strRule1)
'				If Not rsR1.eof Then
'					response.write rsR1("IllegalRule")
'				End If
'				rsR1.close
'				Set rsR1=Nothing 
			End If 
			If trim(rsfound("Rule3"))<>"" then
				response.write "<br>"&trim(rsfound("Rule3"))
'				strRule1="select * from Law where ItemID='"&trim(rsfound("Rule3"))&"' and Version=2"
'				Set rsR1=conn.execute(strRule1)
'				If Not rsR1.eof Then
'					response.write rsR1("IllegalRule")
'				End If
'				rsR1.close
'				Set rsR1=Nothing 
			End If 
			response.write "</td>"
			response.write "<td align=""left""> "&trim(rsfound("RuleSpeed"))&" / "&trim(rsfound("IllegalSpeed")) &"&nbsp;</td>"




	
			response.write "</tr>"
			rsfound.movenext
		wend
	%>
</table>
</body>
</html>
<%conn.close%>