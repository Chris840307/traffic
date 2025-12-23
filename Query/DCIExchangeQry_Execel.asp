<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_資料交換紀錄.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

	strSQL="select a.SN,a.BillSN,a.RecordDate,a.ReturnMarkType,a.FileName,a.DCIReturnStatusID,a.ExchangeTypeID,a.DciErrorCarData,a.DCIErrorIDdata,b.ChName,a.BillNo,a.CarNo,a.BillTypeID,h.illegaladdressid,h.illegaladdress,a.EXCHANGEDATE,a.RecordMemberID,a.seqNo,a.BatchNumber,c.Content as BillTypeName,d.DCIReturn,d.StatusContent,d.DCIRETURNSTATUS,e.DCIActionName,f.DCIreturn as CarErrorSN,f.StatusContent as CarErrorContent,g.DCIreturn as DCIErrorSN,g.StatusContent as DCIErrorContent,i.UnitName from (select * from DCILog"&request("strDCISQL")&") a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g,BillBase h,UnitInfo i where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=h.SN and h.billUnitID=i.UnitID "&request("TempSQL")&" order by a.BillSN,a.ExchangeDate,a.BillNo"
	
	set rsfound=conn.execute(strSQL)

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing

If  sys_City="台南市" Then
	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

	strI="insert into Log values((select max(sn)+1 from log),360,"&Trim(Session("User_ID"))&",'"&Trim(Session("Ch_Name"))&"','"&userip&"',sysdate,'上傳下載資料查詢(匯出EXCEL):"&Replace(strSQL,"'","""")&"')"
	'response.write strI
	Conn.execute strI
End If 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>DCI 資料交換紀錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" border="1">
	<tr>
		<td align="center"><strong>DCI 資料交換紀錄</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<td>建檔日期</td>
					<td>建檔人員</td>
					<td>舉發單位</td>
					<td>類別</td>					
					<td>舉發單號</td>
					<td>車號</td>
					<%If sys_City="基隆市" Then%>
						<td>違規地點</td>
					<%end if%>
					<td>DCI作業</td>
					<td>應到案處所</td>
					<td>廠牌</td>
					<td>顏色</td>
					<td>車藉狀況</td>
					<td>結果</td>
					<td>訊息</td>
					<td>操作</td>
					<td>交換時間</td>					
					<td>檔案名稱/序號</td>
					<td>批號欄位</td>					
				</tr>
				<%
					ReturnMarkType=split("3,4,5,Y",",")
					ReturnMarkName=Split("單退,寄存,公示,撤消",",")
					while Not rsfound.eof
						response.write "<tr bgcolor='#FFFFFF' align='center'>"
						response.write "<td style='mso-number-format:""\@"";'>"&gInitDT(trim(rsfound("RecordDate")))&hour(rsfound("RecordDate"))&"</td>"
						response.write "<td>"&rsfound("ChName")&"</td>"
						response.write "<td align=""left"">"&rsfound("UnitName")&"</td>"
						response.write "<td>"&rsfound("BillTypeName")&"</td>"						
						response.write "<td>"&rsfound("BillNo")&"</td>"
						response.write "<td style='mso-number-format:""\@"";'>"&rsfound("CarNo")&"</td>"
						
						If sys_City="基隆市" Then
							response.write "<td>"
								If trim(rsfound("illegaladdressid")) <> "" Then Response.Write trim(rsfound("illegaladdressid"))&" - "
								Response.Write trim(rsfound("illegaladdress"))
							Response.Write "</td>"
						end if

						if trim(rsfound("ExchangeTypeID"))="N" then
							response.write "<td>"
							for arr=0 to Ubound(ReturnMarkType)
								if trim(ReturnMarkType(arr))=trim(rsfound("ReturnMarkType")) then
									response.write ReturnMarkName(arr)
									exit for
								end if
							next
							if arr>Ubound(ReturnMarkType) then response.write "送達註記"
							response.write "&nbsp;</td>"
						else
							response.write "<td>"&rsfound("DCIActionName")&"</td>"
						end if

						'StrBass="select b.Content as CarTypeName,c.Content as CarColor,d.Content as Rule4Name from BillBaseDCIReturn a,(select ID,Content from DCICode where TypeID=5) b,(select ID,Content from DCICode where TypeID=4) c,(select ID,Content from DCICode where TypeID=10) d where a.DciReturnCarType=b.ID(+) and a.DciReturnCarColor=c.ID(+) and a.Rule4=d.ID(+) and a.BillNo='"&rsfound("BillNo")&"' and a.CarNo='"&rsfound("CarNo")&"'"

						StrBass="select a.A_Name,a.DciReturnCarColor,c.ID as CarStatusID,c.Content as CarStatusName,d.ID as Rule4,d.Content as Rule4Name,e.DCIStationName from (select * from BillBaseDCIReturn where EXCHANGETYPEID='A'  and CarNo='"&rsfound("CarNo")&"') a,(select ID,Content from DCICode where TypeID=10) c,(select ID,Content from DCICode where TypeID=10) d,Station e where a.DCIReturnCarStatus=c.ID(+) and a.Rule4=d.ID(+) and a.DCIReturnStation=e.DCIStationID(+)"
						set rsCarType=conn.execute(strBass)
						Sys_DciReturnCarColor="":Sys_DCIStationName="":Sys_A_Name="":Sys_CarStatusID="":Sys_CarStatusName="":Sys_Rule4="":Sys_Rule4Name="":Sys_CarColorID="":Sys_CarColorName=""
						if not rsCarType.eof then
							Sys_DciReturnCarColor=trim(rsCarType("DciReturnCarColor"))
							Sys_DCIStationName=trim(rsCarType("DCIStationName"))
							Sys_A_Name=trim(rsCarType("A_Name"))
							Sys_CarStatusID=trim(rsCarType("CarStatusID"))
							Sys_CarStatusName=trim(rsCarType("CarStatusName"))
							Sys_Rule4=trim(rsCarType("Rule4"))
							Sys_Rule4Name=trim(rsCarType("Rule4Name"))
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
						next


						response.write "<td align=""left""> "&Sys_DCIStationName &"</td>"
						response.write "<td> "& Sys_A_Name &"</td>"
						response.write "<td>"& Sys_CarColorName &"</td>"
						response.write "<td align=""left"">"&Sys_Rule4Name&"</td>"
		


						if trim(rsfound("DCIRETURNSTATUS"))="1" then
							response.write "<td>正常</td>"
						elseif trim(rsfound("DCIRETURNSTATUS"))="-1" then
							response.write "<td><font color=""red"">異常</font></td>"
						else
							response.write "<td>未處理</td>"
						end if

						DCIerror="":dciSQL=""
						if trim(rsfound("DCIReturnStatusID"))="00" then
							if trim(rsfound("DciErrorCarData"))<>"" then
								dciSQL="'"&rsfound("DciErrorCarData")&"'"
							end if
							if trim(rsfound("DCIErrorIDdata"))<>"" then
								if trim(dciSQL)<>"" then
									dciSQL=dciSQL&",'"&rsfound("DCIErrorIDdata")&"'"
								else
									dciSQL="'"&rsfound("DCIErrorIDdata")&"'"
								end if
							end if
							if trim(dciSQL)<>"" then
								strSQL="select DCIReturn,StatusContent from DCIReturnStatus where DCIActionID='"&rsfound("ExchangeTypeID")&"E' and DCIReturn in("&dciSQL&")"
								set rsdci=conn.execute(strSQL)
								while Not rsdci.eof
									if trim(DCIerror)<>"" then DCIerror=trim(DCIerror)&","
									DCIerror=trim(DCIerror)&rsdci("DCIReturn")&". "&rsdci("StatusContent")
									rsdci.movenext
								wend
								rsdci.close
							end if
						end if
						if trim(rsfound("BillTypeID"))="2" then
							strSQL="select ID,Content from DCICode where TypeID=10 and ID in(Select Rule4 from BillBaseDCIReturn where BillNo='"&rsfound("BillNo")&"' and CarNo='"&rsfound("CarNo")&"')"

							set rsdci=conn.execute(strSQL)
							while Not rsdci.eof
								if trim(DCIerror)<>"" then DCIerror=trim(DCIerror)&","
								DCIerror=trim(DCIerror)&rsdci("ID")&". "&rsdci("Content")
								rsdci.movenext
							wend
							rsdci.close
						end if

						Message=rsfound("DCIReturn")&". "&rsfound("StatusContent")
						'if trim(DCIerror)<>"" then Message=Message&"<br>"&DCIerror
						if trim(rsfound("CarErrorSN"))<>"" then Message=Message&"<br>"&rsfound("CarErrorSN")&". "&rsfound("CarErrorContent")
						if trim(rsfound("DCIErrorSN"))<>"" then Message=Message&"<br>"&rsfound("DCIErrorSN")&". "&rsfound("DCIErrorContent")

						response.write "<td class=""font10"" nowrap>"
						response.write Message
						response.write "</td>"

						if trim(rsfound("ExchangeTypeID"))="E" then
							response.write "<td>"
							strDelReason="select * from BillDeleteReason a,DciCode b " &_
								" where a.BillSN="&trim(rsfound("BillSn"))&" and b.TypeID=3" &_
								" and a.DelReason=b.ID"
							set rsDelReason=conn.execute(strDelReason)
							if not rsDelReason.eof then
								response.write trim(rsDelReason("Content"))
								if trim(rsDelReason("Note"))<>"" then
									response.write "("&trim(rsDelReason("Note"))&")"
								end if
							end if
							rsDelReason.close
							set rsDelReason=nothing
							response.write "</td>"
						else
							response.write "<td></td>"
						end if

						response.write "<td>"&rsfound("EXCHANGEDATE")&"</td>"						
						response.write "<td>"&trim(rsfound("FileName"))&"&nbsp;<font color=""Red"">"&trim(rsfound("seqNo"))&"</font></td>"
						response.write "<td>"&rsfound("BatchNumber")&"</td>"						
						response.write "</tr>"
						rsfound.movenext
					wend
				%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close%>