<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_舉發單資料.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<%
Server.ScriptTimeout = 65000
Response.flush
%>
<%
	'檢查是否可進入本系統
	'AuthorityCheck(234)

	strSQL=Session("BillSQL")
	'response.write  strSQL
	'response.end
	set rsfound=conn.execute(strSQL)

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing

If  sys_City="台南市" Then
	userip = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
	If trim(userip) = "" Then userip = Request.ServerVariables("REMOTE_ADDR") 

	strI="insert into Log values((select nvl(max(Sn),0)+1 from Log),360,"&Trim(Session("User_ID"))&",'"&Trim(Session("Ch_Name"))&"','"&userip&"',sysdate,'舉發單資料維護(匯出EXCEL):"&Replace(strSQL,"'","""")&"')"
	'response.write strI
	Conn.execute strI
End If 

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單查詢</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" border="1">
	<tr>
		<td align="center"><strong>舉發單紀錄</strong></td>
	</tr>
	<tr>
		<td>
			<table width="95%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<td>類別</td>
					<td>違規日期</td>
					<td>違規時間</td>
					<td>舉發單號</td>
					<td>車號</td>
					<td>簡式車種</td>
					
					<td>駕駛人ID</td>
					<td>駕駛人姓名</td>
					<td>車主姓名</td>
					<td>詳細車種</td>
					<td>舉發員警</td>
				<%if sys_City = "苗栗縣" then %>
					<td>舉發上層單位</td>
				<%End If %>
					<td>舉發單位</td>
					<td>違規地點</td>
					<td>法條一</td>
				<%if sys_City = "苗栗縣" Or sys_City = "彰化縣" then %>
					<td>違規事實一</td>
				<%End If %>
					<td>法條二</td>
				<%if sys_City = "苗栗縣" Or sys_City = "彰化縣" then %>
					<td>違規事實二</td>
				<%End If %>
					<td>罰款一</td>
					<td>罰款二</td>					
					<td>填單日期</td>
					<td>應到案日期</td>
					<td>應到案處所</td>
					<td>建檔日期</td>
					<td>入案日期</td>
					<td>代保管物件</td>
				<%if sys_City = "台中市" then %>
					<td>告示單號</td>
				<%End If %>
				<%if sys_City = "南投縣" then %>
					<td>交通事故種類</td>
					<td>民眾檢舉日期</td>
				<%End If %>
				<%if sys_City = "台東縣" then %>
					<td>入案批號</td>
				<%End If %>
				<%if sys_City = "彰化縣" then %>
					<td>郵寄日期</td>
					<td>送達日期</td>
				<%End If %>
					<td>操作</td>
					<!-- <th width="6%">罰款</th> -->
					<!-- <th width="8%">DCI</th> -->
				</tr>
				<%
					If Not rsfound.Bof Then rsfound.MoveFirst 
					While Not rsfound.Eof
						Response.flush
						chname="":chRule="":ForFeit=""
						if rsfound("BillMem1")<>"" then	chname=rsfound("BillMem1")
						if rsfound("BillMem2")<>"" then	chname=chname&"/"&rsfound("BillMem2")
						if rsfound("BillMem3")<>"" then	chname=chname&"/"&rsfound("BillMem3")
						if rsfound("BillMem4")<>"" then	chname=chname&"/"&rsfound("BillMem4")

						response.write "<tr bgcolor='#FFFFFF' align='center' "
						response.write ">"
					'類別
						response.write "<td>"
					if trim(rsfound("BillBaseTypeID"))="0" then
						strBTypeVal="select Content from DCIcode where TypeID=2 and ID='"&trim(rsfound("BillTypeID"))&"'"
						set rsBTypeVal=conn.execute(strBTypeVal)
						if not rsBTypeVal.eof then
							response.write rsBTypeVal("Content")
						end if
						rsBTypeVal.close
						set rsBTypeVal=nothing
					else
						if trim(rsfound("BillTypeID"))="1" then
							response.write "慢車行人道路障礙"
						elseif trim(rsfound("BillTypeID"))="2" then
							response.write "行人"
						elseif trim(rsfound("BillTypeID"))="3" then
							response.write "道路障礙"
						end if
					end if
						response.write "</td>"
					'違規日期
						response.write "<td>"
				if sys_City = "台南市" then 
						if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
							response.write rsfound("IllegalDate")
						else
							response.write "&nbsp;"
						end if
				else
						if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
							response.write gInitDT(rsfound("IllegalDate"))&"&nbsp;"
						else
							response.write "&nbsp;"
						end if
				end if
						response.write "</td>"
					'違規時間
						response.write "<td>"
						if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
							response.write Right("00"&hour(trim(rsfound("IllegalDate"))),2)&Right("00"&minute(trim(rsfound("IllegalDate"))),2)&"&nbsp;"
						else
							response.write "&nbsp;"
						end if
						response.write "</td>"
					'舉發單號
						response.write "<td>"&rsfound("BillNo")&"</td>"
					'車號
						response.write "<td>"&rsfound("CarNo")&"&nbsp;</td>"
					'簡式車種1汽車 / 2拖車/ 3重機/ 4輕機 
						response.write "<td>"
						if trim(rsfound("CarSimpleID"))="1" then
							response.write "汽車"
						elseif trim(rsfound("CarSimpleID"))="2" then
							response.write "拖車"
						elseif trim(rsfound("CarSimpleID"))="3" then
							response.write "重機"
						elseif trim(rsfound("CarSimpleID"))="4" then
							response.write "輕機"
						elseif trim(rsfound("CarSimpleID"))="5" then
							response.write "動力機械"
						elseif trim(rsfound("CarSimpleID"))="6" then
							response.write "臨時車牌"
						end if
						response.write "</td>"
					'駕駛人ID
						response.write "<td>"
						if trim(rsfound("DriverID"))<>"" and not isnull(rsfound("DriverID")) then
							response.write trim(rsfound("DriverID"))
						else
							response.write "&nbsp;"
						end if
						response.write "</td>"
					'駕駛人姓名
						OwnerName=""
						CarType=""
						DCICaseInDate=""
						response.write "<td>"
						strDName="select Driver,Owner,DciReturnCarType,DCICaseInDate from BillBaseDciReturn where BillNo='"&trim(rsfound("BillNo"))&"' and CarNo='"&trim(rsfound("CarNo"))&"' and ExchangeTypeID='W'"
						set rsDName=conn.execute(strDName)
						if not rsDName.eof then
							response.write rsDName("Driver")
							OwnerName=trim(rsDName("Owner"))
							CarType=trim(rsDName("DciReturnCarType"))
							DCICaseInDate=trim(rsDName("DCICaseInDate"))
						else
							response.write "&nbsp;"
						end if
						rsDName.close
						set rsDName=nothing
						response.write "</td>"
					'車主姓名
						response.write "<td>"
							if OwnerName="" then
								response.write "&nbsp;"
							else
								response.write funcCheckFont(OwnerName,20,0)
							end if
						response.write "</td>"
					'車種
						response.write "<td>"
							if CarType<>"" then
								strCarType="select Content from DciCode where TypeID='5' and ID='"&CarType&"'"
								set rsCarType=conn.execute(strCarType)
								if not rsCarType.eof then
									response.write trim(rsCarType("Content"))
								end if
								rsCarType.close
								set rsCarType=nothing
							else
								response.write "&nbsp;"
							end if
						response.write "</td>"
					'舉發員警
						response.write "<td>"&chname&"</td>"
					if sys_City = "苗栗縣" Then
						response.write "<td>"
						if trim(rsfound("BillUnitID"))<>"" and not isnull(rsfound("BillUnitID")) then
							strUnit="select (select UnitName from UnitInfo where UnitID=a.UnitTypeID) as UnitName from UnitInfo a where a.UnitID='"&trim(rsfound("BillUnitID"))&"'"
							set rsUnit=conn.execute(strUnit)
							if not rsUnit.eof then
								response.write trim(rsUnit("UnitName"))
							end if
							rsUnit.close
							set rsUnit=nothing
						else
							response.write "&nbsp;"
						end If
						response.write "</td>"
					End If 
					'舉發單位
						response.write "<td>"
						if trim(rsfound("BillUnitID"))<>"" and not isnull(rsfound("BillUnitID")) then
							strUnit="select UnitName from UnitInfo where UnitID='"&trim(rsfound("BillUnitID"))&"'"
							set rsUnit=conn.execute(strUnit)
							if not rsUnit.eof then
								response.write trim(rsUnit("UnitName"))
							end if
							rsUnit.close
							set rsUnit=nothing
						else
							response.write "&nbsp;"
						end if
						response.write "</td>"
					'違規地點
						response.write "<td>"
						response.write rsfound("IllegalAddress")
						response.write "</td>"
					'法條一
						response.write "<td>"
						response.write rsfound("Rule1")
						response.write "</td>"
					if sys_City = "苗栗縣" Or sys_City = "彰化縣" then
						response.write "<td>"
						strCarImple=""
						if left(trim(rsfound("Rule1")),4)="2110" or left(trim(rsfound("Rule1")),4)="2210" or trim(rsfound("Rule1"))="4310102" or trim(rsfound("Rule1"))="4310103" or trim(rsfound("Rule1"))="4310104" then
							if trim(rsfound("CarSimpleID"))=1 or trim(rsfound("CarSimpleID"))=2 then
								strCarImple=" and CarSimpleID in ('5','0')"
							elseif trim(rsfound("CarSimpleID"))=3 or trim(rsfound("CarSimpleID"))=4 then
								strCarImple=" and CarSimpleID in ('3','0')"
							else
								strCarImple=""
							end if
						end if
						strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rsfound("Rule1"))&"' and Version='2'"&strCarImple&" order by CarSimpleID Desc"
						set rsR1=conn.execute(strR1)
						if not rsR1.eof then 
							response.write trim(rsR1("IllegalRule"))
						end if
						rsR1.close
						set rsR1=nothing
						response.write "</td>"
					End If
					'法條二
						response.write "<td>"
						response.write rsfound("Rule2")
						response.write "</td>"
					if sys_City = "苗栗縣" Or sys_City = "彰化縣" then
						response.write "<td>"
						strCarImple=""
						if left(trim(rsfound("Rule2")),4)="2110" or left(trim(rsfound("Rule2")),4)="2210" or trim(rsfound("Rule2"))="4310102" or trim(rsfound("Rule2"))="4310103" or trim(rsfound("Rule2"))="4310104" then
							if trim(rsfound("CarSimpleID"))=1 or trim(rsfound("CarSimpleID"))=2 then
								strCarImple=" and CarSimpleID in ('5','0')"
							elseif trim(rsfound("CarSimpleID"))=3 or trim(rsfound("CarSimpleID"))=4 then
								strCarImple=" and CarSimpleID in ('3','0')"
							else
								strCarImple=""
							end if
						end if
						strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rsfound("Rule2"))&"' and Version='2'"&strCarImple&" order by CarSimpleID Desc"
						set rsR1=conn.execute(strR1)
						if not rsR1.eof then 
							response.write trim(rsR1("IllegalRule"))
						end if
						rsR1.close
						set rsR1=Nothing
						response.write "</td>"
					End If
					'罰款一
						response.write "<td>"
						response.write rsfound("ForFeit1")
						response.write "</td>"
					'罰款二
						response.write "<td>"
						response.write rsfound("ForFeit2")
						response.write "</td>"												
					'填單日期
						response.write "<td>"
				if sys_City = "台南市" then 
						if trim(rsfound("BillFillDate"))<>"" and not isnull(rsfound("BillFillDate")) then
							response.write rsfound("BillFillDate")
						else
							response.write "&nbsp;"
						end if
				else
						if trim(rsfound("BillFillDate"))<>"" and not isnull(rsfound("BillFillDate")) then
							response.write gInitDT(rsfound("BillFillDate"))&"&nbsp;"
						else
							response.write "&nbsp;"
						end if
				end if
						response.write "</td>"
					
					if trim(rsfound("BillBaseTypeID"))="0" then	'攔停逕舉
						strDealLine="select b.DeallineDate,a.DCIReturnStation MemberStation from BillBaseDCIReturn a,BillBase b where a.EXCHANGETYPEID='W' and b.SN="&trim(rsfound("SN"))&" and b.CarNo=a.CarNo(+) and b.BillNo=a.BillNo(+)"
					else	'行人慢車攤販
						strDealLine="select DeallineDate,MemberStation from PasserBase where SN="&trim(rsfound("SN"))
					end if 
						set rsDealline=conn.execute(strDealLine)
						if not rsDealline.eof then
							BillDeallineDate=gInitDT(rsDealline("DeallineDate"))&"&nbsp;"
							if trim(rsfound("BillBaseTypeID"))="0" then	'攔停逕舉到案處所
								strStation="select DciStationName as StationName from Station where DciStationID='"&trim(rsDealline("MemberStation"))&"'"
							else	'行人慢車攤販到案處所
								strStation="select UnitName as StationName from UnitInfo where UnitID='"&trim(rsDealline("MemberStation"))&"'"
							end if 
							set rsStation=conn.execute(strStation)
							if not rsStation.eof then
								BillMemberStation=trim(rsStation("StationName"))
							end if
							rsStation.close
							set rsStation=nothing
						else
							BillDeallineDate="&nbsp;"
							BillMemberStation="&nbsp;"
						end if
						rsDealline.close
						set rsDealline=nothing
					'應到案日期
						response.write "<td>"
						response.write BillDeallineDate
						response.write "</td>"
					'應到案處所
						response.write "<td>"
						response.write BillMemberStation
						response.write "</td>"
					'建檔日期
						response.write "<td>"
						if trim(rsfound("RecordDate"))<>"" and not isnull(rsfound("RecordDate")) then
							response.write gInitDT(rsfound("RecordDate"))&"　"&right("00"&hour(rsfound("RecordDate")),2)&":"&right("00"&minute(rsfound("RecordDate")),2)&":"&right("00"&Second(rsfound("RecordDate")),2)
						else
							response.write "&nbsp;"
						end if
						response.write "</td>"
					'入案日期
						response.write "<td>"
						response.write DCICaseInDate&"&nbsp;"					
						response.write "</td>"
					'代保管物件
						response.write "<td>"
						strSql="select b.Content from BILLFASTENERDETAIL a,DCICode b where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BIllSN="&trim(rsfound("SN"))&" and a.CarNo='"&trim(rsfound("CarNo"))&"'"
						set rsfast=conn.execute(strsql)
						fastring=""
						while Not rsfast.eof
							if trim(fastring)<>"" then fastring=fastring&","
							fastring=fastring&rsfast("Content")
							rsfast.movenext
						wend
						rsfast.close					
						Response.write fastring
						response.write "</td>"			
					if sys_City = "台中市" Then
					'告示單號
						response.write "<td>"
						strBR="select * from BillReportNo " &_
							" where BillSn="&trim(rsfound("SN"))
						set rsBR=conn.execute(strBR)
						If Not rsBR.eof Then 
							response.write Trim(rsBR("ReportNo"))
						End If 
						rsBR.close
						set rsBR=nothing
						response.write "</td>"
					End If	
					if sys_City = "南投縣" then 
						JurgeDay_Temp="&nbsp;"
					'交通事故種類
						response.write "<td>"
						str5="select trafficaccidenttype,JurgeDay from billbase where sn="&trim(rsfound("SN"))
						Set rs5=conn.execute(str5)
						If Not rs5.eof Then
							If Trim(rsfound("billbasetypeid"))="0" And trim(rsfound("billtypeid"))="1" then
								If Trim(rs5("trafficaccidenttype"))="" Or IsNull(rs5("trafficaccidenttype")) Then
									response.write "&nbsp;"
								Else 
									response.write "A"&Trim(rs5("trafficaccidenttype"))
								End If
							End If 
							
							if Trim(rs5("JurgeDay"))<>"" and not isnull(rs5("JurgeDay")) then
								JurgeDay_Temp=gInitDT(rs5("JurgeDay"))&"&nbsp;"
							end if
						Else
							response.write "&nbsp;"
						End If
						rs5.close
						Set rs5=Nothing 
						response.write "</td>"
					'民眾檢舉日
						response.write "<td>"
						response.write JurgeDay_Temp
						response.write "</td>"
					End If
					if sys_City = "台東縣" Then
					'入案批號
						response.write "<td>"
						strBR="select * from Dcilog " &_
							" where BillSn="&trim(rsfound("SN"))&" and ExchangeTypeID='W'"
						set rsBR=conn.execute(strBR)
						If Not rsBR.eof Then 
							response.write Trim(rsBR("BatchNumber"))
						End If 
						rsBR.close
						set rsBR=nothing
						response.write "</td>"
					End If	
					if sys_City = "彰化縣" then		
						response.write "<td>"
						mdate=""
						sdate=""
						strBR="select * from billmailhistory " &_
							" where BillSn="&trim(rsfound("SN"))
						set rsBR=conn.execute(strBR)
						If Not rsBR.eof Then 
							mdate=gInitDT(Trim(rsBR("MAILDATE")))
							If Trim(rsBR("SIGNDATE"))<>"" Then
								sdate=gInitDT(Trim(rsBR("SIGNDATE")))
							ElseIf Trim(rsBR("OPENGOVDATE"))<>"" Then
								sdate=gInitDT(Trim(rsBR("OPENGOVDATE")))
							ElseIf Trim(rsBR("STOREANDSENDMailDate"))<>"" Then
								sdate=gInitDT(Trim(rsBR("STOREANDSENDMailDate")))
							End If 
							response.write mdate
						End If 
						rsBR.close
						set rsBR=nothing
						response.write "</td>"
						response.write "<td>"
							response.write sdate
						response.write "</td>"
					End If
					'操作
						response.write "<td>"
						if trim(rsfound("RecordStateID"))="-1" then
							'已刪除之告發單則顯示出刪除原因
							strRea="select a.Note,b.Content from BillDeleteReason a,DCIcode b where a.BillSN="&trim(rsfound("SN"))&" and b.TypeID=3 and b.ID=a.DelReason"
							set rsRea=conn.execute(strRea)
							if not rsRea.eof then
								response.write rsRea("Content")
								if trim(rsRea("Note"))<>"" and not isnull(rsRea("Note")) then
									response.write "("&trim(rsRea("Note"))&")"
								end if
							end if
							rsRea.close
							set rsRea=nothing
						else
							Response.Write rsfound("Note")
						end if
							response.write "</td>"
					
						response.write "</tr>"
					rsfound.MoveNext
					Wend
					rsfound.close
					set rsfound=nothing
				%>
				
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%
conn.close
set conn=nothing
%>