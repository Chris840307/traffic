<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

	strSQL="select (select min(BillSN) from DCILog"&request("strDCISQL")&" and ExchangeTypeID<>'E' and carno=b.carno) BillSN,a.CarNo,a.Owner,a.OwnerAddress,a.OwnerNotifyAddress,a.DriverHomeAddress,b.ExchangeDate,c.OwnerAddress OwnerHomeAddress,c.DriverAddress from (select CarNo,Owner,OwnerAddress,OwnerNotifyAddress,DriverHomeAddress from BillbaseDCIReturn where CarNo in(select distinct CarNo from DCILog"&request("strDCISQL")&" and ExchangeTypeID<>'E') and ExchangetypeID='A') a,(select distinct CarNo,ExchangeDate from DCILog"&request("strDCISQL")&" and ExchangeTypeID<>'E') b,(select distinct CarNo,OwnerAddress,DriverAddress from Billbase where sn in(select BillSN from DCILog"&request("strDCISQL")&" and ExchangeTypeID<>'E')) c where a.CarNo=b.CarNo and a.CarNo=c.CarNo order by a.CarNo,b.ExchangeDate"

	set rsfound=conn.execute(strSQL)
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
		<td align="center"><strong>停管車籍清冊</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="1" cellpadding="1" cellspacing="1">
				<tr>					
					<td align="center">序號</td>
					<td>車號</td>
					<td>停車時間</td>
					<td>車主姓名</td>
					<td>地址</td>
					<td>地址調整</td>
				</tr>
				<%
					filecnt=0
					while Not rsfound.eof
						AddressKind="":Ctrl=0:arr_AddressKind="":tmp_addr=""
						ONfadd="":owradd="":Drhadd="":ONfaddID="":owraddID="":DrhaddID=""

						If not ifnull(rsfound("OwnerNotifyAddress")) Then
							ONfadd=mid(trim(rsfound("OwnerNotifyAddress")),4)
							ONfaddID="OwnerNotifyAddress"

						end if

						If not ifnull(rsfound("OwnerHomeAddress")) Then
							owradd=trim(rsfound("OwnerHomeAddress"))
							owraddID="OwnerHomeAddress"

						elseIf not ifnull(rsfound("OwnerAddress")) Then
							owradd=trim(rsfound("OwnerAddress"))
							owraddID="OwnerAddress"

						End if

						If not ifnull(rsfound("DriverAddress")) Then
							Drhadd=trim(rsfound("DriverAddress"))
							DrhaddID="DriverAddress"

						elseIf not ifnull(rsfound("DriverHomeAddress")) Then
							Drhadd=trim(rsfound("DriverHomeAddress"))
							DrhaddID="DriverHomeAddress"

						End if
		
						If (ONfadd = owradd) and (ONfadd = Drhadd) and (owradd = Drhadd) Then
							If not ifnull(ONfadd) Then
								Ctrl=1
								AddressKind=ONfaddID
							End if

						elseIf (ONfadd <> owradd) and (ONfadd = Drhadd) and (owradd <> Drhadd) Then
							If (not ifnull(ONfadd)) and (not ifnull(owradd)) Then
								Ctrl=2
								AddressKind=ONfaddID & "," & owraddID
							elseIf not ifnull(ONfadd) Then
								Ctrl=1
								AddressKind=ONfaddID
							elseIf not ifnull(owradd) Then
								Ctrl=1
								AddressKind=owraddID
							End if

						elseIf (ONfadd = owradd) and (ONfadd <> Drhadd) and (owradd <> Drhadd) Then
							If (not ifnull(ONfadd)) and (not ifnull(Drhadd)) Then
								Ctrl=2
								AddressKind=ONfaddID & "," & DrhaddID
							elseIf not ifnull(ONfadd) Then
								Ctrl=1
								AddressKind=ONfaddID
							elseIf not ifnull(Drhadd) Then
								Ctrl=1
								AddressKind=DrhaddID
							End if

						elseIf (ONfadd <> owradd) and (ONfadd <> Drhadd) and (owradd = Drhadd) Then
							If (not ifnull(ONfadd)) and (not ifnull(owradd)) Then
								Ctrl=2
								AddressKind=ONfaddID & "," & owraddID
							elseIf not ifnull(ONfadd) Then
								Ctrl=1
								AddressKind=ONfaddID
							elseIf not ifnull(owradd) Then
								Ctrl=1
								AddressKind=owraddID
							End if

						elseIf (ONfadd <> owradd) and (ONfadd <> Drhadd) and (owradd <> Drhadd) Then
							If not ifnull(ONfadd) Then
								Ctrl=Ctrl+1
								AddressKind=ONfaddID
							end if

							If not ifnull(owradd) Then
								Ctrl=Ctrl+1
								If not ifnull(AddressKind) Then
									AddressKind=AddressKind & "," & owraddID
								else
									AddressKind=owraddID
								End if
							end if

							If not ifnull(Drhadd) Then
								Ctrl=Ctrl+1
								If not ifnull(AddressKind) Then
									AddressKind=AddressKind & "," & DrhaddID
								else
									AddressKind=DrhaddID
								End if
							end if

						End if
						If Ctrl > 1 Then 
							arr_AddressKind=split(AddressKind,",")

							filecnt=filecnt+1
							response.write "<tr>"
							response.write "<td rowspan="""&Ctrl&""" align=""center"">"
							Response.Write filecnt
							Response.Write "</td>"
							response.write "<td rowspan="""&Ctrl&""">"
							Response.Write trim(rsfound("CarNo"))
							Response.Write "</td>"

							response.write "<td rowspan="""&Ctrl&""">"
							strT="select * from billbase where sn in (select billsn from Dcilog where CarNo='"&trim(rsfound("CarNo"))&"' and exchangedate=to_date('"&Year(trim(rsfound("ExchangeDate")))&"/"&month(trim(rsfound("ExchangeDate")))&"/"&day(trim(rsfound("ExchangeDate")))&" "&hour(trim(rsfound("ExchangeDate")))&":"&minute(trim(rsfound("ExchangeDate")))&":"&second(trim(rsfound("ExchangeDate")))&"','YYYY/MM/DD/HH24/MI/SS'))"
							'response.write strT
							Set rsT=conn.execute(strT)
							If Not rsT.eof Then
								If Not IsNull(rsT("IllegalDate")) then
								response.write Year(trim(rsT("IllegalDate")))-1911&"/"&month(trim(rsT("IllegalDate")))&"/"&day(trim(rsT("IllegalDate")))&" "&hour(trim(rsT("IllegalDate")))&":"&minute(trim(rsT("IllegalDate")))
								End if
							End If
							rsT.close
							Set rsT=Nothing 
							Response.Write "</td>"

							response.write "<td rowspan="""&Ctrl&""">"
							Response.Write funcCheckFont(trim(rsfound("Owner")),25,1)
							Response.Write "</td>"

							response.write "<td rowspan="""&Ctrl&""">"
							Response.Write "<a href=""AddressUpdate.asp?sys_CarNo="&trim(rsfound("CarNo"))&"&FileName=&BillSN="&trim(rsfound("BillSN"))&""" target=""_blank"">地址調整</a>"
							Response.Write "</td>"

							For i = 1 to Ctrl
								If i > 1 Then response.write "<tr>"
								response.write "<td>"

								If i = 1 Then Response.Write "<br>"

								If trim(arr_AddressKind(i-1)) = "OwnerAddress" Then
									Response.Write "(車)"
									Response.Write funcCheckFont(rsfound("OwnerAddress"),25,1)

								elseif trim(arr_AddressKind(i-1)) = "OwnerHomeAddress" Then
									Response.Write "(車)"
									Response.Write funcCheckFont(rsfound("OwnerHomeAddress"),25,1)

								elseIf trim(arr_AddressKind(i-1)) = "OwnerNotifyAddress" Then
									Response.Write "(通)"
									Response.Write funcCheckFont(rsfound("OwnerNotifyAddress"),25,1)

								elseIf trim(arr_AddressKind(i-1)) = "DriverHomeAddress" Then
									Response.Write "(戶)"
									Response.Write funcCheckFont(rsfound("DriverHomeAddress"),25,1)

								elseIf trim(arr_AddressKind(i-1)) = "DriverAddress" Then
									Response.Write "(戶)"
									Response.Write funcCheckFont(rsfound("DriverAddress"),25,1)

								end if

								Response.Write "</td>"
								response.write "</tr>"
							Next
						End if 
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