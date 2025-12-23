<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%

	strSQL="select a.CarNo,a.Owner,a.OwnerAddress,a.DriverHomeAddress,c.OwnerNotifyAddress,b.BillNo,b.ExchangeDate from (select CarNo,Owner,OwnerAddress,DriverHomeAddress from BillbaseDCIReturn where CarNo in(select distinct CarNo from DCILog"&request("strDCISQL")&" and ExchangeTypeID<>'E') and ExchangetypeID='A') a,(select BillNo,CarNo,ExchangeDate from DCILog"&request("strDCISQL")&" and ExchangeTypeID<>'E') b,(select CarNo,OwnerAddress OwnerNotifyAddress from BillbaseDCIReturn where BillNo in(select BillNo from DCILog"&request("strDCISQL")&" and ExchangeTypeID<>'E') and ExchangetypeID='W') c where a.CarNo=b.CarNo and a.CarNo=c.CarNo order by a.CarNo,b.ExchangeDate"

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
					<td>單號</td>
					<td>車號</td>
					<td>車主姓名</td>
					<td>地址</td>
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


						If not ifnull(rsfound("DriverHomeAddress")) Then
							Drhadd=trim(rsfound("DriverHomeAddress"))
							DrhaddID="DriverHomeAddress"

						End if
		
						If ONfadd = Drhadd Then
							If not ifnull(ONfadd) Then
								Ctrl=1
								AddressKind=ONfaddID
							End if


						elseIf ONfadd <> Drhadd Then
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

						End if

						arr_AddressKind=split(AddressKind,",")

						filecnt=filecnt+1
						response.write "<tr>"
						response.write "<td rowspan="""&Ctrl&""" align=""center"">"
						Response.Write filecnt
						Response.Write "</td>"
						response.write "<td rowspan="""&Ctrl&""">"
						Response.Write trim(rsfound("BillNo"))
						Response.Write "</td>"
						response.write "<td rowspan="""&Ctrl&""">"
						Response.Write trim(rsfound("CarNo"))
						Response.Write "</td>"

						response.write "<td rowspan="""&Ctrl&""">"
						Response.Write trim(rsfound("Owner"))
						Response.Write "</td>"

						For i = 1 to Ctrl
							If i > 1 Then response.write "<tr>"
							response.write "<td>"

							If trim(arr_AddressKind(i-1)) = "OwnerNotifyAddress" Then
								Response.Write "(通)"
								Response.Write rsfound("OwnerNotifyAddress")

							elseIf trim(arr_AddressKind(i-1)) = "DriverHomeAddress" Then
								Response.Write "(戶)"
								Response.Write rsfound("DriverHomeAddress")
							end if

							Response.Write "</td>"
							response.write "</tr>"
						Next
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