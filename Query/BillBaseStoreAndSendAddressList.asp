<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
	
	fMnoth=month(now)
	if fMnoth<10 then fMnoth="0"&fMnoth
	fDay=day(now)
	if fDay<10 then	fDay="0"&fDay
	fname=year(now)&fMnoth&fDay&"戶籍地址補正清冊.xls"
	Response.AddHeader "Content-Disposition", "filename="&fname
	response.contenttype="application/x-msexcel; charset=MS950"
	
	strCity="select value from Apconfigure where id=30"
	set rsCity=conn.execute(strCity)
	titlePage=trim(rsCity("value"))
	rsCity.close

	if UCase(request("Sys_BatchNumber"))<>"" then
		tmp_BatchNumber=split(UCase(request("Sys_BatchNumber")),",")
		for i=0 to Ubound(tmp_BatchNumber)
			if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
			if i=0 then
				Sys_BatchNumber=trim(Sys_BatchNumber)&tmp_BatchNumber(i)
			else
				Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&tmp_BatchNumber(i)
			end if
			if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(Sys_BatchNumber)&"'"
		next
		'strwhere=" and a.BatchNumber in('"&Sys_BatchNumber&"')"
	end if

	if request("Sys_BatchNumber")<>"" then

		strwhere=" BillSn in (select BillSn from DciLog where batchnumber in('"&Sys_BatchNumber&"') and exchangetypeid='N')"
	end If 

	strSQL="select BillNo,CarNo,MailNumber,MailDate,(select Content from DCICode where TypeID=7 and ID=BillMailHistory.UserMarkResonID) UserMarkResonName" & _
	",(select BillTypeID from BillBase where Sn=BillMailHistory.BillSN) BillTypeID" & _
	",(select DciReturnStation from BillBaseDciReturn where ExchangeTypeID='W' and BillNo=BillMailHistory.BillNo and CarNo=BillMailHistory.CarNo) DciReturnStation" & _
	",(select Owner from BillBaseDciReturn where ExchangeTypeID='W' and BillNo=BillMailHistory.BillNo and CarNo=BillMailHistory.CarNo) Owner" & _
	",(select Driver from BillBaseDciReturn where ExchangeTypeID='W' and BillNo=BillMailHistory.BillNo and CarNo=BillMailHistory.CarNo) Driver" & _
	",(select DriverID from BillBaseDciReturn where ExchangeTypeID='W' and BillNo=BillMailHistory.BillNo and CarNo=BillMailHistory.CarNo) DriverID" & _
	",(select OwnerZIP||'@'||OwnerAddress from BillBase where Sn=BillMailHistory.BillSN) OwnerAddress" & _
	",(select DriverZip||'@'||DriverAddress from BillBase where Sn=BillMailHistory.BillSN) DriverAddress" & _
	",(select (select Content from dcicode where TypeID=5 and ID=BillBaseDciReturn.DCIReturnCarType) content from BillBaseDciReturn where ExchangeTypeID='W' and BillNo=BillMailHistory.BillNo and CarNo=BillMailHistory.CarNo) CarType" & _
	" from BillMailHistory where " & strwhere &" order by DciReturnStation"

	set rsfound=conn.execute(strSQL)
	

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>戶籍地址補正清冊</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" border="0">
	<tr>
		<td align="center"><strong><%=titlePage%>違規案件委外作業-二次寄存戶籍地址補正確認清冊</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" border="2" cellpadding="0" cellspacing="0">
				<tr>					
					<td align="center">序號</td>
					<td align="center">監理站</td>
					<td align="center">單號</td>
					<td align="center">車號</td>
					<td align="center">車種</td>
					<td align="center">車主姓名</td>
					<td align="center">違規人證號</td>
					<td align="center">原掛號碼</td>
					<td align="center">原郵寄日</td>
					<td align="center">退件原因</td>
					<td align="center">原投遞地址</td>
					<td align="center">戶籍地址</td>
				</tr>
				<%
					filecnt=0
					while Not rsfound.eof


						Owner=trim(rsfound("Owner")):DriverID=trim(rsfound("DriverID"))
						If trim(rsfound("Driver")) <> "" and trim(rsfound("DriverID"))<>"" Then
							Owner=trim(rsfound("Driver"))
						End If 
						
						
					
						StationName=""
						strSQL="select DciStationName from Station where DciStationID='" & trim(rsfound("DciReturnStation")) & "'"
						set rs=conn.execute(strSQL)
						If not rs.eof Then StationName=trim(rs("DciStationName"))
						rs.close

						Sys_OwnerAddress="":tmpOwnerAddress="":Sys_OwnerZipName=""
						Sys_DriverAddress="":tmpDriverAddress="":Sys_DriverZipName=""

						If not ifnull(replace(rsfound("OwnerAddress"),"@","")) Then
							if trim(rsfound("BillNo")) <>"F13367617" then 

								tmpOwnerAddress=split(rsfound("OwnerAddress"),"@")
								Sys_OwnerZipName=""
								If not ifnull(trim(tmpOwnerAddress(0))) Then

									strSQL="select zipName from zip where zipid="&trim(tmpOwnerAddress(0))
									set rs=conn.execute(strSQL)
									If not rs.eof Then Sys_OwnerZipName=trim(rs("zipName"))
									rs.close
								end If 

								Sys_OwnerAddress=trim(tmpOwnerAddress(0))&" "&Sys_OwnerZipName&replace(tmpOwnerAddress(1),Sys_OwnerZipName,"")
							end if
						End if 
 
				

						If not ifnull(replace(rsfound("DriverAddress"),"@","")) Then
							if trim(rsfound("BillNo")) <>"F13367617" then 
								Sys_DriverAddress=""

								tmpDriverAddress=split(rsfound("DriverAddress"),"@")

								If not ifnull(trim(tmpDriverAddress(0))) Then
									strSQL="select zipName from zip where zipid="&trim(tmpDriverAddress(0))
									set rs=conn.execute(strSQL)
									If not rs.eof Then Sys_DriverZipName=trim(rs("zipName"))
									rs.close
								End if 

								Sys_DriverAddress=trim(tmpDriverAddress(0))&" "&Sys_DriverZipName&replace(tmpDriverAddress(1),Sys_DriverZipName,"")
							end if
						End if 
	
					
						filecnt=filecnt+1
						response.write "<tr>"
						response.write "<td align=""center"">"
						Response.Write filecnt
						Response.Write "</td>"

						response.write "<td>"
						Response.Write StationName
						Response.Write "</td>"

						response.write "<td>"
						Response.Write trim(rsfound("BillNo"))
						Response.Write "</td>"

						response.write "<td>"
						Response.Write trim(rsfound("CarNo"))
						Response.Write "</td>"

						response.write "<td>"
						Response.Write trim(rsfound("CarType"))
						Response.Write "</td>"

						response.write "<td>"
						Response.Write trim(Owner)
						Response.Write "</td>"

						response.write "<td>"
						Response.Write trim(DriverID)
						Response.Write "</td>"

						response.write "<td>"
						Response.Write trim(rsfound("MailNumber"))
						Response.Write "</td>"

						response.write "<td>"
						Response.Write gInitDT(trim(rsfound("MailDate")))
						Response.Write "</td>"

						response.write "<td>"
						Response.Write trim(rsfound("UserMarkResonName"))
						Response.Write "</td>"

						response.write "<td>"
						If trim(rsfound("BillTypeID")) = 1 Then
							Response.Write Sys_DriverAddress

						else
							Response.Write Sys_OwnerAddress

						End if 						
						Response.Write "</td>"

						response.write "<td>"
						Response.Write Sys_DriverAddress
						Response.Write "</td>"

						response.write "</tr>"
						rsfound.movenext

					wend
				rsfound.close
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td align="center">
			前揭違規案件業已建檔傳送資料庫，請查核無誤後，於603表蓋章，本件留存備查。
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close%>