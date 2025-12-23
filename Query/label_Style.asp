<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舉發單信封黏貼標籤</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
</head>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<body>
<%
Server.ScriptTimeout = 12000
strCity    = "select value from Apconfigure where id=31"
set rsCity = conn.execute(strCity)
sys_City   = trim(rsCity("value"))
rsCity.close
set rsCity = nothing
'sys_City="彰化縣"
'--------------------------------------------------------------------------------------------------------------------
'管轄郵遞區號
strCode		= "select value from apconfigure where name='管轄郵遞區號'"
set rsCode	= conn.execute(strCode)
if not rsCode.eof then
	Code	= trim(rsCode("value"))
end If
rsCode.close
set rsCode	= nothing	

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
'------------------------------------------------------------------------------------------------
If sys_City="彰化縣" Then 

	tempSQL = " where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) " & _
			  " and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) " & _
			  " and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' " & _
			  " and a.DciErrorCarData Not in ('1','3','9','a','j','A','H','K','L','T','V','n') " & _
			  " and a.DciReturnStatusID<>'n' " & request("sys_strSQL") & ") or ((a.BillTypeID='1' or f.UseTool=8) and a.ExchangeTypeID=d.DCIActionID(+) " & _
			  " and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) " & _
			  " and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 " & _
			  " and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and f.RecordStateId <> -1 " & request("sys_strSQL") & ")"

	strBil  = "select a.BillSN,a.RecordMemberID,f.RecordDate from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d," & _
			  " (select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g," & _
			  " (select * from DciReturnStatus where DciActionID='WE') h " & tempSQL & " order by a.RecordMemberID,f.RecordDate"

	set rsbil 	= conn.execute(strBil)
	PBillSN 	= ""
	while Not rsbil.eof
		if trim(PBillSN) <> "" then PBillSN = trim(PBillSN) & ","
		PBillSN = PBillSN & rsbil("BillSN")
		rsbil.movenext
	wend
	Set rsbil = Nothing

	if (Instr(request("Sys_BatchNumber"),"N") > 0) and trim(PBillSN) <> "" then
		strSQL 		= "Select BillSN from BillMailHistory where BillSN in(" & PBillSN & ") order by UserMarkDate"
		set rshis 	= conn.execute(strSQL)
		PBillSN 	= ""
		while Not rshis.eof
			if trim(PBillSN) <> "" then PBillSN = trim(PBillSN) & ","
			PBillSN = PBillSN & rshis("BillSN")
			rshis.movenext
		wend
		rshis.close
	End if
	
	PBillSN = Split(trim(PBillSN),",")
Else
	PBillSN = Split(trim(request("PBillSN")),",")
End if

for i=0 to Ubound(PBillSN)

if cint(i)>0 and i mod 5=0 then response.write "<div class=""PageNext"">&nbsp;</div>"

if i mod 5=0 then 

'---------------------------------------------------------------------------------------
strBill	= "select b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress,c.StationName from billbasedcireturn a" & _
		  ",Billbase b,Station c where a.DciReturnStation=c.DCIStationID(+) and a.BillNO=b.BillNo and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN=" & PBillSN(i)

set rsBill = conn.execute(strBill)
GetMailAddress		= ""
Billno				= ""
StationName			= ""
CarNo				= ""
Owner				= ""
Sys_BillTypeID		= ""
OwnerZip			= ""
Sys_DriverHomeAddress=""
Sys_DriverHomeZip	= ""
Sys_Driver			= ""
Sys_BillNo_BarCode	= ""
ZipName2			= ""

if not rsBill.eof then
	if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" then
		ZipName=""
	else
		strZip		= "select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
		set rsZip	= conn.execute(strZip)
		if not rsZip.eof then
			ZipName	= trim(rsZip("ZipName"))
		end if
		rsZip.close
		set rsZip	= nothing
		
		strZip		= "select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
		set rsZip	= conn.execute(strZip)
		if not rsZip.eof then
			ZipName2= trim(rsZip("ZipName"))
		end if
		rsZip.close
		set rsZip	= nothing
	end if
	
	GetMailAddress	= ZipName&trim(rsBill("OwnerAddress"))
			
	Billno			= trim(rsBill("Billno"))
	StationName		= trim(rsBill("StationName"))
	CarNo			= trim(rsBill("CarNo"))
	Owner			= trim(rsBill("Owner"))
	Sys_BillTypeID	= trim(rsBill("BillTypeID"))
	OwnerZip		= trim(rsBill("OwnerZip"))
			
	Sys_DriverHomeAddress	= ZipName2&trim(rsBill("DriverHomeAddress"))
    Sys_DriverHomeZip		= trim(rsBill("DriverHomeZip"))
	Sys_Driver				= trim(rsBill("Driver"))
	Sys_DriverHomeAddress	= replace(Sys_DriverHomeAddress,ZipName2&ZipName2,ZipName2)
	Sys_BillNo_BarCode		= BillNo
	GetMailAddress			= replace(GetMailAddress,ZipName&ZipName,ZipName)

	If sys_City = "彰化縣" or sys_City = "高雄市" or sys_City = "屏東縣" or sys_City="嘉義市" Then
		DelphiASPObj.GenSendStoreBillno BillNo,0,50,160,1
	else
		DelphiASPObj.GenSendStoreBillno BillNo,0,50,160
	end if
            	
	If sys_City = "台東縣" Then 
		If Driver = "" Or Sys_DriverHomeAddress = "" Then 
			strBill		= "select a.Owner,a.OwnerZip,a.OwnerAddress,a.DriverHomeZip,a.DriverHomeAddress,a.Driver from billbasedcireturn a where a.ExchangeTypeID='A' and a.Carno='"&CarNo&"'"
			set rsBill	= conn.execute(strBill)
			
			if not rsBill.eof Then
				ZipName	= ""
				strZip	= "select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip	= conn.execute(strZip)
				if not rsZip.eof then
					ZipName = trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=Nothing

				GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
				GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)

				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
					if not rsZip.eof then
						ZipName2=trim(rsZip("ZipName"))
					end if
				rsZip.close
				set rsZip=nothing


				Owner				  = trim(rsBill("Owner"))
				OwnerZip			  = trim(rsBill("OwnerZip"))
				Sys_DriverHomeZip	  = trim(rsBill("DriverHomeZip"))
				Sys_DriverHomeAddress = ZipName2&trim(rsBill("DriverHomeAddress"))
				Sys_DriverHomeAddress = replace(Sys_DriverHomeAddress,ZipName2&ZipName2,ZipName2)

				Sys_Driver=trim(rsBill("Driver"))

			End if
		End If 
	End if
end If
rsBill.close
set rsBill=nothing	
'-------------------------------------------------------------------------------------
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_Owner=""
Sys_OwnerZip=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_BillNo=""
Sys_CarNo=""
Sys_Rule1=""
Sys_Rule2=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)

Sys_Level1=0:Sys_Level2=0

if Not rsfound.eof then Sys_Level1=trim(rsfound("FORFEIT1"))
if Not rsfound.eof then Sys_Level2=trim(rsfound("FORFEIT2"))
Sum_Level=cdbl(funTnumber(Sys_Level1))+cdbl(funTnumber(Sys_Level2))

If Sys_BillTypeID="1" Then
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Instr(request("Sys_BatchNumber"),"N")>0 then
	Sys_OwnerAddress=""
	Sys_OwnerZip=""
else
	If Sys_BillTypeID="1" Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_OwnerZip	=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip	=trim(rsfound("OwnerZip"))
	End if
end if

If ifnull(Sys_OwnerAddress) Then
If Sys_BillTypeID="1" Then
	strSql="select * from BillbaseDCIReturn where  BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
else
	strSql="select * from BillbaseDCIReturn where CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A'"
End if
	set rsdata=conn.execute(strsql)
	If Sys_BillTypeID="1" Then
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))
	else
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
	End if

	if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
		if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))

		If ifnull(Sys_OwnerAddress) Then
			if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
			if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
		end if
	else
		If Sys_BillTypeID="1" Then
			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("DriverHomeZip"))
		else
			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
		End if
	end if
	rsdata.close
end if

If ifnull(Sys_OwnerAddress) Then
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=replace(trim(rszip("ZipName")),"台","臺")
if Not rszip.eof then Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(rszip("ZipName")),"")
if Not rszip.eof then Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(Sys_OwnerZipName),"")
rszip.close

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")

If sys_City="彰化縣" or sys_City="高雄市" or sys_City = "屏東縣" or sys_City="嘉義市" Then
	DelphiASPObj.GenSendStoreBillno BillNo,0,60,160,1
else
	DelphiASPObj.GenSendStoreBillno BillNo,0,60,160
end if

%>
	
	<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%
				'宜蘭停管單退要抓戶籍
				if (Sys_BillTypeID="1") and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_Driver)="" Then 				
					response.write funcCheckFont(Owner,18,1)
				  Else
					response.write funcCheckFont(Sys_Driver,18,1) 
				  End if
				else 
					response.write funcCheckFont(Owner,18,1)
				end if
				%>
				<%if sys_City="彰化縣" then
				response.write "　先生/女士"
				end if
				%>
				</font></td>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="120" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>""  then
				  If Trim(Sys_DriverHomeZip)="" Then 
					response.write OwnerZip 
				  Else
					response.write Sys_DriverHomeZip 
				  End if
				else 
					response.write Sys_OwnerZip  
				end if
				%>&nbsp;<%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>""  then 
				  If Trim(Sys_DriverHomeAddress)="" Then 
					response.write funcCheckFont(GetMailAddress,18,1)
				  Else
					response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
				  End if
				else 
					If Trim(Sys_DriverHomeAddress)="" Then 
						response.write replace(Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) ,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
					else
						response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
					end if
					
				end If 
				
				'if sys_City="彰化縣" and Session("UnitLevelID")>"1" then
				'	response.write "(罰鍰："&Sum_Level&"元)"
				'end if 
					%></font>
         　</td>
		</tr>
	</table>
	<% 

	'---------------------------------------------------------------------------------------
	if (i+1 < Ubound(PBillSN)) or (i+1 = Ubound(PBillSN))then 
		strBill="select b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress," &_
				"c.StationName from billbasedcireturn a,Billbase b,Station c where a.DciReturnStation=c.DCIStationID(+) and a.BillNO=b.BillNo " &_
				" and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN=" & PBillSN(i+1)

		set rsBill=conn.execute(strBill)
			GetMailAddress		 =""
			Billno				 =""
			StationName			 =""
			CarNo				 =""
			Owner				 =""
			Sys_BillTypeID		 =""
			OwnerZip			 =""
		    Sys_DriverHomeAddress=""
            Sys_DriverHomeZip	 =""
    	    Sys_Driver			 =""
			Sys_BillNo_BarCode	 =""
			ZipName2			 =""
			
		if not rsBill.eof then
		
	    	if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" then
				ZipName=""
			else
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
			end if
			
			GetMailAddress		  = ZipName&trim(rsBill("OwnerAddress"))
			
			Billno				  = trim(rsBill("Billno"))
			StationName			  = trim(rsBill("StationName"))
			CarNo				  = trim(rsBill("CarNo"))
			Owner				  = trim(rsBill("Owner"))
			Sys_BillTypeID		  = trim(rsBill("BillTypeID"))
			OwnerZip			  = trim(rsBill("OwnerZip"))
			
			Sys_DriverHomeAddress = ZipName2&trim(rsBill("DriverHomeAddress"))
            Sys_DriverHomeZip	  = trim(rsBill("DriverHomeZip"))
			Sys_Driver			  = trim(rsBill("Driver"))

			Sys_BillNo_BarCode	  = BillNo
			Sys_DriverHomeAddress = replace(Sys_DriverHomeAddress,ZipName2&ZipName2,ZipName2)
			GetMailAddress		  = replace(GetMailAddress,ZipName&ZipName,ZipName)
            If sys_City="彰化縣" or sys_City="高雄市" or sys_City = "屏東縣" or sys_City="嘉義市" Then
				DelphiASPObj.GenSendStoreBillno BillNo,0,50,160,1
			else
				DelphiASPObj.GenSendStoreBillno BillNo,0,50,160
			end if
			If sys_City="台東縣" Then 
				If Driver="" Or Sys_DriverHomeAddress="" Then 
				strBill="select a.Owner,a.OwnerZip,a.OwnerAddress,a.DriverHomeZip,a.DriverHomeAddress,a.Driver from billbasedcireturn a where a.ExchangeTypeID='A' and a.Carno='"&CarNo&"'"

				set rsBill=conn.execute(strBill)
					if not rsBill.eof Then
						ZipName=""
							strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=Nothing

							GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
							GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)

							strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName2=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing


						Owner				  = trim(rsBill("Owner"))
						OwnerZip			  = trim(rsBill("OwnerZip"))
						Sys_DriverHomeZip	  = trim(rsBill("DriverHomeZip"))
						Sys_DriverHomeAddress = ZipName2&trim(rsBill("DriverHomeAddress"))
						Sys_DriverHomeAddress = replace(Sys_DriverHomeAddress,ZipName2&ZipName2,ZipName2)

						Sys_Driver=trim(rsBill("Driver"))

					End if
				End If 
			End if
		end If
	rsBill.close
	set rsBill=nothing	
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+1)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_Owner=""
Sys_OwnerZip=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_BillNo=""
Sys_CarNo=""
Sys_Rule1=""
Sys_Rule2=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)

If Sys_BillTypeID="1" Then
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	If ifnull(Sys_OwnerAddress) Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
else
	If Sys_BillTypeID="1" Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	End if
end if

If ifnull(Sys_OwnerAddress) Then
	strSql="select * from BillbaseDCIReturn where CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A'"
	set rsdata=conn.execute(strsql)
	If Sys_BillTypeID="1" Then
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))
	else
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
	End if

	if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
	rsdata.close
end if

If ifnull(Sys_OwnerAddress) Then
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=replace(trim(rszip("ZipName")),"台","臺")
if Not rszip.eof then Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(rszip("ZipName")),"")
if Not rszip.eof then Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(Sys_OwnerZipName),"")
rszip.close

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")

If sys_City="彰化縣" or sys_City="高雄市" or sys_City = "屏東縣" or sys_City="嘉義市" Then
	DelphiASPObj.GenSendStoreBillno BillNo,0,60,160,1

else
	DelphiASPObj.GenSendStoreBillno BillNo,0,60,160

end if

'-------------------------------------------------------------------------------------
	%>
		<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1") and trim(Sys_DriverHomeAddress)<>"" Then
				  If Trim(Sys_Driver)="" Then 				
					response.write funcCheckFont(Owner,18,1) 
				  Else
					response.write funcCheckFont(Sys_Driver,18,1)  
				  End if
				else
					response.write funcCheckFont(Owner,18,1) 				
				end if
				%>				
				<%if sys_City="彰化縣" then
				response.write "　先生/女士"
				end if
				%>
				</font></td>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="120" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_DriverHomeZip)="" Then 
					response.write OwnerZip 
				  Else
					response.write Sys_DriverHomeZip 
				  End if
				else 
					response.write Sys_OwnerZip
				end if
				%>&nbsp;<%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_DriverHomeAddress)="" Then 
					response.write funcCheckFont(GetMailAddress,18,1) 
				  Else
					response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
				  End if
				else
					If Trim(Sys_DriverHomeAddress)="" Then 
						response.write replace(Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) ,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
					else
						response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
					end if					
				end If 
				
				'if sys_City="彰化縣" and Session("UnitLevelID")>"1" then
				'	response.write "(罰鍰："&Sum_Level&"元)"
				'end if 
				%></font>
         　</td>
		</tr>
	</table>
	<% 
end if	
	'---------------------------------------------------------------------------------------
	if (i+2 < Ubound(PBillSN)) or (i+2 = Ubound(PBillSN))then 
strBill="select b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress,c.StationName from billbasedcireturn a,Billbase b,Station c where a.DciReturnStation=c.DCIStationID(+) and a.BillNO=b.BillNo and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i+2)
set rsBill=conn.execute(strBill)
			GetMailAddress=""
			Billno=""
			StationName=""
			CarNo=""
			Owner=""
			Sys_BillTypeID=""
			OwnerZip=""
			 Sys_DriverHomeAddress=""
             Sys_DriverHomeZip=""
    	     Sys_Driver=""
			Sys_BillNo_BarCode=""
			ZipName2=""
		if not rsBill.eof then
		if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" then
				ZipName=""
			else
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
			end if
				GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
			
			Billno=trim(rsBill("Billno"))
			StationName=trim(rsBill("StationName"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))
			Sys_BillTypeID=trim(rsBill("BillTypeID"))
			OwnerZip=trim(rsBill("OwnerZip"))
			
			 Sys_DriverHomeAddress=ZipName2&trim(rsBill("DriverHomeAddress"))
                 Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
			     Sys_Driver=trim(rsBill("Driver"))

			Sys_BillNo_BarCode=BillNo
										Sys_DriverHomeAddress=replace(Sys_DriverHomeAddress,ZipName2&ZipName2,ZipName2)
										GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)
            	If sys_City="彰化縣" or sys_City="高雄市" or sys_City = "屏東縣" or sys_City="嘉義市" Then
				DelphiASPObj.GenSendStoreBillno BillNo,0,50,160,1

			else
				DelphiASPObj.GenSendStoreBillno BillNo,0,50,160

			end If 

			If sys_City="台東縣" Then 
				If Driver="" Or Sys_DriverHomeAddress="" Then 
				strBill="select a.Owner,a.OwnerZip,a.OwnerAddress,a.DriverHomeZip,a.DriverHomeAddress,a.Driver from billbasedcireturn a where a.ExchangeTypeID='A' and a.Carno='"&CarNo&"'"

				set rsBill=conn.execute(strBill)
					if not rsBill.eof Then
						ZipName=""
							strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=Nothing

							GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
							GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)

							strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName2=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing


						Owner=trim(rsBill("Owner"))
						OwnerZip=trim(rsBill("OwnerZip"))
						 Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
						 Sys_DriverHomeAddress=ZipName2&trim(rsBill("DriverHomeAddress"))
							Sys_DriverHomeAddress=replace(Sys_DriverHomeAddress,ZipName2&ZipName2,ZipName2)

						 Sys_Driver=trim(rsBill("Driver"))

					End if
				End If 
			End if
		end If
	rsBill.close
	set rsBill=nothing	
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+2)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_Owner=""
Sys_OwnerZip=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_BillNo=""
Sys_CarNo=""
Sys_Rule1=""
Sys_Rule2=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)

If Sys_BillTypeID="1" Then
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	If ifnull(Sys_OwnerAddress) Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
else
	If Sys_BillTypeID="1" Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	End if
end if

If ifnull(Sys_OwnerAddress) Then
	strSql="select * from BillbaseDCIReturn where CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A'"
	set rsdata=conn.execute(strsql)
	If Sys_BillTypeID="1" Then
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))
	else
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
	End if

	if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
	rsdata.close
end if

If ifnull(Sys_OwnerAddress) Then
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=replace(trim(rszip("ZipName")),"台","臺")
if Not rszip.eof then Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(rszip("ZipName")),"")
if Not rszip.eof then Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(Sys_OwnerZipName),"")
rszip.close

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")

If sys_City="彰化縣" or sys_City="高雄市" or sys_City = "屏東縣" or sys_City="嘉義市" Then
	DelphiASPObj.GenSendStoreBillno BillNo,0,60,160,1

else
	DelphiASPObj.GenSendStoreBillno BillNo,0,60,160

end if
'-------------------------------------------------------------------------------------
	%>
	
		<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1")  and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_Driver)="" Then 				
					response.write funcCheckFont(Owner,18,1) 
				  Else
					response.write funcCheckFont(Sys_Driver,18,1)  
				  End if
				else 
					response.write funcCheckFont(Owner,18,1)
				end if
				%>				
				<%if sys_City="彰化縣" then
				response.write "　先生/女士"
				end if
				%>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="120" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_DriverHomeZip)="" Then 
					response.write OwnerZip 
				  Else
					response.write Sys_DriverHomeZip 
				  End if
				else 
					response.write Sys_OwnerZip 
				end if
				%>&nbsp;<%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕"))  and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_DriverHomeAddress)="" Then 
					response.write funcCheckFont(GetMailAddress,18,1)
				  Else
					response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
				  End if
				else 
					If Trim(Sys_DriverHomeAddress)="" Then 
						response.write replace(Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) ,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
					else
						response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
					end if
				end If 

				'if sys_City="彰化縣" and Session("UnitLevelID")>"1" then
				'	response.write "(罰鍰："&Sum_Level&"元)"
				'end if 
				%></font>
         　</td>
		</tr>
	</table>
	<% 
	'---------------------------------------------------------------------------------------
	end if
	if (i+3 < Ubound(PBillSN)) or (i+3 = Ubound(PBillSN))then 
strBill="select b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress,c.StationName from billbasedcireturn a,Billbase b,Station c where a.DciReturnStation=c.DCIStationID(+) and a.BillNO=b.BillNo and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i+3)
set rsBill=conn.execute(strBill)
			GetMailAddress=""
			Billno=""
			StationName=""
			CarNo=""
			Owner=""
			Sys_BillTypeID=""
			OwnerZip=""
			 Sys_DriverHomeAddress=""
             Sys_DriverHomeZip=""
    	     Sys_Driver=""
			Sys_BillNo_BarCode=""
			ZipName2=""
		if not rsBill.eof then
		
	    	if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣" then
				ZipName=""
			else
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
			end if
				GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
			
			Billno=trim(rsBill("Billno"))
			StationName=trim(rsBill("StationName"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))
			Sys_BillTypeID=trim(rsBill("BillTypeID"))
			OwnerZip=trim(rsBill("OwnerZip"))
			
			Sys_DriverHomeAddress=ZipName2&trim(rsBill("DriverHomeAddress"))
            Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
			Sys_Driver=trim(rsBill("Driver"))

			Sys_BillNo_BarCode=BillNo
										Sys_DriverHomeAddress=replace(Sys_DriverHomeAddress,ZipName2&ZipName2,ZipName2)
										GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)
            	If sys_City="彰化縣" or sys_City="高雄市" or sys_City = "屏東縣" or sys_City="嘉義市" Then
				DelphiASPObj.GenSendStoreBillno BillNo,0,50,160,1

			else
				DelphiASPObj.GenSendStoreBillno BillNo,0,50,160

			end If 

			If sys_City="台東縣" Then 
				If Driver="" Or Sys_DriverHomeAddress="" Then 
				strBill="select a.Owner,a.OwnerZip,a.OwnerAddress,a.DriverHomeZip,a.DriverHomeAddress,a.Driver from billbasedcireturn a where a.ExchangeTypeID='A' and a.Carno='"&CarNo&"'"

				set rsBill=conn.execute(strBill)
					if not rsBill.eof Then
						ZipName=""
							strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=Nothing

							GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
							GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)

							strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName2=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing


						Owner=trim(rsBill("Owner"))
						OwnerZip=trim(rsBill("OwnerZip"))
						 Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
						 Sys_DriverHomeAddress=ZipName2&trim(rsBill("DriverHomeAddress"))
							Sys_DriverHomeAddress=replace(Sys_DriverHomeAddress,ZipName2&ZipName2,ZipName2)

						 Sys_Driver=trim(rsBill("Driver"))

					End if
				End If 
			End if
		end If
	rsBill.close
	set rsBill=nothing	
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+3)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_Owner=""
Sys_OwnerZip=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_BillNo=""
Sys_CarNo=""
Sys_Rule1=""
Sys_Rule2=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)

If Sys_BillTypeID="1" Then
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	If ifnull(Sys_OwnerAddress) Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
else
	If Sys_BillTypeID="1" Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	End if
end if

If ifnull(Sys_OwnerAddress) Then
	strSql="select * from BillbaseDCIReturn where CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A'"
	set rsdata=conn.execute(strsql)
	If Sys_BillTypeID="1" Then
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))
	else
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
	End if

	if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
	rsdata.close
end if

If ifnull(Sys_OwnerAddress) Then
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=replace(trim(rszip("ZipName")),"台","臺")
if Not rszip.eof then Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(rszip("ZipName")),"")
if Not rszip.eof then Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(Sys_OwnerZipName),"")
rszip.close

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")

If sys_City="彰化縣" or sys_City="高雄市" or sys_City = "屏東縣" or sys_City="嘉義市" Then
	DelphiASPObj.GenSendStoreBillno BillNo,0,60,160,1

else
	DelphiASPObj.GenSendStoreBillno BillNo,0,60,160

end if
'-------------------------------------------------------------------------------------
	%>
	
		<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1") and trim(Sys_DriverHomeAddress)<>"" then
				  If Trim(Sys_Driver)="" Then 				
					response.write funcCheckFont(Owner,18,1) 
				  Else
					response.write funcCheckFont(Sys_Driver,18,1)  
				  End if
				else
					response.write funcCheckFont(Owner,18,1) 
				end if
				%>				
				<%if sys_City="彰化縣" then
				response.write "　先生/女士"
				end if
				%>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="120" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>"" then
				  If Trim(Sys_DriverHomeZip)="" Then 
					response.write OwnerZip 
				  Else
					response.write Sys_DriverHomeZip 
				  End if

				else 
					response.write Sys_OwnerZip  
				end if
				%>&nbsp;<%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕")) and trim(Sys_DriverHomeAddress)<>"" then 
				  If Trim(Sys_DriverHomeAddress)="" Then 
					response.write funcCheckFont(GetMailAddress,18,1)
				  Else
					response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
				  End if
				else 
					If Trim(Sys_DriverHomeAddress)="" Then 
						response.write replace(Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) ,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
					else
						response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
					end if
				end If 

				'if sys_City="彰化縣" and Session("UnitLevelID")>"1" then
				'	response.write "(罰鍰："&Sum_Level&"元)"
				'end if 
				%></font>
         　</td>
		</tr>
	</table>
	<% 
	'---------------------------------------------------------------------------------------
	end if
	if (i+4 < Ubound(PBillSN)) or (i+4 = Ubound(PBillSN))then 
strBill="select b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress,c.StationName from billbasedcireturn a,Billbase b,Station c where a.DciReturnStation=c.DCIStationID(+) and a.BillNO=b.BillNo and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i+4)
set rsBill=conn.execute(strBill)
			GetMailAddress=""
			Billno=""
			StationName=""
			CarNo=""
			Owner=""
			Sys_BillTypeID=""
			OwnerZip=""
			 Sys_DriverHomeAddress=""
             Sys_DriverHomeZip=""
    	     Sys_Driver=""
			Sys_BillNo_BarCode=""
			ZipName2=""
		if not rsBill.eof then
		
	    	if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣"  then
				ZipName=""
			else
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing
			end if
				GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
			
			Billno=trim(rsBill("Billno"))
			StationName=trim(rsBill("StationName"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))
			Sys_BillTypeID=trim(rsBill("BillTypeID"))
			OwnerZip=trim(rsBill("OwnerZip"))
			
			Sys_DriverHomeAddress=ZipName2&trim(rsBill("DriverHomeAddress"))
            Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
			Sys_Driver=trim(rsBill("Driver"))

			Sys_BillNo_BarCode=BillNo
										Sys_DriverHomeAddress=replace(Sys_DriverHomeAddress,ZipName2&ZipName2,ZipName2)
										GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)
            	If sys_City="彰化縣" or sys_City="高雄市" or sys_City = "屏東縣" or sys_City="嘉義市" Then
				DelphiASPObj.GenSendStoreBillno BillNo,0,50,160,1

			else
				DelphiASPObj.GenSendStoreBillno BillNo,0,50,160

			end If 

			If sys_City="台東縣" Then 
				If Driver="" Or Sys_DriverHomeAddress="" Then 
				strBill="select a.Owner,a.OwnerZip,a.OwnerAddress,a.DriverHomeZip,a.DriverHomeAddress,a.Driver from billbasedcireturn a where a.ExchangeTypeID='A' and a.Carno='"&CarNo&"'"

				set rsBill=conn.execute(strBill)
					if not rsBill.eof Then
						ZipName=""
							strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=Nothing

							GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
							GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)

							strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
							set rsZip=conn.execute(strZip)
							if not rsZip.eof then
								ZipName2=trim(rsZip("ZipName"))
							end if
							rsZip.close
							set rsZip=nothing


						Owner=trim(rsBill("Owner"))
						OwnerZip=trim(rsBill("OwnerZip"))
						 Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
						 Sys_DriverHomeAddress=ZipName2&trim(rsBill("DriverHomeAddress"))
							Sys_DriverHomeAddress=replace(Sys_DriverHomeAddress,ZipName2&ZipName2,ZipName2)

						 Sys_Driver=trim(rsBill("Driver"))

					End if
				End If 
			End if
		end If
	rsBill.close
	set rsBill=nothing	
strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&PBillSN(i+4)&" and ExchangetypeID='W'"
set rsbil=conn.execute(strBil)
Sys_Owner=""
Sys_OwnerZip=""
Sys_OwnerZipName=""
Sys_OwnerAddress=""
Sys_BillNo=""
Sys_CarNo=""
Sys_Rule1=""
Sys_Rule2=""

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)

If Sys_BillTypeID="1" Then
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Sys_Owner=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	If ifnull(Sys_OwnerAddress) Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	end if
else
	If Sys_BillTypeID="1" Then
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
	End if
end if

If ifnull(Sys_OwnerAddress) Then
	strSql="select * from BillbaseDCIReturn where CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='A'"
	set rsdata=conn.execute(strsql)
	If Sys_BillTypeID="1" Then
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Driver"))
	else
		if Not rsdata.eof then Sys_Owner=trim(rsdata("Owner"))
	End if

	if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

			if Not rsdata.eof then Sys_OwnerAddress=trim(rsdata("OwnerAddress"))
			if Not rsdata.eof then Sys_OwnerZip=trim(rsdata("OwnerZip"))
	rsdata.close
end if

If ifnull(Sys_OwnerAddress) Then
	if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
	if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
end if

strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=replace(trim(rszip("ZipName")),"台","臺")
if Not rszip.eof then Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(rszip("ZipName")),"")
if Not rszip.eof then Sys_OwnerAddress=replace(Sys_OwnerAddress,trim(Sys_OwnerZipName),"")
rszip.close

Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")

If sys_City="彰化縣" or sys_City="高雄市" or sys_City = "屏東縣" or sys_City="嘉義市" Then
	DelphiASPObj.GenSendStoreBillno BillNo,0,60,160,1

else
	DelphiASPObj.GenSendStoreBillno BillNo,0,60,160

end If 
'-------------------------------------------------------------------------------------
	%>
	
		<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1") and trim(Sys_DriverHomeAddress)<>"" then
				  If Trim(Sys_Driver)="" Then 				
					response.write funcCheckFont(Owner,18,1) 
				  Else
					response.write funcCheckFont(Sys_Driver,18,1) 
				  End if
				else 
					response.write funcCheckFont(Owner,18,1)
				end if
				%>				
				<%if sys_City="彰化縣" then
				response.write "　先生/女士"
				end if
				%>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="50" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕"))  and trim(Sys_DriverHomeAddress)<>""  then 
				  If Trim(Sys_DriverHomeZip)="" Then 
					response.write OwnerZip 
				  Else
					response.write Sys_DriverHomeZip 
				  End if
				else
					response.write Sys_OwnerZip 
				end if
				%>&nbsp;<%
				if (Sys_BillTypeID="1" or (sys_City="宜蘭縣" and trim(Session("Ch_Name"))="楊玉燕"))  and trim(Sys_DriverHomeAddress)<>""  then 
				  If Trim(Sys_DriverHomeAddress)="" Then 
					response.write funcCheckFont(GetMailAddress,18,1)
				  Else
					response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
				  End if
				else 
					If Trim(Sys_DriverHomeAddress)="" Then 
						response.write replace(Sys_OwnerZipName&funcCheckFont(Sys_OwnerAddress,18,1) ,Sys_OwnerZipName&Sys_OwnerZipName,Sys_OwnerZipName)
					else
						response.write funcCheckFont(Sys_DriverHomeAddress,18,1)
					end if
				end If 

				'if sys_City="彰化縣" and Session("UnitLevelID")>"1" then
				'	response.write "(罰鍰："&Sum_Level&"元)"
				'end if 
				%></font>
         　</td>
		</tr>
	</table>
	
<% 
End if
end if
%>
<%next%>
</body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="../smsx.cab#Version=6,1,432,1">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
printWindow(true,5.08,5.08,5.08,5.08);
</script></p>