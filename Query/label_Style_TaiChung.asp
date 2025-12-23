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

<body>

<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
'sys_City="台中縣"
'--------------------------------------------------------------------------------------------------------------------
'登入者、單位地址
	strUNit="select UnitName,Address from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
	set rsUNit=conn.execute(strUNit)
		if not rsUNit.eof then
			UnitName=trim(rsUNit("UnitName"))
			Address=trim(rsUNit("Address"))
		end If
	rsUNit.close
	set rsUNit=nothing	
'縣市名稱	
strCityName="select value from apconfigure where name='縣市名稱'"
set rsCityName=conn.execute(strCityName)
		if not rsCityName.eof then
			CityName=trim(rsCityName("value"))
		end If
	rsCityName.close
	set rsCityName=nothing	
'管轄郵遞區號
strCode="select value from apconfigure where name='管轄郵遞區號'"
set rsCode=conn.execute(strCode)
		if not rsCode.eof then
			Code=trim(rsCode("value"))
		end If
	rsCode.close
	set rsCode=nothing	
	Set DelphiASPObj = Server.CreateObject("GenBarCode.BarCodeASP")
'--------------------------------------------------------------------------------------------------------------------

PBillSN=split(trim(request("PBillSN")),",")
for i=0 to Ubound(PBillSN)

if cint(i)>0  then response.write "<div class=""PageNext"">&nbsp;</div>"



'---------------------------------------------------------------------------------------
strBill="select b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress,c.StationName from billbasedcireturn a,Billbase b,Station c where a.DciReturnStation=c.DCIStationID(+) and a.BillNO=b.BillNo and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i)

set rsBill=conn.execute(strBill)
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
			
            	DelphiASPObj.GenSendStoreBillno BillNo,0,50,160
			
		end If
	rsBill.close
	set rsBill=nothing	
'-------------------------------------------------------------------------------------



%>
	
	<table border="0" width="80%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td height="100" width="60%">
				<p align="center">
				<font face="標楷體" size="5"><%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>""   then response.write Sys_Driver else response.write Owner%></font></td>
    <td>
              <p align="left">
              <img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg"><font face="標楷體" size="2"><%if sys_City="台中縣" then  response.write StationName%></font></td>
    		</tr>
	    	<tr>
			<td height="120" align="left" valign="top" colspan="2">
				<font face="標楷體" size="5"><%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>""  then response.write Sys_DriverHomeZip else response.write OwnerZip  %>&nbsp;<%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>""   then  response.write Sys_DriverHomeAddress else response.write GetMailAddress %></font>
         　</td>
		</tr>
	</table>
<%next%>
</body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="../smsx.cab#Version=6,1,432,1">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
printWindow(true,25,50,5.08,5.08);
</script></p>