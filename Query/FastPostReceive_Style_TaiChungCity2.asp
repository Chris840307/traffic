<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>中華郵政掛號郵件收回執</title>
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
Server.ScriptTimeout=6000
PBillSN=split(trim(request("PBillSN")),",")
for i=0 to Ubound(PBillSN)
			GetMailAddress="" :Sys_DriverHomeAddress="" : Sys_DriverHomeZip=""
			Sys_Driver="" :Billno="" :CarNo="" :Owner=""
			Zip1="" :Zip2="":Zip3=""
			Sys_BillTypeID="" :	Zip11="":Zip21="":Zip31="":MailNumber=""
if cint(i)>0 and i mod 3=0 then response.write "<div class=""PageNext"">&nbsp;</div>"
if i mod 3=0 then 
strBill="select  d.unitname,b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress from billbasedcireturn a,Billbase b ,Unitinfo d where a.BillNO=b.BillNo and a.CarNo=b.Carno and a.ExchangeTypeID='W' and b.billunitid=d.unitid and b.SN="&PBillSN(i)
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
			end if
				GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
			
			Billno=trim(rsBill("Billno"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))
			Sys_BillTypeID=trim(rsBill("BillTypeID"))
			OwnerZip=trim(rsBill("OwnerZip"))

			unitname=trim(rsBill("unitname"))						

			OwnerZip=trim(rsBill("OwnerZip"))
			 Sys_DriverHomeAddress=trim(rsBill("DriverHomeAddress"))
                 Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
			     Sys_Driver=trim(rsBill("Driver"))
			
			Sys_BillNo_BarCode=BillNo
			
'			If sys_City<>"台中縣" Then
            	DelphiASPObj.GenSendStoreBillno BillNo,0,50,160
'             else
'            	Sys_BillNo_BarCode=Sys_BillNo_BarCode&"_4"
'            end if	
			
			
		end If
	rsBill.close
	set rsBill=nothing	
'-------------------------------------------------------------------------------------
strMailNumber="select MailNumber from BillMailHistory where BillSN="&PBillSN(i)
set rsMailNumber=conn.execute(strMailNumber)
		if not rsMailNumber.eof then
			MailNumber=trim(rsMailNumber("MailNumber"))
		end If
	rsMailNumber.close
	set rsMailNumber=nothing	

if (i< Ubound(PBillSN)) or (i = Ubound(PBillSN))then 
%>
<div id="R1" style="position:relative;">

	<!-- MSTableType="layout" -->
	<tr>
		<td colspan="2" valign="top">
		<p>　</td>
		<td rowspan="5" valign="top" style="border-style: solid; border-width: 1px" width="420">
		
		
	</p>
		
		
<div style="position: absolute; width: 424px; height: 326px; z-index: 9; left: 198px; top: 37px" id="layer14">
	<table border="0" width="100%" id="table1" cellspacing="0" height="100%" cellpadding="0">
		<tr>
			<td height="58">
			<div style="position: absolute; width: 344px; height: 20px; z-index: 5; left: 19px; top: 54px" id="layer29">
				<font face="標楷體"><%if Sys_BillTypeID="1"   then response.write Sys_Driver else response.write Owner%></font></div>



			<div style="position: absolute; width: 228px; height: 17px; z-index: 6; left: 29px; top: 277px" id="layer33">
				<font face="標楷體"><%=UnitName%></font></div>
			　</td>
		</tr>
		<tr>
			<td height="103">
			<div style="position: absolute; width: 225px; height: 35px; z-index: 6; left: 18px; top: 74px" id="layer30">
				<font face="標楷體"><%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>"" then response.write DriverHomeZip else response.write OwnerZip  %>&nbsp;<%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>""   then  response.write Sys_DriverHomeAddress else response.write GetMailAddress %></div>
			<div style="position: absolute; width: 303px; height: 31px; z-index: -1; left: 77px; top: 32px" id="layer31">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="..\BarCodeImage\<%=Sys_BillNo_BarCode%>.jpg">
</div>



			<div style="position: absolute; width: 130px; height: 17px; z-index: 16; left: 158px; top: 14px" id="layer32">
				<font face="標楷體"><%=MailNumber%></font></div>
			　</td>
		</tr>
		<tr>
			<td>
			　</td>
		</tr>
	</table>
</div>
　</td>
		<td height="67">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<p>　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		</p>
　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td valign="top" width="110">
		<p>　</td>
		<td valign="top" width="95">
		</p>
　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<p>　</td>
		<td height="67" width="102">　</td>
	</tr>

			</div>
<%
end if
	if (i+1 < Ubound(PBillSN)) or (i+1 = Ubound(PBillSN))then 
				GetMailAddress="" :Sys_DriverHomeAddress="" : Sys_DriverHomeZip=""
			Sys_Driver="" :Billno="" :CarNo="" :Owner=""
			Zip1="" :Zip2="":Zip3=""
			Sys_BillTypeID="" :	Zip11="":Zip21="":Zip31="":MailNumber=""

strBill="select  d.unitname,b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress from billbasedcireturn a,Billbase b ,billmailhistory c,Unitinfo d where a.BillNO=b.BillNo and a.CarNo=b.Carno and a.ExchangeTypeID='W' and b.billunitid=d.unitid and b.SN="&PBillSN(i+1)
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
			end if
				GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
			
			Billno=trim(rsBill("Billno"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))
			Sys_BillTypeID=trim(rsBill("BillTypeID"))
			OwnerZip=trim(rsBill("OwnerZip"))

			unitname=trim(rsBill("unitname"))						

			OwnerZip=trim(rsBill("OwnerZip"))
			 Sys_DriverHomeAddress=trim(rsBill("DriverHomeAddress"))
                 Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
			     Sys_Driver=trim(rsBill("Driver"))
			
			Sys_BillNo_BarCode=BillNo
			
'			If sys_City<>"台中縣" Then
            	DelphiASPObj.GenSendStoreBillno BillNo,0,50,160
'             else
'            	Sys_BillNo_BarCode=Sys_BillNo_BarCode&"_4"
'            end if	
			
			
		end If
	rsBill.close
	set rsBill=nothing	
'-------------------------------------------------------------------------------------
strMailNumber="select MailNumber from BillMailHistory where BillSN="&PBillSN(i+1)
set rsMailNumber=conn.execute(strMailNumber)
		if not rsMailNumber.eof then
			MailNumber=trim(rsMailNumber("MailNumber"))
		end If
	rsMailNumber.close
	set rsMailNumber=nothing	

%>
<div id="R1" style="position:relative;">

	<!-- MSTableType="layout" -->
	<tr>
		<td colspan="2" valign="top">
		<p>　</td>
		<td rowspan="5" valign="top" style="border-style: solid; border-width: 1px" width="420">
		
		
	</p>
		
		
<div style="position: absolute; width: 424px; height: 326px; z-index: 9; left: 199px; top: 138px" id="layer14">
	<table border="0" width="100%" id="table1" cellspacing="0" height="100%" cellpadding="0">
		<tr>
			<td height="58">
			<div style="position: absolute; width: 344px; height: 20px; z-index: 5; left: 15px; top: 46px" id="layer29">
				<font face="標楷體"><%if Sys_BillTypeID="1"   then response.write Sys_Driver else response.write Owner%></font></div>



			<div style="position: absolute; width: 228px; height: 17px; z-index: 6; left: 29px; top: 268px" id="layer33">
				<font face="標楷體"><%=UnitName%></font></div>
			　</td>
		</tr>
		<tr>
			<td height="103">
			<div style="position: absolute; width: 225px; height: 35px; z-index: 6; left: 15px; top: 65px" id="layer30">
				<font face="標楷體"><%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>"" then response.write DriverHomeZip else response.write OwnerZip  %>&nbsp;<%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>""   then  response.write Sys_DriverHomeAddress else response.write GetMailAddress %></div>
			<div style="position: absolute; width: 303px; height: 31px; z-index: -1; left: 77px; top: 23px" id="layer31">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="..\BarCodeImage\<%=Sys_BillNo_BarCode%>.jpg">
</div>



			<div style="position: absolute; width: 130px; height: 17px; z-index: 16; left: 158px; top: 6px" id="layer32">
				<font face="標楷體"><%=MailNumber%></font></div>
			　</td>
		</tr>
		<tr>
			<td>
			　</td>
		</tr>
	</table>
</div>
　</td>
		<td height="67">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<p>　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		</p>
　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td valign="top" width="110">
		<p>　</td>
		<td valign="top" width="95">
		</p>
　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<p>　</td>
		<td height="67" width="102">　</td>
	</tr>

			</div>
<%
end if
	if (i+2 < Ubound(PBillSN)) or (i+2 = Ubound(PBillSN))then 
			GetMailAddress="" :Sys_DriverHomeAddress="" : Sys_DriverHomeZip=""
			Sys_Driver="" :Billno="" :CarNo="" :Owner=""
			Zip1="" :Zip2="":Zip3=""
			Sys_BillTypeID="" :	Zip11="":Zip21="":Zip31="":MailNumber=""
strBill="select  d.unitname,b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress from billbasedcireturn a,Billbase b ,billmailhistory c,Unitinfo d where a.BillNO=b.BillNo and a.CarNo=b.Carno and a.ExchangeTypeID='W' and b.billunitid=d.unitid and b.SN="&PBillSN(i+2)
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
			end if
				GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
			
			Billno=trim(rsBill("Billno"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))
			Sys_BillTypeID=trim(rsBill("BillTypeID"))
			OwnerZip=trim(rsBill("OwnerZip"))

			unitname=trim(rsBill("unitname"))						

			OwnerZip=trim(rsBill("OwnerZip"))
			 Sys_DriverHomeAddress=trim(rsBill("DriverHomeAddress"))
                 Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
			     Sys_Driver=trim(rsBill("Driver"))
			
			Sys_BillNo_BarCode=BillNo
			
'			If sys_City<>"台中縣" Then
            	DelphiASPObj.GenSendStoreBillno BillNo,0,50,160
'             else
'            	Sys_BillNo_BarCode=Sys_BillNo_BarCode&"_4"
'            end if	
			
			
		end If
	rsBill.close
	set rsBill=nothing	
'-------------------------------------------------------------------------------------
strMailNumber="select MailNumber from BillMailHistory where BillSN="&PBillSN(i+2)
set rsMailNumber=conn.execute(strMailNumber)
		if not rsMailNumber.eof then
			MailNumber=trim(rsMailNumber("MailNumber"))
		end If
	rsMailNumber.close
	set rsMailNumber=nothing	

%>
<div id="R1" style="position:relative;">

	<!-- MSTableType="layout" -->
	<tr>
		<td colspan="2" valign="top">
		<p>　</td>
		<td rowspan="5" valign="top" style="border-style: solid; border-width: 1px" width="420">
		
		
	</p>
		
		
<div style="position: absolute; width: 424px; height: 326px; z-index: 9; left: 198px; top: 220px" id="layer14">
	<table border="0" width="100%" id="table1" cellspacing="0" height="100%" cellpadding="0">
		<tr>
			<td height="58">
			<div style="position: absolute; width: 344px; height: 20px; z-index: 5; left: 13px; top: 57px" id="layer29">
				<font face="標楷體"><%if Sys_BillTypeID="1"   then response.write Sys_Driver else response.write Owner%></font></div>



			<div style="position: absolute; width: 228px; height: 17px; z-index: 6; left: 29px; top: 281px" id="layer33">
				<font face="標楷體"><%=UnitName%></font></div>
			　</td>
		</tr>
		<tr>
			<td height="103">
			<div style="position: absolute; width: 225px; height: 35px; z-index: 6; left: 13px; top: 77px" id="layer30">
				<font face="標楷體"><%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>"" then response.write DriverHomeZip else response.write OwnerZip  %>&nbsp;<%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>""   then  response.write Sys_DriverHomeAddress else response.write GetMailAddress %></div>
			<div style="position: absolute; width: 303px; height: 31px; z-index: -1; left: 77px; top: 35px" id="layer31">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="..\BarCodeImage\<%=Sys_BillNo_BarCode%>.jpg">
</div>



			<div style="position: absolute; width: 130px; height: 17px; z-index: 16; left: 158px; top: 17px" id="layer32">
				<font face="標楷體"><%=MailNumber%></font></div>
			　</td>
		</tr>
		<tr>
			<td>
			　</td>
		</tr>
	</table>
</div>
　</td>
		<td height="67">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<p>　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		</p>
　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td valign="top" width="110">
		<p>　</td>
		<td valign="top" width="95">
		</p>
　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<p>　</td>
		<td height="67" width="102">　</td>
	</tr>

			</div>
<%end if%>
</div>
<%
end if

	if (i mod 100)=0 then response.flush

next
%>

</body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="../smsx.cab#Version=6,1,432,1">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
printWindow(true,5.08,5.08,5.08,5.08);
</script></p>