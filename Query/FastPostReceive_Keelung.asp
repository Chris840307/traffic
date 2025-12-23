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
<!--#include virtual="traffic/Common/AllFunction.inc"-->
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
strBill="select b.Billno,b.CarNo,a.Owner,a.Driver,a.OwnerZip,a.DriverHomeAddress,a.DriverHomeZip,a.OwnerZip,a.OwnerAddress,substr(a.OwnerZip,1,1) as Zip1,substr(a.OwnerZip,2,1) as Zip2,substr(a.OwnerZip,3,1) as Zip3,substr(a.DriverHomeZip,1,1) as Zip11,substr(a.DriverHomeZip,2,1) as Zip21,substr(a.DriverHomeZip,3,1) as Zip31,b.BillTypeID from billbasedcireturn a,Billbase b where a.BillNO=b.BillNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i)
set rsBill=conn.execute(strBill)
		if not rsBill.eof then
		      		      ZipName=""
						strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing

			      ZipName2=""
						strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing


				GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))


				if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣"  then
				Sys_DriverZipName=""
			else
				strDZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsDZip=conn.execute(strDZip)
				if not rsDZip.eof then
					Sys_DriverZipName=trim(rsDZip("ZipName"))
				end if
				rsDZip.close
				set rsDZip=nothing
			end If
			


                 Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
			     Sys_Driver=trim(rsBill("Driver"))
Sys_DriverHomeAddress=trim(rsBill("DriverHomeAddress"))
			Billno=trim(rsBill("Billno"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))

			Zip1=trim(rsBill("Zip1"))
			Zip2=trim(rsBill("Zip2"))
			Zip3=trim(rsBill("Zip3"))
						DriverHomeZip=trim(rsBill("DriverHomeZip"))
						OwnerZip=trim(rsBill("OwnerZip"))	
			Sys_BillTypeID=trim(rsBill("BillTypeID"))

			Zip11=trim(rsBill("Zip11"))
			Zip21=trim(rsBill("Zip21"))
			Zip31=trim(rsBill("Zip31"))
            Sys_BillNo_BarCode=BillNo 
          	DelphiASPObj.GenSendStoreBillno BillNo,0,41,160


		end If
	rsBill.close
	set rsBill=nothing	
'-------------------------------------------------------------------------------------
If Instr(request("Sys_BatchNumber"),"N")>0 then
	strMailNumber="select StoreAndSendMailNumber as MailNumber from BillMailHistory where BillSN="&PBillSN(i)
else
	strMailNumber="select MailNumber from BillMailHistory where BillSN="&PBillSN(i)
end if

set rsMailNumber=conn.execute(strMailNumber)
if not rsMailNumber.eof then
	MailNumber=trim(rsMailNumber("MailNumber"))&" 200016 36"
end If
rsMailNumber.close
set rsMailNumber=nothing
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
strSql="select * from BillbaseDCIReturn where BillNo='"&trim(Billno)&"' and CarNo='"&trim(CarNo)&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)
If Sys_BillTypeID=1 Then
	if Not rsFound.eof then Owner=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Owner=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsFound.eof then GetMailAddress=ZipName2&trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then OwnerZip=trim(rsfound("DriverHomeZip"))
else
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then GetMailAddress=ZipName2&trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then OwnerZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then GetMailAddress=ZipName&trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then OwnerZip=trim(rsfound("OwnerZip"))
	End if
end if

If ifnull(GetMailAddress) Then
	strSql="select * from BillbaseDCIReturn where CarNo='"&trim(CarNo)&"' and ExchangetypeID='A'"
	set rsdata=conn.execute(strsql)
	If Sys_BillTypeID=1 Then
		if Not rsdata.eof then Owner=trim(rsdata("Driver"))
	else
		if Not rsdata.eof then Owner=trim(rsdata("Owner"))
	End if

	if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		if Not rsdata.eof then GetMailAddress=ZipName2&trim(rsdata("DriverHomeAddress"))
		if Not rsdata.eof then OwnerZip=trim(rsdata("DriverHomeZip"))
	else
		If Sys_BillTypeID=1 Then
			if Not rsdata.eof then GetMailAddress=ZipName2&trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then OwnerZip=trim(rsdata("DriverHomeZip"))
		else
			if Not rsdata.eof then GetMailAddress=ZipName&trim(rsdata("OwnerAddress"))
			if Not rsdata.eof then OwnerZip=trim(rsdata("OwnerZip"))
		End if
	end if
	rsdata.close
end if

If ifnull(GetMailAddress) Then
	if Not rsfound.eof then Owner=trim(rsfound("Owner"))
	if Not rsfound.eof then GetMailAddress=ZipName&trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then OwnerZip=trim(rsfound("OwnerZip"))
end if
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if (i< Ubound(PBillSN)) or (i = Ubound(PBillSN))then 
            ZipName=replace(ZipName,"臺","台")
			GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)
			GetMailAddress=funcCheckFont(replace(GetMailAddress&"","臺","台"),16,1)
            Sys_DriverZipName    =replace(Sys_DriverZipName,"臺","台")
			Sys_DriverHomeAddress=replace(Sys_DriverZipName&Sys_DriverHomeAddress,Sys_DriverZipName&Sys_DriverZipName,Sys_DriverZipName)
			Sys_DriverHomeAddress=funcCheckFont(replace(Sys_DriverHomeAddress&"","臺","台"),16,1)   
Sys_Driver=funcCheckFont(Sys_Driver,16,1)			
Owner=funcCheckFont(Owner,16,1)
%>
<div id="R1" style="position:relative;">
<table border="2" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="608" id="AutoNumber1" height="278">
  <tr>
    <td width="604" colspan="4" height="20">
    <p align="center"><font face="標楷體">中華郵政掛號郵件收回執</font></td>
  </tr>
  <tr>
    <td width="604" colspan="4" height="15">　　　<font face="標楷體">郵件種類　　　　　　　　　　號碼<%=MailNumber%>　　<font size="2">(由郵局收寄人員填寫)</font></font></td>
  </tr>
  <tr>
    <td width="93" height="63"> <b>&nbsp;<font size="2">收件人姓名<br>&nbsp;地址</font></b><font size="2"><font face="標楷體">(請寄</font><br>
	</font><font size="2" face="標楷體">&nbsp;件人填寫)</font></td>
    <td width="407" height="63" colspan="2">
    <table border="0" width="100%" id="table3" height="60" cellspacing="0" cellpadding="0">
		<tr>
			<td width="348">
    <font face="標楷體">&nbsp;&nbsp;&nbsp;<%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write DriverHomeZip else response.write OwnerZip%>&nbsp;&nbsp;

    </font></td>
			<td rowspan="3" valign="top">
<font size="4" face="標楷體">小姐 </font>
			<br>
<font size="4" face="標楷體">先生</font></td>
		</tr>
		<tr>
			<td width="348">
    　<font face="標楷體">&nbsp;&nbsp;&nbsp;<%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write Sys_Driver else response.write Owner%></font> </td>
		</tr>
		<tr>
			<td width="348" height="26">

			<div style="position: absolute; width: 548px; height: 24px; z-index: 8; left: 86px; top: 83px" id="layer31">
    　<font face="標楷體"><%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then  response.write Sys_DriverHomeAddress else response.write GetMailAddress %></font></td>
    </div>

		</tr>
	</table>
	</td>
    <td width="102" height="242" rowspan="3" valign="bottom">
    <div id="L2" style="position:absolute; left:503px;top:42px;width:215px; height:274px">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="105" id="AutoNumber2" height="266">
      <tr>
        <td width="103" height="119" valign="bottom">
<font face="標楷體">收寄局郵戳</font></td>
      </tr>
      <tr>
        <td width="103" height="148" valign="bottom">
<font face="標楷體">投遞後郵戳</font></td>
      </tr>
    </table>
    </div>
    </td>
  </tr>
  <tr>
    <td width="93" height="180" rowspan="2"><b><font size="5" face="標楷體">&nbsp;投&nbsp;遞</font></b><br><br><b><font size="5" face="標楷體">&nbsp;記&nbsp;要</font></b></td>
    <td width="77" height="76">
    <p align="center"><font face="標楷體">請收件<br>人填寫</font></td>
    <td width="329" height="76">
	<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td>
<font face="標楷體" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 年　　月　　日收到第　　　 &nbsp;&nbsp;&nbsp; 號 </font>
			</td>
		</tr>
		<tr>
			<td>
<font face="標楷體" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 掛號郵件壹件 </font>
			</td>
		</tr>
		<tr>
			<td><font face="標楷體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;□本人&nbsp;□代收</font></td>
		</tr>
		<tr>
			<td height="18">
<font face="標楷體" size="2">&nbsp;&nbsp; 收件人</font></td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;&nbsp; 蓋　章</font><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="標楷體">投遞士戳</font>
</font>
</td>
		</tr>
	</table>
	</td>
  </tr>
  <tr>
    <td width="407" height="105" colspan="2">
	<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="0">
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;&nbsp; (供查詢時填寫)</font><font size="2"> </font>
			</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;□經查上述郵件已於　　年　　月　　日妥投</font><font size="2">
</font>
</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 君收訖</font><font size="2">
</font>
</td>
		</tr>
		<tr>
			<td height="18">
<font size="2" face="標楷體">&nbsp;&nbsp; 該機構收發單位代收訖</font><font size="2"> </font>
</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;□附上原掛號收據影印本一件　請查收</font><font size="2"> </font>
</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;□</font><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="標楷體">郵　局</font>
</font>
</td>
		</tr>
		<tr>
			<td height="18">
			<p align="left">
<font face="標楷體" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 年　　月　　日</font></td>
		</tr>
	</table>
	</td>
  </tr>
  </table>

    
    
<!--第三區-->

<font face="標楷體" size="3">該回執聯請退回<%=UnitName%>&nbsp;&nbsp;<%=Code%>&nbsp;&nbsp;<%=Address%></font>
<br><br>
			

			<div style="position: absolute; width: 235px; height: 33; z-index: 8; left: 283px; top: 44px" id="layer31">
<img src="..\BarCodeImage\<%=Sys_BillNo_BarCode%>.jpg">
</div>


<%
end if
	if (i+1 < Ubound(PBillSN)) or (i+1 = Ubound(PBillSN))then 
strBill="select b.Billno,b.CarNo,a.Owner,a.Driver,a.OwnerZip,a.DriverHomeAddress,a.OwnerAddress,a.DriverHomeZip,a.OwnerZip,substr(a.OwnerZip,1,1) as Zip1,substr(a.OwnerZip,2,1) as Zip2,substr(a.OwnerZip,3,1) as Zip3,substr(a.DriverHomeZip,1,1) as Zip11,substr(a.DriverHomeZip,2,1) as Zip21,substr(a.DriverHomeZip,3,1) as Zip31,b.BillTypeID from billbasedcireturn a,Billbase b where a.BillNO=b.BillNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i+1)
set rsBill=conn.execute(strBill)
		if not rsBill.eof then
		      		      ZipName=""
						strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing

		      		      ZipName2=""
						strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing

				GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
				 Sys_DriverHomeAddress=trim(rsBill("DriverHomeAddress"))
                 Sys_DriverHomeZip=ZipName2&trim(rsBill("DriverHomeZip"))
			     Sys_Driver=trim(rsBill("Driver"))

			Billno=trim(rsBill("Billno"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))

			Zip1=trim(rsBill("Zip1"))
			Zip2=trim(rsBill("Zip2"))
			Zip3=trim(rsBill("Zip3"))
						DriverHomeZip=trim(rsBill("DriverHomeZip"))
						OwnerZip=trim(rsBill("OwnerZip"))						

			Sys_BillTypeID=trim(rsBill("BillTypeID"))

			Zip11=trim(rsBill("Zip11"))
			Zip21=trim(rsBill("Zip21"))
			Zip31=trim(rsBill("Zip31"))
			            Sys_BillNo_BarCode=BillNo 
			          	DelphiASPObj.GenSendStoreBillno BillNo,0,41,160
		end If
	rsBill.close
	set rsBill=nothing	
'-------------------------------------------------------------------------------------
If Instr(request("Sys_BatchNumber"),"N")>0 then
	strMailNumber="select StoreAndSendMailNumber as MailNumber from BillMailHistory where BillSN="&PBillSN(i+1)
else
	strMailNumber="select MailNumber from BillMailHistory where BillSN="&PBillSN(i+1)
end if

set rsMailNumber=conn.execute(strMailNumber)
if not rsMailNumber.eof then
	MailNumber=trim(rsMailNumber("MailNumber"))&" 200016 36"
end If
rsMailNumber.close
set rsMailNumber=nothing
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
strSql="select * from BillbaseDCIReturn where BillNo='"&trim(Billno)&"' and CarNo='"&trim(CarNo)&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)
If Sys_BillTypeID=1 Then
	if Not rsFound.eof then Owner=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Owner=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsFound.eof then GetMailAddress=ZipName2&trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then OwnerZip=trim(rsfound("DriverHomeZip"))
else
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then GetMailAddress=ZipName2&trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then OwnerZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then GetMailAddress=ZipName&trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then OwnerZip=trim(rsfound("OwnerZip"))
	End if
end if

If ifnull(GetMailAddress) Then
	strSql="select * from BillbaseDCIReturn where CarNo='"&trim(CarNo)&"' and ExchangetypeID='A'"
	set rsdata=conn.execute(strsql)
	If Sys_BillTypeID=1 Then
		if Not rsdata.eof then Owner=trim(rsdata("Driver"))
	else
		if Not rsdata.eof then Owner=trim(rsdata("Owner"))
	End if

	if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		if Not rsdata.eof then GetMailAddress=trim(rsdata("DriverHomeAddress"))
		if Not rsdata.eof then OwnerZip=trim(rsdata("DriverHomeZip"))
	else
		If Sys_BillTypeID=1 Then
			if Not rsdata.eof then GetMailAddress=ZipName2&trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then OwnerZip=trim(rsdata("DriverHomeZip"))
		else
			if Not rsdata.eof then GetMailAddress=trim(rsdata("OwnerAddress"))
			if Not rsdata.eof then OwnerZip=trim(rsdata("OwnerZip"))
		End if
	end if
	rsdata.close
end if

If ifnull(GetMailAddress) Then
	if Not rsfound.eof then Owner=trim(rsfound("Owner"))
	if Not rsfound.eof then GetMailAddress=ZipName&trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then OwnerZip=trim(rsfound("OwnerZip"))
end if
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ZipName=replace(ZipName,"臺","台")
			GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)
			GetMailAddress=funcCheckFont(replace(GetMailAddress&"","臺","台"),16,1)
            Sys_DriverZipName    =replace(Sys_DriverZipName,"臺","台")
			Sys_DriverHomeAddress=replace(Sys_DriverZipName&Sys_DriverHomeAddress,Sys_DriverZipName&Sys_DriverZipName,Sys_DriverZipName)
			Sys_DriverHomeAddress=funcCheckFont(replace(Sys_DriverHomeAddress&"","臺","台"),16,1)   
Sys_Driver=funcCheckFont(Sys_Driver,16,1)			
Owner=funcCheckFont(Owner,16,1)
%>
<table border="2" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="608" id="table7" height="274">
  <tr>
    <td width="604" colspan="4" height="20">
    <p align="center"><font face="標楷體">中華郵政掛號郵件收回執</font></td>
  </tr>
  <tr>
    <td width="604" colspan="4" height="15">　　　<font face="標楷體">郵件種類　　　　　　　　　　號碼<%=MailNumber%>　　<font size="2">(由郵局收寄人員填寫)</font></font></td>
  </tr>
  <tr>
    <td width="93" height="59"> <b><font size="2">&nbsp;收件人姓名<br>&nbsp;地址</font></b><font size="2"><font face="標楷體">(請寄</font><br>
	</font><font size="2" face="標楷體">&nbsp;件人填寫)</font></td>
    <td width="407" height="59" colspan="2">
    <table border="0" width="100%" id="table8" height="53" cellspacing="0" cellpadding="0">
		<tr>
			<td width="336">
    <font face="標楷體">&nbsp;&nbsp;&nbsp;<%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write DriverHomeZip else response.write OwnerZip%>
    </font></td>
			<td rowspan="3" valign="top">
<font size="4" face="標楷體">小姐 </font>
<br>
<font size="4" face="標楷體">先生</font></td>
		</tr>
		<tr>
			<td width="336">
    　<font face="標楷體">&nbsp;&nbsp;&nbsp;<%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write Sys_Driver else response.write Owner%></font> </td></td>
		</tr>
		<tr>
			<td width="336">
    　			
	<div style="position: absolute; width: 548px; height: 24px; z-index: 8; left: 86px; top: 431px" id="layer31">
    　<font face="標楷體"><%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then  response.write Sys_DriverHomeAddress else response.write GetMailAddress %></font>
    </div></td>
		</tr>
	</table>
	</td>
    <td width="102" height="238" rowspan="3" valign="bottom">
    　</td>
  </tr>
  <tr>
    <td width="93" height="180" rowspan="2"><b><font size="5" face="標楷體">&nbsp;投&nbsp;遞</font></b><br><br><b><font size="5" face="標楷體">&nbsp;記&nbsp;要</font></b></td>
    <td width="77" height="76">
    <p align="center"><font face="標楷體">請收件<br>人填寫</font></td>
    <td width="329" height="76">
	<table border="0" width="100%" id="table10" cellspacing="0" cellpadding="0">
		<tr>
			<td>
<font face="標楷體" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 年　　月　　日收到第　　　 &nbsp;&nbsp;&nbsp; 號 </font>
			</td>
		</tr>
		<tr>
			<td>
<font face="標楷體" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;掛號郵件壹件 </font>
			</td>
		</tr>
		<tr>
			<td><font face="標楷體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;□本人&nbsp;□代收
		</tr>
		<tr>
			<td height="18">
<font face="標楷體" size="2">&nbsp;&nbsp; 收件人</font></td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;&nbsp; 蓋　章</font><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="標楷體">投遞士戳</font>
</font>
</td>
		</tr>
	</table>
	</td>
  </tr>
  <tr>
    <td width="407" height="105" colspan="2">
	<table border="0" width="100%" id="table11" cellspacing="0" cellpadding="0">
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;&nbsp; (供查詢時填寫)</font><font size="2"> </font>
			</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;□經查上述郵件已於　　年　　月　　日妥投</font><font size="2">
</font>
</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 君收訖</font><font size="2">
</font>
</td>
		</tr>
		<tr>
			<td height="18">
<font size="2" face="標楷體">&nbsp;&nbsp; 該機構收發單位代收訖</font><font size="2"> </font>
</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;□附上原掛號收據影印本一件　請查收</font><font size="2"> </font>
</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;□</font><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="標楷體">郵　局</font>
</font>
</td>
		</tr>
		<tr>
			<td height="18">
			<p align="left">
<font face="標楷體" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 年　　月　　日</font></td>
		</tr>
	</table>
	</td>
  </tr>
  </table>



<font face="標楷體" size="3">該回執聯請退回<%=UnitName%>&nbsp;&nbsp;<%=Code%>&nbsp;&nbsp;<%=Address%></font>
			<br><br>
			<div style="position: absolute; width: 233px; height: 33; z-index: 8; left: 272px; top: 393px" id="layer32">
<img src="..\BarCodeImage\<%=Sys_BillNo_BarCode%>.jpg">
</div>
<div id="L3" style="position:absolute; left:503px;top:393px;width:215px; height:270px">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="105" id="table12" height="260">
      <tr>
        <td width="103" height="119" valign="bottom">
<font face="標楷體">收寄局郵戳</font></td>
      </tr>
      <tr>
        <td width="103" height="142" valign="bottom">
<font face="標楷體">投遞後郵戳</font></td>
      </tr>
    </table>
    </div>
<%
end if

	if (i+2 < Ubound(PBillSN)) or (i+2 = Ubound(PBillSN))then 
strBill="select b.Billno,b.CarNo,a.Owner,a.Driver,a.OwnerZip,a.DriverHomeAddress,a.OwnerAddress,a.DriverHomeZip,a.OwnerZip,substr(a.OwnerZip,1,1) as Zip1,substr(a.OwnerZip,2,1) as Zip2,substr(a.OwnerZip,3,1) as Zip3,substr(a.DriverHomeZip,1,1) as Zip11,substr(a.DriverHomeZip,2,1) as Zip21,substr(a.DriverHomeZip,3,1) as Zip31,b.BillTypeID from billbasedcireturn a,Billbase b where a.BillNO=b.BillNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i+2)
set rsBill=conn.execute(strBill)
		if not rsBill.eof then
		      ZipName=""
						strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing

		      ZipName2=""
						strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=nothing

				GetMailAddress=ZipName&trim(rsBill("OwnerAddress"))
				 Sys_DriverHomeAddress=ZipName2&trim(rsBill("DriverHomeAddress"))
                 Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
			     Sys_Driver=trim(rsBill("Driver"))

			Billno=trim(rsBill("Billno"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))

			Zip1=trim(rsBill("Zip1"))
			Zip2=trim(rsBill("Zip2"))
			Zip3=trim(rsBill("Zip3"))
						DriverHomeZip=trim(rsBill("DriverHomeZip"))
						OwnerZip=trim(rsBill("OwnerZip"))						

			Sys_BillTypeID=trim(rsBill("BillTypeID"))

			Zip11=trim(rsBill("Zip11"))
			Zip21=trim(rsBill("Zip21"))
			Zip31=trim(rsBill("Zip31"))
			            Sys_BillNo_BarCode=BillNo 
			          	DelphiASPObj.GenSendStoreBillno BillNo,0,41,160
		end If
	rsBill.close
	set rsBill=nothing	
'-------------------------------------------------------------------------------------
If Instr(request("Sys_BatchNumber"),"N")>0 then
	strMailNumber="select StoreAndSendMailNumber as MailNumber from BillMailHistory where BillSN="&PBillSN(i+2)
else
	strMailNumber="select MailNumber from BillMailHistory where BillSN="&PBillSN(i+2)
end if

set rsMailNumber=conn.execute(strMailNumber)
if not rsMailNumber.eof then
	MailNumber=trim(rsMailNumber("MailNumber"))&" 200016 36"
end If
rsMailNumber.close
set rsMailNumber=nothing
'------------------------------------------------------------------------------------------------------------------------------------------------------------------------
strSql="select * from BillbaseDCIReturn where BillNo='"&trim(Billno)&"' and CarNo='"&trim(CarNo)&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)
If Sys_BillTypeID=1 Then
	if Not rsFound.eof then Owner=trim(rsfound("Driver"))
else
	if Not rsFound.eof then Owner=trim(rsfound("Owner"))
End if
if Not rsfound.eof then Sys_DriverID=trim(rsfound("DriverID"))
if Instr(request("Sys_BatchNumber"),"N")>0 then
	if Not rsFound.eof then GetMailAddress=ZipName2&trim(rsfound("DriverHomeAddress"))
	if Not rsFound.eof then OwnerZip=trim(rsfound("DriverHomeZip"))
else
	If Sys_BillTypeID=1 Then
		if Not rsFound.eof then GetMailAddress=ZipName2&trim(rsfound("DriverHomeAddress"))
		if Not rsFound.eof then OwnerZip=trim(rsfound("DriverHomeZip"))
	else
		if Not rsFound.eof then GetMailAddress=ZipName&trim(rsfound("OwnerAddress"))
		if Not rsFound.eof then OwnerZip=trim(rsfound("OwnerZip"))
	End if
end if

If ifnull(GetMailAddress) Then
	strSql="select * from BillbaseDCIReturn where CarNo='"&trim(CarNo)&"' and ExchangetypeID='A'"
	set rsdata=conn.execute(strsql)
	If Sys_BillTypeID=1 Then
		if Not rsdata.eof then Owner=trim(rsdata("Driver"))
	else
		if Not rsdata.eof then Owner=trim(rsdata("Owner"))
	End if

	if Not rsdata.eof then Sys_DriverID=trim(rsdata("DriverID"))

	if Instr(request("Sys_BatchNumber"),"N")>0 then
		if Not rsdata.eof then GetMailAddress=ZipName2&trim(rsdata("DriverHomeAddress"))
		if Not rsdata.eof then OwnerZip=trim(rsdata("DriverHomeZip"))
	else
		If Sys_BillTypeID=1 Then
			if Not rsdata.eof then GetMailAddress=ZipName2&trim(rsdata("DriverHomeAddress"))
			if Not rsdata.eof then OwnerZip=trim(rsdata("DriverHomeZip"))
		else
			if Not rsdata.eof then GetMailAddress=ZipName&trim(rsdata("OwnerAddress"))
			if Not rsdata.eof then OwnerZip=trim(rsdata("OwnerZip"))
		End if
	end if
	rsdata.close
end if

If ifnull(GetMailAddress) Then
	if Not rsfound.eof then Owner=trim(rsfound("Owner"))
	if Not rsfound.eof then GetMailAddress=ZipName&trim(rsfound("OwnerAddress"))
	if Not rsfound.eof then OwnerZip=trim(rsfound("OwnerZip"))
end if
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            ZipName=replace(ZipName,"臺","台")
			GetMailAddress=replace(GetMailAddress,ZipName&ZipName,ZipName)
			GetMailAddress=funcCheckFont(replace(GetMailAddress&"","臺","台"),16,1)
            Sys_DriverZipName    =replace(Sys_DriverZipName,"臺","台")
			Sys_DriverHomeAddress=replace(Sys_DriverZipName&Sys_DriverHomeAddress,Sys_DriverZipName&Sys_DriverZipName,Sys_DriverZipName)
			Sys_DriverHomeAddress=funcCheckFont(replace(Sys_DriverHomeAddress&"","臺","台"),16,1)   
Sys_Driver=funcCheckFont(Sys_Driver,16,1)			
Owner=funcCheckFont(Owner,16,1)
%>

<table border="2" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="608" id="table17" height="274">
  <tr>
    <td width="604" colspan="4" height="20">
    <p align="center"><font face="標楷體">中華郵政掛號郵件收回執</font></td>
  </tr>
  <tr>
    <td width="604" colspan="4" height="15">　　　<font face="標楷體">郵件種類　　　　　　　　　　號碼<%=MailNumber%>　　<font size="2">(由郵局收寄人員填寫)</font></font></td>
  </tr>
  <tr>
    <td width="93" height="59"> <b><font size="2">&nbsp;收件人姓名<br>&nbsp;地址</font></b><font size="2"><font face="標楷體">(請寄</font><br>
	</font><font size="2" face="標楷體">&nbsp;件人填寫)</font></td>
    <td width="407" height="59" colspan="2">
    <table border="0" width="100%" id="table18" height="53" cellspacing="0" cellpadding="0">
		<tr>
			<td width="336">
    <font face="標楷體">&nbsp;&nbsp;&nbsp;<%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write DriverHomeZip else response.write OwnerZip%>
    </font></td>
			<td rowspan="3" valign="top">
<font size="4" face="標楷體">小姐 </font>
			<br>
<font size="4" face="標楷體">先生</font></td>
		</tr>
		<tr>
			<td width="336">
    　<font face="標楷體">&nbsp;&nbsp;&nbsp;<%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write Sys_Driver else response.write Owner%></font> </td></td>
		</tr>
		<tr>
			<td width="336" height="14">
			<div style="position: absolute; width: 548px; height: 24px; z-index: 8; left: 85px; top: 780px" id="layer31">
    　<font face="標楷體"><%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then  response.write Sys_DriverHomeAddress else response.write GetMailAddress %></font></td>
    </div>
		</tr>
	</table>
	</td>
    <td width="102" height="238" rowspan="3" valign="bottom">
    　</td>
  </tr>
  <tr>
    <td width="93" height="180" rowspan="2"><b><font size="5" face="標楷體">&nbsp;投&nbsp;遞</font></b><br><br><b><font size="5" face="標楷體">&nbsp;記&nbsp;要</font></b></td>
    <td width="77" height="76">
    <p align="center"><font face="標楷體">請收件<br>人填寫</font></td>
    <td width="329" height="76">
	<table border="0" width="100%" id="table19" cellspacing="0" cellpadding="0">
		<tr>
			<td>
<font face="標楷體" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 年　　月　　日收到第　　　 &nbsp;&nbsp;&nbsp; 號 </font>
			</td>
		</tr>
		<tr>
			<td>
<font face="標楷體" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;掛號郵件壹件 </font>
			</td>
		</tr>
		<tr>
			<td><font face="標楷體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;□本人&nbsp;□代收
		</tr>
		<tr>
			<td height="18">
<font face="標楷體" size="2">&nbsp;&nbsp; 收件人</font></td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;&nbsp; 蓋　章</font><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="標楷體">投遞士戳</font>
</font>
</td>
		</tr>
	</table>
	</td>
  </tr>
  <tr>
    <td width="407" height="105" colspan="2">
	<table border="0" width="100%" id="table20" cellspacing="0" cellpadding="0">
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;&nbsp; (供查詢時填寫)</font><font size="2"> </font>
			</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;□經查上述郵件已於　　年　　月　　日妥投</font><font size="2">
</font>
</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 君收訖</font><font size="2">
</font>
</td>
		</tr>
		<tr>
			<td height="18">
<font size="2" face="標楷體">&nbsp;&nbsp; 該機構收發單位代收訖</font><font size="2"> </font>
</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;□附上原掛號收據影印本一件　請查收</font><font size="2"> </font>
</td>
		</tr>
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;□</font><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="標楷體">郵　局</font>
</font>
</td>
		</tr>
		<tr>
			<td height="18">
			<p align="left">
<font face="標楷體" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 年　　月　　日</font></td>
		</tr>
	</table>
	</td>
  </tr>
  </table>
<font face="標楷體" size="3">該回執聯請退回<%=UnitName%>&nbsp;&nbsp;<%=Code%>&nbsp;&nbsp;<%=Address%></font>
<div style="position: absolute; width: 226px; height: 33; z-index: 8; left: 270px; top: 740px" id="layer31">
<img src="..\BarCodeImage\<%=Sys_BillNo_BarCode%>.jpg">
</div>
<div id="L4" style="position:absolute; left:503px;top:737px;width:196px; height:283px">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="105" id="table21" height="263">
      <tr>
        <td width="103" height="114" valign="bottom">
<font face="標楷體">收寄局郵戳</font></td>
      </tr>
      <tr>
        <td width="103" height="150" valign="bottom">
<font face="標楷體">投遞後郵戳</font></td>
      </tr>
    </table>
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
printWindow(true,8,5,5.08,0);
</script>