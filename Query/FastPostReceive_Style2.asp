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

Server.ScriptTimeout = 60000 
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
'sys_City="台中縣"
'UnitID="04A7"
set rsCity=nothing
if sys_City="台中縣" then 
	BigUnitName=sys_City&"警察局"
else
	BigUnitName=sys_City&"政府警察局"
end if

'--------------------------------------------------------------------------------------------------------------------
'登入者、單位地址
	strUNit="select UnitName,Address,Tel,UnitLevelID,UnitID from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
	set rsUNit=conn.execute(strUNit)
		if not rsUNit.eof then
			UnitName=trim(rsUNit("UnitName"))
			UnitID=trim(rsUNit("UnitID"))
			Address=trim(rsUNit("Address"))
			UnitLevelID=trim(rsUNit("UnitLevelID"))
			Tel=trim(rsUNit("Tel"))			
		end If
	rsUNit.close
	set rsUNit=nothing	
	
	if trim(UnitLevelID)="3" then 
	strUNit="select UnitName from UnitInfo where UnitTypeID='"&Session("Unit_ID")&"'"
	set rsUNit=conn.execute(strUNit)
		if not rsUNit.eof then
			UnitName=UnitName&trim(rsUNit("UnitName"))
		end If
	rsUNit.close
	set rsUNit=nothing
	
	end If

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
'-----------------------------------------------------------------------------------------------------
PBillSN=""

if UCase(request("Sys_BatchNumber"))<>"" then
	tmp_BatchNumber=split(UCase(request("Sys_BatchNumber")),",")
	for i=0 to Ubound(tmp_BatchNumber)
		if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
		if i=0 then
			Sys_BatchNumber=trim(Sys_BatchNumber)&UCase(trim(tmp_BatchNumber(i)))
		else
			Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&UCase(trim(tmp_BatchNumber(i)))
		end if
		if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(UCase(Sys_BatchNumber))&"'"
	next
	strwhere=" and a.BatchNumber in('"&trim(Sys_BatchNumber)&"')"
end if

if trim(request("Sys_BillNo1"))<>"" and trim(request("Sys_BillNo2"))<>"" then
	strwhere=strwhere&" and a.BillNo between '"&trim(UCase(request("Sys_BillNo1")))&"' and '"&trim(UCase(request("Sys_BillNo2")))&"'"
elseif trim(request("Sys_BillNo1"))<>"" then
	strwhere=strwhere&" and a.BillNo between '"&trim(UCase(request("Sys_BillNo1")))&"' and '"&trim(UCase(request("Sys_BillNo1")))&"'"
elseif trim(request("Sys_BillNo2"))<>"" then
	strwhere=strwhere&" and a.BillNo between '"&trim(UCase(request("Sys_BillNo2")))&"' and '"&trim(UCase(request("Sys_BillNo2")))&"'"
end if

If sys_City="基隆市" then
	KindType="('1','3','9','a','j','A','H','K','L','T')"
elseIf sys_City="台中市" and session("User_ID")=5751 then
	KindType="('1','3','9','a','j','A','H','K','L','T')"
else
	KindType="('1','3','9','a','j','A','H','K','T','n')"
End if
If sys_City="台中市" or sys_City="南投縣" then KindType=KindType&" and a.DciReturnStatusID<>'n'"

tempSQL="where (a.BillTypeID='2' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' and a.DciErrorCarData Not in "&KindType&" "&strwhere&" and NVL(f.EquiPmentID,1)<>-1) or (a.BillTypeID='1' and a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=h.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+) and a.BillSN=f.SN and a.BillNo=f.BillNo and f.RecordStateId <> -1 and d.DCIRETURNSTATUS='1' and a.ExchangeTypeID<>'E' "&strwhere&" and NVL(f.EquiPmentID,1)<>-1)"
'		end if
tempSQL=tempSQL&" and f.EquiPmentID<>-1"

'if trim(request("PBillSN"))="" then '與dci上下查詢不同
chk_MailNumKind=0
if Instr(request("Sys_BatchNumber"),"N")>0 then
	strSQL="select distinct a.BillSN from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h "&tempSQL
	strSQL="Select BillSN from BillMailHistory where BillSN in("&strSQL&") order by UserMarkDate"
	chk_MailNumKind=1
else
	strSQL="select distinct a.BillSN,a.RecordMemberID,f.RecordDate from DCILog a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,BillBase f,(select * from DciReturnStatus where DciActionID='WE') g,(select * from DciReturnStatus where DciActionID='WE') h "&tempSQL&" order by f.RecordDate"
end if

set rsSn=conn.execute(strSQL)
While not rsSn.eof
	If Not ifnull(PbillSN) Then PbillSN=PBillSN&","
	PbillSN=PBillSN&rsSN("BillSN")
	rsSN.movenext	
Wend
rsSN.close
if sys_City="南投縣" then
	PBillSN=split(trim(PBillSN),",")
else
	PBillSN=split(trim(request("PBillSN")),",")
end if

if UnitID="05GF" or UnitID="05BA" or UnitID="05B0" then 
	strbatchnumber=trim(request("Sys_batchnumber"))
end if
iTotal=UBound(PBillSN)
for i=0 to Ubound(PBillSN)
			GetMailAddress="" :Sys_DriverHomeAddress="" : Sys_DriverHomeZip=""
			Sys_Driver="" :Billno="" :CarNo="" :Owner=""
			Zip1="" :Zip2="":Zip3="" :OwnerZip="" :Sys_DriverHomeZip=""
			Sys_BillTypeID="" :	Zip11="":Zip21="":Zip31="":MailNumber=""
			Sys_DriverZipName="" : ZipName="": ZipName2=""
if cint(i)>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
if sys_City="南投縣" Then
	strBill="select c.StoreAndSendMailNumber,b.recordmemberid,b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress from billbasedcireturn a,Billbase b,billmailhistory c where a.BillNO=b.BillNo and a.BillNO=c.BillNo(+) and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i)
else
	strBill="select b.recordmemberid,b.Billno,b.CarNo,a.Owner,a.DriverHomeZip,a.Driver,b.BillTypeID,a.DriverHomeAddress,a.OwnerZip,a.OwnerAddress from billbasedcireturn a,Billbase b where a.BillNO=b.BillNo and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i)
End If

set rsBill=conn.execute(strBill)
icount=0
		if not rsBill.eof then
		
	    	if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣"  then
				ZipName=""
			else
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=replace(trim(rsZip("ZipName")),"台","臺")
				end if
				rsZip.close
				set rsZip=nothing
			end If
			
	    	if sys_City="基隆市" or sys_City="金門縣" or sys_City="澎湖縣"  then
				ZipName2=""
			else
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=replace(trim(rsZip("ZipName")),"台","臺")
				end if
				rsZip.close
				set rsZip=nothing
			end if
			if not isnull(rsBill("OwnerAddress") ) then
				GetMailAddress=ZipName&replace(replace(trim(rsBill("OwnerAddress")),"台","臺"),ZipName,"")
			end if
			Billno=trim(rsBill("Billno"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))
			Sys_BillTypeID=trim(rsBill("BillTypeID"))
			OwnerZip=trim(rsBill("OwnerZip"))
			if sys_City="南投縣" Then
			StoreAndSendMailNumber=trim(rsBill("StoreAndSendMailNumber"))
			End if
			DriverHomeZip=trim(rsBill("DriverHomeZip"))
			recordmemberid=trim(rsBill("recordmemberid"))

				 Sys_DriverHomeAddress=ZipName2&trim(rsBill("DriverHomeAddress"))
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



%>

<div id="R1" style="position:relative;">

	<!-- MSTableType="layout" -->
	<%If sys_City="南投縣" Then %>
	<div style="position: absolute; width: 139px; height: 19px; z-index: 1; left: 500px; top: 14px" id="layer1">
			<font face="標楷體">第<%=i+1%>頁，共<%=itotal+1%>頁</font></div>
	<%End If %>
	<tr>
		<td colspan="2" valign="top">
		<div style="position: absolute; width: 139px; height: 19px; z-index: 1; left: 36px; top: 34px" id="layer1">
			<font face="標楷體">中　華　郵　政</font></div>
		<div style="position: absolute; width: 202px; height: 19px; z-index: 2; left: 2px; top: 56px" id="layer2">
			<font face="標楷體">大宗郵資已付掛號函件收據</font></div>
		<div style="position: absolute; width: 193; height: 32; z-index: 3; left: 1; top: 78" id="layer3">
			<font face="標楷體">第　　　　           　&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 號</font></div>
　</td>
		<td rowspan="5" valign="top" style="border-style: solid; border-width: 1px" width="420">
		
		
<div style="position: absolute; width: 314px; height: 326px; z-index: 9; left: 206px; top: 31px" id="layer14">
	<table border="1" width="100%" id="table1" cellspacing="0" height="100%" cellpadding="0">
		<tr>
			<td height="58">
			<div style="position: absolute; width: 123px; height: 20px; z-index: 1; left: 86px; top: 10px" id="layer15">
				<font face="標楷體">中　華　郵　政</font></div>
			<div style="position: absolute; width: 133px; height: 22px; z-index: 2; left: 82px; top: 31px" id="layer16">
				<font face="標楷體">掛號郵件收件回執</font></div>
			<div style="position: absolute; width: 297px; height: 20px; z-index: 5; left: 12px; top: 67px" id="layer29">
				<font face="標楷體">姓名：<%
				if Sys_BillTypeID="1" and trim(Sys_DriverHomeAddress)<>""  then
					response.write funcCheckFont(Sys_Driver,16,1)
				else
					response.write funcCheckFont(Owner,16,1)
				end if%></font></div>
　</td>
		</tr>
		<tr>
			<td height="103">
			<div style="position: absolute; width: 298px; height: 35px; z-index: 6; left: 11px; top: 89px" id="layer30">
				<font face="標楷體" size="2">地址：<%
				if instr(request("Sys_batchnumber"),"N")>0 and  sys_City="南投縣" then
				  Addr=""
					Addr=DriverHomeZip &"&nbsp;"& funcCheckFont(Sys_DriverHomeAddress,16,1)
					if Trim(Sys_DriverHomeAddress)="" then
					    Addr=OwnerZip &"&nbsp;"& funcCheckFont(GetMailAddress,16,1)
					end if
					response.write Addr

				else
					if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>""  then 
						response.write DriverHomeZip 
					else 
						response.write OwnerZip  
					end if
					%>&nbsp;<%
					if Sys_BillTypeID="1" and trim(Sys_DriverHomeAddress)<>""  then  
						response.write funcCheckFont(Sys_DriverHomeAddress,16,1)
					else 
						response.write funcCheckFont(GetMailAddress,16,1)
					end if
				end if
				%></font></div>
			<div style="position: absolute; width: 278px; height: 24px; z-index: 7; left: 10px; top: 133px" id="layer31">
<font face="標楷體">車號：<%
			if trim(CarNo)<>"" and not isnull(CarNo) then
                             if len(CarNo)>=4 then
				response.write left(CarNo,4)
				response.write left("*************",len(CarNo)-4)
			      else
				response.write (CarNo)
 			     end if
			

			end if 
%>&nbsp;&nbsp;</font>
</div>

			<div style="position: absolute; width: 111; height: 33; z-index: 8; left: 151; top: 110" id="layer31">
<img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg">

</div>

　</td>
		</tr>
		<tr>
			<td>
			<div style="position: absolute; width: 17px; height: 76px; z-index: 3; left: 9px; top: 168px" id="layer17">
				<font face="標楷體">退<br>件<br>原<br>因<br><font size="2">︵<br>請<br>勾<br>選<br>
				︶</font></font><font size="2"> </font>
				</div>
			<div style="position: absolute; width: 17px; height: 123px; z-index: 4; left: 35px; top: 168px" id="layer18">
				<p align="center">□<br>1.<br><font face="標楷體">查<br>無<br>此<br>人</font></div>
			<div style="position: absolute; width: 17px; height: 123px; z-index: 4; left: 56px; top: 168px" id="layer19">
				<p align="center">□<br>2.<br><font face="標楷體">遷<br>移<br>不<br>明</font></div>
			<div style="position: absolute; width: 17px; height: 123px; z-index: 4; left: 77px; top: 168px" id="layer20">
				<p align="center">□<br>3.<br><font face="標楷體">地<br>址<br>欠<br>詳</font></div>
			<div style="position: absolute; width: 17px; height: 123px; z-index: 4; left: 97px; top: 168px" id="layer21">
				<p align="center">□<br>4.<br><font face="標楷體">無<br>此<br>地<br>址</font></div>
			<div style="position: absolute; width: 17px; height: 123px; z-index: 4; left: 117px; top: 168px" id="layer22">
				<p align="center">□<br>5.<br><font face="標楷體">招<br>領<br>逾<br>期</font></div>
			<div style="position: absolute; width: 17px; height: 123px; z-index: 4; left: 137px; top: 168px" id="layer23">
				<p align="center">□<br>6.<br><font face="標楷體">關<br>閉<br>歇<br>業</font></div>
			<div style="position: absolute; width: 17px; height: 123px; z-index: 4; left: 157px; top: 168px" id="layer24">
				<p align="center">□<br>7.<br><font face="標楷體">收<br>件<br>人<br>拒<br>收</font></div>
			<div style="position: absolute; width: 17px; height: 71px; z-index: 3; left: 195px; top: 168px" id="layer25">
				                          <font face="標楷體">回<br>執<br>情<br>況</font>
				</div>
			<div style="position: absolute; width: 17px; height: 123px; z-index: 4; left: 216px; top: 168px" id="layer26">
				<p align="center">□<br>8.<br><font face="標楷體">本<br>人<br>簽<br>收 </font></div>
			<div style="position: absolute; width: 17px; height: 123px; z-index: 4; left: 237px; top: 168px" id="layer27">
				<p align="center">□<br>9.<br><font face="標楷體">非<br>本<br>人<br>簽<br>收</font></div>
			<div style="position: absolute; width: 17px; height: 123px; z-index: 4; left: 257px; top: 168px" id="layer28">
				<p align="center">□<br>10.<br><font face="標楷體">管<br>理<br>人<br>員<br>簽<br>收</font></div>
　</td>
		</tr>
	</table>
</div>
　</td>
		<td height="67">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<div style="position: absolute; width: 198px; height: 15px; z-index: 4; left: 3px; top: 170px" id="layer4">
			<font size="2" face="標楷體">此件已於　年　月　日當面驗明</font></div>
　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<div style="position: absolute; width: 198px; height: 15px; z-index: 4; left: 3px; top: 186px" id="layer5">
			<font face="標楷體" size="2">無訛收妥(請收件人簽名或蓋章)</font></div>
		<div style="position: absolute; width: 198px; height: 15px; z-index: 4; left: 3px; top: 202px" id="layer6">
			<font face="標楷體" size="2">郵資已付掛號　收件人：</font></div>
		<div style="position: absolute; width: 71px; height: 15px; z-index: 4; left: 102px; top: 219px" id="layer7">
			<font size="2" face="標楷體">代收人：</font></div>
		<div style="position: absolute; width: 69px; height: 15px; z-index: 4; left: 102px; top: 234px" id="layer8">
			<font size="2" face="標楷體">關　係：</font></div>
		<div style="position: absolute; width: 99px; height: 15px; z-index: 4; left: 101px; top: 248px" id="layer9">
			<font face="標楷體" size="2">投遞士蓋章：</font></div>
　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td valign="top" width="110">
		<div style="position: absolute; width: 76px; height: 9px; z-index: 6; left: 13px; top: 276px" id="layer11">
			<font face="標楷體"></font></div>
　</td>
		<td valign="top" width="95">
		<div style="position: absolute; width: 82px; height: 20px; z-index: 5; left: 3px; top: 276px" id="layer10">
			<font face="標楷體">原寄局局名</font></div>
		<div style="position: absolute; width: 90px; height: 23px; z-index: 7; left: 103px; top: 275px" id="layer12">
			<font face="標楷體">投遞局郵戳</font></div>
　</td>
		<td height="66">　</td>
	</tr>
	<tr>
		<td colspan="2" valign="top">
		<div style="position: absolute; width: 159px; height: 25px; z-index: 8; left: 6px; top: 320px" id="layer13">
			<font face="標楷體">落地號碼：</font></div>
		<p>　</td>
		<td height="67" width="102">　</td>
	</tr>

			<div style="position: absolute; width: 82px; height: 16px; z-index: 1; left: 521px; top: 199px" id="layer37">
				<font face="標楷體">投遞後郵戳</font></div>
<div style="position: absolute; width: 69px; height: 16px; z-index: 1; left: 528px; top: 122px" id="layer35">
	<font face="標楷體">投遞士戳</font></div>
<div style="position: absolute; width: 203px; height: 323px; z-index: -1; left: -1px; top: 31px" id="layer32">
	<table cellpadding="0" cellspacing="0" width="209" height="100%" border="1">
		<!-- MSTableType="layout" -->
		<tr>
			<td height="73" width="205" colspan="2">　</td>
		</tr>
		<tr>
			<td height="60" width="205" colspan="2">
　</td>
		</tr>
		<tr>
			<td height="103" width="255" colspan="2">　</td>
		</tr>
		<tr>
			<td height="47" width="95">　</td>
			<td height="47" width="108">　</td>
		</tr>
		<tr>
			<td height="41" width="200" colspan="2">　</td>
		</tr>
	</table>
</div>
<div id="L30" style="position:absolute; left:9;top:363;width:571; height:24">
<font face="標楷體" size="4">該回執聯請退回
<%
if sys_City="台中縣" then 
	sDouble="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>雙&nbsp;掛&nbsp;號</b>"
else
	sDouble=""
end if

if trim(UnitLevelID)="1" then 
	response.write replace(BigUnitName,"台","臺")&"交通隊" & sDouble 
else 
	response.write replace(BigUnitName,"台","臺") & UnitName & sDouble 
end if
%>

</font>
    <div style="position: absolute; width: 139px; height: 25; z-index: 1; left: 18px; top: -284px" id="layer38">
<font size="4" face="標楷體">
<% 
'------------------------------------------------------------------------------------------------------------------------------------------------------------
      UnitNum=""
       if Ubound(PBillSN)>20 then 
          Cnt="18"
	   else
          Cnt="16"
	   end if
       if trim(UnitID)="05A7"   and sys_City="南投縣" and instr(request("Sys_batchnumber"),"N")=0 then   '交通隊
		   UnitNum="540000"&"18"
           Sys_mailnumber=GetSN(UnitNum,"NTSMAILNUMBER")
		   response.write Sys_mailnumber
       elseif UnitID="05GF"   and sys_City="南投縣" and instr(request("Sys_batchnumber"),"N")=0 then      '集集
		   UnitNum="540009"&"18"
           Sys_mailnumber=GetSN(UnitNum,"NTSMAILNUMBERGG")
		   response.write Sys_mailnumber
       elseif UnitID="05BA"   and sys_City="南投縣" and instr(request("Sys_batchnumber"),"N")=0 then      '南投分局
		   UnitNum="54000518"
           Sys_mailnumber=GetSN(UnitNum,"ntsubunitmailnumber")
		   response.write Sys_mailnumber
       elseif UnitID="05CB"   and sys_City="南投縣" and instr(request("Sys_batchnumber"),"N")=0 then      '草屯
		   UnitNum="54001018"
           Sys_mailnumber=GetSN(UnitNum,"ntsubunitmailnumber05CB")
		   response.write Sys_mailnumber
	   elseif UnitID="05FG"  and sys_City="南投縣" and instr(request("Sys_batchnumber"),"N")=0 then
	   '竹山
		   UnitNum="540022"&Cnt
           Sys_mailnumber=GetSN(UnitNum,"NTSMAILNUMBERDouSam")
		   response.write Sys_mailnumber
	   elseif sys_City="南投縣" And UnitID="05FG" Then
	    '竹山
   		   UnitNum="540022"&Cnt
           StoreAndSendMailNumber=GetSNN(UnitNum,"NTSMAILNUMBERDouSam")
			response.write StoreAndSendMailNumber
	   elseif sys_City="南投縣" And UnitID<>"05FG" Then
			response.write StoreAndSendMailNumber
	   elseif UnitID="04A7" and sys_City="台中縣" then      '台中縣交通隊
		   UnitNum="42045036"
           Sys_mailnumber=GetSN(UnitNum,"TCMailNumber")
		   response.write Sys_mailnumber
	   end if
'------------------------------------------------------------------------------------------------------------------------------------------------------------
%>
</font>
    </div>

</div>
<%if UnitID="05A7" or UnitID="05GF" or UnitID="05FG" or UnitID="04A7" or UnitID="05BA" or UnitID="05CB" then	%>
<div style="position: absolute; width: 177; height: 33; z-index: 8; left: 2; top: 106" id="layer311">
<%	if instr(request("Sys_batchnumber"),"N")>0 and  sys_City="南投縣" then%>
<img src="../BarCodeImage/11.jpg" width=0 height=0></div>
<%else%>
<img src="../BarCodeImage/<%=Sys_mailnumber%>.jpg"></div>
<%end if%>
<%end if%>

<div id="L30" style="position:absolute; left:111px;top:393px;width:438px; height:24">
<font face="標楷體" size="4"><%if trim(UnitLevelID)="1" then response.write Code%>&nbsp;<%=Address%>&nbsp;<%=Tel%>
<%if sys_City="南投縣" then%>
&nbsp;&nbsp;<%		iTemp=right(year(now())-1911,3) &"/"&_
		right("0"&month(now()),2) &"/"&_
		right("0"&day(now()),2)
'response.write iTemp
%>
<%end if%>
<!-- smith add batch number -->
<%'if UnitID="05GF"  or UnitID="05BA" or UnitID="05B0"  then 
If sys_City="南投縣" Then
'	response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size='2'>" & sBatchNumber & "</font>"
end if
if UnitID="04A7" and sys_City="台中縣" then
    response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size='2'>04A7</font>"
end if
%>
<!---------------------------->
</font><font size="4">
</font>
</div>

<div id="L31" style="position:absolute; left:111px;top:10px;width:438px; height:24">
<font face="標楷體" size="4">
<%if sys_City="南投縣" then%>
&nbsp;&nbsp;<%		iTemp=right(year(now())-1911,3) &"/"&_
		right("0"&month(now()),2) &"/"&_
		right("0"&day(now()),2)
response.write iTemp
%>
<%end if%>

</font><font size="4">
</font>
</div>

<div style="position: absolute; width: 88px; height: 326px; z-index: 11; left: 518px; top: 31px" id="layer33">
	<table border="1" width="100%" height="100%" cellspacing="0" cellpadding="0">
		<tr>
			<td height="30" width="78">
			<div style="position: absolute; width: 81px; height: 16px; z-index: 1; left: 4px; top: 8px" id="layer34">
				<font face="標楷體">收件人蓋章</font></div>
　</td>
		</tr>
		<tr>
			<td height="57" width="78">　</td>
		</tr>
		<tr>
			<td height="25" width="78">　</td>
		</tr>
		<tr>
			<td height="47" width="78">
			<div style="position: absolute; width: 73px; height: 13px; z-index: 2; left: 14px; top: 146px" id="layer36">
				<font face="標楷體" size="2">年　月　日</font></div>
　</td>
		</tr>
		<tr>
			<td height="30" width="78">　</td>
		</tr>
		<tr>
			<td height="132" width="78">　</td>
		</tr>
	</table>
</div>

<div style="position: absolute; width: 148px; height: 23px; z-index: 12; left: 470px; top: 364px" id="layer39">
<% if trim(UnitID)="05A7" then   '交通隊%>

	<font face="標楷體">處理人員：
	<%
		strUNit="select loginid from memberdata where memberid="&recordmemberid
	set rsUNit=conn.execute(strUNit)
		if not rsUNit.eof then
			response.write trim(rsUNit("loginid"))
		end If
	rsUNit.close
	set rsUNit=nothing	
	
	%>
	</font>
<%end if%>	
	</div>
</div>


<%next%>

</body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,5,439,50">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
<%if UnitID="05FG" then%>
	printWindow(true,7,0,0,0);
<%else%>
	printWindow(true,6,0,0,0);
<%end if%>
</script>
<%
function GetSN(UnitNum,SNname)



'---------------------------------------------------------------------------------------------------------------------------------------------------------
        strSQL="select BillSN,MailNumber from BillMailHistory where BillSN="&PBillSN(i)

		set rscnt=conn.execute(strSQL)

		if Not rscnt.eof then
			if ifnull(rscnt("MailNumber")) then
	          strUNit="select "&SNname&".nextval as SN from Dual"
				set rsUNit=conn.execute(strUNit)
				if not rsUNit.eof then
					mailnumber=trim(rsUNit("SN"))
				end If
				rsUNit.close
				set rsUNit=nothing	

				for j=1 to 6-len(trim(mailnumber))
					mailnumber="0" & mailnumber
				next 
		
				Sys_mailnumber=mailnumber&UnitNum

				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&",MailNumber="&Sys_mailnumber&" where BillSN="&PBillSN(i)
				conn.execute(strSQL)
				    DelphiASPObj.GenSendStoreBillno Sys_mailnumber,128,58,200
					  GetSN=Sys_mailnumber	
            else
			   if UnitID="04A7" or UnitID="05A7" or UnitID="05GF" or UnitID="05FG" or UnitID="05BA" or UnitID="05CB" then
			       mailnumber="" 
				   for j=1 to 14-len(trim(rscnt("MailNumber")))
			     		mailnumber="0" & mailnumber 
				   next 		   
				   mailnumber=mailnumber& rscnt("MailNumber")
                   DelphiASPObj.GenSendStoreBillno mailnumber,128,58,200
				   GetSN=mailnumber
			   else
					DelphiASPObj.GenSendStoreBillno rscnt("MailNumber"),128,58,200
				   GetSN=rscnt("MailNumber")			   
			   end if
			end if
		end If
		
		rscnt.close
'---------------------------------------------------------------------------------------------------------------------------------------------------------
end Function

function GetSNN(UnitNum,SNname)



'---------------------------------------------------------------------------------------------------------------------------------------------------------
        strSQL="select BillSN,MailNumber,StoreAndSendMailNumber from BillMailHistory where BillSN="&PBillSN(i)

		set rscnt=conn.execute(strSQL)

		if Not rscnt.eof then
			if ifnull(rscnt("MailNumber")) then
	          strUNit="select "&SNname&".nextval as SN from Dual"
				set rsUNit=conn.execute(strUNit)
				if not rsUNit.eof then
					mailnumber=trim(rsUNit("SN"))
				end If
				rsUNit.close
				set rsUNit=nothing	

				for j=1 to 6-len(trim(mailnumber))
					mailnumber="0" & mailnumber
				next 
		
				Sys_mailnumber=mailnumber&UnitNum

				strSQL="Update BillMailHistory set MailDate="&funGetDate(date,0)&",StoreAndSendMailNumber="&Sys_mailnumber&" where BillSN="&PBillSN(i)
				conn.execute(strSQL)
				    DelphiASPObj.GenSendStoreBillno Sys_mailnumber,128,58,200
					  GetSNN=Sys_mailnumber	
            elseif not ifnull(rscnt("StoreAndSendMailNumber")) then            
			   if UnitID="04A7" or UnitID="05A7" or UnitID="05GF" or UnitID="05FG"  or UnitID="05BA" then
			       mailnumber="" 
				   for j=1 to 14-len(trim(rscnt("StoreAndSendMailNumber")))
			     		mailnumber="0" & mailnumber 
				   next 		   
				   mailnumber=mailnumber& rscnt("StoreAndSendMailNumber")
                   DelphiASPObj.GenSendStoreBillno mailnumber,128,58,200
				   GetSNN=mailnumber
			   else
					DelphiASPObj.GenSendStoreBillno rscnt("StoreAndSendMailNumber"),128,58,200
				   GetSNN=rscnt("StoreAndSendMailNumber")			   
			   end if
			end if
		end If
		
		rscnt.close
'---------------------------------------------------------------------------------------------------------------------------------------------------------
end Function

%>
</p>
<p>　</p>