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
'--------------------------------------------------------------------------------------------------------------------

PBillSN=split(trim(request("PBillSN")),",")
for i=0 to Ubound(PBillSN)
if cint(i)>0 then response.write "<div class=""PageNext"">&nbsp;</div>"

			GetMailAddress="" :Sys_DriverHomeAddress="" : Sys_DriverHomeZip="" : ZipName="" : ZipName2=""
			Sys_Driver="" :Billno="" :CarNo="" :Owner=""
			Zip1="" :Zip2="":Zip3=""
			Sys_BillTypeID="" :	Zip11="":Zip21="":Zip31="":MailNumber=""
if sys_City<>"彰化縣" then
strBill="select b.Billno,b.CarNo,a.Owner,a.Driver,a.OwnerZip,a.DriverHomeAddress,a.DriverHomeZip,a.OwnerAddress,substr(a.OwnerZip,1,1) as Zip1,substr(a.OwnerZip,2,1) as Zip2,substr(a.OwnerZip,3,1) as Zip3,substr(a.DriverHomeZip,1,1) as Zip11,substr(a.DriverHomeZip,2,1) as Zip21,substr(a.DriverHomeZip,3,1) as Zip31,b.BillTypeID from billbasedcireturn a,Billbase b where a.BillNO=b.BillNo and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i)
else
strBill="select b.Billno,b.CarNo,a.Owner,a.Driver,a.OwnerZip,a.DriverHomeAddress,a.DriverHomeZip,a.OwnerAddress,substr(a.OwnerZip,1,1) as Zip1,substr(a.OwnerZip,2,1) as Zip2,substr(a.OwnerZip,3,1) as Zip3,substr(a.DriverHomeZip,1,1) as Zip11,substr(a.DriverHomeZip,2,1) as Zip21,substr(a.DriverHomeZip,3,1) as Zip31,b.BillTypeID from billbasedcireturn a,Billbase b,DCILog c where a.BillNO=b.BillNo and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and a.BillNO=c.BillNo and a.CarNo=c.CarNo and c.DciReturnStatusID<>'n' and b.SN="&PBillSN(i)
end if
set rsBill=conn.execute(strBill)
		if not rsBill.eof Then
		
				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=replace(trim(rsZip("ZipName")),"台","臺")
				end if
				rsZip.close
				set rsZip=Nothing

				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=replace(trim(rsZip("ZipName")),"台","臺")
				end if
				rsZip.close
				set rsZip=Nothing
				
				GetMailAddress=Replace(ZipName&replace(trim(rsBill("OwnerAddress")&""),"台","臺"),ZipName&ZipName,ZipName)
				 Sys_DriverHomeAddress=Replace(ZipName2&replace(trim(rsBill("DriverHomeAddress")&""),"台","臺"),ZipName2&ZipName2,ZipName2)
                 Sys_DriverHomeZip=trim(rsBill("DriverHomeZip"))
			     Sys_Driver=trim(rsBill("Driver"))

			Billno=trim(rsBill("Billno"))
			CarNo=trim(rsBill("CarNo"))
			Owner=trim(rsBill("Owner"))

			Zip1=trim(rsBill("Zip1"))
			Zip2=trim(rsBill("Zip2"))
			Zip3=trim(rsBill("Zip3"))

			Sys_BillTypeID=trim(rsBill("BillTypeID"))

			Zip11=trim(rsBill("Zip11"))
			Zip21=trim(rsBill("Zip21"))
			Zip31=trim(rsBill("Zip31"))
			
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


%>
<div id="R1" style="position:relative;">
<table border="2" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="608" id="AutoNumber1" height="19">
  <tr>
    <td width="604" colspan="4" height="20">
    <p align="center"><font face="標楷體">中華郵政掛號郵件收回執</font></td>
  </tr>
  <tr>
    <td width="604" colspan="4" height="15">　　　<font face="標楷體">郵件種類　　　　　　　　　　號碼　　　　　　　　 　<font size="2">(由郵局收寄人員填寫)</font></font></td>
  </tr>
  <tr>
    <td width="93" height="71"> <b>&nbsp;收件人姓名<br>&nbsp;地址</b><font face="標楷體">(請寄</font><br><font face="標楷體">&nbsp;件人填寫)</font></td>
    <td width="407" height="71" colspan="2">
    　</td>
    <td width="102" height="123" rowspan="3">
    <div id="L2" style="position:absolute; left:503;top:42;width:2121; height:340">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="105" id="AutoNumber2" height="320">
      <tr>
        <td width="105" height="151">　</td>
      </tr>
      <tr>
        <td width="105" height="168">　</td>
      </tr>
    </table>
    </div>
    </td>
  </tr>
  <tr>
    <td width="93" height="249" rowspan="2"><b><font size="5" face="標楷體">&nbsp;投&nbsp;遞</font></b><br><br><b><font size="5" face="標楷體">&nbsp;記&nbsp;要</font></b></td>
    <td width="77" height="69">
    <p align="center"><font face="標楷體" size="4">請收件<br>人填寫</font></td>
    <td width="329" height="69">　</td>
  </tr>
  <tr>
    <td width="407" height="103" colspan="2"></td>
  </tr>
  </table>

<div id="L3" style="position:absolute; left:226;top:119;width:334; height:16">
<font face="標楷體">年　　月　　日收到第　　　 &nbsp;&nbsp;&nbsp; 號 </font>
</div>
<div id="L4" style="position:absolute; left:194;top:138;width:334; height:16">
<font face="標楷體">掛號郵件壹件 </font>
</div>
<div id="L5" style="position:absolute; left:211;top:168;width:52; height:16">
<font face="標楷體">收件人 </font>
</div>
<div id="L6" style="position:absolute; left:211;top:189;width:52; height:16">
<font face="標楷體">蓋　章</font>
</div>
<div id="L7" style="position:absolute; left:261;top:165;width:76; height:125">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="74" id="AutoNumber3" height="40">
  <tr>
    <td width="74" height="40">　</td>
  </tr>
</table>
　</div>
<div id="L8" style="position:absolute; left:400;top:185;width:64; height:32">
<font face="標楷體">投遞士戳</font>
</div>
<div id="L9" style="position:absolute; left:465;top:184;width:7; height:78">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="27" id="AutoNumber4" height="21">
  <tr>
    <td width="52" height="21">　</td>
  </tr>
</table>
　</div>
<!--第三區-->
<div id="L11" style="position:absolute; left:122;top:220;width:223; height:14">
<font size="2" face="標楷體">(供查詢時填寫)</font><font size="2"> </font>
</div>
<div id="L12" style="position:absolute; left:105;top:235;width:406; height:16">
<font face="標楷體">□經查上述郵件已於　　年　　月　　日妥投</font>
</div>
<div id="L13" style="position:absolute; left:233;top:255;width:269; height:16">
<font face="標楷體">君收訖</font>
</div>
<div id="L14" style="position:absolute; left:121;top:277;width:269; height:16">
<font face="標楷體">該機構收發單位代收訖</font>
</div>
<div id="L15" style="position:absolute; left:104;top:298;width:363; height:16">
<font face="標楷體">□附上原掛號收據影印本一件　請查收</font>
</div>
<div id="L16" style="position:absolute; left:104;top:320;width:178; height:16">
<font face="標楷體">□</font>
</div>
<div id="L17" style="position:absolute; left:413;top:321;width:92; height:16">
<font face="標楷體">郵　局</font>
</div>
<div id="L18" style="position:absolute; left:349;top:343;width:113; height:22">
<font face="標楷體">年　　月　　日</font></div>
<div id="L19" style="position:absolute; left:462;top:324;width:35; height:109">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="35" id="AutoNumber5" height="29">
  <tr>
    <td width="50" height="29">　</td>
  </tr>
</table>
</div>

<div id="L20" style="position:absolute; left:3;top:365;width:188; height:16">
<font size="2" face="標楷體"></font><font size="2"> </font>
</div>

<div id="L21" style="position:absolute; left:104;top:50;width:15; height:45">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FF0000" width="11" id="AutoNumber6" height="18" bordercolorlight="#FF0000" bordercolordark="#FF0000">
  <tr>
    <td width="33" height="18"></td>
  </tr>
</table>
</div>
<div id="L22" style="position:absolute; left:120;top:50;width:18; height:60">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FF0000" width="11" id="AutoNumber6" height="18" bordercolorlight="#FF0000" bordercolordark="#FF0000">
  <tr>
    <td width="33" height="18"></td>
  </tr>
</table>
</div>
<div id="L23" style="position:absolute; left:135;top:50;width:17; height:93">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FF0000" width="11" id="AutoNumber6" height="18" bordercolorlight="#FF0000" bordercolordark="#FF0000">
  <tr>
    <td width="33" height="18"></td>
  </tr>
</table>
</div>
<div id="L24" style="position:absolute; left:149;top:52;width:16; height:60">
<font color="#FF0000">－ </font>
</div>
<div id="L25" style="position:absolute; left:183;top:50;width:15; height:60">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FF0000" width="11" id="AutoNumber6" height="18" bordercolorlight="#FF0000" bordercolordark="#FF0000">
  <tr>
    <td width="100" height="18"></td>
  </tr>
</table>
</div>
<div id="L26" style="position:absolute; left:167;top:50;width:15; height:60">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#FF0000" width="11" id="AutoNumber6" height="18" bordercolorlight="#FF0000" bordercolordark="#FF0000">
  <tr>
    <td width="100" height="18"></td>
  </tr>
</table>
</div>

<div id="L27" style="position:absolute; left:450;top:59;width:43; height:26">
<font size="4" face="標楷體">小姐 </font>
</div>
<div id="L28" style="position:absolute; left:450;top:77;width:95; height:18">
<font size="4" face="標楷體">先生 </font>
</div>

<div id="L29" style="position:absolute; left:517;top:172;width:199; height:18">
<font face="標楷體">收寄局郵戳</font><font size="4" face="標楷體"> </font>
</div>
<div id="L30" style="position:absolute; left:517;top:334;width:217; height:18">
<font face="標楷體">投遞後郵戳</font><font size="4" face="標楷體"> </font>
</div>

<div id="L30" style="position:absolute; left:94;top:380;width:515; height:24">
<font face="標楷體" size="5">該回執聯請退回<%=UnitName%> </font>
</div>
<div id="L30" style="position:absolute; left:95;top:406;width:554; height:24">
<font face="標楷體" size="5"><%=Code%>&nbsp;&nbsp;<%=Address%> </font>
</div>

<div id="L30" style="position:absolute; left:211;top:48;width:101; height:24">
<font face="標楷體"><%=Billno%> </font>
</div>
<div id="L30" style="position:absolute; left:350;top:47;width:95; height:16">
<font face="標楷體"><%
			if trim(CarNo)<>"" and not isnull(CarNo) then
				response.write left(CarNo,4)
				response.write left("*************",len(CarNo)-4)
			end if 
%> </font>
</div>
<div id="L30" style="position:absolute; left:310;top:22;width:147; height:16">
<font face="標楷體"><%=MailNumber%> </font>
</div>

<div id="L30" style="position:absolute; left:233;top:68;width:186; height:22">
<font face="標楷體" size="4"><%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write funcCheckFont(Sys_Driver,16,1) else response.write funcCheckFont(Owner,16,1)%></font>
</div>
<div id="L30" style="position:absolute; left:103;top:92;width:399; height:16">
<font face="標楷體" size="2"><%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then  response.write funcCheckFont(Sys_DriverHomeAddress,16,1) else response.write funcCheckFont(GetMailAddress,16,1) %></font>
</div>

<div id="L30" style="position:absolute; left:105;top:52;width:51; height:32">
<font face="標楷體"><%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>""  then response.write Zip11 else response.write Zip1 %>&nbsp;<%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>""   then   response.write Zip21 else response.write Zip2 %>&nbsp;<%if Sys_BillTypeID="1"  and trim(Sys_DriverHomeAddress)<>""  then response.write Zip31 else response.write Zip3  %></font>
</div>

</div>
<%next%>
</body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="../smsx.cab#Version=6,1,432,1">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
printWindow(true,7,10.08,5.08,0);
</script>