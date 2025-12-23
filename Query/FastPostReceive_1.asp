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

PBillSN=split(trim(request("PBillSN")),",")
for i=0 to Ubound(PBillSN)

if cint(i)>0 and i mod 2=0 then response.write "<div class=""PageNext"">&nbsp;</div>"
if i mod 2=0 then 
			GetMailAddress="" :Sys_DriverHomeAddress="" : Sys_DriverHomeZip="": ZipName="" : ZipName2=""
			Sys_Driver="" :Billno="" :CarNo="" :Owner=""
			Zip1="" :Zip2="":Zip3=""
			Sys_BillTypeID="" :	Zip11="":Zip21="":Zip31="":MailNumber=""
strBill="select b.Billno,b.CarNo,a.Owner,a.Driver,a.OwnerZip,a.DriverHomeAddress,a.DriverHomeZip,a.OwnerZip,a.OwnerAddress,substr(a.OwnerZip,1,1) as Zip1,substr(a.OwnerZip,2,1) as Zip2,substr(a.OwnerZip,3,1) as Zip3,substr(a.DriverHomeZip,1,1) as Zip11,substr(a.DriverHomeZip,2,1) as Zip21,substr(a.DriverHomeZip,3,1) as Zip31,b.BillTypeID from billbasedcireturn a,Billbase b where a.BillNO=b.BillNo and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i)
set rsBill=conn.execute(strBill)
		if not rsBill.eof then
						strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=Nothing

				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=Nothing


				GetMailAddress=Replace(ZipName&trim(rsBill("OwnerAddress")),ZipName&ZipName,ZipName)
				 Sys_DriverHomeAddress=Replace(ZipName2&trim(rsBill("DriverHomeAddress")),ZipName2&ZipName2,ZipName2)
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

			If sys_City="彰化縣" or sys_City="高雄市" Then
				DelphiASPObj.GenSendStoreBillno BillNo,0,50,160,1

			else
				DelphiASPObj.GenSendStoreBillno BillNo,0,50,160

			end if
          	
          	
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
<table border="2" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="608" id="AutoNumber1" height="325">
  <tr>
    <td width="604" colspan="4" height="20">
    <p align="center"><font face="標楷體">中華郵政掛號郵件收回執</font></td>
  </tr>
  <tr>
    <td width="604" colspan="4" height="15">　　　<font face="標楷體">郵件種類　　　　　　　　　　號碼<%=MailNumber%>　　<font size="2">(由郵局收寄人員填寫)</font></font></td>
  </tr>
  <tr>
    <td width="93" height="71"> <b>&nbsp;收件人姓名<br>&nbsp;地址</b><font face="標楷體">(請寄</font><br><font face="標楷體">&nbsp;件人填寫)</font></td>
    <td width="407" height="71" colspan="2">
    <table border="0" width="100%" id="table3" height="60" cellspacing="0" cellpadding="0">
		<tr>
			<td width="348">
    <font face="標楷體">&nbsp;&nbsp;&nbsp;<%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write DriverHomeZip else response.write OwnerZip%>&nbsp;&nbsp;
    <%=Billno%>&nbsp;&nbsp;&nbsp; <%=(chstr(left(CarNo,4)&left("*************",len(CarNo)-4)))%> 
    </font></td>
			<td rowspan="2">
<font size="4" face="標楷體">小姐 </font>
			</td>
		</tr>
		<tr>
			<td width="348">
    　<font face="標楷體">&nbsp;&nbsp;&nbsp;<%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write funcCheckFont(Sys_Driver,16,1) else response.write funcCheckFont(Owner,16,1)%></font> </td>
		</tr>
		<tr>
			<td width="348">
    　<font face="標楷體"><%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then  response.write funcCheckFont(Sys_DriverHomeAddress,16,1) else response.write funcCheckFont(GetMailAddress,16,1) %></font></td>
			<td>
<font size="4" face="標楷體">先生</font></td>
		</tr>
	</table>
	</td>
    <td width="102" height="289" rowspan="3" valign="bottom">
    <div id="L2" style="position:absolute; left:503px;top:42px;width:183px; height:340">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="105" id="AutoNumber2" height="298">
      <tr>
        <td width="103" height="151" valign="bottom">
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
    <td width="93" height="219" rowspan="2"><b><font size="5" face="標楷體">&nbsp;投&nbsp;遞</font></b><br><br><b><font size="5" face="標楷體">&nbsp;記&nbsp;要</font></b></td>
    <td width="77" height="69">
    <p align="center"><font face="標楷體" size="4">請收件<br>人填寫</font></td>
    <td width="329" height="69">
	<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0">
		<tr>
			<td>
<font face="標楷體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 年　　月　　日收到第　　　 &nbsp;&nbsp;&nbsp; 號 </font>
			</td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 掛號郵件壹件 </font>
			</td>
		</tr>
		<tr>
			<td>　</td>
		</tr>
		<tr>
			<td height="18">
<font face="標楷體">&nbsp;&nbsp; 收件人</font></td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp;&nbsp; 蓋　章</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="標楷體">投遞士戳</font>
</td>
		</tr>
	</table>
	</td>
  </tr>
  <tr>
    <td width="407" height="117" colspan="2">
	<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="0">
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;&nbsp; (供查詢時填寫)</font><font size="2"> </font>
			</td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp;□經查上述郵件已於　　年　　月　　日妥投</font>
</td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 君收訖</font>
</td>
		</tr>
		<tr>
			<td height="18">
<font face="標楷體">&nbsp;&nbsp; 該機構收發單位代收訖</font>
</td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp;□附上原掛號收據影印本一件　請查收</font>
</td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp;□</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="標楷體">郵　局</font>
</td>
		</tr>
		<tr>
			<td height="18">
			<p align="left">
<font face="標楷體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 年　　月　　日</font></td>
		</tr>
	</table>
	</td>
  </tr>
  </table>

    
    
<!--第三區   去除=Code . 不然分局也會顯示交?對的郵遞區號 正確方式是在單位管理的住址前面加上郵遞區號-->

<font face="標楷體" size="5">該回執聯請退回<%=UnitName%><br>&nbsp;&nbsp;<%=Address%></font>

			<div style="position: absolute; width: 217px; height: 33; z-index: 8; left: 4px; top: 395px" id="layer31">
<img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg">
</div>




<p>　</p>
<p>　</p>
<p>　</p>
<p>　</p>
<p></p><p></p><p></p>
<%
	if (i+1 < Ubound(PBillSN)) or (i+1 = Ubound(PBillSN))then 
strBill="select b.Billno,b.CarNo,a.Owner,a.Driver,a.OwnerZip,a.DriverHomeAddress,a.OwnerAddress,a.DriverHomeZip,a.OwnerZip,substr(a.OwnerZip,1,1) as Zip1,substr(a.OwnerZip,2,1) as Zip2,substr(a.OwnerZip,3,1) as Zip3,substr(a.DriverHomeZip,1,1) as Zip11,substr(a.DriverHomeZip,2,1) as Zip21,substr(a.DriverHomeZip,3,1) as Zip31,b.BillTypeID from billbasedcireturn a,Billbase b where a.BillNO=b.BillNo and a.CarNo=b.CarNo and a.ExchangeTypeID='W' and b.SN="&PBillSN(i+1)
			GetMailAddress="" :Sys_DriverHomeAddress="" : Sys_DriverHomeZip="": ZipName="" : ZipName2=""
			Sys_Driver="" :Billno="" :CarNo="" :Owner=""
			Zip1="" :Zip2="":Zip3=""
			Sys_BillTypeID="" :	Zip11="":Zip21="":Zip31="":MailNumber=""
set rsBill=conn.execute(strBill)
		if not rsBill.eof then
		
						strZip="select ZipName from Zip where ZipID='"&trim(rsBill("OwnerZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=Nothing

				strZip="select ZipName from Zip where ZipID='"&trim(rsBill("DriverHomeZip"))&"'"
				set rsZip=conn.execute(strZip)
				if not rsZip.eof then
					ZipName2=trim(rsZip("ZipName"))
				end if
				rsZip.close
				set rsZip=Nothing


				GetMailAddress=Replace(ZipName&trim(rsBill("OwnerAddress")),ZipName&ZipName,ZipName)
				 Sys_DriverHomeAddress=Replace(ZipName2&trim(rsBill("DriverHomeAddress")),ZipName2&ZipName2,ZipName2)
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

			          	If sys_City="彰化縣" or sys_City="高雄市" Then
							DelphiASPObj.GenSendStoreBillno BillNo,0,50,160,1

						else
							DelphiASPObj.GenSendStoreBillno BillNo,0,50,160

						end if
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
<table border="2" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="608" id="table7" height="325">
  <tr>
    <td width="604" colspan="4" height="20">
    <p align="center"><font face="標楷體">中華郵政掛號郵件收回執</font></td>
  </tr>
  <tr>
    <td width="604" colspan="4" height="15">　　　<font face="標楷體">郵件種類　　　　　　　　　　號碼<%=MailNumber%>　　<font size="2">(由郵局收寄人員填寫)</font></font></td>
  </tr>
  <tr>
    <td width="93" height="71"> <b>&nbsp;收件人姓名<br>&nbsp;地址</b><font face="標楷體">(請寄</font><br><font face="標楷體">&nbsp;件人填寫)</font></td>
    <td width="407" height="71" colspan="2">
    <table border="0" width="100%" id="table8" height="53" cellspacing="0" cellpadding="0">
		<tr>
			<td width="336">
    <font face="標楷體">&nbsp;&nbsp;&nbsp;<%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write DriverHomeZip else response.write OwnerZip%>&nbsp;&nbsp;&nbsp;&nbsp;<%=Billno%>&nbsp;&nbsp;&nbsp; <%=(chstr(left(CarNo,4)&left("*************",len(CarNo)-4)))%>  
    </font></td>
			<td rowspan="2">
<font size="4" face="標楷體">小姐 </font>
			</td>
		</tr>
		<tr>
			<td width="336">
    　<font face="標楷體">&nbsp;&nbsp;&nbsp;<%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then response.write funcCheckFont(Sys_Driver,16,1) else response.write funcCheckFont(Owner,16,1)%></font> </td></td>
		</tr>
		<tr>
			<td width="336">
    　<font face="標楷體"><%if Sys_BillTypeID="1"   and trim(Sys_DriverHomeAddress)<>""  then  response.write funcCheckFont(Sys_DriverHomeAddress,16,1) else response.write funcCheckFont(GetMailAddress,16,1) %></font></td>
			<td>
<font size="4" face="標楷體">先生</font></td>
		</tr>
	</table>
	</td>
    <td width="102" height="289" rowspan="3" valign="bottom">
    　</td>
  </tr>
  <tr>
    <td width="93" height="219" rowspan="2"><b><font size="5" face="標楷體">&nbsp;投&nbsp;遞</font></b><br><br><b><font size="5" face="標楷體">&nbsp;記&nbsp;要</font></b></td>
    <td width="77" height="69">
    <p align="center"><font face="標楷體" size="4">請收件<br>人填寫</font></td>
    <td width="329" height="69">
	<table border="0" width="100%" id="table10" cellspacing="0" cellpadding="0">
		<tr>
			<td>
<font face="標楷體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 年　　月　　日收到第　　　 &nbsp;&nbsp;&nbsp; 號 </font>
			</td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;掛號郵件壹件 </font>
			</td>
		</tr>
		<tr>
			<td>　</td>
		</tr>
		<tr>
			<td height="18">
<font face="標楷體">&nbsp;&nbsp; 收件人</font></td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp;&nbsp; 蓋　章</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="標楷體">投遞士戳</font>
</td>
		</tr>
	</table>
	</td>
  </tr>
  <tr>
    <td width="407" height="117" colspan="2">
	<table border="0" width="100%" id="table11" cellspacing="0" cellpadding="0">
		<tr>
			<td>
<font size="2" face="標楷體">&nbsp;&nbsp; (供查詢時填寫)</font><font size="2"> </font>
			</td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp;□經查上述郵件已於　　年　　月　　日妥投</font>
</td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 君收訖</font>
</td>
		</tr>
		<tr>
			<td height="18">
<font face="標楷體">&nbsp;&nbsp; 該機構收發單位代收訖</font>
</td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp;□附上原掛號收據影印本一件　請查收</font>
</td>
		</tr>
		<tr>
			<td>
<font face="標楷體">&nbsp;□</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font face="標楷體">郵　局</font>
</td>
		</tr>
		<tr>
			<td height="18">
			<p align="left">
<font face="標楷體">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 年　　月　　日</font></td>
		</tr>
	</table>
	</td>
  </tr>
  </table>
<!--第三區   去除 =Code . 不然分局也會顯示交?對的郵遞區號 正確方式是在單位管理的住址前面加上郵遞區號-->

<font face="標楷體" size="5">該回執聯請退回<%=UnitName%><br><%'=Code%>&nbsp;&nbsp;<%=Address%></font>
			<div style="position: absolute; width: 219px; height: 33; z-index: 8; left: 1px; top: 981px" id="layer32">
<img src="../BarCodeImage/<%=Sys_BillNo_BarCode%>.jpg">
</div>
<div id="L3" style="position:absolute; left:503px;top:615px;width:183px; height:340px">
    <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="105" id="table12" height="301">
      <tr>
        <td width="103" height="151" valign="bottom">
<font face="標楷體">收寄局郵戳</font></td>
      </tr>
      <tr>
        <td width="103" height="151" valign="bottom">
<font face="標楷體">投遞後郵戳</font></td>
      </tr>
    </table>
    </div>
<%
end if
%>
</div>
<%
end if
next
%>
</body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="../smsx.cab#Version=6,1,432,1">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
printWindow(true,7,10.08,5.08,0);
</script>



