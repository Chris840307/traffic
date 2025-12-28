<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strSql="select a.SN as BillSN,a.BillNo,b.OpenGovNumber as JudeOGN,c.OpenGovNumber as UrgeOGN,c.UrgeDate,c.BigUnitBossName,c.SubUnitSecBossName,c.ContactTel,c.SendAddress,c.UrgeTypeID,c.ForFeit,a.Driver,a.DriverBirth,a.DriverID,a.DriverZip,a.DriverAddress,a.IllegalDate,a.IllegalAddress,a.DealLineDate,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID from PasserBase a,PasserJude b,PasserUrge c where a.SN="&trim(request("PBillSN"))&" and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+)"

set rsfound=conn.execute(strSql)

if trim(rsfound("UrgeDate"))<>"" then
	UrgeDate=gInitDT(rsfound("UrgeDate"))
else
	UrgeDate=gInitDT(date)
end if
PrintDate=split(gArrDT(date),"-")
strUInfo="select * from UnitInfo where UnitID='"&trim(rsfound("BillUnitID"))&"'"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then
	theSubUnitSecBossName=trim(rsUInfo("SecondManagerName"))
	theBigUnitBossName=trim(rsUInfo("ManageMemberName"))
	theContactTel=trim(rsUInfo("Tel"))
	theBankAccount=trim(rsUInfo("BankAccount"))
	theUnitName=trim(rsUInfo("UnitName"))
end if
if trim(rsfound("SubUnitSecBossName"))<>"" then
	theSubUnitSecBossName=trim(rsfound("SubUnitSecBossName"))
end if
if trim(rsfound("BigUnitBossName"))<>"" then
	theBigUnitBossName=trim(rsfound("BigUnitBossName"))
end if
if trim(rsfound("ContactTel"))<>"" then
	theContactTel=trim(rsfound("ContactTel"))
end if
rsUInfo.close%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>催繳文件</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>

<body>
<form name=myForm method="post">
<center><strong>臺中市警察局（<%=theUnitName%>）<br>違反道路交通管理事件催繳通知書</strong></center>
<table width="645" height="100%" border="1" align="center" cellspacing=0 cellpadding=0>
  <tr align="center">
    <td width="15%" height="44"><span class="style8">發文日期</span></td>
    <td width="24%"><span class="style8"><%=PrintDate(0)%>年<%=PrintDate(1)%>月<%=PrintDate(2)%>日</span></td>
    <td width="12%"><span class="style8">文號</span></td>
    <td width="31%"><span class="style8">警交催字第<%=rsfound("UrgeOGN")%>&nbsp;號</span></td>
    <td width="18%"><span class="style8">局長</span></td>
  </tr>
  <tr align="center">
    <td height="67"><span class="style8">義務人</span></td>
    <td colspan="3"><%=rsfound("Driver")%></td>
    <td rowspan="2" align="center"><%'=trim(theBigUnitBossName)%>&nbsp;</td>
  </tr>
  <tr align="center">
    <td height="61"><span class="style8">住址</span></td>
    <td colspan="3"><%=trim(rsfound("DriverZip"))&trim(rsfound("DriverAddress"))%>&nbsp;</td>
  </tr>
  <tr align="center">
    <td colspan="4" rowspan="4" align="left" valign="top"><p align="left" class="style8 style10">有關台端因違反道路交通管理事件共計x案（詳如清冊），本局分業已依法定程序裁決並送達在案，經依法逕行裁決末在法定期間到案陳述或聲明異議，請接到本通知書後翌日起15日內至本分局繳納，連絡電話：<%=trim(theContactTel)%>或郵政劃撥帳號：<%=trim(theBankAccount)%>或購買匯票郵寄本分局（交通裁決），若逾期仍未繳清罰款，本局將依行政執行法第十一條規定，移送法務部行政執行署所屬行政執行處強制執行。</p>
    <p align="left" class="style8">　　　此致</p>
    <p align="left" class="style8">&nbsp;</p>
    <p align="left" class="style8">&nbsp;</p>
    <p align="left" class="style8">　　　　　　　　　　　　　　　　　　<%=rsfound("Driver")%>　君</p>    </td>
    <td height="37" align="center"><span class="style8">分局長代行</span></td>
  </tr>
  <tr>
    <td height="123" align="center"><%'=trim(theSubUnitSecBossName)%>&nbsp;</td>
  </tr>
  <tr>
    <td height="38" align="center"><span class="style8">承辦人</span></td>
  </tr>
  <tr>
    <td height="108" align="center"><%=Session("Ch_Name")%>&nbsp;</td>
  </tr>
  <tr align="center">
    <td height="47" colspan="5"><span class="style8">中華民國<%=PrintDate(0)%>年<%=PrintDate(1)%>月<%=PrintDate(2)%>日</span></td>
  </tr>
</table>
<input type="Hidden" name="BillSN" value="<%=rsfound("BillSN")%>">
</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}
</script>