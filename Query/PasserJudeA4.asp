<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
fMnoth=month(now)
if fMnoth<10 then
fMnoth="0"&fMnoth
end if
fDay=day(now)
if fDay<10 then
fDay="0"&fDay
end if
fname=year(now)&fMnoth&fDay&"_送達證書.doc"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/msword; charset=MS950" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>送達證書</title>
<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
}
.style2 {font-family: "標楷體"}
.style3 {font-size: 18px}
.style4 {font-family: "標楷體"}
.style5 {font-size: 18px}
.style6 {font-family: "標楷體"; font-size: 18px; }
.style7 {
	font-family: "標楷體";
	font-weight: bold;
	font-size: 22px;
}
.style8 {
	font-family: "標楷體";
	font-size: 16px;
}
.style9 {font-family: "標楷體"}
.style10 {font-size: 16px}
.style11 {font-size: 14px}
.style12 {
	font-size: 14px;
	font-family: "標楷體";
	font-weight: bold;
}
.style13 {font-size: 14px; font-family: "標楷體"; }
.style14 {
	font-size: 30px;
	font-family: "標楷體";
}
.style15 {font-family: "標楷體"; font-size: 28px; }
.style16 {font-family: "標楷體"; font-size: 20px; }
.style17 {font-family: "標楷體"; font-size: 23px; }
.style18 {font-family: "標楷體"; font-size: 24px; }
.style19 {font-size: 24px}
.style20 {font-size: 36px}
.style21 {font-size: 18px}
-->
</style>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<%
BillSN=split(trim(request("PBillSN")),",")

for i=0 to Ubound(BillSN)

strSql="select StoreAndSendMailNumber,MailTypeID,MailDate,MAILCHKNUMBER,FirstBarCode from BillMailHistory where BILLSN="&BillSN(i)

set rs=conn.execute(strSql)
if Not rs.eof then Sys_FirstBarCode=trim(rs("FirstBarCode"))
if Not rs.eof then Sys_StoreAndSendMailNumber=trim(rs("StoreAndSendMailNumber"))
if Not rs.eof then Sys_MailTypeID=trim(rs("MailTypeID"))
if Not rs.eof then Sys_MailDate=trim(rs("MailDate"))
if Not rs.eof then Sys_MAILCHKNUMBER=trim(rs("MAILCHKNUMBER"))
rs.close

strSql="select * from BillBase where SN="&BillSN(i)
set rsSql=conn.execute(strSql)
Sys_BillTypeID=0
if Not rsSql.eof then Sys_BillTypeID=trim(rsSql("BillTypeID"))
if Not rsSql.eof then Sys_Driver=trim(rsSql("Driver"))
if Not rsSql.eof then Sys_DriverID=trim(rsSql("DriverID"))
if Not rsSql.eof then Sys_DriverAddress=trim(rsSql("DriverAddress"))
if Not rsSql.eof then Sys_DriverZip=trim(rsSql("DriverZip"))
rsSql.close

strSQL="select ZipName from Zip where ZipID='"&Sys_DriverZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_DriverZipName=trim(rszip("ZipName"))
rszip.close

strBil="select distinct BillSN,BillNo,CarNo from DCILOG where BillSN="&BillSN(i)
set rsbil=conn.execute(strBil)

strSql="select * from BillbaseDCIReturn where BillNo='"&trim(rsbil("BillNo"))&"' and CarNo='"&trim(rsbil("CarNo"))&"' and ExchangetypeID='W'"
set rsfound=conn.execute(strSql)

if Not rsfound.eof then Sys_Owner=trim(rsfound("Owner"))
if Not rsfound.eof then Sys_OwnerAddress=trim(rsfound("OwnerAddress"))
if Not rsfound.eof then Sys_OwnerZip=trim(rsfound("OwnerZip"))
strSQL="select ZipName from Zip where ZipID='"&Sys_OwnerZip&"'"
set rszip=conn.execute(strSQL)
if Not rszip.eof then Sys_OwnerZipName=trim(rszip("ZipName"))
rszip.close
rsbil.close
%>
<table width="645" border="0">
  <tr>
    <td width="644"><p align="center" class="style7">訴願文書郵務送達證書</p></td>
  </tr>
</table>
<table width="654" height="100%" border="1" cellspacing=0 cellpadding=0>
  <tr>
    <td colspan="2" nowrap> <div align="center" class="style12">受 送 達 人 名 稱 姓 名 地 址</div></td>
    <td colspan="3"><p align="left" class="style13"><%
		if trim(Sys_BillTypeID )="1" then
			response.write Sys_Driver&"<br>"&Sys_DriverZipName&Sys_DriverAddress
		else
			response.write Sys_Owner&"<br>"&Sys_OwnerZipName&Sys_OwnerAddress
		end if
	%>&nbsp;</p></td>
  </tr>
  <tr>
    <td colspan="2"> <p align="center" class="style12">文　　　　　　　　　　　　號</p></td>
    <td colspan="3"> <p align="left" class="style12">　　　 字第　<%=Sys_MAILCHKNUMBER%>　號</p></td>
  </tr>
  <tr>
    <td colspan="2"> <p align="center" class="style12">送 達 文 書 （ 含 案 由 ）</td>
    <td colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td rowspan="2" nowrap> <p align="center" class="style12">原寄郵局日戳</p></td>
    <td rowspan="2" nowrap> <p align="center" class="style12">送達郵局日戳</p></td>
    <td colspan="2"><p align="center" class="style13"> 送達處所（由送達人填記） </p></td>
    <td rowspan="2" nowrap><p align="center" class="style13"> 送達人簽章 </p></td>
  </tr>
  <tr>
    <td colspan="2"><p class="style13">□ 同上記載地址<br>□ 改送： </p>    </td>
  </tr>
  <tr>
    <td rowspan="2">&nbsp;</td>
    <td rowspan="2">&nbsp;</td>
    <td colspan="2"><div align="center" class="style13"> 送達時間（由送達人填記） </div></td>
    <td rowspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" nowrap><p class="style13">中　華　民　國　　　年　　　月　　　日</p>
      <span class="style13">　　　　　　　　　　午　　　時　　　分 </span></td>
  </tr>
  <tr>
    <td colspan="5"><div align="center" class="style13"> 送　　　　　　　　　　達　　　　　　　　　　方　　　　　　　　　　式 </div></td>
  </tr>
  <tr>
    <td colspan="5"><div align="center" class="style13"> 由　　送　　達　　人　　在　　□　　上　　劃　　 ˇ 　　選　　記 </div></td>
  </tr>
  <tr>
    <td colspan="2" valign="top"><div align="center" class="style13">  
      <table width="100%" height="100%" border="0">
        <tr>
          <td width="9%" align="center">□</td>
          <td width="91%" class="style13">已將文書交與應受送達人 </td>
        </tr>
      </table>
    </div></td>
    <td colspan="3"><span class="style13"> □ 本人（簽名或蓋章） </span></td>
  </tr>
  <tr>
    <td colspan="2" valign="top"> 
      <table width="100%" height="100%" border="0">
        <tr>
          <td width="9%" align="center" valign="top" class="style13">□</td>
          <td width="91%"><span class="style13">未獲會晤本人，已將文書交與有辨別事理能力之同居人、受雇人或願代為收受而居住於同一住宅之主人 </span></td>
        </tr>
      </table></td>
    <td colspan="3"><p class="style13">□ 同居人<br>□ 受雇人　　　　　　　　　（簽名或蓋章）<br>□ 居住於同一住宅之主人<br>□ 應送達處所接收郵件人員 </span></td>
  </tr>
  <tr>
    <td colspan="2" valign="top"> 
      <table width="100%" height="100%" border="0">
        <tr>
          <td width="9%" align="center" valign="top" class="style13">□</td>
          <td width="91%" class="style13">應受送達之本人、同居人或 受雇人收領後，拒絕或不能簽名或蓋章者，由送達人記明其事由</td>
        </tr>
      </table></td>
    <td colspan="3"><span class="style13"> □ 本人　　　　　　　　　　　　　　　拒絕收領 </span></td>
  </tr>
  <tr>
    <td height="132" colspan="2"><table width="100%" height="100%" border="0">
        <tr>
          <td height="44" align="center" valign="top" class="style13">□</td>
          <td valign="top" class="style13">未獲會晤本人亦無受領文書之同居人或受雇人，已將該送達文書：</td>
        </tr>
        <tr>
          <td width="9%" height="45" align="center" valign="top" class="style13">□</td>
          <td width="91%" valign="top" class="style13">應受送達人無法律上理由拒絕收領，並有難達留置情事，已將該送達文書： </td>
        </tr>
      </table></td>
    <td class="style13" nowrap>□　存於　　　　警察派出所<br>
						□　寄存於　　　鄉（鎮、市、區）公所<br>
						□　寄存於　　　鄉（鎮、市、區）<br>
						　　　　　　　　村（里）辦公處。<br>
						□　寄存於　　　郵局</td>
    <td colspan="2"><span class="style13"><span class="style13"> 並作送達通知書二份，一份黏貼於應受送達人住居所、事務所或營業所門首，一份□交由鄰居轉交或□置於應受送達人之信箱或其他適當之處所，以為送達。 </span></span></td>
  </tr>
  <tr>
    <td colspan="2"><div align="center"><span class="style8 style11 style9"><span class="style13"> 送　達　人　注　意　事　項 </span></span></div></td>
    <td colspan="3">
      <table width="100%" height="100%" border="0">
        <tr>
          <td height="44" align="center" valign="top" class="style13"><span class="style8 style11 style9">一、</span></td>
          <td valign="top" class="style13"><span class="style13">上述送達方法送達者，送達人應即將本送達證書，提出於交送達之行政機關附卷。 </span></td>
        </tr>
        <tr>
          <td width="9%" height="45" align="center" valign="top" class="style13">二、</td>
          <td width="91%" valign="top" class="style13">法依上述送達方法送達者，送達人應作記載該事由之報告書，提出於交送達之行政機關附卷，並繳回應送達之文書。</td>
        </tr>
      </table></td>
  </tr>
</table>
<p><strong>※</strong><strong>１、本送達證書請繳回○○○（交送達機關）地址： </strong></p>
<p><strong>　２、寄存送達之文書，應保存 3 個月，如未經領取，請退還交送達機關。 </strong></p>
（本證由各機關自行製用；規格 A4 ，※部分建議以紅色套印）
<br><br><br><br><br>
<%next%>
</body>
</html>
