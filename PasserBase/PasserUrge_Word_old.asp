<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strRul="select Value from Apconfigure where ID=3"
set rsRul=conn.execute(strRul)
RuleVer=trim(rsRul("Value"))
rsRul.Close

strState="select Level1,Level4 from law where itemid='"&trim(request("Rule1"))&"' and version="&RuleVer
set rsSql=conn.execute(strState)
If not rsSql.eof Then Level4=trim(rsSql("Level4"))
If not rsSql.eof Then Level1=trim(rsSql("Level1"))
rsSql.close

If not ifnull(Trim(request("DriverID"))) Then
	If Mid(Trim(request("DriverID")),2,1)="1" Then
		Sys_Sex="男"
	elseif Mid(Trim(request("DriverID")),2,1)="2" Then
		Sys_Sex="女"
	End if
End if

Sys_IllegalDate=split(gArrDT(trim(request("IllegalDate"))),"-")

if trim(request("BillFillDate"))<>"" then
	UrgeDate=split(gArrDT(request("BillFillDate")),"-")
else
	UrgeDate=split(gInitDT(date),"-")
end if
PrintDate=split(gArrDT(date),"-")

strSQL="select * from UnitInfo where UnitID='"&trim(request("ArrUnitID"))&"'"
set unit=conn.Execute(strSQL)
If Not unit.eof Then
	theUnitID=trim(unit("UnitID"))
	theUnitName=trim(unit("UnitName"))
	theSubUnitSecBossName=trim(unit("SecondManagerName"))
	theBigUnitBossName=trim(unit("ManageMemberName"))
	theContactTel=trim(unit("Tel"))
	theBankAccount=trim(unit("BankAccount"))
	theBankName=trim(unit("BankName"))
	theUnitAddress=trim(unit("Address"))
end if
unit.close
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>違反道路交通管理處罰條例催繳通知書</title>
<style type="text/css">
<!--
.style1 {
	font-size: 16px;
	font-weight: bold;
	line-height:2;
}
.style2 {font-size: 18px; font-family: "標楷體"; line-height:2;}
.style3 {font-size: 18px; line-height:2;}
.style4 {font-family: "標楷體"; line-height:2;}
.style5 {font-size: 18px; line-height:2;}
.style6 {font-family: "標楷體"; font-size: 18px; line-height:2; }
.style7 {
	font-family: "標楷體";
	font-weight: bold;
	font-size: 22px;
	line-height:2;
}
.style8 {
	font-family: "標楷體";
	font-size: 16px;
	line-height:2;
}
.style9 {font-family: "標楷體"; line-height:2;}
.style10 {font-size: 16px; line-height:2;}
.style11 {font-size: 14px;}
.style12 {
	font-size: 14px;
	font-family: "標楷體";
	font-weight: bold;
	line-height:2;
}
.style13 {font-size: 14px; font-family: "標楷體"; line-height:2; }
.style14 {
	font-size: 22px;
	font-family: "標楷體";
	line-height:1;
}
.style15 {font-family: "標楷體"; font-size: 20px; line-height:1; }
.style16 {font-family: "標楷體"; font-size: 20px; line-height:1; }
.style17 {font-family: "標楷體"; font-size: 20px; line-height:1; }
.style18 {font-family: "標楷體"; font-size: 20px; line-height:2; }
.style19 {font-size: 24px; line-height:2; }
.style20 {font-size: 36px; line-height:2; }
.style21 {font-size: 18px; line-height:2; }
.style22 {font-family: "標楷體"; font-size: 16px;}
.style23 {font-family: "標楷體"; font-size: 14px;}
.style24 {font-family: "標楷體"; font-size: 12px;}
.style25 {font-family: "標楷體"; font-size: 24px;}
.style26 {font-family: "標楷體"; font-size: 10px;}
-->
</style>
<!--#include virtual="traffic/Common/css.txt"-->
</head>

<body>
<table width="645" height="100%" border="1" cellspacing=0 cellpadding=0>
  <tr>
    <td height="62" colspan="4"><div align="center" class="style1">違反道路交通管理處罰條例催繳通知書</div></td>
  </tr>
  <tr>
    <td height="58"><div align="center" class="style3">事　　由</div></td>
    <td colspan="3"><span class="style3">違反道路交通管理事件處罰案</span></td>
  </tr>
  
  <tr>
    <td height="55" align="center"><span class="style3">送達文件</span></td>
    <td><span class="style3">催繳通知書</span></td>
	<td height="55" align="center"><span class="style3">發文日期</span></td>
    <td><span class="style3"><%="民國"&UrgeDate(0)&"年"&UrgeDate(1)&"月"&UrgeDate(2)&"日"%></span></td>
  </tr>
  <tr>
    <td height="61" align="center"><span class="style3">受送達人<br>
    姓　　名</span></td>
    <td colspan="3"><span class="style3">被通知人：<%=request("Driver")%>，性別：<%=Sys_Sex%>，身分證統一號碼：<%=request("DriverID")%></span></td>
  </tr>
  <tr>
    <td height="64" align="center"><span class="style3">送達處所</span></td>
    <td colspan="3"><span class="style3">戶籍地：<%=request("DriverAddress")%></span></td>
  </tr>
 <tr valign="top">
    <td height="224" colspan="4">
		<table border=0 width="100%">
			<tr><td height="81" valign="top"><span class="style21">一、</span></td>
			<td valign="top"><span class="style3">台端年度<%=Sys_IllegalDate(0)%>年至<%=PrintDate(0)%>年違反道路管理事件１件，應繳納新台幣&lt;&lt;金額&gt;&gt;<%=Level1%>　元正，已逾期未繳。（如附表）</span></td></tr>
			<tr><td height="110" valign="top"><span class="style3">二、</span></td>
			<td valign="top"><span class="style3">請於收受本通知書後，十五日內至本分局繳納或用匯票匯本分局，每一案仍維持原罰<%'=rsfound("Level1")%>。右開繳納罰款如未能依時限繳納，已違反行政執行法第四條，金錢給付義務逾期不覆行者，每一違規案件，將對台端採最高罰（<%=Level4%>）並逕送行政執行署，依法強制執行。</span></td></tr>
			<tr><td height="81" valign="top"><span class="style3">三、</span></td>
			<td valign="top"><span class="style3">為顧及台端之權益及本於便民措施，特再通知。（<%=theContactTel%>）</span></td></tr>
			<tr><td height="88" valign="top"><span class="style3">四、</span></td>
			<td valign="top"><span class="style3">戶名：<%=theBankName%>。<br>帳號：<%=theBankAccount%>。</span></td></tr>
	  </table>
  </tr>
  <tr>
    <td height="121" align="center"><span class="style21">舉　發<br>
    單　位</span></td>
    <td><%=Sys_City&"政府警察局<br>"&theUnitName%></td>
    <td colspan="2" rowspan>承辦人：<%response.Write "　　　　　　　　"
	response.Write "<br><br>"
	If sys_City<>"嘉義市" and sys_City<>"澎湖縣" Then
		if Session("Unit_ID") <> "05FG" and Session("Unit_ID") <> "F000" then
			If sys_City<>"宜蘭縣"  and sys_City<>"台南市" and sys_City<>"台南縣" and sys_City<>"花蓮縣" then
				If sys_City<>"嘉義縣" and sys_City<>"高雄縣" then
					response.write "局長："&left(theBigUnitBossName&"　　　　　　　　",5)
				else
					response.write "分局長："&left(theSubUnitSecBossName&"　　　　　　　　",5)
				end if
			end if
		end if
	elseif sys_City="嘉義市" then
		response.write "分局長："&left("　　　　　　　　",5)
	end if%>&nbsp;</span></td>
  </tr>
</table>
</body>
</html>
