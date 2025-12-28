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
fname=year(now)&fMnoth&fDay&"_移送書.doc"
'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/msword; charset=MS950" 

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

sys_City=replace(sys_City,"台中縣","台中市")
sys_City=replace(sys_City,"台南縣","台南市")

showCreditor=false
if sys_City="台中市" or sys_City = "彰化縣" or sys_City = "台南市" or sys_City = "高雄市" or sys_City = "高雄縣" or sys_City="宜蘭縣" or sys_City="基隆市" or sys_City="澎湖縣" or sys_City="屏東縣" then
	showCreditor=true
end If


strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close

BillPageUnit=trim(Request("PasserUnitName"))
chName=Request("operat")
JobName=Request("operatLevel")
If ifnull(JobName) Then jobName="警員"

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

DB_UnitName=trim(Request("ArrUnitName"))
DB_UnitTel=trim(Request("UnitTEL"))
DB_BankName=trim(Request("UnitAccountName"))
DB_BankAccount=trim(Request("UnitAccount"))
PrintDate=split(gArrDT(date),"-")
MakeSureDate=split(gArrDT(DateAdd("d",20,Request("SendDate"))),"-")
LimitDate=split(gArrDT(DateAdd("d",35,Request("SendDate"))),"-")
Sys_Address=Request("DriverAddress")
Sys_Address=Request("DriverZip")&Sys_Address
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>無標題文件</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-family: "標楷體"; font-size: 14px; }
.style2 {font-family: "標楷體"; font-size: 24px;  line-height:2;}
.style3 {font-family: "標楷體"; font-size: 16px; }
-->
</style>
</head>

<body>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="..\smsx.cab#Version=6,1,432,1">
</object>
<table width="90%" height="1%" border="0" cellspacing=0 cellpadding=0>
<tr><td align="right">
		<table width="200" border="1" cellspacing=0 cellpadding=0>
		  <tr>
			<td width="60" class="style1">移送案號</td>
			<td width="134" align="left" class="style1"><%=Request("BillNo")%></td>
		  </tr>
		</table>
	</td>
</tr>
</table>
		<table width="90%" height="90%" border="1" cellspacing=0 cellpadding=0>
		  <tr>
			<td colspan="4" align="left" class="style2">　　　　　<%=thenPasserCity&replace(DB_UnitName,trim(thenPasserCity),"")%>行政執行案件移送書<br>
			<table border="0" width="100%" height="100%" cellspacing="0" cellpadding="0">
			<td width="300">
			<span class="style1">承辦人：<%=JobName&"&nbsp;"&ChName%></span></td><td><span class="style1">發文日期：<%="　"&PrintDate(0)&"年　"&PrintDate(1)&"月　"&PrintDate(2)&"日"%></span></td>
			<tr>
			<td>
			<span class="style1">電話：<%=DB_UnitTel%></span></td><td><span class="style1">發文字號：<%
				response.write BillPageUnit&"交字第"&Request("SendWordNum")&"號"
			%></span></td>
			</table>
			</td>
		  </tr>
		  <tr>
			<td colspan="2" align="center" class="style3" width="55%">義　　　　務　　　　人</td>
			<td colspan="2" align="center" class="style3" width="45%">法定代理人或代表人</td>
		  </tr>
		  <tr>
			<td width="120" class="style3">姓名或名稱</td>
			<td class="style3"><%=Request("Driver")%></td>
			<td colspan="2" class="style3">&nbsp;
			</td>
		  </tr>
		  <tr>
			<td class="style3">出生年月日</td>
			<td class="style3"><%
				if trim(Request("DriverBirth"))<>"" then
					DriverBirth=split(gArrDT(Request("DriverBirth")),"-")
					response.write "　"&DriverBirth(0)&"年　"&DriverBirth(1)&"月　"&DriverBirth(2)&"日"
				end if%>&nbsp;</td>
			<td colspan="2" class="style3">&nbsp;
			</td>
		  </tr>
		  <tr>
			<td class="style3">性　　　　別</td>
			<td class="style3"><%
				If not ifnull(Trim(Request("DriverID"))) Then
					If Mid(Trim(Request("DriverID")),2,1)="1" Then
						response.Write("男")
					elseif Mid(Trim(Request("DriverID")),2,1)="2" Then
						response.Write("女")
					End if
				End if
			%>&nbsp;</td>
			<td colspan="2">&nbsp;</td>
		  </tr>
		  <tr>
			<td class="style3">職　　　　業</td>
			<td>&nbsp;</td>
			<td colspan="2">&nbsp;</td>
		  </tr>
		  <tr>
			<td class="style3" nowrap>身分證統一號<br>碼或營利事業<br>統 一 編 號</td>
			<td class="style3"><%=Request("DriverID")%></td>
			<td colspan="2" class="style3">&nbsp;</td>
		  </tr>
		  <tr>
			<td class="style3" nowrap>住 居 所 或<br>事 務 所 、<br>營 業 所 地<br>址 及 郵 遞<br>區　　　號</td>
			<td class="style3">住：<%=Sys_Address%>&nbsp;<br>
				居：
			</td class="style3">
			<td colspan="2" class="style3">住：&nbsp;<br>
				居：
			</td>
		  </tr>
		  <tr>
			<td rowspan="2" class="style3">執行標的物<br>所　在　地</td>
			<td rowspan="2" class="style3">如附件財產目錄所載</td>
			<td width="124" class="style3">分   署<br>收案日期</td>
			<td width="200" class="style3"><%
				'if trim(rsfound("SendDate"))<>"" then
					'SendDate=split(gArrDT(rsfound("SendDate")),"-")
					'response.write SendDate(0)&"年"&SendDate(1)&"月"&SendDate(2)&"日"
				'end if%>&nbsp;</td>
			</tr>
			<tr>
			<td width="124" class="style3">行政處分或<br>裁定確定日</td>
			<td width="200" class="style3">
				■　<%=MakeSureDate(0)%>年　<%=MakeSureDate(1)%>月　<%=MakeSureDate(2)%>日<br>
				□尚未確定
			</td>
		  </tr>
		  <tr>
			<td rowspan="3" class="style3">義務發生之<br>原因與日期</td>
			<td rowspan="3" class="style3"><%
				response.write "違反道路交通管理處罰條例<br>第"&left(trim(Request("Rule1")),2)&"條"
				if len(trim(Request("Rule1")))>7 then response.write "之"&right(trim(Request("Rule1")),1)
				response.write Mid(trim(Request("Rule1")),3,1)&"項"&Mid(trim(Request("Rule1")),4,2)&"款"&Mid(trim(Request("Rule1")),6,2)&"規定。"
				response.write "<br>"
				if trim(Request("SendDate"))<>"" then
					IllegalDate=split(gArrDT(Request("IllegalDate")),"-")
					response.write IllegalDate(0)&"年"&IllegalDate(1)&"月"&IllegalDate(2)&"日<br>"
					response.write BillPageUnit&"交字第"&Request("PasserNo")&"號"
				end if%>&nbsp;</td>
			<td class="style3">繳&nbsp;納&nbsp;期&nbsp;間<br>&nbsp;屆　滿　日</td>
			<td class="style3">　<%=LimitDate(0)%>年　<%=LimitDate(1)%>月　<%=LimitDate(2)%>日</td>
		  </tr>
		  <tr>
			<td class="style3">徵&nbsp;收&nbsp;期&nbsp;間<br>&nbsp;屆　滿　日</td>
			<td class="style3">　　年　　月　　日</td>
		  </tr>
		  <tr>
			<td class="style3">應納金額</td>
			<td class="style3">新台幣<%=Request("PasserPay")%>元<br>（細目詳如附件）</td>
		  </tr>
		  <tr>
			<td rowspan="2" class="style3">移送法條</td>
			<td rowspan="2" class="style3">
				■依據行政執行法第11條<br>
				■依據道路交通管理處罰條例第<%=left(trim(Request("Rule1")),2)%>條</td>
			<td class="style3">催繳情形</td>
			<td class="style3">□業經催繳<br>
				□未經催繳</td>
		  </tr>
		  <tr>
			<td class="style3">催繳方式</td>
			<td class="style3">□電話催繳<br>
				□明信片或信函方式催繳<br>
				□其他方式（方式為　）</td>
		  </tr>
		  <tr>
			<td class="style3">附件</td>
			<td colspan="3">
				<table border="0" width="100%">
				  <tr>
					<td width="278" class="style3" nowrap>
						<%
						if trim(Request("AttatchTable"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>附表<br>
						<%
						if trim(Request("AttatchJude"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>處分文書、裁定書或義務人依法令負<br>　有義務之證明文件及送達證明文件<br>
						<%
						if trim(Request("AttatchUrge"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>義務人經限期履行而逾期仍不履行<br>　之證明文件及送達證明文件<br>
						戶名：<%=theBankName%></td>
					<td width="209" class="style3" nowrap>
						<%
						if trim(Request("AttatchFortune"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>義務人之財產目錄<br>
						<%
						if trim(Request("AttatchGround"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>土地登記簿謄本<br>
						<%
						if trim(Request("AttatchRegister"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>義務人之戶藉資料<br>
						<%
						if trim(Request("AttatchFileList"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>保全措施之資料<br>
						<%
						if trim(Request("ATTATPOSTAGE"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>執行（債權）憑證<br>
						帳戶：<font size=2><%=theBankAccount%></font>
					</td>
				  </tr>
			  </table>
			</td>
		  </tr>
		  <tr>
			<td class="style3">保全措施</td>
			<td colspan="3" class="style3"><%
						if trim(Request("SAFETOEXIT"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>已限制出境<%
						if trim(Request("SAFEACTION"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>已禁止處分<%
						if trim(Request("SAFEASSURE"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>已提供擔保<%
						if trim(Request("SAFEDETAIN"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>已假扣押<%
						if trim(Request("SAFESHUTSHOP"))<>"" then
							response.write "■"
						else
							response.write "□"
						end if
						%>已勒令停業</td>
		  </tr>
		  <tr>
			<td colspan="4">
				<table border="0" width="100%">
					<tr>
						<td class="style2">　　　此　　　致</td>
					</tr>
					<tr>
						<td class="style2">法務部行政執行署　<%
						If showCreditor Then
							If not ifnull(Request("AgentAddress")) Then
								Response.Write Request("AgentAddress")
							else
								If not ifnull(Request("DriverZip")) Then
									strSQL="select Administrative from zip where zipid='"&trim(Request("DriverZip"))&"'"
									set rszip=conn.execute(strSQL)
									If not rszip.eof Then
										Response.Write replace(rszip("Administrative"),"分署","")
									end if
									rszip.close
								else
									tmpzip=getzip(Request("DriverAddress"))
									If tmpzip<>"null" Then
										strSQL="select Administrative from zip where zipid='"&trim(tmpzip)&"'"

										set rszip=conn.execute(strSQL)
										Response.Write replace(rszip("Administrative"),"分署","")
										rszip.close
									End if								
								End if
							End if

						else
							Response.Write Request("AgentAddress")
						End if
						%>　分署</td>
					</tr>
					<tr><td>&nbsp;</td>
				</table>
			</td>
		  </tr>
		</table>
</body>
</html>
</body>
</html>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
	window.focus();
	printWindow(true,10,10,10,10);
</script>