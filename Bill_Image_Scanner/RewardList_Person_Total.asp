<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--

-->
</style>
<script language="JavaScript">
window.focus();
</script>
<head>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<!--#include virtual="traffic/Common/css.txt"-->
<title>直接執行及共同人員獎勵金年度總額清冊</title>
<%
Server.ScriptTimeout = 186000
Response.flush
 	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing


RewardDate1=""
RewardDate2=""
If Trim(request("sYear"))<>"" Then
	RewardDate1=Trim(request("sYear"))+1911 & "/" & Trim(request("sMonth")) & "/1"
End If 
If Trim(request("eYear"))<>"" Then
	RewardDate2=Trim(request("eYear"))+1911 & "/" & Trim(request("eMonth")) & "/1"
End If 
If RewardDate2<>"" Then
	SqlDate=" and YearMonth between to_date('"&RewardDate1&"','yyyy/mm/dd') and to_date('"&RewardDate2&"','yyyy/mm/dd')"
Else
	SqlDate=" and YearMonth=to_date('"&RewardDate1&"','yyyy/mm/dd')"
End If

strDel1="Delete from RewardMonthTotalData where RunMem="&Session("User_ID")
conn.execute strDel1

Str1="select distinct(Creditid) from Rewardmonthdata where directortogether in ('0','1') and UnitID in ('0807','0820','0830','0840','0850','0851','0852','0853','0854','0855','0856','0857','0858','0859','0860','0861','0862','0863','0864') " & SqlDate
set rs1=conn.execute(Str1)
while Not rs1.eof
	sShouldGetMoney=0
	sRealGetMoney=0
	str2="select sum(ShouldGetMoney) as sShouldGetMoney,sum(RealGetMoney) as sRealGetMoney from Rewardmonthdata where directortogether in ('0','1') and UnitID in ('0807','0820','0830','0840','0850','0851','0852','0853','0854','0855','0856','0857','0858','0859','0860','0861','0862','0863','0864') and CreditID='"&Trim(rs1("Creditid"))&"'" & SqlDate
	Set rs2=conn.execute(str2)
	If Not rs2.eof Then
		If Not isNull(rs2("sShouldGetMoney")) Then
			sShouldGetMoney=CDbl(rs2("sShouldGetMoney"))
		End If
		If Not isNull(rs2("sRealGetMoney")) Then
			sRealGetMoney=CDbl(rs2("sRealGetMoney"))
		End if
	end If
	rs2.close
	Set rs2=nothing
	
	MaxSn=0
	strMaxSn="select Max(Sn) as MaxSN from RewardMonthTotalData"
	Set rsMaxSn=conn.execute(strMaxSn)
	If Not rsMaxSn.eof Then
		If Not IsNull(rsMaxSn("MaxSN")) Then
			MaxSn=CDbl(rsMaxSn("MaxSN"))
		End if
	End If
	rsMaxSn.close
	Set rsMaxSn=nothing

	str3="select * from (select * from Rewardmonthdata where directortogether in ('0','1') and UnitID in ('0807','0820','0830','0840','0850','0851','0852','0853','0854','0855','0856','0857','0858','0859','0860','0861','0862','0863','0864') and CreditID='"&Trim(rs1("Creditid"))&"'" & SqlDate & " order by YearMonth Desc) where rownum<=1"
	Set rs3=conn.execute(str3)
	If Not rs3.eof Then
		sMEMPOINT="null"
		sMEMCASECNT="null"
		sMEMBERID="null"
		If not isnull(rs3("MEMBERID")) Then
			sMEMBERID=trim(rs3("MEMBERID"))
		End If
		If not isnull(rs3("MEMPOINT")) Then
			sMEMPOINT=CDbl(rs3("MEMPOINT"))
		End If
		If not isnull(rs3("MEMCASECNT")) Then
			sMEMCASECNT=CDbl(rs3("MEMPOINT"))
		End If

		strIns1="Insert Into RewardMonthTotalData values('"&Trim(rs3("DIRECTORTOGETHER"))&"'" &_
		",to_date('"&Year(rs3("YEARMONTH"))&"/"&month(rs3("YEARMONTH"))&"/"&day(rs3("YEARMONTH"))&"','yyyy/mm/dd')" &_
		",'"&Trim(rs3("UNITID"))&"','"&Trim(rs3("LOGINID"))&"','"&Trim(rs3("CHNAME"))&"'" &_
		","&sMEMBERID&",'"&Trim(rs3("CREDITID"))&"',"&sShouldGetMoney&","&sRealGetMoney&",sysdate" &_
		","&Trim(Session("User_ID"))&","&MaxSn&",'"&Trim(rs3("SUBUNIT"))&"','"&Trim(rs3("BANKID"))&"'" &_
		",'"&Trim(rs3("BANKACCOUNT"))&"',"&sMEMPOINT&","&Trim(rs3("JOBTITAL"))&"" &_
		","&sMEMCASECNT&","&Trim(Session("User_ID"))&")"
		'response.write strIns1&"<br>"
		conn.execute strIns1
	End If
	rs3.close
	Set rs3=nothing
	
	rs1.movenext
wend
rs1.close
set rs1=Nothing

%>
</head>
<body leftmargin="25" topmargin="5" marginwidth="0" marginheight="0" >
<form name=myForm method="post">
<%
pagecnt=15
PageNo3=0

strUnit="select * from UnitInfo where UnitID in ('0807','0820','0830','0840','0850','0851','0852','0853','0854','0855','0856','0857','0858','0859','0860','0861','0862','0863','0864') order by UnitID"
set rsU=conn.execute(strUnit)
while Not rsU.eof
	
	PageNo=0
	FinalPage=0
	AllTotal=0
	UnitPersonCnt=0
	UnitPersonCnt_Add=0
	strCnt="select count(*) as cnt from RewardMonthTotalData where UnitID='"&trim(rsU("UnitID"))&"' " &_
		" and directortogether in ('0','1')" & SqlDate

	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		UnitPersonCnt=cint(rsCnt("cnt"))
		FinalPage=fix(Cint(rsCnt("cnt"))/ pagecnt + 0.9999999)
	end if
	rsCnt.close
	set rsCnt=nothing

	strQry="select * from RewardMonthTotalData where UnitID='"&trim(rsU("UnitID"))&"' " &_
		" and directortogether in ('0','1')" & SqlDate &_
		" order by SubUnit,DirectorTogether,LoginID"
	set rsQ=conn.execute(strQry)
	while Not rsQ.eof
	
	PageNo=PageNo+1
	if PageNo>1 then
		response.write "<div class=""PageNext"">&nbsp;</div>"
	end if
%>
<table border="0" width="1000" cellpadding="2" cellspacing="0" align="center" >
	<tr>
		<td align="center" colspan="2"><span class="style3"><%=trim(rsU("UnitName"))%></span>&nbsp; &nbsp; &nbsp; &nbsp;<span class="style1">直接人員及共同人員年度總額清冊</span></td>
	</tr>
	<tr>
		<td align="left" valign="bottom" class="style2">列印日期 <%=year(now)&"/"&month(now)&"/"&day(now)%></td>
		<td align="right" class="style2">頁次：<%=PageNo%> / <%=FinalPage%> 頁<br>列印人員：<%=trim(Session("Ch_Name"))%></td>
	</tr>
</table>
<table border="1" width="1000" cellpadding="5" cellspacing="0" align="center" >
	<tr>
		<td align="center" class="style2" width="8%">單位</td>
		<td align="center" class="style2" width="8%">員警代號</td>
		<td align="center" class="style2" width="10%">身分證證號</td>
		<td align="center" class="style2" width="9%">職稱</td>
		<td align="center" class="style2" width="9%">姓名</td>
		<td align="center" class="style2" width="8%">金額</td>
		<td align="center" class="style2" width="8%">局號</td>
		<td align="center" class="style2" width="8%">帳號</td>
		<td align="center" class="style2" width="9%">備考</td>
	</tr>
<%	
	PageTotal=0
	for i=1 to pagecnt
		if rsQ.eof then exit for

		UnitPersonCnt_Add=UnitPersonCnt_Add+1
%>
	<tr>
		<td align="center" class="style2" height="22"><%=trim(rsQ("SubUnit"))%></td>
		<td align="center" class="style2"><%
		if trim(rsQ("LoginID"))="" then
			response.write "&nbsp;"
		else
			response.write trim(rsQ("LoginID"))
		end if
		%></td>
		<td align="center" class="style2"><%
		if trim(rsQ("creditid"))="" then
			response.write "&nbsp;"
		else
			'response.write left(trim(rsQ("creditid")),7)&"***"
			if trim(Session("Unit_ID"))<>"8800" then
			response.write trim(rsQ("creditid"))
			else
			response.write "&nbsp;"
			end if
		end if
		%></td>
		<td align="center" class="style2"><%
		if trim(rsQ("JobTital"))="" then
			response.write "&nbsp;"
		else
			strTitle="select * from Code where ID="&trim(rsQ("JobTital"))
			set rsTitle=conn.execute(strTitle)
			if not rsTitle.eof then
				response.write trim(rsTitle("Content"))
			end if
			rsTitle.close
			set rsTitle=nothing
		end if
		%></td>
		<td align="center" class="style2"><%
		if trim(rsQ("ChName"))="" then
			response.write "&nbsp;"
		else
			response.write trim(rsQ("ChName"))
		end if		
		%></td>
		<td align="center" class="style2"><%
		if isnull(rsQ("RealGetMoney")) then
			response.write "0"
		else
			response.write trim(rsQ("RealGetMoney"))
			PageTotal=PageTotal+cdbl(rsQ("RealGetMoney"))
			AllTotal=AllTotal+cdbl(rsQ("RealGetMoney"))
		end if
		%></td>
		<td align="center" class="style2"><%
		if isnull(rsQ("BankID")) then
			response.write "&nbsp;"
		else
			if trim(Session("Unit_ID"))<>"8800" then
				response.write trim(rsQ("BankID"))
			else
				response.write "&nbsp;"
			end if
		end if	
		%></td>
		<td align="center" class="style2"><%
		if isnull(rsQ("BankAccount")) then
			response.write "&nbsp;"
		else
			if trim(Session("Unit_ID"))<>"8800" then
			response.write trim(rsQ("BankAccount"))
			else
			response.write "&nbsp;"
			end if
		end if	
		%></td>
		<td align="center" class="style2">&nbsp;</td>
	</tr>
<%		Response.flush
		rsQ.movenext
	Next
%>
	<tr>
		<td align="center" class="style2" height="21">小計</td>
		<td colspan="4" class="style2">&nbsp;</td>
		<td align="center" class="style2"><%=PageTotal%></td>
		<td colspan="3" class="style2">&nbsp;</td>
	</tr>
<%
	if cint(UnitPersonCnt_Add) = cint(UnitPersonCnt) then
%>
	<tr>
		<td align="center" class="style2" height="21">總計</td>
		<td colspan="4" class="style2">&nbsp;</td>
		<td align="center" class="style2"><%=AllTotal%></td>
		<td colspan="3" class="style2">&nbsp;</td>
	</tr>
<%
	end if
%>
</table>
<%
	wend
	rsQ.close
	set rsQ=nothing

	rsU.movenext
wend
rsU.close
set rsU=nothing

conn.close
set conn=nothing
%>
</form>
</body>
<script language="JavaScript">
	window.print();
</script>
</html>
