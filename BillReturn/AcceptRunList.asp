<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media="print">
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {font-size: 16px; line-height:1;}
.style2 {font-size: 14px}
.style3 {font-size: 10px}
-->
</style>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<title>逕舉點收清冊</title>
<%
	strUInfo="select * from Apconfigure where ID=40"
	set rsUInfo=conn.execute(strUInfo)
	if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
	rsUInfo.close
	set rsUInfo=nothing

	BasSQL="select a.*,b.UnitName,c.chname BillUnitName,c.LoginID,decode(a.RecordStateID,'0','正常','退件') RecordType,d.ChName AcceptName1,e.Chname AcceptName2 from (select * from BillRunCarAccept where billunitid='"&trim(Request("DB_BillUnitID"))&"' and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)&"  and recorddate="&funGetDate(Request("DB_RecordDate"),1)&") a,UnitInfo b,memberdata c,memberdata d,memberdata e where a.BillUnitID=b.UnitID and a.BillMemID1=c.MemberID(+) and a.RecordMemberID1=d.MemberID(+) and a.RecordMemberID2=e.MemberID(+) order by a.RecordDate"

	strCnt="select count(*) as cnt from ("&BasSQL&")"
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		if trim(rsCnt("cnt"))="0" then
			pagecnt=1
		else
			pagecnt=fix(cdbl(rsCnt("cnt"))/50+0.9999999)
		end if
	end if
	rsCnt.close
	set rsCnt=nothing

	set rsfound=conn.execute(BasSQL)
%>

</head>
<body class="pageprint">
<%
CntSN=1
SumMemy=0
While Not rsfound.eof

if CntSN >1 then response.write "<br><div class=""PageNext""></div>"

tmp_Accept=split(gArrDT(trim(rsfound("AcceptDate"))),"-")
Sys_Accept=tmp_Accept(0)&"年"&tmp_Accept(1)&"月"&tmp_Accept(2)&"日"

%>
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td class="style1" align="center">
		</td>
		<th class="style1" align="center">
			<%=thenPasserCity&rsfound("UnitName")%>
		</th>
		<td class="style1" align="center">
		</td>
	</tr>
	<tr>
		<td class="style1" align="center">
		</td>
		<th class="style1" align="center">
			<%=Sys_Accept&"違反道路交通管理事件逕舉移送表"%>
		</th>
		<td class="style2" align="right">
			<%=fix(cdbl(CntSN)/40+0.9999999)&" of "&pagecnt%>
		</td>
	</tr>
	<tr><td colspan="3">
	<table width="640" border="1" cellspacing="0" cellpadding="0">
		<tr>
			<th nowrap>編號</th>
			<th nowrap>標示單號碼</th>
			<th nowrap>違規日</th>
			<th nowrap>車牌號碼</th>
			<th nowrap>違反條款</th>
			<th nowrap>違規地點</th>
			<th nowrap>舉發員警</th>
			<th>備註</th>
		</tr><%
		for i=1 to 50
			if rsfound.eof then exit for

			Response.Write "<tr>"

			Response.Write "<td class=""style2"" nowrap>"
			Response.Write CntSN
			Response.Write "</td>"

			Response.Write "<td class=""style2"" nowrap>"
			Response.Write trim(rsfound("BillNo"))
			Response.Write "</td>"

			Response.Write "<td class=""style2"" nowrap>"
			Response.Write gInitDT(rsfound("ILLEGALDATE"))
			Response.Write "</td>"

			Response.Write "<td class=""style2"" nowrap>"
			Response.Write trim(rsfound("CarNo"))
			Response.Write "</td>"

			Response.Write "<td class=""style2"" nowrap>"
			Response.Write trim(rsfound("RULE1"))
			Response.Write "</td>"

			Response.Write "<td class=""style2"" nowrap>"
			Response.Write trim(rsfound("ILLEGALADDRESS"))
			Response.Write "</td>"

			Response.Write "<td class=""style2"" nowrap>"
			Response.Write trim(rsfound("BillUnitName"))
			Response.Write "</td>"

			Response.Write "<td class=""style2"">"

			If trim(rsfound("RecordType"))="退件" Then
				Response.Write "退件："
			End if
			
			Response.Write trim(rsfound("NOTE"))&"&nbsp;"
			Response.Write "</td>"
			Response.Write "</tr>"
			CntSN=CntSN+1
			rsfound.MoveNext
		next
	%>
		</table>
	</td></tr>
</table>
<%
Wend
rsfound.close
set rsfound=nothing
%>
共計:   <%=CntSN-1%>  筆

</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
	return win;
}
function funcPrintCaseDataListExecel(){
	//UrlStr="PrintCaseDataList_Execel.asp";
	//newWin(UrlStr,"inputWin",790,550,0,0,"yes","yes","yes","no");
	location='PrintCaseDataList_Excel.asp';
}
function DP(){
	window.focus();
	window.print();
}
</script>
<%
conn.close
set conn=nothing
%>