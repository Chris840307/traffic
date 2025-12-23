<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
.pageprint {
  margin-left: 0mm;
  margin-right: 0mm;
  margin-top: 0mm;
  margin-bottom: 0mm;
}
</style>
<title>行人攤販資料建檔清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 1800
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

	strRul="select Value from Apconfigure where ID=3"
	set rsRul=conn.execute(strRul)
	RuleVer=trim(rsRul("Value"))
	rsRul.Close

	If not ifnull(Request("RecordDate1")) Then

		ArgueDate1=gOutDT(request("RecordDate1"))&" 00:00:00"
		ArgueDate2=gOutDT(request("RecordDate2"))&" 23:59:59"
		
		tmp_meb=""
		If not ifnull(Request("Sys_MemberStation")) Then
			tmp_meb=" and MemberStation in('"&Request("Sys_MemberStation")&"')"
		End if 

		If not ifnull(Session("User_ID")) Then
			tmp_meb=tmp_meb&" and RecordMemberID="&Session("User_ID")
		End if 

		BasSQL="(select sn from passerbase where RecordDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and recordstateid=0"&tmp_meb&") tmpPasser"
	
	End if 

	If BasSQL = "" Then

		If Not ifnull(request("Sys_SendBillSN")) Then

			sys_billsn=request("Sys_SendBillSN")
		elseif Not ifnull(request("hd_BillSN")) Then

			sys_billsn=request("hd_BillSN")
		else

			sys_billsn=request("BillSN")
		End If 

		tmp_billsn=split(sys_billsn,",")

		sys_billsn=""

		For i = 0 to Ubound(tmp_billsn)

			If i >0 then

				If i mod 100 = 0 Then

					sys_billsn=sys_billsn&"@"
				elseif sys_billsn<>"" then

					sys_billsn=sys_billsn&","
				end If 
			end if

			sys_billsn=sys_billsn&tmp_billsn(i)

		Next

		tmpSQL=""

		If Ubound(tmp_billsn) >= 100 Then

			sys_billsn=split(sys_billsn,"@")
			
			For i = 0 to Ubound(sys_billsn)
				
				If tmpSQL <>"" Then tmpSQL=tmpSQL&" union all "
				
				tmpSQL=tmpSQL&"select sn from passerbase where sn in("&sys_billsn(i)&")"
			Next

		else

			tmpSQL="select sn from passerbase where sn in("&sys_billsn&")"

		End if 

		BasSQL="("&tmpSQL&") tmpPasser"
	end if

	strCnt="select count(*) as cnt from PasserBase where Exists(select 'Y' from "&BasSQL&" where SN=PasserBase.SN)"
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		if trim(rsCnt("cnt"))="0" then
			pagecnt=1
		else
			pagecnt=fix(Cint(rsCnt("cnt"))/14+0.9999999)
		end if
	end if
	rsCnt.close
	set rsCnt=nothing

	strSQL="select a.RecordDate,a.SN,a.BillNo,a.IllegalDate,a.IllegalAddress,a.BillUnitID,a.Driver,a.DriverID,a.DRIVERSEX,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.Rule1,a.Rule2,a.BillFillDate,a.DeallineDate,a.MemberStation from PasserBase a where Exists(select 'Y' from "&BasSQL&" where SN=a.SN) order by a.MemberStation,a.RecordDate"

	set rsfound=conn.execute(strSQL)
	tmpSQL=strwhere

%>

</head>
<body class="pageprint">
<form name=myForm method="post">
<%
CaseSN=0
If Not rsfound.Bof Then rsfound.MoveFirst 
While Not rsfound.Eof
	if CaseSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="6" align="center">
			<font size="3">行人攤販資料建檔清冊</font>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="left">列印日期:<%=year(now)-1911&"/"&Right("00"&month(now),2)&"/"&Right("00"&day(now),2)%></td>
		<td colspan="3" align="right">Page <%=fix(CaseSN/14)+1%> of <%=pagecnt%> </td>
	</tr>
	<tr>
	<%if sys_City="苗栗縣" then%>
		<td colspan="3" align="left" colspan="2">登入者:<%=Session("Ch_Name")%></td>
		<td colspan="3" align="right">1 式 2 聯(第 1 聯：分局存查)</td>
	<%else%>
		<td colspan="6" align="left" colspan="2">登入者:<%=Session("Ch_Name")%></td>
	<%end if %>
	</tr>
	<tr><td colspan="6"><hr></td></tr>
	<tr>
		<td width="10%">建檔日期</td>
		<td width="15%">違規日期</td>
		<td width="20%">違規時間</td>
		<td width="15%">舉發單位</td>
		<td width="25%">舉發員警</td>
		<td width="15%">扣件</td>
	</tr>
	<tr>
		<td>舉發單號</td>
		<td colspan="3">違規地點</td>
		<td>法條一＼法條二</td>
		<td>罰鍰一＼罰鍰二</td>
	</tr>
	<tr>
		<td>填單日期</td>
		<td>應到案日期</td>
		<td>駕駛人姓名</td>
		<td>駕駛人ID</td>
		<td colspan="2">到案處所</td>
	</tr>
<%
for i=1 to 10
	if rsfound.eof then exit for
%>
	<tr><td colspan="6"><hr></td></tr>
	<tr>
		<td><%=year(trim(rsfound("RecordDate")))-1911&Right("00"&month(trim(rsfound("RecordDate"))),2)&Right("00"&day(trim(rsfound("RecordDate"))),2)%></td>
		<td width="9%"><%
		if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
			response.write year(trim(rsfound("IllegalDate")))-1911&Right("00"&month(trim(rsfound("IllegalDate"))),2)&Right("00"&day(trim(rsfound("IllegalDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
			response.write Right("00"&hour(trim(rsfound("IllegalDate"))),2)&Right("00"&minute(trim(rsfound("IllegalDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("BillUnitID"))<>"" and not isnull(rsfound("BillUnitID")) then
			strUnit="select UnitName from UnitInfo where UnitID='"&trim(rsfound("BillUnitID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write trim(rsUnit("UnitName"))
			end if
			rsUnit.close
			set rsUnit=nothing
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("BillMemID1"))<>"" and not isnull(rsfound("BillMemID1")) then
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID1"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("LoginId"))&"&nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem1"))
		end If 
		
		if trim(rsfound("BillMemID2"))<>"" and not isnull(rsfound("BillMemID2")) then
			Response.Write ","
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID2"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("LoginId"))&"&nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem2"))
		end If 
		
		if trim(rsfound("BillMemID3"))<>"" and not isnull(rsfound("BillMemID3")) then
			Response.Write ","
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID3"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("LoginId"))&"&nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem3"))
		end if
		%></td>
		<td><%
		strFast="select ConfiscateID,Confiscate from PasserConfiscate" &_
			" where BillSN="&trim(rsfound("SN"))
		set rsFast=conn.execute(strFast)
		while Not rsFast.eof
			response.write trim(rsFast("Confiscate"))
		rsFast.movenext
		wend
		rsFast.close
		set rsFast=nothing
		%></td>
	</tr>
	<tr>
		<td><%
		if trim(rsfound("BillNo"))<>"" and not isnull(rsfound("BillNo")) then
			response.write trim(rsfound("BillNo"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td colspan="3"><%
		if trim(rsfound("illegalAddress"))<>"" and not isnull(rsfound("illegalAddress")) then
			response.write trim(rsfound("illegalAddress"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
			response.write trim(rsfound("Rule1"))
		else
			response.write "&nbsp;"
		end If 
		
		if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
			response.write "＼"&trim(rsfound("Rule2"))
		end if	
		%></td>
		<td><%
		if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
			strSQL="select Level1 from law where itemid='"&trim(rsfound("Rule1"))&"' and VerSion="&RuleVer
			set rslaw=conn.execute(strSQL)
			If not rslaw.eof Then Response.Write trim(rslaw("Level1"))
			rslaw.close
		else
			response.write "&nbsp;"
		end If 

		if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
			strSQL="select Level1 from law where itemid='"&trim(rsfound("Rule2"))&"' and VerSion="&RuleVer
			set rslaw=conn.execute(strSQL)
			If not rslaw.eof Then Response.Write "＼"&trim(rslaw("Level1"))
			rslaw.close
		end If 

		
		%></td>
	</tr>
	<tr>
		<td><%
		if trim(rsfound("BillFillDate"))<>"" and not isnull(rsfound("BillFillDate")) then
			response.write year(trim(rsfound("BillFillDate")))-1911&Right("00"&month(trim(rsfound("BillFillDate"))),2)&Right("00"&day(trim(rsfound("BillFillDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
			response.write year(trim(rsfound("DeallineDate")))-1911&Right("00"&month(trim(rsfound("DeallineDate"))),2)&Right("00"&day(trim(rsfound("DeallineDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("Driver"))<>"" then
			response.write trim(rsfound("Driver"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("DriverID"))<>"" and not isnull(rsfound("DriverID")) then
			response.write trim(rsfound("DriverID"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td colspan="2"><%
		if trim(rsfound("MemberStation"))<>"" and not isnull(rsfound("MemberStation")) then
			response.write trim(rsfound("MemberStation"))&"&nbsp; &nbsp; "
			strStation="select UnitName from UnitInfo where UnitID='"&trim(trim(rsfound("MemberStation")))&"'"
			set rsStation=conn.execute(strStation)
			if not rsStation.eof then
				response.write trim(rsStation("UnitName"))
			end if
			rsStation.close
			set rsStation=nothing
		else
			response.write "&nbsp;"
		end if
		%></td>
	</tr>
	
<%		CaseSN=CaseSN+1
	rsfound.MoveNext
	next
	Response.Write "<tr><td colspan=""6""><hr></td></tr></table>"
Wend
%>


共計:   <%=CaseSN%>  筆
<%if sys_City="苗栗縣" then%> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 分局承辦人&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 委外廠商
<%end if%>
<%
'一式兩份(第二聯)================================================================================================
if sys_City="苗栗縣" then
%>
<div class="PageNext">&nbsp;</div>
<%
CaseSN=0
If Not rsfound.Bof Then rsfound.MoveFirst 
While Not rsfound.Eof
	if CaseSN>0 then response.write "<div class=""PageNext"">&nbsp;</div>"
%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td colspan="6" align="center">
			<font size="3">行人攤販資料建檔清冊</font>
		</td>
	</tr>
	<tr>
		<td colspan="3" align="left">列印日期:<%=year(now)-1911&"/"&Right("00"&month(now),2)&"/"&Right("00"&day(now),2)%></td>
		<td colspan="3" align="right">Page <%=fix(CaseSN/14)+1%> of <%=pagecnt%> </td>
	</tr>
	<tr>
	<%if sys_City="苗栗縣" then%>
		<td colspan="3" align="left" colspan="2">登入者:<%=Session("Ch_Name")%></td>
		<td colspan="3" align="right">1 式 2 聯(第 2 聯：委外留存)</td>
	<%else%>
		<td colspan="6" align="left" colspan="2">登入者:<%=Session("Ch_Name")%></td>
	<%end if %>
	</tr>
	<tr><td colspan="6"><hr></td></tr>
	<tr>
		<td width="10%">建檔日期</td>
		<td width="15%">違規日期</td>
		<td width="20%">違規時間</td>
		<td width="15%">舉發單位</td>
		<td width="25%">舉發員警</td>
		<td width="15%">扣件</td>
	</tr>
	<tr>
		<td>舉發單號</td>
		<td colspan="3">違規地點</td>
		<td>法條一＼法條二</td>
		<td>罰鍰一＼罰鍰二</td>
	</tr>
	<tr>
		<td>填單日期</td>
		<td>應到案日期</td>
		<td>駕駛人姓名</td>
		<td>駕駛人ID</td>
		<td colspan="2">到案處所</td>
	</tr>
<%
for i=1 to 14
	if rsfound.eof then exit for
%>
	<tr><td colspan="6"><hr></td></tr>
	<tr>
		<td><%=year(trim(rsfound("RecordDate")))-1911&Right("00"&month(trim(rsfound("RecordDate"))),2)&Right("00"&day(trim(rsfound("RecordDate"))),2)%></td>
		<td width="9%"><%
		if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
			response.write year(trim(rsfound("IllegalDate")))-1911&Right("00"&month(trim(rsfound("IllegalDate"))),2)&Right("00"&day(trim(rsfound("IllegalDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
			response.write Right("00"&hour(trim(rsfound("IllegalDate"))),2)&Right("00"&minute(trim(rsfound("IllegalDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("BillUnitID"))<>"" and not isnull(rsfound("BillUnitID")) then
			strUnit="select UnitName from UnitInfo where UnitID='"&trim(rsfound("BillUnitID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write trim(rsUnit("UnitName"))
			end if
			rsUnit.close
			set rsUnit=nothing
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("BillMemID1"))<>"" and not isnull(rsfound("BillMemID1")) then
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID1"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("LoginId"))&"&nbsp;  &nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem1"))
		end If
		
		if trim(rsfound("BillMemID2"))<>"" and not isnull(rsfound("BillMemID2")) then
			Response.Write ","
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID2"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("LoginId"))&"&nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem2"))
		end If 
		
		if trim(rsfound("BillMemID3"))<>"" and not isnull(rsfound("BillMemID3")) then
			Response.Write ","
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID3"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("LoginId"))&"&nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem3"))
		end if
		%></td>
		<td><%
		strFast="select ConfiscateID,Confiscate from PasserConfiscate" &_
			" where BillSN="&trim(rsfound("SN"))
		set rsFast=conn.execute(strFast)
		while Not rsFast.eof
			response.write trim(rsFast("Confiscate"))
		rsFast.movenext
		wend
		rsFast.close
		set rsFast=nothing
		%></td>
	</tr>
	<tr>
		<td><%
		if trim(rsfound("BillNo"))<>"" and not isnull(rsfound("BillNo")) then
			response.write trim(rsfound("BillNo"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td colspan="3"><%
		if trim(rsfound("illegalAddress"))<>"" and not isnull(rsfound("illegalAddress")) then
			response.write trim(rsfound("illegalAddress"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
			response.write trim(rsfound("Rule1"))
		else
			response.write "&nbsp;"
		end If 
		
		if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
			response.write "＼"&trim(rsfound("Rule2"))
		end if	
		%></td>
		<td><%
		if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
			strSQL="select Level1 from law where itemid='"&trim(rsfound("Rule1"))&"' and VerSion="&RuleVer
			set rslaw=conn.execute(strSQL)
			If not rslaw.eof Then Response.Write trim(rslaw("Level1"))
			rslaw.close
		else
			response.write "&nbsp;"
		end If 

		if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
			strSQL="select Level1 from law where itemid='"&trim(rsfound("Rule2"))&"' and VerSion="&RuleVer
			set rslaw=conn.execute(strSQL)
			If not rslaw.eof Then Response.Write "＼"&trim(rslaw("Level1"))
			rslaw.close
		end if
		
		
		%></td>
	</tr>
	<tr>
		<td><%
		if trim(rsfound("BillFillDate"))<>"" and not isnull(rsfound("BillFillDate")) then
			response.write year(trim(rsfound("BillFillDate")))-1911&Right("00"&month(trim(rsfound("BillFillDate"))),2)&Right("00"&day(trim(rsfound("BillFillDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
			response.write year(trim(rsfound("DeallineDate")))-1911&Right("00"&month(trim(rsfound("DeallineDate"))),2)&Right("00"&day(trim(rsfound("DeallineDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("Driver"))<>"" then
			response.write trim(rsfound("Driver"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td><%
		if trim(rsfound("DriverID"))<>"" and not isnull(rsfound("DriverID")) then
			response.write trim(rsfound("DriverID"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td colspan="2"><%
		if trim(rsfound("MemberStation"))<>"" and not isnull(rsfound("MemberStation")) then
			response.write trim(rsfound("MemberStation"))&"&nbsp; &nbsp; "
			strStation="select UnitName from UnitInfo where UnitID='"&trim(trim(rsfound("MemberStation")))&"'"
			set rsStation=conn.execute(strStation)
			if not rsStation.eof then
				response.write trim(rsStation("UnitName"))
			end if
			rsStation.close
			set rsStation=nothing
		else
			response.write "&nbsp;"
		end if
		%></td>
	</tr>
	
<%		CaseSN=CaseSN+1
	rsfound.MoveNext
	next

	Response.Write "<tr><td colspan=""6""><hr></td></tr></table>"
Wend

rsfound.close
set rsfound=nothing
%>
共計:   <%=CaseSN%>  筆
	<%if sys_City="苗栗縣" then%> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 分局承辦人&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 委外廠商
	<%end if%>
<%end if%>
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
window.print();

</script>
<%
conn.close
set conn=nothing
%>