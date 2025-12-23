<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
%>
<html>
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.pageprint {
  margin-left: 7mm;
  margin-right: 5.08mm;
  margin-top: 5.08mm;
  margin-bottom: 5.08mm;
}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>逕舉(照片)建檔資料清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)

RecordDate=split(gInitDT(date),"-")
	if trim(request("CallType"))="1" then
		strwhere=" and a.BillStatus in ('0') and a.RecordStateID=0 and a.RecordMemberID="&session("User_ID")
	elseif trim(request("CallType"))="99" Then
		strwhere=" and a.BillStatus in ('0') and a.RecordStateID=0 and a.RecordMemberID="&session("User_ID") &_
			" and JurgeDay is not null"
	elseif trim(request("CallType"))="88" Then '民眾檢舉建檔
		strwhere=" and a.BillStatus in ('1') and a.RecordStateID=0 and a.RecordMemberID="&session("User_ID") &_
			" and JurgeDay is not null"
	elseif trim(request("CallType"))="77" Then	'彰化員警舉發建檔
		strwhere=" and a.BillStatus in ('1') and a.RecordStateID=0 and a.RecordMemberID="&session("User_ID") &_
			" and JurgeDay is null"
	else
		strwhere=Session("PrintCarDataSQL")	
	end if
	'Session.Contents.Remove("PrintCaseDataSQLxls")
	'Session("PrintCaseDataSQLxls")=strwhere	
	
	BillBaseType=""
	if trim(request("CallType"))="88" Or trim(request("CallType"))="77" Then
		BillBaseType="BillBaseTmp"
	Else
		BillBaseType="BillBase"
	End If 

	strCnt="select count(*) as cnt from "&BillBaseType&" a,MemberData b where a.BillTypeID='2' and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by a.RecordDate"
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		if trim(rsCnt("cnt"))="0" then
			pagecnt=1
		else
			pagecnt=fix(Cint(rsCnt("cnt"))/50+0.9999999)
		end if
	end if
	rsCnt.close
	set rsCnt=nothing

	strSQL="select a.RuleSPeed,a.IllegalSpeed , a.SN,a.BillNo,a.CarNo,a.CarSimpleID,a.IllegalDate,a.IllegalAddress,a.BillUnitID,a.DriverID,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.BillFillDate,a.DeallineDate,a.MemberStation,a.equipmentID,a.JurgeDay from "&BillBaseType&" a,MemberData b where a.BillTypeID='2' and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by a.RecordDate"
	set rsfound=conn.execute(strSQL)

	tmpSQL=strwhere
'response.write strSQL
%>

</head>
<body>
<form name="myForm" method="post">
<%
CaseSN=0
If Not rsfound.Bof Then rsfound.MoveFirst 
While Not rsfound.Eof
	if CaseSN>0 then response.write "<div class=""PageNext""></div>"
%>
<table width="<%
If sys_City="苗栗縣" Then
	response.write "1050" 
ElseIf sys_City="台中市" Then
	response.write "100%" 
Else 
	response.write "700" 
End if %>" border="0" cellpadding="1" cellspacing="0">
	<tr>
		<td colspan="2" align="center">
			<font size="3">逕舉(<%
		If trim(request("CallType"))="99" Then
			response.write "民眾檢舉"
		Else
			response.write "照片"
		End If 
			%>)資料建檔清冊</font>
		</td>
	</tr>
	<tr>
		<td align="left"></td>
		<td align="right">Page <%=fix(CaseSN/50)+1%> of <%=pagecnt%> </td>
	</tr>
	<tr>
		<td align="left">登入者:<%=Session("Ch_Name")%></td>
		<td align="right">列印日期:<%=year(now)-1911&"/"&Right("00"&month(now),2)&"/"&Right("00"&day(now),2)%></td>
	</tr>
</table>
	<table width="<%
	If sys_City="苗栗縣" then 
		response.write "1050" 
	ElseIf sys_City="台中市" Then
		response.write "100%" 
	Else 
		response.write "700" 
	End if %>" border="1" cellpadding="0" cellspacing="0">
	<tr>
		<td>
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="5%">編號</td>
			<td width="8%">登入日期</td>
			<td width="10%">車號</td>
			<td width="5%">車種</td>
			<td width="8%">違規日期</td>
			<td width="7%">違規時間</td>
			<td width="<%If sys_City="苗栗縣" then response.write "6%" Else response.write "8%" End if %>">法條一</td>
			<td width="<%If sys_City="苗栗縣" then response.write "6%" Else response.write "8%" End if %>">法條二</td>
			<td width="<%If sys_City="苗栗縣" then response.write "19%" Else response.write "9%" End if %>">舉發員警</td>
			<td width="<%If sys_City="苗栗縣" Or sys_City="台中市" then response.write "19%" Else response.write "23%" End if %>">違規地點 (限速,車速)</td>
		<%if sys_City="台中縣" then%>
			<td width="7%">應到案日期</td>
		<%elseif sys_City="屏東縣" then%>
			<td width="9%">檢舉日期</td>
		<%elseif sys_City="台中市" then%>
			<td width="11%">舉發單位</td>
		<%end if%>
		</tr>
<%if sys_City="基隆市" then%>
		<tr>
			<td ></td>
			<td colspan="2">填單日期</td>
			<td colspan="2">應到案日期</td>
			<td colspan="2">是否郵寄</td>
		</tr>
<%end if%>
		</table>
		</td>
	</tr>
<%
if sys_City="基隆市" Or sys_City="苗栗縣" Then
	printCnt=25
ElseIf sys_City="台南市" Then
	printCnt=25
Else
	printCnt=50
End If

for i=1 to printCnt
	if rsfound.eof then exit for
	CaseSN=CaseSN+1
%>
	<tr>
		<td>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
		<td width="5%"><%=CaseSN%><%
	if sys_City="台南市" Then
		Response.write "<br>&nbsp;"
	End If 
		%></td>
		<td width="8%"><%=year(now)-1911&Right("00"&month(now),2)&Right("00"&day(now),2)%></td>
		<td width="10%"><%
		if trim(rsfound("BillNO"))<>"" then
			response.write trim(rsfound("BillNO"))&"<br>"
		end if
		response.write trim(rsfound("CarNo"))
		%></td>
		<td width="5%"><%
	if sys_City="苗栗縣" Then
		if trim(rsfound("CarSimpleID"))="1" then
			response.write "汽車"
		elseif trim(rsfound("CarSimpleID"))="2" then
			response.write "拖車"
		elseif trim(rsfound("CarSimpleID"))="3" then
			response.write "重機"
		elseif trim(rsfound("CarSimpleID"))="4" then
			response.write "輕機"
		elseif trim(rsfound("CarSimpleID"))="6" then
			response.write "臨時車牌"
		end if 
		
	else
		response.write trim(rsfound("CarSimpleID"))
	end if 
		%></td>
		<td width="8%"><%
		if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
			response.write year(trim(rsfound("IllegalDate")))-1911&Right("00"&month(trim(rsfound("IllegalDate"))),2)&Right("00"&day(trim(rsfound("IllegalDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="7%"><%
		if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
			response.write Right("00"&hour(trim(rsfound("IllegalDate"))),2)&Right("00"&minute(trim(rsfound("IllegalDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="<%If sys_City="苗栗縣" then response.write "6%" Else response.write "8%" End if %>"><%
		if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
			response.write trim(rsfound("Rule1"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="<%If sys_City="苗栗縣" then response.write "6%" Else response.write "8%" End if %>"><%
		if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
			response.write trim(rsfound("Rule2"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="<%If sys_City="苗栗縣" then response.write "19%" Else response.write "9%" End if %>"><%
	If sys_City="苗栗縣" Then
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
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID2"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write "/"&trim(rsMem("LoginId"))&"&nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem2"))
		end If
		if trim(rsfound("BillMemID3"))<>"" and not isnull(rsfound("BillMemID3")) then
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID3"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write "<br>"&trim(rsMem("LoginId"))&"&nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem3"))
		end If
		if trim(rsfound("BillMemID4"))<>"" and not isnull(rsfound("BillMemID4")) then
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID4"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write "/"&trim(rsMem("LoginId"))&"&nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem4"))
		end if
	Else
		if trim(rsfound("BillMem1"))<>"" and not isnull(rsfound("BillMem1")) then
			response.write trim(rsfound("BillMem1"))
		else
			response.write "&nbsp;"
		end if
		if sys_City="台中市" then
			if trim(rsfound("BillMem2"))<>"" and not isnull(rsfound("BillMem2")) then
				response.write "/"&trim(rsfound("BillMem2"))
			end if		
		end if
	End If
		%></td>
		<td width="<%
		if sys_City="苗栗縣" and trim(rsfound("IllegalSPeed"))<>"" Then
			response.write "15%" 
		elseif sys_City="苗栗縣" Or sys_City="台中市" Then
			response.write "19%" 
		else 
			response.write "23%" 
		end if%>"><%
		if trim(rsfound("illegalAddress"))<>"" and not isnull(rsfound("illegalAddress")) then
			response.write trim(rsfound("illegalAddress"))
		else
			response.write "&nbsp;"
		end if
	if sys_City<>"苗栗縣" then
		if trim(rsfound("IllegalSPeed"))<>"" then
			response.write " ( " & rsfound("RuleSPeed")& "," & rsfound("IllegalSPeed") & " ) "
		end if
	end if
		%></td>
	<%if sys_City="苗栗縣" and trim(rsfound("IllegalSPeed"))<>"" then%>
		<td width="4%"><%
		if trim(rsfound("IllegalSPeed"))<>"" then
			response.write " ( " & rsfound("RuleSPeed")& "," & rsfound("IllegalSPeed") & " ) "
		end if
		%></td>
	<%end if%>
	<%if sys_City="台中縣" then%>
		<td width="7%"><%
		if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
			response.write year(trim(rsfound("DeallineDate")))-1911&Right("00"&month(trim(rsfound("DeallineDate"))),2)&Right("00"&day(trim(rsfound("DeallineDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
	<%elseif sys_City="屏東縣" then%>
		<td width="9%"><%
		if trim(rsfound("JurgeDay"))<>"" and not isnull(rsfound("JurgeDay")) then
			response.write year(trim(rsfound("JurgeDay")))-1911&Right("00"&month(trim(rsfound("JurgeDay"))),2)&Right("00"&day(trim(rsfound("JurgeDay"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
	<%elseif sys_City="台中市" then%>
		<td width="11%"><%
		if trim(rsfound("BillUnitID"))<>"" and not isnull(rsfound("BillUnitID")) then
			strU="select * from UnitInfo where unitid='"&trim(rsfound("BillUnitID"))&"'"
			Set rsU=conn.execute(strU)
			If Not rsU.eof Then
				response.write rsU("UnitName")
			End If
			rsU.close
			Set rsU=Nothing 
		else
			response.write "&nbsp;"
		end if
		%></td>
	<%end if%>
		</tr>
<%if sys_City="基隆市" then%>
		<tr>
			<td ></td>
			<td colspan="2"><%
		if trim(rsfound("BillFillDate"))<>"" and not isnull(rsfound("BillFillDate")) then
			response.write year(trim(rsfound("BillFillDate")))-1911&Right("00"&month(trim(rsfound("BillFillDate"))),2)&Right("00"&day(trim(rsfound("BillFillDate"))),2)
		else
			response.write "&nbsp;"
		end if
			%></td>
			<td colspan="2"><%
		if trim(rsfound("DeallineDate"))<>"" and not isnull(rsfound("DeallineDate")) then
			response.write year(trim(rsfound("DeallineDate")))-1911&Right("00"&month(trim(rsfound("DeallineDate"))),2)&Right("00"&day(trim(rsfound("DeallineDate"))),2)
		else
			response.write "&nbsp;"
		end if
			%></td>
			<td colspan="2"><%'是否郵寄
		if trim(rsfound("equipmentID"))<>"" and not isnull(rsfound("equipmentID")) Then
			If trim(rsfound("equipmentID"))="1" Then
				response.write "是"
			Else
				response.write "否"
			End if
			
		else
			response.write "&nbsp;"
		end if	
		%></td>
		</tr>
<%end if%>
	</table>
		</td>
	</tr>

<%		
	rsfound.MoveNext
	next
%>
	</table>
<%
Wend
rsfound.close
set rsfound=nothing
%>
共計:   <%=CaseSN%>  筆

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