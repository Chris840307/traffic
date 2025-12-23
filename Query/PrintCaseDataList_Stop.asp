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
  margin-left: 5.08mm;
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
<title>攔停資料建檔清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
Server.ScriptTimeout = 1800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)

RecordDate=split(gInitDT(date),"-")
	if trim(request("CallType"))="1" then
		strwhere=" and a.papercheck=1 and a.BillStatus in ('0') and a.RecordStateID=0 and a.RecordMemberID="&session("User_ID")
	else
		strwhere=Session("PrintCarDataSQL")	
	end if
	'Session.Contents.Remove("PrintCaseDataSQLxls")
	'Session("PrintCaseDataSQLxls")=strwhere	

	strCnt="select count(*) as cnt from BillBase a,MemberData b where a.BillTypeID<>'2' and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by a.RecordDate"
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

	strSQL="select a.RecordDate,a.SN,a.BillNo,a.CarNo,a.CarSimpleID,a.IllegalDate,a.IllegalAddress,a.BillUnitID,a.DriverBirth,a.DriverID,a.CarAddID,a.equipmentID,SignType,a.SignType,a.BillMemID1,a.BillMem1,a.BillMemID2,a.BillMem2,a.BillMemID3,a.BillMem3,a.BillMemID4,a.BillMem4,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillFillDate,a.DeallineDate,a.MemberStation from BillBase a,MemberData b where a.BillTypeID<>'2' and a.RecordMemberID=b.MemberID(+) "&strwhere&" order by a.RecordDate"
	set rsfound=conn.execute(strSQL)

	tmpSQL=strwhere

%>

</head>
<body>
<form name=myForm method="post">
<%
CaseSN=0
If Not rsfound.Bof Then rsfound.MoveFirst 
While Not rsfound.Eof
	if CaseSN>0 then response.write "<div class=""PageNext""></div>"
%>
<table width="700" border="0" cellpadding="1" cellspacing="0">
	<tr>
		<td colspan="2" align="center">
			<font size="3">攔停資料建檔清冊</font>
		</td>
	</tr>
	<tr>
		<td align="left">列印日期:<%=year(now)-1911&"/"&Right("00"&month(now),2)&"/"&Right("00"&day(now),2)%></td>
		<td align="right">Page <%=fix(CaseSN/14)+1%> of <%=pagecnt%> </td>
	</tr>
	<tr>
		<td align="left" colspan="2">登入者:<%=Session("Ch_Name")%></td>
	</tr>
</table>
	<hr>
	<table width="700" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="9%">建檔日期</td>
		<td width="11%">單號</td>
		<td width="12%">車號</td>
		<td width="9%">簡式車種</td>
		<td width="9%">違規日期</td>
		<td width="9%">違規時間</td>
		<td width="17%">舉發單位</td>
		<td width="15%">舉發員警</td>
		<td width="9%">扣件</td>
	</tr>
	<tr>
		<td colspan="6">違規地點</td>
		<td colspan="3">法條一</td>
	</tr>
	<tr>
		<td>填單日期</td>
		<td>應到案日期</td>
		<td>駕駛人ID</td>
		<td colspan="3">到案處所</td>
		<td colspan="3">法條二</td>
	</tr>
<%
if sys_City="基隆市" Or sys_City="苗栗縣" Then
%>
	<tr>
		<td colspan="2">駕駛人生日</td>
		<td>是否郵寄</td>
		
	<%If sys_City="苗栗縣" then%>
		<td>簽收狀況</td>
		<td colspan="2">砂石車註記</td>
		<td colspan="3">法條三</td>
	<%else%>
		<td colspan="6">簽收狀況</td>
	<%End if%>
	</tr>
<%
End if
%>
	</table>
	<hr>
<%
if sys_City="高雄市" then
	printCnt=13
elseif sys_City="基隆市" Or sys_City="苗栗縣" then
	printCnt=10
else
	printCnt=14
end if

for i=1 to printCnt
	if rsfound.eof then exit for
%>
	<table width="700" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="9%"><%=year(trim(rsfound("RecordDate")))-1911&Right("00"&month(trim(rsfound("RecordDate"))),2)&Right("00"&day(trim(rsfound("RecordDate"))),2)%></td>
		<td width="11%"><%=trim(rsfound("BillNo"))%></td>
		<td width="12%"><%=trim(rsfound("CarNo"))%></td>
		<td width="9%"><%=trim(rsfound("CarSimpleID"))%></td>
		<td width="9%"><%
		if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
			response.write year(trim(rsfound("IllegalDate")))-1911&Right("00"&month(trim(rsfound("IllegalDate"))),2)&Right("00"&day(trim(rsfound("IllegalDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="9%"><%
		if trim(rsfound("IllegalDate"))<>"" and not isnull(rsfound("IllegalDate")) then
			response.write Right("00"&hour(trim(rsfound("IllegalDate"))),2)&Right("00"&minute(trim(rsfound("IllegalDate"))),2)
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="17%"><%
		if trim(rsfound("BillUnitID"))<>"" and not isnull(rsfound("BillUnitID")) then
			strUnit="select UnitName from UnitInfo where UnitID='"&trim(rsfound("BillUnitID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				unitname=trim(rsUnit("UnitName"))
				'三星分局太長mark掉
				if trim(rsUnit("UnitName"))<>"三星分局" then 
					finalunitname=REPLACE(unitname,"三星分局","")
				else
					finalunitname=unitname
				end if				
				response.write finalunitname
			end if
			rsUnit.close
			set rsUnit=nothing
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td width="15%"><%
		if trim(rsfound("BillMemID1"))<>"" and not isnull(rsfound("BillMemID1")) then
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID1"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("LoginId"))&"&nbsp;  &nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem1"))
		end if
		%></td>
		<td width="9%"><%
		strFast="select b.ID,b.Content from BillFastenerDetail a,DciCode b" &_
			" where a.BillSN="&trim(rsfound("SN"))&" and b.TypeID=6 and a.FastenerTypeID=b.ID"
		set rsFast=conn.execute(strFast)
		while Not rsFast.eof
			response.write trim(rsFast("ID"))&trim(rsFast("Content"))
		rsFast.movenext
		wend
		rsFast.close
		set rsFast=nothing
		%></td>
	</tr>
	<tr>
		<td colspan="6"><%
		if trim(rsfound("illegalAddress"))<>"" and not isnull(rsfound("illegalAddress")) then
			response.write trim(rsfound("illegalAddress"))
		else
			response.write "&nbsp;"
		end if
		%></td>
	<%if sys_City="苗栗縣" Then%>
		<td><%
		if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
			response.write trim(rsfound("Rule1"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td colspan="2"><%
		if trim(rsfound("BillMemID2"))<>"" and not isnull(rsfound("BillMemID2")) then
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID2"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("LoginId"))&"&nbsp;  &nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem2"))
		end if
		%></td>
	<%else%>
		<td colspan="3"><%
		if trim(rsfound("Rule1"))<>"" and not isnull(rsfound("Rule1")) then
			response.write trim(rsfound("Rule1"))
		else
			response.write "&nbsp;"
		end if
		if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
			response.write "&nbsp; &nbsp;"&trim(rsfound("Rule3"))
		end if
		%></td>
	<%End If %>
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
		if trim(rsfound("DriverID"))<>"" and not isnull(rsfound("DriverID")) then
			response.write trim(rsfound("DriverID"))
		else
			response.write "&nbsp;"
		end if
		%></td>
		<td colspan="3"><%
		if trim(rsfound("MemberStation"))<>"" and not isnull(rsfound("MemberStation")) then
			response.write trim(rsfound("MemberStation"))&"&nbsp; &nbsp; "
			strStation="select DciStationName from Station where DciStationID='"&trim(trim(rsfound("MemberStation")))&"'"
			set rsStation=conn.execute(strStation)
			if not rsStation.eof then
				response.write trim(rsStation("DciStationName"))
			end if
			rsStation.close
			set rsStation=nothing
		else
			response.write "&nbsp;"
		end if
		%></td>
	<%if sys_City="苗栗縣" Then%>
		<td ><%
		if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
			response.write trim(rsfound("Rule2"))
		else
			response.write "&nbsp;"
		end if		
		%></td>
		<td colspan="2"><%
		if trim(rsfound("BillMemID3"))<>"" and not isnull(rsfound("BillMemID3")) then
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID3"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("LoginId"))&"&nbsp;  &nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem3"))
		end if
		%></td>
	<%else%>
		<td colspan="3"><%
		if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
			response.write trim(rsfound("Rule2"))
		else
			response.write "&nbsp;"
		end if		
		if trim(rsfound("Rule4"))<>"" and not isnull(rsfound("Rule4")) then
			response.write "&nbsp; &nbsp;"&trim(rsfound("Rule4"))
		end if
		%></td>
	<%End if%>
	</tr>
<%
if sys_City="基隆市" Or sys_City="苗栗縣" Then
%>
	<tr>
		<td colspan="2"><%'駕駛人生日
		if trim(rsfound("DriverBirth"))<>"" and not isnull(rsfound("DriverBirth")) Then
			response.write year(trim(rsfound("DriverBirth")))-1911&Right("00"&month(trim(rsfound("DriverBirth"))),2)&Right("00"&day(trim(rsfound("DriverBirth"))),2)
		else
			response.write "&nbsp;"
		end if	
		%></td>
		<td><%'是否郵寄
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
	<%if sys_City="苗栗縣" Then%>
		<td ><%'簽收狀況
		if trim(rsfound("SignType"))<>"" and not isnull(rsfound("SignType")) then
			If trim(rsfound("SignType"))="A" Then
				response.write "簽收"
			Else
				response.write "拒簽收"
			End if
			
		else
			response.write "&nbsp;"
		end if	
		%></td>
		<td colspan="2"><%'砂石車註記
		if trim(rsfound("CarAddID"))="3" then
			response.write "砂石車"
		else
			response.write "&nbsp;"
		end if	
		%></td>
		<td ><%'法條三
		if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
			response.write trim(rsfound("Rule3"))
		else
			response.write "&nbsp;"
		end if		
		%></td>
		<td colspan="2"><%
		if trim(rsfound("BillMemID4"))<>"" and not isnull(rsfound("BillMemID4")) then
			strMem="select LoginId from MemberData where MemberID="&trim(rsfound("BillMemID4"))
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("LoginId"))&"&nbsp;  &nbsp;"
			end if
			rsMem.close
			set rsMem=nothing
			response.write trim(rsfound("BillMem4"))
		end if
		%></td>
	<%else%>
		<td colspan="3"><%'簽收狀況
		if trim(rsfound("SignType"))<>"" and not isnull(rsfound("SignType")) then
			If trim(rsfound("SignType"))="A" Then
				response.write "簽收"
			Else
				response.write "拒簽收"
			End if
			
		else
			response.write "&nbsp;"
		end if	
		%></td>
	<%End If %>
	</tr>
<%
End if
%>
	</table>
	<hr>
<%		CaseSN=CaseSN+1
	rsfound.MoveNext
	next
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
<%
if sys_City<>"基隆市" and trim(Session("User_ID"))<>"10000" then
%>
	window.print();
<%
end if
%>
</script>
<%
conn.close
set conn=nothing
%>