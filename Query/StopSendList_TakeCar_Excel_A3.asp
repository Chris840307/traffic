<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://localhost/traffic/smsx.cab#Version=6,1,432,1">
</object>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
.style2 {font-family:新細明體; color=0044ff; line-height:23px; font-size: 20px}
.style3 {font-family:標楷體; color=0044ff; line-height:18px; font-size: 16px}
.style5 {font-family:新細明體; color=0044ff; line-height:15px; font-size: 12px}
.style6 {font-family:標楷體; color=0044ff; line-height:18px; font-size: 18px}
-->
</style>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>拖吊移送清冊</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<%
Server.ScriptTimeout = 800
Response.flush
%>
<%
'權限
'AuthorityCheck(234)

 'and a.BillTypeID<>'2'
  strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=Nothing
sys_City="高雄市"
%>
<%
'設定固定人員


		strSQL="select memberid from memberdata where loginid='"&request("loginid")&"' and recordstateid=0 and accountstateid=0"
	set rsfound=conn.execute(strSQL)
	If Not rsfound.eof Then 
		RecordMemberID=rsfound("memberid")
	Else
		response.write "輸入人員臂章號碼有誤"
		response.end
	End If


	'頁數
	PageNum=1
	StationArrayTemp=""
	strwhere=" and f.RecordDate between TO_DATE('"&gOutDT(request("StartDate"))&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&gOutDT(request("EndDate"))&" 23:59:59','YYYY/MM/DD/HH24/MI/SS')"

	StationArrayTemp="32"

	strCnt="select count(*) as cnt  from BillBase f where f.BillStatus=9 and f.RecordStateID=0 "&strwhere&" and f.RecordMemberID="&RecordMemberID
	set rsCnt=conn.execute(strCnt)
	if not rsCnt.eof then
		DBcnt=rsCnt("Cnt")
	end if
	rsCnt.close
	set rsCnt=Nothing
	DBcnt=1
%>
</head>
<body>
<form name=myForm method="post">

<%
	strUnitName2="select UnitName from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
	set rsUnitName2=conn.execute(strUnitName2)
	if not rsUnitName2.eof then
		TitleUnitName2=trim(rsUnitName2("UnitName"))
	end if
	rsUnitName2.close
	set rsUnitName2=Nothing
	
	strUnitName2="select chname from MemberData where memberid='"&trim(Session("User_ID"))&"'"
	set rsUnitName2=conn.execute(strUnitName2)
	if not rsUnitName2.eof then
		chname=trim(rsUnitName2("chname"))
	end if
	rsUnitName2.close
	set rsUnitName2=nothing

	strUnitName="select Value from ApConfigure where ID=40"
	set rsUnitName=conn.execute(strUnitName)
	if not rsUnitName.eof then
		TitleUnitName=trim(rsUnitName("value"))&" "&TitleUnitName2
	end if
	rsUnitName.close
	set rsUnitName=nothing

	PrintSNtotal=0	'編號
if request("SNStart")="" and request("SNEnd")="" then 
PrintSN3=0
else
PrintSN3=cdbl(request("SNStart"))-1
end if

	'高雄市交通事件裁決所列表
	if instr(StationArrayTemp,"30")>0 or instr(StationArrayTemp,"31")>0 or instr(StationArrayTemp,"32")>0 then
		DciStationName="高雄市交通事件裁決所"

	
		strCnt="select count(*) as cnt  from BillBase f where f.BillStatus=9 and f.RecordStateID=0 "&strwhere&" and f.RecordMemberID="&RecordMemberID

		set rsCnt=conn.execute(strCnt)
		if not rsCnt.eof then
			if trim(rsCnt("cnt"))="0" then
				pagecnt=1
			else
				pagecnt=fix(Cint(rsCnt("cnt"))/30+0.9999999)
			end if
		end if
		rsCnt.close
		set rsCnt=nothing
		
if request("SNStart")<>"" and request("SNEnd")<>"" then 
		strSQL="select * from (select rownum  no,SN,BillNo,CarNo,CarSimpleID,Loginid,IllegalDate,IllegalAddress,RecordDate,ForFeit1,Rule1,Rule2,Rule3,Rule4,BillUnitID,DealLineDate,DriverID,Loginid2,Billmem1 from (select  f.SN,f.BillNo,f.CarNo,f.CarSimpleID,g.Loginid,f.IllegalDate,f.IllegalAddress,f.RecordDate,f.ForFeit1,f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.BillUnitID,f.DealLineDate,f.DriverID,h.Loginid as Loginid2,f.Billmem1 from BillBase f ,memberdata g,memberdata h where f.BillStatus=9 and f.RecordStateID=0 "&strwhere&" and f.RecordMemberID="&RecordMemberID&" and f.RecordMemberID=h.memberid and f.RecordMemberID=g.memberid and g.recordstateid=0 and g.accountstateid=0  order by f.RecordMemberID,f.RecordDate)) where no between "&request("SNStart")&" and "&request("SNEnd")&" order by no"
else
strSQL="select  f.SN,f.BillNo,f.CarNo,f.CarSimpleID,g.Loginid,f.IllegalDate,f.IllegalAddress,f.RecordDate,f.ForFeit1,f.Rule1,f.Rule2,f.Rule3,f.Rule4,f.BillUnitID,f.DealLineDate" &_
			",f.DriverID,h.Loginid as Loginid2,f.Billmem1 from BillBase f ,memberdata g,memberdata h where f.BillStatus=9 and f.RecordStateID=0 "&strwhere&" and f.RecordMemberID="&RecordMemberID&" and f.RecordMemberID=h.memberid and f.RecordMemberID=g.memberid and g.recordstateid=0 and g.accountstateid=0  order by f.RecordMemberID,f.RecordDate"
end if

		set rs1=conn.execute(strSQL)
ForFeit1=0
		If Not rs1.Bof Then rs1.MoveFirst 
		While Not rs1.Eof
		if PrintSN>0 then
%>
		<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
		If PageNum<>1 Then 	response.write "<div class=""PageNext"">&nbsp;</div>"				
	%></center>
<%

		end If

%>
	<table width="100%" border="0" cellpadding="1" cellspacing="0">
		<tr>
			<td align="center" colspan=16><span class="style2">高雄市政府警察局&nbsp;&nbsp;違規罰單入案系統<br>拖吊已結案件移送表</span></td>
		</tr>
		<tr>
<td align="left" width="80%"><span class="style6">到案地點&nbsp;:&nbsp;高市交通事件裁決中心</td>
<td><span class="style6" width="20%">頁次&nbsp;&nbsp;:&nbsp;&nbsp;<%=fix(PrintSN/30)+1%>&nbsp;of&nbsp;<%=pagecnt%></td>
<tr>

<td width="80%"><span class="style6" width="20%">告發單位&nbsp;:&nbsp;<%=TitleUnitName2%></td>
<td><span class="style6">列印日期 :&nbsp;<%=Right("00"&year(now)-1911,2)%>年<%=Right("00"&month(now),2)%>月<%=Right("00"&day(now),2)%>日</td>
<tr>

<td width="80%"><span class="style6" width="20%">移送批號&nbsp;:&nbsp;<%=Right("00"&year(rs1("RecordDate"))-1911,2)%><%=Right("00"&month(rs1("RecordDate")),2)%><%=Right("00"&day(rs1("RecordDate")),2)%></td>
<td><span class="style6">列印人員&nbsp;:&nbsp;<%=chname%></td>
<tr>

</span></td>
		</tr>
	</table>
	<table width="100%" border="0" cellpadding="1" cellspacing="0">
	<tr>
		<td colspan=17>=========================================================================================================================================================================================</td>
		<tr>
			<td width="2%" align=left><span class="style3">序號</td>
			<td width="3%" align=left><span class="style3">告發單單號</td>
			<td width="3%" align=left><span class="style3">車號</td>
			<td width="2%" align=left><span class="style3">車別</td>
			<td width="2%" align=right><span class="style3">違規日期</td>
			<td width="2%" align=center><span class="style3">時間</td>
			<td width="5%" align=left><span class="style3">違規地點</td>
			<td width="3%" align=left><span class="style3">違規條款</td>
			<td width="2%" align=left><span class="style3">入案</td>
			<td width="3%" align=left><span class="style3">舉發員警</td>
			<td width="3%" align=left><span class="style3">違規金額</td>
			<td width="3%" align=left><span class="style3">建檔日期</td>
		<tr>
		<td colspan=17>-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------</td>
		</tr>
	</tr>

<%		
sn=0
PrintSN2=0

ForFeit2=0
for i=1 to 30

			if rs1.eof then exit For
	
%>
	<tr>
	<td>

		<tr>
<%
			PrintSN=PrintSN+1
			PrintSN2=PrintSN2+1
			PrintSN3=PrintSN3+1			
			PrintSNtotal=PrintSNtotal+1
%>

			<td><span class="style3"><%
			tmp0=""
			For a=1 To 5-Len(PrintSN3) 
			  tmp0=tmp0&"0"
			Next
			response.write tmp0&PrintSN3
			%></span></td>

			<td align=left><span class="style3"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write rs1("BillNO")
			else
				response.write "&nbsp;"
			end if
			%></span></td>

			<td align=left><span class="style3">&nbsp;<%=trim(rs1("CarNo"))%></span></td>
			<td align=left><span class="style3">&nbsp;<%
			If trim(rs1("CarSimpleID"))="1" Then 
				response.write "汽車"
			ElseIf trim(rs1("CarSimpleID"))="2" Then 
				response.write "拖車"
			ElseIf trim(rs1("CarSimpleID"))="3" Then 
				response.write "重機"
			ElseIf trim(rs1("CarSimpleID"))="4" Then 
				response.write "輕機"
			ElseIf trim(rs1("CarSimpleID"))="6" Then 
				response.write "臨時車牌"
			End if
			%></span></td>

			<td align=right><span class="style3"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gInitDT(rs1("IllegalDate"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>

			<td align=center><span class="style3"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write Right("0"&Hour(rs1("IllegalDate")),2)&":"&Right("0"&minute(rs1("IllegalDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></span></td>

			<td align=left><span class="style3"><%
			if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
				response.write "&nbsp;"&trim(rs1("IllegalAddress"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>

			<td align=left><span class="style3"><%
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>

			<td align=left><span class="style3">OK</span></td>

			<td align=left><span class="style3"><%
			if (trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1"))) then
				response.write rs1("BillMem1")
			else
				response.write "&nbsp;"
			end if
			%></span></td>

			<td align=right><span class="style3"><%
			if (trim(rs1("ForFeit1"))<>"" and not isnull(rs1("ForFeit1"))) then
				response.write rs1("ForFeit1")&"&nbsp;&nbsp;"
				ForFeit1=ForFeit1+CDbl(rs1("ForFeit1"))
				ForFeit2=ForFeit2+CDbl(rs1("ForFeit1"))
			else
				response.write "&nbsp;"
			end if
			%></span></td>

			<td align=left><span class="style3"><%
			if (trim(rs1("RecordDate"))<>"" and not isnull(rs1("RecordDate"))) then
				response.write Right("00"&year(rs1("RecordDate"))-1911,2)&Right("00"&month(rs1("RecordDate")),2)&Right("00"&day(rs1("RecordDate")),2)
			else
				response.write "&nbsp;"
			end if
			%></span></td>
		</tr>

		<%If (PrintSN Mod 10=0) And PrintSN<>1 Then %>
			<td colspan=17>-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------</td>
		<tr>
		<%End if%>		

		</td>
		</tr>
<%			
		rs1.MoveNext

		Next

%>
<table border=0 width="100%">
<td width="90%">
	小計：&nbsp;<%=PrintSN2%>  &nbsp;筆
</td>
<td width="10%">
	金額:&nbsp;<%=ForFeit2%>
</td>
</table>

	</table>
<%

		Wend
		rs1.close
		set rs1=nothing
%>


<table border=0 width="100%">
<td colspan=2>
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
</td>
<tr>
<td width="90%">
	總計：&nbsp;<%=PrintSN%>  &nbsp;筆
</td>
<td width="10%">
	總金額:&nbsp;<%=ForFeit1%>
</td>
</table>
<br>
	<center><%
	response.write "<span class=""style5"">第"&PageNum&"頁</span>"
	PageNum=PageNum+1
	%></center>
<%end if%>
</form>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}

printWindow(true,7,5.08,5.08,5.08);
</script>
<%conn.close%>