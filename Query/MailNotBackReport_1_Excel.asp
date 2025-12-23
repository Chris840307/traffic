<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

chkDate=trim(request("chkDate"))
strDate=split("BillFillDate,IllegalDate,RecordDate",",")
strDateName=split("填單日,違規日,建檔日",",")
UserId = Session("User_ID")
startDate_q = Trim(Request("startDate_q"))
endDate_q = Trim(Request("endDate_q"))
unit = Request("unit")
UnitID_q = Request("UnitID_q")
unitList=trim(request("unitSelectlist"))
Batchnumber_q=trim(request("Batchnumber"))
Memlist_q=trim(request("MemSelectlist"))
Server.ScriptTimeout=86400



strUInfo="select * from Apconfigure where ID=40"
set rsUInfo=conn.execute(strUInfo)
if not rsUInfo.eof then thenPasserCity=trim(rsUInfo("value"))
rsUInfo.close

tmpSql=""
'統計日期
if startDate_q<>"" then
	tmpSql = tmpSql & " and "&strDate(chkDate)&" Between To_Date('" & gOutDT(startDate_q)&" 0:0:0" & "','YYYY/MM/DD/HH24/MI/SS') And To_Date('" & gOutDT(endDate_q)&" 23:59:59" & "','YYYY/MM/DD/HH24/MI/SS')"
end if

P_UnitName=thenPasserCity
strSQL="select UnitName from UnitInfo where UnitID='"&UnitID_q&"'"

set rsunit=conn.execute(strSQL)
If Not rsunit.eof Then P_UnitName=trim(rsunit("UnitName"))
rsunit.close
%>
<html>   
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>ExportBase</title>
<style type="text/css">
<!--
body {font-family:標楷體;font-size:12pt}
.style1 {font-family:標楷體;font-size:14pt}
-->
</style>
</head>	 
<body>  

	<table border="0" width="680px" cellpadding="0" cellspacing="0" style="border-collapse: collapse" >
		<tr>				 
			<td><span class="style1"><b><center><%=thenPasserCity%><%=thenPasserUnit%></center></b></span></td>
		</tr>
		<tr>
		   <td><span class="style1"><u><b><center>郵寄未退回清冊</center></b></u></span></td>
		</tr>
		<tr>
		   <td><center><%="("&strDateName(chkDate)&")"%>統計期間: <%=startDate_q%> 至 <%=endDate_q%></center>
		   <br>
			<%="單位名稱:" & P_UnitName & "<br>"%>
		   </td>
		</tr>		
	</table>
	
	<table border="1" width="680px" cellpadding="0" cellspacing="0">	
		<tr>
			<td><B><center>單位</center></B></td>
			<td width="50%"><B><center>件數</center></B></td>

		</tr>
<%
	strU="select * from UnitInfo where UnitID=UnitTypeID and ShowOrder>=0 order by UnitID"
	Set rsU=conn.execute(strU)
	While Not rsU.eof
%>
		<tr>
			<td>
			<%=Trim(rsU("UnitName"))%>
			</td>
			<td>
		<%
		UnitSql=" and BillUnitID in (select UnitID from UnitInfo where UnitTypeID='"&Trim(rsU("UnitID"))&"')"

		BilLBase="select Sn,BillNo,CarNo,BillTypeID,BillUnitID,RecordDate,RecordMemberID,IllegalDate,Owner,OwnerAddress,OwnerZip,Rule1,Rule2,BillMem1,BillMem2,BillMem3 from BillBase where BillNo is not null "&tmpSql&UnitSql&" and recordstateid=0 and billstatus=2 and NVL(EquiPmentID,1)<>-1"	
		if sys_City="台東縣" or sys_City="南投縣" then
			BilLBase=BilLBase&"  and billstatus<>'9'"
		End if		
		'2012/05/4 南投陳淑雲說 監理單位已先入案 n 違規人已先繳結案 L ，不出來，固修改 c.Status in ('Y','S','n','L') 為  c.Status in ('Y','S') by jafe,目前只有南投有改，其他縣市未更新過去
		If sys_City="苗栗縣" Then
			strSQL_Plus=" and c.Status in ('Y','S','n','L') "
		elseIf sys_City="南投縣" Or sys_City="台東縣" Then	'2015/3/13 南投李疑針說要加上 L 已經入案過
			strSQL_Plus=" and c.Status in ('Y','S','L') "
		Else
			strSQL_Plus=" and c.Status in ('Y','S','L') "
		End If 

		strSQL="select count(*) as cnt from ("&BilLBase&") a,UnitInfo b,BillBaseDCIReturn c,BillMailHistory d ,dcilog e where a.billno=e.billno and e.exchangetypeid='W' and e.exchangetypeid=c.exchangetypeid "&strSQL_Plus&" and e.DCIErrorCarData<>'V' and a.BillUnitID=b.UnitID and a.BillNo=c.BillNo(+) and a.CarNo=c.CarNo(+) and a.SN=d.BillSN(+) " 
		'response.write strSQL
		set rsfound=conn.execute(strSQL)
		If Not rsfound.eof then
			response.write rsfound("cnt")
		End If
		rsfound.close
		Set rsfound=nothing
		%>
			</td>
		</tr>
<%
	rsU.movenext
	Wend
	rsU.close 
	Set rsU=nothing
%>
	</table>
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_郵寄未退還統計表.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>	 
</body>
</html>