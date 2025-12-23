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
<%if sys_City<>"雲林縣" then%>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="http://10.104.10.246/traffic/smsx.cab#Version=6,1,432,1">
</object>
<%end if%>
<html>
<head>
<script language="JavaScript">
	window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<title>車籍資料列表</title>
<script type="text/javascript" src="../js/Print.js"></script>
<!--#include virtual="traffic/Common/cssForForm.txt"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"--> 
<%
'權限
'AuthorityCheck(234)

%>
<%
Server.ScriptTimeout = 800
Response.flush

RecordDate=split(gInitDT(date),"-")
	strwhere=Session("PrintCarDataSQLxls")
	dcitype=trim(request("dcitype"))

	strQry=Session("PrintCarDataSQLCheckItem")

	strSQL="select distinct a.SN,a.CarSimpleID,a.IllegalDate,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillMem1,a.ProjectID,a.IllegalAddress,a.RuleSpeed,a.IllegalSpeed,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillNo,a.RuleVer,e.CarNo,e.DciReturnStation,e.DCIReturnCarType,e.A_Name,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerID,e.OwnerAddress,e.OwnerZip,e.Nwner,e.NwnerID,e.NwnerAddress,e.NwnerZip,e.DCIReturnCarStatus,e.DciCounterID from DCILog c,MemberData b,BillBase a,DCIReturnStatus d,BillBaseDCIReturn e where c.BillSN=a.SN and e.ExchangeTypeID='A' and e.Status='S' and a.CarNo=e.CarNo (+) and c.ExchangeTypeID=d.DCIActionID(+) and c.DCIReturnStatusID=d.DCIReturn(+) and c.RecordMemberID=b.MemberID(+) and a.RecordStateID=0 "&strwhere&" order by a.RecordDate"
	set rs1=conn.execute(strSQL)

	ConnExecute "列印車籍清冊="&strQry ,355

	strCnt="select count(*) as cnt from (select distinct a.SN,a.CarSimpleID,a.IllegalDate,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillMem1,a.ProjectID,a.IllegalAddress,a.RuleSpeed,a.IllegalSpeed,a.RecordStateID,a.RecordDate,a.RecordMemberID,a.BillNo,a.RuleVer,e.CarNo,e.DciReturnStation,e.DCIReturnCarType,e.DCIReturnCarColor,e.DriverHomeZip,e.DriverHomeAddress,e.Owner,e.OwnerID,e.OwnerAddress,e.OwnerZip,e.DCIReturnCarStatus,e.DciCounterID from DCILog c,MemberData b,BillBase a,DCIReturnStatus d,BillBaseDCIReturn e where c.BillSN=a.SN and e.ExchangeTypeID='A' and e.Status='S' and a.CarNo=e.CarNo (+) and c.ExchangeTypeID=d.DCIActionID(+) and c.DCIReturnStatusID=d.DCIReturn(+) and c.RecordMemberID=b.MemberID(+) and a.RecordStateID=0 "&strwhere&")"
	set Dbrs=conn.execute(strCnt)
	DBsum=Dbrs("cnt")
	Dbrs.close
	tmpSQL=strwhere


%>
<style type="text/css">
<!--
.style1 {font-size: 14pt;line-height:16pt}
.style3 {font-size: 12pt;line-height:13pt}
<%if sys_City="雲林縣" then%>
.pageprint {
  margin-left: 5.08mm;
  margin-right: 5.08mm;
  margin-top: 5.08mm;
  margin-bottom: 5.08mm;
 }
<%end if%>
-->
</style>
</head>
<body>
<form name=myForm method="post">
<%
	PrintSN=0
	If Not rs1.Bof Then rs1.MoveFirst 
	While Not rs1.Eof
	if PrintSN>0 then response.write "<div class=""PageNext""></div>"

	pagecnt=fix(Cint(trim(DBsum))/10+0.9999999)
%>
	<p align="right">頁次 <%=fix(PrintSN/10)+1%> of <%=pagecnt%></p>
	<table width="1200" border="0" cellpadding="0" cellspacing="0">
		<tr >
			<td colspan="12" align="center"><span class="style1">逕行告發違規資料清冊</span></td>
		</tr>
		<tr>
			<td width="7%" class="style3">違規車號</td>
			<td width="7%" class="style3">違規日期</td>
			<td width="4%" class="style3">時間</td>
			<td width="4%" class="style3">車種</td>
			<td width="11%" class="style3">違規地點</td>
			<td width="18%" class="style3"></td>
			<td width="8%" class="style3">舉發員警</td>
			<td width="7%" class="style3">專案代碼</td>
			<td width="8%" class="style3">詳細車種</td>
			<td width="8%" class="style3">處理狀態</td>
			<td width="10%" class="style3">車籍狀態</td>
			<td width="8%" class="style3">行駕照狀態</td>
		</tr>
		<tr>
			<td class="style3">法條代碼</td>
			<td class="style3">法條內容</td>
			<td class="style3"></td>
			<td class="style3">代碼</td>
			<td class="style3"></td>
			<td class="style3"></td>
			<td class="style3">違規事實</td>
			<td class="style3"></td>
			<td class="style3"></td>
			<td class="style3"></td>
			<td class="style3"></td>
			<td class="style3"></td>
		</tr>
		<tr>
			<td class="style3">車主證號</td>
			<td class="style3"></td>
			<td class="style3" colspan="2">車主姓名</td>
			<!-- <td class="style3"></td> -->
			<td class="style3"></td>
			<td class="style3">車主地址</td>
			<td class="style3"></td>
			<td class="style3"></td>
			<td class="style3">應到案處所</td>
			<td class="style3"></td>
			<td class="style3">顏色</td>
			<td class="style3">廠牌</td>
		</tr>
	</table>
	<hr>
<%		for i=1 to 10
			if rs1.eof then exit for
			PrintSN=PrintSN+1
%>
	<table width="1200" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td width="7%"></td>
			<td width="7%"></td>
			<td width="4%"></td>
			<td width="4%"></td>
			<td width="11%"></td>
			<td width="18%"></td>
			<td width="8%"></td>
			<td width="7%" ></td>
			<td width="8%" ></td>
			<td width="8%" ></td>
			<td width="10%" ></td>
			<td width="8%" ></td>
		</tr>
		<tr>
			<td class="style3"><%
			'違規車號
			if trim(rs1("CarNo"))<>"" and not isnull(rs1("CarNo")) then
				response.write trim(rs1("CarNo"))
			end if
			%></td>
			<td class="style3"><%
			'違規日期
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write year(rs1("IllegalDate"))-1911&"/"&month(rs1("IllegalDate"))&"/"&day(rs1("IllegalDate"))
			end if
			%></td>
			<td class="style3"><%
			'時間
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				if len(hour(rs1("IllegalDate"))) < 2 then 
					sHour = "0" & hour(rs1("IllegalDate"))
				else
					sHour = hour(rs1("IllegalDate"))	
				end if
				if len(minute(rs1("IllegalDate"))) < 2 then 
					sMinute = "0" & minute(rs1("IllegalDate"))
				else
					sMinute = minute(rs1("IllegalDate"))	
				end if
				response.write hour(rs1("IllegalDate"))&":"&minute(rs1("IllegalDate"))
			end if
			%></td>
			<td class="style3"><%
			'車種
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				response.write trim(rs1("CarSimpleID"))
			end if
			%></td>
			<td class="style3" colspan="2"><%
			'違規地點
			if trim(rs1("IllegalAddress"))<>"" and not isnull(rs1("IllegalAddress")) then
				response.write trim(rs1("IllegalAddress"))
			end if
			%></td>
			<td class="style3"><%
			'舉發員警
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
			end if
			%></td>
			<td class="style3"><%
			'專案代碼
			if trim(rs1("ProjectID"))<>"" and not isnull(rs1("ProjectID")) then
				response.write trim(rs1("ProjectID"))
			end if
			%></td>
			<td class="style3"><%
			'詳細車種
			if trim(rs1("DCIReturnCarType"))<>"" and not isnull(rs1("DCIReturnCarType")) then
				strCType="select * from DCIcode where TypeID=5 and ID='"&trim(rs1("DCIReturnCarType"))&"'"
				set rsCType=conn.execute(strCType)
				if not rsCType.eof then
					response.write trim(rsCType("Content"))
				end if
				rsCType.close
				set rsCType=nothing
			end if								
			%></td>
			<td class="style3"><%
			'處理狀態
			strStatus="select ExchangeTypeID,DCIReturnStatusID from DCILog where BillSN="&trim(rs1("SN"))&" order by ExchangeDate Desc"
			set rsStatus=conn.execute(strStatus)
			if not rsStatus.eof then
				strSID="select StatusContent from DCIReturnStatus where DCIactionId='"&trim(rsStatus("ExchangeTypeID"))&"' and DCIreturn='"&trim(rsStatus("DCIReturnStatusID"))&"'"
				set rsSID=conn.execute(strSID)
				if not rsSID.eof then
					response.write trim(rsSID("StatusContent"))
				end if
				rsSID.close
				set rsSID=nothing
			end if
			rsStatus.close
			set rsStatus=nothing
			%></td>
			<td class="style3" ><%
			'車籍狀態
				if trim(rs1("DCIReturnCarStatus"))<>"" and not isnull(rs1("DCIReturnCarStatus")) then
					strCstatus="select Content from DCIcode where TypeID=10 and ID='"&trim(rs1("DCIReturnCarStatus"))&"'"
					set rsCS=conn.execute(strCstatus)
					if not rsCS.eof then
						response.write trim(rsCS("COntent"))
					end if 
					rsCS.close
					set rsCS=nothing
				end if
			%></td>
			<td class="style3" ><%
			'行駕照狀態
				if trim(rs1("DciCounterID"))<>"" and not isnull(rs1("DciCounterID")) then
					If trim(rs1("DciCounterID"))="x" Then
						 response.write "<strong>駕照過期</strong>"
					ElseIf trim(rs1("DciCounterID"))="y" Then
						response.write "<strong>行照過期</strong>"
					ElseIf trim(rs1("DciCounterID"))="v" Then
						response.write "<strong>行駕照過期</strong>"
					End If 
				end if
			%></td>
		</tr>
		<%if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then%>
		<tr>
			<td class="style3"><%
			'法條代碼
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				response.write trim(rs1("Rule1"))
			end if
			%></td>
			<td class="style3" colspan="5"><%
			'法條內容
			if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
				strCarImple=""
				if left(trim(rs1("Rule1")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple=" and CarSimpleID in ('3','0')"
					else
						strCarImple=""
					end if
				end if
				strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple
				set rsR1=conn.execute(strR1)
				if not rsR1.eof then 
					response.write " "&trim(rsR1("IllegalRule"))
				end if
				rsR1.close
				set rsR1=nothing
				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					response.write " "&"("&trim(rs1("Rule4")) & ")"
				end if
			end if
			%></td>

			<td class="style3" colspan="6"><%
			'違規事實
			if (trim(rs1("RuleSpeed"))<>"" and not isnull(rs1("RuleSpeed"))) and (trim(rs1("IllegalSpeed"))<>"" and not isnull(rs1("IllegalSpeed"))) then
				response.write "速限"&trim(rs1("RuleSpeed"))&"公里時速"&trim(rs1("IllegalSpeed"))&"公里，超速"&trim(rs1("IllegalSpeed"))-trim(rs1("RuleSpeed"))&"公里"
			end if
			%></td>
		</tr>
		<%end if%>
		<%if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then%>
		<tr>
			<td class="style3"><%
			'法條代碼
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				response.write trim(rs1("Rule2"))
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				response.write "<br>"&trim(rs1("Rule3"))
			end if
			%></td>
			<td class="style3" colspan="5"><%
			'法條內容
			if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
				strCarImple=""
				if left(trim(rs1("Rule2")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple=" and CarSimpleID in ('3','0')"
					else
						strCarImple=""
					end if
				end if

				strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple
				set rsR1=conn.execute(strR1)
				if not rsR1.eof then 
					response.write " "&trim(rsR1("IllegalRule"))
				end if
				rsR1.close
				set rsR1=nothing
			end if
			if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
				strCarImple=""
				if left(trim(rs1("Rule3")),4)="2110" then
					if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
						strCarImple=" and CarSimpleID in ('5','0')"
					elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
						strCarImple=" and CarSimpleID in ('3','0')"
					else
						strCarImple=""
					end if
				end if

				strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule3"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple
				set rsR1=conn.execute(strR1)
				if not rsR1.eof then 
					response.write "<br> "&trim(rsR1("IllegalRule"))
				end if
				rsR1.close
				set rsR1=nothing
			end if
			%></td>

			<td class="style3" colspan="6"></td>
		</tr>
		<%end if%>
		<tr>
			<td class="style3" colspan="2"><%
			'車主證號
			if trim(rs1("OwnerID"))<>"" and not isnull(rs1("OwnerID")) then
				response.write trim(rs1("OwnerID"))
			end if
			%></td>
			<td class="style3" colspan="3"><%
			'車主姓名
			if trim(rs1("Owner"))<>"" and not isnull(rs1("Owner")) then
				response.write trim(rs1("Owner"))
			end if
			%></td>
			<td class="style3" colspan="3"><%
			'車主地址
			if trim(rs1("OwnerAddress"))<>"" and not isnull(rs1("OwnerAddress")) then
				response.write trim(rs1("OwnerAddress"))
			end if
			%></td>
			<td class="style3" colspan="2"><%
			'應到案處所
			if instr(trim(rs1("DciReturnStation")),"20")>0 or instr(trim(rs1("DciReturnStation")),"21")>0 or instr(trim(rs1("DciReturnStation")),"22")>0 or instr(trim(rs1("DciReturnStation")),"23")>0 or instr(trim(rs1("DciReturnStation")),"24")>0 or instr(trim(rs1("DciReturnStation")),"25")>0 or instr(trim(rs1("DciReturnStation")),"26")>0 then
				response.write "台北市交通事件裁決所"
			elseif instr(StationArrayTemp,"30")>0 or instr(StationArrayTemp,"31")>0 or instr(StationArrayTemp,"32")>0 then
				response.write "高雄市交通事件裁決所"
			else
				strSqlStationName="select DCIstationName from Station where DCIstationID='"&trim(rs1("DciReturnStation"))&"'"
				set rsSN=conn.execute(strSqlStationName)
				if not rsSN.eof then
					response.write trim(rsSN("DCIstationName"))
				end if
				rsSN.close
				set rsSN=nothing
			end if
			%></td>
			<td class="style3"><%
			'顏色
			if trim(rs1("DCIReturnCarColor"))<>"" and not isnull(rs1("DCIReturnCarColor")) then
				ColorLen=cint(Len(rs1("DCIReturnCarColor")))
				for Clen=1 to ColorLen
					colorID=mid(rs1("DCIReturnCarColor"),Clen,1)
					strColor="select * from DCIcode where TypeID=4 and ID='"&trim(colorID)&"'"
					set rsColor=conn.execute(strColor)
					if not rsColor.eof then
						response.write trim(rsColor("Content"))
					end if
					rsColor.close
					set rsColor=nothing
				next
			end if
			%></td>
			<td class="style3"><%
			'廠牌
				if (trim(rs1("A_Name"))<>"" and not isnull(rs1("A_Name"))) then
					response.write trim(rs1("A_Name"))
				end if
			%></td>
		</tr>
	</table>
	<hr>
<%			
		rs1.MoveNext
		next
%>
<%
		Wend
		rs1.close
		set rs1=nothing
%>		
</form>
</body>
<script language="javascript">
//printWindow(true,2,2,2,2);
<%if sys_City="雲林縣" then%>
window.print();
<%end if%>
</script>
</html>
<%conn.close%>