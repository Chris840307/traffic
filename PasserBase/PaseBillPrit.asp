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
fname=year(now)&fMnoth&fDay&"_裁決書.doc"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/msword; charset=MS950" 

strSql="select a.SN as BillSN,a.BillNo,a.Driver,a.DriverBirth,a.DriverID,a.DriverZip,a.DriverAddress,a.IllegalDate,a.IllegalAddress,a.DealLineDate,a.Rule1,a.Rule2,a.Rule3,a.Rule4,a.BillUnitID,b.OpenGovNumber as JudeOGN,b.AgentName as JudeAgentName,b.AgentSex as JudeAgentSex,b.AgentBirth as JudeAgentBirth,b.AgentID as JudeAgentID,b.AgentAddress as JudeAgentAddress,c.OpenGovNumber as UrgeOGN,c.UrgeTypeID,d.OpenGovNumber,d.BigUnitBossName,d.SubUnitSecBossName,d.SendNumber,d.SendDate,d.Agent,d.AgentBirthDate,d.AgentID,d.AgentAddress,d.ForFeit,d.MakeSureDate,d.LimitDate,d.AttatchJude,d.AttatchUrge,d.AttatchFortune,d.AttatchGround,d.AttatchRegister,d.AttatchFileList,d.AttatchTable,d.ATTATPOSTAGE,d.SafeToExit,d.SAFEACTION,d.SAFEASSURE,d.SAFEDETAIN,d.SAFESHUTSHOP from PasserBase a,PasserJude b,PasserUrge c,PasserSend d where a.SN="&trim(request("PBillSN"))&" and a.SN=b.BillSN(+) and a.BillNo=b.BillNo(+) and a.SN=c.BillSN(+) and a.BillNo=c.BillNo(+) and a.SN=d.BillSN(+) and a.BillNo=d.BillNo(+)"
PrintDate=split(gArrDT(date),"-")
set rsfound=conn.execute(strSql)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>無標題文件</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-family: "標楷體"; font-size: 18px; }
.style2 {font-family: "標楷體"; font-size: 14px; }
.style3 {font-family: "標楷體"; font-size: 14px; }
-->
</style>
</head>

<body>
<table width="333" border="0" align="right" cellspacing=0 cellpadding=0>
<tr><td width="84" height="33"><span class="style1"><%'移送案號%></span></td>
<td width="233"><span class="style1">　<%=rsfound("SendNumber")%></span></td></tr>
</table>
<br><br>
<table width="100%" border="0" cellspacing=0 cellpadding=0>
	<tr>
		<td colspan="5" rowspan="2"><span class="style1"></span></td>
		<td width="87" height="34"><span class="style1"><%'發文日期%></span></td>
		<td width="211"><span class="style1">　　<%=PrintDate(0)%>　　　<%=PrintDate(1)%>　　<%=PrintDate(2)%></span></td>
	</tr>
	<tr>
		<td height="39"><span class="style1"><%'發文字號%></span></td>
	  <td><span class="style1" valign="top">　　　　　　　　<%=rsfound("OpenGovNumber")%></span></td>
	</tr>
	<tr>
		<td height="41" colspan="2" align="center"><span class="style1"><%'姓名或名稱%></span></td>
		<td width="121" align="center"><span class="style1"><%'出生年月日%></span></td>
		<td width="95" align="center"><span class="style1"><%'性別%></span></td>
		<td width="115" align="center"><span class="style1"><%'身分證統一編號或%><br>
	    <%'營利事業統一編號%></span></td>
		<td colspan="2" align="center"><span class="style1"><%'住址或事務所、營業所地址及郵遞區號%></span></td>
	</tr>
	<tr>
	  <td width="141" height="56" align="center"><span class="style1">
      <%'義務人%></span></td>
		<td width="203"><span class="style1"><%=rsfound("Driver")%></span></td>
		<td><span class="style2">
	    <%if trim(rsfound("DriverBirth"))<>"" then
				DriverBirth=split(gArrDT(rsfound("DriverBirth")),"-")
				response.write DriverBirth(0)&"年"&DriverBirth(1)&"月"&DriverBirth(2)&"日"
		end if%>
		</span></td>
		<td><span class="style1">
	    <%if Not rsfound.eof then
			If not ifnull(Trim(rsfound("DriverID"))) Then
				If Mid(Trim(rsfound("DriverID")),2,1)="1" Then
					Response.write "男"
				elseif Mid(Trim(rsfound("DriverID")),2,1)="2" Then
					Response.write "女"
				End if
			End if
		end if%>
		</span></td>
		<td><span class="style1"><%=rsfound("DriverID")%></span></td>
		<td><span class="style1"><%'戶籍地%></span></td>
		<td><span class="style1" valign="top"><%=trim(rsfound("DriverZip"))&trim(rsfound("DriverAddress"))%></span></td>
	</tr>
	<tr>
	  <td height="52" align="center"><span class="style1">
      <%'法定代理人%></span></td>
		<td><span class="style1">
	    <%
			if trim(rsfound("Agent"))<>"" then
				response.write rsfound("Agent")
			else
				response.write rsfound("JudeAgentName")
			end if
			%>
		</span></td>
		<td><span class="style1">
	    <%
			if trim(rsfound("AgentBirthDate"))<>"" then
				AgentBirthDate=split(gArrDT(rsfound("AgentBirthDate")),"-")
			else
				AgentBirthDate=split(gArrDT(rsfound("JudeAgentBirth")),"-")
			end if
			if trim(AgentBirthDate(0))<>"" then
				response.write "　"&AgentBirthDate(0)&"年"&AgentBirthDate(1)&"月"&AgentBirthDate(2)&"日"
			end if
			%>
		</span></td>
		<td><span class="style1">
	    <%
			if trim(rsfound("JudeAgentSex"))="1" then
				response.write "男"
			elseif trim(rsfound("JudeAgentSex"))="0" then
				response.write "女"
			end if%>
		</span></td>
		<td><span class="style1">
	    <%
			if trim(rsfound("AgentID"))<>"" then
				response.write rsfound("AgentID")
			else
				response.write rsfound("JudeAgentID")
			end if%>
		</span></td>
		<td><span class="style1"><%'營業地%></span></td>
		<td><span class="style1">
	    <%
			if trim(rsfound("AgentAddress"))<>"" then
				response.write rsfound("AgentAddress")
			else
				response.write rsfound("JudeAgentAddress")
			end if%>
		</span></td>
	</tr>
	<tr>
		<td rowspan="2" align="center"><span class="style1"><%'義務發生之%><br>
	    <%'原因與日期%></span></td>
		<td rowspan="2" Valign="top">		  <span class="style3">
		<%
			if trim(rsfound("Rule1"))<>"" then
				response.write "　　　　　　　　　　　"&left(trim(rsfound("Rule1")),2)
				if len(trim(rsfound("Rule1")))>7 then response.write "之"&right(trim(rsfound("Rule1")),1)
				response.write "<br>"
				response.write "　"&Mid(trim(rsfound("Rule1")),3,1)&"　　　　"&Mid(trim(rsfound("Rule1")),4,2)
				response.write "<br><br>"
			end if	
			'if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
				'response.write "違反道路交通管理處罰條例第"&left(trim(rsfound("Rule2")),2)&"條"
				'if len(trim(rsfound("Rule2")))>7 then response.write "之"&right(trim(rsfound("Rule2")),1)
				'response.write "<br>"
				'response.write "　　　　　　"&Mid(trim(rsfound("Rule2")),3,2)&"項"&Mid(trim(rsfound("Rule2")),5,2)&"款規定"
				'response.write "<br>"
			'end if	
			'if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
				'response.write "違反道路交通管理處罰條例第"&left(trim(rsfound("Rule3")),2)&"條"
				'if len(trim(rsfound("Rule3")))>7 then response.write "之"&right(trim(rsfound("Rule3")),1)
				'response.write "<br>"
				'response.write "　　　　　　"&Mid(trim(rsfound("Rule3")),3,2)&"項"&Mid(trim(rsfound("Rule3")),5,2)&"款規定"
				'response.write "<br>"
			'end if	
			'if trim(rsfound("Rule4"))<>"" and not isnull(rsfound("Rule4")) then
				'response.write "違反道路交通管理處罰條例第"&left(trim(rsfound("Rule4")),2)&"條"
				'if len(trim(rsfound("Rule4")))>7 then response.write "之"&right(trim(rsfound("Rule4")),1)
				'response.write "<br>"
				'response.write "　　　　　　"&Mid(trim(rsfound("Rule4")),3,2)&"項"&Mid(trim(rsfound("Rule4")),5,2)&"款規定"
				'response.write "<br>"
			'end if
			IllegalDate=split(gArrDT(rsfound("IllegalDate")),"-")
			response.write "　　　　　　"&IllegalDate(0)&"　　"&IllegalDate(1)&"　　"&IllegalDate(2)&"<BR>"
			response.write "　　　　　　　　"&rsfound("JudeOGN")
			%>
	    </span></td>
		<td height="51"><span class="style1">
	    <%
			'if trim(rsfound("MakeSureDate"))<>"" then
			'response.write "■確定日期"
			'else
			'response.write "□確定日期"
			'end if%>
		  <br>
        <%
			'if trim(rsfound("LimitDate"))<>"" then
			'response.write "■限繳日期"
			'else
			'response.write "□限繳日期"
			'end if%>	  
      </span></td>
		<td colspan="2" Valign="top"><span class="style2">
	    <%
			if trim(rsfound("MakeSureDate"))<>"" then
				MakeSureDate=split(gArrDT(rsfound("MakeSureDate")),"-")
				response.write "　"&MakeSureDate(0)&"　　　"&MakeSureDate(1)&"　　　　"&MakeSureDate(2)
			end if
			'if trim(rsfound("LimitDate"))<>"" then
				'if trim(rsfound("MakeSureDate"))<>"" then response.write "<BR>"
				'LimitDate=split(gArrDT(rsfound("LimitDate")),"-")
				'response.write "　　"&LimitDate(0)&"　　　　"&LimitDate(1)&"　　　"&LimitDate(2)
			'end if%>
		</span></td>
		<td colspan="2"><span class="style1"><%'執行標的物所在地%></span></td>
	</tr>
	<tr>
		<td height="53"><span class="style1">
	  <%'繳納金額%></span></td>
		<td colspan="2" Valign="top"><span class="style1">　　　　　　　<%=rsfound("ForFeit")%></span></td>
		<td colspan="2" Valign="top"><span class="style1"><%="　"&rsfound("IllegalAddress")%></span></td>
	</tr>
	<tr>
	  <td height="73" align="center"><span class="style1"><%'移送法條%></span></td>
		<td colspan="2"><span class="style1">
		  <%'一、依據行政執行法第十一條、第十三條。%><br>
		  <%
'			chcnt=split("空,二,三,四,五",",")
'			chsum=0
'			if trim(rsfound("Rule1"))<>"" then
'				chsum=chsum+1
'				response.write chcnt(chsum)&"依據道路交通管理處罰條例第"&left(trim(rsfound("Rule1")),2)&"條"
'				if len(trim(rsfound("Rule1")))>7 then response.write "之"&right(trim(rsfound("Rule1")),1)
'				response.write "。"
'				response.write "<br>"
'			end if	
'			if trim(rsfound("Rule2"))<>"" and not isnull(rsfound("Rule2")) then
'				chsum=chsum+1
'				response.write chcnt(chsum)&"依據道路交通管理處罰條例第"&left(trim(rsfound("Rule2")),2)&"條"
'				if len(trim(rsfound("Rule2")))>7 then response.write "之"&right(trim(rsfound("Rule2")),1)
'				response.write "。"
'				response.write "<br>"
'			end if	
'			if trim(rsfound("Rule3"))<>"" and not isnull(rsfound("Rule3")) then
'				chsum=chsum+1
'				response.write chcnt(chsum)&"依據道路交通管理處罰條例第"&left(trim(rsfound("Rule3")),2)&"條"
'				if len(trim(rsfound("Rule3")))>7 then response.write "之"&right(trim(rsfound("Rule3")),1)
'				response.write "。"
'				response.write "<br>"
'			end if	
'			if trim(rsfound("Rule4"))<>"" and not isnull(rsfound("Rule4")) then
'				chsum=chsum+1
'				response.write chcnt(chsum)&"依據道路交通管理處罰條例第"&left(trim(rsfound("Rule4")),2)&"條"
'				if len(trim(rsfound("Rule4")))>7 then response.write "之"&right(trim(rsfound("Rule4")),1)
'				response.write "。"
'				response.write "<br>"
'			end if
			%>
		</span></td>
		<td nowrap><span class="style1"><%'催繳情形%></span></td>
		<td><span class="style1">
	    <%
'			strchk="select count(*) as cnt from PasserJude where BillSN="&rsfound("BillSN")&" and BillNo='"&rsfound("BillNo")&"'"
'			set rschk=conn.execute(strchk)
'			Jodestr=Cint(rschk("cnt"))
'			rschk.close
'			if trim(Jodestr)<>"0" then
'				response.write "■業經催繳<br>"
'				response.write "□未經催繳"
'			else
'				response.write "□業經催繳<br>"
'				response.write "■未經催繳"
'			end if
		%>
		</span></td>
		<td><span class="style1"><%'催繳方式%></span></td>
		<td><span class="style1">
	    <%
'			if trim(rsfound("UrgeTypeID"))="0" then
'				response.write "■電話催繳<br>"
'				response.write "□名信片或信函方式催繳<br>"
'				response.write "□其他方式(方式為：裁決書送達）"
'			elseif trim(rsfound("UrgeTypeID"))="0" then
'				response.write "□電話催繳<br>"
'				response.write "■名信片或信函方式催繳<br>"
'				response.write "□其他方式(方式為：裁決書送達）"
'			else
'				response.write "□電話催繳<br>"
'				response.write "□名信片或信函方式催繳<br>"
'				response.write "■其他方式(方式為：裁決書送達）"
'			end if%>
		</span></td>
	</tr>
	<tr>
		<td align="center"><span class="style1"><%'附件%></span></td>
		<td colspan="6" height="97">
			<span class="style1">
			<%
'			if trim(rsfound("AttatchTable"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'附表%><br>          
			<%
'			if trim(rsfound("AttatchJude"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'處分書裁決書或義務人依法令負有義務之證明文件及送達證明文件%>
			<%
'			if trim(rsfound("AttatchGround"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'土地登記部謄本%><br>
			<%
'			if trim(rsfound("AttatchUrge"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'義務人經限期履行而逾期能不履行之證明文件及送達證明文件%>
			<%
'			if trim(rsfound("AttatchRegister"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'義務人之戶籍資料%><br>
			<%
'			if trim(rsfound("AttatchFortune"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'義務人之財產目錄%>
			<%
'			if trim(rsfound("AttatchFileList"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'磁片電子檔清單%>
			<%
'			if trim(rsfound("ATTATPOSTAGE"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
	  <%'郵資%>	    </span></td>
	</tr>
	<tr>
	  <td height="39" align="center"><span class="style1"><%'保全措施%></span></td>
		<td colspan="6">
			<span class="style1">
			<%
			'if trim(rsfound("SAFETOEXIT"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'已限制出境 %>
			<%
'			if trim(rsfound("SAFEACTION"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'已禁止處分%>
			<%
'			if trim(rsfound("SAFEASSURE"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'已提供擔保 %>
			<%
'			if trim(rsfound("SAFEDETAIN"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'已假扣押%>
			<%
'			if trim(rsfound("SAFESHUTSHOP"))<>"" then
'				response.write "■"
'			else
'				response.write "□"
'			end if%>
			<%'已勒令停業%>
	    </span></td>
	</tr>
	<tr>
		<td height="90" colspan="7">&nbsp;</td>
	</tr>
</table>
</body>
</html>
<script language="javascript">
function DP(){
	window.focus();
	window.print();
}
</script>