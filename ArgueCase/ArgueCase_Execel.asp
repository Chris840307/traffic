<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_申訴案件.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/DB.ini"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=Nothing
'response.write request("SQLstr")
if request("SQLstr")<>"" then
	set rsfound=conn.execute(request("SQLstr"))
end If

Server.ScriptTimeout = 60800
Response.flush
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>申訴案件</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" height="100%" border="1">
	<tr>
		<td align="center"><strong>申訴案件紀錄列表</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%" height="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<td>申訴日期</td>
					<td>舉發單號</td>
					<td>舉發員警</td>
					<td>收文號</td>
					<td>陳述事由</td>
					<td>適用條款</td>
					<td>造成缺失原因</td>
					<td>是否撤銷</td>
					<td>是否結案</td>
				</tr><%
					if request("SQLstr")<>"" then
						while Not rsfound.eof
							chname="":chRule=""
							strB="select * from (select * from BillBaseView where billno='"&Trim(rsfound("BillNo"))&"' order by Recorddate desc) where Rownum<=1"
							Set rsB=conn.execute(strB)
							If Not rsB.eof Then
								
								if rsB("BillMem1")<>"" then	chname=rsB("BillMem1")
								if rsB("BillMem2")<>"" then	chname=chname&"/"&rsB("BillMem2")
								if rsB("BillMem3")<>"" then	chname=chname&"/"&rsB("BillMem3")
								if rsB("Rule1")<>"" then chRule=rsB("Rule1")
								if rsB("Rule2")<>"" then chRule=chRule&"/"&rsB("Rule2")
								if rsB("Rule3")<>"" then chRule=chRule&"/"&rsB("Rule3")
								if rsB("Rule4")<>"" then chRule=chRule&"/"&rsB("Rule4")
							End If
							rsB.close
							Set rsB=Nothing 

							if rsfound("Cancel")="0" then
								chkCancel="是"
							else
								chkCancel="否"
							end if
							if rsfound("Close")="0" then
								chkClose="未處理"
							elseif rsfound("Close")="1" then
								chkClose="結案"
							elseif rsfound("Close")="2" then
								chkClose="待查中"
							end if
							response.write "<tr>"
							response.write "<td>"&rsfound("ArgueDate")&"&nbsp;</td>"
							response.write "<td>"&rsfound("BillNo")&"&nbsp;</td>"
							response.write "<td>"&chname&"</td>"
							response.write "<td>"&rsfound("DocNo")&"&nbsp;</td>"
							
							If sys_City="台南市" Or sys_City="高雄市" Then
								response.write "<td>"&rsfound("ArguerContent")&"</td>"
							Else
								if trim(rsfound("ArguerResonID"))="448" then
									response.write "<td>"&rsfound("ArguerResonName")&"</td>"
								else
									response.write "<td>"&rsfound("ArguerContent")&"</td>"
								end if
							End If 
							
							response.write "<td>"&chRule&"</td>"
							If sys_City="台南市" Or sys_City="高雄市" Then
								if trim(rsfound("ErrorID"))="0" then
									response.write "<td>無缺失</td>"
								else
									response.write "<td>"&rsfound("ErrorConten")&"</td>"
								end If
							else
								if trim(rsfound("ErrorID"))="453" then
									response.write "<td>"&rsfound("ErrorName")&"</td>"
								elseif trim(rsfound("ErrorID"))="0" then
									response.write "<td>無缺失</td>"
								else
									response.write "<td>"&rsfound("ErrorConten")&"</td>"
								end If
							End If 

							response.write "<td>"&chkCancel&"</td>"
							response.write "<td>"&chkClose&"</td>"
							response.write "</tr>"
							rsfound.movenext
						wend
					end if%>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%conn.close%>