<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
fMnoth=month(now)
if fMnoth<10 then fMnoth="0"&fMnoth
fDay=day(now)
if fDay<10 then	fDay="0"&fDay
fname=year(now)&fMnoth&fDay&"_申訴案件ㄧ覽表.xls"
Response.AddHeader "Content-Disposition", "filename="&fname
response.contenttype="application/x-msexcel; charset=MS950" 
%>
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=Nothing

if request("SQLstr")<>"" then
	set rsfound=conn.execute(request("SQLstr"))
end If

Server.ScriptTimeout = 60800
Response.flush
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title> 民眾申述案件管制一覽表</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<body>
<table width="100%" height="100%" border="1">
	<tr>
		<td align="center"><%
		strU="select * from Apconfigure where id=30"
		Set rsU=conn.execute(strU)
		If Not rsU.eof Then
			response.write Trim(rsU("Value"))&" &nbsp;"
		End If 
		rsU.close
		Set rsU=Nothing
		If Trim(request("Sys_Unit"))<>"" Then
			strU="select * from UnitInfo where UnitID='"&Trim(request("Sys_Unit"))&"'"
			Set rsU=conn.execute(strU)
			If Not rsU.eof Then
				response.write Trim(rsU("UnitName"))&" &nbsp;"
			End If 
			rsU.close
			Set rsU=Nothing
			
		End If 
		%>員警舉發交通違規(欄停、逕行舉發)民眾申述案件管制一覽表</td>
	</tr>
	<tr>
		<td>
			<table width="100%" height="100%" border="1" cellpadding="4" cellspacing="1">
				<tr>
					<td>編號</td>
					<td>單位</td>
					<td>舉發<br>方式</td>
					<td>陳情方式</td>									
					<td>來文(列管<br>/發現)日期</td>
					<td>來文機關<br>(列管)文號</td>	
					<td>陳述人<br>(違規人)</td>		
					<td>舉發單號</td>
					<td>車號</td>					
					<td>違規日</td>								
					<td>舉發日</td>								
					<td>違規條款</td>								
					<td>移送管轄機關日期<br>(逕行舉發案件才填)</td>
					<td>陳述事由</td>		
					<td>辦理情形</td>			
					<td>回覆日期 / 文號</td>																					
					<td>舉發員警</td>
					
					<td>違反規定(項)</td>
					<td>違反規定(目)</td>
				<%If sys_City<>"台南市" Then%>
					<td>懲處情形</td>
				<%End If %>
					<td>劣蹟處分</td>
					<td>申誡處分</td>
					<!-- <td>處分(通報)日期</td>
					<td>處分(通報)文號</td> -->
				<%If sys_City="台南市" Then%>
					<td>是否撤單</td>
					<td>備註/撤銷舉發單理由</td>
				<%else%>
					<td>備註</td>
				<%End If %>
					
					<!--
					<td>適用條款</td>
					<td>造成缺失原因</td>
					<td>是否撤銷</td>
					<td>是否結案</td>
					-->
				</tr><%
					if request("SQLstr")<>"" then
							iSN=0
						while Not rsfound.eof
							iSN=iSN+1
							strB="select * from (select * from Billbase where billno='"&Trim(rsfound("BillNo"))&"' order by Recorddate desc) where Rownum<=1"
							Set rsB=conn.execute(strB)
							If Not rsB.eof Then
								chname="":chRule=""
								if rsB("BillMem1")<>"" then	chname=rsB("BillMem1")
								if rsB("BillMem2")<>"" then	chname=chname&"/"&rsB("BillMem2")
								if rsB("BillMem3")<>"" then	chname=chname&"/"&rsB("BillMem3")
								if rsB("Rule1")<>"" then chRule=rsB("Rule1")
								if rsB("Rule2")<>"" then chRule=chRule&"/"&rsB("Rule2")
								if rsB("Rule3")<>"" then chRule=chRule&"/"&rsB("Rule3")
								if rsB("Rule4")<>"" then chRule=chRule&"/"&rsB("Rule4")
								CarNoTmp=rsB("CarNo")
								If Trim(rsB("BillUnitID"))<>"" Then
									strU="select UnitName from UnitInfo where UnitID='"&Trim(rsB("BillUnitID"))&"'"
									Set rsU=conn.execute(strU)
									If not rsU.eof Then
										UnitNameTmp=Trim(rsU("UnitName"))
									End If
									rsU.close
									Set rsU=nothing
								End If 
								if (rsB("Billtypeid")=2) then
									BilltypeNameTmp="逕舉" 						
								else
									BilltypeNameTmp="攔停"
								end if	
								BillTypeIDTmp=trim(rsB("Billtypeid"))
								IllegalDateTmp=ginitdt(rsB("IllegalDate"))
								If Not IsNull(rsB("BillFillDate")) And Trim(rsB("BillFillDate"))<>"" then
									BillFillDateTmp=Year(rsB("BillFillDate"))-1911&Right("00"&month(rsB("BillFillDate")),2)&Right("00"&day(rsB("BillFillDate")),2)
								Else
									BillFillDateTmp="&nbsp;"
								end if
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
							response.write "<td>"& iSN &"</td>"
							response.write "<td>"&UnitNameTmp&"</td>"
							response.write "<td>"& BilltypeNameTmp &"</td>"									
							response.write "<td>"&rsfound("ArgueWay")&"</td>"
							response.write "<td>"
							If Not IsNull(rsfound("ArgueDate")) And Trim(rsfound("ArgueDate"))<>"" then
								response.write Year(rsfound("ArgueDate"))-1911&Right("00"&month(rsfound("ArgueDate")),2)&Right("00"&day(rsfound("ArgueDate")),2)
							Else
								response.write "&nbsp;"
							end if
							response.write "</td>"
							response.write "<td>"&rsfound("reportdeparment")&"/" & rsfound("reportno")&"</td>"
							
							response.write "<td>"&rsfound("Arguer")&"</td>"
							response.write "<td>"&rsfound("BillNo")&"</td>"
							response.write "<td>"&CarNoTmp&"</td>"
							response.write "<td>"&IllegalDateTmp&"</td>"
							response.write "<td>"
							response.write BillFillDateTmp
							response.write "</td>"				
							response.write "<td>"&chRule&"</td>" 
							If trim(BillTypeIDTmp)="2" Then
								response.write "<td>"
								strDci="select DciCaseInDate from billbasedcireturn where BillNo='"&Trim(rsfound("BillNo"))&"' and CarNo='"&Trim(CarNoTmp)&"' and Exchangetypeid='W' "
								Set rsDci=conn.execute(strDci)
								If Not rsDci.eof Then
									response.write Trim(rsDci("DciCaseInDate"))
								End If
								rsDci.close
								Set rsDci=Nothing 
								response.write "</td>"
							Else
								response.write "<td>&nbsp;</td>" 
							End If 
							If sys_City="台南市" Then
								if trim(rsfound("ArguerResonID"))="448" then
									response.write "<td>"&rsfound("ArguerContent")
									If Trim(rsfound("ArguerResonName"))<>"" Then 
										response.write ","&rsfound("ArguerResonName")
									End If
									If Trim(rsfound("ArguerContent2"))<>"" then
										response.write ","&rsfound("ArguerContent2")
									End if
									response.write "</td>"
								else
									response.write "<td>"&rsfound("ArguerContent")&"</td>"
								end if
							Else
								if trim(rsfound("ArguerResonID"))="448" then
									response.write "<td>"&rsfound("ArguerResonName")&"</td>"
								else
									response.write "<td>"&rsfound("ArguerContent")&"</td>"
								end if
							End If 
							If sys_City="台南市" Then
								if trim(rsfound("ErrorID"))="0" then
									response.write "<td>無缺失</td>"
								elseif trim(rsfound("ErrorID"))="453" then
									response.write "<td>"&rsfound("ErrorConten") & "," &Trim(rsfound("ErrorName"))&"</td>"
								else
									response.write "<td>"&rsfound("ErrorConten")&"</td>"
								end If
							Else
								If Trim(rsfound("ReportContent"))<>"" Then
									response.write "<td>"&rsfound("ReportContent")&"</td>"	
								Else
									response.write "<td>"&rsfound("ArguerContent")&"</td>"	
								End If 
							End If 
							
							response.write "<td>"
							If Not IsNull(rsfound("processdate")) And Trim(rsfound("processdate"))<>"" then
								response.write Year(rsfound("processdate"))-1911&Right("00"&month(rsfound("processdate")),2)&Right("00"&day(rsfound("processdate")),2)
							Else
								response.write "&nbsp;"
							end if
							response.write "/" & rsfound("processno")&"</td>"																																
							response.write "<td>"&chname&"</td>"
							response.write "<td>"&rsfound("VIOLATERULE1")&"</td>"	
							response.write "<td>"&rsfound("VIOLATERULE2")&"</td>"	
						If sys_City<>"台南市" Then
							response.write "<td>"&rsfound("Punishment")&"</td>"	
						End If 
							If IsNull(rsfound("BadCnt")) Or Trim(rsfound("BadCnt"))="0" Then
							response.write "<td>&nbsp;</td>"
							Else
							response.write "<td>"&Trim(rsfound("BadCnt"))&"</td>"
							End If
							If IsNull(rsfound("WarnCnt")) Or Trim(rsfound("WarnCnt"))="0" Then
							response.write "<td>&nbsp;</td>"
							Else
							response.write "<td>"&Trim(rsfound("WarnCnt"))&"</td>"
							End If
						If sys_City="台南市" Then
							If Trim(rsfound("Cancel"))="0" Then
								response.write "<td>是</td>"
							Else
								response.write "<td>否</td>"
							End If 
							response.write "<td>"
							If Trim(rsfound("Note"))<>"" Then
								response.write Trim(rsfound("Note"))
							End If 
							If Not IsNull(rsfound("DELBILLREASON")) Then
								strDR="select * from code where id="&Trim(rsfound("DELBILLREASON"))
								Set rsDR=conn.execute(strDR)
								If Not rsDR.eof Then
									response.write " / " & rsDR("Content") 
									If Trim(rsfound("DELBILLREASON"))="811" Then
										response.write "," & rsfound("DelName") 
									End If 
								End If
								rsDR.close
								Set rsDR=Nothing 
							End If 
							
							response.write "</td>"
	
						else
							response.write "<td>"&rsfound("Note")&"</td>"
						End If 
							
							response.write "</tr>"

							Response.flush

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