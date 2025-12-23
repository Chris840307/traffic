<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<%

str_ActDate="全部"

if request("ActDate1")<>"" and request("ActDate2")<>""then

	str_ActDate=request("ActDate1")&" 至 "& request("ActDate2")
end if

strSQL="select UnitName,ChName,sum(PrintCnt) SumPrintCnt from (" & Request("Submit_SQL") &") tabGroup" &_
		" where PrintStatus=1 group by UnitName,ChName"

set RsTemp=conn.execute(strSQL)
%>
<html>   
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>ExportBase</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style3 {font-size: 15px; mso-number-format:"\@";}
.style5 {
	font-size: 11px;
	color: #666666;
}
-->
</style>
</head>	 
<body>    
	<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" align="center" >
		<tr>
			<tr>
				 <td>&nbsp;</td>
			   <td colspan=6><center><span class="style1"><u><b>高雄市政府警察局列印件數統計表</b></u></span></center></td>
				 <td>&nbsp;</td>
			</tr>
			<tr>
				 <td>&nbsp;</td>
			   <td colspan=6><center>統計期間: <%=str_ActDate%></center></td>
				 <td>&nbsp;</td>
			</tr>		
			
		</tr>
	</table>
 <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#808080" >    
   <tr>
     <td bgcolor="#EBFBE3"><B><center>代印單位</center></B></td>
	 <td bgcolor="#EBFBE3"><B><center>代印人員</center></B></td>
     <td bgcolor="#EBFBE3"><B><center>列印件數</center></B></td>
   </tr>                     
<%
While Not RsTemp.Eof%>                             
   <tr>
     <td ><%=RsTemp("UnitName")%></td>
	 <td class="style3"><%=RsTemp("ChName")%></td>
     <td ><%=RsTemp("SumPrintCnt")%></td>
   </tr>   
<%
   RsTemp.MoveNext
Wend
%>     
 </table>    
 </body>
<!-- #include file="../Common/ClearObject.asp" --> 
 <%
	fMnoth=month(now)
	if fMnoth<10 then fMnoth="0"&fMnoth
	fDay=day(now)
	if fDay<10 then	fDay="0"&fDay
	fname=year(now)&fMnoth&fDay&"_列印件數統計表.xls"
	Response.AddHeader "Content-Disposition", "filename="&fname
	response.contenttype="application/x-msexcel; charset=MS950"
 %>
 </html>