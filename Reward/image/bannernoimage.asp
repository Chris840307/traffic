<%
response.write "<table width=""100%""  border=""0"" cellspacing=""0"" cellpadding=""0"">"
'response.write "<tr>"
'response.write "<td background=""../Image/banner.jpg"" height=""57"" colspan=""5"">&nbsp;</td>"
'response.write "</tr>"
response.write "<tr>"
strSQL="select UnitName from UnitInfo where UnitID='"& Session("Unit_ID")&"'"
set rs=conn.execute(strSQL)
response.write "<td width=""650""><div align=""left""><img src=""../Image/dot.gif""> 登入者："&Session("Ch_Name")
response.write "<img src=""../Image/dot.gif"" >單　位："&trim(rs("UnitName"))
sStartDate=gOutDT(ginitdt(now)) & " 00:00:00 "
sEndDate=gOutDT(ginitdt(now)) & " 23:59:59 "
strSQL="Select count(*) billcount from BillBase Where RecordDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID")

set rs=conn.execute(strSQL)

response.write "<img src=""../Image/dot.gif"" >本日建檔數："& rs("billcount") 
strSQL="Select count(*) dcicount from DCILog Where RecordDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID")
set rs=conn.execute(strSQL)
response.write "<img src=""../Image/dot.gif"">傳送數："& rs("dcicount") 


strSQL="select CreateDate from DCICreateFileLog where CreateDate  between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS') " & " and Rownum<10 order by CreateDate desc "
set rs=conn.execute(strSQL)
'response.write strsql
if not rs.eof then 
	response.write "<img src=""../Image/dot.gif"">DCI回傳："& trim(rs("CreateDate")) & "</div></td>"
end if
response.write "<td><div align=""right""><input type=""button"" name=""MoveUp"" value=""功能選單"" onclick=""location='../Traffic_Web_Main.asp'""><img src=""../Image/space.gif"" width=""10""><input type=""button"" name=""MoveUp"" value=""登 出"" onclick=""location='../UserLogout_Contral.asp'""></div></td>"
response.write "</tr>"
response.write "</table>"
%>