<%
response.write "<td width=""800""><div align=""left"">"

'取得登入者
'strSQL="select UnitName from UnitInfo where UnitID='"& Session("Unit_ID")&"'"
'set rssysinfo=conn.execute(strSQL)
response.write "<img src=""../Image/space.gif"" width=""5"" ><img src=""../Image/dot.gif""></img> <span class=""font10"">"&Session("Ch_Name")
'response.write "<img src=""../Image/dot.gif"" ><font size=""13"">單位："&trim(rssysinfo("UnitName")) & "</font>"
'取得今日該使用者建檔與傳送數
sStartDate=gOutDT(ginitdt(now)) & " 00:00:00 "
sEndDate=gOutDT(ginitdt(now)) & " 23:59:59 "
strSQL="Select count(*) billcount from PasserBase Where RecordDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID") & " and RecordStateID <> -1 "

set rssysinfo=conn.execute(strSQL)

response.write "<img src=""../Image/space.gif"" width=""5"" ><img src=""../Image/dot.gif"" ></img><span class=""font10"">本日建檔 </span>" & rssysinfo("billcount") 
'strSQL="Select count(*) dcicount from DCILog Where RecordDate between TO_DATE('"&sStartDate&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&sEndDate&"','YYYY/MM/DD/HH24/MI/SS')" &" and RecordMemberID="& Session("User_ID")

	set rssysinfo=nothing

%>