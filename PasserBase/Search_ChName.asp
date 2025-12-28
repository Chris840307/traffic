<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<%

	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	chName=""

	For i = 1 to len(trim(Request("chName")))
		chName=chName&Mid(trim(Request("chName")),i,1)&"%"
	Next

	strSQL="select distinct Driver from PasserBase where recordstateid=0 and Driver like '%"&chName&"' and rownum<31"

	set rsmem=conn.execute(strSQL)
	strTable=""
	tx="\u0022"

	For i = 1 to 20 
		If rsmem.eof Then exit For 

		If strTable = "" Then strTable="<table border=0 bgcolor="&tx&"#FFFFFF"&tx&" width="&tx&"100"&tx&">"
		
		strTable=strTable&"<tr>"

		strTable=strTable&"<td onclick="&tx&"funCrtVale('"&Request("inObj")&"','"&Request("CrObj")&"','"&trim(rsmem("Driver"))&"');"&tx&" onMouseOver="&tx&"this.style.backgroundColor='#EEEEEE';"&tx&" onMouseOut="&tx&"this.style.backgroundColor='#FFFFFF';"&tx&" nowrap>"
		
		strTable=strTable&trim(rsmem("Driver"))

		strTable=strTable&"</tr>"

		rsmem.movenext	
	Next

	If not rsmem.eof Then strTable=strTable&"<tr><td>.......</td></tr>"

	rsmem.close

	If strTable <>"" Then strTable=strTable&"</table>"

	Response.Write "eval('"&Request("CrObj")&"').innerHTML="""&strTable&""";"
	conn.close
	set conn=nothing
%>
