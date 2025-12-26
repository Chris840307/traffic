<!--#include virtual="traffic/Common/DB.ini"-->
<%
if trim(request("UnitID"))<>"" then
	
strSQL = "select MemberID,ChName from MemberData where UnitID='"&request("UnitID")&"' order by Chname"
	set rscnt=conn.execute(strSQL)
	if Not rscnt.eof then
		chkMemID=0:i=0
		while Not rscnt.eof
			i=i+1
			response.write "document.all['"&request("UnitMem")&"'].options["&i&"]=new Option('"&rscnt("ChName")&"','"&rscnt("MemberID")&"');"& vbcrlf
			if trim(request("MemberID"))<>"" then
				if trim(request("MemberID"))=trim(rscnt("MemberID")) then chkMemID=i
			end if
			rscnt.movenext
		wend
		response.write "document.all['"&request("UnitMem")&"'].options["&chkMemID&"].selected=true;"& vbcrlf
		response.write "document.all['"&request("UnitMem")&"'].length="&i+1&";"& vbcrlf
	else
		response.write "document.all['"&request("UnitMem")&"'].options[1]=new Option('©Ò¦³¤H','');"& vbcrlf
		response.write "document.all['"&request("UnitMem")&"'].length=1;"
	end if
	rscnt.close
	conn.close
end if
%>
