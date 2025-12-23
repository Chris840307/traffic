<!--#include virtual="traffic/Common/DB.ini"-->
<%
if trim(session("ManagerPower"))="1" then
	if trim(request("UnitID"))<>"" then
		
strSQL = "select MemberID,ChName from MemberData where UnitID='"&request("UnitID")&"' order by ChName"
		set rscnt=conn.execute(strSQL)
		if Not rscnt.eof then
			chkMemID=0:i=0
			while Not rscnt.eof
				i=i+1
				response.write "myForm."&request("objName")&".options["&i&"]=new Option('"&rscnt("ChName")&"','"&rscnt("MemberID")&"');"& vbcrlf
				if trim(request("MemberID"))=trim(rscnt("MemberID")) then chkMemID=i
				rscnt.movenext
			wend
			response.write "myForm."&request("objName")&".options["&chkMemID&"].selected=true;"& vbcrlf
			response.write "myForm."&request("objName")&".length="&i+1&";"& vbcrlf
		else
			response.write "myForm."&request("objName")&".options[0]=new Option('請選取','');"& vbcrlf
			response.write "myForm."&request("objName")&".length=1;"
		end if
		rscnt.close
		conn.close
	else
		response.write "myForm."&request("objName")&".options[0]=new Option('請選取','');"& vbcrlf
		response.write "myForm."&request("objName")&".length=1;"
	end if
end if
%>
