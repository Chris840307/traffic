<!--#include virtual="traffic/Common/DB.ini"-->
<%
if trim(request("LoginID"))<>"" then
	
strSQL = "select UnitID,ChName,Max(MemberID) MemberID from MemberData where LoginID='"&request("LoginID")&"' and ACCOUNTSTATEID=-1 and RECORDSTATEID=0 group by UnitID,ChName order by Chname"
	set rscnt=conn.execute(strSQL)
	if Not rscnt.eof then
		UnitID=trim(rscnt("UnitID"))
		MemberID=trim(rscnt("MemberID"))
		response.write "for(i=0;i<document.all["""&request("UnitListName")&"""].length;i++){"		
		response.write "if(document.all["""&request("UnitListName")&"""][i].value=='"&trim(rscnt("UnitID"))&"'){"
		response.write "document.all["""&request("UnitListName")&"""][i].selected=true;"
		response.write "}"
		response.write "}"
	else
		response.write "document.all["""&request("UnitListName")&"""][0].selected=true;"
	end if
	rscnt.close

	
strSQL = "select ChName,MemberID from MemberData where UnitID in('"&UnitID&"') and ACCOUNTSTATEID=-1 and RECORDSTATEID=0 order by Chname,MemberID"
	set rscnt=conn.execute(strSQL)
	if Not rscnt.eof then
		chkMemID=0:i=0
		while Not rscnt.eof
			i=i+1
			response.write "document.all['"&request("MemListName")&"'].options["&i&"]=new Option('"&rscnt("ChName")&"','"&rscnt("MemberID")&"');"& vbcrlf
			if trim(MemberID)<>"" then
				if trim(MemberID)=trim(rscnt("MemberID")) then chkMemID=i
			end if
			rscnt.movenext
		wend
		response.write "document.all['"&request("MemListName")&"'].options["&chkMemID&"].selected=true;"& vbcrlf
		response.write "document.all['"&request("MemListName")&"'].length="&i+1&";"& vbcrlf
	else
		response.write "document.all['"&request("MemListName")&"'].options[1]=new Option('©Ò¦³¤H','');"& vbcrlf
		response.write "document.all['"&request("MemListName")&"'].length=1;"
	end if
	rscnt.close
	conn.close
end if
%>