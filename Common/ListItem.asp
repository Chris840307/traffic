<!--#include virtual="traffic/Common/DB.ini"-->
<%
insertTable="<table width=""645"" border=""1"" cellspacing=0 cellpadding=0>"
if trim(request("LoginID"))<>"" then
	
strSQL = "select UnitID,MemberID,ChName from MemberData where LoginID='"&request("LoginID")&"' and ACCOUNTSTATEID=0 and RECORDSTATEID=0 order by Chname"
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

	
strSQL = "select MemberID,ChName from MemberData where UnitID in('"&UnitID&"') and ACCOUNTSTATEID=0 and RECORDSTATEID=0 order by Chname"
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