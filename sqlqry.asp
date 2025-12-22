<!--#include virtual="TAKECAR/Common/DB.ini"-->

<FORM name="myForm" METHOD="POST">
<body bgcolor="#FEFFCD">
<input type="text" name="t1" value="<%=request("t1")%>">&nbsp;&nbsp;&nbsp;<input type="text" name="F1" value="12">
<br>
<TEXTAREA NAME="bigtext" ROWS="20" COLS="100"></TEXTAREA>
<br>
<input type="button" name="b1" value="dooooooooooooooooooooooooooooooooooooooooo" onclick="myForm.submit()">
<input type="button" name="b2" value="clear" onclick="location.href='sqlqry.asp'">


<table border="1" cellspacing="0" cellpadding="0">
<%

if request("t1")="qry" then
	strsql=request("bigtext")
	response.write "<td colspan=""20"" style=""font-size:10pt"">"&strsql&"</td>"
	response.write "<tr style='cursor:hand;' onmouseover=""this.style.backgroundColor='#91FEFF';"" onmouseout=""this.style.backgroundColor=''"";"">"
	i=0
	set rs=conn.execute(strsql)
	while not rs.eof

		i=i+1
		if i=1 then 
			for j=0 to rs.fields.count-1
				response.write "<td  style=""font-size:"&request("f1")&"pt"">"&rs.fields(j).name&"</td>"
			next
	response.write "<tr style='cursor:hand;' onmouseover=""this.style.backgroundColor='#91FEFF';"" onmouseout=""this.style.backgroundColor=''"";"">"
		else
	response.write "<tr style='cursor:hand;' onmouseover=""this.style.backgroundColor='#91FEFF';"" onmouseout=""this.style.backgroundColor=''"";"">"
		end if
		for j=0 to 	rs.fields.count-1
				response.write "<td style=""font-size:"&request("f1")&"pt"">&nbsp;"&rs.fields(j).value&"</td>"
			
		next
					
		rs.movenext
	wend

end if

if request("t1")="execute" then 
	response.write "<td>"&strsql&"</td>"
	response.write "<tr style='cursor:hand;' onmouseover=""this.style.backgroundColor='#91FEFF';"" onmouseout=""this.style.backgroundColor=''"";"">"
	strsql=request("bigtext")
	conn.execute(strsql)
    response.write "<td>OK</td>"
end if
conn.close()
%>
</FORM>
</html>