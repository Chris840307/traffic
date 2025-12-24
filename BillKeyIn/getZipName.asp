<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<%
	
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	strSQL="select ZipName from Zip where ZipID = '"&trim(request("ZipID"))&"'"
	set rszip=conn.execute(strSQL)
	if Not rszip.eof Then
		response.write "myForm."&trim(Request("getZipName"))&".value='"&rszip("ZipName")&"'+myForm."&trim(Request("getZipName"))&".value.replace('"&rszip("ZipName")&"','');"
	else
		response.write "alert(""¥N½X¿ù»~½Ð½T»{¡C"");"
	end if
	rszip.close
%>
