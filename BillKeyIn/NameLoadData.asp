<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<%
	strSQL="select a.* from (select driver,driverid,DriverZip,driveraddress,IllegalAddressID,illegaladdress,illegaldate,DriverBirth from passerbase where driver='"&request("ChName")&"') a,(select driverID,max(illegaldate) illegaldate from passerbase where driver='"&request("ChName")&"' group by driverID) b where a.illegaldate=b.illegaldate"
	set rs=conn.execute(strSQL)
	if Not rs.eof then
		response.write "myForm.DriverName.value='"&trim(rs("Driver"))&"';"
		response.write "myForm.DriverBrith.value='"&ginitdt(trim(rs("DriverBirth")))&"';"
		response.write "myForm.DriverZip.value='"&trim(rs("DriverZip"))&"';"
		response.write "myForm.DriverAddress.value='"&trim(rs("DriverAddress"))&"';"
		response.write "myForm.IllegalAddressID.value='"&trim(rs("IllegalAddressID"))&"';"
		response.write "myForm.IllegalAddress.value='"&trim(rs("IllegalAddress"))&"';"
	end if
%>
