<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<%
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
	rsCity.close

	strSQL="select a.* from (select driver,driverid,DriverSex,DriverZip,driveraddress,IllegalAddressID,illegaladdress,illegaldate,DriverBirth from passerbase where driverID='"&request("DriverPID")&"' and recordstateid=0) a,(select driverID,max(illegaldate) illegaldate from passerbase where driverID='"&request("DriverPID")&"' and recordstateid=0 group by driverID) b where a.illegaldate=b.illegaldate"
	set rs=conn.execute(strSQL)
	if Not rs.eof then
		response.write "myForm.DriverName.value='"&trim(rs("Driver"))&"';"
		response.write "myForm.DriverBrith.value='"&ginitdt(trim(rs("DriverBirth")))&"';"
		response.write "myForm.DriverPID.value='"&trim(rs("Driverid"))&"';"
		response.write "myForm.DriverZip.value='"&trim(rs("DriverZip"))&"';"
		response.write "myForm.DriverAddress.value='"&trim(rs("DriverAddress"))&"';"
		response.write "myForm.DriverSEX.value='"&trim(rs("DriverSex"))&"';"

		'If sys_City<>"­]®ß¿¤" then
		'	response.write "myForm.IllegalAddressID.value='"&trim(rs("IllegalAddressID"))&"';"
		'	response.write "myForm.IllegalAddress.value='"&trim(rs("IllegalAddress"))&"';"
		'end if
	'else
		'response.write "myForm.DriverBrith.value='';"
		'response.write "myForm.DriverZip.value='';"
		'response.write "myForm.DriverAddress.value='';"
		'response.write "myForm.IllegalAddressID.value='';"
		'response.write "myForm.IllegalAddress.value='';"
	end if
%>
