<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getIllStreet.asp
	'違規地點
	'strAddress="select Address from Street where StreetID='"&trim(request("illAddrID"))&"'"
	strAddress="select Address from Street where StreetID='00012'"
	set rsAddress=conn.execute(strAddress)
	if not rsAddress.eof then
		AddressName=trim(rsAddress("Address"))
	end if
	rsAddress.close
	set rsAddress=nothing


%>

function setIllStreetName(AddressName){
	if (AddressName!=""){
		myForm.IllegalAddressQry.value=AddressName;
	}else{
		myForm.IllegalAddressQry.value="";
	}
}
setIllStreetName("<%=AddressName%>");
<%		

conn.close
set conn=nothing
%>
