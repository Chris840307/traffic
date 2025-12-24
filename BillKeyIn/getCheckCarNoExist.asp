<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getCheckCarNoExist.asp
	'一打一驗檢查車號是否已建檔
	chkCarNoFlag=0

	'檢查是否已經在BillBase裡
	strChkBB="select SN from BillBase where BillTypeID='2' and CarNo='"&trim(request("CarID"))&"' and RecordStateID=0"
	set rsChkBB=conn.execute(strChkBB)
	if not rsChkBB.eof then
		chkCarNoFlag=1
	else
		chkCarNoFlag=0
	end if
	rsChkBB.close
	set rsChkBB=nothing
%>setCheckCarNoExist("<%=chkCarNoFlag%>");
<%
conn.close
set conn=nothing
%>
