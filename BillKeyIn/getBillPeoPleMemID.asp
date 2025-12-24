<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getBillMemID.asp
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	chkMemid=1

	 if asc(mid(trim(request("MemID")),2,1)) > 32 and asc(mid(trim(request("MemID")),2,1)) < 127 then
		chkMemid=0
	 end if

	'舉發人臂章號碼
	if trim(request("MType"))="People" and chkMemid=1 then
		strMem="select a.ChName,a.CreditID,a.MemberID,a.LoginID,a.UnitID,b.UnitName from MemberData a,UnitInfo b where a.UnitID=b.UnitID and a.ChName like '"&trim(request("MemID"))&"' and a.AccountStateID=0 and a.RecordstateID=0"
	else
		strMem="select a.ChName,a.CreditID,a.MemberID,a.LoginID,a.UnitID,b.UnitName from MemberData a,UnitInfo b where a.UnitID=b.UnitID and a.LoginID='"&trim(request("MemID"))&"' and a.AccountStateID=0 and a.RecordstateID=0"
	end if
 
	cnt=0
	set rsMem=conn.execute(strMem)
	while not rsMem.eof
		cnt=cnt+1
		If cnt=1 Then						
			LogInID=trim(rsMem("LoginID"))
			if trim(request("MType"))="People" and chkMemid=0 then LogInID=trim(rsMem("ChName"))
			MemName=trim(rsMem("ChName"))
			MemCreditID=trim(rsMem("CreditID"))
			MemID=trim(rsMem("MemberID"))
			MemUnitID=trim(rsMem("UnitID"))
			MemUnitName=trim(rsMem("UnitName"))
			if trim(request("MType"))="CarS" then 
				MemUnitTypeID=trim(rsMem("UnitTypeID"))
			end if
		end if
		rsMem.movenext
	wend
	rsMem.close
	set rsMem=nothing
	If cnt>1 Then response.write "alert('姓名重覆請按『F5』來選擇警員!!');"
	'If cnt=0 and trim(request("MType"))="People" then
		'If len(trim(request("MemID")))>2 Then
		'	response.write "alert('無此員警資料，請確認人員管理是否有該資料紀錄!!');"
		'End if
	'end if
%>
setPeoPleMemName("<%=trim(request("MType"))%>","<%=trim(request("MemOrder"))%>","<%=MemName%>","<%=MemID%>","<%=LogInID%>","<%=MemUnitID%>","<%=MemUnitName%>");
<%
conn.close
set conn=nothing
%>
