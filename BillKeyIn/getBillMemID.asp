<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getBillMemID.asp
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	UserUnitTypeID=trim(Session("Unit_ID"))
	'使用者所屬上曾單位
	strUT="select UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"'"
	set rsUT=conn.execute(strUT)
	if not rsUT.eof then
		UserUnitTypeID=trim(rsUT("UnitTypeID"))
	end if
	rsUT.close
	set rsUT=nothing

	'舉發人臂章號碼
	if trim(request("MType"))="Car" or trim(request("MType"))="CarS" then
		strMem="select a.ChName,a.CreditID,a.MemberID,a.UnitID,b.UnitName,b.UnitTypeID from MemberData a,UnitInfo b where a.UnitID=b.UnitID and a.LoginID='"&trim(request("MemID"))&"' and a.AccountStateID=0 and a.RecordstateID=0"
	elseif trim(request("MType"))="People" and sys_City<>"高雄縣" then
		strMem="select a.ChName,a.CreditID,a.MemberID,a.LoginID,a.UnitID,b.UnitName,b.UnitTypeID from MemberData a,UnitInfo b where a.UnitID=b.UnitID and a.ChName like '"&trim(request("MemID"))&"' and a.AccountStateID=0 and a.RecordstateID=0"
	else
		strMem="select a.ChName,a.CreditID,a.MemberID,a.LoginID,a.UnitID,b.UnitName,b.UnitTypeID from MemberData a,UnitInfo b where a.UnitID=b.UnitID and a.LoginID='"&trim(request("MemID"))&"' and a.AccountStateID=0 and a.RecordstateID=0"
	end if

	cnt=0
	set rsMem=conn.execute(strMem)
	while not rsMem.eof
		cnt=cnt+1
		If cnt=1 Then
			MemName=trim(rsMem("ChName"))
			if trim(request("MType"))="People" and (sys_City="高雄縣" or sys_City="高雄市" Or sys_City=ApconfigureCityName) then MemName=trim(rsMem("LoginID"))
			MemCreditID=trim(rsMem("CreditID"))
			MemID=trim(rsMem("MemberID"))
			MemUnitID=trim(rsMem("UnitID"))
			MemUnitName=trim(rsMem("UnitName"))
			'if trim(request("MType"))="CarS" then 
				MemUnitTypeID=trim(rsMem("UnitTypeID"))
			'end if
		end if
		rsMem.movenext
	wend
	rsMem.close
	set rsMem=nothing
	If cnt>1 and sys_City<>"台東縣" Then response.write "alert('姓名重覆請按『F5』來選擇警員!!');"
	If cnt=0 and trim(request("MType"))="People" then
		If len(trim(request("MemID")))>2 Then
			response.write "alert('無此員警資料，請確認人員管理是否有該資料紀錄!!');"
		End if
	end if
	UTypeFlag=0
	if sys_City<>"台中市" and (instr(MemUnitName,"保安")<1) then
		if MemUnitTypeID<>UserUnitTypeID then
			UTypeFlag=1
		else
			UTypeFlag=0
		end if
	else
		UTypeFlag=0
	end if

	'雲林縣攔停要判斷是否同分局
	if trim(request("MType"))="CarS" and sys_City="雲林縣" then 
%>
setMemName2("<%=trim(request("MType"))%>","<%=trim(request("MemOrder"))%>","<%=MemName%>","<%=MemID%>","<%=MemUnitID%>","<%=MemUnitName%>","<%=MemUnitTypeID%>","<%=UTypeFlag%>");
<%
	else
%>
setMemName("<%=trim(request("MType"))%>","<%=trim(request("MemOrder"))%>","<%=MemName%>","<%=MemID%>","<%=MemUnitID%>","<%=MemUnitName%>","<%=UTypeFlag%>");
<%	
	end if
conn.close
set conn=nothing
%>
