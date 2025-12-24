<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getPeopleBillNoExist.asp
	'是否為單號輸入驗證
	chkGetBillFlag=0
	chkBillBaseFlag=0
	chkUnitFlag=0
	MLoginID=""
	MMemberID=""
	MMemName=""
	MUnitID=""
	MUnitName=""
	if IsChkGetBillFlag=1 then	'是否要檢查
		'檢查是否有在已領單單號內
		strChkGB="select GetBillSN from GetBillDetail where BillNo='"&trim(request("BillNo"))&"'"
		set rsChkGB=conn.execute(strChkGB)
		if not rsChkGB.eof then
			chkGetBillFlag=1
			'取得員警臂章號碼及單位
			strMemID="select b.LoginID,b.MemberID,b.ChName,c.UnitID,c.UnitName,c.UnitLevelID,c.UnitTypeID from GetBillBase a,MemberData b,UnitInfo c where a.GetBillSN="&trim(rsChkGB("GetBillSN"))&" and a.GetBillMemberID=b.MemberID and b.UnitID=c.UnitID and b.AccountStateID=0 and b.RecordstateID=0"
			set rsMem=conn.execute(strMemID)
			if not rsMem.eof then
				MLoginID=trim(rsMem("LoginID"))
				MMemberID=trim(rsMem("MemberID"))
				MMemName=trim(rsMem("ChName"))
				MUnitID=trim(rsMem("UnitID"))
				MUnitName=trim(rsMem("UnitName"))
				'If trim(rsMem("UnitID"))<>trim(Session("Unit_ID")) and trim(rsMem("UnitTypeID"))<>trim(Session("Unit_ID")) Then chkUnitFlag=1
				if Cint(rsMem("UnitLevelID"))>1 then
					if Cint(rsMem("UnitLevelID"))=2 then
						strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&trim(rsMem("UnitID"))&"' and UnitName like '%分局'"
					elseif Cint(rsMem("UnitLevelID"))>2 then
						strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&trim(rsMem("UnitTypeID"))&"' and UnitName like '%分局'"
					end if
					set rsUnit=conn.execute(strSQL)
					if Not rsUnit.eof then
						SUnitID=trim(rsUnit("UnitID"))
						SUnitName=trim(rsUnit("UnitName"))
					end if
					rsUnit.close
				end if
			end if
			rsMem.close
			set rsMem=nothing
		end if
		rsChkGB.close
		set rsChkGB=nothing
	else
		chkGetBillFlag=1
	end if

	chkBillSN=1
	'檢查是否已經在PasserBase裡

	strChkBB="select SN,BillNo,RecordMemberID from billbase where BillNo='"&trim(request("BillNo"))&"' and recordstateid<>-1 union all select SN,BillNo,RecordMemberID from PasserBase where BillNo='"&trim(request("BillNo"))&"' and recordstateid<>-1"

	set rsChkBB=conn.execute(strChkBB)
	if not rsChkBB.eof then
		if trim(rsChkBB("RecordMemberID"))=trim(Session("User_ID")) then
			chkBillBaseFlag=0
			chkBillSN=trim(rsChkBB("SN"))
		else
			chkBillBaseFlag=1
		end if
	else
		chkBillBaseFlag=2
	end if
	rsChkBB.close
	set rsChkBB=nothing
%>setCheckPeopleBillNoExist("<%=chkGetBillFlag%>","<%=chkBillBaseFlag%>","<%=chkUnitFlag%>","<%=chkBillSN%>","<%=MLoginID%>","<%=MMemberID%>","<%=MMemName%>","<%=MUnitID%>","<%=MUnitName%>","<%=SUnitID%>","<%=SUnitName%>");
<%
conn.close
set conn=nothing
%>
