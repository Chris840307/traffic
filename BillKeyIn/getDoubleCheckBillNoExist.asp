<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getDoubleCheckBillNoExist.asp
	'一打一驗單號輸入驗證
	chkGetBillFlag=0
	chkBillBaseFlag=0
	chkBillBaseTmpFlag=0
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
			strMemID="select b.LoginID,b.MemberID,b.ChName,c.UnitID,c.UnitName from GetBillBase a,MemberData b,UnitInfo c where a.GetBillSN="&trim(rsChkGB("GetBillSN"))&" and a.GetBillMemberID=b.MemberID and b.UnitID=c.UnitID"
			set rsMem=conn.execute(strMemID)
			if not rsMem.eof then
				MLoginID=trim(rsMem("LoginID"))
				MMemberID=trim(rsMem("MemberID"))
				MMemName=trim(rsMem("ChName"))
				MUnitID=trim(rsMem("UnitID"))
				MUnitName=trim(rsMem("UnitName"))
			end if
			rsMem.close
			set rsMem=nothing
		end if
		rsChkGB.close
		set rsChkGB=nothing
	else
		chkGetBillFlag=1
	end if

	'檢查是否已經在BillBase裡
	if chkGetBillFlag=1 then
		strChkBB="select SN from BillBase where BillNo='"&trim(request("BillNo"))&"' and RecordStateID=0"
		set rsChkBB=conn.execute(strChkBB)
		if not rsChkBB.eof then
			chkBillBaseFlag=1
		else
			chkBillBaseFlag=0
		end if
		rsChkBB.close
		set rsChkBB=nothing
	end if

	'檢查是否已經在BillBaseTmp裡
	if chkBillBaseFlag=1 then
		strChkBB="select SN,RecordMemberID from BillBaseTmp where BillNo='"&trim(request("BillNo"))&"' and RecordStateID=0"
		set rsChkBB=conn.execute(strChkBB)
		if not rsChkBB.eof then
			chkBillBaseTmpFlag=0
		else
			chkBillBaseTmpFlag=1
		end if
		rsChkBB.close
		set rsChkBB=nothing
	end if
%>setCheckDoubleBillNoExist("<%=chkGetBillFlag%>","<%=chkBillBaseFlag%>","<%=chkBillBaseTmpFlag%>","<%=MLoginID%>","<%=MMemberID%>","<%=MMemName%>","<%=MUnitID%>","<%=MUnitName%>");
<%
conn.close
set conn=nothing
%>
