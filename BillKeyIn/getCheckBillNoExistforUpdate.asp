<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getCheckBillNoExist.asp
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	'是否為單號輸入驗證
	chkGetBillFlag=0
	chkBillBaseFlag=0
	chkGetBillMemUnit=0
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

			if Session("UnitLevelID")="1" then
				strUnit="select * from UnitInfo where " &_
					" UnitLevelID=1 and UnitID='"&MUnitID&"'"
			else
				strUnit="select * from UnitInfo where " &_
					"(UnitID='"&trim(Session("Unit_ID"))&"' or UnitTypeID='"&trim(Session("Unit_ID"))&"')" &_
					" and UnitID='"&MUnitID&"'"
			end if
				set rsUnit=conn.execute(strUnit)
				if not rsUnit.eof then
					chkGetBillMemUnit=0
				else
					chkGetBillMemUnit=1
				end if
				rsUnit.close
				set rsUnit=nothing
			end if
			rsMem.close
			set rsMem=nothing
		end if
		rsChkGB.close
		set rsChkGB=nothing
	else
		chkGetBillFlag=0
	end if

	chkBillSN=1
	chkBillTypeID=0
	'檢查是否已經在BillBase裡
		strChkBB="select * from BillBase where BillNo='"&trim(request("BillNo"))&"' and RecordStateID=0"
		set rsChkBB=conn.execute(strChkBB)
		if not rsChkBB.eof then
			if trim(rsChkBB("RecordMemberID"))=trim(Session("User_ID")) then
				chkBillBaseFlag=0
				chkBillSN=trim(rsChkBB("SN"))
				chkBillTypeID=trim(rsChkBB("BillTypeID"))
			else
				chkBillBaseFlag=1
			end if
			RecMem2=""
			IllDate2=""
			IllRuleB1=""
			IllRuleB2=""
			strRecMem="select a.UnitName from UnitInfo a,MemberData b where a.UnitID=b.UnitID and b.MemberID="&trim(rsChkBB("RecordMemberID"))
			set rsRecMem=conn.execute(strRecMem)
			if not rsRecMem.eof then
				RecMem2=trim(rsRecMem("UnitName"))
			end if
			rsRecMem.close
			set rsRecMem=nothing
			IllDate2=year(trim(rsChkBB("IllegalDate")))-1911&"/"&month(trim(rsChkBB("IllegalDate")))&"/"&day(trim(rsChkBB("IllegalDate")))&" "&hour(trim(rsChkBB("IllegalDate")))&":"&minute(trim(rsChkBB("IllegalDate")))
			IllRuleB1=trim(rsChkBB("Rule1"))
			IllRuleB2=trim(rsChkBB("Rule2"))
		else
			chkBillBaseFlag=2
		end if
		rsChkBB.close
		set rsChkBB=nothing
%>setCheckBillNoExist("<%=chkGetBillFlag%>","<%=chkBillBaseFlag%>","<%=chkBillSN%>","<%=chkBillTypeID%>","<%=MLoginID%>","<%=MMemberID%>","<%=MMemName%>","<%=MUnitID%>","<%=MUnitName%>","<%=RecMem2%>","<%=IllDate2%>","<%=IllRuleB1%>","<%=IllRuleB2%>","<%=chkGetBillMemUnit%>");
<%
conn.close
set conn=nothing
%>
//檢查單號是否有在GETBILLBASE內
function setCheckBillNoExist(GetBillFlag,BillBaseFlag,BillSN,BillType,MLoginID,MMemberID,MMemName,MUnitID,MUnitName,RecMem2,IllDate2,IllRuleB1,IllRuleB2,chkGetBillMemUnit)
{
	if (GetBillFlag==0){
	<%if sys_City="嘉義縣" or sys_City="宜蘭縣" or sys_City="台東縣" or sys_City="台南縣" then%>
		alert("此單號不存在於領單紀錄中！");
	<%end if%>
		//document.myForm.Billno1.value="";
	}else{
		//if (document.myForm.BillMem1.value==""){
		//	document.myForm.BillMem1.value=MLoginID;
		//	document.myForm.BillMemID1.value=MMemberID;
		//	document.myForm.BillMemName1.value=MMemName;
		//	Layer12.innerHTML=MMemName;
		//	TDMemErrorLog1=0;
		//}
		//if (document.myForm.BillUnitID.value==""){
		//	document.myForm.BillUnitID.value=MUnitID;
		//	Layer6.innerHTML=MUnitName;
		//	TDUnitErrorLog=0;
		//}
	}
	if (BillBaseFlag==1){
		alert("此單號已建檔！！\n建檔單位："+RecMem2+"\n違規時間："+IllDate2+"\n違規法條："+IllRuleB1+" "+IllRuleB2);
		document.myForm.Billno1.select();
	}else if (BillBaseFlag==0){
		alert("此單號已建檔！！\n建檔單位："+RecMem2+"\n違規時間："+IllDate2+"\n違規法條："+IllRuleB1+" "+IllRuleB2);
		document.myForm.Billno1.select();
	}else if (chkGetBillMemUnit==1){
	<%if sys_City<>"台南市" then%>
		alert("建檔單位非領單單位！！");
	<%end if%>
	}
}