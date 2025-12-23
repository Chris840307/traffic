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
			strMemID="select b.LoginID,b.MemberID,b.ChName,c.UnitID,c.UnitName,b.AccountStateID,b.RecordStateID from GetBillBase a,MemberData b,UnitInfo c where a.GetBillSN="&trim(rsChkGB("GetBillSN"))&" and a.GetBillMemberID=b.MemberID and b.UnitID=c.UnitID"
			set rsMem=conn.execute(strMemID)
			if not rsMem.eof then
				MLoginID=trim(rsMem("LoginID"))
				MMemberID=trim(rsMem("MemberID"))
				MMemName=trim(rsMem("ChName"))
				MUnitID=trim(rsMem("UnitID"))
				MUnitName=trim(rsMem("UnitName"))
				if trim(rsMem("AccountStateID"))="-1" or trim(rsMem("RecordStateID"))="-1" then
					MStateID=-1
				else
					MStateID=0
				end if
			if Session("UnitLevelID")="1" then
				if sys_City="嘉義縣" then
					strUnit="select * from UnitInfo where " &_
					"(UnitID='"&trim(Session("Unit_ID"))&"' or UnitTypeID='"&trim(Session("Unit_ID"))&"')" &_
					" and UnitID='"&MUnitID&"'"
				else
					strUnit="select * from UnitInfo where " &_
						" UnitLevelID=1 and UnitID='"&MUnitID&"'"
				end if
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
		strChkBB="select SN,BillTypeID,IllegalDate,Rule1,Rule2,RecordMemberID from BillBase " &_
			" where BillNo='"&trim(request("BillNo"))&"' and RecordStateID=0 " &_
			" union all select SN,BillTypeID,IllegalDate,Rule1,Rule2,RecordMemberID from PasserBase " &_
			" where BillNo='"&trim(request("BillNo"))&"' and RecordStateID=0"
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
		'刪除失敗或未回傳的不能建
		strDci="select * from Dcilog where BillNo='"&trim(request("BillNo"))&"' and ExchangeTypeID='E' " &_
			" and (DciReturnStatusID<>'S' or DciReturnStatusID is null)"
		set rsDci=conn.execute(strDci)
		if not rsDci.eof then
			if isnull(rsDci("DciReturnStatusID")) then
				chkBillBaseFlag=3
			else
				chkBillBaseFlag=4
			end if
		end if
		rsDci.close
		set rsDci=nothing

%>setCheckBillNoExist("<%=chkGetBillFlag%>","<%=chkBillBaseFlag%>","<%=chkBillSN%>","<%=chkBillTypeID%>","<%=MLoginID%>","<%=MMemberID%>","<%=MMemName%>","<%=MUnitID%>","<%=MUnitName%>","<%=RecMem2%>","<%=IllDate2%>","<%=IllRuleB1%>","<%=IllRuleB2%>","<%=chkGetBillMemUnit%>","<%=MStateID%>");
<%
conn.close
set conn=nothing
%>
//檢查單號是否有在GETBILLBASE內
function setCheckBillNoExist(GetBillFlag,BillBaseFlag,BillSN,BillType,MLoginID,MMemberID,MMemName,MUnitID,MUnitName,RecMem2,IllDate2,IllRuleB1,IllRuleB2,chkGetBillMemUnit,MStateID)
{
	if (GetBillFlag==0){
	<%if sys_City="嘉義縣" or sys_City="嘉義市" or sys_City="宜蘭縣" or sys_City="台東縣" or sys_City="台南縣" or sys_City="彰化縣" or sys_City="高雄縣" or sys_City="高雄市" then%>
		alert("此單號不存在於領單紀錄中！");
	<%ElseIf sys_City="雲林縣" then %>
		document.myForm.BillMem1.value="";
		document.myForm.BillMemID1.value="";
		document.myForm.BillMemName1.value="";
		Layer12.innerHTML="";
	<%end if%>
		//document.myForm.Billno1.value="";
	}else{
		//if (document.myForm.BillMem1.value==""){
			document.myForm.BillMem1.value=MLoginID;
			document.myForm.BillMemID1.value=MMemberID;
			document.myForm.BillMemName1.value=MMemName;
			Layer12.innerHTML=MMemName;
			TDMemErrorLog1=0;
		//}
		//if (document.myForm.BillUnitID.value==""){
			document.myForm.BillUnitID.value=MUnitID;
			Layer6.innerHTML=MUnitName;
			TDUnitErrorLog=0;
		//}
	}
	if (BillBaseFlag==1){
		alert("此單號已建檔！！\n建檔單位："+RecMem2+"\n違規時間："+IllDate2+"\n違規法條："+IllRuleB1+" "+IllRuleB2);
		document.myForm.Billno1.select();
	}else if (BillBaseFlag==0){
		alert("此單號已建檔！！\n建檔單位："+RecMem2+"\n違規時間："+IllDate2+"\n違規法條："+IllRuleB1+" "+IllRuleB2);
		document.myForm.Billno1.select();
<%if sys_City<>"台中縣" then%>
	}else if (BillBaseFlag==3){
		alert("此單號，監理站尚未回傳刪除成功，請至 ' 上傳下載查詢系統 ' 確認刪除後再建檔！！");
		document.myForm.Billno1.select();
	}else if (BillBaseFlag==4){
		alert("此單號，監理站回傳刪除異常，請至 ' 上傳下載查詢系統 ' 確認刪除後再建檔！！");
		document.myForm.Billno1.select();
<%end if%>
	}else if (chkGetBillMemUnit==1){
	<%if sys_City<>"台南市" and sys_City<>"台中市" then%>
		alert("建檔單位非領單單位！！");
		<%if sys_City="嘉義縣" or sys_City="高雄縣" then%>
			document.myForm.Billno1.select();
		<%end if%>
	<%end if%>
	}else if (MStateID==-1){
		alert("此領單人員帳號已停用，請至 ' 領單管理系統 ' 確認領單人及單位是否正確!");
	}
}