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
	MZipCode=""
	
	'2021/9/23台中市攔停修改畫面領單檢查不跳錯誤訊息
	TCChkGetBillFlag=1
	If Trim(request("IsCheckGet"))="No" Then
		TCChkGetBillFlag=0	'JS不跳錯誤訊息
	End If 
	
	if IsChkGetBillFlag=1 then	'是否要檢查
		'檢查是否有在已領單單號內
		strChkGB="select GetBillSN from GetBillDetail where BillNo='"&trim(request("BillNo"))&"'"
		set rsChkGB=conn.execute(strChkGB)
		if not rsChkGB.eof then
			chkGetBillFlag=1
			'取得員警臂章號碼及單位
		if sys_City="台中市" then
			strMemID="select b.LoginID,b.MemberID,b.ChName,c.UnitID,c.UnitName,b.AccountStateID,b.RecordStateID,c.ZipCode from GetBillBase a,MemberData b,UnitInfo c where a.GetBillSN="&trim(rsChkGB("GetBillSN"))&" and a.GetBillMemberID=b.MemberID and b.UnitID=c.UnitID"
		Else
			strMemID="select b.LoginID,b.MemberID,b.ChName,c.UnitID,c.UnitName,b.AccountStateID,b.RecordStateID from GetBillBase a,MemberData b,UnitInfo c where a.GetBillSN="&trim(rsChkGB("GetBillSN"))&" and a.GetBillMemberID=b.MemberID and b.UnitID=c.UnitID"
		End If 
			set rsMem=conn.execute(strMemID)
			if not rsMem.eof then
				MLoginID=trim(rsMem("LoginID"))
				MMemberID=trim(rsMem("MemberID"))
				MMemName=trim(rsMem("ChName"))
				MUnitID=trim(rsMem("UnitID"))
				MUnitName=trim(rsMem("UnitName"))
				if sys_City="台中市" Then
					MZipCode=trim(rsMem("ZipCode"))
				End If 
				if trim(rsMem("AccountStateID"))="-1" or trim(rsMem("RecordStateID"))="-1" then
					MStateID=-1
				else
					MStateID=0
				end if
			if Session("UnitLevelID")="1" then
				if sys_City="嘉義縣" Or sys_City="澎湖縣" then
					strUnit="select * from UnitInfo where " &_
					"(UnitID='"&trim(Session("Unit_ID"))&"' or UnitTypeID='"&trim(Session("Unit_ID"))&"')" &_
					" and UnitID='"&MUnitID&"'"
				else
					strUnit="select * from UnitInfo where " &_
						" UnitLevelID=1 and UnitID='"&MUnitID&"'"
				end if
			else
				strUnit="select * from UnitInfo where " &_
					" UnitTypeID in (select UnitTypeID from UnitInfo where UnitID='"&trim(Session("Unit_ID"))&"')" &_
					" and UnitID='"&MUnitID&"'"
			end if
				set rsUnit=conn.execute(strUnit)
				if not rsUnit.eof then
					chkGetBillMemUnit=0
				Else
					'南投交通隊不檢查集集@@@@@@@@@@@@@@@@@
					If sys_City="南投縣" And trim(Session("Unit_ID"))="05A7" And (Trim(MUnitID)="05G3" or Trim(MUnitID)="05G4" Or Trim(MUnitID)="05G5" Or Trim(MUnitID)="05G8" Or Trim(MUnitID)="05G9" Or Trim(MUnitID)="05GB" Or Trim(MUnitID)="05GD" Or Trim(MUnitID)="05GF" Or Trim(MUnitID)="05G7" Or Trim(MUnitID)="05G2" Or Trim(MUnitID)="05G6" Or Trim(MUnitID)="05GA" Or Trim(MUnitID)="05G0" Or Trim(MUnitID)="05G1" Or Trim(MUnitID)="05GG" Or Trim(MUnitID)="05GC") Then
						chkGetBillMemUnit=0
					Else
						chkGetBillMemUnit=1
					End If 
					'chkGetBillMemUnit=1
				end if
				rsUnit.close
				set rsUnit=nothing
			end if
			rsMem.close
			set rsMem=nothing
		end if
		rsChkGB.close
		set rsChkGB=Nothing
		If sys_City="台南市" And Left(trim(request("BillNo")),2)="SY" Then
			chkGetBillFlag=1
		End If 
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
		strDci="select * from (select DciReturnStatusID from (select * from Dcilog where BillNo='"&trim(request("BillNo"))&"' and ExchangeTypeID='E' order by ExchangeDate Desc) where Rownum<=1) where (DciReturnStatusID<>'S' or DciReturnStatusID is null)"
		'strDci="select * from Dcilog where BillNo='"&trim(request("BillNo"))&"' and ExchangeTypeID='E' " &_
		'	" and (DciReturnStatusID<>'S' or DciReturnStatusID is null)"
		set rsDci=conn.execute(strDci)
		if not rsDci.eof then
			if isnull(rsDci("DciReturnStatusID")) then
				chkBillBaseFlag=3
			else
				chkBillBaseFlag=4
			end if
		end if
		rsDci.close
		set rsDci=Nothing
		
		If sys_City="台中市" then
			If chkBillBaseFlag=2 And Trim(request("AcceptBatchNumber"))<>"" Then
				strSQL1="select * from BillStopCarAccept where Batchnumber='"&Trim(request("AcceptBatchNumber"))&"'" &_
					" and BillNo='"&Trim(request("BillNo"))&"' and RecordStateID=0"
				Set rs1=conn.execute(strSQL1)
				If rs1.eof Then
					chkBillBaseFlag=5
				End If
				rs1.close
				Set rs1=Nothing 
			End If 
		End If 

		'chkBillBaseFlag=1,0 已建檔
		'chkBillBaseFlag=2 未建檔
		'chkBillBaseFlag=3,4 刪除失敗或未回傳
		'chkBillBaseFlag=5 台中市檢查是否有在登記簿裡

%>setCheckBillNoExist("<%=chkGetBillFlag%>","<%=chkBillBaseFlag%>","<%=chkBillSN%>","<%=chkBillTypeID%>","<%=MLoginID%>","<%=MMemberID%>","<%=MMemName%>","<%=MUnitID%>","<%=MUnitName%>","<%=RecMem2%>","<%=IllDate2%>","<%=IllRuleB1%>","<%=IllRuleB2%>","<%=chkGetBillMemUnit%>","<%=MStateID%>");
<%		
	conn.close
	set conn=Nothing

%>
//檢查單號是否有在GETBILLBASE內
function setCheckBillNoExist(GetBillFlag,BillBaseFlag,BillSN,BillType,MLoginID,MMemberID,MMemName,MUnitID,MUnitName,RecMem2,IllDate2,IllRuleB1,IllRuleB2,chkGetBillMemUnit,MStateID)
{
	if (GetBillFlag==0){
	<%if sys_City="嘉義縣" or sys_City="嘉義市" or sys_City="宜蘭縣" or sys_City="台東縣" or sys_City="台南縣" or sys_City="彰化縣" then%>
		alert("此單號不存在於領單紀錄中！");
	<%elseIf sys_City="台南市" or sys_City="基隆市" or (sys_City="台中市" And TCChkGetBillFlag=1) then%>
		alert("此單號不存在於領單紀錄中，如無領單記錄不可建檔！");
		document.myForm.Billno1.select();
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
		<%if sys_City="台中市" then%>
			document.myForm.IllegalZip.value="<%=MZipCode%>";
		<%end if%>
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
<%if sys_City="台中市" then%>
	}else if (BillBaseFlag==5){
		alert("此單號，登記簿沒有登打資料！！");
<%end if%>
	}else if (chkGetBillMemUnit==1){
	<%if sys_City<>"台南市" and sys_City<>"台中市" then%>
		alert("建檔單位非領單單位！！");
		<%if sys_City="嘉義縣" or sys_City="高雄縣" or sys_City="基隆市" then%>
			document.myForm.Billno1.select();
		<%end if%>
	<%end if%>
	}else if (MStateID==-1){
		alert("此領單人員帳號已停用，請至 ' 領單管理系統 ' 確認領單人及單位是否正確!");
	}
}