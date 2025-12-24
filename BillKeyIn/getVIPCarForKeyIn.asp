<!--#include virtual="/traffic/Common/db.ini"-->
<%
' 檔案名稱： getVIPCarForKeyIn.asp
	'是否為特殊用車

		'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing
	
	SpecNote=""
	if (trim(Session("SpecUser"))="1" or sys_City="台南市") and trim(request("BillType"))="2" then
		strVIP="select * from SpecCar where CarNo='"&trim(request("CarID"))&"' and RecordStateID<>-1"
		set rsVIP=conn.execute(strVIP)
		if not rsVIP.eof then
			CarCnt=1
			SpecNote=trim(rsVIP("Note"))
		else
			CarCnt=0
		end if
		rsVIP.close
		set rsVIP=nothing

	else
		CarCnt=0
	end if
	'檢查車號在幾天內有重複
	RepeatCnt=0
	RecUnit1=""
	IllDate1=""
	IllTime1=""
	IllRuleA1=""
	IllRuleA2=""
	BillFillMemName=""
	BillFillMemUnit=""
	strRep="select * from BillBase where CarNo='"&trim(request("CarID"))&"' and RecordStateID<>-1"
	set rsRep=conn.execute(strRep)
	If Not rsRep.Bof Then rsRep.MoveFirst 
	While Not rsRep.Eof
		if DateDiff("d",rsRep("RecordDate"),now) < CarNoRepeatDate then
			RepeatCnt=RepeatCnt+1
			strRecUnit="select a.UnitName from UnitInfo a,MemberData b where a.UnitID=b.UnitID and b.MemberID="&trim(rsRep("RecordMemberID"))
			set rsRecUnit=conn.execute(strRecUnit)
			if not rsRecUnit.eof then
				RecUnit1=trim(rsRecUnit("UnitName"))
			else
				RecUnit1=""
			end if
			rsRecUnit.close
			set rsRecUnit=nothing
			IllDate1=year(trim(rsRep("IllegalDate")))-1911&"/"&month(trim(rsRep("IllegalDate")))&"/"&day(trim(rsRep("IllegalDate")))&" "&hour(trim(rsRep("IllegalDate")))&":"&minute(trim(rsRep("IllegalDate")))
			IllRuleA1=trim(rsRep("Rule1"))
			IllRuleA2=trim(rsRep("Rule2"))
			IllStreet=trim(rsRep("IllegalAddress"))
			BillFillMemName=trim(rsRep("BillMem1"))
			if trim(rsRep("BillMem2"))<>"" and not isnull(rsRep("BillMem2")) then
				BillFillMemName=BillFillMemName&"，"&trim(rsRep("BillMem2"))
			end if
			if trim(rsRep("BillMem3"))<>"" and not isnull(rsRep("BillMem3")) then
				BillFillMemName=BillFillMemName&"，"&trim(rsRep("BillMem3"))
			end if
			strBillUnit="select UnitName from UnitInfo where UnitID='"&trim(rsRep("BillUnitID"))&"'"
			set rsBillUnit=conn.execute(strBillUnit)
			if not rsBillUnit.eof then
				BillFillMemUnit=trim(rsBillUnit("UnitName"))
			else
				BillFillMemUnit=""
			end if
			rsBillUnit.close
			set rsBillUnit=nothing
		end if
	rsRep.MoveNext
	Wend
	rsRep.close
	set rsRep=nothing
%>setVIPCar("<%=CarCnt%>","<%=RepeatCnt%>","<%=CarNoRepeatDate%>","<%=RecUnit1%>","<%=IllDate1%>","<%=IllRuleA1%>","<%=IllRuleA2%>","<%=IllStreet%>","<%=BillFillMemName%>","<%=BillFillMemUnit%>");
<%
conn.close
set conn=nothing
%>
//是否為特殊用車
function setVIPCar(CarCnt,RepeatCnt,CarNoRepeatDate,RecUnit1,IllDate1,IllRuleA1,IllRuleA2,IllStreet,BillFillMemName,BillFillMemUnit)
{
	if (CarCnt > 0){
		Layer7.innerHTML="＊業管車輛";
		TDVipCarErrorLog=1;
<%if sys_City="雲林縣" then %>
		alert("此車牌為業管車輛!");
		document.myForm.CarNo.select();
<%elseif sys_City="高雄市" Or sys_City=ApconfigureCityName then %>
		alert("此車牌為業管車輛。原因 ：<%=SpecNote%>");
		//document.myForm.CarNo.select();
<%end if%>
	}else{
		Layer7.innerHTML=" ";
		TDVipCarErrorLog=0;
	}
<%if sys_City<>"雲林縣" then%>
	if (RepeatCnt>0){
		alert("此車號已在"+CarNoRepeatDate+"天內建檔\n建檔單位："+RecUnit1+"\n違規時間："+IllDate1+"\n違規地點："+IllStreet+"\n違規法條："+IllRuleA1+" "+IllRuleA2+"\n舉發員警："+BillFillMemName+"\n舉發單位："+BillFillMemUnit+"\n請確認是否有重複建檔!")
		//document.myForm.CarSimpleID.focus();
	}
<%end if%>
}
