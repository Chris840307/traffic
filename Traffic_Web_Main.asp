<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing
	

%>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
body {
  font-family: "Microsoft JhengHei", Arial, sans-serif!important;
}
#Layer1 {
	position:absolute;
	width:209px;
	height:38px;
	z-index:2;
	top: 13px;
}
.style1 {font-size: 14px}
.style2 {font-size: 13px; }
#Layer2 {
	position:absolute;
	width:566px;
	height:38px;
	z-index:3;
	top: 15px;
}
#LayerTime {
	z-index:2;
}
#Layer151 {
	position:absolute;
	width:209px;
	height:38px;
	z-index:2;
}
#D1 {
	BACKGROUND-COLOR: #F4F6F8; 
	BORDER-BOTTOM: white 2px outset; 
	BORDER-LEFT: white 2px outset; 
	BORDER-RIGHT: white 2px outset; 
	BORDER-TOP: white 2px outset; 
	LEFT: 0px; POSITION: absolute; 
	TOP: 0px; VISIBILITY: hidden; 
	WIDTH: 150px; 
	layer-background-color: #F4F6F8;
	
	z-index:5;
}
-->
</style>
<head>
<!--#include virtual="traffic/Common/css.txt"-->
<title>宏謙科技實業有限公司-入案管理系統</title>
<%
memName=Session("Ch_Name")
GroupID=Session("Group_ID")
UnitNo=Session("Unit_ID")

 

	if sys_City<>"台中縣" then
		ArgueDate1=DateAdd("d",-10,date) & " 0:0:0"
		ArgueDate2=date & " 23:29:59" 
		strDelErr="select * from Dcilog where ExchangeTypeID='E' and (DciReturnStatusID<>'S')" &_
			" and ExchangeDate between TO_DATE('"&ArgueDate1&"','YYYY/MM/DD/HH24/MI/SS')" &_
			" and TO_DATE('"&ArgueDate2&"','YYYY/MM/DD/HH24/MI/SS') and RecordMemberID="&Session("User_ID")
		set rsDelErr=conn.execute(strDelErr)
		If Not rsDelErr.Bof Then rsDelErr.MoveFirst 
		While Not rsDelErr.Eof
			if (trim(rsDelErr("DciReturnStatusID"))<>"S" and not isnull(rsDelErr("DciReturnStatusID"))) Then
				ISDel=1
				strCheckErr="select * from (select * from Dcilog where ExchangeTypeID='E' and billsn="&trim(rsDelErr("BillSn"))&" order by ExchangeDate Desc) where Rownum<=1 "
				Set rsCE=conn.execute(strCheckErr)
				If Not rsCE.eof Then 
					If trim(rsCE("DciReturnStatusID"))="S" or isnull(rsCE("DciReturnStatusID")) Then 
						ISDel=0
					End If 
				End If
				rsCE.close
				Set rsCE=Nothing 
				If ISDel=1 Then 
					if trim(rsDelErr("DciReturnStatusID"))="n" then
						strUpd="Update BillBase set RecordStateID=0,BillStatus='2' where RecordStateID<>0 and Sn="&trim(rsDelErr("BillSn"))
						conn.execute strUpd
					else
						strUpd="Update BillBase set RecordStateID=0,BillStatus='2' where RecordStateID<>0 and Sn="&trim(rsDelErr("BillSn"))
						conn.execute strUpd
					end If
				End If 
			end if
		rsDelErr.MoveNext
		Wend
		rsDelErr.close
		set rsDelErr=nothing
	end if
	

'1100715(監理站不知道哪時加的)=====================================================================
	strChkL2="select * from Law where itemid ='1300305' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="insert into law values('1300305','0','引擎號碼與原登記位置不符(二次以上本行為)',3600,3900,4300,4800,'V','0','0','0','0','0',to_date('2021/7/15','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('1300306','0','引擎號碼與原登記模型不符(二次以上本行為)',3600,3900,4300,4800,'V','0','0','0','0','0',to_date('2021/7/15','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('1300307','0','車身號碼與原登記位置不符(二次以上本行為)',3600,3900,4300,4800,'V','0','0','0','0','0',to_date('2021/7/15','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('1300308','0','車身號碼與原登記模型不符(二次以上本行為)',3600,3900,4300,4800,'V','0','0','0','0','0',to_date('2021/7/15','yyyy/mm/dd'),'2',0,sysdate,null,null)"
		conn.execute strInsL2

	End if
	rsChkL2.close
	Set rsChkL2=Nothing
'1101001=====================================================================
	strChkL2="select * from Law where itemid ='2910113' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="insert into law values('2910113','0','裝載貨物超過規定長度肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910114','0','裝載貨物超過規定寬度肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('2910115','0','裝載貨物超過規定高度肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910116','0','裝載貨物超過規定長度肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('2910117','0','裝載貨物超過規定寬度肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910118','0','裝載貨物超過規定高度肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910119','0','裝載貨物超過規定長度肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910120','0','裝載貨物超過規定寬度肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910121','0','裝載貨物超過規定高度肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910122','5','裝載貨物超過規定長度應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910122','6','裝載貨物超過規定長度應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('2910123','5','裝載貨物超過規定寬度應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910123','6','裝載貨物超過規定寬度應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('2910124','5','裝載貨物超過規定高度應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910124','6','裝載貨物超過規定高度應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('2910125','0','裝載貨物超過規定長度肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910126','0','裝載貨物超過規定寬度肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910127','0','裝載貨物超過規定高度肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910128','0','裝載貨物超過規定長度肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910129','0','裝載貨物超過規定寬度肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910130','0','裝載貨物超過規定高度肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910131','0','裝載貨物超過規定長度肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910132','0','裝載貨物超過規定寬度肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910133','0','裝載貨物超過規定高度肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('2910233','0','裝載整體物品有超重未請領臨時通行證肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910234','0','裝載整體物品有超長未請領臨時通行證肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910235','0','裝載整體物品有超寬未請領臨時通行證肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910236','0','裝載整體物品有超高未請領臨時通行證肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910237','0','裝載整體物品有超重未懸掛危險標識肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910238','0','裝載整體物品有超長未懸掛危險標識肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910239','0','裝載整體物品有超寬未懸掛危險標識肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910240','0','裝載整體物品有超高未懸掛危險標識肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910241','0','裝載整體物品有超重未請領臨時通行證肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910242','0','裝載整體物品有超長未請領臨時通行證肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910243','0','裝載整體物品有超寬未請領臨時通行證肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910244','0','裝載整體物品有超高未請領臨時通行證肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910245','0','裝載整體物品有超重未懸掛危險標識肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910246','0','裝載整體物品有超長未懸掛危險標識肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910247','0','裝載整體物品有超寬未懸掛危險標識肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910248','0','裝載整體物品有超高未懸掛危險標識肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('2910249','0','裝載整體物品有超重未請領臨時通行證肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910250','0','裝載整體物品有超長未請領臨時通行證肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910251','0','裝載整體物品有超寬未請領臨時通行證肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910252','0','裝載整體物品有超高未請領臨時通行證肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910253','0','裝載整體物品有超重未懸掛危險標識肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910254','0','裝載整體物品有超長未懸掛危險標識肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910255','0','裝載整體物品有超寬未懸掛危險標識肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910256','0','裝載整體物品有超高未懸掛危險標識肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910257','5','裝載整體物品有超重未請領臨時通行證應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910257','6','裝載整體物品有超重未請領臨時通行證應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910258','5','裝載整體物品有超長未請領臨時通行證應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910258','6','裝載整體物品有超長未請領臨時通行證應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910259','5','裝載整體物品有超寬未請領臨時通行證應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910259','6','裝載整體物品有超寬未請領臨時通行證應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910260','5','裝載整體物品有超高未請領臨時通行證應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910260','6','裝載整體物品有超高未請領臨時通行證應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910261','5','裝載整體物品有超重未懸掛危險標識應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910261','6','裝載整體物品有超重未懸掛危險標識應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910262','5','裝載整體物品有超長未懸掛危險標識應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910262','6','裝載整體物品有超長未懸掛危險標識應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910263','5','裝載整體物品有超寬未懸掛危險標識應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910263','6','裝載整體物品有超寬未懸掛危險標識應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910264','5','裝載整體物品有超高未懸掛危險標識應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910264','6','裝載整體物品有超高未懸掛危險標識應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910265','0','裝載整體物品有超重未請領臨時通行證肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910266','0','裝載整體物品有超長未請領臨時通行證肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910267','0','裝載整體物品有超寬未請領臨時通行證肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910268','0','裝載整體物品有超高未請領臨時通行證肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910269','0','裝載整體物品有超重未懸掛危險標識肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910270','0','裝載整體物品有超長未懸掛危險標識肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910271','0','裝載整體物品有超寬未懸掛危險標識肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910272','0','裝載整體物品有超高未懸掛危險標識肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910273','0','裝載整體物品有超重未請領臨時通行證肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910274','0','裝載整體物品有超長未請領臨時通行證肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910275','0','裝載整體物品有超寬未請領臨時通行證肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910276','0','裝載整體物品有超高未請領臨時通行證肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910277','0','裝載整體物品有超重未懸掛危險標識肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910278','0','裝載整體物品有超長未懸掛危險標識肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910279','0','裝載整體物品有超寬未懸掛危險標識肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910280','0','裝載整體物品有超高未懸掛危險標識肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910281','0','裝載整體物品有超重未請領臨時通行證肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910282','0','裝載整體物品有超長未請領臨時通行證肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910283','0','裝載整體物品有超寬未請領臨時通行證肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910284','0','裝載整體物品有超高未請領臨時通行證肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910285','0','裝載整體物品有超重未懸掛危險標識肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910286','0','裝載整體物品有超長未懸掛危險標識肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910287','0','裝載整體物品有超寬未懸掛危險標識肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910288','0','裝載整體物品有超高未懸掛危險標識肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910327','0','裝載危險物品未請領臨時通行證肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910328','0','裝載危險物品未依規定懸掛或黏貼危險物品標誌及標示牌肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910329','0','裝載危險物品罐槽車之罐槽體未檢驗合格肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910330','0','裝載危險物品運送人員未經專業訓練合格肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910331','0','裝載危險物品運送人員不遵守有關安全規定肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910332','0','裝載危險物品未請領臨時通行證肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910333','0','裝載危險物品未依規定懸掛或黏貼危險物品標誌及標示牌肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910334','0','裝載危險物品罐槽車之罐槽體未檢驗合格肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910335','0','裝載危險物品運送人員未經專業訓練合格肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910336','0','裝載危險物品運送人員不遵守有關安全規定肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910337','0','裝載危險物品未請領臨時通行證肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910338','0','裝載危險物品未依規定懸掛或黏貼危險物品標誌及標示牌肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910339','0','裝載危險物品罐槽車之罐槽體未檢驗合格肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910340','0','裝載危險物品運送人員未經專業訓練合格肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910341','0','裝載危險物品運送人員不遵守有關安全規定肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910342','0','裝載危險物品未請領臨時通行證應歸責於汽車駕駛人',9000,9000,9000,9000,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910343','5','裝載危險物品未依規定懸掛或黏貼危險物品標誌及標示牌應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910343','6','裝載危險物品未依規定懸掛或黏貼危險物品標誌及標示牌應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910344','0','裝載危險物品罐槽車之罐槽體未檢驗合格應歸責於汽車駕駛人',9000,9000,9000,9000,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910345','0','裝載危險物品運送人員未經專業訓練合格應歸責於汽車駕駛人',9000,9000,9000,9000,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910346','5','裝載危險物品運送人員不遵守有關安全規定應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910346','6','裝載危險物品運送人員不遵守有關安全規定應歸責於汽車駕駛人',4500,4900,5800,6700,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910347','0','裝載危險物品未請領臨時通行證肇事致人受傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910348','0','裝載危險物品未依規定懸掛或黏貼危險物品標誌及標示牌肇事致人受傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910349','0','裝載危險物品罐槽車之罐槽體未檢驗合格肇事致人受傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910350','0','裝載危險物品運送人員未經專業訓練合格肇事致人受傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910351','0','裝載危險物品運送人員不遵守有關安全規定肇事致人受傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910352','0','裝載危險物品未請領臨時通行證肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910353','0','裝載危險物品未依規定懸掛或黏貼危險物品標誌及標示牌肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910354','0','裝載危險物品罐槽車之罐槽體未檢驗合格肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910355','0','裝載危險物品運送人員未經專業訓練合格肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910356','0','裝載危險物品運送人員不遵守有關安全規定肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910357','0','裝載危險物品未請領臨時通行證肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910358','0','裝載危險物品未依規定懸掛或黏貼危險物品標誌及標示牌肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910359','0','裝載危險物品罐槽車之罐槽體未檢驗合格肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910360','0','裝載危險物品運送人員未經專業訓練合格肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910361','0','裝載危險物品運送人員不遵守有關安全規定肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910403','0','貨車裝載不依規定者肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910404','0','聯結汽車之裝載不依規定者肇事致人受傷',10000,11000,13000,15000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910405','0','貨車裝載不依規定者肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910406','0','聯結汽車之裝載不依規定者肇事致人重傷',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910407','0','貨車裝載不依規定者肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910408','0','聯結汽車之裝載不依規定者肇事致人死亡',18000,18000,18000,18000,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910409','0','貨車裝載不依規定者應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910410','0','聯結汽車之裝載不依規定者應歸責於汽車駕駛人',3000,3300,3900,4500,'0','2','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910411','0','貨車裝載不依規定者肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910412','0','聯結汽車之裝載不依規定者肇事致人受傷應歸責於汽車駕駛人',10000,11000,13000,15000,'0','0','L','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910413','0','貨車裝載不依規定者肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910414','0','聯結汽車之裝載不依規定者肇事致人重傷應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910415','0','貨車裝載不依規定者肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910416','0','聯結汽車之裝載不依規定者肇事致人死亡應歸責於汽車駕駛人',18000,18000,18000,18000,'0','0','2','3','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910503','0','汽車牽引拖架不依規定肇事致人受傷',10000,11000,13000,15000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910504','0','汽車附掛拖車不依規定肇事致人受傷',10000,11000,13000,15000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910505','0','汽車牽引拖架不依規定肇事致人重傷',18000,18000,18000,18000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910506','0','汽車附掛拖車不依規定肇事致人重傷',18000,18000,18000,18000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910507','0','汽車牽引拖架不依規定肇事致人死亡',18000,18000,18000,18000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910508','0','汽車附掛拖車不依規定肇事致人死亡',18000,18000,18000,18000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910603','0','大貨車裝載貨櫃超出車身之外肇事致人受傷',10000,11000,13000,15000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910604','0','大貨車裝載貨櫃未依規定裝置聯鎖設備肇事致人受傷',10000,11000,13000,15000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910605','0','大貨車裝載貨櫃超出車身之外肇事致人重傷',18000,18000,18000,18000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910606','0','大貨車裝載貨櫃未依規定裝置聯鎖設備肇事致人重傷',18000,18000,18000,18000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910607','0','大貨車裝載貨櫃超出車身之外肇事致人死亡',18000,18000,18000,18000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910608','0','大貨車裝載貨櫃未依規定裝置聯鎖設備肇事致人死亡',18000,18000,18000,18000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2910702','0','汽車未經核准附掛拖車行駛肇事致人受傷',10000,11000,13000,15000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910703','0','汽車未經核准附掛拖車行駛肇事致人重傷',18000,18000,18000,18000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('2910704','0','汽車未經核准附掛拖車行駛肇事致人死亡',18000,18000,18000,18000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('2930002','0','違反第29條第1項第1款至第4款情形，應歸責於汽車駕駛人，汽車所有人仍應記違規紀錄1次',0,0,0,0,'V','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="update Law set IllegalRule='非屬汽車範圍而行駛於道路上之動力機械未依規定請領臨時通行證(無車號)',Level1=3000,Level2=3300,Level3=3900,Level4=4500 where itemid='3210001' and version=2"
		conn.execute strInsL2

		strInsL2="insert into law values('3210003','0','非屬汽車範圍而行駛於道路上之動力機械未依規定請領臨時通行證(不宜行駛公路)',3000,3300,3900,4500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('3210004','0','非屬汽車範圍而行駛於道路上之動力機械未依規定請領臨時通行證(有車號)',3000,3300,3900,4500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

	End if
	rsChkL2.close
	Set rsChkL2=Nothing
'獎勵金記得要加法條點數==================================================================
	strInsL2="update Law set Recordstateid=0 where itemid='8210201' and version=2"
	conn.execute strInsL2
'1110401=====================================================================
	'If asfd="dsafs" then
	strChkL2="select * from Law where itemid ='35300169' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="insert into law values('35300169','3','汽機車駕駛人駕駛汽機車，於十年內酒精濃度超過規定標準第2次',90000,90000,90000,90000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35300169','4','汽機車駕駛人駕駛汽機車，於十年內酒精濃度超過規定標準第2次',120000,120000,120000,120000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35300170','3','汽機車駕駛人駕駛汽機車肇事致人重傷，且於十年內酒精濃度超過規定標準第2次',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300170','4','汽機車駕駛人駕駛汽機車肇事致人重傷，且於十年內酒精濃度超過規定標準第2次',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35300171','3','汽機車駕駛人駕駛汽機車肇事致人死亡，且於十年內酒精濃度超過規定標準第2次',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300171','4','汽機車駕駛人駕駛汽機車肇事致人死亡，且於十年內酒精濃度超過規定標準第2次',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35300172','3','汽機車駕駛人駕駛汽機車，於十年內酒精濃度超過規定標準第2次(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300172','4','汽機車駕駛人駕駛汽機車，於十年內酒精濃度超過規定標準第2次(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35300173','3','汽機車駕駛人駕駛汽機車肇事致人重傷，且於十年內酒精濃度超過規定標準第2次(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300173','4','汽機車駕駛人駕駛汽機車肇事致人重傷，且於十年內酒精濃度超過規定標準第2次(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35300174','3','汽機車駕駛人駕駛汽機車肇事致人死亡，且於十年內酒精濃度超過規定標準第2次(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300174','4','汽機車駕駛人駕駛汽機車肇事致人死亡，且於十年內酒精濃度超過規定標準第2次(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300175','3','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食毒品第2次',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35300175','4','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食毒品第2次',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300176','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食毒品第2次',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300176','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食毒品第2次',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2		

		strInsL2="insert into law values('35300177','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食毒品第2次',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300177','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食毒品第2次',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300178','3','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食毒品第2次(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300178','4','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食毒品第2次(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300179','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食毒品第2次(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300179','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食毒品第2次(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35300180','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食毒品第2次(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300180','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食毒品第2次(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300181','3','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食迷幻藥第2次',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300181','4','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食迷幻藥第2次',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300182','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食迷幻藥第2次',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300182','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食迷幻藥第2次',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35300183','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食迷幻藥第2次',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300183','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食迷幻藥第2次',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
				
		strInsL2="insert into law values('35300184','3','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食迷幻藥第2次(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300184','4','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食迷幻藥第2次(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
				
		strInsL2="insert into law values('35300185','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食迷幻藥第2次(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300185','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食迷幻藥第2次(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
				
		strInsL2="insert into law values('35300186','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食迷幻藥第2次(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300186','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食迷幻藥第2次(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
			
		strInsL2="insert into law values('35300187','3','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食麻醉藥品第2次',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300187','4','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食麻醉藥品第2次',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
				
		strInsL2="insert into law values('35300188','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食麻醉藥品第2次',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300188','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食麻醉藥品第2次',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
				
		strInsL2="insert into law values('35300189','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食麻醉藥品第2次',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300189','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食麻醉藥品第2次',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
			
		strInsL2="insert into law values('35300190','3','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食麻醉藥品第2次(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300190','4','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食麻醉藥品第2次(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
				
		strInsL2="insert into law values('35300191','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食麻醉藥品第2次(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300191','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食麻醉藥品第2次(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
				
		strInsL2="insert into law values('35300192','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食麻醉藥品第2次(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300192','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食麻醉藥品第2次(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300193','3','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食管制藥品第2次',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300193','4','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食管制藥品第2次',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300194','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食管制藥第2次',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300194','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食管制藥第2次',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300195','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食管制藥第2次',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300195','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食管制藥第2次',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300196','3','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食管制藥品第2次(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300196','4','汽機車駕駛人駕駛汽機車，於十年內經測試檢定有吸食管制藥品第2次(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300197','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食管制藥第2次(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300197','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內經測試檢定有吸食管制藥第2次(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300198','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食管制藥第2次(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300198','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內經測試檢定有吸食管制藥第2次(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300199','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食毒品)',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300199','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食毒品)',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300200','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300200','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300201','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300201','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300202','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300202','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300203','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食毒品)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300203','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食毒品)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300204','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300204','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300205','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300205','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35300206','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300206','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)',120000,120000,120000,120000,'0','0','9','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35300207','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食毒品)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300207','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食毒品)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300208','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300208','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300209','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300209','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300210','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300210','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300211','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食毒品)(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300211','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食毒品)(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300212','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300212','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300213','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300213','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300214','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300214','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300215','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食毒品)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300215','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食毒品)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300216','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300216','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300217','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300217','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300218','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300218','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300219','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食毒品)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300219','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食毒品)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300220','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300220','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食迷幻藥)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300221','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300221','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食麻醉藥品)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300222','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300222','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先酒駕、後吸食管制藥品)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300223','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食毒品、後酒駕)',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300223','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食毒品、後酒駕)',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300224','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300224','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300225','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300225','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300226','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300226','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)',120000,120000,120000,120000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300227','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食毒品、後酒駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300227','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食毒品、後酒駕)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300228','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300228','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300229','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300229','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300230','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300230','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300231','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食毒品、後酒駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300231','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食毒品、後酒駕)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300232','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300232','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300233','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300233','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300234','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300234','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)',120000,120000,120000,120000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300235','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食毒品、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300235','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食毒品、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300236','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300236','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300237','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300237','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300238','3','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300238','4','汽機車駕駛人駕駛汽機車，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300239','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食毒品、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300239','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食毒品、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300240','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300240','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300241','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300241','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300242','3','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300242','4','汽機車駕駛人駕駛汽機車肇事致人重傷，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300243','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食毒品、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300243','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食毒品、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300244','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300244','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食迷幻藥、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300245','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300245','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食麻醉藥品、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300246','3','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)(無照)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300246','4','汽機車駕駛人駕駛汽機車肇事致人死亡，於十年內違反第一項第2次(先吸食管制藥品、後酒駕)(無照)',120000,120000,120000,120000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300247','0','汽機車駕駛人駕駛汽機車，於十年內違反第一項規定第3次以上(酒駕)',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300248','0','汽機車駕駛人駕駛汽機車致人重傷，於十年內違反第一項規定第3次以上(酒駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300249','0','汽機車駕駛人駕駛汽機車致人死亡，於十年內違反第一項規定第3次以上(酒駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300250','0','汽機車駕駛人駕駛汽機車，於十年內違反第一項規定(無照)第3次以上(酒駕)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300251','0','汽機車駕駛人駕駛汽機車致人重傷，於十年內違反第一項規定(無照)第3次以上(酒駕)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300252','0','汽機車駕駛人駕駛汽機車致人死亡，於十年內違反第一項規定(無照)第3次以上(酒駕)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300253','0','汽機車駕駛人駕駛汽機車，於十年內違反第一項規定第3次以上(藥駕)',90000,90000,90000,90000,'0','0','2','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300254','0','汽機車駕駛人駕駛汽機車致人重傷，於十年內違反第一項規定第3次以上(藥駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300255','0','汽機車駕駛人駕駛汽機車致人死亡，於十年內違反第一項規定第3次以上(藥駕)',90000,90000,90000,90000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300256','0','汽機車駕駛人駕駛汽機車，於十年內違反第一項規定(無照)第3次以上(藥駕)',90000,90000,90000,90000,'0','0','0','3','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300257','0','汽機車駕駛人駕駛汽機車致人重傷，於十年內違反第一項規定(無照)第3次以上(藥駕)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35300258','0','汽機車駕駛人駕駛汽機車致人死亡，於十年內違反第一項規定(無照)第3次以上(藥駕)',90000,90000,90000,90000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403001','0','接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品',180000,180000,180000,180000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403002','0','接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品',180000,180000,180000,180000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403003','0','接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品',180000,180000,180000,180000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403004','0','接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品',180000,180000,180000,180000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403005','0','接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品',180000,180000,180000,180000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403006','0','接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403007','0','接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403008','0','接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403009','0','接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403010','0','接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403011','0','接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403012','0','接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403013','0','接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403014','0','接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403015','0','接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403016','0','接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品(無駕駛執照)',180000,180000,180000,180000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403017','0','接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品(無駕駛執照)',180000,180000,180000,180000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403018','0','接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品(無駕駛執照)',180000,180000,180000,180000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403019','0','接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品(無駕駛執照)',180000,180000,180000,180000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403020','0','接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品(無駕駛執照)',180000,180000,180000,180000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403021','0','接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403022','0','接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403023','0','接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403024','0','接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403025','0','接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403026','0','接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403027','0','接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403028','0','接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403029','0','接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35403030','0','接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404001','0','發生交通事故後，在接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品',180000,180000,180000,180000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404002','0','發生交通事故後，在接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品',180000,180000,180000,180000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404003','0','發生交通事故後，在接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品',180000,180000,180000,180000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404004','0','發生交通事故後，在接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品',180000,180000,180000,180000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404005','0','發生交通事故後，在接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品',180000,180000,180000,180000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404006','0','發生交通事故後，在接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404007','0','發生交通事故後，在接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404008','0','發生交通事故後，在接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404009','0','發生交通事故後，在接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404010','0','發生交通事故後，在接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404011','0','發生交通事故後，在接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404012','0','發生交通事故後，在接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404013','0','發生交通事故後，在接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404014','0','發生交通事故後，在接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404015','0','發生交通事故後，在接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404016','0','發生交通事故後，在接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品(無駕駛執照)',180000,180000,180000,180000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404017','0','發生交通事故後，在接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品(無駕駛執照)',180000,180000,180000,180000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404018','0','發生交通事故後，在接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品(無駕駛執照)',180000,180000,180000,180000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404019','0','發生交通事故後，在接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品(無駕駛執照)',180000,180000,180000,180000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404020','0','發生交通事故後，在接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品(無駕駛執照)',180000,180000,180000,180000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404021','0','發生交通事故後，在接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404022','0','發生交通事故後，在接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404023','0','發生交通事故後，在接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404024','0','發生交通事故後，在接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404025','0','發生交通事故後，在接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人重傷(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404026','0','發生交通事故後，在接受酒精濃度測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404027','0','發生交通事故後，在接受毒品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404028','0','發生交通事故後，在接受迷幻藥測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404029','0','發生交通事故後，在接受麻醉藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35404030','0','發生交通事故後，在接受管制藥品測試檢定前，吸食服用含酒精之物、毒品、迷幻藥、麻醉藥品及其相類似之管制藥品，且肇事致人死亡(無駕駛執照)',180000,180000,180000,180000,'0','0','0','9','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500013','0','汽機車駕駛人駕駛汽機車，於十年內第2次違反第4項規定',360000,360000,360000,360000,'0','0','2','5','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500014','0','汽機車駕駛人駕駛汽機車，於十年內第2次違反第4項規定，肇事致人重傷',360000,360000,360000,360000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500015','0','汽機車駕駛人駕駛汽機車，於十年內第2次違反第4項規定，肇事致人死亡',360000,360000,360000,360000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500016','0','汽機車駕駛人駕駛汽機車，於十年內第2次違反第4項規定(無照)',360000,360000,360000,360000,'0','0','0','5','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500017','0','汽機車駕駛人駕駛汽機車，於十年內第2次違反第4項規定，肇事致人重傷(無照)',360000,360000,360000,360000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500018','0','汽機車駕駛人駕駛汽機車，於十年內第2次違反第4項規定，肇事致人死亡(無照)',360000,360000,360000,360000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500019','0','汽機車駕駛人駕駛汽機車，於十年內第3次以上違反第4項規定',180000,180000,180000,180000,'0','0','2','5','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500020','0','汽機車駕駛人駕駛汽機車，於十年內第3次以上違反第4項規定，肇事致人重傷',180000,180000,180000,180000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500021','0','汽機車駕駛人駕駛汽機車，於十年內第3次以上違反第4項規定，肇事致人死亡',180000,180000,180000,180000,'0','0','2','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500022','0','汽機車駕駛人駕駛汽機車，於十年內第3次以上違反第4項規定(無照)',180000,180000,180000,180000,'0','0','0','5','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500023','0','汽機車駕駛人駕駛汽機車，於十年內第3次以上違反第4項規定，肇事致人重傷(無照)',180000,180000,180000,180000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35500024','0','汽機車駕駛人駕駛汽機車，於十年內第3次以上違反第4項規定，肇事致人死亡(無照)',180000,180000,180000,180000,'0','0','0','9','0','6,g',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700009','3','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.15-0.25(未含))而不禁駛',15000,16500,19500,22500,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700009','5','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.15-0.25(未含))而不禁駛',30000,33000,39000,45000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700009','6','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.15-0.25(未含))而不禁駛',33000,36000,42500,49000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700010','3','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.25-0.4(未含))而不禁駛',22500,24500,29000,33500,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700010','5','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.25-0.4(未含))而不禁駛',37500,40500,47500,54500,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700010','6','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.25-0.4(未含))而不禁駛',42000,46000,54000,62000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700011','3','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.4-0.55(未含))而不禁駛',45000,49500,58500,67500,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700011','5','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.4-0.55(未含))而不禁駛',60000,66000,78000,90000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700011','6','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.4-0.55(未含))而不禁駛',65000,71500,84000,97000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700012','3','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.55以上)而不禁駛',67500,74000,87500,90000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700012','5','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.55以上)而不禁駛',85000,93500,110000,120000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700012','6','汽機車所有人明知駕駛人酒精濃度超過規定標準(0.55以上)而不禁駛',100000,110000,120000,120000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700013','3','汽機車所有人明知駕駛人吸食毒品駕駛汽車而不予禁止駕駛',90000,90000,90000,90000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700013','4','汽機車所有人明知駕駛人吸食毒品駕駛汽車而不予禁止駕駛',120000,120000,120000,120000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700014','3','汽機車所有人明知駕駛人吸食迷幻藥駕駛汽車而不予禁止駕駛',90000,90000,90000,90000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700014','4','汽機車所有人明知駕駛人吸食迷幻藥駕駛汽車而不予禁止駕駛',120000,120000,120000,120000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700015','3','汽機車所有人明知駕駛人吸食麻醉藥品駕駛汽車而不予禁止駕駛',90000,90000,90000,90000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700015','4','汽機車所有人明知駕駛人吸食麻醉藥品駕駛汽車而不予禁止駕駛',120000,120000,120000,120000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700016','3','汽機車所有人明知駕駛人吸食管制藥品駕駛汽車而不予禁止駕駛',90000,90000,90000,90000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35700016','4','汽機車所有人明知駕駛人吸食管制藥品駕駛汽車而不予禁止駕駛',120000,120000,120000,120000,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35800003','3','汽機車駕駛人酒精濃度超過規定標準(0.25-0.55(未含))滿18歲之同車乘客',6000,6600,7800,9000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35800003','5','汽機車駕駛人酒精濃度超過規定標準(0.25-0.55(未含))滿18歲之同車乘客',9000,9900,11000,12000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35800003','6','汽機車駕駛人酒精濃度超過規定標準(0.25-0.55(未含))滿18歲之同車乘客',10000,11000,13000,15000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35800004','3','汽機車駕駛人酒精濃度超過規定標準(0.55以上)滿18歲之同車乘客',9000,9900,11000,12000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35800004','5','汽機車駕駛人酒精濃度超過規定標準(0.55以上)滿18歲之同車乘客',10000,11000,13000,15000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35800004','6','汽機車駕駛人酒精濃度超過規定標準(0.55以上)滿18歲之同車乘客',15000,15000,15000,15000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900002','0','汽機車駕駛人有第35條第1項第1款之情形',0,0,0,0,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900003','0','汽機車駕駛人有第35條第1項第2款之情形',0,0,0,0,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900004','0','汽機車駕駛人有第35條第3項之情形',0,0,0,0,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900005','0','汽機車駕駛人有第35條第4項第1款之情形',0,0,0,0,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900006','0','汽機車駕駛人有第35條第4項第2款之情形',0,0,0,0,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900007','0','汽機車駕駛人有第35條第4項第3款之情形',0,0,0,0,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900008','0','汽機車駕駛人有第35條第4項第4款之情形',0,0,0,0,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900009','0','汽機車駕駛人有第35條第5項之情形',0,0,0,0,'V','0','X','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900010','0','汽機車駕駛人有第35條第1項第1款之情形因而肇事致人重傷',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900011','0','汽機車駕駛人有第35條第1項第2款之情形因而肇事致人重傷',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900012','0','汽機車駕駛人有第35條第3項之情形因而肇事致人重傷',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900013','0','汽機車駕駛人有第35條第4項第1款之情形因而肇事致人重傷',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900014','0','汽機車駕駛人有第35條第4項第2款之情形因而肇事致人重傷',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900015','0','汽機車駕駛人有第35條第4項第3款之情形因而肇事致人重傷',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900016','0','汽機車駕駛人有第35條第4項第4款之情形因而肇事致人重傷',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900017','0','汽機車駕駛人有第35條第5項之情形因而肇事致人重傷',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900018','0','汽機車駕駛人有第35條第1項第1款之情形因而肇事致人死亡',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900019','0','汽機車駕駛人有第35條第1項第2款之情形因而肇事致人死亡',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900020','0','汽機車駕駛人有第35條第3項之情形因而肇事致人死亡',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900021','0','汽機車駕駛人有第35條第4項第1款之情形因而肇事致人死亡',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900022','0','汽機車駕駛人有第35條第4項第2款之情形因而肇事致人死亡',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900023','0','汽機車駕駛人有第35條第4項第3款之情形因而肇事致人死亡',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900024','0','汽機車駕駛人有第35條第4項第4款之情形因而肇事致人死亡',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		strInsL2="insert into law values('35900025','0','汽機車駕駛人有第35條第5項之情形因而肇事致人死亡',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('351000031','3','汽車駕駛人不依規定駕駛配備車輛點火自動鎖定裝置汽車',60000,66000,78000,90000,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('351000031','5','汽車駕駛人不依規定駕駛配備車輛點火自動鎖定裝置汽車',90000,99000,110000,120000,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('351000031','6','汽車駕駛人不依規定駕駛配備車輛點火自動鎖定裝置汽車',120000,120000,120000,120000,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('352000021','3','汽車駕駛人依第35條之1第1項規定申請登記而不依規定使用車輛點火自動鎖定裝置',10000,11000,13000,15000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('352000021','5','汽車駕駛人依第35條之1第1項規定申請登記而不依規定使用車輛點火自動鎖定裝置',20000,22000,26000,30000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('352000021','6','汽車駕駛人依第35條之1第1項規定申請登記而不依規定使用車輛點火自動鎖定裝置',30000,30000,30000,30000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('353000011','3','車輛點火自動鎖定裝置由他人代為使用解鎖',6000,6600,7800,9000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('353000011','5','車輛點火自動鎖定裝置由他人代為使用解鎖',9000,9900,11000,12000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('353000011','6','車輛點火自動鎖定裝置由他人代為使用解鎖',12000,12000,12000,12000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
	
	End if
	rsChkL2.close
	Set rsChkL2=Nothing

	strChkL2="select * from Law where itemid ='3510' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="insert into law values('3510','0','租賃車業者已盡告知第35條處罰規定之義務，汽機車駕駛人依其駕駛車輛所處罰鍰加罰二分之一',0,0,0,0,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

	End if
	rsChkL2.close
	Set rsChkL2=Nothing
'1110801=====================================================================
	strChkL2="select * from Law where itemid ='3110009' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="insert into law values('3110009','0','汽車行駛於一般道路上營業大客車駕駛人未依規定繫安全帶',2000,2000,2000,2000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('3120023','0','汽車行駛於高速公路未依規定繫安全帶(二人以上)—營業大客車、計程車或租賃車輛有代僱駕駛人',6000,6000,6000,6000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('3120024','0','汽車行駛於快速公路未依規定繫安全帶(二人以上)—營業大客車、計程車或租賃車輛有代僱駕駛人',6000,6000,6000,6000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('3120025','0','營業大客車行駛於高速公路上其四歲以上乘客經告知仍未繫安全帶(罰乘客)',3000,3300,3900,4500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('3120026','0','營業大客車行駛於快速公路上其四歲以上乘客經告知仍未繫安全帶(罰乘客)',3000,3300,3900,4500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('18100051','1','汽車未依規定裝設防止捲入裝置',12000,13000,15000,18000,'V','0','0','0','0','5',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('18100051','2','汽車未依規定裝設防止捲入裝置',14000,15000,18000,21000,'V','0','0','0','0','5',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('18100061','0','二種以上設備同時違反第18條之1第1項規定',16000,17000,19000,22000,'V','0','0','0','0','5',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('18200051','0','汽車裝設之防止捲入裝置無法正常運作，未於行車前改善，仍繼續行駛',9000,9900,11000,13000,'V','0','0','0','0','5',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('18200061','0','二種以上設備同時違反第18條之1第2項規定',12000,13000,15000,18000,'V','0','0','0','0','5',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('8519903','0','違規處罰，以主要駕駛人為被通知人',0,0,0,0,'+','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('8519904','0','違規處罰，以主要駕駛人為被通知人，不記點',0,0,0,0,'+','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('8519905','0','違規處罰，以長租車租用人為被通知人',0,0,0,0,'+','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
	End if
	rsChkL2.close
	Set rsChkL2=Nothing
'1111130微電車=====================================================================
	strChkL2="select * from Law where itemid ='32000101' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="insert into law values('6920003','0','人力行駛車輛，未依規定辦理登記，領取證照即行駛道路',300,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6920004','0','獸力行駛車輛，未依規定辦理登記，領取證照即行駛道路',300,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950001','0','個人行動器具未依直轄市、縣（市）政府所定規格、指定行駛路段、時間、速度限制、安全注意及其他管理事項規定行駛',1200,1600,1600,1600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950002','0','個人行動器具未依直轄市、縣（市）政府所定規格、指定行駛路段、時間、速度限制、安全注意及其他管理事項規定行駛，肇事致人受傷',2400,3200,3200,3200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950003','0','個人行動器具未依直轄市、縣（市）政府所定規格、指定行駛路段、時間、速度限制、安全注意及其他管理事項規定行駛，肇事致人重傷或死亡',3600,3600,3600,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950004','0','個人行動器具違反道路交通管理處罰條例慢車章節規定',1200,1600,1600,1600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950005','0','個人行動器具違反道路交通管理處罰條例慢車章節規定，肇事致人受傷',2400,3200,3200,3200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('6950006','0','個人行動器具違反道路交通管理處罰條例慢車章節規定，肇事致人重傷或死亡',3600,3600,3600,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7000002','0','微型電動二輪車，經依規定淘汰並公告禁止行駛後仍行駛',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7000003','0','微型電動二輪車以外其他慢車，經依規定淘汰並公告禁止行駛後仍行駛',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7110002','0','經型式審驗合格，電動輔助自行車，未黏貼審驗合格標章，於道路行駛',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7120002','0','未經型式審驗合格，電動輔助自行車，於道路行駛',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210003','0','微型電動二輪車，未經核准，擅自變更裝置',1200,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210004','0','微型電動二輪車以外其他慢車，未經核准，擅自變更裝置',300,500,500,500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210005','0','微型電動二輪車，不依規定保持煞車之良好與完整',1200,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210006','0','微型電動二輪車以外其他慢車，不依規定保持煞車之良好與完整',300,500,500,500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210007','0','微型電動二輪車，不依規定保持鈴號之良好與完整',1200,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210008','0','微型電動二輪車以外其他慢車，不依規定保持鈴號之良好與完整',300,500,500,500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210009','0','微型電動二輪車，不依規定保持燈光之良好與完整',1200,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210010','0','微型電動二輪車以外其他慢車，不依規定保持燈光之良好與完整',300,500,500,500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210011','0','微型電動二輪車，不依規定保持反光裝置之良好與完整',1200,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7210012','0','微型電動二輪車以外其他慢車，不依規定保持反光裝置之良好與完整',300,500,500,500,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7220003','0','微型電動二輪車，於道路行駛或使用，擅自增、減、變更行駛速率以外之電子控制裝置或原有規格',2500,2700,2700,2700,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7220004','0','電動輔助自行車，於道路行駛或使用，擅自增、減、變更行駛速率以外之電子控制裝置或原有規格',1800,2000,2000,2000,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7220005','0','微型電動二輪車，於道路行駛或使用，擅自增、減、變更與行駛速率相關之電子控制裝置或原有規格',5400,5400,5400,5400,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7220006','0','電動輔助自行車，於道路行駛或使用，擅自增、減、變更與行駛速率相關之電子控制裝置或原有規格',5400,5400,5400,5400,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310105','0','微型電動二輪車，不在劃設之慢車道通行',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310106','0','微型電動二輪車以外其他慢車，不在劃設之慢車道通行',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310107','0','微型電動二輪車，無正當理由在未劃設慢車道之道路不靠右側路邊行駛',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310108','0','微型電動二輪車以外其他慢車，無正當理由在未劃設慢車道之道路不靠右側路邊行駛',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310203','0','微型電動二輪車，不在規定之地區路線行駛',400,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310204','0','微型電動二輪車以外其他慢車，不在規定之地區路線行駛',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310205','0','微型電動二輪車，不在規定時間內行駛',400,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310206','0','微型電動二輪車以外其他慢車，不在規定時間內行駛',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310303','0','微型電動二輪車，不依規定轉彎',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310304','0','微型電動二輪車以外其他慢車，不依規定轉彎',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310305','0','微型電動二輪車，不依規定超車',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310306','0','微型電動二輪車以外其他慢車，不依規定超車',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310307','0','微型電動二輪車，不依規定停車',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310308','0','微型電動二輪車以外其他慢車，不依規定停車',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310309','0','微型電動二輪車，不依規定通過交岔路口',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310310','0','微型電動二輪車以外其他慢車，不依規定通過交岔路口',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310403','0','微型電動二輪車，在道路上爭先、爭道',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310404','0','微型電動二輪車以外其他慢車，在道路上爭先、爭道',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310405','0','微型電動二輪車，在道路上以其他危險方式駕車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310406','0','微型電動二輪車以外其他慢車，在道路上以其他危險方式駕車',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310503','0','微型電動二輪車，在夜間行車未開啟燈光',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310504','0','微型電動二輪車以外其他慢車，在夜間行車未開啟燈光',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310607','0','微型電動二輪車，以手持方式使用行動電話，進行撥接、通話、數據通訊',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310608','0','微型電動二輪車以外其他慢車，以手持方式使用行動電話，進行撥接、通話、數據通訊',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310609','0','微型電動二輪車，以手持方式使用行動電話有礙駕駛安全之行為',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310610','0','微型電動二輪車以外其他慢車，以手持方式使用行動電話有礙駕駛安全之行為',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310611','0','微型電動二輪車，以手持方式使用電腦，進行撥接、通話、數據通訊',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310612','0','微型電動二輪車以外其他慢車，以手持方式使用電腦，進行撥接、通話、數據通訊',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310613','0','微型電動二輪車，以手持方式使用電腦有礙駕駛安全之行為',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310614','0','微型電動二輪車以外其他慢車，以手持方式使用電腦有礙駕駛安全之行為',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310615','0','微型電動二輪車，以手持方式使用其他相類功能裝置進行撥接、通話、數據通訊',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310616','0','微型電動二輪車以外其他慢車，以手持方式使用其他相類功能裝置進行撥接、通話、數據通訊',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310617','0','微型電動二輪車，以手持方式使用其他相類功能裝置有礙駕駛安全之行為',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7310618','0','微型電動二輪車以外其他慢車，以手持方式使用其他相類功能裝置有礙駕駛安全之行為',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320005','0','微型電動二輪車，駕駛人吐氣酒精濃度達每公升0.15毫克以上，未滿0.25毫克或血液中酒精濃度達百分之0.03以上，未滿0.05',1600,1800,1800,1800,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320006','0','微型電動二輪車以外其他慢車，駕駛人吐氣酒精濃度達每公升0.15毫克以上，未滿0.25毫克或血液中酒精濃度達百分之0.03以上，未滿0.05',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320007','0','微型電動二輪車，駕駛人吐氣酒精濃度達每公升0.25毫克以上或血液中酒精濃度達百分之0.05以上',2400,2400,2400,2400,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320008','0','微型電動二輪車以外其他慢車，駕駛人吐氣酒精濃度達每公升0.25毫克以上或血液中酒精濃度達百分之0.05以上',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320009','0','微型電動二輪車，經測試檢定，有吸食毒品',2400,2400,2400,2400,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320010','0','微型電動二輪車以外其他慢車，經測試檢定，有吸食毒品',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320011','0','微型電動二輪車，經測試檢定，有吸食迷幻藥',2400,2400,2400,2400,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320012','0','微型電動二輪車以外其他慢車，經測試檢定，有吸食迷幻藥',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320013','0','微型電動二輪車，經測試檢定，有吸食麻醉藥品',2400,2400,2400,2400,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320014','0','微型電動二輪車以外其他慢車，經測試檢定，有吸食麻醉藥品',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320015','0','微型電動二輪車，經測試檢定，有吸食管制藥品',2400,2400,2400,2400,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7320016','0','微型電動二輪車以外其他慢車，經測試檢定，有吸食管制藥品',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7330002','0','微型電動二輪車，駕駛人拒絕接受酒精濃度測試之檢定',4800,4800,4800,4800,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7330003','0','微型電動二輪車以外其他慢車，駕駛人拒絕接受酒精濃度測試之檢定',4800,4800,4800,4800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7340002','0','微型電動二輪車，駕駛人未依規定戴安全帽',300,300,300,300,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410102','0','微型電動二輪車，不服從執行交通勤務警察之指揮',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410103','0','微型電動二輪車以外其他慢車，不服從執行交通勤務警察之指揮',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410104','0','微型電動二輪車，不依標誌之指示',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410105','0','微型電動二輪車以外其他慢車，不依標誌之指示',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410106','0','微型電動二輪車，不依標線之指示',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410107','0','微型電動二輪車以外其他慢車，不依標線之指示',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410108','0','微型電動二輪車，不依號誌之指示',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410109','0','微型電動二輪車以外其他慢車，不依號誌之指示',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410202','0','微型電動二輪車，在同一慢車道上，不按遵行之方向行駛',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410203','0','微型電動二輪車以外其他慢車，在同一慢車道上，不按遵行之方向行駛',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410302','0','微型電動二輪車，不依規定，擅自穿越快車道',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410303','0','微型電動二輪車以外其他慢車，不依規定，擅自穿越快車道',500,700,700,700,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410402','0','微型電動二輪車，不依規定停放車輛',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410403','0','微型電動二輪車以外其他慢車，不依規定停放車輛',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410502','0','微型電動二輪車，違規行駛人行道',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410503','0','微型電動二輪車以外其他慢車，違規行駛人行道',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410504','0','微型電動二輪車，在快車道行駛',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410505','0','微型電動二輪車以外其他慢車，在快車道行駛',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410603','0','微型電動二輪車，聞消防車、警備車、救護車、工程救險車、毒性化學物質災害事故應變車之警號不立即避讓',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410604','0','微型電動二輪車以外其他慢車，聞消防車、警備車、救護車、工程救險車、毒性化學物質災害事故應變車之警號不立即避讓',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410703','0','微型電動二輪車，行經行人穿越道有行人穿越時，未讓行人優先通行',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410704','0','微型電動二輪車以外其他慢車，行經行人穿越道有行人穿越時，未讓行人優先通行',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410705','0','微型電動二輪車，行駛至交岔路口轉彎時，未讓行人優先通行',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410706','0','微型電動二輪車以外其他慢車，行駛至交岔路口轉彎時，未讓行人優先通行',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410803','0','微型電動二輪車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行',800,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410804','0','微型電動二輪車以外其他慢車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410903','0','微型電動二輪車，聞或見大眾捷運系統車輛之聲號或燈光，不依規定避讓',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410904','0','微型電動二輪車以外其他慢車，聞或見大眾捷運系統車輛之聲號或燈光，不依規定避讓',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410905','0','微型電動二輪車，聞或見大眾捷運系統車輛之聲號或燈光，在後跟隨迫近',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7410906','0','微型電動二輪車以外其他慢車，聞或見大眾捷運系統車輛之聲號或燈光，在後跟隨迫近',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7420002','0','微型電動二輪車，行近行人穿越道，遇有攜帶白手杖或導盲犬之視覺功能障礙者時，不暫停讓視覺功能障礙者先行通過',1000,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7420003','0','微型電動二輪車以外其他慢車，行近行人穿越道，遇有攜帶白手杖或導盲犬之視覺功能障礙者時，不暫停讓視覺功能障礙者先行通過',600,800,800,800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430003','0','微型電動二輪車，違規行駛人行道，導致視覺功能障礙者受傷',1600,1800,1800,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430004','0','微型電動二輪車以外其他慢車，違規行駛人行道，導致視覺功能障礙者受傷',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430005','0','微型電動二輪車，行駛快車道，導致視覺功能障礙者受傷',1600,1800,1800,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430006','0','微型電動二輪車以外其他慢車，行駛快車道，導致視覺功能障礙者受傷',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430007','0','微型電動二輪車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行，導致視覺功能障礙者受傷',1600,1800,1800,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430008','0','微型電動二輪車以外其他慢車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行，導致視覺功能障礙者受傷',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430009','0','微型電動二輪車，違規行駛人行道，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430010','0','微型電動二輪車以外其他慢車，違規行駛人行道，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430011','0','微型電動二輪車，行駛快車道，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430012','0','微型電動二輪車以外其他慢車，行駛快車道，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430013','0','微型電動二輪車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7430014','0','微型電動二輪車以外其他慢車，於設置有必要之標誌或標線供慢車行駛之人行道上，未讓行人優先通行，導致視覺功能障礙者死亡',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500001','0','微型電動二輪車，在鐵路平交道，不遵看守人員指示，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500002','0','微型電動二輪車以外其他慢車，在鐵路平交道，不遵看守人員指示，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500003','0','微型電動二輪車，在鐵路平交道，警鈴已響、閃光號誌已顯示，或遮斷器開始放下，仍強行闖越，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500004','0','微型電動二輪車以外其他慢車，在鐵路平交道，警鈴已響、閃光號誌已顯示，或遮斷器開始放下，仍強行闖越，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500005','0','微型電動二輪車，在無看守人員管理或無遮斷器、警鈴及閃光號誌設備之鐵路平交道，設有警告標誌或跳動路面，不依規定暫停，逕行通過，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500006','0','微型電動二輪車以外其他慢車，在無看守人員管理或無遮斷器、警鈴及閃光號誌設備之鐵路平交道，設有警告標誌或跳動路面，不依規定暫停，逕行通過，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500007','0','微型電動二輪車，在鐵路平交道超車，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500008','0','微型電動二輪車以外其他慢車，在鐵路平交道超車，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500009','0','微型電動二輪車，在鐵路平交道迴車，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500010','0','微型電動二輪車以外其他慢車，在鐵路平交道迴車，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500011','0','微型電動二輪車，在鐵路平交道倒車，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500012','0','微型電動二輪車以外其他慢車，在鐵路平交道倒車，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500013','0','微型電動二輪車，在鐵路平交道臨時停車，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500014','0','微型電動二輪車以外其他慢車，在鐵路平交道臨時停車，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500015','0','微型電動二輪車，在鐵路平交道停車，未肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500016','0','微型電動二輪車以外其他慢車，在鐵路平交道停車，未肇事',1200,1400,1400,1400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500017','0','微型電動二輪車，在鐵路平交道，不遵看守人員指示，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500018','0','微型電動二輪車以外其他慢車，在鐵路平交道，不遵看守人員指示，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500019','0','微型電動二輪車，在鐵路平交道，警鈴已響、閃光號誌已顯示，或遮斷器開始放下，仍強行闖越，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500020','0','微型電動二輪車以外其他慢車，在鐵路平交道，警鈴已響、閃光號誌已顯示，或遮斷器開始放下，仍強行闖越，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500021','0','微型電動二輪車，在無看守人員管理或無遮斷器、警鈴及閃光號誌設備之鐵路平交道，設有警告標誌或跳動路面，不依規定暫停，逕行通過，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500022','0','微型電動二輪車以外其他慢車，在無看守人員管理或無遮斷器、警鈴及閃光號誌設備之鐵路平交道，設有警告標誌或跳動路面，不依規定暫停，逕行通過，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500023','0','微型電動二輪車，在鐵路平交道超車，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500024','0','微型電動二輪車以外其他慢車，在鐵路平交道超車，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500025','0','微型電動二輪車，在鐵路平交道迴車，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500026','0','微型電動二輪車以外其他慢車，在鐵路平交道迴車，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500027','0','微型電動二輪車，在鐵路平交道倒車，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500028','0','微型電動二輪車以外其他慢車，在鐵路平交道倒車，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500029','0','微型電動二輪車，在鐵路平交道臨時停車，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500030','0','微型電動二輪車以外其他慢車，在鐵路平交道臨時停車，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500031','0','微型電動二輪車，在鐵路平交道停車，因而肇事',2400,2400,2400,2400,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7500032','0','微型電動二輪車以外其他慢車，在鐵路平交道停車，因而肇事',2000,2200,2200,2200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610102','0','微型電動二輪車，慢車乘坐人數超過規定數額',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610103','0','微型電動二輪車以外其他慢車，乘坐人數超過規定數額',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610202','0','微型電動二輪車，裝載貨物超過規定重量',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610203','0','微型電動二輪車以外其他慢車，裝載貨物超過規定重量',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610204','0','微型電動二輪車，裝載貨物超出車身一定限制',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610205','0','微型電動二輪車以外其他慢車，裝載貨物超出車身一定限制',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610302','0','微型電動二輪車，裝載容易滲漏、飛散、有惡臭氣味貨物',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610303','0','微型電動二輪車以外其他慢車，裝載容易滲漏、飛散、有惡臭氣味貨物',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610304','0','微型電動二輪車，裝載危險性貨物不嚴密封固或不為適當之裝置',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610305','0','微型電動二輪車以外其他慢車，裝載危險性貨物不嚴密封固或不為適當之裝置',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610402','0','微型電動二輪車，裝載禽、畜重疊',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610403','0','微型電動二輪車以外其他慢車，裝載禽、畜重疊',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610404','0','微型電動二輪車，裝載禽、畜倒置',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610405','0','微型電動二輪車以外其他慢車，裝載禽、畜倒置',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610502','0','微型電動二輪車，裝載貨物不捆紮結實',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610503','0','微型電動二輪車以外其他慢車，裝載貨物不捆紮結實',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610602','0','微型電動二輪車，上、下乘客不緊靠路邊妨礙交通',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610603','0','微型電動二輪車以外其他慢車，上、下乘客不緊靠路邊妨礙交通',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610604','0','微型電動二輪車，裝卸貨物不緊靠路邊妨礙交通',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610605','0','微型電動二輪車以外其他慢車，裝卸貨物不緊靠路邊妨礙交通',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610703','0','微型電動二輪車，牽引其他車輛',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610704','0','微型電動二輪車以外其他慢車，牽引其他車輛',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610705','0','微型電動二輪車，攀附車輛隨行',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7610706','0','微型電動二輪車以外其他慢車，攀附車輛隨行',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620102','0','腳踏自行車，附載幼童，駕駛人未滿18歲',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620103','0','電動輔助自行車，附載幼童，駕駛人未滿18歲',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620202','0','腳踏自行車，附載之幼童年齡超過規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620203','0','電動輔助自行車，附載之幼童年齡超過規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620204','0','腳踏自行車，附載之幼童體重超過規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620205','0','電動輔助自行車，附載之幼童體重超過規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620302','0','腳踏自行車或電動輔助自行車，附載幼童，不依規定使用合格之兒童座椅',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620303','0','附載幼童，不依規定使用合格腳踏自行車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620304','0','附載幼童，不依規定使用合格之電動輔助自行車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620402','0','腳踏自行車，附載幼童，違反第76條第2項第1款至第3款以外附載幼童之規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7620403','0','電動輔助自行車，附載幼童，違反第76條第2項第1款至第4款以外附載幼童之規定',300,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7810102','0','行人不依標誌標線號誌之指示或警察指揮',500,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7810203','0','行人不在劃設之人行道通行',500,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7810204','0','無正當理由在未劃設人行道之道路不靠邊通行',500,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7810302','0','行人不依規定擅自穿越車道',500,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('7810402','0','行人於交通頻繁之道路或鐵路平交道附近阻礙交通',500,500,500,500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000101','0','非屬汽車、動力機械及個人行動器具範圍之動力載具於快車道以外之道路範圍行駛或使用',1200,1300,1400,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000111','0','非屬汽車、動力機械及個人行動器具範圍之動力運動休閒器材於快車道以外之道路範圍行駛或使用',1200,1300,1400,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000121','0','非屬汽車、動力機械及個人行動器具範圍之其他相類之動力器具於快車道以外之道路範圍行駛或使用',1200,1300,1400,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000131','0','非屬汽車、動力機械及個人行動器具範圍之動力載具於快車道行駛或使用',2000,2200,2400,2600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000141','0','非屬汽車、動力機械及個人行動器具範圍之動力運動休閒器材於快車道行駛或使用',2000,2200,2400,2600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000151','0','非屬汽車、動力機械及個人行動器具範圍之其他相類之動力器具於快車道行駛或使用',2000,2200,2400,2600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000161','0','非屬汽車、動力機械及個人行動器具範圍之動力載具於道路上行駛或使用因而肇事',2800,3000,3300,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000171','0','非屬汽車、動力機械及個人行動器具範圍之動力運動休閒器材於道路上行駛或使用因而肇事',2800,3000,3300,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('32000181','0','非屬汽車、動力機械及個人行動器具範圍之其他相類之動力器具於道路上行駛或使用因而肇事',2800,3000,3300,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71101011','0','經型式審驗合格，微型電動二輪車，未依規定領用牌照，於道路行駛',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71101021','0','未經型式審驗合格，微型電動二輪車，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71102011','0','經型式審驗合格，微型電動二輪車，使用偽造或變造之牌照，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71102021','0','未經型式審驗合格，微型電動二輪車，使用偽造或變造之牌照，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','8、1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71103011','0','經型式審驗合格，微型電動二輪車，牌照借供他車使用，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71103021','0','經型式審驗合格，微型電動二輪車，使用他車牌照，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71103031','0','未經型式審驗合格，微型電動二輪車，使用他車牌照，於道路行駛',3600,3600,3600,3600,'V','0','0','0','0','8、1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71104011','0','經型式審驗合格，微型電動二輪車，已領有牌照而未懸掛，於道路行駛',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71104021','0','經型式審驗合格，微型電動二輪車，已領有牌照而不依指定位置懸掛，於道路行駛',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71105011','0','經型式審驗合格，微型電動二輪車，牌照業經註銷，仍懸掛該註銷牌照行駛道路',1800,2000,2000,2000,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71105021','0','經型式審驗合格，微型電動二輪車，牌照業經註銷，無牌照行駛道路',1800,2000,2000,2000,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71106011','0','經型式審驗合格，微型電動二輪車，牌照遺失不報請該管主管機關補發，經舉發後仍不辦理而行駛道路',1200,1400,1400,1400,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300011','0','經型式審驗合格，微型電動二輪車，未依規定領用牌照，於道路停車',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300021','0','未經型式審驗合格，微型電動二輪車，於道路停車',3600,3600,3600,3600,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300031','0','經型式審驗合格，微型電動二輪車，使用偽造或變造之牌照，於道路停車',3600,3600,3600,3600,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300041','0','未經型式審驗合格，微型電動二輪車，使用偽造或變造之牌照，於道路停車',3600,3600,3600,3600,'V','0','0','0','0','8、1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300051','0','經型式審驗合格，微型電動二輪車，使用他車牌照，於道路停車',3600,3600,3600,3600,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300061','0','未經型式審驗合格，微型電動二輪車，使用他車牌照，於道路停車',3600,3600,3600,3600,'V','0','0','0','0','8、1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300071','0','經型式審驗合格，微型電動二輪車，已領有牌照而未懸掛，於道路停車',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300081','0','經型式審驗合格，微型電動二輪車，已領有牌照而不依指定位置懸掛，於道路停車',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300091','0','經型式審驗合格，微型電動二輪車，牌照業經註銷，仍懸掛該註銷牌照於道路停車',1500,1700,1700,1700,'V','0','0','0','0','8、f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300101','0','經型式審驗合格，微型電動二輪車，牌照業經註銷，無牌照於道路停車',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71300111','0','經型式審驗合格，微型電動二輪車，牌照遺失不報請該管主管機關補發，經舉發後仍不辦理，於道路停車',1200,1400,1400,1400,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71400011','0','經型式審驗合格並黏貼審驗合格標章，微型電動二輪車，未於本條例111年4月19日修正施行後2年內依規定登記、領用、懸掛牌照，於道路行駛',1500,1700,1700,1700,'V','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71100012','0','微型電動二輪車，損毀牌照，使不能辨認其牌號',900,1800,1800,1800,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71100022','0','微型電動二輪車，變造牌照，使不能辨認其牌號',900,1800,1800,1800,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71100032','0','微型電動二輪車，塗抹污損牌照，使不能辨認其牌號',900,1800,1800,1800,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71100042','0','微型電動二輪車，安裝其他器具之方式，使不能辨認其牌號',900,1800,1800,1800,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71201012','0','微型電動二輪車，牌照遺失，不報請補發',300,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71201022','0','微型電動二輪車，牌照破損，不報請換發或重新申請',300,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71202012','0','微型電動二輪車，牌照污穢，不洗刷清楚，非行車途中因遇雨、雪道路泥濘所致',150,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('71202022','0','微型電動二輪車，牌照為他物遮蔽',150,300,300,300,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72000041','0','微型電動二輪車，於道路行駛或使用，行駛速率超過每小時25公里，未超過35公里',900,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72000051','0','微型電動二輪車，於道路行駛或使用，行駛速率超過每小時35公里，未超過45公里',1200,1500,1500,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72000061','0','微型電動二輪車，於道路行駛或使用，行駛速率超過每小時45公里',1500,1800,1800,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72100012','0','未滿14歲之人，駕駛微型電動二輪車',1000,1200,1200,1200,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72100022','0','未滿14歲之人，駕駛個人行動器具',600,800,800,800,'0','0','0','0','0','f',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72200012','0','微型電動二輪車租賃業者，未於租借予駕駛人前，教導駕駛人車輛操作方法及道路行駛規定',800,1200,1200,1200,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('72200022','0','個人行動器具租賃業者，未於租借予駕駛人前，教導駕駛人車輛操作方法及道路行駛規定',600,800,800,800,'V','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

	End if
	rsChkL2.close
	Set rsChkL2=Nothing

	strChkL2="select * from Law where itemid ='29300012' and illegalrule='汽車裝載貨物超過核定之總重量、總聯結重量' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If Not rsChkL2.eof Then
		strInsL2="update Law set illegalrule='汽車裝載貨物超過核定之總重量、總聯結重量' where itemid ='29300012' and version=2"
		conn.execute strInsL2
	
	End if
	rsChkL2.close
	Set rsChkL2=Nothing
'1120331=====================================================================
	strChkL2="select * from Law where itemid ='4420003' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="insert into law values('4420003','3','駕駛汽車行經行人穿越道有行人穿越時，不暫停讓行人先行通過',1200,1300,1400,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4420003','4','駕駛汽車行經行人穿越道有行人穿越時，不暫停讓行人先行通過',3600,3600,3600,3600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4430003','3','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',2400,2600,2800,3100,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4430003','5','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',4800,5200,6400,7200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4430003','6','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',7200,7200,7200,7200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4430004','3','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',2400,2600,2800,3100,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4430004','5','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',4800,5200,6400,7200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4430004','6','汽車駕駛人駕駛汽車行近行人穿越道遇有攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',7200,7200,7200,7200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510904','3','支線道車不讓幹線道車先行',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510904','5','支線道車不讓幹線道車先行',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510904','6','支線道車不讓幹線道車先行',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510905','3','少線道車不讓多線道車先行',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510905','5','少線道車不讓多線道車先行',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510905','6','少線道車不讓多線道車先行',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510906','3','車道數相同時，左方車不讓右方車先行',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510906','5','車道數相同時，左方車不讓右方車先行',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4510906','6','車道數相同時，左方車不讓右方車先行',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511507','3','行經無號誌交岔路口不依規定',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511507','5','行經無號誌交岔路口不依規定',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511507','6','行經無號誌交岔路口不依規定',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511508','3','行經無號誌交岔路口不依標誌指示',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511508','5','行經無號誌交岔路口不依標誌指示',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511508','6','行經無號誌交岔路口不依標誌指示',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511509','3','行經無號誌交岔路口不依標線指示',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511509','5','行經無號誌交岔路口不依標線指示',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511509','6','行經無號誌交岔路口不依標線指示',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511510','3','行經巷道不依規定',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511510','5','行經巷道不依規定',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511510','6','行經巷道不依規定',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511511','3','行經巷道不依標誌指示',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511511','4','行經巷道不依標誌指示',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511511','5','行經巷道不依標誌指示',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511512','3','行經巷道不依標線指示',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511512','5','行經巷道不依標線指示',1500,1600,1700,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4511512','6','行經巷道不依標線指示',1800,1800,1800,1800,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4820003','3','汽車駕駛人轉彎時，除禁止行人穿越路段外，不暫停讓行人優先通行',1200,1300,1400,1500,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4820003','4','汽車駕駛人轉彎時，除禁止行人穿越路段外，不暫停讓行人優先通行',3600,3600,3600,3600,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830003','3','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',2400,2600,2800,3100,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830003','5','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',4800,5200,6400,7200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830003','6','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶白手杖之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',7200,7200,7200,7200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830004','3','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',2400,2600,2800,3100,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830004','5','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',4800,5200,6400,7200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4830004','6','汽車駕駛人轉彎時除禁止行人穿越路段外行近攜帶導盲犬之視覺功能障礙者不暫停讓視覺功能障礙者先行通過',7200,7200,7200,7200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('6020304','3','不遵守道路交通號誌之指示(遇閃光紅燈未停車再開)',1200,1300,1400,1500,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('6020304','5','不遵守道路交通號誌之指示(遇閃光紅燈未停車再開)',1500,1600,1700,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('6020304','6','不遵守道路交通號誌之指示(遇閃光紅燈未停車再開)',1800,1800,1800,1800,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('6020305','0','不遵守道路交通號誌之指示(其他)',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

	End if
	rsChkL2.close
	Set rsChkL2=Nothing

'1130630=====================================================================
	if now>="2024/6/30" then
	strChkL2="select * from Law where itemid ='5000205' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		
		strInsL2="insert into law values('5000205','0','倒車前未顯示倒車燈光',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5000206','0','倒車時不注意其他車輛或行人',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5000305','6','大型汽車無人在後指引時，不先測明車後有足夠之地位',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5000306','6','大型汽車無人在後指引時，不促使行人避讓',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510115','3','在橋樑臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510115','4','在橋樑臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510116','3','在隧道臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510116','4','在隧道臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510117','3','在圓環臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510117','4','在圓環臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510118','3','在障礙物對面臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510118','4','在障礙物對面臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510119','3','在快車道臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510119','4','在快車道臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510120','3','在騎樓以外之人行道臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510120','4','在騎樓以外之人行道臨時停車',600,600,600,600,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510121','3','在騎樓臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510121','4','在騎樓臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510122','3','在行人穿越道臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('5510122','4','在行人穿越道臨時停車',600,600,600,600,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510211','3','在消防車出入口五公尺內臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510211','4','在消防車出入口五公尺內臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510212','3','在交岔路口十公尺內臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510212','4','在交岔路口十公尺內臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510213','3','在公共汽車招呼站十公尺內臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510213','4','在公共汽車招呼站十公尺內臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510407','3','併排臨時停車',500,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5510407','4','併排臨時停車',600,600,600,600,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610118','3','在公共汽車招呼站十公尺內停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610118','4','在公共汽車招呼站十公尺內停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610119','3','在橋樑停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610119','5','在橋樑停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610119','6','在橋樑停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610120','3','在隧道停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610120','5','在隧道停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610120','6','在隧道停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610121','3','在圓環停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610121','5','在圓環停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610121','6','在圓環停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610122','3','在障礙物對面停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610122','5','在障礙物對面停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610122','6','在障礙物對面停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610123','3','在行人穿越道停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610123','5','在行人穿越道停車',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610123','6','在行人穿越道停車',1200,1200,1200,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610124','3','在快車道停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610124','5','在快車道停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610124','6','在快車道停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610125','3','在交岔路口十公尺內停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610125','5','在交岔路口十公尺內停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610125','6','在交岔路口十公尺內停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610126','3','在消防車出入口五公尺內停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610126','5','在消防車出入口五公尺內停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610126','6','在消防車出入口五公尺內停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610127','3','在騎樓以外之人行道停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610127','5','在騎樓以外之人行道停車',900,1000,1100,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610127','6','在騎樓以外之人行道停車',1200,1200,1200,1200,'0','1','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610312','3','在消防栓之前停車',600,700,800,900,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610312','5','在消防栓之前停車',900,1000,1100,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('5610312','6','在消防栓之前停車',1200,1200,1200,1200,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200271','3','機車駕駛人行駛道路以手持方式使用行動電話進行撥接',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200281','3','機車駕駛人行駛道路以手持方式使用行動電話進行通話',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200291','3','機車駕駛人行駛道路以手持方式使用行動電話進行數據通訊',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200301','3','機車駕駛人行駛道路以手持方式使用行動電話進行有礙駕駛安全之行為',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200311','3','機車駕駛人行駛道路以手持方式使用電腦進行撥接',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200321','3','機車駕駛人行駛道路以手持方式使用電腦進行通話',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200331','3','機車駕駛人行駛道路以手持方式使用電腦進行數據通訊',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200341','3','機車駕駛人行駛道路以手持方式使用電腦進行有礙駕駛安全之行為',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200351','3','機車駕駛人行駛道路以手持方式使用相類功能裝置進行撥接',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200361','3','機車駕駛人行駛道路以手持方式使用相類功能裝置進行通話',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200371','3','機車駕駛人行駛道路以手持方式使用相類功能裝置進行數據通訊',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('31200381','3','機車駕駛人行駛道路以手持方式使用相類功能裝置進行有礙駕駛安全之行為',1000,1000,1000,1000,'0','0','0','0','0','0',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2


		strInsL2="update law set illegalrule='駕駛汽車行近行人穿越道有行人穿越時，不暫停讓行人先行通過' where itemid='4420004'"
		conn.execute strInsL2

		strInsL2="update law set recordstateid=-1 where itemid in ('1610301','1610401','1610402','1610403','1610404','1610501','1610502','2110101','2110102','2110201','2110301','2110302','2110303','2110304','2110305','2110306','2110401','2110402','2110501','2130001','2130002','2130003','2130004','2130005','2130006','2130007','2130008','2150001','2150002','2150003','2150004','2150005','2150006','2150007','2150008','2150009','2150010','2150011','2150012','2150013','2150014','2150015','2210501','2230001','2230002','2230003','2230004','2230005','2230006','2230007','2230008','2230009','2230010','2230011','2230012','2230013','2230014','2230015','2300201','2440001','3310101','3310102','3310103','3310104','3310105','3310106','3310107','3310108','3310109','3310110','3310111','3310112','3310113','3310114','3310115','3310116','3310117','3310118','3310119','3310120','3310121','3310122','3310123','3310124','3310125','3310126','3310127','3310128','3310129','3310130','3310131','3310132','3310201','3310202','3310203','3310204','3310401','3310402','3310403','3310404','3310607','3310701','3310702','3310703','3310704','3310705','3310706','3310707','3310708','3310709','3310710','3310711','3310712','3310713','3310714','3310715','3310901','3310902','3311105','3311106','3311601','3311602','3311603','3311604','3311605','3311606','3311701','3311702','3400001','3400002','3400003','3400004','3400005','3400006','3400013','3400014','3400015','3400016','3400017','3400018','3400019','3810001','3810002','3810003','4200001','4200002','4200003','4310105','4310106','4310107','4310108','4310113','4310114','4310210','4310211','4310212','4310213','4310214','4310215','4310216','4310217','4310218','4310219','4310220','4310221','4310222','4310223','4310224','4310225','4310226','4310227','4310228','4310229','4310230','4310231','4310232','4310233','4310234','4310235','4310236','4310237','4310238','4310239','4310307','4310308','4310309','4310310','4310311','4310312','4310313','4310314','4310315','4310316','4310317','4310318','4310319','4310320','4310321','4310322','4310323','4310324','4310401','4310402','4310403','4310404','4310405','4310406','4310407','4310408','4310409','4310410','4310411','4310412','4310413','4310414','4310415','4310416','4310417','4310418','4330003','4330008','4330013','4330018','4330029','4330030','4330031','4330032','4340011','4340014','4340035','4340044','4340045','4340056','4340057','4420003','4430003','4430004','4700101','4700102','4700103','4700104','4700105','4700106','4700107','4700108','4700109','4700110','4700111','4700112','4700201','4700202','4700203','4700204','4700205','4700206','4700207','4700208','4700209','4700210','4700211','4700212','4700301','4700302','4700303','4700304','4700305','4700306','4700401','4700402','4700403','4700404','4700501','4700502','4700503','4700504','4700505','4700506','4700507','4700508','4820003','4830003','4830004','5000201','5000202','5000301','5000302','5320001','5510101','5510102','5510103','5510104','5510105','5510106','5510107','5510201','5510202','5510203','5510401','5510404','5610102','5610103','5610310','5620002','6020304','6020305','6110401','6110403','6130001','6130002','6330001','6330002','6330003','6330004','6720011','6820004','7430003','7430005','7430007','7430009','7430010','7430011','7430012','7430013','7430014','8230101','8230102','8230103','8230104','8310101','8310201','8310301','8410101','8540001','8540002','8540003','8540004','21101021','21102021','21103021','21104021','21104041','21105021','21105041','21106021','21106041','21106061','21107081','30101001','30101002','30101003','30101004','30107001','30107002','30107003','30107004','31100011','31100021','31100031','31100041','31100051','31100061','31100071','31100081','31100091','31100101','31100111','31100121','31100131','31100141','31200011','31200021','31200031','31200041','31200051','31200061','31200071','31200081','31200091','31200101','31200111','31200121','31200131','31200141','56000011','56000021','56000031','56000041','56000051','56000061','56000071','56000081','56000091','56000101','56000111','56000121','56000131','56000141','56000151','56000161','56000171','56000181','56000191','56000201','56000211','56000221','56000231','56000241','63000011','63100011','351000031','352000021','353000011','4410201','4000007','4000010','4000013','4000016','4810103','4810104','4810105','4810113','4810114','4810115','4810201','4810301','4810401','4810402','4810501','4810502','4810601','4810602','4810701','4810702','4810703','5610116','5510204','5510205','5610110')"
		conn.execute strInsL2

		
	End if
	rsChkL2.close
	Set rsChkL2=Nothing
	end if

	strChkL2="select * from Law where itemid ='6910101' and recordstateid=-1 and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		strInsL2="update law set recordstateid=-1 where itemid in ('6910101','6920001','6920002','7010101','7000001','7110101','7100001','7210101','7210102','7200001','7200002','7220001','7220002','7300101','7300102','7300201','7300301','7300401','7300501','7310103','7310104','7310202','7310302','7310402','7310502','7310601','7310602','7310603','7310604','7310605','7310606','7320002','7320003','7320004','7330001','7340001','7410101','7410201','7410301','7410401','7410501','7410601','7410602','7410701','7410702','7410801','7410802','7410901','7410902','7420001','7430001','7430002','7510101','7510102','7610101','7610201','7610301','7610401','7610501','7610601','7610702','7620101','7620201','7620301','7620401','7710001','7710002','7710003','7710004','7710005','7710006','7710007','7710008','7710009','7710010','7710011','7710012','7710013','7710014','7710101','7720001','7720002','7720101','7810101','7810201','7810202','7810301','7810401','8110101','32000011','32000021','32000031','32000041','32000051','32000061','32000071','32000081','32000091','72000011','72000021','72000031','81101011','81101021','35300079','35300080','35300081','35300082','35300083','35300084','35300085','35300086','35300087','35300088','35300089','35300090','35300091','35300092','35300093','35300094','35300095','35300096','35300097','35300098','35300099','35300100','35300101','35300102','35300103','35300104','35300105','35300106','35300107','35300108','35300109','35300110','35300111','35300112','35300113','35300114','35300115','35300116','35300117','35300118','35300119','35300120','35300121','35300122','35300123','35300124','35300125','35300126','35300127','35300128','35300129','35300130','35300131','35300132','35300133','35300134','35300135','35300136','35300137','35300138','35300139','35300140','35300141','35300142','35300143','35300144','35300145','35300146','35300147','35300148','35300149','35300150','35300151','35300152','35300153','35300154','35300155','35300156','35300157','35300158','35300159','35300160','35300161','35300162','35300163','35300164','35300165','35300166','35300167','35300168','35500001','35500002','35500003','35500004','35500005','35500006','35500007','35500008','35500009','35500010','35500011','35500012','35700001','35700002','35700003','35700004','35700005','35700006','35700007','35700008','35800001','35800002','35900001','351000011','351000011','351000021','352000011','3120019','3120020','18100041','18200041','2910104','2910105','2910106','2910107','2910108','2910109','2910110','2910111','2910112','2910209','2910210','2910211','2910212','2910213','2910214','2910215','2910216','2910217','2910218','2910219','2910220','2910221','2910222','2910223','2910224','2910225','2910226','2910227','2910228','2910229','2910230','2910231','2910232','2910307','2910310','2910311','2910312','2910313','2910314','2910315','2910316','2910317','2910318','2910319','2910320','2910324','2910325','2910326','2930001','4310503','4310504','4310505','4310506','4340001','4340002','4340003','4340004','4340005','4340006','4340007','4340008','4340017','4340018','4340021','4340022','4340023','4340024','4340025','4340026','4340033','4340034','18300011','3510148','3510149','3510150','3510151','3510152','3510153','3510154','3510155','3510156','3510157','3510158','3510161','3510162','3510163','3510170','3510173','3510174','3510175','3510176','3510177','3510178','3510179','3510180','3510181','3510182','3510183','3510184','3510185','3510186','3510187','3510188','3510189','3510190','3510191','3510192','3510193','3510194','3510195','3510196','3510197','3510301','3510302','3510303','3510304','3510305','3510306','3510307','3510308','3510309','3520001','3520002','3520003','3520004','3520005','3520006','3520007','3520008','3520009','3520010','3520011','3520012','3520029','3520030','3520031','3520032','3520033','3520034','3520035','3520036','3520037','3520038','3520039','3520040','3520041','3520042','3520043','3520044','3520045','3520046','3520047','3520048','3520049','3520050','3520051','3520052','3520053','3520054','3520055','3520056','3520057','3520058','3520059','3520060','3520061','3520062','3520063','3520064','3520065','3520066','3530016','3530017','3530018','3530019','3530020','3530021','3530022','3530023','3530024','3530025','3530026','3530027','3530028','3530029','3530030','3530031','3530032','3530033','3530034','3530035','3530036','3530037','3530038','3530039','3530040','3530041','3530042','3530043','3530044','3530045','3530046','3530047','3530048','3530049','3530050','3530051','3530052','3530053','3530054','3530055','3530056','3530057','3530058','3530059','3530060','3530061','3530062','3530063','3530064','3530065','3530066','3530067','3530068','3530069','3530070','3530071','3530072','3530073','3530074','3530075','3530076','3530077','3530078','3530079','3530080','3530081','3530082','3530083','3530084','3530085','3530086','3530087','3530088','3530089','3530090','3530091','3530092','3530093','3540031','3540032','3540033','3540034','3540035','3540036','3540037','3540038','3540039','3540040','3540041','3540042','3540043','3540044','3540045','3540046','3540047','3540048','3540049','3540050','3540051','3540052','3540053','3540054','3540055','3540056','3540057','3540058','3540059','3540060','3540061','3540062','3540063','3540064','3540065','3540066','3540067','3540068','3540069','3540070','3540071','3540072','3540073','3540074','3540075','3540076','3540077','3540078','3540079','3540080','3540081','3540082','3540083','3540084','3540085','3540086','3540087','3540088','3540089','3540090','3560012','3560013','3560014','3560015','7310701','7320001','35010201','35010202','35010203','35010204','35010205','35010206','35010207','35010208','35010209','35010210','35010211','35010212','35010213','35010214','35010215','35010216','35010217','35010218','35010219','35010220','35010221','35010222','35010223','35010224','35010225','35010226','35010227','35010228','35010229','35010230','35010231','35010232','35010233','35010234','35010235','35010236','35010237','35010238','35010239','35010240','35020001','35020002','35020003','35020004','35020005','35020006','35020007','35020008','35020009','35020010','35020011','35020012','35020013','35020014','35020015','35020016','35060001','35060002','35060003','35060004','4810101','4810102','4810111','4810112','4420002','4430001','4430002','4510901','4510902','4510903','4511501','4511502','4511503','4511504','4511505','4511506','4820002','4830001','4830002','6020303')"
		conn.execute strInsL2
	End if
	rsChkL2.close
	Set rsChkL2=Nothing

if now>="2025/6/30" then
	strChkL2="select * from Law where itemid ='4440013' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		
		strInsL2="insert into law values('4440013','0','汽車駕駛人有違反44條第2項規定之情形，因而肇事致人受傷',18000,20000,24000,30000,'0','0','L','0','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('4440014','0','汽車駕駛人有違反44條第2項規定之情形，因而肇事致人受傷(無駕駛執照)',18000,20000,24000,30000,'0','0','0','1','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4440015','0','汽車駕駛人有違反44條第2項規定之情形，因而肇事致人重傷',36000,36000,36000,36000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4440016','0','汽車駕駛人有違反44條第2項規定之情形，因而肇事致人重傷(無駕駛執照)',36000,36000,36000,36000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4440017','0','汽車駕駛人有違反44條第3項規定之情形，因而肇事致人受傷',18000,20000,24000,30000,'0','0','L','0','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4440018','0','汽車駕駛人有違反44條第3項規定之情形，因而肇事致人受傷(無駕駛執照)',18000,20000,24000,30000,'0','0','0','1','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4440019','0','汽車駕駛人有違反44條第3項規定之情形，因而肇事致人重傷',36000,36000,36000,36000,'0','0','2','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('4440020','0','汽車駕駛人有違反44條第3項規定之情形，因而肇事致人重傷(無駕駛執照)',36000,36000,36000,36000,'0','0','0','3','0','6',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
	

	End if
	rsChkL2.close
	Set rsChkL2=Nothing

	strChkL2="select * from Law where itemid ='35900026' and version=2"
	Set rsChkL2=conn.execute(strChkL2)
	If rsChkL2.eof Then
		
		strInsL2="insert into law values('35900026','0','汽機車駕駛人有第35條第1、3、4、5項之情形之一',0,0,0,0,'V','0','X','0','0','8',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2
		
		strInsL2="insert into law values('35900027','0','汽機車駕駛人有第35條第1、3、4、5項之情形之一因而肇事致人重傷',0,0,0,0,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="insert into law values('35900028','0','汽機車駕駛人有第35條第1、3、4、5項之情形之一因而肇事致人死亡',36000,36000,36000,36000,'V','0','0','0','0','1',sysdate,'2',0,sysdate,null,null)"
		conn.execute strInsL2

		strInsL2="update law set recordstateid=-1 where itemid in ('5400110','5400111','5400112','5400113','5400114','5400115','5400116','5400117','5400118','5400207','5400208','5400209','5400210','5400211','5400212','5400316','5400317','5400318','5400319','5400320','5400321','5400322','5400323','5400324','5400325','5400326','5400327','5400328','5400329','5400330','29400012','29400022','29400032','29400042','29400052')"
		conn.execute strInsL2

	End if
	rsChkL2.close
	Set rsChkL2=Nothing
	
	
end if

	

'獎勵金記得要加法條點數==================================================================

%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</head>
<body leftmargin="25" topmargin="5" marginwidth="0" marginheight="0" <%
	if sys_City<>"台中市" and sys_City<>"台中縣" then
%>
		onLoad="init()"
<%
	end if
%>>
<% if Session("Group_ID") <> "8984" and Session("Group_ID") <> "8985" and Session("Group_ID") <> "8987" then %>
<div id="D1" style="width:350px">
<%
	if sys_City<>"台中市" and sys_City<>"台中縣" then
%>
  <table border="0" width="350">
    <TBODY> 

	<tr>
		<!-- 處理進度 -->
		<td valign="top"> 
			<table width="100%" border="0">
				<tr bgcolor="#CAD6E2">
					<td >處理進度(1~68條)</td>
				</tr>
				<tr bgColor="#FFFFFF">
					<td id="YestodayLayer">

					</td>
				</tr>
				<tr bgColor="#FFFFFF">
					<td id="TodayLayer">
					
					</td>
				</tr>
			</table>
		</td>
		<td align="middle" bgColor="#C36A2D" rowSpan="2" width="20" >
			<font color="#FFFFFF"> 
			個<br>人<br>資<br>料
			</font>
		</td>
    </tr>
    <tr>
		<!-- 上傳紀錄 -->
		<td valign="top" id="UpLoadLayer"> 

		</td>
    </tr>
    </TBODY> 
	<%If sys_City="高雄市X" then%>
	<tr>
		<td height="5"></td>
	</tr>
	<tr>
	<td><iframe frameborder="0" width="320px" height="200px" src="chat/onlinelist.asp"></iframe></td>
		<td align="middle" bgColor="#FF0000" rowSpan="2" width="20" >
			<font color="#FFFFFF"> 
			線<br>上<br>人<br>員
			</font>
		</td>
<%End If%>
  </table>
<%
	end if
%>
</div>
<% end if %>
<font color="#000000"><b>
<%

strUnit="select * from UnitInfo where UnitID='"&UnitNo&"'"
set rsUnit=conn.execute(strUnit)
UnitName=rsUnit("UnitName")
rsUnit.close
set rsUnit=nothing

%>
</b>
</font>
<div style="position: absolute;top:0px;">
<img src="/traffic/image/banner.jpg">
</div>
<table width='980' border='0' align="center" >
<%If sys_City="高雄市" then%>

	<tr>
		<td colspan="5">
		<!--  -->
		<!--2024 01 19 smith mark src= ArgueCase/CaseNotifyPublic.asp-->
		<iframe frameborder="0" width="970px" height="150px" src=""></iframe>
		</td>
	</tr>
<%End If%>

	<tr>
		<td colspan="5" height="40">
			<div id="Layer1" style="top:60px;">
			<table width="230" align="left" border="0">
				<tr>
					<td class="style1" nowrap><strong>
					<%if sys_City<>"台中市" and sys_City<>"高雄市" and sys_City<>"苗栗縣" then%>
					<!--<div style="text-align:right;width:100%"><img src="Image/dot.gif" ><font style="font-size:12pt;line-height:20px;font-weight:bold;">電話客服時間 週一~週五(國定假日除外)</font> <font style="font-size:14pt;line-height:24px;font-weight:bold;">上午09:00~12:00 下午01:00~05:00</font></div>-->
					
					<%elseif sys_City="台中市" then %>
					<div style="text-align:right;width:100%"><img src="Image/dot.gif" ><font style="font-size:14pt;line-height:24px;font-weight:bold;">有關舉發案件處理流程問題(如退件等)，請聯繫交通大隊執法組 04-23289524，謝謝</font></div>
					
					<%end if%>
						<%=UnitName%>
						<img src="image/space.gif" alt="" width="5" height="2" border="0" align="baseline">
						<%=memName%><%
			'If sys_City="屏東縣" then
				strMpID="select MpoliceID from memberdata where memberid="& Trim(session("User_ID"))
				Set rsMpID=conn.execute(strMpID)
				If Not rsMpID.eof Then
					response.write "&nbsp; "&Trim(rsMpID("MpoliceID"))
				End If 
				rsMpID.close
				Set rsMpID=Nothing 
			'End if
				%></strong>
						
						<img src="image/space.gif" alt="" width="5" height="1" border="0" align="baseline">
						<input type="button" name="b10010" value="修改個人資料" onclick="funMember();" style="font-size: 9pt; width: 90px; height:23px;">
				  		<!--<img src="Image/dot.gif" >
				  		<a href="SystemInit.htm" target="_blank" class="style2">** 首次使用系統注意事項 ** </a>			  -->
				  	
				    	<!--<img src="Image/dot.gif" ><font class="style2">
				   		客服 :<%
					If sys_City="高雄市" Then
						response.write "(請優先打警用)警用 3985 找工程師、"
					End If 
						%> (02) 2790-0989
				      <img src="Image/space.gif" width="5" height="1" >
				      <img src="Image/dot.gif" ondblclick="location='UserLogout_Contral.asp'">
				      傳真 : (02) 2790-3616 		<img src="Image/space.gif" width="5" height="1" >
						<img src="Image/dot.gif" ><font class="style2">信箱<b>  178hyndai@gmail.com </b></font>
					<br>-->
					<%if sys_City<>"台中市" and sys_City<>"高雄市" and sys_City<>"苗栗縣" then%>
					<!-- <div style="text-align:right;width:100%"><font style="font-size:12pt;line-height:20px;font-weight:bold;">電話客服時間 週一~週五(國定假日除外)</font> <font style="font-size:14pt;line-height:24px;font-weight:bold;">上午09:00~12:00 下午01:00~05:00</font></div> -->
					<%elseif sys_City="台中市" then %>
					<!-- <div style="text-align:right;width:100%"><font style="font-size:12pt;line-height:20px;font-weight:bold;">有關舉發案件處理流程問題(如退件等)，請聯繫交通大隊執法組 04-23289524，謝謝</font></div> -->
					<%end if%>
				    <%
					If sys_City="高雄市" Then
					%>
						<br/><div style="text-align:center;width:100%"><img src="Image/dot.gif" ><a href="NOTICEMain.asp" target="_blank" style="color: #FF0000;font-size: 32px;">公告訊息(請點選)</a></div>	
					<%
					End If 
					%>
				      
				      <!---加上 檢查 flash_recovery_area_usage 是否快滿的顯示--->
				
				      
					</td>
				</tr>
				
				<tr>
					<td class="style1" nowrap><strong>
					
					<img src="image/space.gif" alt="" width="5" height="1" border="0" align="baseline">
					<%
'					strGName="select Content from Code where TypeID=10 and ID="&trim(GroupID)
'					set rsGName=conn.execute(strGName)
'					if not rsGName.eof then
'						response.write trim(rsGName("Content"))
'					end if
'					rsGName.close
'					set rsGName=nothing
									
					%></strong>
					</td>
				</tr>
				
			</table>
			</div>
				<img src="image/space.gif" alt="" width="300" height="2" border="0" align="baseline">
				<!-- <input type="button" name="update1" value="Update" onclick="funcUpdate();"> -->
					<!-- <input type="button" name="b10010" value="登 出" onclick="location='UserLogout_Contral.asp'" > --></span>
		     <!-- <div id="Layer1">
		     <span class="style1">
			  	<table width="181" align="left">
				<tr>
				<td width="93" align="center" class="style2">相片剩餘空間：</td>
				<td width="85" align="right" class="style2"> --><%
'			  set fs=Server.CreateObject("Scripting.FileSystemObject")
'			  FileName=Server.MapPath("storespace.ini")
'			  
'			  if fs.FileExists(FileName) then
'			  	set txtf=fs.OpenTextFile(FileName)
'				if not txtf.atEndOfStream then
'					PicFree=txtf.ReadAll
'					response.write PicFree
'				end if
'				txtf.close
'			  	set txtf=nothing
'			  end if
'			  set fs=nothing
			  %>
			  
			  <!-- &nbsp;MB
			  	</td>
				</tr>
				<tr>
				<td align="center" class="style2">資料剩餘空間：</td>
				<td align="right" class="style2"> --><%
'			  strDB="SELECT a.Tablespace_Name,a.Bytes / 1024 / 1024 TotalMb," &_
'			  	"b.Bytes / 1024 / 1024 UsedMb,c.Bytes / 1024 / 1024 FreeMb," &_
'				"(c.Bytes * 100) / a.Bytes ""% FREE""  FROM Sys.Sm$ts_Avail a, Sys.Sm$ts_Used b,  Sys.Sm$ts_Free c" &_
'				" WHERE a.Tablespace_Name = b.Tablespace_Name  AND a.Tablespace_Name = c.Tablespace_Name" &_
'				" And a.Tablespace_Name='TRAFFIC'"
'				set rsDB=conn.execute(strDB)
'				if not rsDB.eof then
'					response.write rsDB("FreeMb")
'				end if
'				rsDB.close
'				set rsDB=nothing
			  %>
			  
			  <!-- &nbsp;MB
			  	</td>
				</tr>
			  </table>
			  </span> 
		    </div> -->

		    

					<!--
				  <img src="Image/dot.gif" >
				  <a href="Help.rar" class="style2">使用手冊 V3 (0504) </a>
				  -->	
<br><br><br><br><br><br>

<hr>
<font size="3"><strong><div id="LayerTime" ondblclick="NowDownProcess();"></div></strong></font>
	  <td>
 
	</tr>
	<tr>
		<td colspan="5" class="style3">
		<%if sys_City<>"高雄市" and sys_City<>"台中市" and sys_City<>"苗栗縣" then%>
			<!--<span style="font-size:18pt;line-height:26px;color:#CC0000">
			如有系統問題請多加利用客服信箱(<b>  178hyndai@gmail.com </b>)作回報，儘量不要使用客服電話回報問題，感謝您的配合
			</span>
				</br>-->
		<%End if%>
			<%if sys_City<>"苗栗縣" and sys_City<>"新竹市" and not (sys_City="花蓮縣" And Trim(session("Group_ID"))="301012") then%>
			<font color="red"><strong>快速查詢</strong></font>
			<input type="hidden" name="creditidhidden" value="@@@<%=Session("Credit_ID")%>^^^" >
			&nbsp;<strong>舉發單號</strong>&nbsp;<input type="text" name="BillNo" size="10" maxlength="9" onkeyup="EnterBillQry();">
			&nbsp;<strong>車號</strong>&nbsp;<input type="text" name="CarNo" size="8" maxlength="9" onkeyup="EnterBillQry();">
			<!-- 花蓮拖吊場限制查詢條件 -->
			<% if Session("Credit_ID")="0000" AND sys_City="花蓮縣" then %>
				&nbsp;<strong></strong>&nbsp;<input type="hidden" name="IllegalName" size="12" maxlength="30">
				&nbsp;<strong></strong>&nbsp;<input type="hidden" name="IllegalID" size="10" maxlength="10">	

             <% else %>
				&nbsp;<strong>姓名</strong>&nbsp;<input type="text" name="IllegalName" size="12" maxlength="30"  onkeyup="EnterBillQry();">
				&nbsp;<strong>身份證號</strong>&nbsp;<input type="text" name="IllegalID" size="10" maxlength="10" onkeyup="EnterBillQry();">	


			 <% end if %>
			<strong>查詢事由</strong>
			<select name="QryReason" >
				<option value="" >請選擇</option>
			<%If sys_City<>"屏東縣" then%>
				<option value="資料檢核" <%If Trim(request("QryReason"))="資料檢核" Then response.write "selected" End if%>>資料檢核</option>
			<%End If %>
				<option value="執行業務" <%If Trim(request("QryReason"))="執行業務" Then response.write "selected" End if%>>執行業務</option>
			<%If sys_City="屏東縣" then%>
				<option value="民眾申訴(來電)" <%If Trim(request("QryReason"))="民眾申訴(來電)" Then response.write "selected" End if%>>民眾申訴(來電)</option>
				<option value="民眾申訴(臨櫃)" <%If Trim(request("QryReason"))="民眾申訴(臨櫃)" Then response.write "selected" End if%>>民眾申訴(臨櫃)</option>
				<option value="民眾申訴(公文)" <%If Trim(request("QryReason"))="民眾申訴(公文)" Then response.write "selected" End if%>>民眾申訴(公文)</option>
			<%else%>
				<option value="民眾申訴" <%If Trim(request("QryReason"))="民眾申訴" Then response.write "selected" End if%>>民眾申訴</option>
			<%End If %>	
			<%If sys_City="屏東縣" then%>
				<option value="資料檢核" <%If Trim(request("QryReason"))="資料檢核" Then response.write "selected" End if%>>資料檢核</option>
			<%End If %>
			<%If sys_City<>"嘉義縣" then%>
				<option value="事故處理" <%If Trim(request("QryReason"))="事故處理" Then response.write "selected" End if%>>事故處理</option>
			<%End If %>	
			<%If sys_City<>"台中市" and sys_City<>"屏東縣" and sys_City<>"嘉義縣" and sys_City<>"新竹市" and sys_City<>"雲林縣" and sys_City<>"嘉義市" then%>
				<option value="偵查刑案" <%If Trim(request("QryReason"))="偵查刑案" Then response.write "selected" End if%>>偵查刑案</option>
			<%End If %>
			
			</select>
			<%If sys_City="高雄市" or sys_City="屏東縣" or sys_City="台東縣" then%>
			<br>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; 
			<strong>代查詢人</strong>
						<input type="text" size="12" value="" name="ForChkMen" onkeyup="EnterBillQry();" >
			<%end if%>
			<input type="button" value="查詢" <%
			If sys_City="台東縣" And Trim(session("Unit_ID"))="Z000" Then

					
			Else
				response.write "onclick='getBillData();'"
			End if
			%> <%
			If sys_City="台東縣" Then
				If Trim(session("Unit_ID"))="Z000" Then
					response.write "disabled"
				End If 
			End if
			%>>
			<%if sys_City<>"高雄市" then%>
			<span class='style2'><b>可跨分局查詢</b></span>
			<%end If%>
			<%end If%>
			
			
			<%if sys_City="台中市" Then%>
			<br>			
			<a href="BillUnitMem.asp" target="_blank" class="style2">** 各單位舉發單職名章主管人員 ** </a>
			&nbsp;&nbsp;
			<a href="MailNotBack.doc" target="_blank" class="style2">** 郵寄未退回清冊使用手冊 ** </a>	
			&nbsp; &nbsp; &nbsp; &nbsp;&nbsp;
			<strong>告示單號</strong>&nbsp;<input type="text" name="ReportNo" size="15" onkeyup="EnterBillQry();" >
			
			<%End if%>
			<%'If sys_City="彰化縣" then%>
			<!--<br>
			<font color="red"><strong>&nbsp; &nbsp; &nbsp;未入案案件及入案異常案件，無法使用快速查詢，請至『舉發單資料維護系統』查詢案件
			<br>-->
			<!-- <font style="font-size: 28px;line-height:32px;"> -->
			<!--&nbsp; &nbsp; &nbsp;如要查詢案件入案異常原因，請至『上傳下載資料查詢系統』查詢-->
			<!-- </font> -->
			</strong></font>
			<br>
			<%if sys_City="基隆市" or sys_City="台南市" then%>
			<a href="Edgeset.asp" target="_blank"  style="font-size: 20pt;line-height:26px;">** Edge瀏覽器相容性設定方式 **</a>
			<%End If%>
			<%If ((Trim(Session("Credit_ID"))="N121946889" Or Trim(Session("Credit_ID"))="P120936942" Or Trim(Session("Credit_ID"))="E121544459" Or Trim(Session("Credit_ID"))="M121135787" Or Trim(Session("Credit_ID"))="Z016") And sys_City="台中市") Or Session("Credit_ID")="A000000000" Or (sys_City="新竹市" And Trim(Session("Credit_ID"))="G121048936") Or (sys_City="南投縣" And Trim(session("Unit_ID"))="05A7") or sys_City="彰化縣" then%>
			<a href="BillKeyIn/Update_Report_IllegalTime.asp" target="_blank" class="style2">** 強制修改違規時間 ** </a>	
			<%End If %>
			<%If sys_City="南投縣" And (Trim(session("Credit_ID"))="L224336134" Or Trim(session("Credit_ID"))="M121815995" Or Trim(session("Credit_ID"))="5PUF3SS5" Or Trim(session("Credit_ID"))="W37YHGEH" Or Trim(session("Credit_ID"))="M122213628" Or Trim(session("Credit_ID"))="M122439300" Or Trim(session("Credit_ID"))="M122033148" Or Trim(session("Credit_ID"))="A000000000") then%>
				<input type="button" value="建檔同車號案件檢查" onclick="window.open('BillKeyIn/setCheckCarRule.asp','winother1','width=500,height=350,left=100,top=50,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no')" style="font-size: 9pt; width: 160px; height: 27px">
			<%End if%>
			<%If sys_City="新竹市" And (Trim(session("Credit_ID"))="G121048936" Or Trim(session("Credit_ID"))="A000000000") then%>
				<input type="button" value="建檔同車號案件檢查" onclick="window.open('BillKeyIn/setCheckCarRule.asp','winother1','width=500,height=350,left=100,top=50,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no')" style="font-size: 9pt; width: 160px; height: 27px">
			<%End if%>
			<%if (Session("Credit_ID")<>"0000" And Trim(session("Group_ID"))<>"301012" AND sys_City="花蓮縣") Or sys_City="台東縣" then%>
			<br>
			催繳單號
			<input type="text" class="btn1" size="16" maxlength="16" value="" name="StopBillNo" onkeyup="EnterBillQry_Stop();">
			催繳車號
			<input name="StopCarNo" type="text" value="" size="8" maxlength="9" class="btn1" onkeyup="EnterBillQry_Stop();">
			<input type="button" name="btStopBill" value="催繳單查詢" onclick="Selt_Stop();">
			<%end if%>
			
			<%If sys_City="苗栗縣" then%>
				<input type="button" value="批號處理進度" onclick="Selt_BatchNumber();">
			<%End If %>
			<%If sys_City="屏東縣" then%>
			<a href="WebSet1030401.doc" target="_blank" class="style2">建檔系統不會自動帶應到案日期處理方式</a>
			<%End If %>
			<%If sys_City="金門縣" then%>
			<a href="Use1061110.pdf" target="_blank" style="font-size: 12pt;">交通執法系統操作手冊(如要下載，請按右鍵另存目標)</a>	
			<%End If %> 
			<%if sys_City="高雄市" or sys_City<>"新竹市" then%> 
			<Br><strong><font color="#336600">拖吊已結案件查詢</font> </strong>
			<strong>單號</strong>&nbsp;<input type="text" name="TakeCarBillNo" size="12" maxlength="9" onkeyup="value=value.toUpperCase();">
			<strong>車號</strong>&nbsp;<input type="text" name="TakeCarCarNo" size="12" maxlength="9" onkeyup="value=value.toUpperCase();">
			<strong>查詢事由</strong>
			<select name="QryReason2" >
				<option value="" >請選擇</option>
				<option value="資料檢核" <%If Trim(request("QryReason2"))="資料檢核" Then response.write "selected" End if%>>資料檢核</option>
				<option value="執行業務" <%If Trim(request("QryReason2"))="執行業務" Then response.write "selected" End if%>>執行業務</option>
				<option value="民眾申訴" <%If Trim(request("QryReason2"))="民眾申訴" Then response.write "selected" End if%>>民眾申訴</option>
				<option value="事故處理" <%If Trim(request("QryReason2"))="事故處理" Then response.write "selected" End if%>>事故處理</option>
				<option value="偵查刑案" <%If Trim(request("QryReason2"))="偵查刑案" Then response.write "selected" End if%>>偵查刑案</option>
			<%If sys_City<>"台中市" and sys_City<>"屏東縣" and sys_City<>"嘉義縣" and sys_City<>"新竹市" then%>
				<option value="偵查刑案" <%If Trim(request("QryReason"))="偵查刑案" Then response.write "selected" End if%>>偵查刑案</option>
			<%End If %>
			</select>
			<input type="button" value="拖吊已結案件查詢" onclick='getTakeCarBillData()'>
			<%End If %>
			<%if sys_City="台中市" then%>
				<input type="button" value="7天內刪除案件" onclick="window.open('WeekDeleteCase.asp','winMap181-1','width=900,height=550,left=0,top=0,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no')" style="font-size: 9pt; width: 100px; height: 27px">
			<%end if%>
				
			<%if Session("Credit_ID")="A000000000" then%>
							<font color="red">v$flash_recovery_area_usage</font>
							<%
								strDB="select PERCENT_SPACE_USED from v$flash_recovery_area_usage  where File_Type='ARCHIVELOG'"
								set rsDB=conn.execute(strDB)
								if not rsDB.eof Then
									if cint(rsDB("PERCENT_SPACE_USED"))>80 Then
										response.write "<font color='red' style='line-height:48px;font-size:40pt;'>"&cint(rsDB("PERCENT_SPACE_USED"))&"%</font>"
									Else
										response.write "<font color='red' style='line-height:28px;font-size:20pt;'>"&cint(rsDB("PERCENT_SPACE_USED"))&"%</font>"
									End if
									
									if cint(rsDB("PERCENT_SPACE_USED"))>80 and Session("Credit_ID")="A000000000" then
							%>
									<script language="JavaScript">
										alert("Oracle 暫存區快滿了，請趕快清一清");
									</script>
							<%
									end if
								end if
								rsDB.close
								set rsDB=Nothing
								'response.write "<font color='red' style='font-size:40pt;'>ooo</font>"
			  			%>			
							&nbsp; &nbsp; <font color="red"> v$session</font>
							<%
								strDB="Select count(*) as cnt from v$session"
								set rsDB=conn.execute(strDB)
								if not rsDB.eof Then
										response.write "<font color='red' style='line-height:28px;font-size:20pt;'>"&cint(rsDB("cnt"))&"</font>"

								end if
								rsDB.close
								set rsDB=Nothing
								'response.write "<font color='red' style='font-size:40pt;'>ooo</font>"
			  			%>		
							<!--<input type="button" value="資料庫Session" onclick="window.open('trafficDBCheck.asp','winother','width=900,height=550,left=0,top=0,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no')" style="font-size: 9pt; width: 100px; height: 27px">-->
			<%if sys_City="xxx" then%>
				<input type="button" value="事故系統佔用資源" onclick="window.open('tjdbCPU.asp','winother','width=400,height=350,left=0,top=0,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=yes,toolbar=no')" style="font-size: 9pt; width: 100px; height: 27px">
			<%end If%>
			<%end if%>
			<%if sys_City="台中市" And Trim(Session("Group_ID"))="200" then %>
			<a href="trafficdata.html" target="_blank" style="line-height:28px;font-size:18pt;">違規簡訊服務資訊</a>	
			<%end if%>
		</td>
	</tr>
<%if sys_City="台南市x" And Trim(Session("Group_ID"))="200" then%>
	<tr>
		<td colspan="5">
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr>
			    <td bgcolor="#FFCC33">
				<font size="4"><strong>系統警示</strong>
				</font>	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			    </td>
			</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="5">
		<table width="100%" align="left" border="0">
			<tr>
				<td width="40%">
				十天內建檔後超過四天內未入案之案件共 <strong><%
			LimitDate=DateAdd("d",-4,date)
			TenDate=DateAdd("d",-10,date)
			strA="select count(*) as cnt from billbase where billstatus in ('0','1') and recordstateid=0 and recorddate between to_date('"&TenDate&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and to_date('"&LimitDate&" 23:59:59','YYYY/MM/DD/HH24/MI/SS') and SN not in (select BillSN from ALERTCHECK where TypeID='1')"
			Set rsA=conn.execute(strA)
			If Not rsA.eof Then
				response.write rsA("cnt")
			End If
			rsA.close
			Set rsA=Nothing 
				%></strong> 筆
				</td>
				<td width="60%">
				<input type="button" value="檢視" onclick='window.open("Check_NotCaseIn.asp","Check_NotCaseIn","left=100,top=50,location=0,width=860,height=580,resizable=yes,scrollbars=yes,status=yes")'>
				</td>
			</tr>
			<tr>
				<td>
				十天內建檔後僅做車籍查詢就刪除之案件共 <strong><%
			strA="select count(*) as cnt from billbase where billno is null and sn in (select billsn from dcilog where exchangetypeid='A') and recordstateid=-1 and Recorddate > to_date('"&TenDate&" 0:0:0','YYYY/MM/DD/HH24/MI/SS') and SN not in (select BillSN from ALERTCHECK where TypeID='2')"
			Set rsA=conn.execute(strA)
			If Not rsA.eof Then
				response.write rsA("cnt")
			End If
			rsA.close
			Set rsA=Nothing 
				%></strong> 筆 
				</td>
				<td>
				<input type="button" value="檢視" onclick='window.open("Check_QryCar.asp","Check_QryCar","left=100,top=50,location=0,width=860,height=580,resizable=yes,scrollbars=yes,status=yes")'>
				</td>
			</tr>
			<tr>
				<td>
				十天內密碼輸入錯誤三次後系統鎖定之帳號共 <strong><%
			strA="select count(*) as cnt from Memberdata where AccountStateID=-1 and DelMemberID=99999 and LeaveJOBDate > to_date('"&TenDate&" 0:0:0','YYYY/MM/DD/HH24/MI/SS')"
			Set rsA=conn.execute(strA)
			If Not rsA.eof Then
				response.write rsA("cnt")
			End If
			rsA.close
			Set rsA=Nothing 
				%></strong> 筆 
				</td>
				<td>
				<input type="button" value="檢視" onclick='window.open("Check_UserLock.asp","Check_UserLock","left=100,top=50,location=0,width=860,height=580,resizable=yes,scrollbars=yes,status=yes")'>
				</td>
			</tr>
		</table>
		</td>
	</tr>
<%End if%>
<%
if Session("Group_ID") = "8984" or Session("Group_ID") = "8985" or Session("Group_ID") = "8987" then
	strFuncGroup="select * from Code where TypeID=19 and ID!=510 and Id!=512 and ID!=515 and ID!= 514 and id!=511 and id!=509 order by ShowOrder"
else
	strFuncGroup="select * from Code where TypeID=19 order by ShowOrder"
end if

set rsFuncGroup=conn.execute(strFuncGroup)
If Not rsFuncGroup.Bof Then rsFuncGroup.MoveFirst 
While Not rsFuncGroup.Eof
%>
	<tr><td width="20%"></td><td width="20%"></td><td width="20%"></td><td width="20%"></td><td width="20%"></td><tr>
	<tr>
		<td colspan="5">
		<table width='100%' border='1' align="center" cellpadding="5" cellspacing="0">
			<tr>

			    <td bgcolor="#EBF5FF">



				<font size="4"><strong><%=trim(rsFuncGroup("Content"))%></strong>
				</font>	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<% if sys_city="高雄市" or sys_City="苗栗縣" then
						ShowTimeGroupID=516
					else
						ShowTimeGroupID=512
					end if
					if Trim(rsFuncGroup("ID"))=Trim(ShowTimeGroupID) then
				%>
					
				<%  end if
				
					if rsFuncGroup("ID")="509" then
				%>
					<br>
					<font size="5">
					<strong>
					<img src="Image/dot.gif"></img>
					<a href="send.doc" target="_blank" >民眾簽收 簡要說明.doc</a>
					<img src="Image/dot.gif"></img>
					<a href="opengov.doc" target="_blank" >公示送達 簡要說明.doc</a>
					<img src="Image/dot.gif"></img><!--
					<a href="storeandsend.doc" target="_blank" >寄存送達 簡要說明.doc</a>-->
					</strong>
					
					</font>		
				<%  end if				
					if (rsFuncGroup("ID")="515" or rsFuncGroup("ID")="516" ) and sys_city="高雄市" then
				%>
					<br>
					<font size="5">
					<strong>
					<img src="Image/dot.gif"></img>
					<a href="CaseProcedure.asp" target="_blank" >案件舉發流程</a>
					</strong>
					
					</font>		
				<%  end if		
				%>
			    </td>
			</tr>
		</table>
		</td>
	</tr>
<%
	s=1
	strFunc="select a.* from FunctionPageData a,FunctionData b where a.SystemID=b.SystemID and b.GroupID='"&trim(GroupID)&"' and a.SystemGroupID="&trim(rsFuncGroup("ID"))&" and b.Function='1' order by ShowOrder"
	set rsFunc=conn.execute(strFunc)
	If Not rsFunc.Bof Then rsFunc.MoveFirst 
	While Not rsFunc.Eof

	if s=1 then
		response.write "<tr>"
	end if
%>	<td width="190" height="180">
<div id="<%=rsFunc("SystemID")%>" style="width:155px; height:170px; z-index:1 ;">  
		<table id="<%="table"&rsFunc("SystemID")%>" width='100%' border='0' align="center" >
		<tr><td id="td1" align="center">
		<a onclick="OpenSystem('<%=rsFunc("URLLocation")%>','<%=rsFunc("SystemID")%>');" onMouseOver="DivColorChange('<%="table"&rsFunc("SystemID")%>');" onMouseOut="DivColorChange2('<%="table"&rsFunc("SystemID")%>');">
		<%
		if trim(rsFunc("ImageLocation"))="" or isnull(rsFunc("ImageLocation")) then
			picName="tmp.jpg"
		else
			picName=rsFunc("ImageLocation")
		end if
		%>
		<img src="image/<%=picName%>" alt="" width="128" height="128" border="0" align="baseline">
		  <br><font size="4"><%
		  strCode="select * from Code where ID="&trim(rsFunc("SystemID"))
		  set rsCode=conn.execute(strCode)
		  if not rsCode.eof then
			response.write trim(rsCode("Content"))		
		  end if
		  rsCode.close
		  set rsCode=nothing
		  %></font></a>
		</td></tr>
		</table>
</div>
	</td>
<%	
	if s=5 then
		response.write "</tr>"
		s=1
	else
		s=s+1
	end if

	rsFunc.MoveNext
	Wend

rsFuncGroup.MoveNext
Wend
rsFuncGroup.close
set rsFuncGroup=nothing
%>
  <%
	rsFunc.close
	set rsFunc=nothing
	strSQL="select TO_Char(sysdate,'YYYY/MM/DD HH24:MI:SS') tmpDate from dual"
	set rstime=conn.execute(strSQL)
	SysDate=rstime("tmpDate")
	rstime.close

%>
</table>

</body>
<script type="text/javascript" src="./js/date.js"></script>
<script language="JavaScript">
<%
'if sys_City="金門縣" Then
	Modifydate=""
	PassWordTemp=""
	showUpdateMemberdateFlag=0
	strMem="select * from memberdata where memberid=" & Trim(session("User_ID")) & " and recordstateid=0 and accountstateid=0"
	Set rsMem=conn.execute(strMem)
	If Not rsMem.eof Then
		If Trim(rsMem("ModifyTime"))="" Then
			Modifydate=Trim(rsMem("ModifyTime"))
		Else
			Modifydate=Trim(rsMem("RecordDate"))
		End If 
		
		If sys_City="澎湖縣" Or sys_City="基隆市" Or sys_City="高雄市" Or sys_City="金門縣" Or sys_City="台東縣" Or sys_City="彰化縣" Or sys_City="台中市" Or sys_City="屏東縣" Or sys_City="嘉義縣" Or sys_City="雲林縣" Or sys_City="嘉義市" Or sys_City="新竹市" Then
			PassWordTemp=encrypt(Trim(rsMem("PassWord")))
		Else
			PassWordTemp=Trim(rsMem("PassWord"))
		End If 
	End If 
	rsMem.close
	Set rsMem=Nothing 
	If DateDiff("d",Modifydate,now)>90 Then
		response.write "alert('您的密碼已超過三個月( " & DateDiff("d",Modifydate,now) & " 天)未更換，請儘速修改您的密碼!!');"
		showUpdateMemberdateFlag=1
	elseif Len(PassWordTemp)<8 then
		response.write "alert('密碼長度至少為<8>碼，請儘速修改您的密碼!!');"
		showUpdateMemberdateFlag=1
	else
		chkUp=0
		chkDown=0
		chkInt=0
		chkMark=0
		for i=1 to Len(PassWordTemp)
			if Asc(Mid(Trim(PassWordTemp), i, 1))>=65 and Asc(Mid(Trim(PassWordTemp), i, 1))<=90 then
				chkUp=1
			end if
			if Asc(Mid(Trim(PassWordTemp), i, 1))>=97 and Asc(Mid(Trim(PassWordTemp), i, 1))<=122 then
				chkDown=1
			end if 
			if Asc(Mid(Trim(PassWordTemp), i, 1))>=48 and Asc(Mid(Trim(PassWordTemp), i, 1))<=57 then
				chkInt=1
			end if 
			if (Asc(Mid(Trim(PassWordTemp), i, 1))>=33 and Asc(Mid(Trim(PassWordTemp), i, 1))<=47) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=58 and Asc(Mid(Trim(PassWordTemp), i, 1))<=64) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=91 and Asc(Mid(Trim(PassWordTemp), i, 1))<=96) or (Asc(Mid(Trim(PassWordTemp), i, 1))>=123 and Asc(Mid(Trim(PassWordTemp), i, 1))<=126) then
				chkMark=1
			end if
		next
		if chkUp=0 or chkDown=0 or chkInt=0 or chkMark=0 then
			response.write "alert('密碼需包含英文、數字、特殊符號及大小寫混和，請儘速修改您的密碼!!');"
			showUpdateMemberdateFlag=1
		end if
	End If 
'End If 

if sys_City="台中市" then
	StrChkNoPass="select count(*) as cnt from billbaseTmp where Recordstateid=0 " &_
		" and BillStatus='7' and checkflag='2' and recordmemberid="&trim(Session("User_ID"))
	Set rsCNP=conn.execute(StrChkNoPass)
	If Not rsCNP.eof Then
		if cdbl(rsCNP("cnt"))>0 then
%>
		alert("您有 <%=rsCNP("cnt")%> 筆，個人建檔審核未通過案件，請至影像建檔審核系統確認!");
<%
		end if 
	End If 
	rsCNP.close
	Set rsCNP=Nothing 
end if


conn.close
set conn=nothing
%>

function EnterBillQry(){
	document.all.BillNo.value=document.all.BillNo.value.toUpperCase();
	document.all.CarNo.value=document.all.CarNo.value.toUpperCase();
	document.all.IllegalID.value=document.all.IllegalID.value.toUpperCase();
<%if sys_City="台中市" then%>
	document.all.ReportNo.value=document.all.ReportNo.value.toUpperCase();
<%end if%>
	if (event.keyCode==13){
		getBillData();
	}
}
<%if sys_City="高雄市" or sys_City="新竹市" then%>
function getTakeCarBillData(){
	if (document.all.TakeCarBillNo.value.length < 9 && document.all.TakeCarBillNo.value!=""){
		alert("舉發單號小於九碼！");
	
	}else if (document.all.TakeCarBillNo.value=="" && document.all.TakeCarCarNo.value==""){
		alert("必須填入拖吊單號或拖吊車號！");
	}else if (document.all.QryReason2.value==""){
		alert("因資安審查規定，查詢必須選擇查詢事由！");
	}else{
		UrlStr="../traffic/Query/BillBaseData_Detail_TakeCar_KSC.asp?BillNo="+document.all.TakeCarBillNo.value+"&CarNo="+document.all.TakeCarCarNo.value+"&QryReason2="+document.all.QryReason2.value;
		newWin(UrlStr,"winMap",800,550,50,10,"yes","yes","yes","no");
	}
}
<%end if%>

function getBillData(){
<%If sys_City="台東縣" And Trim(session("Unit_ID"))="Z000" Then%>	

<%else%>
	if (document.all.BillNo.value.length < 9 && document.all.BillNo.value!=""){
		alert("舉發單號小於九碼！");
	<%if sys_City="台東縣" then%>
	}else if (document.all.BillNo.value==""){
		alert("因資安審查規定，查詢必須輸入單號！");
	<%end if %>
	<%if sys_City="台中市" then%>
	}else if (document.all.CarNo.value=="" && document.all.BillNo.value=="" && document.all.IllegalName.value=="" && document.all.IllegalID.value=="" && document.all.ReportNo.value==""){
		alert("必須填入單號或車號或違規人或告示單號！");
	<%else%>
	}else if (document.all.CarNo.value=="" && document.all.BillNo.value=="" && document.all.IllegalName.value=="" && document.all.IllegalID.value=="" ){
		alert("必須填入單號或車號或違規人！");
	<%end if %>	
	}else if (document.all.QryReason.value==""){
		alert("因資安審查規定，查詢必須選擇查詢事由！");
		
	}else{
	<%if sys_City="台中市" then%>
		UrlStr="../traffic/Query/BillBaseData_Detail_Main.asp?BillNo="+document.all.BillNo.value+"&CarNo="+document.all.CarNo.value+"&IllegalName="+document.all.IllegalName.value+"&IllegalID="+document.all.IllegalID.value+"&ReportNo="+document.all.ReportNo.value+"&QryReason="+document.all.QryReason.value;
		newWin(UrlStr,"winMap",800,550,50,10,"yes","yes","yes","no");
	<%elseif sys_City="高雄市" or sys_City="屏東縣" or sys_City="台東縣" then%>
		UrlStr="/traffic/Query/BillBaseData_Detail_Main.asp?BillNo="+document.all.BillNo.value+"&CarNo="+document.all.CarNo.value+"&IllegalName="+document.all.IllegalName.value+"&IllegalID="+document.all.IllegalID.value+"&QryReason="+document.all.QryReason.value+"&ForChkMen="+document.all.ForChkMen.value;
		newWin(UrlStr,"winMap",800,550,50,10,"yes","yes","yes","no");
	<%else%>
		UrlStr="/traffic/Query/BillBaseData_Detail_Main.asp?BillNo="+document.all.BillNo.value+"&CarNo="+document.all.CarNo.value+"&IllegalName="+document.all.IllegalName.value+"&IllegalID="+document.all.IllegalID.value+"&QryReason="+document.all.QryReason.value;
		newWin(UrlStr,"winMap",800,550,50,10,"yes","yes","yes","no");
	<%end if %>
	}
<%end if%>
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
	win.focus();
}

function OpenSystem(PageUrl,Sn){
	SCheight=screen.availHeight;
	SCWidth=screen.availWidth;
	UrlStr=PageUrl;
	newWin(UrlStr,Sn,SCWidth,SCheight,0,0,"yes","no","yes","no");
}
function DivColorChange(DivNo){
	eval(DivNo).border="1";
}
function DivColorChange2(DivNo){
	eval(DivNo).border="0";
}
function funMember(){
	UrlStr="UserDataEdit.asp";
	newWin(UrlStr,"winMapfunMember",800,450,50,10,"yes","no","yes","no");
}
function change_Time(){
//	var time=new Date();
//	t_Hour=time.getHours();
//	t_Minute=time.getMinutes();
//	t_Second=time.getSeconds();
	//runServerScript("getServerTime.asp");

	//LayerTime.innerHTML="目前時間  "+t_Hour+"："+t_Minute;
	//setTimeout(change_Time,60000);
}
function funOpenWindow(){
	//跳出視窗的網址
	UrlStr="NOTICEMain.asp";
	newWin(UrlStr,"winMap111",600,450,50,10,"yes","no","yes","no");
<%'sys_City="台中市" or 
if (sys_City="基隆市" and trim(Session("Unit_ID"))="0207") or (sys_City="苗栗縣" and (trim(Session("Credit_ID"))="A000000000" or trim(Session("Credit_ID"))="JENIFER" or trim(Session("Credit_ID"))="TIFFANY" or trim(Session("Credit_ID"))="YESLYN")) or (sys_City="彰化縣" and trim(Session("Group_ID"))=199) then%>
	newWin("WeekDeleteCase.asp","winMap181",760,600,80,10,"yes","no","yes","no");
<%end if%>

<%if sys_City="基隆市" then%>
	newWin("chkNotMaildate.asp","winMap182",760,600,80,10,"yes","no","yes","no");
<%end if%>

<%if sys_City="台中市" then%>
	<%if trim(Session("Group_ID"))="200" or trim(Session("Group_ID"))="201" or trim(Session("Group_ID"))="202" or trim(Session("Group_ID"))="9334" then%>
	//newWin("GetNotBillNo.asp","GetNotBillNo",860,600,110,10,"yes","no","yes","no");
	<%end if %>
<%end if %>
<%
	if showUpdateMemberdateFlag=1 then
		response.write "funMember();"
	end if 
%>
}
//登入就跳視窗
<%if sys_City<>"高雄市" then%>
//funOpenWindow();
<%end if%>

function menuIn() //隱藏
{
        if(n4) {
                clearTimeout(out_ID)
                if( menu.left > menuH*-1+30+10 ) {  
                        menu.left -= 14
                        in_ID = setTimeout("menuIn()", 1)
                }
                else if( menu.left > menuH*-1+30 ) {
                        menu.left--
                        in_ID = setTimeout("menuIn()", 1)
                }
        }
        else { 
                clearTimeout(out_ID)
                if( menu.pixelLeft > menuH*-1+30+10 ) {
                        menu.pixelLeft -= 14
                        in_ID = setTimeout("menuIn()", 1) 
                }
                else if( menu.pixelLeft > menuH*-1+30 ) {
                        menu.pixelLeft--
                        in_ID = setTimeout("menuIn()", 1)
                }
        }
}
function menuOut() //顯示
{
        if(n4) {
                clearTimeout(in_ID)
                if( menu.left < -10) { 
                        menu.left += 4
                        out_ID = setTimeout("menuOut()", 1)
                }
                else if( menu.left < 0) { 
                        menu.left++
                        out_ID = setTimeout("menuOut()", 1)
                }
                
        }
        else { 
                clearTimeout(in_ID)
                if( menu.pixelLeft < -10) {
                        menu.pixelLeft += 2
                        out_ID = setTimeout("menuOut()", 1)
                }
                else if( menu.pixelLeft < 0 ) {
                        menu.pixelLeft++
                        out_ID = setTimeout("menuOut()", 1)
                }
        }
}
function fireOver() { 
        clearTimeout(F_out)	       
        F_over = setTimeout("menuOut()", 10) 
}
function fireOut() { 
        clearTimeout(F_over)
         F_out = setTimeout("menuIn()", 10)
}
function init() {
        if(n4) {
                menu = document.D1
                menuH = menu.document.width
                menu.left = menu.document.width*-1+30 
                menu.onmouseover = menuOut
                menu.onmouseout = menuIn
				menu.visibility = "visible"
        }
        else if(e4) {
                menu = D1.style
                menuH = D1.offsetWidth
                //D1.style.pixelLeft = D1.offsetWidth*-1+20
                D1.onmouseover = fireOver
                D1.onmouseout = fireOut
				D1.onclick = fireOut
                D1.style.visibility = "visible"
        }
		UpdateLayer();
}
function UpdateLayer(){
	//UpLoadLayer.innerHTML="";
	runServerScript("UpdateMainLayer.asp");
	setTimeout(UpdateLayer,1200000);
	//alert("1");
}
F_over=F_out=in_ID=out_ID=null
n4 = 0;
e4 = 1;
var procesID='<%=Session("Credit_ID")%>'

function DownProcess(nowtime){
	<%if sys_City="花蓮縣" or sys_City="苗栗縣" or sys_City="台東縣" or sys_City="雲林縣" or sys_City="台中市" or sys_City="高雄市" then%>
		//runServerScript("/traffic/BillReturn/SystemDownloadFile.asp?nowTime="+nowtime);
	<%end if%>
}

function NowDownProcess(){
	<%if sys_City="花蓮縣" or sys_City="苗栗縣" or sys_City="台東縣" or sys_City="雲林縣" or sys_City="台中市" or sys_City="高雄市" then%>
		//var nowtime="<%=year(SysDate)&"/"&Month(SysDate)&"/"&Day(SysDate)&" "&hour(SysDate)&"："&minute(SysDate)&"："&second(SysDate)%>";
		//runServerScript("/traffic/BillReturn/SystemDownloadFile.asp");
		//alert("已重新處理");
	<%end if%>
}

function SQLDownProcess(nowtime){
	<%if sys_City="台中縣" or sys_City="雲林縣" or sys_City="屏東縣"then%>
//		if(procesID=='A000000000'){
			//alert("/traffic/BillReturn/T-SQL.asp?nowTime="+nowtime);
			runServerScript("/traffic/BillReturn/T-SQL.asp?nowTime="+nowtime);
//		}
	<%end if%>
}

function ExportDB(){
	//runServerScript("OracleReturn.asp");
	newWin("/traffic/BillReturn/OracleReturn.aspx","winMap181",760,600,50,10,"yes","no","yes","no");
}

function funcUpdate(){
	newWin("/traffic/Update.asp","winMapUpdate",760,600,50,10,"yes","no","yes","no");
}
<%if (Session("Credit_ID")<>"0000" AND sys_City="花蓮縣") Or sys_City="台東縣" then%>
function EnterBillQry_Stop(){
	StopBillNo.value=StopBillNo.value.toUpperCase();
	StopCarNo.value=StopCarNo.value.toUpperCase();
	if (event.keyCode==13){
		Selt_Stop();
	}
}

function Selt_Stop(){
	if (StopCarNo.value=="" && StopBillNo.value==""){
		alert("必須填入單號或車號或違規人！");
	}else{
		window.open("../traffic/Query/<%
		if (sys_City="台東縣") then
			response.write "StopBillBaseData_Detail_TaiDung.asp"
		else
			response.write "StopBillBaseData_Detail.asp"
		end if 
		%>?BillNo="+StopBillNo.value+"&CarNo="+StopCarNo.value,"WebPage2","left=0,top=0,location=0,width=980,height=555,resizable=yes,scrollbars=yes,menubar=yes,status=yes");
	}
	
}
<%end if%>
<%if sys_City="苗栗縣" then%>
function Selt_BatchNumber(){
	window.open("../traffic/Query/BatchNumberQry.asp","WebPagedx2","left=0,top=0,location=0,width=980,height=555,resizable=yes,scrollbars=yes,menubar=yes,status=yes");
}
<%end if %>
var mydate = new Date("<%=year(SysDate)&"/"&Month(SysDate)&"/"&Day(SysDate)&" "&hour(SysDate)&":"&minute(SysDate)&":"&second(SysDate)%>");
var mySec=<%=second(SysDate)%>;

function Selt_Time(){
	var nowtime='';
	mySec=mySec+1;
	mydate.setSeconds(mySec);

	nowtime = ("0"+mydate.getHours()).substr(("0"+mydate.getHours()).length-2,2);

	nowtime = nowtime + " : " + ("0"+mydate.getMinutes()).substr(("0"+mydate.getMinutes()).length-2,2);
	nowtime = nowtime + " : " + ("0"+mydate.getSeconds()).substr(("0"+mydate.getSeconds()).length-2,2);
	LayerTime.innerText = "系統時間 "+nowtime;
	//if (mydate.getMinutes()%1==0&&mydate.getSeconds()==1){DownProcess(nowtime);}
	if (mySec==60){mySec=0;}
}
var oInterval = setInterval("Selt_Time()", 1000);

<%
if sys_City<>"高雄市" and sys_City<>"台中市" and sys_City<>"苗栗縣" and sys_City<>"南投縣" then
%>
//newWin("SystemBulletin.asp","",860,600,50,10,"yes","no","yes","no");
<%end if %>
</script>

</html>
