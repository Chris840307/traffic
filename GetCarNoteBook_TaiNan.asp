<%
'fname=year(now)&fMnoth&fDay&"_交寄大宗.doc"
'Response.AddHeader "Content-Disposition", "filename="&fname
'response.contenttype="application/msword; charset=MS950" 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>領車通知單</title>
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style2 {font-family: "標楷體"; font-size: 16px;line-height:20px}
.style3 {font-family: "標楷體"; font-size: 14px;line-height:20px}
.style4 {font-family: "標楷體"; font-size: 20px;line-height:20px}
.style5 {font-family: "標楷體"; font-size: 12px;line-height:20px}
-->
</style>
</head>
<!--#include virtual="traffic/takecar/Common/DB.ini"-->
<!--#include virtual="/traffic/Takecar/Common/AllFunction.inc"-->
<body>
<%
		strBill="select tel,Address from traffic.UnitInfo where UnitID='"&session("Unit_ID")&"'"
		set rsBill=conn.execute(strBill)
				if Not rsBill.eof then
					tel           = trim(rsBill("tel"))
					Address       = trim(rsBill("Address"))
				End if
		Set rsBill=Nothing

    ReceiveSN=split(request("sn"),",")
    for i=0 to ubound(ReceiveSN)
	if cint(i)>0 then 
		response.write "<div class=""PageNext"">&nbsp;</div>"
               
        end if

	carno="" : illegalDateTime="" : DealLineDate="" : owner="" : rule1="" : rule2="" : rule3=""

	If request("times")="1" Then 

		strBill="select a.NowKeepUnitID,a.rulecode,a.UnitName,a.UnitID,a.TableTypeID,a.InCarTypeID,a.ownerName,a.carno,a.Billno,a.ownerAddr,a.indatetime,a.CarTypeID from takebase a where  a.sn="&ReceiveSN(i) &" order by sn"
		set rsBill=conn.execute(strBill)
				if Not rsBill.eof Then
					UnitID          = trim(rsBill("UnitID"))				
					TableTypeID     = trim(rsBill("TableTypeID"))				
					CarTypeName     = GetCarTypeToName(trim(rsBill("CarTypeID")))
					Billno          = trim(rsBill("Billno"))				
					CarNo           = trim(rsBill("CarNo"))
					InCarTypeID     = trim(rsBill("InCarTypeID"))
					indatetime      = trim(rsBill("indatetime"))
					ownerName       = trim(rsBill("ownerName"))
					ownerAddr       = trim(rsBill("ownerAddr"))
					rulecode       = trim(rsBill("rulecode"))
					nowKeepUnitID       = trim(rsBill("nowKeepUnitID"))
				End if
		Set rsBill=Nothing

'		strBill="select nwner from takeCarDciReturn where carno='"&carno&"'"
'		set rsBill=conn.execute(strBill)
'				if Not rsBill.eof then
'					ownerName=trim(rsBill("nwner"))
'				End if
'		Set rsBill=Nothing
	Else
		strBill="select a.nowKeepUnitID,a.rulecode,a.UnitID,a.TableTypeID,a.InCarTypeID,a.ownerName,a.carno,a.Billno,a.DriverAddr,a.indatetime,a.CarTypeID from takebase a where  a.sn="&ReceiveSN(i) &" order by sn"
		set rsBill=conn.execute(strBill)
				if Not rsBill.eof Then
					UnitID          = trim(rsBill("UnitID"))				
					TableTypeID     = trim(rsBill("TableTypeID"))				
					CarTypeName     = GetCarTypeToName(trim(rsBill("CarTypeID")))
					Billno          = trim(rsBill("Billno"))				
					CarNo           = trim(rsBill("CarNo"))
					InCarTypeID     = trim(rsBill("InCarTypeID"))
					indatetime      = trim(rsBill("indatetime"))
					ownerName       = trim(rsBill("ownerName"))
					ownerAddr       = trim(rsBill("DriverAddr"))
					rulecode       = trim(rsBill("rulecode"))
					nowKeepUnitID       = trim(rsBill("nowKeepUnitID"))
				End if
		Set rsBill=Nothing

		'strBill="select nwner from takeCarDciReturn where carno='"&carno&"'"
		'set rsBill=conn.execute(strBill)
		'		if Not rsBill.eof then
		'			ownerName=trim(rsBill("nwner"))
		'		End if
		'Set rsBill=Nothing

	End if

%>
<div style="position:relative;">
<!--車主姓名回執聯中-->
<div style="position:absolute; left:320px; top:190px;height:260px;writing-mode:tb-rl;text-align=left" class="style2">
<%=ownername%>
</div>

<!--車主地址回執聯中-->
<div style="position:absolute; left:250px; top:180px;height:200px;writing-mode:tb-rl;text-align=left" class="style2">
<%=owneraddr%>
</div>


<!--車主姓名回執聯右-->
<div style="position:absolute; left:570px; top:220px;height:260px;writing-mode:tb-rl;text-align=left " class="style4">
<%=ownername%>
</div>

<!--車主地址回執聯中-->
<div style="position:absolute; left:700px; top:40px;height:460px;writing-mode:tb-rl;text-align=left " class="style4">
<%=owneraddr%>
</div>


<!--投遞後後郵戳-->
<div style="position:absolute; left:20px; top:470px;" class="style3">
<%=GetCDateTime(indatetime) & GetSpace(10) & GetTakeCarUnitName(nowKeepUnitID)%>
</div>

<div style="position:absolute; left:20px; top:490px;" class="style3">
<%=Billno & GetSpace(2) & Carno & GetSpace(7) & GetUnitAddr(nowKeepUnitID)%>
</div>

<div style="position:absolute; left:100px; top:550px;" class="style4">
<b><%="領"& GetSpace(2) &"車"& GetSpace(2) &"通"& GetSpace(2) &"知"& GetSpace(2) &"單"%></b>
</div>

<%

tmpnum="三"
if nowKeepUnitID="07C1" then tmpnum="一"

%>
<div style="position:absolute; left:400px; top:910px;text-align=left" class="style3">
臺南市政府警察局交通警察大隊<br>
第<%=tmpnum%>中隊&nbsp;<%=GetTakeCarUnitName(nowKeepUnitID)%>&nbsp;&nbsp;<%=GetCDate(goutdt(request("NoteDate")))%>
</div>

<div style="position:absolute; left:80px; top:600px;text-align=left" class="style3">
<%if InCarTypeID="2" Then%>
    <%If RuleCode="21" Or RuleCode="23" Or RuleCode="24" then%>
			<%=GetSpace(4)%>台端所有<%=CarNo%>號 <%=CarTypeName%> 於 <%=Year(indatetime)-1911%> 年 <%=Right("0"&Month(indatetime),2)%> 月 <%=Right("0"&day(indatetime),2)%> 日 <%=Right("0"&hour(indatetime),2)%> 時 <%=Right("0"&minute(indatetime),2)%> 分<br>
			違反道路交通管理處罰條例經本局執勤人員依法製單舉發，並將該車代保管在案，迄今尚未到指定處理<br>
			機關辦理。請接到本通知單15日內持至取締單位辦理發還手續(請取締單位於第一聯「收據聯」受理員警<br>
			職名章及主管職名章處蓋章)，再持該取締單位審核蓋章之「臺南市政府警察局舉發違反交通管理事件車<br>
			輛移置保管單第一聯(收據聯)」及車輛行照或相關證件、駕駛執照或身分證至本局違規車輛保管場，辦<br>
			理領車手續，另號牌經註銷、吊銷禁止行駛之車輛請以載具領回切勿行駛道路，逾期未領者，當依法公<br>
			告拍賣。
    <%else%>
		<%=GetSpace(4)%>因「酒後駕車」經本局&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;分局(大隊)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;派出所(警備隊)製單舉發並代保管車輛在<br>
		案，車號：<%=Carno%><br>
		請持下列證件：<br>
		一、繳納酒後駕車罰鍰收據正本。<br>
		二、臺南市政府警察局舉發違反交通管理事件車輛移置保管單第一聯(收據聯)。<br>
		三、車輛行照或相關證件。<br>
		四、駕駛執照(代保管車輛由領有駕照之人駕駛)。<br>
		至取締單位辦理發還手續(請取締單位於第一聯「收據聯」受理員警職名章及主管職名章處蓋章)，再在<br>
		持該取締單位審核蓋章之「臺南市政府警察局舉發違反交通管理事件車輛移置保管單第一聯(收據<br>
		聯)」及證件至本局違規車輛保管場，辦理領車手續，逾期未領者，當依法公告拍賣。
	<%End if%>
<%else '拖吊%>
		<%=GetSpace(4)%>台端所有<%=CarNo%>號 <%=CarTypeName%> 於 <%=Year(indatetime)-1911%> 年 <%=Right("0"&Month(indatetime),2)%> 月 <%=Right("0"&day(indatetime),2)%> 日 <%=Right("0"&hour(indatetime),2)%> 時 <%=Right("0"&minute(indatetime),2)%> 分<br>
		於臺南市<%If rulecode="54" Then response.write "停車場管理自治條例" Else response.write "違規停車"%>
		，經本局依規定拖吊至 <%=GetTakeCarUnitName(Session("Unit_ID"))%> 保管，迄今已多日且尚未領回，請於函到15日內攜
		帶本通知書、行車執照或新領號牌申請書、駕駛執照（或身分證）等證件至該場繳費領車，逾期未領者，當依法公告拍賣。
<%End if%>



</div>
</div>
<%Next%>



</body>

</html>
<object id=factory style="display:none"
classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814"
codebase="../smsx.cab#Version=6,1,432,1">
</object>
<script type="text/javascript" src="../js/Print.js"></script>
<script language="javascript">
printWindow(true,5.08,3.5,5.08,0);
</script></p>
<%conn.close%>