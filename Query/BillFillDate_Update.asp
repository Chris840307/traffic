<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include file="sqlDCIExchangeData.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html onkeydown="KeyDown()">
<script language="JavaScript">
	window.focus();
</script>
<head>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<style type="text/css">
<!--
.style4 {
	font-size: 12px
}
.style5 {
	font-size: 20px;
	color: #FF0000;
}
-->
</style>
<title>逕舉監理所入案</title>
<% Server.ScriptTimeout = 3800 %>
<%
'tmpSQL=Session("BillSQLforReport")
tmpSQL=replace(trim(request("DciLogSQLforReport")),"@!@","%")
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing

if trim(request("kinds"))="Report_CaseIn" then
	strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
	set rsSN=conn.execute(strSN)
	if not rsSN.eof then
		theBatchTime=(year(now)-1911)&"W"&trim(rsSN("SN"))
	end if
	rsSN.close
	set rsSN=Nothing
	
	if sys_City="高雄市" Or sys_City=ApconfigureCityName or (trim(request("chkQueryCar"))="1" and sys_City="台中市") Then
		'車籍查尋的批號
		strSN2="select DCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN2=conn.execute(strSN2)
		if not rsSN2.eof then
			theBatchTimeQryCar=(year(now)-1911)&"A"&trim(rsSN2("SN"))
		end if
		rsSN2.close
		set rsSN2=nothing
	End If
	
	BillNoFirst=""	'第一筆單號
	BillNoLast=""	'最後一筆單號
	BillCount=0	'入案數
	isDoubleCase=0 '是否重覆入案
	if tmpSQL="" then
		strToDCI=""
	else
		strToDCI="select a.SN,a.IllegalDate,a.BillTypeID,a.BillNo,a.CarNo,a.BillUnitID,a.BillStatus,a.RecordDate,a.RecordMemberID from BillBase a,MemberData b where a.RecordStateID<>-1 and a.RecordMemberID=b.MemberID(+)"&tmpSQL&" order by a.RecordDate"
	end if
	'response.write strToDCI
	'response.end
	set rsToDCI=conn.execute(strToDCI)
	If Not rsToDCI.Bof Then
		rsToDCI.MoveFirst
	else
%>
<script language="JavaScript">
	alert("無可進行入案之舉發單！");
	window.close();
</script>
<%
	end if
	Do While Not rsToDCI.Eof
		BillCount=BillCount+1
		if sys_City="高雄市" Or sys_City=ApconfigureCityName or (trim(request("chkQueryCar"))="1" and sys_City="台中市") Then '車籍查尋
			funcCarDataCheck conn,trim(rsToDCI("SN")),"",trim(rsToDCI("BillTypeID")),trim(rsToDCI("CarNo")),trim(rsToDCI("BillUnitID")),trim(rsToDCI("RecordDate")),trim(rsToDCI("RecordMemberID")),theBatchTimeQryCar
		End If
		
		BillNoLast=funcBillToDCICaseIn(conn,trim(rsToDCI("SN")),trim(rsToDCI("BillNo")),trim(rsToDCI("BillTypeID")),trim(rsToDCI("CarNo")),trim(rsToDCI("BillUnitID")),trim(rsToDCI("RecordDate")),trim(rsToDCI("RecordMemberID")),theBatchTime,sys_City)

		if trim(request("FillFlag"))="Y" and trim(request("BillFillDate"))<>"" and trim(request("DealLineDate"))<>"" then
			strUpd="Update BillBase set BillFillDate="&funGetDate(gOutDT(trim(request("BillFillDate"))),1)&" ,DeallineDate="&funGetDate(gOutDT(trim(request("DealLineDate"))),1)&" where Sn="&trim(rsToDCI("SN"))
			conn.execute strUpd
		end If
				
		if BillCount=1 then
			BillNoFirst=BillNoLast
		End If 
		'if BillCount=1 Or (BillCount Mod 50 = 0) then
			strDblCase="select count(*) as cnt from BillBase where RecordstateID=0 and billno='"&BillNoLast&"'"
			Set rsDblCase=conn.execute(strDblCase)
			If Not rsDblCase.eof Then
				If CInt(rsDblCase("cnt"))>1 Then
					isDoubleCase=1
				End If 
			End If
			rsDblCase.close
			Set rsDblCase=Nothing 
			'response.write BillNoLast&" - "&BillCount&"<br>"
		'end If
		
		If isDoubleCase=1 Then
			Exit Do 
		End If 
	rsToDCI.MoveNext
	loop

	if trim(request("HelpPrint"))="1" then
		strInsP="Insert Into BillPrintJob(BatchNumber) " &_
			" values('"&trim(theBatchTime)&"')"
		conn.execute strInsP
	end if

	If Not rsToDCI.Bof Then
		If isDoubleCase=1 Then
%>
<script language="JavaScript">
		alert("入案失敗：舉發單號已重覆，請儘速聯絡工程師為您處理！");
		window.close();
</script>
<%		
		Else 
%>
<script language="JavaScript">
	opener.myForm.submit();
	window.open("BatchAlert.asp?BatchNo=<%=theBatchTime%>&FistNo=<%=BillNoFirst%>&LastNo=<%=BillNoLast%>&BillCount=<%=BillCount%>&BatchNoQryCar=<%=theBatchTimeQryCar%>","winAlert1","width=250,height=250,left=350,top=250,scrollbars=no,menubar=no,resizable=no,fullscreen=no,status=no,toolbar=no");

	//alert("入案處理完成，批號：<%=theBatchTime%> \n起始單號：<%=BillNoFirst%> \n結束單號：<%=BillNoLast%> \n共計：<%=BillCount%> 筆");
	window.close();
</script>
<%		End If 
	end if
	rsToDCI.close
	set rsToDCI=nothing
end if
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td>逕舉監理所入案
				<br>
				<span class="style4">
				<strong>(勾選下方選項</strong>以及輸入日期後，系統會把該批舉發單設定相同的<strong>填單日</strong> 以及 <strong>應到案日</strong>)
				</span>
				</td>
			</tr>
			<tr>
				<td>
				<%if sys_City="基隆市" then%>
					<input type="hidden" name="FillFlag" value="N" >
					<input type="hidden" name="BillFillDate" value="" size="8" maxlength="7" >
					<input type="hidden" name="DealLineDate" value="" size="8" maxlength="7">
					<span class="style5"><strong>因審計室查核規定，填單日不可修改為建檔日後，如有疑問請洽詢交通隊承辦人。</strong></span>
				<%else%>
					<input type="checkbox" name="FillFlag" value="Y" <%
					if sys_City<>"高雄市" and sys_City<>ApconfigureCityName And sys_City<>"苗栗縣" And sys_City<>"台中市" then
						response.write "checked"
					end if
					%>>修改填單日期
					<input type="text" name="BillFillDate" value="<%=ginitdt(date)%>" size="8" maxlength="7" onBlur="getDealLineDate()">
					
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;應到案日期
					<input type="text" name="DealLineDate" value="" size="8" maxlength="7">
					<br><font color="red" size="5" >(打勾才會改)</font>
				<%End If%>
				</td>
			</tr>


			<tr>
				<td bgcolor="#EBFBE3" align="center">
					<input type="button" value="確 定" name="b1" onclick="funReport_CaseIn();">
					
					<%if sys_City="台中市" then%>
					&nbsp; &nbsp; 
					<input type="checkbox" name="chkQueryCar" value="1">附車籍查詢
					<%end if %>
					<input type="hidden" value="" name="kinds">
				</td>
			</tr>
			<tr>
				<td bgcolor="#EBFBE3" align="center" id="LayerUp">
					
				</td>
			</td>
		</table>

	</form>
<%
conn.close
set conn=nothing
%>
</body>

<script language="JavaScript">
function funReport_CaseIn(){
	var TodayDate=<%=ginitdt(date)%>;
	if (myForm.FillFlag.checked==true && myForm.BillFillDate.value==""){
		alert("請輸入填單日期!");
	}else if(myForm.FillFlag.checked==true && myForm.DealLineDate.value==""){
		alert("請輸入應到案日期!");
	}else if(myForm.FillFlag.checked==true && dateCheck(myForm.BillFillDate.value)==false){
		alert("填單日期輸入錯誤!");
	}else if(myForm.FillFlag.checked==true && myForm.BillFillDate.value.substr(0,1)=="0" ){
		alert("填單日期輸入錯誤，請直接輸入年份，開頭不須補0!");
	}else if(myForm.FillFlag.checked==true && myForm.BillFillDate.value.substr(0,1)=="9" && myForm.BillFillDate.value.length==7 ){
		alert("填單日期輸入錯誤!");
	}else if(myForm.FillFlag.checked==true && myForm.BillFillDate.value.substr(0,1)=="1" && myForm.BillFillDate.value.length==6 ){
		alert("填單日期輸入錯誤!");
	}else if(myForm.FillFlag.checked==true && dateCheck(myForm.DealLineDate.value)==false){
		alert("應到案日期輸入錯誤!");
	}else if(myForm.FillFlag.checked==true && myForm.DealLineDate.value.substr(0,1)=="0" ){
		alert("應到案日期輸入錯誤，請直接輸入年份，開頭不須補0!");
	}else if(myForm.FillFlag.checked==true && myForm.DealLineDate.value.substr(0,1)=="9" && myForm.DealLineDate.value.length==7 ){
		alert("應到案日期輸入錯誤!");
	}else if(myForm.FillFlag.checked==true && myForm.DealLineDate.value.substr(0,1)=="1" && myForm.DealLineDate.value.length==6 ){
		alert("應到案日期輸入錯誤!");
<%if sys_City<>"宜蘭縣" and sys_City<>"台南縣" and sys_City<>"嘉義市" then%>
	}else if(myForm.FillFlag.checked==true && eval(TodayDate) < eval(myForm.BillFillDate.value)){
		alert("填單日期不得比今天晚!");
<%end if%>
	}else if (myForm.FillFlag.checked==true && !ChkIllegalDate(myForm.BillFillDate.value)){
		alert("填單日期已超過三個月期限");
	}else if (myForm.FillFlag.checked==true && !ChkIllegalDateKS(myForm.DealLineDate.value,45)){
		alert("應到案日期與填單日期相差過大，請確認是否有誤。");
	}else{
		if (myForm.FillFlag.checked==true && !ChkIllegalDateDay("B",15)){
			if(confirm('您輸入的填單日期大於今天15天，是否確定要入案？')){
				myForm.b1.disabled=true;
				LayerUp.innerHTML="資料處理中，請勿將本視窗關閉，以及請勿重新整理網頁!!";
				myForm.kinds.value="Report_CaseIn";
				myForm.submit();
			}
		}else if (myForm.FillFlag.checked==true && !ChkIllegalDateDay("S",-7)){
			if(confirm('您輸入的填單日期小於今天 7 天，是否確定要入案？')){
				myForm.b1.disabled=true;
				LayerUp.innerHTML="資料處理中，請勿將本視窗關閉，以及請勿重新整理網頁!!";
				myForm.kinds.value="Report_CaseIn";
				myForm.submit();
			}
		}else{
			myForm.b1.disabled=true;
			LayerUp.innerHTML="資料處理中，請勿將本視窗關閉，以及請勿重新整理網頁!!";
			myForm.kinds.value="Report_CaseIn";
			myForm.submit();
		}
	}	
}
//逕舉由填單日期帶入應到案日期
function getDealLineDate(){
	myForm.BillFillDate.value=myForm.BillFillDate.value.replace(/[^\d]/g,'');
	BFillDateTemp=myForm.BillFillDate.value;
	if (BFillDateTemp.length >= 6){
		Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
		Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
		Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
		var BFillDate=new Date(Byear,Bmonth-1,Bday);
		var DLineDate=new Date()
		DLineDate=DateAdd("d",<%
		if sys_City="高雄縣" or sys_City="高雄市" or sys_City="苗栗縣" Or sys_City=ApconfigureCityName then
			response.write getReportDealDateValue
		else
			response.write "45"
		end if
		%>,BFillDate);
		Dyear=parseInt(DLineDate.getFullYear())-1911;
		Dmonth=parseInt(DLineDate.getMonth())+1;
		Dday=DLineDate.getDate();
		Dyear=Dyear.toString();
		if (Dmonth < 10){
			Dmonth="0"+Dmonth;
		}
		if (Dday < 10){
			Dday="0"+Dday;
		}
		myForm.DealLineDate.value=Dyear+Dmonth+Dday;
	}
}
//檢查應到案日誤差
function ChkIllegalDateKS(IllDate,addDay){
	BFillDateTemp=myForm.BillFillDate.value;
	Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
	Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
	Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
	var BFillDate=new Date(Byear,Bmonth-1,Bday);
	var DLineDate=new Date()
	DLineDate=DateAdd("d",addDay,BFillDate);
	Dyear=parseInt(DLineDate.getFullYear())-1911;
	Dmonth=parseInt(DLineDate.getMonth())+1;
	Dday=DLineDate.getDate();
	Dyear=Dyear.toString();
	if (Dmonth < 10){
		Dmonth="0"+Dmonth;
	}
	if (Dday < 10){
		Dday="0"+Dday;
	}

	if (parseInt(myForm.DealLineDate.value)>parseInt(Dyear+Dmonth+Dday) ){
		return false;
	}else{
		return true;
	}
}
function ChkIllegalDateDay(IllDate,addDay){
	BFillDateTemp="<%=(year(now)-1911)&right("00"&month(now),2)&right("00"&day(now),2)%>";
	Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
	Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
	Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
	var BFillDate=new Date(Byear,Bmonth-1,Bday);
	var DLineDate=new Date()
	DLineDate=DateAdd("d",addDay,BFillDate);
	Dyear=parseInt(DLineDate.getFullYear())-1911;
	Dmonth=parseInt(DLineDate.getMonth())+1;
	Dday=DLineDate.getDate();
	Dyear=Dyear.toString();
	if (Dmonth < 10){
		Dmonth="0"+Dmonth;
	}
	if (Dday < 10){
		Dday="0"+Dday;
	}
	if (IllDate=="B"){
		if (parseInt(myForm.BillFillDate.value)>parseInt(Dyear+Dmonth+Dday) ){
			return false;
		}else{
			return true;
		}
	}else{
		if (parseInt(myForm.BillFillDate.value)<parseInt(Dyear+Dmonth+Dday) ){
			return false;
		}else{
			return true;
		}
	}
}

function KeyDown(){ 

		if (event.keyCode==116){	//F5鎖死
			event.keyCode=0;   
			event.returnValue=false;   
		}
	}
document.onselectstart=new Function("return false"); 
document.onselect=new Function("return false"); 
document.oncontextmenu=new Function("return false");

getDealLineDate()
</script>
</html>
