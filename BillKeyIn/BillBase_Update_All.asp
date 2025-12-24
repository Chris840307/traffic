<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<title>舉發單批次修改</title>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<%
'檢查是否可進入本系統
'AuthorityCheck(223)
'==========cookie==========
'填單人代碼
theRecordMemberID=trim(Session("User_ID"))
'=========fucntion=========
function DateFormatChange(changeDate)
	DateFormatChange=funGetDate(gOutDT(changeDate),1)
	

end function

'==========================
'修改告發單
strwhere=Session("PrintCarDataSQL")	
'response.write strwhere
if trim(request("kinds"))="DB_insert" then

	strUpd=""

	'填單日期
	if trim(request("BillFillDate"))<>"" then
		theBillFillDate=DateFormatChange(trim(request("BillFillDate")))
		if strUpd<>"" then
			strUpd=strUpd&",BillFillDate="&theBillFillDate
		else
			strUpd="BillFillDate="&theBillFillDate
		end if
	end if
	'應到案日期
	if trim(request("DealLineDate"))<>"" then
		theDealLineDate=DateFormatChange(trim(request("DealLineDate")))
		if strUpd<>"" then
			strUpd=strUpd&",DealLineDate="&theBillFillDate
		else
			strUpd="DealLineDate="&theBillFillDate
		end if
	end if
	'限速
	if trim(request("RuleSpeed"))<>"" then
		if strUpd<>"" then
			strUpd=strUpd&",RuleSpeed='"&trim(request("RuleSpeed"))&"'"
		else
			strUpd="RuleSpeed='"&trim(request("RuleSpeed"))&"'"
		end if
	end if
	'違規地點代碼
	if trim(request("IllegalAddressID"))<>"" then
		if strUpd<>"" then
			strUpd=strUpd&",IllegalAddressID='"&trim(request("IllegalAddressID"))&"'"
		else
			strUpd="IllegalAddressID='"&trim(request("IllegalAddressID"))&"'"
		end if
	end if
	'違規地點
	if trim(request("IllegalAddress"))<>"" then
		if strUpd<>"" then
			strUpd=strUpd&",IllegalAddress='"&trim(request("IllegalAddress"))&"'"
		else
			strUpd="IllegalAddress='"&trim(request("IllegalAddress"))&"'"
		end if
	end if
	'舉發人
	if trim(request("BillMem1"))<>"" then
		if strUpd<>"" then
			strUpd=strUpd&",BillMemID1='"&trim(request("BillMemID1"))&"',BillMem1='"&trim(request("BillMemName1"))&"'" &_
				",BillMemID2='"&trim(request("BillMemID2"))&"',BillMem2='"&trim(request("BillMemName2"))&"'" &_
				",BillMemID3='"&trim(request("BillMemID3"))&"',BillMem3='"&trim(request("BillMemName3"))&"'" &_
				",BillMemID4='"&trim(request("BillMemID4"))&"',BillMem4='"&trim(request("BillMemName4"))&"'"
		else
			strUpd="BillMemID1='"&trim(request("BillMemID1"))&"',BillMem1='"&trim(request("BillMemName1"))&"'" &_
				",BillMemID2='"&trim(request("BillMemID2"))&"',BillMem2='"&trim(request("BillMemName2"))&"'" &_
				",BillMemID3='"&trim(request("BillMemID3"))&"',BillMem3='"&trim(request("BillMemName3"))&"'" &_
				",BillMemID4='"&trim(request("BillMemID4"))&"',BillMem4='"&trim(request("BillMemName4"))&"'"
		end if
	end if
	'舉發單位
	if trim(request("BillUnitID"))<>"" then
		if strUpd<>"" then
			strUpd=strUpd&",BillUnitID='"&trim(request("BillUnitID"))&"'"
		else
			strUpd="BillUnitID='"&trim(request("BillUnitID"))&"'"
		end if
	end if

	strLoop="select a.SN,a.IllegalDate from BillBaseView a,MemberData b where a.RecordMemberID=b.MemberID(+) "&strwhere
	set rsLoop=conn.execute(strLoop)
	If Not rsLoop.Bof Then rsLoop.MoveFirst 
	While Not rsLoop.Eof
		'違規日期
		if trim(request("IllegalDate"))<>"" then
			if strUpd<>"" then
				theIllegalDate=funGetDate(gOutDT(request("IllegalDate") ) &" "&hour(rsLoop("IllegalDate"))&":"&minute(rsLoop("IllegalDate")),1)
				strID=",IllegalDate=TO_DATE('"&gOutDT(request("IllegalDate"))&" "&hour(rsLoop("IllegalDate"))&":"&minute(rsLoop("IllegalDate"))&":0','YYYY/MM/DD/HH24/MI/SS')"
			else
				strID="IllegalDate=TO_DATE('"&gOutDT(request("IllegalDate"))&" "&hour(rsLoop("IllegalDate"))&":"&minute(rsLoop("IllegalDate"))&":0','YYYY/MM/DD/HH24/MI/SS')"
			end if
		end if
		strBillUpd="update BillBase set "&strUpd&strID&" where SN="&trim(rsLoop("SN"))
		'response.write strBillUpd
		conn.execute strBillUpd
		ConnExecute strBillUpd,353
	rsLoop.MoveNext
	Wend
	rsLoop.close
	set rsLoop=nothing
%>
<script language="JavaScript">
	alert("資料修改完成");
	opener.myForm.submit();
	window.close();
</script>
<%
end if


%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

	<form name="myForm" method="post">
		<table width='800' border='1' align="center" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4" align="left">舉發單批次修改</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right" width="15%">違規日期</td>
				<td align="left" width="25%">
				<input type="text" size="10" maxlength="6" value="" name="IllegalDate" onfocus="this.select()" onkeydown="funTextControl(this);" style=ime-mode:disabled onblur="getDealLineDate()">
				</td>
				<td bgcolor="#FFFFCC" align="right" width="15%">填單日期</td>
				<td align="left" width="45%">
				<input type="text" size="10" value="" maxlength="6" name="BillFillDate" onBlur="getDealLineDate_Stop();" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">應到案日期</td>
				<td align="left">
				<input type="text" size="10" maxlength="6" name="DealLineDate" value="" onBlur="value=value.replace(/[^\d]/g,'')" onkeydown="funTextControl(this);" style=ime-mode:disabled>
				</td>
				<td bgcolor="#FFFFCC" align="right">限速</td>
				<td align="left">
				<input type="text" size="10" name="RuleSpeed" onBlur="RuleSpeedforLaw()" onkeydown="funTextControl(this);" style=ime-mode:disabled  value="">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">違規地點代碼</td>
				<td align="left">
				<input type="text" size="10" value="" name="IllegalAddressID" onKeyUp="getillStreet();" onkeydown="funTextControl(this);" style=ime-mode:disabled onblur="getillStreet2();">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Street.asp","WebPage3","left=0,top=0,location=0,width=700,height=455,resizable=yes,scrollbars=yes")'>
				</td>
				<td bgcolor="#FFFFCC" align="right">違規地點</td>
				<td align="left">
				<input type="text" size="34" value="" name="IllegalAddress" style=ime-mode:active onkeyup="AutoGetIllStreet();" onkeydown="funTextControl(this);">
				</td>
			</tr>
			<tr>
				<td align="right"></td>
				<td align="left"></td>
				<td align="right"></td>
				<td align="left"></td>
			<tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">舉發人代碼1</td>
				<td align="left">
				<input type="text" size="5" name="BillMem1" onkeyup="getBillMemID1();" onkeydown="funTextControl(this);" style=ime-mode:disabled value="">
				<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=1","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer12" style="position:absolute ; width:90px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID1">
					<input type="hidden" value="" name="BillMemName1">
				</td>
				<td bgcolor="#FFFFCC" align="right">舉發人代碼2</td>
				<td align="left">
					<input type="text" size="5" name="BillMem2" onkeyup="getBillMemID2();" onkeydown="funTextControl(this);" style=ime-mode:disabled value="">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=2","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer13" style="position:absolute ; width:90px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID2">
					<input type="hidden" value="" name="BillMemName2">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">舉發人代碼3</td>
				<td align="left">
					<input type="text" size="5" name="BillMem3" onkeyup="getBillMemID3();" onkeydown="funTextControl(this);" style=ime-mode:disabled value="">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=3","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer14" style="position:absolute ; width:90px; height:30px; z-index:0;layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID3">
					<input type="hidden" value="" name="BillMemName3">
				</td>
				<td bgcolor="#FFFFCC" align="right">舉發人代碼4</td>
				<td align="left">
					<input type="text" size="5" name="BillMem4" onkeyup="getBillMemID4();" onkeydown="funTextControl(this);" style=ime-mode:disabled value="">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_MemID.asp?MemOrder=4","WebPage1","left=0,top=0,location=0,width=700,height=555,resizable=yes,scrollbars=yes")'>
					<div id="Layer17" style="position:absolute ; width:90px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					<input type="hidden" value="" name="BillMemID4">
					<input type="hidden" value="" name="BillMemName4">
				</td>
			</tr>
			<tr>
				<td align="right"></td>
				<td align="left"></td>
				<td align="right"></td>
				<td align="left"></td>
			<tr>
			<tr>
				<td bgcolor="#FFFFCC" align="right">舉發單位</td>
				<td align="left" colspan="3">
					<input type="text" size="5" name="BillUnitID" onkeyup="getUnit();" onkeydown="funTextControl(this);" style=ime-mode:disabled value="">
					<img src="../Image/BillkeyInButton.jpg" width="25" height="23" onclick='window.open("Query_Unit.asp?SType=U","WebPage2","left=0,top=0,location=0,width=700,height=575,resizable=yes,scrollbars=yes")'>
					<span class="style5">
					<div id="Layer6" style="position:absolute ; width:227px; height:30px; z-index:0; layer-background-color: #CCFFFF; border: 1px none #000000;"></div>
					</span>
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" align="center" colspan="4">
				<input type="button" name="btn1" value="整批修改" onclick="InsertBillVase()">
				<input type="button" name="btn2" value="離開" onclick="window.close()">
				</td>
			</tr>
		</table>
		<br>
		<center><font size="5" color="red">此作業會把目前所選擇的案件資料<b>整批修改</b>為上述欄位資料 </font></center>
		<input type="hidden" value="" name="kinds">
		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
var TDLawNum=0;
var TDLawErrorLog1=0;
var TDLawErrorLog2=0;
var TDLawErrorLog3=0;
var TDLawErrorLog4=0;
var TDStationErrorLog=0;
var TDUnitErrorLog=0;
var TDFastenerErrorLog1=0;
var TDFastenerErrorLog2=0;
var TDFastenerErrorLog3=0;
var TDMemErrorLog1=0;
var TDMemErrorLog2=0;
var TDMemErrorLog3=0;
var TDMemErrorLog4=0;
var ChkCarIlldateFlag=0;
var TDProjectIDErrorLog=0;
var TDVipCarErrorLog=0;
var TodayDate=<%=ginitdt(date)%>;
MoveTextVar("IllegalDate,BillFillDate||DealLineDate,RuleSpeed||IllegalAddressID,IllegalAddress||BillMem1,BillMem2||BillMem3,BillMem4||BillUnitID");
//修改告發單
function InsertBillVase(){
	var error=0;
	var errorString="";
	if (myForm.IllegalDate.value!=""){
		if (!dateCheck( myForm.IllegalDate.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：違規日期輸入錯誤。";
		}
	}
	if (myForm.BillFillDate.value!=""){
		if (!dateCheck( myForm.BillFillDate.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：填單日期輸入錯誤。";
		}
	}
	if (myForm.DealLineDate.value!=""){
		if (!dateCheck( myForm.DealLineDate.value )){
			error=error+1;
			errorString=errorString+"\n"+error+"：應到案日期輸入錯誤。";
		}
	}
	if (myForm.IllegalDate.value=="" && myForm.BillFillDate.value=="" && myForm.DealLineDate.value=="" && myForm.RuleSpeed.value=="" && myForm.IllegalAddressID.value=="" && myForm.IllegalAddress.value=="" && myForm.BillMem1.value=="" && myForm.BillMem2.value=="" && myForm.BillMem3.value=="" && myForm.BillMem4.value=="" && myForm.BillUnitID.value==""){
		error=error+1;
		errorString=errorString+"\n"+error+"：請填入任一欄位。";
	}
	if (error==0){
		if(confirm('是否確定要修改整批舉發單列表上的舉發單資料？')){
			myForm.kinds.value="DB_insert";
			myForm.submit();
		}
	}else{
		alert(errorString);
	}
}

function getDealLineDate(){
	if(TodayDate < myForm.IllegalDate.value){
		alert("違規日期不得大於今天!!");
		myForm.IllegalDate.select();
	}
	myForm.IllegalDate.value=myForm.IllegalDate.value.replace(/[^\d]/g,'');

}
function getDealLineDate_Stop(){
	myForm.BillFillDate.value=myForm.BillFillDate.value.replace(/[^\d]/g,'');
}
function RuleSpeedforLaw(){
	myForm.RuleSpeed.value=myForm.RuleSpeed.value.replace(/^[^\d]+|[^\d.]|,+$/g,'');
}
//違規地點代碼(ajax)
function getillStreet(){
<%if sys_City<>"基隆市" and sys_City<>"彰化縣" then%>
		myForm.IllegalAddressID.value=myForm.IllegalAddressID.value.toUpperCase();
<%end if%>
	if (event.keyCode!=13){
		if (event.keyCode==116){	
			event.keyCode=0;
			OstreetID=myForm.IllegalAddressID.value;

			window.open("Query_Street.asp?OstreetID="+OstreetID,"WebPage_Street_People2","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
		}else if (myForm.IllegalAddressID.value.length > 2){
			var illAddrNum=myForm.IllegalAddressID.value;
			runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
		}
	}
}
//違規地點代碼OnBlur
function getillStreet2(){
	if (myForm.IllegalAddress.value==""){
		if (myForm.IllegalAddressID.value.length > 2){
			var illAddrNum=myForm.IllegalAddressID.value;
			runServerScript("getIllStreet.asp?illAddrID="+illAddrNum);
		}
	}
}
function AutoGetIllStreet(){	//按F5可以直接顯示相關路段
	if (event.keyCode==116){	
		event.keyCode=0;
		Ostreet=myForm.IllegalAddress.value;
		window.open("Query_Street.asp?OStreet="+Ostreet,"WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
}
//舉發人一(ajax)
function getBillMemID1(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=1","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem1.value.length > 2){
		var BillMemNum=myForm.BillMem1.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=1&MemID="+BillMemNum);
	}else if (myForm.BillMem1.value.length <= 2 && myForm.BillMem1.value.length > 0){
		Layer12.innerHTML=" ";
		myForm.BillMemID1.value="";
		myForm.BillMemName1.value="";
		TDMemErrorLog1=1;
	}else{
		Layer12.innerHTML=" ";
		myForm.BillMemID1.value="";
		myForm.BillMemName1.value="";
		TDMemErrorLog1=0;
	}
}
//舉發人二(ajax)
function getBillMemID2(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=2","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem2.value.length > 2){
		var BillMemNum=myForm.BillMem2.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=2&MemID="+BillMemNum);
	}else if (myForm.BillMem2.value.length <= 2 && myForm.BillMem2.value.length > 0){
		Layer13.innerHTML=" ";
		myForm.BillMemID2.value="";
		myForm.BillMemName2.value="";
		TDMemErrorLog2=1;
	}else{
		Layer13.innerHTML=" ";
		myForm.BillMemID2.value="";
		myForm.BillMemName2.value="";
		TDMemErrorLog2=0;
	}
}
//舉發人三(ajax)
function getBillMemID3(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=3","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem3.value.length > 2){
		var BillMemNum=myForm.BillMem3.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=3&MemID="+BillMemNum);
	}else if (myForm.BillMem3.value.length <= 2 && myForm.BillMem3.value.length > 0){
		Layer14.innerHTML=" ";
		myForm.BillMemID3.value="";
		myForm.BillMemName3.value="";
		TDMemErrorLog3=1;
	}else{
		Layer14.innerHTML=" ";
		myForm.BillMemID3.value="";
		myForm.BillMemName3.value="";
		TDMemErrorLog3=0;
	}
}
//舉發人四(ajax)
function getBillMemID4(){
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_MemID.asp?MemOrder=4","WebPage_Street_People","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillMem4.value.length > 2){
		var BillMemNum=myForm.BillMem4.value;
		runServerScript("getBillMemID.asp?MType=Car&MemOrder=4&MemID="+BillMemNum);
	}else if (myForm.BillMem4.value.length <= 2 && myForm.BillMem4.value.length > 0){
		Layer17.innerHTML=" ";
		myForm.BillMemID4.value="";
		myForm.BillMemName4.value="";
		TDMemErrorLog4=1;
	}else{
		Layer17.innerHTML=" ";
		myForm.BillMemID4.value="";
		myForm.BillMemName4.value="";
		TDMemErrorLog4=0;
	}
}
//舉發單位(ajax)
function getUnit(){
	myForm.BillUnitID.value=myForm.BillUnitID.value.toUpperCase();
	if (event.keyCode==116){	
		event.keyCode=0;
		window.open("Query_Unit.asp?SType=U","WebPage_Station12","left=0,top=0,location=0,width=800,height=660,resizable=yes,scrollbars=yes");
	}
	if (myForm.BillUnitID.value.length > 1){
		var BillUnitNum=myForm.BillUnitID.value;
		runServerScript("getBillUnitID.asp?BillUnitID="+BillUnitNum);
	}else{
		Layer6.innerHTML=" ";
		TDUnitErrorLog=1;
	}
}

function funTextControl(obj){
	if (event.keyCode==13){ //Enter換欄
		event.keyCode=0;
		event.returnValue=false;
		CodeEnter(obj.name);
	}	
	/*if (event.keyCode==37){ //左換欄
		event.keyCode=0;
		event.returnValue=false;
		CodeMoveLeft(obj.name);
	}*/else if (event.keyCode==38){ //上換欄
		event.keyCode=0;
		event.returnValue=false;
		CodeMoveLeft(obj.name);
	}/*else if (event.keyCode==39){ //右換欄
		event.keyCode=0;
		event.returnValue=false;
		CodeMoveRight(obj.name);
	}*/else if (event.keyCode==40){ //下換欄
		event.keyCode=0;
		event.returnValue=false;
		CodeMoveRight(obj.name);
	}
}
</script>
</html>
