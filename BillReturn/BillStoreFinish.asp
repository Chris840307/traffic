<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>寄存送達期滿註記</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
if sys_City="花蓮縣" then
	CarName="姓名"
else
	CarName="車號"
end if

if trim(request("DB_Selt"))="Selt" then
	Sys_BillNo=Split(trim(request("item")),",")
	Sys_BackCause=Split(request("Sys_BackCause"),",")
	Sys_BackDate=Split(request("Sys_BackDate"),",")
	EffectDate=Split(request("EffectDate"),",")
	Sys_MailStation=Split(request("MailStation"),",")
	Sys_Now=now
	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" and trim(Sys_BackDate(i))<>"" then
			Sys_Now=DateAdd("s",1,Sys_Now)
			strSQL="select BillNo from MailStationReturn where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
			set rschk=conn.execute(strSQL)
			If Not rschk.eof Then
				strSQL="Update MailStationReturn set ArriveDate="&funGetDate(gOutDT(EffectDate(i)),0)&",MailStation='"&trim(Sys_MailStation(i))&"',ReturnDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",ResonID='"&trim(Sys_BackCause(i))&"',UserMarkMemberID="&Session("User_ID")&",UserMarkDate="&funGetDate((Sys_Now),1)&" where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
				conn.execute(strSQL)
			else
				strSQL="Insert Into MailStationReturn(BillNo,ArriveDate,MailStation,ReturnDate,ResonID,UserMarkMemberID,UserMarkDate) values('"&trim(Ucase(Sys_BillNo(i)))&"',"&funGetDate(gOutDT(EffectDate(i)),0)&",'"&trim(Sys_MailStation(i))&"',"&funGetDate(gOutDT(Sys_BackDate(i)),0)&",'"&trim(Sys_BackCause(i))&"',"&Session("User_ID")&","&funGetDate((Sys_Now),1)&")"
				conn.execute(strSQL)
			end if
			rschk.close
		end if
	next
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
Sys_BackCause="<input type=Hidden name='Sys_BackCause' value='1'>"
%>
<form name=myForm method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
	<font size="3" color="deepred">請確認該資料曾經做過 <b>寄存送達註記，並曾上傳監理站</b> &nbsp;&nbsp;&nbsp;&nbsp; <a Href="storefinishhelp.doc">使用說明下載</a></font>
	</tr>
	<tr>
		<td height="45" bgcolor="#FFCC33" class="pagetitle"><strong>寄存送達期滿註記
		<font size="3" color="deepred">此功能主要提供整批退回紅單移送監理站使用，並不會上傳監理站，完成後請至 各式清冊/舉發單列印 功能下方以註記日期列印 清冊</font></strong></td>
	
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						預設單退日期統一為&nbsp;<input name="Sys_BackDateMain" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BackDateMain');">
						&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuDate('Sys_BackDateMain','Sys_BackDate');">
						&nbsp;<font size="2">非必要選項,也可以由下方設定各舉發單不同的單退日期</font>
						<br>
						預設送達日期統一為&nbsp;<input name="Sys_DefEffectDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_DefEffectDate');">
						&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuDate('Sys_DefEffectDate','EffectDate');">
						&nbsp;<font size="2">非必要選項,也可以由下方設定各舉發單不同的送達日期</font>
						<br>
						預設郵局統一為&nbsp;
						<input name="Sys_mailStation" type="text" class="btn1" size="17" maxlength="15">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuDate('Sys_mailStation','mailStation');">
						&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的郵局</font>
						<hr>
						&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funSelt();">
						<img src="space.gif" width="9" height="8">
						<input type="button" name="insert" value="再多30筆" onClick="insertRow(fmyTable)">
						<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">寄存送達註記列表<img src="space.gif" width="12" height="8"> <b>郵局 / 送達日 為非必填欄位. 其餘為必填欄位<br></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<table id='fmyTable' width='978' border='0' bgcolor='#FFFFFF'>
				<tr bgcolor="#ffffff">
					<td align='center' bgcolor="#ffffff" nowrap></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td height="35" bgcolor="#FFDD77">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="">
<input type="Hidden" name="chkcnt" value="">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var cunt=0;
function insertRow(isTable){
	for(i=0;i<=29;i++){
		Rindex = isTable.rows.length;
		if (isTable.rows.length > 0) {
		    Cindex = isTable.rows[Rindex - 1].cells.length;
		} else {
		    Cindex = 0;
		}
		if (Rindex == 0 || Cindex == 1) {
		    nextRow = isTable.insertRow(Rindex);
		    txtArea = nextRow.insertCell(0);
		} else {
		    if (cunt == 0) {
		        Cindex = 0;
		        isTable.rows[Rindex - 1].deleteCell();
		    }
		    txtArea = isTable.rows[Rindex - 1].insertCell(Cindex);
		}
		cunt++;
		//txt_nameStr = "item"+cunt;
		txtArea.innerHTML ="單號<input type=text name='item' size=10 class='btn1' onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;<%=Sys_BackCause%>&nbsp;&nbsp;單退日<input type=text name='Sys_BackDate' size=8 class='btn1' onkeyup='funkeyChk(this);' onkeydown='keyBackDate("+cunt+");'>&nbsp;&nbsp;送達日<input type=text name='EffectDate' size=8 class='btn1' onkeyup='funkeyChk(this);' onkeydown='keyEffectDate("+cunt+");'>&nbsp;&nbsp;<%=CarName%><input type=text name='CarNo' size=8 class='btn1' readOnly>&nbsp;&nbsp;郵局<input type=text name='mailStation' size=15 class='btn1' maxlength=25 onkeydown='keyMailStation("+cunt+");'>";
	}
}
function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}
function keyFunction(itemcnt) {
	myForm.item[itemcnt-1].value=myForm.item[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9||myForm.item[itemcnt-1].length>=9) {
		if (chkBillNo(itemcnt)){
			if (myForm.item[itemcnt-1].value!=''){
				myForm.chkcnt.value=itemcnt;
				runServerScript("chkStoreFinish.asp?BillNo="+myForm.item[itemcnt-1].value);
			}
		}else{
			alert("單號長度必須為9碼!!");
		}
	}
}

function keyBackDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		myForm.EffectDate[itemcnt-1].focus();
	}
}

function keyEffectDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		myForm.mailStation[itemcnt-1].focus();
	}
}

function keyMailStation(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		if(itemcnt<myForm.item.length){
			myForm.item[itemcnt].focus();
		}
	}
}

function DeleteRow(isTable){
	if(isTable.rows.length>0){
		Rindex = isTable.rows.length;
		Cindex = isTable.rows(Rindex-1).cells.length;
		if(Cindex==1){
			cunt--;
			isTable.rows(Rindex-1).deleteCell();
			isTable.deleteRow();
		}else{
			cunt--;
			isTable.rows(Rindex-1).deleteCell();
		}
	}
}


function funStation(){
	for(i=0;i<myForm.mailStation.length;i++){
		myForm.mailStation[i].value=myForm.Sys_mailStation.value;
	}
}

function funSelt(){
	var err=0;
	for(i=0;i<myForm.item.length;i++){
		if(myForm.item[i].value!=''){
			if(myForm.Sys_BackDate[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 行單退日期不可空白!!");
				break;
			}
		}
	}
	if(err==0){
		myForm.DB_Selt.value="Selt";
		myForm.submit();
	}
}
for(j=0;j<=3;j++){
	insertRow(fmyTable);
}
function funDefuDate(obj,defName){
	for(i=0;i<eval("myForm."+defName).length;i++){
		eval("myForm."+defName+"["+i+"]").value=eval("myForm."+obj).value;
	}
}
</script>