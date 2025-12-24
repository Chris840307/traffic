<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>停管催繳-二次郵寄設定</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

if trim(request("DB_Selt"))="Selt" then
	Sys_BillNo=Split(trim(request("item")),",")
	Sys_CarNo=Split(trim(request("Sys_CarNo")),",")
	Sys_DirverAddress=Split(request("DirverAddress"),",")
	Sys_DriverZip=Split(request("DriverZip"),",")
	Sys_Driver=Split(request("Driver"),",")
	Sys_SendKind=Split(request("SendKind"),",")
	Sys_DeallineDate=Split(request("Sys_DeallineDate"),",")
	Sys_MailDate=Split(request("Sys_MailDate"),",")
	Sys_Now=now

	sys_maiNumberSN=""

	If sys_City="台東縣" Then
		sys_maiNumberSN="MailNumber_Stop_Sn"

	elseif sys_City="花蓮縣" then
		sys_maiNumberSN="stopcarbillprintno"

	end If 
	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" and trim(Sys_DirverAddress(i))<>"" and trim(Sys_SendKind(i))="1" then

			Sys_Now=DateAdd("s",1,Sys_Now)
			if Not ifnull(Sys_DirverAddress(i)) then
				strSQL="Select CarNo from StopCarSendAddress where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
				set rsbill=conn.execute(strSQL)
				If rsBill.eof Then
					strSQL="Insert Into StopCarSendAddress Values('"&trim(Ucase(Sys_BillNo(i)))&"','"&Sys_CarNo(i)&"','1',"&funGetDate((Sys_Now),1)&")"
					conn.execute(strSQL)
				else
					strSQL="Update StopCarSendAddress set UserMarkDate="&funGetDate((Sys_Now),1)&" where BillNo='"&trim(Sys_BillNo(i))&"'"
					conn.execute(strSQL)
				end if
				rsbill.close


				MailNo=""

				strSQL="select LPAD("& sys_maiNumberSN &".NextVal,6,'0') cmt from Dual"
				set rsnum=conn.execute(strSQL)
				MailNo=trim(rsnum("cmt"))
				rsnum.close

				strSQL="Update StopBillMailHistory set StoreAndSendMailNumber='"&MailNo&"' where BillNo='"&trim(Sys_BillNo(i))&"'"

				conn.execute(strSQL)
				

				strSQL="update billbase set DriverAddress='"&Sys_DirverAddress(i)&"',DriverZip='"&Sys_DriverZip(i)&"' where ImageFileNameB='"&trim(Sys_BillNo(i))&"'"

				conn.execute(strSQL)

				If not ifnull(Sys_Driver(i)) Then
					strSQL="update billbase set Driver='"&trim(Sys_Driver(i))&"' where ImageFileNameB='"&trim(Sys_BillNo(i))&"'"

					conn.execute(strSQL)
				End if

				If not ifnull(Sys_DeallineDate(i)) Then
					strSQL="update billbase set DeallineDate="&funGetDate(gOutDT(Sys_DeallineDate(i)),0)&" where ImageFileNameB='"&trim(Sys_BillNo(i))&"'"

					conn.execute(strSQL)
				End if

				If not ifnull(Sys_MailDate(i)) Then
					strSQL="update StopBillMailHistory set StoreAndSendMailDate="&funGetDate(gOutDT(Sys_MailDate(i)),0)&" where BillNo='"&trim(Sys_BillNo(i))&"'"

					conn.execute(strSQL)
				End if

			end if
		end if
	next
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
%>
<form name=myForm method="post">
<table width="100%" border="0" bgcolor="#ffffff">

	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						&nbsp;&nbsp;寄發種類代碼<input name="Sys_defSendKind" type="text" class="btn1" size="10" maxlength="15">
						&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefSendKind();"><font size="2">&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的寄發種類</font>
						<br>
						&nbsp;&nbsp;郵寄日期<input name="Sys_defMailDate" type="text" class="btn1" size="10" maxlength="15">
						&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="fundefMailDate();"><font size="2">&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的郵寄日</font>
						<br>
						&nbsp;&nbsp;繳費期限<input name="Sys_defDeallineDate" type="text" class="btn1" size="10" maxlength="15">
						&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="fundefDeallineDate();"><font size="2">&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的繳費期限</font>
						<br><br>
						&nbsp;&nbsp;<input type="button" name="btnOK" value="確定儲存" onclick="funSelt();">
						<img src="space.gif" width="9" height="8">
						<input type="button" name="insert" value="再多30筆" onClick="insertRow(fmyTable)">
						<br>
						<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
						<font size=3 color=red><br><b>寄發種類代碼：1表示需第二次寄發，2表示不需再次寄發</b></font>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">停管催繳-二次郵寄設定 ( 輸入完成按Enter可自動跳到下一格 )</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<table id='fmyTable' width='100%' border='0' bgcolor='#FFFFFF'>
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
		txtArea.innerHTML ="單號<input type=text name='item' size=8 class='btn1' maxlength=16 onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;車號<input type=text name='Sys_CarNo' size=8 class='btn1' onkeydown='keyFunction("+cunt+");' ReadOnly>　　一次區號<input type=text name='OwnerZip' size=2 class='btn1' onkeydown='funOwnerZip("+cunt+");' MaxLength=3>&nbsp;&nbsp;一次地址<input type=text name='OwnerAddress' size=35 class='btn1' onkeydown='funAddress("+cunt+");'>&nbsp;&nbsp;寄發種類代碼<input type=text name='SendKind' size=1 class='btn1' onkeydown='funSendKind("+cunt+");';' MaxLength=1><br>郵寄日期<input type=text name='Sys_MailDate' size=6 class='btn1' onkeyup='chknumber(this);' onkeydown='keyMailDate("+cunt+");' maxlength='7'>繳費期限<input type=text name='Sys_DeallineDate' size=6 class='btn1' onkeyup='chknumber(this);' onkeydown='keyBackDate("+cunt+");' maxlength='7'>&nbsp;戶藉區號<input type=text name='DriverZip' size=2 class='btn1' onkeydown='funDriverZip("+cunt+");' MaxLength=3>&nbsp;&nbsp;戶藉地址<input type=text name='DirverAddress' size=35 class='btn1' onkeydown='funAddress("+cunt+");'>&nbsp;&nbsp;寄件人<input type=text name='Driver' size=8 class='btn1' onkeydown='funDriver("+cunt+");'>&nbsp;&nbsp;<input type=Hidden name='Sys_ZipName' value=''><hr>";
	}
}
function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}

function keyMailDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		if(itemcnt<myForm.Sys_MailDate.length){
			myForm.Sys_DeallineDate[itemcnt-1].focus();
		}
	}
}

function keyBackDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		if(itemcnt<myForm.Sys_DeallineDate.length){
			myForm.item[itemcnt].focus();
		}
	}
}

function keyFunction(itemcnt) {
	myForm.item[itemcnt-1].value=myForm.item[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9) {
		if (myForm.item[itemcnt-1].value!=''){
			myForm.chkcnt.value=itemcnt;

			myForm.item[itemcnt-1].value=("000000000000000"+myForm.item[itemcnt-1].value).substr(("000000000000000"+myForm.item[itemcnt-1].value).length-16,16);

			runServerScript("chkStopToAddOwnerAdder.asp?Sys_BillNo="+myForm.item[itemcnt-1].value);
		}
	}
}

function funOwnerZip(itemcnt) {
	myForm.OwnerZip[itemcnt-1].value=myForm.OwnerZip[itemcnt-1].value.replace(/[^\d]/g,'');
	if(myForm.OwnerZip[itemcnt-1].value.length<4){
		if (event.keyCode==13||event.keyCode==9) {
			if(itemcnt<myForm.OwnerZip.length){
				myForm.OwnerAddress[itemcnt-1].focus();
			}
			myForm.chkcnt.value=itemcnt;
			runServerScript("chkStoreAndSendZip.asp?Zip="+myForm.OwnerZip[itemcnt-1].value);
		}
	}else{
		alert('郵地區號錯誤!!');
	}
}

function funDriverZip(itemcnt) {
	myForm.DriverZip[itemcnt-1].value=myForm.DriverZip[itemcnt-1].value.replace(/[^\d]/g,'');
	if(myForm.DriverZip[itemcnt-1].value.length<4){
		if (event.keyCode==13||event.keyCode==9) {
			if(itemcnt<myForm.DriverZip.length){
				myForm.DirverAddress[itemcnt-1].focus();
			}
			myForm.chkcnt.value=itemcnt;
			runServerScript("chkStopDriverStoreAndSendZip.asp?Zip="+myForm.DriverZip[itemcnt-1].value);
		}
	}else{
		alert('郵地區號錯誤!!');
	}
}

function funZipName(itemcnt) {
	runServerScript("chkStoreAndSendZip.asp?Zip="+myForm.OwnerZip[itemcnt-1].value);
}
function funSendKind(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(myForm.SendKind[itemcnt-1].value>3 || myForm.SendKind[itemcnt-1].value<1){
			alert('寄發種類不可大於3');
		}else{
			myForm.item[itemcnt].focus();
		}
	}
}

function funAddress(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		myForm.Driver[itemcnt-1].focus();
	}
}

function funDriver(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.Driver.length){
			myForm.SendKind[itemcnt-1].focus();
		}
	}
}

function fundefMailDate(){
	for(i=0;i<myForm.Sys_MailDate.length;i++){
		myForm.Sys_MailDate[i].value=myForm.Sys_defMailDate.value;
	}
}

function fundefDeallineDate(){
	for(i=0;i<myForm.Sys_DeallineDate.length;i++){
		myForm.Sys_DeallineDate[i].value=myForm.Sys_defDeallineDate.value;
	}
}

function funDefSendKind(){
	for(i=0;i<myForm.SendKind.length;i++){
		myForm.SendKind[i].value=myForm.Sys_defSendKind.value;
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


function funSelt(){
	var err=0;
	for(i=0;i<myForm.item.length;i++){
		if(myForm.item[i].value!=''){
			if(myForm.SendKind[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 寄發種類不可空白!!");
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
</script>