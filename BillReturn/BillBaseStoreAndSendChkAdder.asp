<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>單退(寄存)二次郵寄</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 16px; color:#ff0000; }
-->
</style>
</HEAD>
<BODY>
<%
Server.ScriptTimeout=6000
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
	Sys_BillNo=Split(Ucase(trim(request("item"))),",")
	Sys_StoreAndSendSendDate=Split(request("StoreAndSendSendDate"),",")
	Sys_BackCause=Split(request("Sys_BackCause"),",")
	Sys_mailNumber=Split(request("mailNumber"),",")
	Sys_OwnerAddress=Split(request("OwnerAddress"),",")
	Sys_OwnerZip=Split(request("OwnerZip"),",")
	Sys_DriverAddress=Split(request("DriverAddress"),",")
	Sys_DriverZip=Split(request("DriverZip"),",")
	Sys_ChkSend=Split(request("ChkSend"),",")
	EffectDate=Split(request("EffectDate"),",")
	Sys_Now=DateAdd("n", -5, now)
	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" then
			Sys_Now=DateAdd("s",1,Sys_Now)

			if trim(Sys_BackCause(i))="5" or trim(Sys_BackCause(i))="6" or trim(Sys_BackCause(i))="7" or trim(Sys_BackCause(i))="T" then
				Sys_BackCauseTmp=trim(Sys_BackCause(i))
				Str_BackCauseSQL=""
			else
				Sys_BackCauseTmp="T"
				strSqlBack="select Content from DciCode where TypeID=7 and ID='"&trim(Sys_BackCause(i))&"'"
				set rsSqlBack=conn.execute(strSqlBack)
				if not rsSqlBack.eof then
					Str_BackCauseSQL=",Note=Note || '退回原因："& trim(rsSqlBack("Content"))&"'"
				end if
				rsSqlBack.close
				set rsSqlBack=nothing
			end if

			if trim(Sys_OwnerAddress(i))<>"" then
				strSQL="Update BillBase set OwnerZIP='"&trim(Sys_OwnerZip(i))&"',OwnerAddress='"&chstr(trim(Sys_OwnerAddress(i)))&"',DriverZip='"&trim(Sys_DriverZip(i))&"',DriverAddress='"&chstr(trim(Sys_DriverAddress(i)))&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"' and recordstateid=0"

				conn.execute(strSQL)

				strSQL="Update BillBaseDciReturn set DriverHomeZIP='"&trim(Sys_DriverZip(i))&"',DriverHomeAddress='"&chstr(trim(Sys_DriverAddress(i)))&"',DriverCounty=(select SubStr(ZipName,1,3) from Zip where ZipID='"&trim(Sys_DriverZip(i))&"') where CarNo=(Select CarNo from billbase where billno='"&trim(Ucase(Sys_BillNo(i)))&"' and recordstateid=0) and ExchangetypeID='A'"

				conn.execute(strSQL)

				strSQL="Update BillBaseDciReturn set DriverHomeZIP='"&trim(Sys_DriverZip(i))&"',DriverHomeAddress='"&chstr(trim(Sys_DriverAddress(i)))&"',DriverCounty=(select SubStr(ZipName,1,3) from Zip where ZipID='"&trim(Sys_DriverZip(i))&"') where CarNo=(Select CarNo from billbase where billno='"&trim(Ucase(Sys_BillNo(i)))&"' and recordstateid=0) and billno='"&trim(Ucase(Sys_BillNo(i)))&"' and ExchangetypeID='W'"

				conn.execute(strSQL)
			end if

			if trim(Sys_ChkSend(i))="是" then

				strSQL="Update BillMailHistory set StoreAndSendSendDate="&funGetDate(gOutDT(Sys_StoreAndSendSendDate(i)),0)&" where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
				conn.execute(strSQL)

				strSQL="Update BillMailHistory set ReturnResonID='"&Sys_BackCauseTmp&"',MailReturnDate="&funGetDate(gOutDT(Sys_StoreAndSendSendDate(i)),0)&",ReturnRecordMemberID="&Session("User_ID")&",ReturnReCordDate="&funGetDate((Sys_Now),1)&",UserMarkMemberID="&Session("User_ID")&",UserMarkDate="&funGetDate((Sys_Now),1)&",UserMarkResonID='"&Sys_BackCauseTmp&"',UserMarkReturnDate="&funGetDate(gOutDT(Sys_StoreAndSendSendDate(i)),0)&",MailTypeID=null where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"

				conn.execute(strSQL)
			else
				strSQL="Update BillMailHistory set StoreAndSendReturnResonID='"&trim(Sys_BackCause(i))&"',StoreAndSendMailReturnDate="&funGetDate(gOutDT(Sys_StoreAndSendSendDate(i)),0)&",StoreAndSendRecordMemberID="&Session("User_ID")&",StoreAndSendReCordDate="&funGetDate((Sys_Now),1)&",UserMarkMemberID="&Session("User_ID")&",UserMarkDate="&funGetDate((Sys_Now),1)&",UserMarkResonID='"&trim(Sys_BackCause(i))&"',StoreAndSendEffectDate="&funGetDate(gOutDT(EffectDate(i)),0)&",UserMarkReturnDate="&funGetDate(gOutDT(Sys_StoreAndSendSendDate(i)),0)&",MailTypeID=2 where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"

				conn.execute(strSQL)

			end if
			
			
			if not ifnull(Sys_mailNumber(i)) then
				strSQL="Update BillMailHistory set mailNumber='"&trim(Sys_mailNumber(i))&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
				conn.execute(strSQL)
			end if
			strSQL="Update Billbase set BillStatus=3 "&Str_BackCauseSQL&" where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"' and RecordStateID=0"
			conn.execute(strSQL)
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
		<td height="27" bgcolor="#FFCC33" class="pagetitle"><strong>寄存送達證書 戶籍地址補正 </strong>
		<img src="../Image/dot.gif"></img>
		<a href="../storechkAddrss.doc" target="_blank" >單退二次郵寄簡要說明.doc</a>
		<br><bR>
		透過此功能可修正寄存送達郵記之正確地址。 自動抓取監理站民眾領牌登錄之戶籍地址部份則不適用此功能<br>
		
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						預設退件原因統一為&nbsp;
						<select name="Sys_BackCauseMain" class="btn1">
							<%
						if sys_City="南投縣" then
							strSQL="select ID,Content from DCICode where TypeID=7 and ID in ('5','6','7','T','1','2','3','4','M','P')"
						else
							strSQL="select ID,Content from DCICode where TypeID=7 and ID in ('5','6','7','T')"
						end if
							set rs1=conn.execute(strSQL)
							seltarr="":seltName="":seltIndex=-1
							while Not rs1.eof
								response.write "<option value="""&rs1("ID")&""""
								response.write ">"&rs1("Content")&"</option>"

								seltarr=seltarr&"<option value='"&rs1("ID")&"'>"&rs1("Content")&"</option>"

								seltIndex=seltIndex+1
								seltName=seltName&seltIndex&"."&rs1("Content")&"　"

								rs1.movenext
							wend
							rs1.close

							titleStr=""
							BackCause_btn="<input name='Sys_BackCauseIndex' type=Hidden class='btn1' size=1 maxlength=2 onkeyup=funBackCauseIndex(this,'Sys_BackCause',""+cunt+"");>"
							if sys_City="高雄市" or sys_City="高港局" then
								titleStr="<br><span class=""style1"">"&seltName&"</span>"
								BackCause_btn="<input name='Sys_BackCauseIndex' type=text class='btn1' size=1 maxlength=2 onkeyup=funBackCauseIndex(this,'Sys_BackCause',""+cunt+"");>"
							end if%>
						</select>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuSelt();">
						&nbsp;<font size="2">非必要選項,也可以由下方設定各舉發單不同的單退原因</font>
						<br>
						預設單退日期統一為&nbsp;<input name="Sys_StoreAndSendSendDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_StoreAndSendSendDate');">
						&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuDate('Sys_StoreAndSendSendDate','StoreAndSendSendDate');">
						&nbsp;<font size="2">非必要選項,也可以由下方設定各舉發單不同的單退日期</font>
						<br>
						預設大宗掛號統一為&nbsp;
						<input name="Sys_Number" type="text" class="btn1" size="10" maxlength="15">
						&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funNumber();"><font size="2">&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的大宗掛號</font>
						<br>
						預設送達日期統一為&nbsp;<input name="Sys_DefEffectDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_DefEffectDate');">
						&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuDate('Sys_DefEffectDate','EffectDate');">
						&nbsp;<font size="2">非必要選項,也可以由下方設定各舉發單不同的送達日期</font>
						<br><br>
						&nbsp;&nbsp;<input type="button" name="btnOK" value="確定儲存" onclick="funSelt();">
						<img src="space.gif" width="9" height="8">
						<input type="button" name="insert" value="再多30筆" onClick="insertRow(fmyTable)">
						<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">二次郵寄紀錄列表( 輸入完成按Enter可自動跳到下一格 )　　<font color="red" size="3">請輸入<b>郵遞區號 </b>取得鄉鎮市 或是 戶籍地址內輸入<b>鄉鎮市</b></font> <br><%=titleStr%></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<Div style="overflow:auto;width:100%;height:400px;background:#FFFFFF">
				<table id='fmyTable' width='120%' border='0' bgcolor='#FFFFFF'>
					<tr bgcolor="#ffffff">
						<td align='center' bgcolor="#ffffff" nowrap></td>
					</tr>
				</table>
			</Div>
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
		if(isTable.rows.length>0){
		    Cindex = isTable.rows[Rindex-1].cells.length;
		}else{
		    Cindex=0;
		}
		if(Rindex==0||Cindex==1){
		    nextRow = isTable.insertRow(Rindex);
		    txtArea = nextRow.insertCell(0);
		}else{
		    if(cunt==0){
		        Cindex=0;
		        isTable.rows[Rindex-1].deleteCell();
		    }
		    txtArea =isTable.rows[Rindex-1].insertCell(Cindex);
		}
		cunt++;
		//txt_nameStr = "item"+cunt;
		var cnt_num=("0000"+cunt).substr(("0000"+cunt).length-3,3);
		txtArea.innerHTML =cnt_num+".&nbsp;單號<input type=text name='item' size=6 class='btn1' onkeydown='keyFunction("+cunt+");' onFocus='chkinput(this);'>&nbsp;&nbsp;單退日<input type=text name='StoreAndSendSendDate' size=3 class='btn1' onkeyup='funkeyChk(this);' onkeydown='funSendDate("+cunt+");' maxlength='7'>&nbsp;&nbsp;原因<%=BackCause_btn%><select name='Sys_BackCause' class='btn1' onkeydown='keyBackCause("+cunt+");'><%=seltarr%></select>&nbsp;&nbsp;掛號碼<input type=text name='mailNumber' size=5 class='btn1' onkeydown='keyMailNumber("+cunt+");' maxlength='20'>&nbsp;&nbsp;送達日<input type=text name='EffectDate' size=3 class='btn1' onkeyup='chknumber(this);' onkeydown='keyEffectDate("+cunt+");' maxlength='7'>&nbsp;&nbsp;<%=CarName%><input type=text name='CarNo' size=5 class='btn1' readOnly><br>　&nbsp;再寄：<input type=text name='ChkSend' size=3 class='btn1' readOnly>&nbsp;&nbsp;(戶)地址<input type=text name='DriverZip' size=1 class='btn1' onkeydown='funDriverZip("+cunt+");' MaxLength=3>&nbsp;&nbsp;<input type=text name='DriverAddress' size=60 class='btn1' onkeydown='funDriverAddress("+cunt+");' onblur='funMouseDriverAddress("+cunt+");' maxlength='48'><br>　　　　　　第一次郵寄地<input type=text name='OwnerZip' size=1 class='btn1' onkeydown='funOwnerZip("+cunt+");' MaxLength=3>&nbsp;&nbsp;<input type=text name='OwnerAddress' size=60 class='btn1' onkeydown='funOwnerAddress("+cunt+");' funMouseOwnerAddress='funOwnerAddress("+cunt+");' maxlength='48'><hr>";
	}
}

function keyEffectDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		if(itemcnt<myForm.item.length){
			myForm.item[itemcnt].focus();
		}
	}
}

function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}
function chkinput(obj) {
	obj.style.imeMode="disabled";
}
function keyFunction(itemcnt) {
	//myForm.item[itemcnt-1].value=myForm.item[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9) {
		if (chkBillNo(itemcnt)){
			if (myForm.item[itemcnt-1].value!=''){
				myForm.chkcnt.value=itemcnt;
				runServerScript("chkStoreAndSendChkAdder.asp?BillNo="+myForm.item[itemcnt-1].value);
			}
		}else{
			alert("單號長度必須為9碼!!");
		}
	}
}

function funDriverZip(itemcnt) {
	myForm.OwnerZip[itemcnt-1].value=myForm.OwnerZip[itemcnt-1].value.replace(/[^\d]/g,'');
	if(myForm.OwnerZip[itemcnt-1].value.length<4){
		if (event.keyCode==13||event.keyCode==9) {
			if(itemcnt<myForm.OwnerZip.length){
				myForm.DriverAddress[itemcnt-1].focus();
			}
			myForm.chkcnt.value=itemcnt;
			runServerScript("chkStoreAndSendZipChkDriverAddr.asp?DriverZip="+myForm.DriverZip[itemcnt-1].value);
		}
	}else{
		alert('郵地區號錯誤!!');
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
			runServerScript("chkStoreAndSendZipChkOwnerAddr.asp?OwnerZip="+myForm.OwnerZip[itemcnt-1].value);
		}
	}else{
		alert('郵地區號錯誤!!');
	}
}

function funDriverAddress(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		myForm.chkcnt.value=itemcnt;
		runServerScript("chkStoreAndSendZipChkDriverAddr.asp?ZipName="+myForm.DriverAddress[itemcnt-1].value);
	}
}
function funMouseDriverAddress(itemcnt) {
	myForm.chkcnt.value=itemcnt;
	runServerScript("chkStoreAndSendZipChkDriverAddr.asp?ZipName="+myForm.DriverAddress[itemcnt-1].value);
}

function funOwnerAddress(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		myForm.chkcnt.value=itemcnt;
		runServerScript("chkStoreAndSendZipChkOwnerAddr.asp?ZipName="+myForm.OwnerAddress[itemcnt-1].value);
	}
}

function funMouseOwnerAddress(itemcnt) {
	myForm.chkcnt.value=itemcnt;
	runServerScript("chkStoreAndSendZipChkOwnerAddr.asp?ZipName="+myForm.OwnerAddress[itemcnt-1].value);
}

function funSendDate(itemcnt) {
	<%If sys_City="南投縣" then%>
		if (event.keyCode==13||event.keyCode==9) {
			if(itemcnt<myForm.StoreAndSendSendDate.length){
				myForm.mailNumber[itemcnt-1].focus();
			}
		}
	<%else%>
		if (event.keyCode==13||event.keyCode==9) {
			if(itemcnt<myForm.StoreAndSendSendDate.length){
				myForm.item[itemcnt].focus();
			}
		}
	<%end if%>
}
function keyMailNumber(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.mailNumber.length){
			myForm.item[itemcnt].focus();
		}
	}
}

function funkeymove(itemcnt) {
	/*if(event.keyCode==40){
		if(myForm.StoreAndSendMailDate.length>itemcnt){
			myForm.StoreAndSendMailDate[itemcnt].focus();
		}
	}else if(event.keyCode==38){
		if(itemcnt>1){
			myForm.StoreAndSendMailDate[itemcnt-2].focus();
		}
	}*/
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.StoreAndSendMailDate.length){
			myForm.OwnerAddress[itemcnt-1].focus();
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
function funDefuSelt(){
	for(i=0;i<myForm.Sys_BackCause.length;i++){
		myForm.Sys_BackCause[i].selectedIndex=myForm.Sys_BackCauseMain.selectedIndex;
	}
}
function funNumber(){
	for(i=0;i<myForm.mailNumber.length;i++){
		myForm.mailNumber[i].value=myForm.Sys_Number.value;
	}
}
function funSelt(){
	var err=0;
	for(i=0;i<myForm.item.length;i++){
		if(myForm.item[i].value!=''){
			if(myForm.StoreAndSendSendDate[i].value==''||myForm.StoreAndSendSendDate[i].value.length<6){
				err=1;
				alert("第 "+(i+1)+" 行單退日期不可空白或格式錯誤!!");
				break;
			}
			if(myForm.ChkSend[i].value=='否'){
				if(myForm.EffectDate[i].value==''||myForm.EffectDate[i].value.length<6){
					err=1;
					alert("第 "+(i+1)+" 行不二次郵寄，送達日期不可空白或格式錯誤!!");
					break;
				}
			}
		}
	}

	if(err==0){
		myForm.DB_Selt.value="Selt";
		myForm.submit();
	}
}
function funBackCauseIndex(obj,strobj,itemcnt){
	var selectLen=eval("myForm."+strobj+"["+(itemcnt-1)+"]").length;
	chknumber(obj);
	if(obj.value!=''&&obj.value<selectLen){
		eval("myForm."+strobj+"["+(itemcnt-1)+"]").options[obj.value].selected=true;
	}else if(obj.value!=''){
		alert("超出範圍請重新填寫!!");
	}
	if (event.keyCode==13||event.keyCode==9) {
		myForm.mailNumber[itemcnt-1].focus();
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