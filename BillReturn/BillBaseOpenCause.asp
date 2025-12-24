<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>公示送達-原因註記</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<%

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

if trim(request("DB_Selt"))="Selt" then
	Sys_BillNo=Split(Ucase(request("item"))&" ",",")
	Sys_BackCause=Split(request("Sys_BackCause")&" ",",")
	Sys_BackDate=Split(request("Sys_BackDate")&" ",",")
	Sys_mailNumber=Split(request("mailNumber")&" ",",")
	Sys_Now=DateAdd("n", -5, now)

	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" and trim(Sys_BackDate(i))<>"" then
			Sys_Now=DateAdd("s",1,Sys_Now)
			strSQL="Update BillMailHistory set OpenGovStationID='"&request("Sys_Station")&"',OpenGovResonID='"&trim(Sys_BackCause(i))&"',OpenGovMailReturnDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",OpenGovRecordMemberID="&Session("User_ID")&",OpenGovReCordDate="&funGetDate((Sys_Now),1)&",UserMarkMemberID="&Session("User_ID")&",UserMarkDate="&funGetDate((Sys_Now),1)&",UserMarkResonID='"&trim(Sys_BackCause(i))&"',UserMarkReturnDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",MailTypeID=6 where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
			conn.execute(strSQL)
			strSQL="Update Billbase set BillStatus=3 where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
			conn.execute(strSQL)
			if trim(Sys_mailNumber(i))<>"" then
				strSQL="Update BillMailHistory set mailNumber='"&trim(Sys_mailNumber(i))&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
				conn.execute(strSQL)
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
		<td height="27" bgcolor="#FFCC33" class="pagetitle"><strong>公示送達-原因註記</strong><br>
			使用者可選擇是否需要先針對單退舉發單依據 舉發單 應到案處所 以及 舉發單位 進行分類,<br>後續在下方使用條碼刷入舉發單號時，
			系統會自動偵測該舉發單的<br>應到案處所 與 舉發單位 與使用者選取的分類條件是否相同. 
						
			</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						
						分類條件&nbsp;&nbsp;<font size="2">(非必要選取)</font><br>
						 應到案處所&nbsp;
						<select name="Sys_Station" class="btn1">
							<option value="">請選取</option>
							<%strSQL="select StationID,DCIStationID,DCIStationName from Station"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("StationID")&""""
								response.write ">"&rs1("DCIStationID")&","&rs1("DCIStationName")&"</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
						舉發單位&nbsp;
						<select name="Sys_UnitID" class="btn1">
							<option value="">請選取</option>
							<%strSQL="select UnitID,UnitName from UnitInfo"
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								response.write "<option value="""&rs1("UnitID")&""""
								response.write ">"&rs1("UnitID")&","&rs1("UnitName")&"</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
						<br>						
						預設退件原因統一為&nbsp;
						<select name="Sys_BackCauseMain" class="btn1">
							<%strSQL="select ID,Content from DCICode where TypeID=7 and ID in('1','2','3','4','V','8')"
							set rs1=conn.execute(strSQL)
							seltarr=""
							while Not rs1.eof
								response.write "<option value="""&rs1("ID")&""""
								response.write ">"&rs1("Content")&"</option>"

								seltarr=seltarr&"<option value='"&rs1("ID")&"'>"&rs1("Content")&"</option>"
								rs1.movenext
							wend
							rs1.close%>
						</select>
						&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuSelt();"><font size="2">&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的單退原因</font>
						<br>
						預設單退日期統一為&nbsp;
						<input name="Sys_BackDateMain" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BackDateMain');">
						&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuDate();"><font size="2">&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的單退日期</font>
						<br>
						預設大宗掛號統一為&nbsp;
						<input name="Sys_Number" type="text" class="btn1" size="10" maxlength="15">
						&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funNumber();"><font size="2">&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的大宗掛號</font>
						<br>
						&nbsp;&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funSelt();">
						<img src="space.gif" width="9" height="8">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="insert" value="再多30筆" onClick="insertRow(fmyTable)">
						<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">單退紀錄列表</td>
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
			<input type="button" name="btnOK1" value="確定存檔" onclick="funSelt();">
			<%
				Response.Write "<input type=""button"" name=""insert2"" value=""再多30筆"" onClick=""insertRow(fmyTable)"">"
			%>
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="">
<input type="Hidden" name="chkcnt" value="">
</form>

<form name="upForm" method="post">

	<input type="Hidden" name="item" value="">
	<input type="Hidden" name="Sys_BackCause" value="">
	<input type="Hidden" name="Sys_BackDate" value="">
	<input type="Hidden" name="mailNumber" value="">

	<input type="Hidden" name="DB_Selt" value="">
</form>

</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var cunt=0;
function insertRow(isTable){
	<%
		cnt=29
	%>
	var cnt=<%=cnt%>;
	
	for(i=0;i<=cnt;i++){
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
		txtArea.innerHTML =cnt_num+".&nbsp;單號<input type=text name='item' size=10 class='btn1' onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;原因<select name='Sys_BackCause' class='btn1'><%=seltarr%></select>&nbsp;&nbsp;單退日期<input type=text name='Sys_BackDate' size=10 class='btn1' onkeyup='funkeyChk(this);' onkeydown='keyBackDate("+cunt+");' maxlength='7'>&nbsp;&nbsp;大宗掛號碼<input type=text name='mailNumber' size=10 class='btn1' onkeydown='keyMailNumber("+cunt+");' maxlength='20'>&nbsp;&nbsp;車號<input type=text name='CarNo' size=10 class='btn1' readOnly>";
	}
}
function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}
function keyFunction(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if (chkBillNo(itemcnt)){
			if (myForm.item[itemcnt-1].value!=''){
				myForm.chkcnt.value=itemcnt;
				runServerScript("chkBillNo.asp?BillNo="+myForm.item[itemcnt-1].value);
			}
		}else{
			alert("單號長度必須為9碼!!");
		}
	}
}

function keyBackDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		myForm.mailNumber[itemcnt-1].focus();
	}
}

function keyMailNumber(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		if(itemcnt<myForm.mailNumber.length){
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
function funDefuSelt(){
	for(i=0;i<myForm.Sys_BackCause.length;i++){
		myForm.Sys_BackCause[i].selectedIndex=myForm.Sys_BackCauseMain.selectedIndex;
	}
}

function funDefuDate(){
	for(i=0;i<myForm.Sys_BackDate.length;i++){
		myForm.Sys_BackDate[i].value=myForm.Sys_BackDateMain.value;
	}
}

function funNumber(){
	for(i=0;i<myForm.mailNumber.length;i++){
		myForm.mailNumber[i].value=myForm.Sys_Number.value;
	}
}
function funSelt(){
	var err=0;

	var item='';
	var Sys_BackCause='';
	var Sys_BackDate='';
	var mailNumber='';

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
		for(i=0;i<myForm.item.length;i++){
			if(myForm.item[i].value!=''){
				if(item!=''){
					item=item+',';
					Sys_BackCause=Sys_BackCause+',';
					Sys_BackDate=Sys_BackDate+',';
					mailNumber=mailNumber+',';
				}
				item=item + myForm.item[i].value;
				Sys_BackCause=Sys_BackCause + myForm.Sys_BackCause[i].value;
				Sys_BackDate=Sys_BackDate + myForm.Sys_BackDate[i].value;
				mailNumber=mailNumber + myForm.mailNumber[i].value;
			}
		}

		upForm.item.value=item;
		upForm.Sys_BackCause.value=Sys_BackCause;
		upForm.Sys_BackDate.value=Sys_BackDate;
		upForm.mailNumber.value=mailNumber;

		upForm.DB_Selt.value="Selt";
		upForm.submit();
	}
}

<%
	Response.Write "for(j=0;j<=3;j++){insertRow(fmyTable);}"
%>
</script>