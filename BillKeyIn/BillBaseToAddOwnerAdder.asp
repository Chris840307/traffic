<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/banner.asp"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>二次郵寄前註記</TITLE>
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
	Sys_OwnerAddress=Split(request("OwnerAddress"),",")
	Sys_OwnerZip=Split(request("OwnerZip"),",")
	Sys_ZipName=Split(request("Sys_ZipName"),",")
	Sys_Now=now
	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" and trim(Sys_OwnerAddress(i))<>"" then
			if Not ifnull(Sys_OwnerAddress(i)) then
				tmp_ZipID="":tmp_ZipName="":Sys_tempAddress=""
				If ifnull(Sys_OwnerZip(i)) Then
					Sys_tempAddress=replace(left(trim(Sys_OwnerAddress(i)),5),"臺","台")
					strSQL="select ZipID,ZipName from Zip where ZipName like '"&Sys_tempAddress&"%'"
					set rszip=conn.execute(strSQL)
					If Not rszip.eof Then
						tmp_ZipID=rszip("ZipID")
						tmp_ZipName=rszip("ZipName")
					end if
					rszip.close
				else
					tmp_ZipID=Sys_OwnerZip(i)
					tmp_ZipName=Sys_ZipName(i)
				end if

'				strSQL="Select BillTypeID from BillBase where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
'				set rsbill=conn.execute(strSQL)
'				If trim(rsbill("BillTypeID"))="1" Then
					strSQL="Update BillBaseDciReturn set DriverHomeZIP='"&trim(tmp_ZipID)&"',DriverHomeAddress='"&replace(trim(Sys_OwnerAddress(i)),trim(tmp_ZipName),"")&"',DriverCounty='"&left(trim(tmp_ZipName),3)&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
					conn.execute(strSQL)
'				else
'					strSQL="Update BillBaseDciReturn set OwnerZip='"&trim(tmp_ZipID)&"',OwnerAddress='"&replace(trim(Sys_OwnerAddress(i)),trim(tmp_ZipName),"")&"',OwnerCounty='"&left(trim(tmp_ZipName),3)&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
'					conn.execute(strSQL)
'				end if
'				rsbill.close
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
		<td height="27" bgcolor="#FFCC33" class="pagetitle"><strong>寄存送達證書 戶籍地址補正 </strong> <br><bR>
		透過此功能可修正寄存送達郵記之正確地址。 自動抓取監理站民眾領牌登錄之戶籍地址部份則不適用此功能<br>
		<br>
		<font Size="3"><b>作業流程 :</b> <b>1</b>. 查詢 正確戶籍地址 <img src="space.gif" width="29" height="8"> <b> 2. </b> 篩選出該批欲作戶籍補正資料
		 <img src="space.gif" width="39" height="8"><b>3</b>. 至 寄存送達證書 戶籍地址補正 功能進行補正 <br><img src="space.gif" width="79" height="8">
		 <b>4</b>. 經由 單退註記-寄存送達  逕行註記 <img src="space.gif" width="39" height="8"><b>5</b>. 經由 上傳監理站-單退 進行上傳，取得批號 <br>
		 
		 <img src="space.gif" width="81" height="8"><b>6</b>. 等待下載完成後以該批號列印 大宗清冊/函件 以及 送達證書
		</font>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						&nbsp;&nbsp;<input type="button" name="btnOK" value="匯入地址資料" onclick="funAddressSelt();"><br>
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
		<td height="26" bgcolor="#FFCC33">二次郵寄紀錄列表 ( 輸入完成按Enter可自動跳到下一格 )</td>
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
		if(isTable.rows.length>0){
			Cindex = isTable.rows(Rindex-1).cells.length;
		}else{
			Cindex=0;
		}
		if(Rindex==0||Cindex==1){
			nextRow = isTable.insertRow(Rindex);
			txtArea = nextRow.insertCell(0);
		}else{
			if(cunt==0){
				Cindex=0;
				isTable.rows(Rindex-1).deleteCell();
			}
			txtArea =isTable.rows(Rindex-1).insertCell(Cindex);
		}
		cunt++;
		//txt_nameStr = "item"+cunt;
		txtArea.innerHTML ="單號<input type=text name='item' size=6 class='btn1' onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;郵地區號<input type=text name='OwnerZip' size=2 class='btn1' onkeydown='funZip("+cunt+");' MaxLength=3>&nbsp;&nbsp;戶籍地址<input type=text name='OwnerAddress' size=35 class='btn1' onkeydown='funAddress("+cunt+");'>&nbsp;&nbsp;<%=CarName%><input type=text name='CarNo' size=5 class='btn1' readOnly><input type=Hidden name='Sys_ZipName' value=''>";
	}
}
function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}
function keyFunction(itemcnt) {
	myForm.item[itemcnt-1].value=myForm.item[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9||myForm.item[itemcnt-1].value>=9) {
		if (chkBillNo(itemcnt)){
			if (myForm.item[itemcnt-1].value!=''){
				myForm.chkcnt.value=itemcnt;
				runServerScript("chkStoreAndSendOwnerAdder.asp?BillNo="+myForm.item[itemcnt-1].value);
			}
		}else{
			alert("單號長度必須為9碼!!");
		}
	}
}

function funZip(itemcnt) {
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

function funZipName(itemcnt) {
	runServerScript("chkStoreAndSendZip.asp?Zip="+myForm.OwnerZip[itemcnt-1].value);
}

function funAddress(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		myForm.chkcnt.value=itemcnt;
		runServerScript("chkStoreAndSendZip.asp?ZipName="+myForm.OwnerAddress[itemcnt-1].value);
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

function funAddressSelt(){
	UrlStr="AddressSendStyle.asp";
	myForm.action=UrlStr;
	myForm.target="Address";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funSelt(){
	var err=0;
	for(i=0;i<myForm.item.length;i++){
		if(myForm.item[i].value!=''){
			if(myForm.OwnerAddress.value==''){
				err=1;
				alert("第 "+(i+1)+" 行戶藉地址不可空白!!");
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