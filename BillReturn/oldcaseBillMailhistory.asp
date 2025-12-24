<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>舊資料送達註記 </TITLE>
</HEAD>
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strfiletitle="select value from Apconfigure where id=100"
set rsfiletitle=conn.execute(strfiletitle)
filetitle=trim(rsfiletitle("value"))
rsfiletitle.close

if trim(request("DB_Selt"))="Selt" then
	Sys_BillNo=Split(UCase(request("BillNo")),",")
	Sys_CarNo=Split(UCase(request("CarNo")),",")
	Sys_BackCause=Split(request("Sys_BackCause"),",")
	Sys_BackDate=Split(request("Sys_BackDate"),",")
	Sys_DeliverNo=Split(request("Sys_DeliverNo"),",")
	Sys_Now=now

	strSQL="select NVL(Max(FileNameSeq),6) as FileNameSeq from OldCaseBillMailHistory where RecordDate between TO_DATE('"&date()&" :00:00:00','YYYY/MM/DD HH24:MI:SS') and TO_DATE('"&date()&" :23:59:59','YYYY/MM/DD HH24:MI:SS')"
	set rs=conn.execute(strSQL)
	Sys_FileNameSeq=cdbl(rs("FileNameSeq"))+1
	If Sys_FileNameSeq<10 Then
		Sys_FileSeq=Sys_FileNameSeq
	else
		Sys_FileSeq=Chr((Sys_FileNameSeq+55))
	End if
	Sys_FileName=filetitle&gInitDT(date()) & Sys_FileSeq &".X.F"
	rs.close

	strSQL="select LoginID from MemberData where MemberID="&Session("User_ID")
	set rsmem=conn.execute(strSQL)
	Sys_LoginID=trim(rsmem("LoginID"))
	rsmem.close

	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" and trim(Sys_BackDate(i))<>"" then
			strSQL="select count(*) cmt from OldCaseBillMailHistory where BillNo='"&trim(Sys_BillNo(i))&"' and CarNo='"&trim(Sys_CarNo(i))&"'"

			set rscmt=conn.execute(strSQL)
			cmt=cdbl(rscmt("cmt"))
			rscmt.close

			If cmt = 0 Then
				strSQL="Insert into OldCaseBillMailHistory (SninDCIFile,BillNo,CarNo,ReaSonID,LoginID,DOCNumber,ProcessDate,RecordMemberID,RecordDate,FileName,FileNameSeq) values('"&right("00000"&(i+1),5)&"','"&trim(Sys_BillNo(i))&"','"&UCase(trim(Sys_CarNo(i)))&"','"&trim(Sys_BackCause(i))&"','"&Sys_LoginID&"','"&trim(Sys_DeliverNo(i))&"',"&funGetDate(gOutDT(Sys_BackDate(i)),0)&","&Session("User_ID")&","&funGetDate((Sys_Now),1)&",'"&Sys_FileName&"','"&Sys_FileNameSeq&"')"

				conn.execute(strSQL)

			else
				strSQL="Update OldCaseBillMailHistory set SninDCIFile='"&right("00000"&(i+1),5)&"',ReaSonID='"&trim(Sys_BackCause(i))&"',DOCNumber='"&trim(Sys_DeliverNo(i))&"',FileName='"&Sys_FileName&"',FileNameSeq='"&Sys_FileNameSeq&"',RecordDate="&funGetDate((Sys_Now),1)&" where BillNo='"&trim(Sys_BillNo(i))&"'"

				conn.execute(strSQL)
			End if
		end if
	next
	Dim objFSO, objTextFile
	Dim sRead, sReadLine, sReadAll
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.CreateTextFile(Server.mappath("\UpProcess\")&"\"&Sys_FileName, True)
	strSQL="select * from OldCaseBillMailHistory where FileName like '"&Sys_FileName&"%' order by SninDCIFile"
	set rstxt=conn.execute(strSQL)
	While not rstxt.eof
		strInput=""
		strInput=strInput&left(left(trim(rstxt("SninDCIFile")),5)&"               ",6)
		strInput=strInput&left(left(trim(rstxt("CarNo")),8)&"               ",9)
		strInput=strInput&left(left(trim(rstxt("BillNo")),9)&"               ",10)
		strInput=strInput&left(left(trim(rstxt("ReaSonID")),3)&"               ",4)
		strInput=strInput&left("               ",11)
		strInput=strInput&left("0"&"               ",2)
		strInput=strInput&left("X"&"               ",2)
		strInput=strInput&left(left(trim(rstxt("LoginID")),6)&"               ",7)
		strInput=strInput&left(left(trim(rstxt("DOCNumber")),9)&"               ",10)
		strInput=strInput&left(right("0"&gInitDT(trim(rstxt("ProcessDate"))),7)&"               ",8)
		objTextFile.WriteLine(strInput)
		rstxt.movenext
	Wend
	objTextFile.Close
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
%>
<BODY>
<form name=myForm method="post">
<table border="0" width="100%" bgcolor="#ffffff">
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr><td>
					<table>
					<tr>
						<td><font size="5"><b>舊系統資料送達註記 </b></font> <br><br>	</td>
						<td>( 新系統資料請勿在此註記 ) </td>
					</tr>
					<tr>
						<td>預設收受/送達原因統一為&nbsp;</td>
						<td>
							<select name="Sys_BackCauseMain" class="btn1">
								<option value="1">簽收</option>
								<option value="F">寄存郵局</option>
								<option value="D">公示送達</option>
								<option value="Y">徹消送達</option>
							</select><%
							seltarr="<option value='1'>簽收</option><option value='F'>寄存郵局</option><option value='D'>公示送達</option><option value='Y'>徹消送達</option>"
							%>
						</td>
						<td><input type="button" name="btnDefu" value="確定" onclick="funBackCauseMain();"></td>
						<td><font size="2">非必要選項,也可以由下方設定各舉發單不同的收受原因</font></td>
					</tr><tr>
						<td>預設送達/公示日期&nbsp;</td>
						<td>
							<input name="Sys_BackDateMain" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BackDateMain');">
						</td>
						<td><input type="button" name="btnDefu" value="確定" onclick="funDefuDate();"></td>
						<td><font size="2">非必要選項,也可以由下方設定各舉發單不同的收受日期</font></td>
					</tr><tr>
						<td><br><input type="button" name="btnOK" value="確定存檔" onclick="funSelt();"></td>
						<td><input type="button" name="insert" value="再多30筆" onClick="insertRow(fmyTable)"></td>
					</tr>
				</table>
				</td></tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">公示/送達 紀錄列表</td>
	</tr>
	<tr>
		<td height="26"> ＊以下資料皆為必填欄位　＊每日交通隊與分局加起來<b>最多可處理20次</b>　＊目前使用Ｘ窗口代號上傳到監理站</td>
	</tr>	
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<table id='fmyTable' width="110%" border='0' bgcolor='#FFFFFF'>
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
		txtArea.innerHTML ="單號&nbsp;<input type=text name='BillNo' size=10 class='btn1' onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;車號&nbsp;<input type=text name='CarNo' size=8 class='btn1' onkeydown='funCarNo("+cunt+");'>&nbsp;&nbsp;原因&nbsp;<select name='Sys_BackCause' class='btn1' onkeydown='funBackCause("+cunt+");'><%=seltarr%></select>&nbsp;&nbsp;簽收 / 寄存 / 公示日期&nbsp;<input type=text name='Sys_BackDate' size=7 class='btn1' onkeydown='funBackDate("+cunt+");'>&nbsp;&nbsp;文號&nbsp;<input type=text name='Sys_DeliverNo' size=10 class='btn1' onkeydown='funDeliverNo("+cunt+");'>";
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

function keyFunction(itemcnt) {
	myForm.BillNo[itemcnt-1].value=myForm.BillNo[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9||myForm.BillNo[itemcnt-1].value.length>=9) {
		if (myForm.BillNo[itemcnt-1].value.length==9){
			myForm.chkcnt.value=itemcnt;
			runServerScript("chkOldBillNo.asp?BillNo="+myForm.BillNo[itemcnt-1].value);
		}else{
			alert("單號長度必須為9碼!!");
		}
	}
}
function funCarNo(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.CarNo.length){
			myForm.Sys_BackCause[itemcnt-1].focus();
		}
	}
}
function funBackCause(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(myForm.Sys_BackCause[itemcnt-1].selectedIndex<2){
			myForm.Sys_DeliverNo[itemcnt-1].value=myForm.BillNo[itemcnt-1].value;
		}
		if(itemcnt<myForm.Sys_BackCause.length){
			myForm.Sys_BackDate[itemcnt-1].focus();
		}
	}
}
function funBackDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.Sys_BackDate.length){
			myForm.Sys_DeliverNo[itemcnt-1].focus();
		}
	}
}
function funDeliverNo(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.Sys_DeliverNo.length){
			myForm.BillNo[itemcnt].focus();
		}
	}
}

function funBackCauseMain(){
	for(i=0;i<myForm.Sys_BackCause.length;i++){
		myForm.Sys_BackCause[i].selectedIndex=myForm.Sys_BackCauseMain.selectedIndex;
		if(myForm.Sys_BackCauseMain.selectedIndex<2){
			myForm.Sys_DeliverNo[i].value=myForm.BillNo[i].value;
		}
	}
}

function funDefuDate(){
	for(i=0;i<myForm.Sys_BackDate.length;i++){
		myForm.Sys_BackDate[i].value=myForm.Sys_BackDateMain.value;
	}
}

function funSelt(){
	var err=0;
	for(i=0;i<myForm.BillNo.length;i++){
		if(myForm.BillNo[i].value!=''){
			if(myForm.Sys_BackDate[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 行公示/送達日期不可空白!!");
				break;
			}else if(myForm.CarNo[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 行車號不可空白!!");
				break;
			}else if(myForm.Sys_BackCause[i].value=='D'){
				if(myForm.Sys_DeliverNo[i].value==''){
					err=1;
					alert("第 "+(i+1)+" 行送達文號不可空白!!");
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

for(j=0;j<=3;j++){
	insertRow(fmyTable);
}

</script>