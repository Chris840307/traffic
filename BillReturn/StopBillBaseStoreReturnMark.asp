<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>催繳-寄存送達</TITLE>
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
	Sys_mailNumber=Split(request("mailNumber"),",")
	Sys_mailStation=Split(request("mailStation"),",")
	EffectDate=Split(request("EffectDate"),",")
	Sys_Now=now

	if sys_City="台東縣" then
		sys_BatNo="S"&gInitDT(date)&right("00"&hour(now),2)&right("00"&minute(now),2)
	end if

	
	'jafe------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	'response.write request("Sys_JpgFile") &"aa"
	Sys_JpgFile=Split(request("Sys_JpgFile"),",")
	'jafe------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		

	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" and trim(Sys_BackDate(i))<>"" then

			Sys_Now=DateAdd("s",1,Sys_Now)
			'strSQL="Select MailReturnDate from BillMailHistory where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
			'set rs=conn.execute(strSQL)
			strSQL="Update StopBillMailHistory set ReturnResonID='"&trim(Sys_BackCause(i))&"',MailReturnDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",mailStation='"&trim(Sys_mailStation(i))&"',ReturnRecordMemberID="&Session("User_ID")&",ReturnReCordDate="&funGetDate((Sys_Now),1)&",UserMarkMemberID="&Session("User_ID")&",UserMarkDate="&funGetDate((Sys_Now),1)&",UserMarkResonID='"&trim(Sys_BackCause(i))&"',UserMarkReturnDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",MailTypeID=null ,StoreAndSendEffectDate="&funGetDate(gOutDT(EffectDate(i)),0)&" where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"

			conn.execute(strSQL)
			if trim(Sys_mailNumber(i))<>"" then
				strSQL="Update StopBillMailHistory set mailNumber='"&trim(Sys_mailNumber(i))&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
				conn.execute(strSQL)

			end if
			strSQL="Update Billbase set BillStatus=3 where ImageFIleNameB='"&trim(Ucase(Sys_BillNo(i)))&"'"  & " and Recordstateid=0"
			conn.execute(strSQL)
			'rs.close

			if sys_City="台東縣" then

				stopBackDate=gOutDT(EffectDate(i))

				'strSQL="update billbase set DealLineDate="&funGetDate(DateAdd("d",7,stopBackDate),0)&" where ImageFIleNameB='"&trim(Ucase(Sys_BillNo(i)))& "' and Recordstateid=0 and DealLineDate <= "&funGetDate(DateAdd("d",7,stopBackDate),0)
				'conn.execute(strSQL)

				'sys_BatNo="S"&gInitDT(date)&right("00"&hour(now),2)&right("00"&minute(now),2)

				strSQL="Update Billbase set Note=Note||'批號:"&sys_BatNo&"' where ImageFIleNameB='"&trim(Ucase(Sys_BillNo(i)))& "'" & " and Recordstateid=0"

				conn.execute(strSQL)

			end if


		end if
	next
	if sys_City="花蓮縣" or sys_City="台東縣" then
	'jafe----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		'固定位置
		fp="f:\\ScannerImport" 
        finDir="f:\\ScannerImport\\催繳\\"
        set fso=Server.CreateObject("Scripting.FileSystemObject")
        for i=0 to Ubound(Sys_JpgFile)
		Sys_BillNo(i)=trim(Sys_BillNo(i))
		Sys_JpgFile(i)=trim(Sys_JpgFile(i))
        next
        set fod=fso.GetFolder(fp)
        set fic=fod.Files
		'if fso.fileexists(fp)=true then 
			for i=0 to Ubound(Sys_JpgFile)
				if trim(Sys_BillNo(i))<>"" and trim(Sys_BackDate(i))<>""  and trim(Sys_JpgFile(i))<>"" then
					sSQL = "select BillAttatchImage_seq.nextval as SN from Dual"
					set oRST = Conn.execute(sSQL)
					if not oRST.EOF then
						sMaxSN = oRST("SN")
					end if
					oRST.close
					
					if fso.FolderExists(finDir & year(date)-1911) = false then
						fso.CreateFolder(finDir & year(date)-1911) 
					end if
					
					if fso.FolderExists(finDir & year(date)-1911 & "\\" & right("0" & month(date),2)) = false then
						fso.CreateFolder(finDir & year(date)-1911 & "\\" & right("0"&month(date),2))
					end if
					
					fDir=year(date)-1911 & "/" & right("0"&month(date),2) & "/"
				'	response.write Sys_JpgFile(i)
					FileDirAndName="/img/scan/催繳/" & fDir & Sys_BillNo(i) & "_" & replace(Sys_JpgFile(i),Sys_BillNo(i) & "_","")
					
					strInsert="insert into BillAttatchImage(SN,FileName,BillNo,TypeID,RecordMemberID,RecordDate,RecordStateID)" & _
							  " values("&sMaxSN&",'"&FileDirAndName&"','" & Sys_BillNo(i) & "','0','"& trim(session("User_ID")) &"',SYSDATE,0)"
					conn.execute(strInsert)
					mDir=finDir & year(date)-1911 & "\\" & right("0"&month(date),2) & "\\"
					'response.write Sys_JpgFile(i)

					fso.CopyFile fp & "\\" & Sys_JpgFile(i), mDir & Sys_BillNo(i) & "_" & replace(Sys_JpgFile(i),Sys_BillNo(i) & "_",""),true
					fso.DeleteFile(fp & "\\" & Sys_JpgFile(i))
				end if
			Next	
		'end if
	'jafe-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	end If 

	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	if sys_City="台東縣" then Response.Write "alert('批號："&sys_BatNo&"');"
	Response.write "</script>"
end if
%>
<form name=myForm method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle"><strong>催繳-寄存送達</strong><br>
		<!--	使用者可選擇是否需要先針對單退舉發單依據 舉發單 應到案處所 以及 舉發單位 進行分類,<br>後續在下方使用條碼刷入舉發單號時，
			系統會自動偵測該舉發單的<br>應到案處所 與 舉發單位 與使用者選取的分類條件是否相同. 
			-->
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
						預設退件原因&nbsp;
						<select name="Sys_BackCauseMain" class="btn1">
							<%strSQL="select ID,Content from DCICode where TypeID=7 and ID in('5','6','7','T')"
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
						&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuSelt();">
						&nbsp;<font size="2">非必要選項,也可以由下方設定各舉發單不同的單退原因</font>
						<br>
						預設單退日期&nbsp;<input name="Sys_BackDateMain" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BackDateMain');">
						&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuDate();">
						&nbsp;<font size="2">非必要選項,也可以由下方設定各舉發單不同的單退日期</font>
						<br>
						預設送達日期&nbsp;<input name="Sys_EffectDateMain" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_EffectDateMain');">
						&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funEffectDate();">
						&nbsp;<font size="2">非必要選項,也可以由下方設定各舉發單不同的單退日期</font>
						<br>
						預設大宗掛號&nbsp;
						<input name="Sys_Number" type="text" class="btn1" size="10" maxlength="15">
						&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funNumber();"><font size="2">&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的大宗掛號</font>
						<br>
						預設郵局為&nbsp;
						<input name="Sys_mailStation" type="text" class="btn1" size="10" maxlength="15">
						<input type="button" name="btnDefu" value="確定" onclick="funStation();">
						<font size="2">
							非必要選項,也可以由下方設定各舉發單不同的郵局
						</font>
						<br>
						&nbsp;&nbsp;<input type="button" name="btnOK" value="確定存檔" onclick="funSelt();">
						<img src="space.gif" width="9" height="8">
						<input type="button" name="insert" value="再多30筆" onClick="insertRow(fmyTable)">
						<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">寄存送達紀錄列表&nbsp;
		<%if sys_City="花蓮縣" or sys_City="台東縣" then%>
			<input type="button" name="btnLoad" value="掃描匯入" onclick="funScannerImport();"></td>
		<%end if%>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<Div style="overflow:auto;width:100%;height:330px;background:#FFFFFF">
				<table id='fmyTable' width='1200' border='0' bgcolor='#FFFFFF'>
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
<input type="Hidden" name="SaveError" value="">
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
		txtArea.innerHTML =cnt_num+"<Span ID='UrlJpg'>催繳單號</span>&nbsp;<input type=text name='item' size=16 class='btn1' onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;原因<select name='Sys_BackCause' class='btn1' onkeydown='keyBackCause("+cunt+");'><%=seltarr%></select>&nbsp;&nbsp;單退日<input type=text name='Sys_BackDate' size=4 class='btn1' onkeyup='funkeyChk(this);' onkeydown='keyBackDate("+cunt+");'>&nbsp;&nbsp;送達日<input type=text name='EffectDate' size=4 class='btn1' onkeyup='funkeyChk(this);' onkeydown='keyEffectDate("+cunt+");'>&nbsp;&nbsp;大宗碼<input type=text name='mailNumber' size=3 class='btn1' onkeydown='keyMailNumber("+cunt+");'>&nbsp;&nbsp;<%=CarName%><input type=text name='CarNo' size=4 class='btn1' readOnly>&nbsp;&nbsp;郵局<input type=text name='mailStation' size=4 class='btn1'><input type=Hidden name='DeallineDate'><input type='Hidden' name='Sys_JpgFile' value=''>";
	}
}
function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}
function keyFunction(itemcnt) {
	myForm.item[itemcnt-1].value=myForm.item[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9||myForm.item[itemcnt-1].length>=16) {

		if (myForm.item[itemcnt-1].value!=''){
			myForm.chkcnt.value=itemcnt;

			myForm.item[itemcnt-1].value=("000000000000000"+myForm.item[itemcnt-1].value).substr(("000000000000000"+myForm.item[itemcnt-1].value).length-16,16);
				
			runServerScript("StopchkAcceptBillNo.asp?BillNo="+myForm.item[itemcnt-1].value);
			myForm.Sys_BackDate[itemcnt-1].focus();
		}
	}
}

function keyBackCause(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		myForm.Sys_BackDate[itemcnt-1].focus();
	}
}

function keyMailNumber(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.mailNumber.length){
			myForm.item[itemcnt].focus();
		}
	}
}

function keyBackDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		myForm.EffectDate[itemcnt-1].focus();
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
function funDefuDate(){
	for(i=0;i<myForm.Sys_BackDate.length;i++){
		myForm.Sys_BackDate[i].value=myForm.Sys_BackDateMain.value;
	}
}
function funEffectDate(){
	for(i=0;i<myForm.EffectDate.length;i++){
		myForm.EffectDate[i].value=myForm.Sys_EffectDateMain.value;
	}
}
function keyEffectDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
	
		runServerScript("chkStopMailDate.asp?EffectDate="+myForm.EffectDate[itemcnt-1].value);
		myForm.mailNumber[itemcnt-1].focus();
		if(myForm.EffectDate[itemcnt-1].value!=''&&myForm.DeallineDate[itemcnt-1].value!=''){
			if(eval(myForm.EffectDate[itemcnt-1].value)>eval(myForm.DeallineDate[itemcnt-1].value)){
				alert("該筆資料單退日期大於繳費日期"+myForm.DeallineDate[itemcnt-1].value);
			}
		}
	}
}

function funStation(){
	for(i=0;i<myForm.mailStation.length;i++){
		myForm.mailStation[i].value=myForm.Sys_mailStation.value;
	}
}

// jafe--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function funScannerImport(){
		myForm.DB_Selt.value="Import";
		myForm.submit();
}
// jafe____________________________________________________________________________________________________________

<%
if sys_City="花蓮縣" or sys_City="台東縣" then
	'jafe-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	if trim(request("DB_Selt"))="Import" then
	'固定位置
	fp="f:\\ScannerImport" 

			set fso=Server.CreateObject("Scripting.FileSystemObject")

			set fod=fso.GetFolder(fp)
			set fic=fod.Files
	   
		i=-1
		response.write "myForm.SaveError.value="""";"
		tmpBillno=""

		For Each fil In fic
			if UCase(fso.GetExtensionName(fil.Name)) ="JPG" Or UCase(fso.GetExtensionName(fil.Name)) ="JPEG" then
				i=i+1
				if i >=90 then
					if (i) mod 30 =0 then response.write "insertRow(fmyTable);"
				end if
				Sys_tmpBillNo=Split(fil.Name,"_")

				tmpBillno=Sys_tmpBillNo(0)
				if not ifnull(Sys_tmpBillNo(0)) then tmpBillno=right("0000000000000000"&Sys_tmpBillNo(0),16)

				response.write "myForm.item[" & i & "].value='" & tmpBillno &"';" 
				response.write "myForm.chkcnt.value="&(i+1)&";"
				response.write "UrlJpg[" & i & "].innerHTML=""<a href='\\ScannerImport\\"&fil.Name&"' TARGET ='_blank'>催繳單號</a>"";" 
				response.write "myForm.Sys_JpgFile[" & i & "].value='" & fil.Name &"';" 

				strSQL = "select a.CarNo,a.DeallineDate,a.billFillDate,b.DCIStationName,b.StationID,c.UnitID,c.UnitName from BillBase a,Station b,UnitInfo c where a.BillUnitID=c.UnitID(+) and a.RecordStateID=0 and a.ImageFileNameB='"&trim(tmpBillno)&"'"

				set rscnt=conn.execute(strSQL)
				if Not rscnt.eof then
					response.write "myForm.Sys_BackDate[" & i & "].value='"& gInitDT(DateAdd("d",getBillReturnDate,rscnt("billFillDate")))&"';"
					response.write "myForm.DeallineDate[" & i & "].value='"& gInitDT(trim(rscnt("DeallineDate"))) &"';"
					response.write "myForm.CarNo[" & i & "].value='"&trim(rscnt("CarNo"))&"';"
				end if 
				rscnt.close

			end if
		Next
		
	end if
	'jafe-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
end if
%>

</script>