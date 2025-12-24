<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>催繳 - 收受註記</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
if sys_City="花蓮縣" then
	CarName="車號"
else
	CarName="車號"
end if

if trim(request("DB_Selt"))="Selt" then
	Sys_BillNo=Split(request("item"),",")
	Sys_BackCause=Split(request("Sys_BackCause"),",")
	Sys_BackDate=Split(request("Sys_BackDate"),",")
	Sys_mailNumber=Split(request("mailNumber"),",")
	Sys_mailStation=Split(request("mailStation"),",")
	Sys_signman=Split(request("signman"),",")
	Sys_Now=now

	if sys_City="台東縣" then
		sys_BatNo="A"&gInitDT(date)&right("00"&hour(now),2)&right("00"&minute(now),2)
	end if

	'jafe------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	'response.write request("Sys_JpgFile") &"aa"
	Sys_JpgFile=Split(request("Sys_JpgFile"),",")
	'jafe------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		
	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" and trim(Sys_BackDate(i))<>"" then
			
			Sys_Now=DateAdd("s",1,Sys_Now)
			strSQL="Update StopBillMailHistory set ReturnResonID='"&trim(Sys_BackCause(i))&"',MailReturnDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",ReturnRecordMemberID="&Session("User_ID")&",ReturnReCordDate="&funGetDate((Sys_Now),1)&",UserMarkMemberID="&Session("User_ID")&",UserMarkDate="&funGetDate((Sys_Now),1)&",UserMarkResonID='"&trim(Sys_BackCause(i))&"',UserMarkReturnDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",mailStation='"&trim(Sys_mailStation(i))&"',signman='"&trim(Sys_signman(i))&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
			conn.execute(strSQL)
			
			if trim(Sys_mailNumber(i))<>"" then
				strSQL="Update StopBillMailHistory set mailNumber='"&trim(Sys_mailNumber(i))&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
				conn.execute(strSQL)
			end if
			strSQL="Update Billbase set BillStatus=7 where ImageFIleNameB='"&trim(Ucase(Sys_BillNo(i)))& "'" & " and Recordstateid=0"
			conn.execute(strSQL)

			if sys_City="台東縣" then

				stopBackDate=gOutDT(Sys_BackDate(i))

				'strSQL="update billbase set DealLineDate="&funGetDate(DateAdd("d",7,stopBackDate),0)&" where ImageFIleNameB='"&trim(Ucase(Sys_BillNo(i)))& "' and Recordstateid=0 and DealLineDate <= "&funGetDate(DateAdd("d",7,stopBackDate),0)

				'conn.execute(strSQL)

				'sys_BatNo="A"&gInitDT(date)&right("00"&hour(now),2)&right("00"&minute(now),2)

				strSQL="Update Billbase set Note=Note||'批號:"&sys_BatNo&"' where ImageFIleNameB='"&trim(Ucase(Sys_BillNo(i)))& "'" & " and Recordstateid=0"

				conn.execute(strSQL)

			end if

		end if
	next
	if sys_City="花蓮縣" or sys_City="台東縣" then
		If Ubound(Sys_JpgFile)>=0 Then
			'jafe----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
			'固定位置
			fp="F:\\ScannerImport\\A000" 
			finDir="F:\\ScannerImport\\催繳\\"
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
		end if
	'jafe-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	end If 

	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	if sys_City="台東縣" then Response.Write "alert('批號："&sys_BatNo&"');"
	Response.write "</script>"
end if
%>
<BODY>
<form name=myForm method="post">
<table border="0" width="110%" bgcolor="#ffffff">
	<!--<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle"><strong>催繳-收受註記</strong></td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						預設收受原因統一為&nbsp;
						<select name="Sys_BackCauseMain_Bak" class="btn1">
							<%strSQL="select ID,Content from DCICode where TypeID=7 and ID in('A','B','C','D')"
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
						收受日期&nbsp;<input name="Sys_BackDateMain_Bak" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BackDateMain_Bak');">
						<br>
						整批案件註記收受
						&nbsp;&nbsp;&nbsp;&nbsp;
						郵寄日期&nbsp;<input name="Sys_SendDate_Bak1" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_SendDate_Bak1');">&nbsp;
						～
						&nbsp;<input name="Sys_SendDate_Bak2" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_SendDate_Bak2');">&nbsp;&nbsp;
						<input type="button" name="btnOK" value="確定" onclick="funBatSelt();">
					</td>
				</tr>
			</table>
		</td>
	</tr>-->
	<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle"><strong>催繳-收受註記</strong><br>
			<!--使用者可選擇是否需要先針對收受舉發單依據 舉發單 應到案處所 以及 舉發單位 進行分類,<br>後續在下方使用條碼刷入舉發單號時，
			系統會自動偵測該舉發單的應到案處所 與 舉發單位 與使用者選取的分類條件是否相同. -->
						
			</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr><td>
					<table>
					<tr>
						<td colspan=3><b>分類條件&nbsp;&nbsp;<font size="2">(非必要選取)</b></font></td>
					</tr><tr>
						<td>應到案處所&nbsp;</td>
						<td colspan=2>
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
						舉發單位&nbsp;</td>
						<td colspan=2>
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
						</td>						
					</tr>
					<tr>
						<td><b>整批統一設定資訊</b>	</td>
					</tr>
					<tr>
						<td>預設收受原因為&nbsp;</td>
						<td>
							<select name="Sys_BackCauseMain" class="btn1">
								<%strSQL="select ID,Content from DCICode where TypeID=7 and ID in ('A','B','C')"
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
						</td>
						<td><input type="button" name="btnDefu" value="確定" onclick="funDefuSelt();"></td>
						<td><font size="2">非必要選項,也可以由下方設定各舉發單不同的收受原因</font></td>
					</tr><tr>
						<td>預設收受日期&nbsp;</td>
						<td>
							<input name="Sys_BackDateMain" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BackDateMain');">
						</td>
						<td><input type="button" name="btnDefu" value="確定" onclick="funDefuDate();"></td>
						<td><font size="2">非必要選項,也可以由下方設定各舉發單不同的收受日期</font></td>
					</tr><tr>
						<td>預設大宗掛號為&nbsp;</td>
						<td><input name="Sys_Number" type="text" class="btn1" size="10" maxlength="15"></td>
						<td><input type="button" name="btnDefu" value="確定" onclick="funNumber();"></td>
						<td><font size="2">非必要選項,也可以由下方設定各舉發單不同的大宗掛號</font></td>
					</tr><tr>
						<td>預設郵局為&nbsp;</td>
						<td><input name="Sys_mailStation" type="text" class="btn1" size="10" maxlength="15"></td>
						<td><input type="button" name="btnDefu" value="確定" onclick="funStation();"></td>
						<td><font size="2">非必要選項,也可以由下方設定各舉發單不同的郵局</font></td>
					</tr><tr>
						<td><input type="button" name="btnOK" value="確定存檔" onclick="funSelt();"></td>
						<td><input type="button" name="insert" value="再多30筆" onClick="insertRow(fmyTable)"></td>
					</tr>
				</table>
				</td></tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">收受 紀錄列表&nbsp;
			<%if sys_City="花蓮縣" or sys_City="台東縣" then%>
				<input type="button" name="btnLoad" value="掃描匯入" onclick="funScannerImport();"></td>
			<%end if%>
			<br>
			<strong>未上傳前如果註記錯誤存檔時，可以再註記一次，蓋掉原本錯誤的紀錄。 </strong>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<Div style="overflow:auto;width:100%;height:330px;background:#FFFFFF">
				<table id='fmyTable' width="110%" border='0' bgcolor='#FFFFFF'>
					<tr bgcolor="#ffffff">
						<td align='center' bgcolor="#ffffff" nowrap>目前無新增項目 <b>( 掛號碼 / 郵局 / 代收人 為 非必填項目 )</b></td>
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
		txtArea.innerHTML =cnt_num+"<Span ID='UrlJpg'>催繳單號</span>&nbsp;<input type=text name='item' size=16 class='btn1' onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;原因<select name='Sys_BackCause' class='btn1'><%=seltarr%></select>&nbsp;&nbsp;收受日<input type=text name='Sys_BackDate' size=4 class='btn1' onkeydown='funBackDate("+cunt+");'>&nbsp;&nbsp;<%=CarName%><input type=text name='CarNo' size=5 class='btn1' ReadOnly>&nbsp;&nbsp;應到案日<input type=text name='DeallineDate' size=4 class='btn1' ReadOnly>&nbsp;&nbsp;掛號碼<input type=text name='mailNumber' size=4 class='btn1' onkeydown='funmailNumber("+cunt+");'>&nbsp;&nbsp;郵局<input type=text name='mailStation' size=4 class='btn1'>&nbsp;&nbsp;代收人<input type=text name='signman' size=4 class='btn1'><input type='Hidden' name='Sys_JpgFile' value=''>";
	}
}

function funBackDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.Sys_BackDate.length){
			myForm.mailNumber[itemcnt-1].focus();
		}
		if(myForm.Sys_BackDate[itemcnt-1].value!=''&&myForm.DeallineDate[itemcnt-1].value!=''){
			if(eval(myForm.Sys_BackDate[itemcnt-1].value)>eval(myForm.DeallineDate[itemcnt-1].value)){
				alert("該筆資料收受日期大於繳費日期");
			}
		}
	}
}
function keyFunction(itemcnt) {
	myForm.item[itemcnt-1].value=myForm.item[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9||myForm.item[itemcnt-1].length>=16) {

			if (myForm.item[itemcnt-1].value!=''){
				myForm.chkcnt.value=itemcnt;

				myForm.item[itemcnt-1].value=("000000000000000"+myForm.item[itemcnt-1].value).substr(("000000000000000"+myForm.item[itemcnt-1].value).length-16,16);

				runServerScript("StopchkAcceptBillNo.asp?BillNo="+myForm.item[itemcnt-1].value);
			}

	}
}
function funmailNumber(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
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
function funNumber(){
	for(i=0;i<myForm.mailNumber.length;i++){
		myForm.mailNumber[i].value=myForm.Sys_Number.value;
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
				alert("第 "+(i+1)+" 行收受日期不可空白!!");
				break;
			}
		}
	}
	if(err==0){
		myForm.DB_Selt.value="Selt";
		myForm.submit();
	}
}

function funBatSelt(){
	var err=0;
	if(myForm.Sys_BackDateMain_Bak.value==''){
		err=1;
		alert("整批收受日期不可空白!!");
	}else if(myForm.Sys_SendDate_Bak1.value==''){
		err=1;
		alert("整批郵寄日期不可空白!!");
	}else if(myForm.Sys_SendDate_Bak2.value==''){
		err=1;
		alert("整批郵寄日期不可空白!!");
	}
	if(err==0){
		myForm.DB_Selt.value="SeltBat";
		myForm.submit();
	}
}

// jafe--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function funScannerImport(){
		myForm.DB_Selt.value="Import";
		myForm.submit();
}
// jafe____________________________________________________________________________________________________________


for(j=0;j<=3;j++){
	insertRow(fmyTable);
}
function funDefuDate(){
	for(i=0;i<myForm.Sys_BackDate.length;i++){
		myForm.Sys_BackDate[i].value=myForm.Sys_BackDateMain.value;
	}
}

<%
if sys_City="花蓮縣" or sys_City="台東縣" then
	'jafe-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	if trim(request("DB_Selt"))="Import" then

	'固定位置
		fp="F:\\ScannerImport\\A000"

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
				response.write "UrlJpg[" & i & "].innerHTML=""<a href='\\ScannerImport\\A000\\"&fil.Name&"' TARGET ='_blank'>催繳單號</a>"";" 
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