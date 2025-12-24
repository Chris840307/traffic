<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/bannernodata.asp"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<TITLE>舉發單-收受註記</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 16px; color:#ff0000; }
.style2 {font-size: 10px; }
-->
</style>
</HEAD>
<%
Server.ScriptTimeout=6000
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
if sys_City="花蓮縣" then
	CarName="姓名"
	KindNo="查證文號"
else
	CarName="車號"
	KindNo="郵局"
end if

if sys_City="彰化縣" then
	BackCauseMainString="'A','B','C'"
else
	BackCauseMainString="'A','B','C','D'"
end if

if trim(request("DB_Selt"))="Selt" then
	Sys_BillNo=Split(Ucase(request("item"))&" ",",")
	Sys_BackCause=Split(request("Sys_BackCause")&" ",",")
	Sys_BackDate=Split(request("Sys_BackDate")&" ",",")
	Sys_mailNumber=Split(request("mailNumber")&" ",",")
	Sys_mailStation=Split(request("mailStation")&" ",",")
	Sys_signman=Split(request("signman")&" ",",")
	'jafe------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	'response.write request("Sys_JpgFile") &"aa"
	Sys_JpgFile=Split(request("Sys_JpgFile"),",")
	'jafe------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		
	Sys_Now=DateAdd("n", -5, now)
	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" and trim(Sys_BackDate(i))<>"" then
			Sys_Now=DateAdd("s",1,Sys_Now)
			strSQL="Update BillMailHistory set SignResonID='"&trim(Sys_BackCause(i))&"',SignDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",SignRecordMemberID="&Session("User_ID")&",ReturnReCordDate="&funGetDate((Sys_Now),1)&",UserMarkMemberID="&Session("User_ID")&",UserMarkDate="&funGetDate((Sys_Now),1)&",UserMarkResonID='"&trim(Sys_BackCause(i))&"',UserMarkReturnDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",mailStation='"&trim(Sys_mailStation(i))&"',signman='"&trim(Sys_signman(i))&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
			conn.execute(strSQL)
			if trim(Sys_mailNumber(i))<>"" then
					strSQL="Update BillMailHistory set mailNumber='"&trim(Sys_mailNumber(i))&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
					conn.execute(strSQL)
			end if
			strSQL="Update Billbase set BillStatus=7 where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"' and RecordStateID=0"
			conn.execute(strSQL)
		end if
	next
	
'jafe----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		'固定位置
		fp="F:\\ScannerImport" 
        finDir="F:\\ScannerImport\\finish\\"
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
					FileDirAndName="/img/scan/" & fDir & Sys_BillNo(i) & "_" & replace(Sys_JpgFile(i),Sys_BillNo(i) & "_","")
					
					strInsert="insert into BillAttatchImage(SN,FileName,BillNo,TypeID,RecordMemberID,RecordDate,RecordStateID)" & _
							  " values("&sMaxSN&",'"&FileDirAndName&"','" & Sys_BillNo(i) & "','0','"& trim(session("User_ID")) &"',SYSDATE,0)"
					conn.execute(strInsert)
					mDir=finDir & year(date)-1911 & "\\" & right("0"&month(date),2) & "\\"
					'response.write Sys_JpgFile(i)
					fso.MoveFile fp & "\\" & Sys_JpgFile(i), mDir & Sys_BillNo(i) & "_" & replace(Sys_JpgFile(i),Sys_BillNo(i) & "_","")
				end if
			Next	
		'end if
	'jafe----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------	
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
%>
<BODY>
<form name=myForm method="post">
<table border="0" width="100%" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle">
			<strong>舉發單-收受 / 第一次雙掛號寄存郵局 註記</strong>					
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
								<%
								chkUnit=""
								if sys_City="基隆市" then
									strSQL="select UnitID from UnitInfo where UnitTypeID in(select UnitTypeID from UnitInfo where UnitID='"&Session("Unit_ID")&"')"

									set rs1=conn.execute(strSQL)
									while Not rs1.eof
										If not ifnull(chkUnit) Then
											chkUnit=chkUnit&"@"
										End if
										chkUnit=chkUnit&trim(rs1("UnitID"))

										rs1.movenext
									wend
									rs1.close
								end if

								strSQL="select UnitID,UnitName from UnitInfo"
								set rs1=conn.execute(strSQL)
								Response.Write "<option value="""&chkUnit&""">請選取</option>"
								while Not rs1.eof
									response.write "<option value="""&rs1("UnitID")&""""
									response.write ">"&rs1("UnitID")&","&rs1("UnitName")&"</option>"
									rs1.movenext
								wend
								rs1.close

								%>
							</select>
						</td>						
					</tr>
					<tr>
						<td><b>整批統一設定資訊</b>	</td>
					</tr>
					<tr>
						<td>預設收受/送達原因統一為&nbsp;</td>
						<td>
							<select name="Sys_BackCauseMain" class="btn1">
								<%strSQL="select ID,Content from DCICode where TypeID=7 and ID in ("&BackCauseMainString&")"
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
						</td>
						<td><input type="button" name="btnDefu" value="確定" onclick="funDefuSelt();"></td>
						<td><font size="2">非必要選項,也可以由下方設定各舉發單不同的收受原因</font></td>
					</tr><tr>
						<td>預設收受/送達日期&nbsp;</td>
						<td>
							<input name="Sys_BackDateMain" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
							<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BackDateMain');">
						</td>
						<td><input type="button" name="btnDefu" value="確定" onclick="funDefuDate();"></td>
						<td><font size="2">非必要選項,也可以由下方設定各舉發單不同的收受日期</font></td>
					</tr><tr>
						<td>預設大宗掛號統一為&nbsp;</td>
						<td><input name="Sys_Number" type="text" class="btn1" size="10" maxlength="15"></td>
						<td><input type="button" name="btnDefu" value="確定" onclick="funNumber();"></td>
						<td><font size="2">非必要選項,也可以由下方設定各舉發單不同的大宗掛號</font></td>
					</tr><tr>
						<td>預設<%=KindNo%>統一為&nbsp;</td>
						<td><input name="Sys_mailStation" type="text" class="btn1" size="10" maxlength="15"></td>
						<td><input type="button" name="btnDefu" value="確定" onclick="funStation();"></td>
						<td><font size="2">非必要選項,也可以由下方設定各舉發單不同的郵局</font></td>
					</tr><tr>
						<td><input type="button" name="btnOK" value="確定存檔" onclick="funSelt();"></td>
						<td><%
								Response.Write "<input type=""button"" name=""insert"" value=""再多30筆"" onClick=""insertRow(fmyTable)"">"
							%>
						</td>
					</tr>
				</table>
				</td></tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">收受/送達 紀錄列表&nbsp;<input type="button" name="btnLoad" value="掃描匯入" onclick="funScannerImport();">　　　<strong>未上傳前如果註記錯誤存檔時，可以再註記一次，蓋掉原本錯誤的紀錄。 </strong><br><%=titleStr%>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<Div style="overflow:auto;width:100%;height:330px;background:#FFFFFF">
				<table id='fmyTable' width="120%" border='0' bgcolor='#FFFFFF'>
					<tr bgcolor="#ffffff">
						<td align='center' bgcolor="#ffffff" nowrap>目前無新增項目 <b>( 掛號碼 / <%=KindNo%> / 代收人 為 非必填項目 )					
						<%  
							if sys_City="台東縣" then 
							  response.write "結案日期可輸入於郵局欄位"
							end if
						%>
						</b></td>
					</tr>
				</table>
			</Div>
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
	<input type="Hidden" name="mailStation" value="">
	<input type="Hidden" name="signman" value="">
	<input type="Hidden" name="Sys_JpgFile" value="">
	<input type="Hidden" name="DB_Selt" value="">
	<input type="Hidden" name="SaveError" value="">
</form>

</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var cunt=0;
function insertRow(isTable){
	<%cnt=29%>
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
		txtArea.innerHTML =cnt_num+".&nbsp;<Span ID='UrlJpg'>單號</span>&nbsp;<input type=text name='item' size=7 class='btn1' onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;原因<%=BackCause_btn%><select name='Sys_BackCause' class='btn1'><%=seltarr%></select>&nbsp;&nbsp;收受日<input type=text name='Sys_BackDate' size=4 class='btn1' onkeydown='funBackDate("+cunt+");'>&nbsp;&nbsp;<%=CarName%><input type=text name='CarNo' size=8 class='btn1' ReadOnly>&nbsp;&nbsp;應到案日<input type=text name='DeallineDate' size=4 class='btn1' ReadOnly>&nbsp;&nbsp;掛號<input type=text name='mailNumber' size=10 class='btn1' maxlength='20' onkeydown='funmailNumber("+cunt+");'>&nbsp;&nbsp;代收<input type=text name='signman' size=4 class='btn1'>&nbsp;&nbsp;<%=KindNo%><input type=text name='mailStation' size=4 class='btn1'><input type='Hidden' name='Sys_JpgFile' value=''>";
	}
}

function funBackDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.Sys_BackDate.length){
			myForm.mailNumber[itemcnt-1].focus();
		}
	}
}
function keyFunction(itemcnt) {
	//myForm.item[itemcnt-1].value=myForm.item[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9) {
		if (chkBillNo(itemcnt)){
			if (myForm.item[itemcnt-1].value!=''){
				myForm.chkcnt.value=itemcnt;
				runServerScript("chkAcceptBillNo.asp?BillNo="+myForm.item[itemcnt-1].value);
			}
		}else{
			alert("單號長度必須為9碼!!");
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
	var js_Item="";
	var js_BackCause="";
	var js_BackDate="";
	var js_mailNumber="";
	var js_mailStation="";
	var js_signman="";
	var js_Sys_JpgFile="";

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
		for(i=0;i<myForm.item.length;i++){
			if(myForm.item[i].value!=''){
				if(js_Item!=''){
					js_Item=js_Item+',';
					js_BackCause=js_BackCause+',';
					js_BackDate=js_BackDate+',';
					js_mailNumber=js_mailNumber+',';
					js_mailStation=js_mailStation+',';
					js_signman=js_signman+',';
					js_Sys_JpgFile=js_Sys_JpgFile+',';
				}
				js_Item=js_Item + myForm.item[i].value;
				js_BackCause=js_BackCause + myForm.Sys_BackCause[i].value;
				js_BackDate=js_BackDate + myForm.Sys_BackDate[i].value;
				js_mailNumber=js_mailNumber + myForm.mailNumber[i].value;
				js_mailStation=js_mailStation + myForm.mailStation[i].value;
				js_signman=js_signman + myForm.signman[i].value;
				js_Sys_JpgFile=js_Sys_JpgFile + myForm.Sys_JpgFile[i].value;
			}
		}

		upForm.item.value=js_Item;
		upForm.Sys_BackCause.value=js_BackCause;
		upForm.Sys_BackDate.value=js_BackDate;
		upForm.mailNumber.value=js_mailNumber;
		upForm.mailStation.value=js_mailStation;
		upForm.signman.value=js_signman;
		upForm.Sys_JpgFile.value=js_Sys_JpgFile;
		upForm.DB_Selt.value="Selt";
		if(upForm.SaveError.value!=''){
			if(confirm(upForm.SaveError.value+"是否要繼續執行?")){
				upForm.submit();
			}
		}else{
			upForm.submit();
		}
	}
}
// jafe--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function funScannerImport(){
		myForm.DB_Selt.value="Import";
		myForm.submit();
}
// jafe____________________________________________________________________________________________________________

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


<%
	Response.Write "for(j=0;j<=3;j++){insertRow(fmyTable);}"
%>

function funDefuDate(){
	for(i=0;i<myForm.Sys_BackDate.length;i++){
		myForm.Sys_BackDate[i].value=myForm.Sys_BackDateMain.value;
	}
}

<%

	strRul="select Value from Apconfigure where ID=2"
	set rsRul=conn.execute(strRul)
	chkBillno=left(trim(rsRul("Value")),1)
	rsRul.Close

'jafe-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if trim(request("DB_Selt"))="Import" then
'固定位置
fp="F:\\ScannerImport" 

        set fso=Server.CreateObject("Scripting.FileSystemObject")

        set fod=fso.GetFolder(fp)
        set fic=fod.Files
   
    i=-1
	response.write "upForm.SaveError.value="""";"

    For Each fil In fic
		if UCase(fso.GetExtensionName(fil.Name)) ="JPG" Or UCase(fso.GetExtensionName(fil.Name)) ="JPEG" then
			i=i+1
			if i >=90 then
				if (i) mod 30 =0 then response.write "insertRow(fmyTable);"
			end if
			Sys_tmpBillNo=Split(fil.Name,"_")

			if trim(Sys_tmpBillNo(0))<>"" then
				if left(trim(Sys_tmpBillNo(0)),1)=chkBillno then
					response.write "myForm.item[" & i & "].value='" & Sys_tmpBillNo(0) &"';" 
				end if
			end if

			response.write "myForm.chkcnt.value="&(i+1)&";"
			response.write "UrlJpg[" & i & "].innerHTML=""<a href='\\ScannerImport\\"&fil.Name&"?nowTime=" & now & "' TARGET ='_blank'>單號</a>"";" 
			response.write "myForm.Sys_JpgFile[" & i & "].value='" & fil.Name &"';" 

			tb_BillBase="select BillNo,CarNo,BillUnitID,DeallineDate,billFillDate from BillBase where BillNo='"&trim(Ucase(Sys_tmpBillNo(0)))&"' and RecordStateID=0"

			tb_BillbaseDCIReturn="select BillNo,CarNo,Owner,DciReturnStation from BillbaseDCIReturn where BillNo='"&trim(Ucase(Sys_tmpBillNo(0)))&"'"
			
			strSQL = "select a.DeallineDate,a.billFillDate,d.DciReturnStation,b.DCIStationName,b.StationID,c.UnitID,c.UnitName,d.CarNo,d.Owner from ("&tb_BillBase&") a,Station b,UnitInfo c,("&tb_BillbaseDCIReturn&") d where a.BillNo=d.BillNo and a.CarNo=d.CarNo and d.DCIReturnStation=b.DCIStationID(+) and a.BillUnitID=c.UnitID(+)"

			set rscnt=conn.execute(strSQL)
			if Not rscnt.eof then
				errStr=""
				tmpSql="select BillNo,ExchangeTypeID,ReturnMarkType,DCIReturnStatusID,ExchangeDate from DCILog where billsn=(select sn from billbase where billno='"&trim(Ucase(Sys_tmpBillNo(0)))&"' and RecordStateID=0) and billno is not null order by ExchangeDate Desc"
				strSQL="select * from ("&tmpSql&") DCILogTmp where rownum=1"

				set rschk=conn.execute(strSQL)
				If Not rschk.eof Then
					If trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="7"  and trim(rschk("DCIReturnStatusID"))="S" Then
						errStr=errStr&"單號："&rschk("BillNo")&"該筆舉發單有做過收受註記!!\n"

					elseIf trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="4"  and trim(rschk("DCIReturnStatusID"))="S" Then
						errStr=errStr&"單號："&rschk("BillNo")&"該筆舉發單有做過寄存送達!!\n"

					elseIf trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="5"  and trim(rschk("DCIReturnStatusID"))="S" Then
						errStr=errStr&"單號："&rschk("BillNo")&"該筆舉發單有做過公示送達!!\n"
					 '單退不用測銷送達

					elseif trim(rschk("ExchangeTypeID"))="N" and trim(rschk("DCIReturnStatusID"))="n" Then
						errStr=errStr&"單號："&rschk("BillNo")&"該筆舉發單已結案請至上傳下載資料查詢確認!!\n"
					End If

				End if
				rschk.close
				if errStr<>"" then response.write "upForm.SaveError.value=upForm.SaveError.value+'"&errStr&"';"

				response.write "myForm.Sys_BackDate[" & i & "].value='"& gInitDT(DateAdd("d",getBillReturnDate,rscnt("billFillDate")))&"';"
				response.write "myForm.DeallineDate[" & i & "].value='"& gInitDT(trim(rscnt("DeallineDate"))) &"';"
				response.write "myForm.CarNo[" & i & "].value='"&trim(rscnt("CarNo"))&"';"
			end if 
			rscnt.close

        end if
    Next
	
end if
'jafe-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
%>
</script>