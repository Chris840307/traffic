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
	'jafe-------------------------------------------------------------------------------------------------------------------------------------------------------
	Sys_JpgFile=Split(request("Sys_JpgFile"),",")
	'jafe----------------------------------------------------------------------------------------------------------------------------------------------------
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
	Next

	If Ubound(Sys_JpgFile)>=0 Then
		'jafe------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
		'固定位置
		fp="F:\\ScannerImport" 
        finDir="F:\\ScannerImport\\公示\\"

        set fso=Server.CreateObject("Scripting.FileSystemObject")

		if Session("Unit_ID")="Z000" then

			if fso.FolderExists(fp & "\\A000") then
				fp="F:\\ScannerImport\\A000"
			end If 
		else

			if fso.FolderExists(fp & "\\" & Session("Unit_ID")) then
				fp="F:\\ScannerImport\\" & Session("Unit_ID")
			end If 
		end if

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
					FileDirAndName="/img/scan/公示/" & fDir & Sys_BillNo(i) & "_" & replace(Sys_JpgFile(i),Sys_BillNo(i) & "_","")
					
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
	'jafe----------------------------------------------------------------------------------------------------------------------------------------------------------------		
	End if 
		
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
						&nbsp;&nbsp;<input type="button" name="btnOK" value="匯入地址資料" onclick="funAddressSelt();"><br>
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
		<td height="26" bgcolor="#FFCC33">單退紀錄列表&nbsp;<%
			Response.Write "<input type=""button"" name=""btnOK"" value=""上傳送達掃描檔(JPG)"" onclick=""funUPfile();"">"
		%>
			<input type="button" name="btnLoad" value="掃描匯入" onclick="funScannerImport();">
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<table id='fmyTable' width='978' border='0' bgcolor='#ffffff'>
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
	<input type="Hidden" name="Sys_JpgFile" value="">
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

	isTable=document.all.fmyTable;

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
		txtArea.innerHTML =cnt_num+".&nbsp;<Span ID='UrlJpg'>單號</span>&nbsp;<input type=text name='item' size=10 class='btn1' onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;原因<select name='Sys_BackCause' class='btn1'><%=seltarr%></select>&nbsp;&nbsp;單退日期<input type=text name='Sys_BackDate' size=10 class='btn1' onkeyup='funkeyChk(this);' onkeydown='keyBackDate("+cunt+");' maxlength='7'>&nbsp;&nbsp;大宗掛號碼<input type=text name='mailNumber' size=10 class='btn1' onkeydown='keyMailNumber("+cunt+");' maxlength='20'>&nbsp;&nbsp;車號<input type=text name='CarNo' size=10 class='btn1' readOnly><input type='Hidden' name='Sys_JpgFile' value=''>";
	}
}
function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	winopen.focus();
	return win;
}

function funUPfile(){
	newWin("","UPfile",700,550,50,10,"yes","yes","yes","no");
	UrlStr="upload/default.asp";
	myForm.action=UrlStr;
	myForm.target="UPfile";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funAddressSelt(){
	newWin("","Address",700,550,50,10,"yes","yes","yes","no");
	UrlStr="BillBaseOpenReturnMarkSendStyle_miaoli.asp";
	myForm.action=UrlStr;
	myForm.target="Address";
	myForm.submit();
	myForm.action="";
	myForm.target="";
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
	var js_Sys_JpgFile="";
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
					js_Sys_JpgFile=js_Sys_JpgFile+',';
				}
				item=item + myForm.item[i].value;
				Sys_BackCause=Sys_BackCause + myForm.Sys_BackCause[i].value;
				Sys_BackDate=Sys_BackDate + myForm.Sys_BackDate[i].value;
				mailNumber=mailNumber + myForm.mailNumber[i].value;
				js_Sys_JpgFile=js_Sys_JpgFile + myForm.Sys_JpgFile[i].value;
			}
		}

		upForm.item.value=item;
		upForm.Sys_BackCause.value=Sys_BackCause;
		upForm.Sys_BackDate.value=Sys_BackDate;
		upForm.mailNumber.value=mailNumber;
		upForm.Sys_JpgFile.value=js_Sys_JpgFile;
		upForm.DB_Selt.value="Selt";
		upForm.submit();
	}
}
// jafe------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
function funScannerImport(){
		myForm.DB_Selt.value="Import";
		myForm.submit();

}
// jafe____________________________________________________________________________________________________________

<%

	strRul="select Value from Apconfigure where ID=2"
	set rsRul=conn.execute(strRul)
	chkBillno=left(trim(rsRul("Value")),1)
	rsRul.Close

	Response.Write "for(j=0;j<=3;j++){insertRow(fmyTable);}"
'jafe---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
if trim(request("DB_Selt"))="Import" then
'固定位置
	fp="F:\\ScannerImport" 

	set fso=Server.CreateObject("Scripting.FileSystemObject")

	if Session("Unit_ID")="Z000" then

		if fso.FolderExists(fp & "\\A000") then
			fp="F:\\ScannerImport\\A000"
		end If 

	else

		if fso.FolderExists(fp & "\\" & Session("Unit_ID")) then
			fp="F:\\ScannerImport\\" & Session("Unit_ID")
		end If 

	end if

	set fod=fso.GetFolder(fp)
	set fic=fod.Files
   
    i=-1
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

			response.write "UrlJpg[" & i & "].innerHTML=""<a href='"&replace(fp,"F:","")&"\\"&fil.Name&"' TARGET ='_blank'>單號</a>"";" 
			response.write "myForm.Sys_JpgFile[" & i & "].value='" & fil.Name &"';" 



			tb_BillBase="select BillNo,CarNo,BillUnitID from BillBase where BillNo='"&trim(Ucase(Sys_tmpBillNo(0)))&"' and RecordStateID=0"

			tb_BillbaseDCIReturn="select BillNo,CarNo,DCIReturnStation,Owner from BillbaseDCIReturn where BillNo='"&trim(Ucase(Sys_tmpBillNo(0)))&"'"

			strSQL = "select (select max(billstatus) from billbase where billno='"&trim(Ucase(Sys_tmpBillNo(0)))&"') billstatus,d.DCIReturnStation,b.DCIStationName,b.StationID,c.UnitID,c.UnitName,d.CarNo,d.Owner from ("&tb_BillBase&") a,Station b,UnitInfo c,("&tb_BillbaseDCIReturn&") d where a.BillNo=d.BillNo and a.CarNo=d.CarNo and d.DCIReturnStation=b.DCIStationID(+) and a.BillUnitID=c.UnitID(+)"

			set rscnt=conn.execute(strSQL)
			if Not rscnt.eof then

				tmpSql="select BillNo,ExchangeTypeID,ReturnMarkType,DCIReturnStatusID,ExchangeDate from DCILog where billsn=(select sn from billbase where billno='"&trim(Ucase(Sys_tmpBillNo(0)))&"' and RecordStateID=0) and billno is not null order by ExchangeDate Desc"
				strSQL="select * from ("&tmpSql&") DCILogTmp where rownum=1"

				set rschk=conn.execute(strSQL)
				If Not rschk.eof Then
					If trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="7" and trim(rschk("DCIReturnStatusID"))="S" Then
						errStr=errStr&"單號："&rschk("BillNo")&"該筆舉發單有做過收受註記，請先由舉發單維護系統進行撤銷送達!!\n"

					elseIf trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="4"  and trim(rschk("DCIReturnStatusID"))="S" Then
						errStr=errStr&"單號："&rschk("BillNo")&"該筆舉發單有做過寄存送達!!\n"

					elseIf trim(rschk("ExchangeTypeID"))="N" and trim(rschk("ReturnMarkType"))="5"  and trim(rschk("DCIReturnStatusID"))="S" Then
						errStr=errStr&"單號："&rschk("BillNo")&"該筆舉發單有做過公示送達!!\n"

					elseif trim(rschk("ExchangeTypeID"))="N" and trim(rschk("DCIReturnStatusID"))="n" Then
						errStr=errStr&"單號："&rschk("BillNo")&"該筆舉發單已結案請至上傳下載資料查詢確認!!\n"

					elseif trim(rscnt("billstatus"))="9" Then
						errStr=errStr&"單號："&rschk("BillNo")&"該筆舉發單已結案請至舉發單資料維護查詢確認!!\n"

					End If

				End if
				rschk.close

				response.write "myForm.CarNo[" & i & "].value='"&trim(rscnt("Owner"))&"';"
			else
				errStr=errStr&"單號："&Sys_tmpBillNo(0)&"無此單號!!\n"
			end if 
			rscnt.close
			
        end if
    Next
	if errStr<>"" then response.write "alert('"&errStr&"');"
end if
'jafe-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

%>
</script>