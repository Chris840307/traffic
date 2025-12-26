<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>寄存送達註記</TITLE>
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
	Sys_BillNo=Split(Ucase(trim(request("item")))&" ",",")
	Sys_BackCause=Split(request("Sys_BackCause")&" ",",")
	Sys_BackDate=Split(request("Sys_BackDate")&" ",",")
	EffectDate=Split(request("EffectDate")&" ",",")
	MailDate=Split(request("MailDate")&" ",",")
	Sys_mailStation=Split(request("mailStation")&" ",",")
	Sys_mailNumber=Split(request("mailNumber")&" ",",")
	Sys_Now=DateAdd("n", -5, now)

	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" and trim(Sys_BackDate(i))<>"" then
			Sys_Now=DateAdd("s",1,Sys_Now)
			strSQL="Update BillMailHistory set StoreAndSendReturnResonID='"&trim(Sys_BackCause(i))&"',StoreAndSendMailReturnDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",StoreAndSendRecordMemberID="&Session("User_ID")&",StoreAndSendReCordDate="&funGetDate((Sys_Now),1)&",UserMarkMemberID="&Session("User_ID")&",UserMarkDate="&funGetDate((Sys_Now),1)&",UserMarkResonID='"&trim(Sys_BackCause(i))&"',UserMarkReturnDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&",StoreAndSendEffectDate="&funGetDate(gOutDT(EffectDate(i)),0)&",StoreAndSendMailDate="&funGetDate(gOutDT(MailDate(i)),0)&",MailTypeID=2,mailStation='"&trim(Sys_mailStation(i))&"' where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"'"
			conn.execute(strSQL)
			if trim(Sys_mailNumber(i))<>"" then
				strSQL="Update BillMailHistory set StoreAndSendMailNumber='"&trim(Sys_mailNumber(i))&"' where BillNo='"&trim(Sys_BillNo(i))&"'"
				conn.execute(strSQL)
			end if
			strSQL="Update Billbase set BillStatus=3 where BillNo='"&trim(Ucase(Sys_BillNo(i)))&"' and RecordStateID=0"
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
		<td height="27" bgcolor="#1BF5FF" class="pagetitle"><strong>寄存送達註記</strong><br>
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
						<hr><%
						if sys_City="台南市" then
							Sys_BackCause="<input type=Hidden name='Sys_BackCause' value='T'>"
						else
							Response.Write "預設送達原因統一為&nbsp;"
							Response.Write "<select name=""Sys_BackCauseMain"" class=""btn1"">"

							strSQL="select ID,Content from DCICode where TypeID=7 and ID in('5','6','7','T')"
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
							Response.Write "</select>"

							Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
							Response.Write "<input type=""button"" name=""btnDefu"" value=""確定"" onclick=""funDefuSelt();"">"

							Response.Write "&nbsp;<font size=""2"">非必要選項,也可以由下方設定各舉發單不同的單退原因</font><br>"
							
							titleStr=""
							BackCause_btn="<input name='Sys_BackCauseIndex' type=Hidden class='btn1' size=1 maxlength=2 onkeyup=funBackCauseIndex(this,'Sys_BackCause',""+cunt+"");>"
							if sys_City="高雄市" or sys_City="高港局" then
								titleStr="<br><span class=""style1"">"&seltName&"</span>"
								BackCause_btn="<input name='Sys_BackCauseIndex' type=text class='btn1' size=1 maxlength=2 onkeyup=funBackCauseIndex(this,'Sys_BackCause',""+cunt+"");>"
							end if

							Sys_BackCause="原因"&BackCause_btn&"<select name='Sys_BackCause' class='btn1'>"&seltarr&"</select>"
						end if%>
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
						預設期滿日期統一為&nbsp;<input name="Sys_DefMailDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_DefMailDate');">
						&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funDefuDate('Sys_DefMailDate','MailDate');">
						&nbsp;<font size="2">非必要選項,也可以由下方設定各舉發單不同的送達日期</font>
						<br>
						預設大宗掛號統一為&nbsp;
						<input name="Sys_Number" type="text" class="btn1" size="10" maxlength="15">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funNumber();"><font size="2">&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的大宗掛號</font>
						<br>
						預設郵局統一為&nbsp;
						<input name="Sys_mailStation" type="text" class="btn1" size="10" maxlength="15">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<input type="button" name="btnDefu" value="確定" onclick="funStation();">
						&nbsp;&nbsp;非必要選項,也可以由下方設定各舉發單不同的郵局</font>
						<hr>
						&nbsp;&nbsp;<input type="button" name="btnOK" value="儲存" onclick="funSelt();">
						<img src="space.gif" width="9" height="8">
						<%
						Response.Write "<input type=""button"" name=""insert"" value=""再多30筆"" onClick=""insertRow(fmyTable)"">"
						%>
						<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#1BF5FF">寄存送達註記列表<img src="space.gif" width="12" height="8"> <b>期滿日,掛號碼,郵局 為 非必填 欄位</b>. 其餘為必填欄位<br><%=titleStr%></td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<Div style="overflow:auto;width:100%;height:270px;background:#FFFFFF">
				<table id='fmyTable' width='105%' border='0' bgcolor='#FFFFFF'>
					<tr bgcolor="#ffffff">
						<td align='center' bgcolor="#ffffff" nowrap></td>
					</tr>
				</table>
			</Div>
		</td>
	</tr>
	<tr align="center">
		<td height="35" bgcolor="#1BF5FF">
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
	<input type="Hidden" name="EffectDate" value="">
	<input type="Hidden" name="MailDate" value="">
	<input type="Hidden" name="mailStation" value="">
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
		txtArea.innerHTML =cnt_num+".&nbsp;單號<input type=text name='item' size=6 class='btn1' onkeydown='keyFunction("+cunt+");'>&nbsp;&nbsp;<%=Sys_BackCause%>&nbsp;&nbsp;單退日<input type=text name='Sys_BackDate' size=3 class='btn1' onkeyup='chknumber(this);' onkeydown='keyBackDate("+cunt+");' maxlength='7'>&nbsp;&nbsp;送達日<input type=text name='EffectDate' size=3 class='btn1' onkeyup='chknumber(this);' onkeydown='keyEffectDate("+cunt+");' maxlength='7'>&nbsp;&nbsp;期滿日<input type=text name='MailDate' size=3 class='btn1' onkeydown='keyMailDate("+cunt+");' maxlength='7'>&nbsp;&nbsp;掛號碼<input type=text name='mailNumber' size=3 class='btn1' onkeydown='keyMailNumber("+cunt+");' maxlength='20'>&nbsp;&nbsp;<%=CarName%><input type=text name='CarNo' size=4 class='btn1' readOnly>&nbsp;&nbsp;郵局<input type=text name='mailStation' size=4 class='btn1' maxlength='45' onkeydown='keymailStation("+cunt+");'>";
	}
}

function keyFunction(itemcnt) {
	//myForm.item[itemcnt-1].value=myForm.item[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9) {
		if (chkBillNo(itemcnt)){
			if (myForm.item[itemcnt-1].value!=''){
				myForm.chkcnt.value=itemcnt;
				runServerScript("chkReturnMark_two.asp?BillNo="+myForm.item[itemcnt-1].value);
			}
		}else{
			alert("單號長度必須為9碼!!");
		}
	}
}

function keyMailNumber(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.mailNumber.length){
			myForm.item[itemcnt].focus();
		}
	}
}

function keymailStation(itemcnt) {
	if (event.keyCode==13||event.keyCode==9) {
		if(itemcnt<myForm.mailStation.length){
			myForm.item[itemcnt].focus();
		}
	}
}

function keyBackDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		myForm.EffectDate[itemcnt-1].focus();
	}
}

function keyMailDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		myForm.mailNumber[itemcnt-1].focus();
	}
}

function keyEffectDate(itemcnt) {
	if (event.keyCode==13||event.keyCode==9){
		myForm.chkcnt.value=itemcnt;
		runServerScript("chkMailDate.asp?EffectDate="+myForm.EffectDate[itemcnt-1].value);
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

	var item='';
	var Sys_BackCause='';
	var Sys_BackDate='';
	var EffectDate='';
	var MailDate='';
	var mailStation='';
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
	for(i=0;i<myForm.item.length;i++){
		if(myForm.item[i].value!=''){
			if(myForm.EffectDate[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 行送達日期不可空白!!");
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
					EffectDate=EffectDate+',';
					MailDate=MailDate+',';
					mailStation=mailStation+',';
					mailNumber=mailNumber+',';
				}
				item=item + myForm.item[i].value;
				Sys_BackCause=Sys_BackCause + myForm.Sys_BackCause[i].value;
				Sys_BackDate=Sys_BackDate + myForm.Sys_BackDate[i].value;
				EffectDate=EffectDate + myForm.EffectDate[i].value;
				MailDate=MailDate + myForm.MailDate[i].value;
				mailStation=mailStation + myForm.mailStation[i].value;
				mailNumber=mailNumber + myForm.mailNumber[i].value;
			}
		}

		upForm.item.value=item;
		upForm.Sys_BackCause.value=Sys_BackCause;
		upForm.Sys_BackDate.value=Sys_BackDate;
		upForm.EffectDate.value=EffectDate;
		upForm.MailDate.value=MailDate;
		upForm.mailStation.value=mailStation;
		upForm.mailNumber.value=mailNumber;

		upForm.DB_Selt.value="Selt";
		upForm.submit();
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

function funDefuDate(obj,defName){
	for(i=0;i<eval("myForm."+defName).length;i++){
		eval("myForm."+defName+"["+i+"]").value=eval("myForm."+obj).value;
	}
	
	if(obj=='Sys_DefEffectDate'){
		for(i=0;i<eval("myForm."+defName).length;i++){
			eval("myForm.MailDate["+i+"]").value=DatePro(eval("myForm."+obj).value);
		}
	}
}
</script>

<script language="vbscript"> 
	Function DatePro(iDate)
		dim strDate
		strDate=trim(gInitDT(DateAdd("d",90,gOutDT(iDate))))
		DatePro=strDate
	End Function
	Function gOutDT(iDate)
		if trim(iDate)<>"" then
			Dim tTemp
			tTemp=DateSerial(Left(iDate,len(iDate)-4)+1911,Mid(iDate,len(iDate)-3,2),right(iDate,2))
			gOutDT = tTemp
		else
			gOutDT = ""
		end if
	End Function
	Function gInitDT(iDate)
		Dim iTemp 
		if trim(iDate)<>"" then
			iTemp=right(year(iDate)-1911,3) &_
			right("0"&month(iDate),2) &_
			right("0"&day(iDate),2)
			gInitDT = iTemp
		else
			gInitDT = ""
		end if
	End Function
</script>