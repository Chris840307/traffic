<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>攔停點收系統</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
</HEAD>
<BODY>
<%
Server.ScriptTimeout=6000
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

Server.ScriptTimeout=6000
strCity="select value from Apconfigure where id=3"
set rsCity=conn.execute(strCity)
sys_RuleVer=trim(rsCity("value"))
rsCity.close

if trim(request("DB_Selt"))="Selt" then
	Sys_BillNo=Split(Ucase(trim(request("item"))),",")
	Sys_CarNo=Split(Ucase(trim(request("CarNo"))),",")
	Sys_illegalDate=Split(trim(request("illegalDate")),",")
	Sys_Rule1=Split(trim(request("Rule1")),",")
	Sys_DriverName=Split(trim(request("DriverName")),",")
	Sys_Fastener1=Split(trim(request("Fastener1")),",")
	Sys_BillMemID1=Split(trim(request("BillMemID1")),",")
	Sys_Rule2=Split(trim(request("Rule2")),",")
	Sys_Fastener2=Split(trim(request("Fastener2")),",")
	Sys_chkBackBillBase=Split(trim(request("Sys_BackBillBase")),",")
	Sys_Note=Split(trim(request("Note")),",")
	Sys_BillUnitid=Split(Trim(Request("BillUnitID")),",")
	Sys_AcceptDate=Trim(Request("AcceptDate"))

	Sys_now=funGetDate(now,1)
	
	For i = 0 to Ubound(Sys_BillNo)
		If not ifnull(Sys_BillNo(i)) Then
			strSQL="select count(1) cmt from BillStopCarAccept where BillNo='"&trim(Sys_BillNo(i))&"' and recordstateid=0"
			
			set rsnt=conn.execute(strSQL)

			If cdbl(rsnt("cmt"))=0 Then
				strSQL="Insert into BillStopCarAccept(BillNo,CarNo,BillUnitID,IllegalDate,AcceptDate,Rule1,Driver,FastenerTypeID1,BILLMEMID1,RULEVER,RECORDSTATEID,RECORDDATE) values('"&trim(Sys_BillNo(i))&"','"&trim(Sys_CarNo(i))&"','"&trim(Sys_BillUnitid(i))&"',"&funGetDate(gOutDT(Sys_illegalDate(i)),0)&","&funGetDate(gOutDT(Sys_AcceptDate),0)&",'"&trim(Sys_Rule1(i))&"','"&trim(Sys_DriverName(i))&"','"&trim(Sys_Fastener1(i))&"',"&trim(Sys_BillMemID1(i))&",'"&trim(sys_RuleVer)&"',0,"&Sys_now&")"
				conn.execute(strSQL)
			else
				strSQL="Update BillStopCarAccept set CarNo='"&trim(Sys_CarNo(i))&"',BillUnitID='"&trim(Sys_BillUnitid(i))&"',IllegalDate="&funGetDate(gOutDT(Sys_illegalDate(i)),0)&",AcceptDate="&funGetDate(gOutDT(Sys_AcceptDate),0)&",Rule1='"&trim(Sys_Rule1(i))&"',Driver='"&trim(Sys_DriverName(i))&"',FastenerTypeID1='"&trim(Sys_Fastener1(i))&"',BILLMEMID1="&trim(Sys_BillMemID1(i))&",RuleVer='"&trim(sys_RuleVer)&"' where billno='"&trim(Sys_BillNo(i))&"' and recordstateid=0"

				conn.execute(strSQL)
			End if
			updstr=""
			If not ifnull(Sys_chkBackBillBase(i)) Then
				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"recordstateid=-1,Note='"&Sys_Note(i)&"'"
			End if

			If not ifnull(Sys_Rule2(i)) Then
				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"Rule2='"&Sys_Rule2(i)&"'"
			End if

			If not ifnull(Sys_Fastener2(i)) Then
				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"FastenerTypeID2='"&Sys_Fastener2(i)&"'"
			End if

			If trim(Session("UnitLevelID"))>1 and (not ifnull(Request("chkAccept"))) Then
				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"RecordMemberID1="&Session("User_ID")
			elseIf trim(Session("UnitLevelID"))=1 and (not ifnull(Request("chkAccept"))) Then
				If not ifnull(updstr) Then updstr=updstr&","
				updstr=updstr&"RecordMemberID1="&Session("User_ID")&",RecordMemberID2="&Session("User_ID")
			End if
			
			If not ifnull(updstr) Then
				strSQL="Update BillStopCarAccept set "&updstr&" where billno='"&trim(Sys_BillNo(i))&"' and recordstateid=0"
				conn.execute(strSQL)
			End if

			rsnt.close
		end if
	Next
	Response.write "<script>"
	Response.Write "alert('簽收送件完成！');"
	Response.write "</script>"
end if
%>
<form name="myForm" method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle">
			<strong>攔停點收系統</strong>
			<a href="./Upaddress/CheckAccept.doc"><font size="3" color="blue"><u>點收件系統使用說明</u></font></a>
		</td>
	</tr>
	<tr>
		<td>預設違規日期&nbsp;
			<input name="Sys_BackDateMain" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
			<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BackDateMain');">
			<input type="button" name="btnDefu" value="確定" onclick="funDefuDate();">
			<font size="2">非必要選項,也可以由下方設定各舉發單不同的收受日期</font>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" name="btnOK" value="匯入攔停資料" onclick="funChkSelt();">
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						<table border="0">
							<tr>
								<td>
									尚未點收案件<span id='BillBaseOrder'></span><br>
									<table width="450" border="1" cellpadding="0" cellspacing="0">
										<tr bgcolor="#EBFBE3" align="center">
											<th width="20%">點收日</th>
											<th width="30%">送件單位</th>
											<th width="15%">件數</th>
											<th width="35%">點收</th>
										</tr>
										<tr>
											<td colspan="4">
												<Div style="overflow:auto;width:100%;height:100px;background:#FFFFFF">
												<table width="100%" border="1" cellpadding="1" cellspacing="0"><%

												strSQL="select a.AcceptDate,a.BillUnitID,a.RecordDate,a.cmt,b.UnitName from (select AcceptDate,BillUnitID,RecordDate,count(1) cmt from BillStopCarAccept where billunitid in(select unitid from unitinfo where unittypeid=(select unittypeid from unitinfo where unitid='"&trim(Session("Unit_ID"))&"')) and RecordMemberID1 is null and RecordMemberID2 is null and recordstateid=0 group by AcceptDate,BillUnitID,RecordDate) a,UnitInfo b where a.BillUnitID=b.UnitID"

												set rs=conn.execute(strSQL)
												While not rs.eof
													recorddate=datevalue(rs("recorddate"))&" "&hour(rs("recorddate"))&":"&minute(rs("recorddate"))&":"&second(rs("recorddate"))

													Response.Write "<tr align=""center"">"
													Response.Write "<td width=""20%"">"&gInitDT(rs("AcceptDate"))&"</th>"
													Response.Write "<td width=""30%"">"&rs("UnitName")&"</th>"
													Response.Write "<td width=""15%"">"&rs("cmt")&"</th>"
													Response.Write "<td width=""35%""><input type=""button"" name=""btnAcc"" value=""查閱"" onclick=""funAcceptLoad('"&gInitDT(rs("AcceptDate"))&"','"&rs("BillUnitID")&"','"&rs("UnitName")&"','"&rs("cmt")&"','"&recorddate&"');""></th>"
													Response.Write "</tr>"
													rs.movenext
												Wend
												rs.close
												%>
												</table>
												</div>
											</td>
										</tr>
									</table>
								</td>
								<td>
									&nbsp;&nbsp;點收日期
									<input type=text name='AcceptDate' size="5" class='btn1' maxlength='7' value="<%=gInitDT(now)%>">
									<input type="button" name="datestr" value="..." onclick="OpenWindow('AcceptDate');"><br><br>

									<input class='btn1' type='checkbox' name='chkAccept' value='1'>已點收
									
									&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" name="btnOK" value="確定儲存" onclick="funSelt();">
									<img src="space.gif" width="9" height="8">
									<input type="button" name="insert" value="再多30筆" onClick="insertRow(fmyTable)">
									<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
									<br>
									<a href="./Upaddress/AcceptSample.xls"><font size="4" color="blue"><u>舉發員警攔停點收檔下載</u></font></a>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td>
						<%
							Response.Write "<font size=3 color=""Red""><B>"
							Response.Write "扣件代碼"
							Response.Write "</B></font><br>"

							Response.Write "<font size=3 color=""Red"">"
							strSQL="select ID,Content from DCICode where typeid=6 order by ID"
							set rscode=conn.execute(strSQL)
							While not rsCode.eof
								
								Response.Write rscode("ID")&"："&rscode("Content")
								
								If trim(rscode("ID"))="9" Then
									Response.Write "<BR>"
								elseIf trim(rscode("ID"))="I" Then
									Response.Write "<BR>"
								else
									Response.Write "，"
								end if
											
								rsCode.movenext
							Wend
							Response.Write "</font>"
							rsCode.close
						%>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">攔停點收紀錄列表 ( 輸入完成按Enter可自動跳到下一格 )</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<Div style="overflow:auto;width:100%;height:330px;background:#FFFFFF">
				<table id='fmyTable' width='100%' border='0' bgcolor='#FFFFFF'>
					<tr bgcolor="#ffffff">
						<td align='center' bgcolor="#ffffff" nowrap></td>
					</tr>
				</table>
			</div>
		</td>
	</tr>
	<tr align="center">
		<td height="35" bgcolor="#FFDD77">
			<input type="button" name="btnOK1" value="確定存檔" onclick="funSelt();">
			<input type="button" name="insert2" value="再多30筆" onClick="insertRow(fmyTable)">
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
		var cnt_num=("0000"+cunt).substr(("0000"+cunt).length-3,3);

		txtArea.innerHTML =cnt_num+"單號<input type=text name='item' size=6 class='btn1' onkeydown='keyFunction("+cunt+");' maxlength='9'>&nbsp;&nbsp;車號<input type=text name='CarNo' size=5 class='btn1' onkeydown='keyCarNo("+cunt+");'>&nbsp;&nbsp;違規日<input type=text name='illegalDate' size=3 class='btn1' onkeydown='KeyillegalDate("+cunt+");' maxlength='7'>&nbsp;&nbsp;法條1<input type=text name='Rule1' size=5 class='btn1' onkeydown='KeyRule1("+cunt+");'>&nbsp;&nbsp;違規人姓名<input type=text name='DriverName' size=2 class='btn1' onkeydown='KeyDriverName("+cunt+");' maxlength='4'>&nbsp;&nbsp;扣件1<input type=text name='Fastener1' size=2 class='btn1' onkeydown='KeyFastener1("+cunt+");'><input type=hidden name='BillMemID1'><input type='hidden' name='BillUnitID'>舉發員警<input type=text name='BillMemName' size=2 class='btn1' onkeydown='KeyBillMem1("+cunt+");'  onkeyup='getBillMem1("+cunt+");'><span id='BillMemName1'></span><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;法條2<input type=text name='Rule2' size=5 class='btn1' onkeydown='KeyRule2("+cunt+");'>&nbsp;&nbsp;扣件2<input type=text name='Fastener2' size=2 class='btn1' onkeydown='KeyFastener2("+cunt+");'><input class='btn1' type='checkbox' name='chkBackBillBase' value='-1' onclick='funChkBackBillBase("+cunt+");'><input type='hidden' name='Sys_BackBillBase'>退件原因<input type=text name='Note' size=52 class='btn1' onkeydown='KeyNote("+cunt+");' disabled><hr>";
	}
}
function funChkSelt(){
	UrlStr="BillBaseCheckAcceptSendStyle.asp";
	myForm.action=UrlStr;
	myForm.target="ChkSelt";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funAcceptLoad(AcceptDate,UnitID,UnitName,Cmt,RecordDate){
	BillBaseOrder.innerHTML="<font size=3 color='Red'>『目前查閱"+AcceptDate+":"+UnitName+":"+Cmt+"件』</font>";

	runServerScript("getStopCarAcceptData.asp?AcceptDate="+AcceptDate+"&UnitID="+UnitID+"&RecordDate="+RecordDate);

	for(i=Cmt;i<myForm.item.length;i++){
		if(myForm.CarNo[i].value==''){
			break;
		}
		myForm.item[i].value='';
				
		myForm.CarNo[i].value='';

		myForm.illegalDate[i].value='';

		myForm.Rule1[i].value='';

		myForm.DriverName[i].value='';

		myForm.Fastener1[i].value='';

		myForm.BillMemName[i].value='';

		myForm.BillMemID1[i].value='';

		myForm.BillUnitID[i].value='';

		BillMemName1[i].innerHTML='';

		myForm.Sys_BackBillBase[i].value='';

		myForm.chkBackBillBase[i].checked=false;

		myForm.Note[i].disabled=true;

		myForm.Note[i].value='';

	}
}
function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}
function keyFunction(itemcnt) {
	//myForm.item[itemcnt-1].value=myForm.item[itemcnt-1].value.toUpperCase();
	if (event.keyCode==13||event.keyCode==9) {
		if (!chkBillNo(itemcnt)){
			alert("單號長度必須為9碼!!");
		}else{
			myForm.CarNo[itemcnt-1].focus();
		}
	}
}
function keyCarNo(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.illegalDate[itemcnt-1].focus();
	}
}
function KeyillegalDate(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Rule1[itemcnt-1].focus();
	}
}

function KeyRule1(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.DriverName[itemcnt-1].focus();
	}
}

function KeyDriverName(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Fastener1[itemcnt-1].focus();
	}
}

function KeyFastener1(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.BillMemName[itemcnt-1].focus();
	}
}
function KeyBillMem1(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Rule2[itemcnt-1].focus();
	}
}
function getBillMem1(itemcnt){
	runServerScript("CheckStopCarAcceptMemID.asp?LoginID="+myForm.BillMemName[itemcnt-1].value+"&itemcnt="+(itemcnt-1));
}
function KeyRule2(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.Fastener2[itemcnt-1].focus();
	}
}
function KeyFastener2(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}
function KeyNote(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.item[itemcnt].focus();
	}
}

function funChkBackBillBase(itemcnt){
	if(myForm.chkBackBillBase[itemcnt-1].checked){
		myForm.Note[itemcnt-1].disabled=false;
		myForm.Sys_BackBillBase[itemcnt-1].value="-1";
	}else{
		myForm.Note[itemcnt-1].disabled=true;
		myForm.Sys_BackBillBase[itemcnt-1].value="";
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
			if(myForm.CarNo[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 行車號不可空白!!");
				break;

			}else if(myForm.CarNo[i].value!=''&&myForm.CarNo[i].value.indexOf("-",0)<0){
				err=1;
				alert("第 "+(i+1)+" 行車號格式錯誤!!");
				break;

			}else if(myForm.Rule1[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 行法條不可空白!!");
				break;

			}else if(myForm.Rule1[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 行法條不可空白!!");
				break;

			}else if(myForm.DriverName[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 行違規人不可空白!!");
				break;

			}else if(myForm.Fastener1[i].value!=''&& "1,2,3,4,5,6,7,8,9,A,B,C,D,E,F,G,H,I,J,K,L".indexOf(myForm.Fastener1[i].value,0)<0){
				err=1;
				alert("第 "+(i+1)+" 行扣件代碼錯誤!!");
				break;
			
			}else if(myForm.BillMemID1[i].value==''||myForm.BillUnitID[i].value==''){
				err=1;
				alert("第 "+(i+1)+" 行舉發員警錯誤!!");
				break;

			}
			myForm.Note[i].disabled=false;
		}
	}
	if(myForm.AcceptDate.value==''){
		err=1;
		alert("點收日期不可空白!!");

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