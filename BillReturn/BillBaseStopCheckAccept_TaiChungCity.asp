<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE>攔停登記簿系統</TITLE>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
<!--
.style1 {font-size: 10px; color:#ff0000; }
.btn3{
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
}
-->
</style>
</HEAD>
<BODY>
<%
Server.ScriptTimeout=6000

Function ChkNum(strValue)
	if ISNull(strValue) or trim(strValue)="" or IsEmpty(strValue) then
		ChkNum="null"
	else
		ChkNum=strValue
	end if
End Function

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

strCity="select value from Apconfigure where id=3"
set rsCity=conn.execute(strCity)
sys_RuleVer=trim(rsCity("value"))
rsCity.close

if trim(request("DB_Selt"))="Selt" then
	Sys_BatChNumber=trim(Request("Sys_BatChNumber"))
	Sys_Old_BatchNumber=trim(Request("Old_BatchNumber"))
	Sys_BillNo=Split(Ucase(trim(request("BillNo"))),",")
	Sys_Fastener1=Split(trim(request("FastenerTypeID1")),",")
	Sys_Fastener2=Split(trim(request("FastenerTypeID2")),",")
	Sys_Note=Split(trim(request("Note")),",")
	Sys_chkBackBillBase=Split(trim(request("Sys_BackBillBase")),",")
	Sys_BillBaseDel=Split(trim(request("Sys_BillBaseDel")),",")
	Sys_AcceptDate=Split(trim(request("AcceptDate")),",")
	Sys_BackDate=Split(trim(request("BackDate")),",")
	old_BillNo=Split(trim(request("old_BillNo")),",")

	Sys_Now=now
	
	For i = 0 to Ubound(Sys_BillNo)
		If not ifnull(Sys_BillNo(i)) Then
			
			If not ifnull(old_BillNo(i)) Then

				strSQL="Update BillStopCarAccept set BatchNumber='"&Sys_BatChNumber&"',BillNo='"&trim(Sys_BillNo(i))&"',FastenerTypeID1='"&trim(Sys_Fastener1(i))&"',FastenerTypeID2='"&trim(Sys_Fastener2(i))&"',Note='"&trim(Sys_Note(i))&"',RecordStateID=0,AcceptDate="&funGetDate(gOutDT(Sys_AcceptDate(i)),0)&" where BillNo='"&trim(old_BillNo(i))&"' and BatchNumber='"&Sys_Old_BatchNumber&"'"

				conn.execute(strSQL)
				
			else
				Sys_Now=DateAdd("s",1,Sys_Now)

				strSQL="insert into BillStopCarAccept(BillNo,BATCHNUMBER,FastenerTypeID1,FastenerTypeID2,NOTE,RECORDMEMBERID1,RECORDSTATEID,RECORDDATE1,AcceptDate) values('"&trim(Sys_BillNo(i))&"','"&trim(Sys_BatChNumber)&"','"&trim(Sys_Fastener1(i))&"','"&trim(Sys_Fastener2(i))&"','"&trim(Sys_Note(i))&"',"&Session("User_ID")&",0,"&funGetDate((Sys_Now),1)&","&funGetDate(gOutDT(Sys_AcceptDate(i)),0)&")"

				conn.execute(strSQL)
			End if 

			If not ifnull(Sys_chkBackBillBase(i)) Then
				strSQL="Update BillStopCarAccept set RecordStateID=-1,BackDate="&funGetDate(gOutDT(Sys_BackDate(i)),0)&" where BillNo='"&trim(Sys_BillNo(i))&"' and BatchNumber='"&Sys_BatChNumber&"'"

				conn.execute(strSQL)
			End if 

			If not ifnull(Sys_BillBaseDel(i)) Then
				strSQL="delete BillStopCarAccept where BillNo='"&trim(Sys_BillNo(i))&"' and BatchNumber='"&Sys_BatChNumber&"'"

				conn.execute(strSQL)
			End if 

		end if
	Next
	Response.write "<script>"
	Response.Write "alert('簽收送件完成！');"
	If ifnull(old_BillNo(0)) Then Response.Write "location='BillBaseStopCheckAccept_TaiChungCity.asp';"
	Response.write "</script>"
end if
%>
<form name="myForm" method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="37" bgcolor="#FFCC33" class="pagetitle">
			<strong>攔停登記簿系統</strong>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						<table border="0">
							<tr>
								<td><%
										tmp_AcceptDate1="":tmp_AcceptDate2=""
										tmp_BackDate1="":tmp_BackDate2=""

										If not ifnull(Request("Sys_AcceptDate1")) Then

											tmp_AcceptDate1=Request("Sys_AcceptDate1")
											tmp_AcceptDate2=Request("Sys_AcceptDate2")
										else

											tmp_AcceptDate1=Request("myForm_AcceptDate1")
											tmp_AcceptDate2=Request("myForm_AcceptDate2")

										End if 

										
										If not ifnull(Request("Sys_BackDate1")) Then

											tmp_BackDate1=Request("Sys_BackDate1")
											tmp_BackDate2=Request("Sys_BackDate2")
										else

											tmp_BackDate1=Request("myForm_BackDate1")
											tmp_BackDate2=Request("myForm_BackDate2")
										End if 

									%>
									查詢批號：
									<input type="text" name='Search_BatChNumber' size="8" class='btn1' value="<%=Request("Search_BatChNumber")%>">
									&nbsp;&nbsp;
									單號：
									<input type="text" name='Search_BillNo' size="8" class='btn1' value="<%=Request("Search_BillNo")%>"><br>
									收件日期：
									<input type="text" name='myForm_AcceptDate1' size="8" class='btn1' maxlength='7' value="<%=tmp_AcceptDate1%>" onkeyup="chknumber(this);">
									<input type="button" name="datestr" value="..." class="btn3" style="width:15px; height:20px;" onclick="OpenWindow('myForm_AcceptDate1');">
									～
									<input type="text" name='myForm_AcceptDate2' size="8" class='btn1' maxlength='7' value="<%=tmp_AcceptDate2%>" onkeyup="chknumber(this);">
									<input type="button" name="datestr" value="..." class="btn3" style="width:15px; height:20px;" onclick="OpenWindow('myForm_AcceptDate2');">
									&nbsp;&nbsp;單位代碼：
									<input type="text" name='Search_Unit' size="8" class='btn1' value="<%=Request("Search_Unit")%>"><br>
									退件日期：
									<input type="text" name='myForm_BackDate1' size="8" class='btn1' maxlength='7' value="<%=tmp_BackDate1%>" onkeyup="chknumber(this);">
									<input type="button" name="datestr" value="..." class="btn3" style="width:15px; height:20px;" onclick="OpenWindow('myForm_BackDate1');">
									～
									<input type="text" name='myForm_BackDate2' size="8" class='btn1' maxlength='7' value="<%=tmp_BackDate2%>" onkeyup="chknumber(this);">
									<input type="button" name="datestr" value="..." class="btn3" style="width:15px; height:20px;" onclick="OpenWindow('myForm_BackDate2');">
									<input type="button" name="selt" value="查詢" onclick="funQry();">
									<input type="button" name="selt" value="退件清冊" onclick="funBackList();">
									
									<input type="button" name="cancel" value="清除" onClick="location='BillBaseStopCheckAccept_TaiChungCity.asp'">
									<table style="width:600px;" border="1" cellpadding="0" cellspacing="0">&nbsp;&nbsp;
										<tr bgcolor="#EBFBE3" align="center" style="height:30px;">
											<td style="width:100px;">批號</td>
											<!--<td style="width:77px;">建檔人</td>-->
											<td style="width:50px;">件數</td>
											<td style="width:50px;">退件數</td>
											<td style="width:105px;">操作</td>
										</tr>
										<tr>
											<td colspan="5">
												<Div style="overflow:auto;width:100%;height:100px;background:#FFFFFF">
												<table width="100%" border="1" cellpadding="1" cellspacing="0"><%
												CarCode=",1,2,3,4,5,6,":chkExp=1:chkTime=1

												strWhere=""

												If not ifnull(tmp_AcceptDate2) Then
													ExchangeDate1=gOutDT(tmp_AcceptDate1)&" 0:0:0"
													ExchangeDate2=gOutDT(tmp_AcceptDate2)&" 23:59:59"

													strWhere="where BatchNumber in(select distinct BatchNumber from BillStopCarAccept where AcceptDate between to_date('"&ExchangeDate1&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&ExchangeDate2&"','YYYY/MM/DD/HH24/MI/SS'))"
												End if 

												If not ifnull(Request("Search_BatChNumber")) Then

													If strWhere <> "" Then 
														strWhere=strWhere&" and "
													else
														strWhere=strWhere&" where "
													end if
													
													If instr(Request("Search_BatChNumber"),",") > 0 Then

														strWhere=strWhere&" BatchNumber in('"&trim(Request("Search_BatChNumber"))&"')"
													else

														strWhere=strWhere&" BatchNumber='"&trim(Request("Search_BatChNumber"))&"'"
													End if 

													'strWhere=strWhere&" ',"&trim(Request("Search_BatChNumber"))&",' like '%,'||BatchNumber||',%'"
												End if 

												If not ifnull(Request("Search_BillNo")) Then

													If strWhere <> "" Then 
														strWhere=strWhere&" and "
													else
														strWhere=strWhere&" where "
													end if

													strWhere=strWhere&"BatchNumber in(select distinct BatchNumber from BillStopCarAccept where BillNo='"&trim(Request("Search_BillNo"))&"')"
												End if 

												If not ifnull(Request("Search_Unit")) Then

													If strWhere <> "" Then 
														strWhere=strWhere&" and "
													else
														strWhere=strWhere&" where "
													end if
													
													If instr(Request("Search_Unit"),",") > 0 Then

														strWhere=strWhere&" SubStr(BatchNumber,1,2) in('"&trim(Request("Search_Unit"))&"')"
													else

														strWhere=strWhere&" SubStr(BatchNumber,1,2)='"&trim(Request("Search_Unit"))&"'"
													End if 

													'strWhere=strWhere&" ',"&trim(Request("Search_Unit"))&",' like '%,'||SubStr(BatchNumber,1,2)||',%'"
												End if 

												If not ifnull(tmp_BackDate2) Then

													ExchangeDate1=gOutDT(tmp_BackDate1)&" 0:0:0"
													ExchangeDate2=gOutDT(tmp_BackDate2)&" 23:59:59"

													If strWhere <> "" Then 
														strWhere=strWhere&" and "
													else
														strWhere=strWhere&" where "
													end if

													strWhere=strWhere&"BatchNumber in(select distinct BatchNumber from BillStopCarAccept where BackDate between to_date('"&ExchangeDate1&"','YYYY/MM/DD/HH24/MI/SS') and to_date('"&ExchangeDate2&"','YYYY/MM/DD/HH24/MI/SS'))"
												End if 
												
												If strWhere = "" Then
													strWhere="where BatchNumber in(select distinct BatchNumber from BillStopCarAccept where RecordDate1 between to_date('"&date&" 00:00:00','YYYY/MM/DD/HH24/MI/SS') and to_date('"&date&" 23:59:59','YYYY/MM/DD/HH24/MI/SS'))"
												End if 

												strSQL="select BatchNumber,Sum(DeCode(RecordStateID,0,1,0)) SCnt,Sum(DeCode(RecordStateID,-1,1,0)) DCnt from BillStopCarAccept "&strWhere&" group by BatchNumber"

												set rs=conn.execute(strSQL)
												While not rs.eof
													Response.Write "<tr align=""center"">"

													Response.Write "<td style=""width:123px;"">"&rs("BatchNumber")&"</th>"
													'Response.Write "<td style=""width:95px;"">"&rs("ChName")&"</th>"
													Response.Write "<td style=""width:61px;"">"&rs("SCnt")&"</th>"
													Response.Write "<td style=""width:60px;"">"&rs("DCnt")&"</th>"
													Response.Write "<td style=""width:130px;"">"
													Response.Write "<input type=""button"" name=""btnAcc"" class=""btn3"" style=""width:80px;height:25px;font-size:16px;"" value=""詳細"" onclick=""funAcceptLoad('"&rs("BatchNumber")&"');"">&nbsp;&nbsp;"

													Response.Write "<input type=""button"" name=""btnDec"" class=""btn3"" style=""width:80px;height:25px;font-size:16px;"" value=""清冊"" onclick=""funPrintDec('"&rs("BatchNumber")&"');"""
													If cdbl(rs("SCnt")) = 0 Then Response.Write "disabled"
													Response.Write ">"
													

													'Response.Write "&nbsp;<input type=""button"" name=""Update"" class=""btn3"" style=""width:40px;height:25px;font-size:16px;"" value=""列印"" class=""btn3"" onclick=""funAcceptRunList('"&rs("BatchNumber")&"');"">"

													Response.Write "</td></tr>"

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
								預設收件日統一為&nbsp;
								<input name="Def_AcceptDate" type="text" class="btn1" size="10" maxlength="7"  onkeyup='chknumber(this);'>
								&nbsp;&nbsp;&nbsp;
								<input type="button" name="btnDefu" value="確定" onclick="funDef_AcceptDate();"><br>
								預設退件日統一為&nbsp;
								<input name="Def_BackDate" type="text" class="btn1" size="10" maxlength="7"  onkeyup='chknumber(this);'>
								&nbsp;&nbsp;&nbsp;
								<input type="button" name="btnDefu" value="確定" onclick="funDef_BackDate();"><br>


								<hr>
								批號：<input type="text" name='Sys_BatChNumber' ID='Sys_BatChNumber' size="18" class='btn1' maxlength='8' value="" onkeyup="JS_chkBatchNumber(this.value);"><br>
								<span ID="chkBatchNumberMessg" style="font-family: 標楷體; font-size: 18px; color: #ff0000;"></span>
								<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
								</td>
							</tr>

						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">攔停登記簿紀錄列表 ( 輸入完成按Enter可自動跳到下一格 )
		<img src="space.gif" width="29" height="8">
		<input type="button" name="insert" class="btn3" style="width:80px;height:25px;font-size:16px;" value="再多50筆" onClick="insertRow(fmyTable)">
		&nbsp;&nbsp;&nbsp;
		<input type="button" name="btnOK1" class="btn3" style="width:80px;height:25px;font-size:16px;" value="確定存檔" onclick="funSelt();">
		<br>
			<%
			Response.Write "<BR><B>"
			Response.Write "車種代碼：『1汽車、2拖車、3重機/550cc以上、4輕機、5動力機械、6臨時車牌』；"
			Response.Write "</B>"
			%>
		</td>
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
			<input type="button" name="btnOK1" class="btn3" style="width:80px;height:25px;font-size:16px;" value="確定存檔" onclick="funSelt();">
			<input type="button" name="insert2" class="btn3" style="width:80px;height:25px;font-size:16px;" value="再多50筆" onClick="insertRow(fmyTable)">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="">
<input type="Hidden" name="chkcnt" value="">
<input type="Hidden" name="DB_AcceptDate" value="">
<input type="Hidden" name="DB_BillUnitID" value="">
<input type="Hidden" name="DB_RecordMemberID1" value="">
<input type="Hidden" name="DB_RecordMemberID2" value="">
<input type="Hidden" name="chkBatchNumber" value="">
<input type="Hidden" name="Old_BatchNumber" value="">
</form>

<form name="upForm" method="post">
	<input type="Hidden" name="Sys_AcceptDate1" value="">
	<input type="Hidden" name="Sys_AcceptDate2" value="">
	<input type="Hidden" name="Sys_BackDate1" value="">
	<input type="Hidden" name="Sys_BackDate2" value="">
	<input type="Hidden" name="Search_BatChNumber" value="">
	<input type="Hidden" name="Search_BillNo" value="">
	<input type="Hidden" name="Search_Unit" value="">
	<input type="Hidden" name="BatChNumber" value="">
	<input type="Hidden" name="DB_Selt" value="">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="JavaScript">
var cunt=0;

function insertRow(isTable){
	for(i=0;i<=49;i++){
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

		if(cnt_num%2==0){txtArea.style.backgroundColor ="#EEEEEE";}

		txtArea.innerHTML =
		"<b>" + cnt_num + "</b>&nbsp;&nbsp;"+
		"<span style='color:#ff0000;'>*</span>" +
		"單號&nbsp;&nbsp;<input type=text name='BillNo' size=10 class='btn1' onkeyup='UpperCase(this);' onkeydown='keyBillNo("+cunt+");'>" +
		"&nbsp&nbsp" +
		"收件日<input type=text name='AcceptDate' size=10 class='btn1' onkeyup='chknumber(this);' onkeydown='KeyAcceptDate("+cunt+");'>" +
		"<br>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp" +
		"扣件1<select name='FastenerTypeID1'><option value=''>請選擇</option><option value='1'>行照</option><option value='2'>駕照</option><option value='3'>告發單</option>" +
		"<option value='4'>牌一面</option><option value='5'>牌二面</option><option value='6'>空車</option><option value='7'>其它</option><option value='8'>錄影</option>" +
		"<option value='9'>照片</option><option value='A'>汽駕</option><option value='B'>機駕</option><option value='C'>本牌乙</option><option value='D'>本牌二</option>" +
		"<option value='E'>移牌乙</option><option value='F'>移牌二</option><option value='G'>臨時牌</option><option value='H'>試車牌</option><option value='I'>執登證</option>" +
		"<option value='J'>測速器</option><option value='K'>喇叭</option><option value='L'>車鑰匙</option></select>" +
		"&nbsp&nbsp" +
		"扣件2<select name='FastenerTypeID2'><option value=''>請選擇</option><option value='1'>行照</option><option value='2'>駕照</option><option value='3'>告發單</option>" +
		"<option value='4'>牌一面</option><option value='5'>牌二面</option><option value='6'>空車</option><option value='7'>其它</option><option value='8'>錄影</option>" +
		"<option value='9'>照片</option><option value='A'>汽駕</option><option value='B'>機駕</option><option value='C'>本牌乙</option><option value='D'>本牌二</option>" +
		"<option value='E'>移牌乙</option><option value='F'>移牌二</option><option value='G'>臨時牌</option><option value='H'>試車牌</option><option value='I'>執登證</option>" +
		"<option value='J'>測速器</option><option value='K'>喇叭</option><option value='L'>車鑰匙</option></select>" +
		"&nbsp;&nbsp;" +		
		"審核說明<input type=text name='Note' size=32 class='btn1' onkeydown='KeyNote("+cunt+");'>" +
		"&nbsp;&nbsp;" +
		"退件<input class='btn1' type='checkbox' name='chkBackBillBase' value='-1' onclick='funChkBackBillBase("+cunt+");'>" +
		"&nbsp&nbsp" +
		"退件日<input type=text name='BackDate' size=10 class='btn1' onkeyup='chknumber(this);' onkeydown='KeyBackDate("+cunt+");'>" +
		"&nbsp&nbsp" +
		"刪除<input class='btn1' type='checkbox' name='chkBillBaseDel' value='-1' onclick='funChkBillDel("+cunt+");'>" +
		"<input type='hidden' name='Sys_BackBillBase'>" +
		"<input type='hidden' name='Sys_BillBaseDel'>" +
		"<input type='hidden' name='old_BillNo' value=''>" +
		"<hr>";
	}
}
function funChkSelt(){
	UrlStr="BillBaseCheckRunAcceptSendStyle_miaoli.asp";
	myForm.action=UrlStr;
	myForm.target="ChkSelt";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funAcceptRunList(AcceptDate,BillUnitID,RecordMemberID1,chkType){
	var UnitLevelID='<%=session("UnitLevelID")%>';

	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordMemberID1.value="";
	myForm.DB_RecordMemberID2.value="";

	if(chkType=='0'){
		myForm.DB_RecordMemberID1.value=RecordMemberID1;
	}else{
		myForm.DB_RecordMemberID2.value=RecordMemberID1;
	}

	UrlStr="AcceptRunList.asp";
	
	myForm.action=UrlStr;
	myForm.target="PrintAccept";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPrintBatOver(){
	myForm.DB_AcceptDate.value="";
	myForm.DB_BillUnitID.value="";
	myForm.DB_RecordMemberID1.value="";
	myForm.DB_RecordMemberID2.value="";

	myForm.DB_Selt.value="PrintBatOver";
	myForm.submit();
}

function funPrintDec(BatNum){

	upForm.BatChNumber.value=BatNum;

	UrlStr="AcceptStopList_TaiChungCity.asp";
	
	upForm.action=UrlStr;
	upForm.target="PrintAccept";
	upForm.submit();
	upForm.action="";
	upForm.target="";
}

function funSaveCheck(AcceptDate,BillUnitID,RecordMemberID1,chkType){
	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordMemberID1.value="";
	myForm.DB_RecordMemberID2.value="";

	if(chkType=='0'){
		myForm.DB_RecordMemberID1.value=RecordMemberID1;
	}else{
		myForm.DB_RecordMemberID2.value=RecordMemberID1;
	}

	myForm.DB_Selt.value="SaveCheck";
	myForm.submit();
}

function funSaveBat(){

	myForm.DB_Selt.value="SaveBat";
	myForm.submit();
}

function JS_chkBatchNumber(BatChNumber){	
	if(BatChNumber.length > 7){
		chkBatchNumberMessg.innerHTML="";
		myForm.chkBatchNumber.value=0;
		runServerScript("chkStopAcceptBatchNumber.asp?BatChNumber="+BatChNumber);
	}
}

function funAcceptLoad(BatChNumber){
	myForm.Sys_BatChNumber.value=BatChNumber;

	myForm.Old_BatchNumber.value=BatChNumber;

	for(i=0;i<myForm.BillNo.length;i++){

		myForm.BillNo[i].value='';

		myForm.AcceptDate[i].value='';

		myForm.FastenerTypeID1[i].value='';

		myForm.FastenerTypeID2[i].value='';
		
		myForm.Sys_BackBillBase[i].value='';
		
		myForm.chkBackBillBase[i].checked=false;

		myForm.Sys_BillBaseDel[i].value='';

		myForm.chkBillBaseDel[i].checked=false;

		myForm.Note[i].value='';

		myForm.BackDate[i].value='';
		
		myForm.old_BillNo[i].value='';

	}

	runServerScript("getStopCarAcceptData_TaiChungCity.asp?BatchNumber="+BatChNumber+"&objcnt="+myForm.BillNo.length);

	
}
function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}

function keyBillNo(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		if (myForm.BillNo[itemcnt-1].value.length!=9){
			alert("單號長度必須為9碼!!");
		}else{
			myForm.BillNo[itemcnt].focus();
		}			
	}
}

function KeyAcceptDate(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.BillNo[itemcnt].focus();
	}
}

function funDef_AcceptDate(itemcnt){
	for(i=0;i<myForm.AcceptDate.length;i++){
		myForm.AcceptDate[i].value=myForm.Def_AcceptDate.value;
	}
}

function KeyBackDate(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.BillNo[itemcnt].focus();
	}
}

function funDef_BackDate(itemcnt){
	for(i=0;i<myForm.BackDate.length;i++){
		myForm.BackDate[i].value=myForm.Def_BackDate.value;
	}
}

function KeyCarSimpleID(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.BillNo[itemcnt].focus();
	}
}


function KeyNote(itemcnt){
	if (event.keyCode==13||event.keyCode==9){
		myForm.BillNo[itemcnt].focus();
	}
}

function funChkBackBillBase(itemcnt){
	if(myForm.chkBackBillBase[itemcnt-1].checked){
		myForm.Sys_BackBillBase[itemcnt-1].value="-1";
	}else{
		myForm.Sys_BackBillBase[itemcnt-1].value="";
	}
}

function funChkBillDel(itemcnt){
	if(myForm.chkBillBaseDel[itemcnt-1].checked){
		myForm.Sys_BillBaseDel[itemcnt-1].value="-1";
	}else{
		myForm.Sys_BillBaseDel[itemcnt-1].value="";
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

function funBatBack(){
	for(i=0;i<myForm.item.length;i++){
		if(myForm.illegalDate[i].value!=''){
			myForm.chkBackBillBase[i].click();
			funChkBackBillBase(i+1);

			if(myForm.chkBackBillBase[i].checked){
				myForm.Note[i].value=myForm.BatNote.value;
			}else{
				myForm.Note[i].value="";
			}
		}
	}
}

function funSelt(){
	var err=0;
	var errmsg="";

	if(myForm.Sys_BatChNumber.value==''){
		err=1;
		errmsg=errmsg+"批號必須填寫!!\n";

	}else if(myForm.Sys_BatChNumber.value.length<8){
		err=1;
		errmsg=errmsg+"批號格式不正確!!\n";
	}

	if(myForm.chkBatchNumber.value > 0 && myForm.Old_BatchNumber.value==""){
		err=1;
		errmsg=errmsg+"批號已存在!!\n";

	}else if(myForm.chkBatchNumber.value > 0 && myForm.Old_BatchNumber.value!="" && myForm.Old_BatchNumber.value!=myForm.Sys_BatChNumber.value){
		err=1;
		errmsg=errmsg+"批號已存在!!\n";
	}

	for(i=0;i<myForm.BillNo.length;i++){
		if(myForm.BillNo[i].value!=''){

			if (myForm.BillNo[i].value.length!=9){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行單號長度必須為9碼!!\n";
			}

			if(myForm.AcceptDate[i].value==''){
				err=1;
				errmsg=errmsg+"第 "+(i+1)+" 行收件日期不可空白!!\n";
			}

			if(myForm.chkBackBillBase[i].checked==true){
				if(myForm.BackDate[i].value==''){
					err=1;
					errmsg=errmsg+"第 "+(i+1)+" 行退件日期不可空白!!\n";
				}
			}

			if(err!=0){
				break;
			}
		}
	}

	if(myForm.BillNo[0].value!=''){
		if(err==0){
			if(myForm.Old_BatchNumber.value==""){myForm.Old_BatchNumber.value=myForm.Sys_BatChNumber.value;}
			myForm.DB_Selt.value="Selt";
			myForm.submit();
		}else{
			alert(errmsg);
		}
	}
}

function funQry(){
	var err=0;
	var errmsg="";
	upForm.Sys_AcceptDate1.value=myForm.myForm_AcceptDate1.value;
	upForm.Sys_AcceptDate2.value=myForm.myForm_AcceptDate2.value;
	upForm.Search_BatChNumber.value=myForm.Search_BatChNumber.value;
	upForm.Search_BillNo.value=myForm.Search_BillNo.value;
	upForm.Search_Unit.value=myForm.Search_Unit.value;
	upForm.Sys_BackDate1.value=myForm.myForm_BackDate1.value;
	upForm.Sys_BackDate2.value=myForm.myForm_BackDate2.value;
	upForm.DB_Selt.value="Qry";
	upForm.submit();

}

function funBackList(){
	var err=0;
	var errmsg="";
	upForm.Sys_AcceptDate1.value=myForm.myForm_AcceptDate1.value;
	upForm.Sys_AcceptDate2.value=myForm.myForm_AcceptDate2.value;
	upForm.Search_BatChNumber.value=myForm.Search_BatChNumber.value;
	upForm.Search_BillNo.value=myForm.Search_BillNo.value;
	upForm.Search_Unit.value=myForm.Search_Unit.value;
	upForm.Sys_BackDate1.value=myForm.myForm_BackDate1.value;
	upForm.Sys_BackDate2.value=myForm.myForm_BackDate2.value;

	if(upForm.Sys_BackDate1.value!=''&&upForm.Sys_BackDate2.value!=''){

		UrlStr="AcceptStopBackList_TaiChungCity.asp";
	
		upForm.action=UrlStr;
		upForm.target="PrintAccept";
		upForm.submit();
		upForm.action="";
		upForm.target="";
	}

}

function UpperCase(obj){
	if(obj.value!=obj.value.toUpperCase()){
		obj.value=obj.value.toUpperCase();
	}

	if (obj.value.length>9){
		obj.value=obj.value.substr(0,9);
	}
}


for(j=0;j<=3;j++){
	insertRow(fmyTable);
}
</script>

<%
if trim(request("DB_Selt"))="Selt" then
	If not ifnull(old_BillNo(0)) Then
		Response.write "<script>"
		Response.Write "funAcceptLoad('"&Sys_BatChNumber&"');"
		Response.write "</script>"
	end If 
end if
%>