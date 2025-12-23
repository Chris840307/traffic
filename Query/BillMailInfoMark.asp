<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>攔停/逕舉大宗條碼資料註記</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
if trim(request("DB_Add"))="Add" then
	tmp_ChkPasse=Split(trim(request("ChkPasse")),",")
	Sys_ChkPasse="":chkPassBillNo=""

	For i = 0 to Ubound(tmp_ChkPasse)
		If not ifnull(trim(tmp_ChkPasse(i))) Then

			if not ifnull(Sys_ChkPasse) then Sys_ChkPasse=Sys_ChkPasse&"','"
			Sys_ChkPasse=Sys_ChkPasse&trim(tmp_ChkPasse(i))
		
		End if 		
	Next
	
	if trim(UCase(request("Sys_BillNo1")))<>"" then
		BillStartNumber=trim(request("Sys_BillNo1")):BillEndNumber=trim(request("Sys_BillNo2"))
		if trim(BillEndNumber)="" then BillEndNumber=BillStartNumber

		if trim(BillStartNumber)<>"" then
			for i=1 to len(BillStartNumber)
				if IsNumeric(mid(BillStartNumber,i,1)) then
					Sno=MID(BillStartNumber,1,i-1)
					Tno=MID(BillStartNumber,i,len(BillStartNumber))
					exit for
				end if
			next
		end if
		if trim(BillEndNumber)<>"" then
			for i=1 to len(BillEndNumber)
				if IsNumeric(mid(BillEndNumber,i,1)) then
					Sno2=MID(BillEndNumber,1,i-1)
					Tno2=MID(BillEndNumber,i,len(BillEndNumber))
					exit for
				end if
			next
		end if

		If not ifnull(Sys_ChkPasse) Then chkPassBillNo=" and a.BillNo not in('"&Sys_ChkPasse&"')"	
		if Instr(request("Sys_BatchNumber"),"N")>0 then
			strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,c.UserMarkDate from DCILog a,BillBase b,billmailhistory c where SUBSTR(a.BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(a.BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"'"&chkPassBillNo&" and a.DciReturnStatusID not in('n','k') and a.BillSN=b.SN and a.BillNo=b.BillNo and a.billsn=c.billsn and a.billno=c.billno and b.RecordStateID=0 and NVL(b.EquiPmentID,1)<>-1 order by c.UserMarkDate"
		else
			strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,b.RecordDate from DCILog a,BillBase b where SUBSTR(a.BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(a.BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"'"&chkPassBillNo&" and a.DciReturnStatusID not in('N') and a.BillSN=b.SN and a.BillNo=b.BillNo and b.RecordStateID=0 and NVL(b.EquiPmentID,1)<>-1 order by b.RecordDate"
			
			if sys_City="基隆市" then
				strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,b.RecordDate,c.OwnerCounty from DCILog a,BillBase b,BillBaseDciReturn c where SUBSTR(a.BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(a.BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"'"&chkPassBillNo&" and a.DciReturnStatusID not in('N') and a.BillSN=b.SN and a.BillNo=b.BillNo and b.RecordStateID=0 and a.BillNO=c.BillNo and a.CarNo=c.CarNO and c.ExchangeTypeID='W' and NVL(b.EquiPmentID,1)<>-1 order by c.OwnerCounty,b.RecordDate"
			end if
		end if
	elseif trim(UCase(request("Sys_BatchNumber")))<>"" then
		tmp_BatchNumber=split(UCase(request("Sys_BatchNumber")),",")
		for i=0 to Ubound(tmp_BatchNumber)
			if i>0 then Sys_BatchNumber=trim(Sys_BatchNumber)&","
			if i=0 then
				Sys_BatchNumber=trim(Sys_BatchNumber)&tmp_BatchNumber(i)
			else
				Sys_BatchNumber=trim(Sys_BatchNumber)&"'"&tmp_BatchNumber(i)
			end if
			if i<Ubound(tmp_BatchNumber) then Sys_BatchNumber=trim(Sys_BatchNumber)&"'"
		next

		If not ifnull(Sys_ChkPasse) Then chkPassBillNo=" and a.BillNo not in('"&Sys_ChkPasse&"')"		
		if Instr(request("Sys_BatchNumber"),"N")>0 then
			strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,c.UserMarkDate from DCILog a,BillBase b,billmailhistory c where a.BatchNumber in('"&Sys_BatchNumber&"')"&chkPassBillNo&" and a.DciReturnStatusID not in('n','k') and a.BillSN=b.SN and a.BillNo=b.BillNo and a.billsn=c.billsn and a.billno=c.billno and b.RecordStateID=0 and NVL(b.EquiPmentID,1)<>-1 order by c.UserMarkDate"
		else
			strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,b.RecordDate from DCILog a,BillBase b where a.BatchNumber in('"&Sys_BatchNumber&"')"&chkPassBillNo&" and a.BillSN=b.SN and a.DciReturnStatusID not in('N') and a.BillNo=b.BillNo and b.RecordStateID=0 and NVL(b.EquiPmentID,1)<>-1 order by b.RecordDate"

			if sys_City="基隆市" then
				strSQL="select distinct a.BillSN,a.BillNo,a.CarNo,b.RecordDate,c.OwnerCounty from DCILog a,BillBase b,BillBaseDciReturn c where a.BatchNumber in('"&Sys_BatchNumber&"')"&chkPassBillNo&" and a.BillSN=b.SN and a.BillNo=b.BillNo and a.DciReturnStatusID not in('N') and b.RecordStateID=0 and a.BillNO=c.BillNo and a.CarNo=c.CarNO and c.ExchangeTypeID='W' and NVL(b.EquiPmentID,1)<>-1 order by c.OwnerCounty,b.RecordDate"
			end if
		end if
	end If 
	if trim(strSQL)<>"" then set rsload=conn.execute(strSQL)
	if trim(request("Sys_mailNumber1"))<>"" then
		Sys_mailNumber=0
		Sys_mailNumberUnitID=trim(request("Sys_mailNumber2"))
		cnt=0:cunt=len(trim(request("Sys_mailNumber1")))
		while Not rsload.eof
			Sys_mailNumber=CDbl(request("Sys_mailNumber1"))+cnt
			'cunt=len(trim(request("Sys_mailNumber1")))+len(cnt+1)-1
			
			Sys_mailNumber=right("00000000"&Sys_mailNumber,cunt)&Sys_mailNumberUnitID
			BillSN=rsload("BillSN"):BillNo=rsload("BillNo"):CarNo=rsload("CarNo")

			strSQL="select count(*) as cnt from BillMailHistory where BillSN="&BillSN
			set rscnt=conn.execute(strSQL)
			if CDbl(rscnt("cnt"))>0 then
				rscnt.close
				if Instr(request("Sys_BatchNumber"),"N")>0 then
					strSQL="Update BillMailHistory set StoreANDSendSendDate="&funGetDate(gOutDT(request("Sys_MailDate")),0)&",StoreAndSendMailNumber='"&Sys_mailNumber&"' where BillSN="&BillSN
					conn.execute(strSQL)
				else
					strSQL="Update BillMailHistory set MailDate="&funGetDate(gOutDT(request("Sys_MailDate")),0)&",MailNumber='"&Sys_mailNumber&"' where BillSN="&BillSN
					conn.execute(strSQL)
				end if
				
			else
				rscnt.close
				if Instr(request("Sys_BatchNumber"),"N")>0 then
					strSQL="Insert into BillMailHistory(BillSN,BillNo,CarNo,StoreANDSendSendDate,StoreAndSendMailNumber) values("&BillSN&",'"&BillNo&"','"&CarNo&"',"&funGetDate(gOutDT(request("Sys_MailDate")),0)&","&Sys_mailNumber&")"
					conn.execute(strSQL)
				else
					strSQL="Insert into BillMailHistory(BillSN,BillNo,CarNo,MailDate,MailNumber) values("&BillSN&",'"&BillNo&"','"&CarNo&"',"&funGetDate(gOutDT(request("Sys_MailDate")),0)&","&Sys_mailNumber&")"
					conn.execute(strSQL)
				end if
			end if

			cnt=cnt+1
			rsload.movenext
		wend
	elseif trim(strSQL)<>"" then
		while Not rsload.eof
			BillSN=rsload("BillSN"):BillNo=rsload("BillNo"):CarNo=rsload("CarNo")
			strSQL="select count(*) as cnt from BillMailHistory where BillSN="&BillSN
			set rscnt=conn.execute(strSQL)
			if CDbl(rscnt("cnt"))>0 then
				rscnt.close
				if Instr(request("Sys_BatchNumber"),"N")>0 then
					strSQL="Update BillMailHistory set StoreANDSendSendDate="&funGetDate(gOutDT(request("Sys_MailDate")),0)&" where BillSN="&BillSN
					conn.execute(strSQL)
				else
					strSQL="Update BillMailHistory set MailDate="&funGetDate(gOutDT(request("Sys_MailDate")),0)&" where BillSN="&BillSN
					conn.execute(strSQL)
				end if
			else
				rscnt.close
				if Instr(request("Sys_BatchNumber"),"N")>0 then
					strSQL="Insert into BillMailHistory(BillSN,BillNo,CarNo,StoreANDSendSendDate) values("&BillSN&",'"&BillNo&"','"&CarNo&"',"&funGetDate(gOutDT(request("Sys_MailDate")),0)&")"
					conn.execute(strSQL)
				else
					strSQL="Insert into BillMailHistory(BillSN,BillNo,CarNo,MailDate) values("&BillSN&",'"&BillNo&"','"&CarNo&"',"&funGetDate(gOutDT(request("Sys_MailDate")),0)&")"
					conn.execute(strSQL)
				end if
			end if
			rsload.movenext
		wend
	end if
	rsload.close
	text_mailNumber=""
	text_BillNo=Split(trim(request("item")),",")
	text_mailNumber=Split(request("mailNumber")&" ",",")
	text_ChkPassBillNo=Split(trim(request("Sys_ChkPassBillNo"))&" ",",")
	Sys_Now=now
	for i=0 to Ubound(text_BillNo)
		if trim(text_mailNumber(i))<>"" and trim(text_ChkPassBillNo(i))="" then
			if Instr(request("Sys_BatchNumber"),"N")>0 then
				txtSQL="select a.BillSN,a.BillNo,a.CarNo,c.UserMarkDate from DCILog a,BillBase b,billmailhistory c where a.BIllNo='"&trim(text_BillNo(i))&"' and a.DciReturnStatusID not in('n','k') and a.BillSN=b.SN and a.BillNo=b.BillNo and a.billsn=c.billsn and a.billno=c.billno and b.RecordStateID=0 order by c.UserMarkDate"
			else
				txtSQL="select a.BillSN,a.BillNo,a.CarNo,b.RecordDate from DCILog a,BillBase b where a.BIllNo='"&trim(text_BillNo(i))&"' and a.BillSN=b.SN and a.BillNo=b.BillNo and a.DciReturnStatusID not in('N') and b.RecordStateID=0 order by b.RecordDate"

				if sys_City="基隆市" then
					txtSQL="select a.BillSN,a.BillNo,a.CarNo,b.RecordDate,c.OwnerCounty from DCILog a,BillBase b,BillBaseDciReturn c where a.BIllNo='"&trim(text_BillNo(i))&"' and a.BillSN=b.SN and a.BillNo=b.BillNo and b.RecordStateID=0 and a.BillNO=c.BillNo and a.CarNo=c.CarNO and a.DciReturnStatusID not in('N') and c.ExchangeTypeID='W' order by c.OwnerCounty,a.BillNo"
				end if
			end if
			set rsload=conn.execute(txtSQL)
			while Not rsload.eof
				BillSN=rsload("BillSN"):BillNo=rsload("BillNo"):CarNo=rsload("CarNo")
				strSQL="select count(*) as cnt from BillMailHistory where BillSN="&BillSN
				set rscnt=conn.execute(strSQL)
				if CDbl(rscnt("cnt"))>0 then
					rscnt.close
					if Instr(request("Sys_BatchNumber"),"N")>0 then
						strSQL="Update BillMailHistory set StoreANDSendSendDate="&funGetDate(gOutDT(request("Sys_MailDate")),0)&",StoreAndSendMailNumber='"&trim(text_mailNumber(i))&"' where BillSN="&BillSN
						conn.execute(strSQL)
					else
						strSQL="Update BillMailHistory set MailDate="&funGetDate(gOutDT(request("Sys_MailDate")),0)&",MailNumber='"&trim(text_mailNumber(i))&"' where BillSN="&BillSN
						conn.execute(strSQL)
					end if
				else
					rscnt.close
					if Instr(request("Sys_BatchNumber"),"N")>0 then
						strSQL="Insert into BillMailHistory(BillSN,BillNo,CarNo,StoreANDSendSendDate,StoreAndSendMailNumber) values("&BillSN&",'"&BillNo&"','"&CarNo&"',"&funGetDate(gOutDT(request("Sys_MailDate")),0)&",'"&Sys_mailNumber&"')"
						conn.execute(strSQL)
					else
						strSQL="Insert into BillMailHistory(BillSN,BillNo,CarNo,MailDate,MailNumber) values("&BillSN&",'"&BillNo&"','"&CarNo&"',"&funGetDate(gOutDT(request("Sys_MailDate")),0)&",'"&Sys_mailNumber&"')"
						conn.execute(strSQL)
					end if
				end if
				cnt=cnt+1
				rsload.movenext
			wend
			rsload.close
		end if
	next
	%>
	<script language="JavaScript">
		alert ("產生完成!!");
		self.close();
	</script><%
end if
%>
<body>
<form name="myForm" method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>郵寄日期/大宗條碼資料註記</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>作業批號</font></td>
					<td>
						<input name="Sys_BatchNumber" class="btn1" value="<%=UCase(request("Sys_BatchNumber"))%>" type="text" size="15" maxlength="20">
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>舉發單號</font></td>
					<td><%response.write trim(request("Sys_BillNo1"))
						if trim(request("Sys_BillNo2"))<>"" then response.write "～"&trim(request("Sys_BillNo2"))%>
						<input name="Sys_BillNo1" class="btn1" value="<%=trim(request("Sys_BillNo1"))%>" type="hidden" size="15" maxlength="20">
						<input name="Sys_BillNo2" class="btn1" value="<%=trim(request("Sys_BillNo2"))%>" type="hidden" size="15" maxlength="20">
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4><strong>逕舉案件</strong> 郵寄日期</font></td>
					<td colspan="3">
						<input name="Sys_MailDate" class="btn1" value="<%
							if trim(request("Sys_MailDate"))<>"" then
								response.write request("Sys_MailDate")
							else
								response.write gInitDT(date)
							end if
						%>" type="text" size="10" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_MailDate');">
					</td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>貼條啟始號碼</font></td>
					<td>大宗碼：<input name="Sys_mailNumber1" class="btn1" value="<%=request("Sys_mailNumber1")%>" type="text" size="15" maxlength="20">　單位代碼及局碼：<input name="Sys_mailNumber2" class="btn1" value="<%=request("Sys_mailNumber2")%>" type="text" size="20" maxlength="30"><br><font color="red"><strong>*折封機逕舉舉發單不需填</strong></font></td>
				</tr>
				<tr>
					<td align="right" nowrap bgcolor="#FFFF99"><font size=4>是否自動帶入單號</font></td>
					<td>
						<input type="checkbox" name="ChkAutoBillNo" value="1" onclick="AutoBillNo();">
						自動帶入單號
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input name="btnadd" type="button" value=" 確 定 " onclick="funAdd();"> 
			<input name="btnexit" type="button" value=" 關 閉 " onclick="funExt();">
			<%
				Response.Write "<input type=""button"" name=""insert"" value=""再多30筆"" onClick=""insertRow(fmyTable)"">"
			%>
			<img src="space.gif" width="20" height="5">
		</td>
	</tr>
</table>

<input type="Hidden" name="DB_Selt" value="">
<input type="Hidden" name="chkcnt" value="">
<input type="Hidden" name="DB_Add" value="<%=request("DB_Add")%>">
<input type="Hidden" name="PBillSN" value="<%=request("PBillSN")%>">

<input type="Hidden" name="ChkPasse" value="">
<input type="Hidden" name="item" value="">
<input type="Hidden" name="mailNumber" value="">
<input type="Hidden" name="Sys_ChkPassBillNo" value="">

<hr>
</form>

<form name="AddForm" method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle"><strong>可依據需求設定或修正各舉發單大宗貼條碼 </strong><font size=2>(由下方輸入單號與貼條碼後點選確定按鈕即可，郵寄日期也可由上方欄位輸入)</font></td>
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
		</td>
	</tr>
</table>

</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
var cunt=0;
function AutoBillNo(){
	if (myForm.ChkAutoBillNo.checked==true){
		myForm.chkcnt.value=cunt;
	runServerScript("chkAutoBillNo.asp?Sys_BatchNumber="+myForm.Sys_BatchNumber.value+"&Sys_BillNo1="+myForm.Sys_BillNo1.value+"&Sys_BillNo2="+myForm.Sys_BillNo2.value+"&chkcnt="+myForm.chkcnt.value);
	}else{
		myForm.chkcnt.value="";
		for(i=0;i<AddForm.item.length;i++){
			AddForm.item[i].value='';
		}
	}
}

function insertRow(isTable){
	var objtemp="";
	<%
		cnt=29
	%>
	var cnt=<%=cnt%>;
	
	for(i=0;i<=cnt;i++){
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
		txtArea.innerHTML =cunt+".&nbsp;單號<input type=text name='item' size=10 class='btn1' onkeydown='keyFunction(1,"+cunt+");'>&nbsp;&nbsp;大宗掛號碼<input type=text name='mailNumber' size=10 onkeydown='keyFunction(2,"+cunt+");' class='btn1'>&nbsp;&nbsp;車號<input type=text name='CarNo' size=10 class='btn1' readOnly>&nbsp;&nbsp;略過<input class='btn1' type='checkbox' name='ChkPasse' value='' onclick=funChkPasse("+cunt+");><input type=hidden name='Sys_ChkPassBillNo' value=''>";
	}
}
function funChkPasse(itemcnt){
	if(AddForm.ChkPasse[itemcnt-1].checked){
		AddForm.ChkPasse[itemcnt-1].value=AddForm.item[itemcnt-1].value;
		AddForm.Sys_ChkPassBillNo[itemcnt-1].value=AddForm.item[itemcnt-1].value;
	}else{
		AddForm.ChkPasse[itemcnt-1].value='';
		AddForm.Sys_ChkPassBillNo[itemcnt-1].value='';
	}
}
function funkeyChk(obj) {
	obj.value=obj.value.replace(/[^\d]/g,'');
}
function keyFunction(objname,itemcnt) {
	if (event.keyCode==13||event.keyCode==9||AddForm.item[itemcnt-1].value>=9) {
		if(objname==1){
			if (chkBillNo(itemcnt)){
				if (AddForm.item[itemcnt-1].value!=''){
					myForm.chkcnt.value=itemcnt;
					runServerScript("chkBillMailNo.asp?BillNo="+AddForm.item[itemcnt-1].value);
				}
			}else{
				alert("單號長度必須為9碼!!");
			}
		}else if(objname==2){
			if(myForm.chkcnt.value<AddForm.item.length){
				AddForm.item[itemcnt].focus();
			}
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

function funSelt(){
	myForm.DB_Selt.value="Selt";
	myForm.submit();
}

<%
	Response.Write "for(j=0;j<=3;j++){insertRow(fmyTable);}"
%>

function funAdd(){
	var err=0;
	var ChkPasse="";
	var item="";
	var mailNumber="";
	var Sys_ChkPassBillNo="";

	if(myForm.Sys_MailDate.value==''){
		alert("郵寄日期必須填寫!!");
		err=1;
	}else if(!dateCheck(myForm.Sys_MailDate.value)){
		err=1;
		alert("郵寄日期輸入不正確!!");
	}else if(myForm.Sys_BatchNumber.value==""){
		err=1;
		alert("批號必須填寫!!");
	}else{
		for(i=0;i<AddForm.item.length;i++){
			if(AddForm.item[i].value!=''){
				if(item!=''){
					ChkPasse=ChkPasse+',';
					item=item+',';
					mailNumber=mailNumber+',';
					Sys_ChkPassBillNo=Sys_ChkPassBillNo+',';
				}
				ChkPasse=ChkPasse + AddForm.ChkPasse[i].value;
				item=item + AddForm.item[i].value;
				mailNumber=mailNumber + AddForm.mailNumber[i].value;
				Sys_ChkPassBillNo=Sys_ChkPassBillNo + AddForm.Sys_ChkPassBillNo[i].value;
			}
		}

		myForm.ChkPasse.value=ChkPasse;
		myForm.item.value=item;
		myForm.mailNumber.value=mailNumber;
		myForm.Sys_ChkPassBillNo.value=Sys_ChkPassBillNo;

		myForm.DB_Add.value="Add";
		myForm.submit();
	}
}
function funExt() {
	if(confirm("是否關閉維護系統?")){
		window.close();
	}
}
</script>
<%conn.close%>