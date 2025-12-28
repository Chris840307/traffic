<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>送達紀錄</title>
<!--#include virtual="traffic/Common/css.txt"-->
</head>
<%
Sys_SenderMemID=0

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

if trim(request("DB_State"))="Update" then
	ArgueDate1=gOutDT(request("edit_ArrivedDate"))
	MailDate1=gOutDT(request("edit_MailDate"))
	If not ifnull(request("edit_SenderMemID")) Then Sys_SenderMemID=request("edit_SenderMemID")
	strSQL="Update PassersEndArrived set ArrivedDate="&funGetDate(ArgueDate1,0)&",SenderMemID="&Sys_SenderMemID&",MailDate="&funGetDate(MailDate1,0)&",SendMailStation='"&trim(request("edit_SendMailStation"))&"',ArriveType="&trim(request("edit_ArriveType"))&",ReturnResonID='"&trim(request("edit_ReturnResonID"))&"',Note='"&trim(request("edit_Note"))&"' where SN="&request("SN")
	conn.execute(strSQL)
end If 

if trim(request("DB_State"))="Add" then
	strSQL="select Max(SN) as cnt from PassersEndArrived"
	set rscnt=conn.execute(strSQL)
	PasserSN=1
	if Not isnull(rscnt("cnt")) then
		PasserSN=cdbl(rscnt("cnt"))+1
	end if
	rscnt.close
	ArgueDate1=gOutDT(request("ArrivedDate"))
	MailDate1=gOutDT(request("MailDate"))
	If not ifnull(request("SenderMemID")) Then Sys_SenderMemID=request("SenderMemID")
	strSQL="insert into PassersEndArrived(SN,PasserSN,ArrivedDate,SenderMemID,RecordmemberID,MailDate,SendMailStation,ArriveType,ReturnResonID,Note) values("&PasserSN&","&request("BillSN")&","&funGetDate(ArgueDate1,0)&","&Sys_SenderMemID&","&Session("User_ID")&","&funGetDate(MailDate1,0)&",'"&trim(request("Sys_SendMailStation"))&"',"&trim(request("Sys_ArriveType"))&",'"&trim(request("Sys_ReturnResonID"))&"','"&trim(request("Sys_Note"))&"')"
	conn.execute(strSQL)
end if
if request("DB_State")="Del" then
	strSQL="Delete from PassersEndArrived where SN="&request("SN")
	conn.execute strSQL
end If 

strSQL="select billno,illegaldate from passerbase where sn="&request("BillSN")
set rspass=conn.execute(strSQL)
sys_billno=rspass("billno")
sys_illegaldate=gInitDT(rspass("illegaldate"))
rspass.close

strSQL="Select a.*,b.Chname from PassersEndArrived a,MemberData b where a.SenderMemID=b.MemberID(+) and a.PasserSN="&request("BillSN")
set rsload=conn.execute(strSQL)
%>
<BODY onkeydown="KeyDown()">
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33">送達紀錄　　　<font color="red"><b>違規日期：<%=sys_illegaldate%>
		&nbsp;&nbsp;&nbsp;&nbsp;
		單號：<%=sys_billno%></b></font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td>
						<table width="100%" border="0">
							<tr>
								<td>送達日期</td>
								<td nowrap>
									<input name="ArrivedDate" class="btn1" type="text" value="<%=trim(request("ArrivedDate"))%>" size="10" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
									<input type="button" name="datestr" value="..." onclick="OpenWindow('ArrivedDate');">
								</td>
								<td nowrap>
									送達單位
								</td>
								<td>
									<%=UnSelectUnitOption("UnitID","SenderMemID")%>
								</td>
								<td nowrap>
									送達人員
								</td>
								<td>
									<%=UnSelectMemberOption("UnitID","SenderMemID")%>
								</td>
							</tr>
							<tr>
								<td nowrap>
									大宗掛號碼
								</td>
								<td>
									<input name="Sys_SendMailStation" class="btn1" type="text" value="<%=trim(request("Sys_SendMailStation"))%>" size="10" maxlength="50">
								</td>
								<td>
									送達類別
								</td>
								<td>
									<select name="Sys_ArriveType" class="btn1">
										<option Value="0"<%If trim(request("Sys_ArriveType"))="0" Then response.Write(" selected")%>>裁決</option>
										<option Value="1"<%If trim(request("Sys_ArriveType"))="1" Then response.Write(" selected")%>>催告</option>
										<option Value="2"<%If trim(request("Sys_ArriveType"))="2" Then response.Write(" selected")%>>郵寄</option>
									</select>
								</td>
								<td>
									送達狀況
								</td>
								<td>
									<select name="Sys_ReturnResonID" onchange="AcceptReson();" class="btn1">
										<option Value="">請選擇</option>
										<option Value="2"<%If trim(request("Sys_ReturnResonID"))="2" Then response.Write(" selected")%>>寄存</option>
										<option Value="0"<%If trim(request("Sys_ReturnResonID"))="0" Then response.Write(" selected")%>>收受</option>
										<option Value="1"<%If trim(request("Sys_ReturnResonID"))="1" Then response.Write(" selected")%>>公示</option>
									</select>
									<select name="Sys_AcceptReson" onchange="if(eval(myForm.Sys_ReturnResonID.value)==0){myForm.Sys_Note.value=this.value;}else{myForm.Sys_Note.value='';}" class="btn1">
										<option Value="本人">本人</option>
										<option Value="代收">代收</option>
									</select>
								</td>
							</tr>
							<tr>
								<td>郵寄日期</td>
								<td nowrap>
									<input name="MailDate" class="btn1" type="text" value="<%=trim(request("MailDate"))%>" size="10" maxlength="10" onkeyup="value=value.replace(/[^\d]/g,'')">
									<input type="button" name="datestr" value="..." onclick="OpenWindow('MailDate');">
								</td>
								<td>
									原因
								</td>
								<td colspan=2>
									<input name="Sys_Note" class="btn1" type="text" value="<%=trim(request("Sys_Note"))%>" size="30" maxlength="50">
								</td>
							
							</tr>
							<tr>
								<td colspan=4 align="center">
									<input type="button" name="btnAdd" value="新增" onclick="funAdd();">
									<input name="btnexit" type="button" value=" 關 閉 " onclick="funExit();">
									<input name="btnexit" type="button" value=" 產生送達證書 " onclick="funBillDeliver();">
									<%
									If not (Sys_City = "嘉義縣") Then

										Response.Write "<input name=""btnexit"" type=""button"" value="" 產生舉發單 "" onclick=""funPrintLegal();"">"
									End if 

									If Sys_City = "台中市" Then

										Response.Write "<input name=""btnexit"" type=""button"" value="" 產生存根聯 "" onclick=""funPrintA4();"">"
									End if 
									%>									
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" class="style3">送達紀錄列表</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th>送達日期</th>
					<th>送達人員</th>
					<th>郵寄日期</th>
					<th>大宗掛號碼</th>
					<th>送達類別</th>
					<th>送達狀況</th>
					<th>退件原因</th>
					<th>操作</th>
				</tr><%
				str_ArriveType=split("裁決,催告,郵寄",",")
				Sys_ReturnResonID=split("收受,公示,寄存",",")
				while Not rsload.eof
					response.write "<tr align='center' bgcolor='#FFFFFF'"
					lightbarstyle 0
					response.write ">"
					if trim(request("Edit_SN"))<>trim(rsload("SN")) then
						response.write "<td>"
						If not ifnull(rsload("Imagefilename")) Then						
							Response.Write "<a href=""./Picture/"&trim(rsload("Imagefilename"))&""" target=""_blank"">"
						end if

						Response.Write gInitDT(rsload("ArrivedDate"))

						If not ifnull(rsload("Imagefilename")) Then	Response.Write "</a>"

						Response.Write "</td>"

						response.write "<td>"&rsload("Chname")&"</td>"
						response.write "<td>"&gInitDT(rsload("MailDate"))&"</td>"
						response.write "<td>"&rsload("SendMailStation")&"</td>"

						response.write "<td>"
						If Not IfNull(rsload("ArriveType")) Then response.Write(str_ArriveType(rsload("ArriveType")))
						response.Write "</td>"

						response.write "<td>"
						If Not IfNull(rsload("ReturnResonID")) Then response.Write(Sys_ReturnResonID(rsload("ReturnResonID")))
						response.Write "</td>"
						response.write "<td>"&rsload("Note")&"</td>"

						response.write "<td>"
						response.write "<input type=""button"" name=""Update"" value=""修改"" onclick=""funEdit('"&rsload("SN")&"');"">"
						response.write "<input type=""button"" name=""Del"" value=""刪除"" onclick=""funDel('"&rsload("SN")&"');"">"
						response.write "<input type=""button"" name=""btnMap"" value=""送達證書影像檔上傳"" onclick=""funMap('"&rsload("SN")&"');"">"
						response.write "</td>"
					else
						response.write "<tr align='center' bgcolor='#FFFFFF'"
						lightbarstyle 0
						response.write ">"
						response.write "<td>"
						response.write "<input name=""edit_ArrivedDate"" class=""btn1"" type=""text"" value="""&gInitDT(rsload("ArrivedDate"))&""" size=""10"" maxlength=""10"" onkeyup=""value=value.replace(/[^\d]/g,'')"">"
						response.write "<input type=""button"" name=""datestr"" value=""..."" onclick=""OpenWindow('edit_ArrivedDate');"">"
						response.write "</td>"

						response.write "<td>"
						response.write "<select name=""edit_SenderMemID"" class=""btn1"">"
						response.write "<option Value="""">請選擇</option>"
						strSQL="select MemberID,Chname from MemberData"
						set rs1=conn.execute(strSQL)
						while Not rs1.eof
							response.write "<option value="""&rs1("MemberID")&""""
							if trim(rs1("MemberID"))=trim(rsload("SenderMemID")) then response.write " Selected"
							response.write ">"&rs1("Chname")&"</option>"
							rs1.movenext
						wend
						rs1.close
						response.write "</select>"
						response.write "</td>"

						response.write "<td>"
						response.write "<input name=""edit_MailDate"" class=""btn1"" type=""text"" value="""&gInitDT(rsload("MailDate"))&""" size=""10"" maxlength=""10"" onkeyup=""value=value.replace(/[^\d]/g,'')"">"
						response.write "<input type=""button"" name=""datestr"" value=""..."" onclick=""OpenWindow('edit_MailDate');"">"
						response.write "</td>"

						response.write "<td>"
						response.write "<input name=""edit_SendMailStation"" class=""btn1"" type=""text"" value="""&rsload("SendMailStation")&""" size=""10"" maxlength=""50"">"
						response.write "</td>"

						response.write "<td>"
						response.write "<select name=""edit_ArriveType"" class=""btn1"">"
						response.Write "<option Value=""0"""
						If trim(rsload("ArriveType"))="0" Then response.Write(" selected")
						response.Write ">裁決</option>"

						response.Write "<option Value=""1"""
						If trim(rsload("ArriveType"))="1" Then response.Write(" selected")
						response.Write ">催告</option>"
						
						Response.Write "<option Value=""2"""
						If trim(rsload("ArriveType"))="2" Then response.Write(" selected")
						response.Write ">郵寄</option>"

						response.write "</select>"
						response.write "</td>"

						response.write "<td>"
						response.write "<select name=""edit_ReturnResonID"" class=""btn1"">"
						response.Write "<option Value="""""
						If trim(rsload("ReturnResonID"))="" Then response.Write(" selected")
						response.Write ">請選擇</option>"

						response.Write "<option Value=""0"""
						If trim(rsload("ReturnResonID"))="0" Then response.Write(" selected")
						response.Write ">收受</option>"

						response.Write "<option Value=""1"""
						If trim(rsload("ReturnResonID"))="1" Then response.Write(" selected")
						response.Write ">公示</option>"

						response.Write "<option Value=""2"""
						If trim(rsload("ReturnResonID"))="2" Then response.Write(" selected")
						response.Write ">寄存</option>"

						response.write "</select>"
						response.write "</td>"

						response.write "<td>"
						response.write "<input name=""edit_Note"" class=""btn1"" type=""text"" value="""&rsload("Note")&""" size=""10"" maxlength=""50"">"
						response.write "</td>"

						response.write "<td>"
						response.write "<input type=""button"" name=""Update"" value=""確定"" onclick=""funUpdate('"&rsload("SN")&"');"">"
						response.write "<input type=""button"" name=""Canal"" value=""取消"" onclick=""funEdit('');"">"
						response.write "</td>"
					end if
					response.write "</tr>"
					rsload.movenext
				wend%>
			</table>
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_State" value="">
<input type="Hidden" name="BillSN" value="<%=request("BillSN")%>">
<input type="Hidden" name="Edit_SN" value="">
<input type="Hidden" name="SN" value="">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">

var Sys_City="<%=sys_City%>";

<%response.write "UnitMan('UnitID','SenderMemID','"&request("SenderMemID")&"');"%>
if(eval(myForm.Sys_ReturnResonID.value)==0){myForm.Sys_Note.value=myForm.Sys_AcceptReson.value;}else{myForm.Sys_Note.value='';}
function KeyDown(){ 
	if (event.keyCode==116){	//F5鎖死
		event.keyCode=0;   
		event.returnValue=false;   
	}
}
function funAdd(){
	var err=0;
	var sys_illegaldate='<%=sys_illegaldate%>';
	if(myForm.ArrivedDate.value==""){
		err=1;
		alert("送達日必須輸入!!");
	}else if(myForm.ArrivedDate.value!=""){
		if(!dateCheck(myForm.ArrivedDate.value)){
			err=1;
			alert("送達日輸入不正確!!");
		}
	}
	if(err==0){
		if(myForm.Sys_ArriveType.value==""){
			err=1;
			alert("送達類別必須選取!!");
		}
	}

	if(err==0){
		if(eval(myForm.ArrivedDate.value)<=eval(sys_illegaldate)){
			err=1;
			alert("送達日必須大於違規日!!");
		}
	}

	if(err==0){
		myForm.DB_State.value='Add';
		myForm.submit();
	}
}
function AcceptReson(){

	if(eval(myForm.Sys_ReturnResonID.value)==0){

		myForm.Sys_Note.value=myForm.Sys_AcceptReson.value;
	}else{
		
		myForm.Sys_Note.value='';
	}
}
function funEdit(SN){
	myForm.Edit_SN.value=SN;
	myForm.submit();
}
function funDel(SN){
	if(confirm('確定刪除此筆紀錄嗎？')){
		myForm.SN.value=SN;
		myForm.DB_State.value='Del';
		myForm.submit();
	}
}
function funMap(SN){

	var UrlStr="SendStyle.asp?SN="+SN;
	newWin(UrlStr,"winMap",700,150,50,10,"yes","yes","no","no");
}

function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}


function funPrintLegal(){	

	var UrlStr="";	
	
	if(Sys_City=='花蓮縣'){
		UrlStr="../PasserQuery/BillPrintLegal_YiLan_chromat_1110817.asp";
		//UrlStr="../PasserQuery/BillPrintLegal_CHCG_1110817.asp";

	}else if(Sys_City=='嘉義縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_Chiayi_1110817.asp";

	}else if(Sys_City=='高雄市'){
		
		UrlStr="../PasserQuery/BillPrintLegal_KaoHsiungCity_1110817.asp";
	}else if(Sys_City=='基隆市'){
		
		UrlStr="../PasserQuery/BillPrintLegal_KeeLung_1110817.asp";
	}else if(Sys_City=='金門縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_KMA_1110817.asp";
	}else if(Sys_City=='苗栗縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_miaoli_1110817.asp";
	}else if(Sys_City=='南投縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_NanTou_1110817.asp";
	}else if(Sys_City=='屏東縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_PingTung_1110817.asp";
	}else if(Sys_City=='台南市'){
		
		UrlStr="../PasserQuery/BillPrintLegal_TaiNanCity_1110817.asp";
	}else if(Sys_City=='宜蘭縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_YiLan_chromat_1110817.asp";
	}else if(Sys_City=='雲林縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_Yunlin_1110817.asp";
	}else if(Sys_City=='彰化縣'){
		
		UrlStr="../PasserQuery/BillPrintLegal_CHCG_1110817.asp";
	}else if(Sys_City=='臺東縣'){
		
		UrlStr="../PasserQuery/BillPrintsTaiTung_chromat_1110817.asp";
	}else if(Sys_City=='台中市'){
		
		UrlStr="../PasserQuery/BillPrints_TaiChungCity_1110817.asp";
	}else if(Sys_City=='連江縣'){
		
		UrlStr="../PasserQuery/BillPrints_lattice_MU.asp";
	}else if(Sys_City=='澎湖縣'){
		
		UrlStr="../PasserQuery/BillPrints_a4_penghu1120118.asp";
	}

	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funPrintA4(){	

	var UrlStr="";	
	
	if(Sys_City=='台中市'){
		
		UrlStr="../PasserQuery/BillPrintsA4_TaiChungCity_1120419.asp";
	}

	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}


function funBillDeliver(){

	var UrlStr="BillBase_Deliver_Word.asp";

	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function funUpdate(SN){
	var err=0;
	if(myForm.edit_ArrivedDate.value==""){
		err=1;
		alert("送達日必須輸入!!");
	}else if(myForm.edit_ArrivedDate.value!=""){
		if(!dateCheck(myForm.edit_ArrivedDate.value)){
			err=1;
			alert("送達日輸入不正確!!");
		}
	}
	if(err==0){
		myForm.SN.value=SN;
		myForm.DB_State.value='Update';
		myForm.submit();
	}
}
function funExit(){
	opener.myForm.submit(); 
	self.close();
}
</script>
<%
conn.close
set conn=nothing
%>