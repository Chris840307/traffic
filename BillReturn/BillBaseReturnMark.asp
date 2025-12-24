<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/banner.asp"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<TITLE> 退件管理 </TITLE>
<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

if trim(request("DB_Selt"))="Selt" then
	Sys_BillNo=Split(Ucase(request("item"))&" ",",")
	Sys_BackCause=Split(request("Sys_BackCause")&" ",",")

	for i=0 to Ubound(Sys_BillNo)
		if trim(Sys_BillNo(i))<>"" then
			strSQL="Update BillMailHistory set ReturnResonID='"&Sys_BackCause(i)&"',MailReturnDate="&funGetDate(gOutDT(request("Sys_BackDate")),0)&",ReturnRecordMemberID="&Session("User_ID")&",ReturnReCordDate="&funGetDate((now),1)&",UserMarkMemberID="&Session("User_ID")&",UserMarkDate="&funGetDate((now),1)&",UserMarkResonID='"&Sys_BackCause(i)&"',UserMarkReturnDate="&funGetDate(gOutDT(request("Sys_BackDate")),0)&" where BillNo='"&Sys_BillNo(i)&"'"

			conn.execute(strSQL)

			strSQL="Update Billbase set BillStatus=3 where BillNo='"&Sys_BillNo(i)&"'"
			conn.execute(strSQL)
		end if
	next
	Response.write "<script>"
	Response.Write "alert('儲存完成！');"
	Response.write "</script>"
end if
%>
</HEAD>
<BODY>
<form name=myForm method="post">
<table width="100%" border="0" bgcolor="#ffffff">
	<tr>
		<td height="27" bgcolor="#FFCC33" class="pagetitle">退件管理</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#CCCCCC">
			<table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
					<td nowrap>
						退件原因&nbsp;
						<select name="Sys_BackCauseMain" class="btn1">
							<%strSQL="select ID,Content from DCICode where TypeID=7"
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
						<input type="button" name="btnDefu" value="預設為退件原因" onclick="funDefuSelt();">&nbsp;&nbsp;&nbsp;
						<br>
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
						退件日期&nbsp;<input name="Sys_BackDate" type="text" class="btn1" size="10" maxlength="11" onkeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('Sys_BackDate');">
						&nbsp;&nbsp;<input type="button" name="btnOK" value="確定" onclick="funSelt();">
						<img src="space.gif" width="9" height="8">
						<%
							Response.Write "<input type=""button"" name=""insert2"" value=""再多30筆"" onClick=""insertRow(fmyTable)"">"
						%>
						<!--<input type="button" name="Delete" value="減少1筆" onClick="DeleteRow(fmyTable)">-->
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td height="26" bgcolor="#FFCC33">退件紀錄列表</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0" valign="top">
			<table id='fmyTable' width='978' border='0' bgcolor='#FFFFFF'>
				<tr bgcolor="#ffffff">
					<td align='center' bgcolor="#ffffff" nowrap>目前無新增項目</td>
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
</form>

<form name="upForm" method="post">

	<input type="Hidden" name="item" value="">
	<input type="Hidden" name="Sys_BackCause" value="">

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
		
		txtArea.innerHTML ="單號<input type=text name='item' size=10 class='btn1'>&nbsp;&nbsp;原因<select name='Sys_BackCause' class='btn1'><%=seltarr%></select>";
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
function funSelt(){
	var error=0;

	var item='';
	var Sys_BackCause='';

	if(myForm.Sys_BackDate.value==""){
		error=1;
		alert("退件日必須要填!!");
	}else if(!dateCheck(myForm.Sys_BackDate.value)){
		error=1;
		alert("退件日輸入不正確!!");
	}else{

		for(i=0;i<myForm.item.length;i++){
			if(myForm.item[i].value!=''){
				if(item!=''){
					item=item+',';
					Sys_BackCause=Sys_BackCause+',';
				}
				item=item + myForm.item[i].value;
				Sys_BackCause=Sys_BackCause + myForm.Sys_BackCause[i].value;
			}
		}

		upForm.item.value=item;
		upForm.Sys_BackCause.value=Sys_BackCause;

		upForm.DB_Selt.value="Selt";
		upForm.submit();
	}
}

<%
	Response.Write "for(j=0;j<=3;j++){insertRow(fmyTable);}"
%>
</script>