<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
SQL="select /*+ INDEX(md IDX_MEMBERDATA1) INDEX(gb IDX_GETBILLBASE2) */ " & _
    "ut.UNITNAME,md.LoginID,md.AccountStateID,gb.GetBillDate,gb.BillReturnDate,gb.BillStartNumber,gb.BILLENDNUMBER, " & _
    "gb.CounterfoiReturn,gb.note,ut.unitid, " & _
    "gb.GetBillMemberID,gb.DispatchMemberID " & _
    "from getbillbase gb,memberdata md,unitinfo ut " & _
    "where gb.GetBillMemberID=md.MemberID and gb.RecordStateID <> -1 " & _
    "and md.UnitID=ut.UnitID and gb.GETBILLSN=" & Request("GETBILLSN")
set RsUpd1=Server.CreateObject("ADODB.RecordSet")
RsUpd1.open SQL,Conn,3,3

sql = "Select ChName , MemberID from MemberData where memberid='" & RsUpd1("DispatchMemberID") & "'"

set RsUpd3=Server.CreateObject("ADODB.RecordSet")
RsUpd3.open sql,Conn,3,3

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>領單管理-資料異動</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function datacheck()
{
	var result ;
	var rtnChkBillNum ;
	
  if(document.all.GetBillDate.value=="")   
  {
    alert('請選擇領單日期!!');
    return false;  
  }	
  
  if(document.all.UnitID.value=="")   
  {
    alert('請選擇領單單位!!');
    return false;  
  }
  
  if(document.all.GetBillMemberID.value=="")   
  {
    alert('請選擇領單人員!!');
    return false;  
  }  
  
  if(document.all.BillStartNumber.value=="")   
  {
    alert('請輸入舉發單起始碼!!');
    return false;  
  }else{
     result = chkBillNumber(document.all.BillStartNumber,"[舉發單起始碼] 格式錯誤!!"); 
     if (result != "Y") return false;
  } 
   
  if(document.all.BillEndNumber.value=="")   
  {
    alert('請輸入舉發單截止碼!!');
    return false;  
  }else{
     result = chkBillNumber(document.all.BillEndNumber,"[舉發單截止碼] 格式錯誤!!"); 
     if (result != "Y") return false;      	
  } 
   
	if(document.all.ReturnType[0].checked){
		document.all.CounterfoiReturn.value=1;

	}else if (document.all.ReturnType[1].checked){
		document.all.CounterfoiReturn.value=0;

	}else{
		alert('請選擇是否仍在使用中!!');
		return false;  	
	} 
  
  rtnChkBillNum = ValidateBillNumbers (document.all.BillStartNumber,document.all.BillEndNumber,'Y');
  switch (rtnChkBillNum){
     case 1:
        alert("[舉發單起始碼]與[舉發單截止碼]之前三碼不一致!!");
        return false;
        break;
     case 2:	
        alert("[舉發單起始碼]不得大於[舉發單截止碼]!!");
        return false; 
        break;    
  }
}
-->
</Script>

<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {font-size: 14px}
.style2 {font-size: 18px}
.style3 {font-size: 15px}
-->
</style></head>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<script language=javascript src='../js/GetBill.js'></script>
<body>
<FORM NAME="updGetBillBase" ACTION="GetBill_mdy.asp" METHOD="POST" onSubmit="return datacheck();">  
	<input type="hidden" name="BillCount">	
	<input type="hidden" name="CounterfoiReturn">
	<input type="hidden" name="chekMemID">
	<input type="hidden" name="tag" value="<%=request("tag")%>"> 
	<input type="hidden" name="GETBILLSN" value="<%=request("GETBILLSN")%>">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">領單管理-資料異動</span></td>
  </tr>
  
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="11%" bgcolor="#FFFFCC"><div align="right"><span class="font12">領單日期           </span></div></td>
        <td width="89%">
        	<input type='text' size='7' maxlength='7' id='GetBillDate' name='GetBillDate' value="<%=gInitDT(RsUpd1("getbilldate"))%>" class="btn1">
        		<input type="button" name="datestra" value="..." onclick="OpenWindow('GetBillDate');">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="font12">發放人員</span></div></td>
        <td><span class="font12"><%=RsUpd3("ChName")%></span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" ><span class="font12">領單單位          </span></div></td>
        <td>
			<input name="LevelUnit" type="radio" onClick="funCounterReceive();" value="0" <%
				if ifnull(Request("LevelUnit")) and trim(RsUpd1("AccountStateID"))="-1" then
					response.write "checked"
				elseif Request("LevelUnit")="0" then
					 response.write "checked"
				end if%>>
          </span><span class="font12">已離職<span class="font10">
		  &nbsp;&nbsp;

		  <input name="LevelUnit" type="radio" onClick="funCounterReceive();" value="1" <%
				if ifnull(Request("LevelUnit")) and trim(RsUpd1("AccountStateID"))="0" then
					response.write "checked"
				elseif Request("LevelUnit")="1" then
					 response.write "checked"
				end if%>>
          </span><span class="font12">現任中<span class="font10">
		  &nbsp;&nbsp;
			<%
				if ifnull(Request("LevelUnit")) and trim(RsUpd1("AccountStateID"))="-1" then
					strtmp="<select name=""UnitID"" ID=""UnitID"" class=""btn1"" onchange=""UnitLaverMan('UnitID','GetBillMemberID');"">"
				elseif Request("LevelUnit")="0" then
					strtmp="<select name=""UnitID"" ID=""UnitID"" class=""btn1"" onchange=""UnitLaverMan('UnitID','GetBillMemberID');"">"
				else
					strtmp="<select name=""UnitID"" ID=""UnitID"" class=""btn1"" onchange=""UnitMan('UnitID','GetBillMemberID');"">"
				end if
				strSQL="select UnitID,UnitName from UnitInfo order by UnitTypeID,UnitName"
				strtmp=strtmp+"<option value="""">所有單位</option>"
				set rs1=conn.execute(strSQL)
				while Not rs1.eof
					strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
					if trim(rs1("UnitID"))=trim(RsUpd1("UnitID")) then
						strtmp=strtmp+" selected"
					end if
					strtmp=strtmp+">"&rs1("UnitID")&" - "&rs1("UnitName")&"</option>"
					rs1.movenext
				wend
				rs1.close
				strtmp=strtmp+"</select>"
				response.write strtmp
			%>
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" ><span class="font12">領單人員          </span></div></td>
        <td>
			<%
			if ifnull(Request("LevelUnit")) and trim(RsUpd1("AccountStateID"))="-1" then
				response.write UnLaverSelectMemberOption("UnitID","GetBillMemberID")
			elseif Request("LevelUnit")="0" then
				response.write UnLaverSelectMemberOption("UnitID","GetBillMemberID")
			else
				response.write UnSelectMemberOption("UnitID","GetBillMemberID")
			end if
			%>
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"> <div align="right" ><span class="font12">舉發單號          </span></div></td>
        <td>
        	<!--<input name="BillStartNumber" type="text" size="10" maxlength="9" onKeyDown='lockString(this);' onKeyUp='lockString(this);'>-->
        	<input name="BillStartNumber" readOnly type="text" value="<%=RsUpd1("billstartnumber")%>" size="10" maxlength="9" onBlur="javascript:this.innerText=this.value.toUpperCase();" class="btn1">
          ~
          <input name="BillEndNumber" readOnly type="text" value="<%=RsUpd1("billendnumber")%>" size="10" maxlength="9" onBlur="javascript:this.innerText=this.value.toUpperCase();" class="btn1">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="font12">使用狀態 </span></div><br></td>
        <td><span class="font12">
		  <input name="ReturnType" type="radio" <%
			If not ifnull(request("counterfoireturn")) Then
				if trim(request("counterfoireturn"))="1" then response.write "checked"
			else
				if trim(RsUpd1("counterfoireturn"))="1" then response.write "checked"
			end if%> onclick="funchkBillNo();">
          </span><span class="font12">使用完畢 <span class="font10"><font color="gray">(員警繳回)</font></span>
          &nbsp;&nbsp;
          <input name="ReturnType" type="radio" <%
			If not ifnull(request("counterfoireturn")) Then
				if trim(request("counterfoireturn"))="0" then response.write "checked"
			else
				if trim(RsUpd1("counterfoireturn"))="0" then response.write "checked"
			end if%>>
  				<span class="font12">  使用中 </span><br>
  				<!--smith 0611加上繳回日期-->
  				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;繳回日期
  				 <input type='text' size='7' maxlength='7' id='BillReturnDate' name='BillReturnDate' value="<%=gInitDT(RsUpd1("BillReturnDate"))%>" class="btn1">
        		</td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="font12">備註</span></div></td>
      		<td>
          <textarea name="Note" cols="50" rows="3" onKeyDown="calStr(this,50);" onKeyUp="calStr(this,50);"><%=RsUpd1("note")%></textarea>
          <span class="smallBlock">剩餘字數: <input readOnly size=3 name="nbchars" class="smallBlock"></span>
          </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
        <input type="submit" name="Submit423" value="確 定">
        <span class="style3"><img src="space.gif" width="9" height="8"></span>       
        <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉">
</p>    </td>
  </tr>

</table>
</FORM>
</body>
<script type="text/javascript" src="../js/date.js"></script>
<SCRIPT LANGUAGE="JavaScript" >
<%
if ifnull(Request("LevelUnit")) and trim(RsUpd1("AccountStateID"))="-1" then
	response.write "UnitLaverMan('UnitID','GetBillMemberID','"&CStr(RsUpd1("GetBillMemberID"))&"');"
	'Response.write "document.all.chekLaverChMemID.value='"&CStr(RsUpd1("LoginID"))&"';"

	response.write "ListItemLaver('UnitID','GetBillMemberID');"
elseif Request("LevelUnit")="0" then
	response.write "UnitLaverMan('UnitID','GetBillMemberID','"&CStr(RsUpd1("GetBillMemberID"))&"');"
	'Response.write "document.all.chekLaverChMemID.value='"&CStr(RsUpd1("LoginID"))&"';"

	response.write "ListItemLaver('UnitID','GetBillMemberID');"
else
	response.write "UnitMan('UnitID','GetBillMemberID','"&CStr(RsUpd1("GetBillMemberID"))&"');"
	'Response.write "document.all.chekChMemID.value='"&CStr(RsUpd1("LoginID"))&"';"

	response.write "ListItem('UnitID','GetBillMemberID');"
end if
%>

function funCounterReceive(){
	updGetBillBase.onSubmit="";
	updGetBillBase.action="";
	updGetBillBase.target="";
	updGetBillBase.submit();
}
function funchkBillNo(){
	if(updGetBillBase.BillStartNumber.value!=''&&updGetBillBase.BillEndNumber.value!=''){
		updGetBillBase.Submit423.disabled=true;
		runServerScript("chkBillBase.asp?BillStartNumber="+updGetBillBase.BillStartNumber.value+"&BillEndNumber="+updGetBillBase.BillEndNumber.value);
	}
}
</script> 
</html>
<!-- #include file="../Common/ClearObject.asp" -->
