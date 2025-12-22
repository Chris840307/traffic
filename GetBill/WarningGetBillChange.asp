<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
SQL="select /*+ INDEX(md IDX_MEMBERDATA1) INDEX(gb IDX_GETBILLBASE2) */ " & _
    "ut.UNITNAME,md.ChName,gb.GetBillDate,gb.BillStartNumber,gb.BILLENDNUMBER, " & _
    "gb.CounterfoiReturn,gb.note,ut.unitid, " & _
    "gb.GetBillMemberID,gb.DispatchMemberID " & _
    "from Warninggetbillbase gb,memberdata md,unitinfo ut " & _
    "where gb.GetBillMemberID=md.MemberID and gb.RecordStateID != -1 " & _
    "and md.UnitID=ut.UnitID and gb.GETBILLSN=" & Request("GETBILLSN")
set RsUpd1=Server.CreateObject("ADODB.RecordSet")
RsUpd1.open SQL,Conn,3,3

sqlUnit = "Select UnitName , UnitID from UnitInfo"
set RsUnit=Server.CreateObject("ADODB.RecordSet")
RsUnit.open sqlUnit,Conn,3,3

sql = "Select ChName , MemberID from MemberData where UnitID='" & RsUpd1("UnitID") & "'"
set RsUpd2=Server.CreateObject("ADODB.RecordSet")
RsUpd2.open sql,Conn,3,3

sql = "Select ChName , MemberID from MemberData where memberid='" & RsUpd1("DispatchMemberID") & "'"
set RsUpd3=Server.CreateObject("ADODB.RecordSet")
RsUpd3.open sql,Conn,3,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>警告單管理-資料異動</title>
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
  }/*else if(isNaN(document.all.BillStartNumber.value)){
	 alert('[舉發單起始碼] 格式錯誤!!');
     if (result != "Y") return false;
  } */
   
  if(document.all.BillEndNumber.value=="")   
  {
    alert('請輸入舉發單截止碼!!');
    return false;  
  }/*else if(isNaN(document.all.BillEndNumber.value)){
	 alert('[舉發單截止碼] 格式錯誤!!');
     if (result != "Y") return false;
  } */
   
  /*if(document.all.ReturnType[0].checked){
     document.all.CounterfoiReturn.value=1;
  }else if (document.all.ReturnType[1].checked){
  	 document.all.CounterfoiReturn.value=0;
  }else{
    alert('請選擇是否仍在使用中!!');
    return false;  	
  } */
  
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
<FORM NAME="WarningupdGetBillBase" ACTION="WarningGetBill_mdy.asp" METHOD="POST" onSubmit="return datacheck();">  
	<input type="hidden" name="BillCount">	
	<input type="hidden" name="CounterfoiReturn" value="<%=RsUpd1("counterfoireturn")%>">
	<input type="hidden" name="tag" value="<%=request("tag")%>"> 
	<input type="hidden" name="GETBILLSN" value="<%=request("GETBILLSN")%>">
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">警告單管理-資料異動</span></td>
  </tr>
  
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="11%" bgcolor="#FFFFCC"><div align="right"><span class="font12">領單日期           </span></div></td>
        <td width="89%">
        	<input type='text' size='10' id='GetBillDate' name='GetBillDate' value="<%=gInitDT(RsUpd1("getbilldate"))%>" class="btn1">
        		<input type="button" name="datestra" value="..." onclick="OpenWindow('GetBillDate');">
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="font12">發放人員</span></div></td>
        <td><span class="font12"><%=RsUpd3("ChName")%></span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" ><span class="font12">接續使用單位          </span></div></td>
        <td>
			<%=UnSelectUnitOption("UnitID","GetBillMemberID")%>
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" ><span class="font12">接續使用員警          </span></div></td>
        <td>
			<%=UnSelectMemberOption("UnitID","GetBillMemberID")%>
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"> <div align="right" ><span class="font12">舉發單號          </span></div></td>
        <td>
        	<!--<input name="BillStartNumber" type="text" size="10" maxlength="9" onKeyDown='lockString(this);' onKeyUp='lockString(this);'>-->
        	<input name="BillStartNumber" type="text" value="<%=RsUpd1("billstartnumber")%>" size="10" maxlength="11" onBlur="javascript:this.innerText=this.value.toUpperCase();" class="btn1">
          ~
          <input name="BillEndNumber" type="text" value="<%=RsUpd1("billendnumber")%>" size="10" maxlength="11" onBlur="javascript:this.innerText=this.value.toUpperCase();" class="btn1" readOnly>
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
        <span class="style3"><img src="space.gif" width="9" height="8"></span>        <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉">
</p>    </td>
  </tr>

</table>
</FORM>
</body>
<SCRIPT LANGUAGE="JavaScript" >
<%response.write "UnitMan('UnitID','GetBillMemberID','"&CStr(RsUpd1("GetBillMemberID"))&"');"%>
</script> 
</html>
<!-- #include file="../Common/ClearObject.asp" -->
