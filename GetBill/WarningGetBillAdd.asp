<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<%
sqlUnit = "Select UnitName , UnitID from UnitInfo"
set RsUnit=Server.CreateObject("ADODB.RecordSet")
RsUnit.open sqlUnit,Conn,3,3

sql = "Select ChName , MemberID from MemberData where UnitID='" & Request("UnitID") & "'"
set RsUpd1=Server.CreateObject("ADODB.RecordSet")
RsUpd1.open sql,Conn,3,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>領單管理-資料新增</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<script language=javascript src='../js/WarningGetBill.js'></script>
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

  if(document.all.GetBillMemberID.value=="")   
  {
    alert('請選擇領單人員!!');
    return false;  
  } 
  
  if(document.all.UnitID.value=="")   
  {
    alert('請選擇領單單位!!');
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
    return false;  
  }
  if(document.all.ReturnType[0].checked){
     document.all.CounterfoiReturn.value=0;
	 document.all.BillIn.value=0;
  }else if (document.all.ReturnType[1].checked){
  	 document.all.CounterfoiReturn.value=0;
	 document.all.BillIn.value=1;
  }else if (document.all.ReturnType[2].checked){
  	 document.all.CounterfoiReturn.value=0;
	 document.all.BillIn.value=2;
  }

  if (isNaN(document.all.BillEndNumber.value)){
	  rtnChkBillNum = ValidateBillNumbers (document.all.BillStartNumber,document.all.BillEndNumber,'Y');
	  switch (rtnChkBillNum){
		 /*case 1:
			alert("[舉發單起始碼]與[舉發單截止碼]之前三碼不一致!!");
			return false;
			break;*/
		 case 2:	
			alert("[舉發單起始碼]不得大於[舉發單截止碼]!!");
			return false; 
			break;    
	  }
  }else{
	document.all.BillCount.value =document.all.BillEndNumber.value-1;
	return true;
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
.style5 {
	font-size: 11px;
	color: #666666;
}
-->
</style></head>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<body>
<%
if Session("Msg")<>"" then
	 Response.write "<font  color='Red' size='2'>" & Session("Msg") & "</font>"
	 Session("Msg") = ""
end if	
%>	
<FORM NAME="addGetBillBase" ACTION="WarningGetBill_mdy.asp" METHOD="POST" onSubmit="return datacheck();" onkeydown="funTextControl();">  	
	<input type="hidden" name="BillCount">
	<input type="hidden" name="CounterfoiReturn">
	<input type="hidden" name="BillIn">
	<input type="hidden" name="tag" value="NEW"> 
<table width="100%" height="70%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">警告單管理-資料異動</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td width="11%" bgcolor="#FFFFCC"><div align="right" ><span class="font12">領單日期           </span></div></td>
        <td width="89%">																													<!--<%=request("GetBillDate")%> -->
        	<input type='text' size='10' id='GetBillDate' name='GetBillDate' value='<%=gInitDT(now)%>'  class="btn1">
        	 <!-- <input type="button" name="datestra" value="..." onclick="OpenWindow('GetBillDate');"> -->
        	<font size="2">* 可以使用Enter 鍵直接跳到領單人員</font>
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" ><span class="font12">發放人員</span></div></td>
        <td><span class="font12"><%=Session("Ch_Name")%></span></td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" ><span class="font12">領單單位          </span></div></td>
        <td>
			<input name="LevelUnit" type="radio" onClick="funCounterReceive();" value="0" <%
				If not ifnull(request("LevelUnit")) Then
					if trim(request("LevelUnit"))="0" then response.write "checked"
				end if%>>
			</span><span class="font12">已離職<span class="font10">
		  &nbsp;&nbsp;
			<input name="LevelUnit" type="radio" onClick="funCounterReceive();" value="1" <%
				if trim(request("LevelUnit"))<>"0" then response.write "checked"
				%>>
			</span><span class="font12">現任中<span class="font10">
		  &nbsp;&nbsp;
			<%
				If trim(request("LevelUnit"))="0" Then
					strtmp="<select name=""UnitID"" ID=""UnitID"" class=""btn1"" onchange=""UnitLaverMan('UnitID','GetBillMemberID');"">"
				else
					strtmp="<select name=""UnitID"" ID=""UnitID"" class=""btn1"" onchange=""UnitMan('UnitID','GetBillMemberID');"">"
				end if
				strSQL="select UnitID,UnitName from UnitInfo order by UnitTypeID,UnitName"
				strtmp=strtmp+"<option value="""">所有單位</option>"
				set rs1=conn.execute(strSQL)
				while Not rs1.eof
					strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
					if trim(rs1("UnitID"))=trim(request(UnitName)) then
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
        <td bgcolor="#FFFFCC"><div align="right"><span class="font12">領單人員          </span></div></td>
        <td><%
			If trim(request("LevelUnit"))="0" Then
				response.write UnLaverSelectMemberOption("UnitID","GetBillMemberID")
			else
				response.write UnSelectMemberOption("UnitID","GetBillMemberID")
			end if
			%>
			<span class="style5">*可輸入人員代碼帶出人員與所屬單位</span>
		</td>
      </tr>

      <tr>
        <td bgcolor="#FFFFCC"> <div align="right" ><span class="font12">舉發單號          </span></div></td>
        <td>
        	<!--<input name="BillStartNumber" type="text" size="10" maxlength="9" onKeyDown='lockString(this);' onKeyUp='lockString(this);'>-->
        	起始號
        	<input name="BillStartNumber" value='<%=request("BillStartNumber")%>' type="text" size="10" maxlength="11" onBlur="javascript:this.innerText=this.value.toUpperCase();" class="btn1">

			～<input name="BillEndNumber" value='<%=request("BillEndNumber")%>' type="text" size="10" maxlength="11"class="btn1">截止號
			<br>
			<font size="4" color="red"><B>*新增標示單時請連同100A或(NO123456)一併填寫</B></font>
        </td>
      </tr>
	
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="font12">使用狀況 </span></div></td>
        <td>
	
       <!--      <input name="ReturnType" type="hidden <% 'if (request("counterfoireturn")="1") then response.write "checked" end if%>> -->
     <!--      </span></span><span class="style1"><span class="style3"> 使用完畢-->
          <input name="ReturnType" type="radio" <%if (request("counterfoireturn")="0" or request("counterfoireturn")="") then response.write "checked" end if%>>
   <span class="font12">員警領取使用</span>  <input name="ReturnType" type="radio">
   <span class="font12">入庫 </span>  <input name="ReturnType" type="radio">
   <span class="font12">出庫 </span> <span class="style4"> ( 領取舉發單數量多會需要較長處理時間 . 請耐心等候 )</span> </td>
      </tr>

      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="font12">備註</span></div></td>
        <td><span class="style1"><span class="style3">
          <textarea name="Note" cols="50" rows="3" onKeyDown="calStr(this,50);" onKeyUp="calStr(this,50);"><%=request("note")%></textarea>
          <span class="smallBlock">剩餘字數: <input readOnly size=3 name="nbchars" class="smallBlock"></span>
          </span></span></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
         <input type="Submit" name="Submit423" value="確 定">
        <span class="style3"><img src="space.gif" width="9" height="8"></span>        <input type="button" name="Submit4232" onClick="javascript:window.opener.location.reload();window.close();" value="關 閉">
</p>    </td>
  </tr>

</table>
</FORM>
</body>
<script type="text/javascript" src="../js/date.js"></script>
<SCRIPT LANGUAGE="JavaScript" >
<%response.write "UnitMan('UnitID','GetBillMemberID','"&request("GetBillMemberID")&"');"%>
var objcnt=0;
var space=",";
var textObj="GetBillDate,chekChMemID,BillStartNumber,BillEndNumber";
var temp_Arr=textObj.split(space);
document.all[temp_Arr[objcnt]].select();
function funTextControl(){
	if (objcnt<3){
		if (event.keyCode==13){ //Enter換欄
			event.keyCode=0;
			event.returnValue=false;
			objcnt=objcnt+1;
			document.all[temp_Arr[objcnt]].focus();
		}
	}
}

function funCounterReceive(){
	addGetBillBase.onSubmit="";
	addGetBillBase.action="";
	addGetBillBase.target="";
	addGetBillBase.submit();
}
</script> 
</html>
<!-- #include file="../Common/ClearObject.asp" -->
