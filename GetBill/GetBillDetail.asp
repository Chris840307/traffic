<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>領單管理-資料異動</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
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
<!-- #include file="../Common/checkFunc.inc"-->
<Script language="JavaScript">
<!--	
function qryCheck()
{
	var chkResult ;
	var rtnChkBillNum ;
	  
  if ((document.GetBillDetail.BillStartNumber_q.value=="")){
  	alert('請輸入舉發單起始碼!!');
  	document.GetBillDetail.BillStartNumber_q.focus();
  	return false;
  }  
  if ((document.GetBillDetail.BillEndNumber_q.value=="")){
  	alert('請輸入舉發單截止碼!!');
  	document.GetBillDetail.BillEndNumber_q.focus();
  	return false;
  }

  if (document.GetBillDetail.BillStartNumber_q.value!=""){
     chkResult = chkBillNumber(document.all.BillStartNumber_q,"[舉發單起始碼] 格式錯誤!!"); 
     if (chkResult != "Y") return false;
  }
  if (document.GetBillDetail.BillEndNumber_q.value!=""){
     chkResult = chkBillNumber(document.all.BillEndNumber_q,"[舉發單截止碼] 格式錯誤!!"); 
     if (chkResult != "Y") return false;    
  }  
   
  if (document.GetBillDetail.BillStartNumber_q.value!="" && document.GetBillDetail.BillEndNumber_q.value!=""){  
     rtnChkBillNum = ValidateBillNumbers (document.GetBillDetail.BillStartNumber_q,document.GetBillDetail.BillEndNumber_q,'N');
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
}	

function sendQry(){
	var rtn;
	var titleA="<%=Request("billstartnumber")%>".substr(1,3);
	var numA="<%=Request("billstartnumber")%>".substr(4,6);

	var titleB="<%=Request("billendnumber")%>".substr(1,3);
	var numB="<%=Request("billendnumber")%>".substr(4,6);

	var titleC=document.all.BillStartNumber_q.value.substr(1,3);
	var numC=document.all.BillStartNumber_q.value.substr(4,6);

	var titleD=document.all.BillEndNumber_q.value.substr(1,3);
	var numD=document.all.BillEndNumber_q.value.substr(4,6);

	rtn = true;
	rtn = qryCheck();
	if (rtn!=false){
		if(numA>numC||numB<numD){
			alert("查詢單號超出領單範圍!!");
			rtn = false;
		}
	}
	if (rtn!=false){
     var form_A= document.forms[0];
     form_A.qryType.value="2";
     form_A.action = "GetBillDetail.asp";
     form_A.submit();		
	}
}
function sendDefUpdate(){
	var rtn;
	var param ;
	var billstartnumber;
	var billendnumber;

	var titleA="<%=Request("billstartnumber")%>".substr(1,3);
	var numA="<%=Request("billstartnumber")%>".substr(4,6);

	var titleB="<%=Request("billendnumber")%>".substr(1,3);
	var numB="<%=Request("billendnumber")%>".substr(4,6);

	var titleC=document.all.BillStartNumber_q.value.substr(1,3);
	var numC=document.all.BillStartNumber_q.value.substr(4,6);

	var titleD=document.all.BillEndNumber_q.value.substr(1,3);
	var numD=document.all.BillEndNumber_q.value.substr(4,6);

	rtn = true;
	rtn = qryCheck();
	if (rtn!=false){
		if(numA>numC||numB<numD){
			alert("查詢單號超出領單範圍!!");
			rtn = false;
		}
	}
	 if (GetBillDetail.BillStartNumber_q.value==''||GetBillDetail.BillEndNumber_q.value==''){
		billstartnumber=GetBillDetail.billstartnumber.value;
		billendnumber=GetBillDetail.billendnumber.value;
	 }else{
		billstartnumber=GetBillDetail.BillStartNumber_q.value;
		billendnumber=GetBillDetail.BillEndNumber_q.value;
	 }
	 if(rtn!=false){
		 if(confirm("是否整批設定特殊狀態?")){
			param = "GetBillDetail_mdy.asp?billstartnumber=" + billstartnumber + "&billendnumber=" + billendnumber + "&noteContent=" + GetBillDetail.DefNote.value + "&BillStateId=" + GetBillDetail.DefBillStateID.value ;
			exportExcel(param,'UpdGetBillDetail') ;
		}
	}
}

function sendUpdate(getBillSn,billNo,noteContent,BillStateId){
	 var param ;
	 param = "GetBillDetail_mdy.asp?getBillSn=" + getBillSn + "&billNo=" + billNo + "&noteContent=" + noteContent + "&BillStateId=" + BillStateId ;
   exportExcel(param,'UpdGetBillDetail') ;	
}

function sendLoss(){
	var rtn;
	var BillStartNumber_q;
	var BillEndNumber_q;
	var BillStartNumber;
	var BillEndNumber;
	
  rtn = true;
	if ((document.GetBillDetail.BillStartNumber_q.value!="") || (document.GetBillDetail.BillEndNumber_q.value!="")) 
	  rtn = qryCheck();  
	//rtn = qryCheck();
	if (rtn!=false){
		 BillStartNumber_q = document.GetBillDetail.BillStartNumber_q.value;
		 BillEndNumber_q = document.GetBillDetail.BillEndNumber_q.value;
		 BillStartNumber = '<%=Request("billstartnumber")%>';
		 BillEndNumber = '<%=Request("billendnumber")%>';
		 exportExcel('BillAudit2.asp?BillStartNumber_q=' + BillStartNumber_q + '&BillEndNumber_q=' + BillEndNumber_q + '&BillStartNumber=' + BillStartNumber + '&BillEndNumber=' + BillEndNumber,'BillAudit2');
	}
}
/*
function sendUpdate(content,stateId){
	var form_A= document.forms[0];
	form_A.NoteContent.value = content ;
	form_A.BillStateId.value = stateId ;
	form_A.action = "GetBillDetail_mdy.asp";
	form_A.submit();
}
*/
-->
</Script>
<body>
<%
if Session("Msg")<>"" then
	 Response.write "<font  color='Red' size='2'>" & Session("Msg") & "</font>"
	 Session("Msg") = ""
end if	

qryType = Request("qryType")
Select Case qryType
	Case "1" :
    billstartnumber = Request("billstartnumber")
    billendnumber = Request("billendnumber")    
    sql = "Select a.*, c.ChName as GetBillChName , d.UnitName from GetBillDetail a , GetBillBase b , MemberData c, UnitInfo d Where a.GetBillSn=" & Request("GETBILLSN") & " and a.billno between " & _
          "'" & billstartnumber & "' and '" & billendnumber & "'" &_
          " and a.GetBillSN=b.GetBillSN and b.GetBillMemberID=c.MemberID "&_
             " and c.UnitID=d.UnitID "
  Case "2" :
    billstartnumber = trim(Request("BillStartNumber_q"))
    billendnumber = Request("BillEndNumber_q")
		
    'sql = "Select a.* , b. from GetBillDetail Where billno between " & _
    '      "'" & billstartnumber & "' and '" & billendnumber & "'"  & _ 
    '      " And BillStateID <> 463 "
		sql = "Select a.* , c.ChName as GetBillChName  , d.UnitName from GetBillDetail a , GetBillBase b , MemberData c, UnitInfo d Where a.billno between " & _
          "'" & billstartnumber & "' and '" & billendnumber & "'"  & _ 
          " and a.GetBillSN=b.GetBillSN and b.GetBillMemberID=c.MemberID "  & _  
          " and c.UnitID=d.UnitID "
		  
End Select
sql = sql & " Order By 2 , BillStateID desc"

Session("DetailSQL") = sql
set Rs=Server.CreateObject("ADODB.RecordSet")
rs.cursorlocation = 3
rs.open SQL,Conn,3,1
%>	
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">領單管理-資料異動</span></td>
  </tr>

<FORM NAME="GetBillDetail" ACTION="" METHOD="POST">  
	<input type="hidden" name="qryType" value="<%=qryType%>"> 
	<input type="hidden" name="GETBILLSN" value="<%=Request("GETBILLSN")%>"> 
	<input type="hidden" name="getbilldate" value="<%=Request("getbilldate")%>"> 
	<input type="hidden" name="chname" value="<%=Request("chname")%>"> 
	<input type="hidden" name="billstartnumber" value="<%=Request("billstartnumber")%>"> 
	<input type="hidden" name="billendnumber" value="<%=Request("billendnumber")%>"> 
	<input type="hidden" name="nbchars" >
	<input type="hidden" name="NoteContent">
	<input type="hidden" name="BillStateId">
	
  <tr>
    <td height="32" bgcolor="#CCCCCC">
    	 <table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
          <tr>
            <td height="21"><span class="font12">領單日期 : <%=Request("getbilldate")%> &nbsp;&nbsp;&nbsp;
            	   領單人員 : <%=Request("chname")%> <img src="space.gif" width="20" height="5">&nbsp;&nbsp;&nbsp;
            	   舉發單號 : <%=billstartnumber%>~<%=billendnumber%><img src="space.gif" width="12" height="5"> 
            	   舉發單號 <input name="BillStartNumber_q" type="text" value="<%=trim(Request("BillStartNumber_q"))%>" size="10" maxlength="9" class="btn1">
                  &nbsp;~&nbsp;<input name="BillEndNumber_q" type="text" value="<%=trim(Request("BillEndNumber_q"))%>" size="10" maxlength="9" class="btn1">
                <img src="space.gif" width="20" height="5">            
                <input type="button" onclick="sendQry();" name="Submit3" value="查詢" <%=ReturnPermission(CheckPermission(221,1))%>>
                <input type="hidden" name="Submit33" onclick="sendLoss()" value="漏號稽核"></span>
				<br><%
				strSQL = "Select id,content From Code Where TypeId=17 and ID in(555,463,461,462,460,464,459) order by showorder"
				set rsdef=conn.execute(strSQL)
				response.write "<select name=""DefBillStateID"">"
				while Not rsdef.eof
					Response.Write "<option value='" & trim(rsdef("id")) & "'>" & trim(rsdef("content")) & "</option>"
					rsdef.movenext
				wend
				rsdef.close
				response.write "</select>"%>&nbsp;&nbsp;&nbsp;
				<input name="DefNote" type="text" value="" size="30" class="btn1">
				<input type="button" onclick="sendDefUpdate();" name="Submit3" value="特殊狀態整批設定" <%=ReturnPermission(CheckPermission(221,1))%>>
            </td>
          </tr>
       </table>
    </td>
  </tr>
</FORM>  
  <tr>
    <td height="26" bgcolor="#FFCC33"><span class="pagetitle">領取舉發單列表</span></td>
  </tr>
<%
if not rs.EOF then
	actionPage=cint(0 & trim(request("page"))) 
	if actionPage < 1 then actionPage=1
	rs.PageSize=PageSize
	if actionPage > rs.PageCount then actionPage=rs.PageCount
	rs.AbsolutePage=actionPage 
%>  
  <tr>
    <td bgcolor="#E0E0E0">
    	 <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="1">
          <tr bgcolor="#EBFBE3" height="30">
            <th width="10%" nowrap ><span class="font12">舉發單號</span></th>
            <th width="15%" nowrap><span class="font12">領取單位</span></th>
            <th width="8%" nowrap><span class="font12">領單人員</span></th>
            <!-- <th width="10%" height="15" nowrap><span class="font12">設定</span></th>-->
            <!-- <th width="16%" nowrap><span class="font12">特殊說明紀錄人員</span></th> -->
            <th width="10%" nowrap><span class="font12">記錄時間</span></th>
            <th width="37%" height="15" nowrap><span class="font12"> 特殊狀態記錄</span></th>
          </tr>

<%        
  sqlTemp = "Select id,content From Code Where TypeId=17 and ID in(555,463,461,462,460,464,459) order by showorder"
  Set RsTemp=Server.CreateObject("ADODB.RecordSet")
  
	RsTemp.cursorlocation = 3
  RsTemp.open sqlTemp,Conn,3,1 
  Set dicObj = Server.CreateObject("Scripting.Dictionary")
  dicObj.RemoveAll
  While Not RsTemp.Eof
     idStr = RsTemp("id")
     contentStr = RsTemp("content")
     dicObj.Add idStr,contentStr
     RsTemp.MoveNext
  Wend
  if RsTemp.state then RsTemp.close
  
  
  
	for I=1 to rs.pagesize
	   RecordDateTemp = ""
	   ChName = ""
	   if (rs("RecordDate") & "") <> "" Then
	      RecordDateTemp = gInitDT(rs("RecordDate"))
	   end if
	   RecordMemberID = rs("RecordMemberID") & ""
	   if RecordMemberID <> "" Then
	      sqlTemp = "Select chname From MemberData Where MemberID=" & Int(RecordMemberID)
	      RsTemp.cursorlocation = 3	      
        RsTemp.open sqlTemp,Conn,3,1
        if Not RsTemp.Eof Then
           ChName = RsTemp("ChName")
           if RsTemp.state then RsTemp.close
        end if
	   End If
	   BillStateId = rs("BillStateId")
%>      
         <tr bgcolor="#FFFFFF" height="30" onMouseOver="this.className='listtitle2'" onMouseOut="this.className='listtitle1'">
           <td><span class="font12"><div align="center"><%=rs("BillNo")%></div></span></td>
           <td><span class="font12"><div align="center"><%=rs("UnitName")%></div></span></td>
           <td><span class="font12"><div align="center"><%=rs("GetBillChName")%></div></span></td>

            <!--<td><%=ChName%></td>-->
           <td><span class="font12"><div align="center"><%=RecordDateTemp%></div></span></td>
           <td height="10"><img src="../image/space.gif"></img>
              <select name="BillStateID_<%=I%>">
<%
         keysTemp = dicObj.Keys
         itemsTemp = dicObj.Items
         keyStr = ""
         itemStr = ""
         For k = 0 To dicObj.Count - 1
           if (CStr(keysTemp(k))=CStr(BillStateId)) Then
           	  strTemp = " selected "
           else
           	  strTemp = ""
           end if
           keyStr = keysTemp(k)
           itemStr = itemsTemp(k)
           Response.Write "<option value='" & keyStr & "'" & strTemp & ">" & itemStr & "</option>"
         Next
%>
                 </select>
             <input name="NoteContent_<%=I%>" type="text" value="<%=rs("NoteContent")%>" size="21" maxlength="50" onKeyDown="calStr(this,50);" onKeyUp="calStr(this,50);" class="btn1">
             <input type="button" name="Submit4" value="確定" onclick="sendUpdate('<%=rs("GetBillSn")%>','<%=rs("BillNo")%>',document.all.NoteContent_<%=I%>.value,document.all.BillStateID_<%=I%>.value);" <%=ReturnPermission(CheckPermission(221,3))%>>
           </td>
         </tr>
<%              
		rs.Movenext              
		If rs.EOF then exit for              
	next              
%> 
       </table>
    </td>
  </tr>
<%
urlParam = "&qryType=" & Request("qryType") & "&GETBILLSN=" & Request("GETBILLSN") & "&chname=" & Request("chname") & _
           "&billstartnumber=" & Request("billstartnumber") & "&billendnumber=" & Request("billendnumber") & _
           "&BillStartNumber_q=" & Request("BillStartNumber_q") & "&BillEndNumber_q=" & Request("BillEndNumber_q") & _
           "&getbilldate=" & Request("getbilldate")
%>			  
  <tr>
    <td align="center" height="35" bgcolor="#FFDD77">    	
    	 <font size="2"><%ShowPageLink actionPage,rs.PageCount,"GetBillDetail.asp",urlParam%></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
       <input type="button" name="Submit423" value="轉換成Excel" onclick="exportExcel('saveDetailExcel.asp','saveDetailExcel')">
    </td>
  </tr>
<% else %>    
  <tr>
  	 <td align="center" >        
	      <center><font  color="Red" size="2">              
	<%              
	Response.Write "目前查無任何資料 ..."              
	%>              
	      </font></center><br> 
	   </td>
	</tr>             
<%              
end if              
rs.close              
set rs = nothing              
%>     

</table>
</body>
</html>
<!-- #include file="../Common/ClearObject.asp" -->
