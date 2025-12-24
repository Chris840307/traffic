<!-- #include file="../Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
	SQL="select SN,BillNo,CarNo,IllegalDate from BillBase where SN =" & Request("BillSN") 
	set RsBillBase=Server.CreateObject("ADODB.RecordSet")
	RsBillBase.open SQL,Conn,3,3
	if not RsBillBase.eof then
		'先看是否已經有該筆退件資料	
		SQL="select BillNo,CarNo,MailDate,MailTypeID,MailNumber,MailReturnDate, ReturnResonID  " & _
				",StoreAndSendMailDate,StoreAndSendMailNumber,StoreAndSendMailReturnDate,StoreAndSendReturnResonID,ReturnRecordDate,ReturnRecordMemberID" & _
				" from BillMailHistory where BillSN =" & Request("BillSN") 
		set RsBillMailHistory=conn.execute(SQL)
		if RsBillMailHistory.eof then
	  		sql = "Insert into BillMailHistory (BillSN,BillNo,CarNo ) " & _
				   "values (" & Request("BillSN") & ",'" & RsBillBase("BillNo") & "','" & RsBillBase("CarNo") & "')" 

	  		Conn.Execute(sql)
	  		RsBillMailHistory.Close
			SQL="select BillNo,CarNo,MailDate,MailTypeID,MailNumber,MailReturnDate, ReturnResonID  " & _
					",StoreAndSendMailDate,StoreAndSendMailNumber,StoreAndSendMailReturnDate,StoreAndSendReturnResonID,ReturnRecordDate,ReturnRecordMemberID" & _
					" from BillMailHistory where BillSN =" & Request("BillSN") 
			set RsBillMailHistory=conn.execute(SQL)
			'Session("Msg")="該筆舉發單尚未郵寄無法退件 "
			'Response.write "<script>"
			'Response.Write "alert('該筆舉發單尚未郵寄無法退件！');"
			'Response.write "self.close();"
			'Response.write "</script>"
		end if
	else
			'Session("Msg")="該筆舉發單資料不存在,請確認  "
			Response.write "<script>"
			Response.Write "alert('該筆舉發單尚未郵寄無法退件！');"
			Response.write "self.close();"
			Response.write "</script>"		
	end if

	'退件原因
	SQL = "Select ID,Content from DCICode where TypeID=7 "
    set RsReturnReson=conn.execute(SQL)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>退件管理</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function datacheck()
{
	var result ;
	
  if(document.all.MailReturnDate.value=="")   
  {
    alert('請輸入退件日期');
    return false;  
  }	
	if(document.all.MailReturnDate.value!=""){
		if(!dateCheck(document.all.MailReturnDate.value)){
			alert("第一次退件日期輸入不正確!!");
				return false;  
		}
	}
	if(document.all.StoreAndSendMailReturnDate.value!=""){
		if(document.all.StoreAndSendMailReturnDate.value!=""){
			if(!dateCheck(document.all.StoreAndSendMailReturnDate.value)){
				alert("第二次退件日期輸入不正確!!");
					return false;  
			}
		}
	}
	
  /* 
  if(document.all.ReturnReson.value=="")   
  {
    alert('請選擇退件原因');
    return false;  
  }
  
  if(document.all.ReturnDate.value=="")   
  {
    alert('請輸入退件日期');
    return false;  
  } 

  */
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
.style3 {font-size: 15px}
-->
</style></head>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<!-- #include file="../Common/checkFunc.inc"-->
<body>
<%
if Session("Msg")<>"" then
	 Response.write "<font  color='Red' size='15'>" & Session("Msg") & "</font>"	
 
end if	

%>		
<FORM NAME="updBillReturn" ACTION="BillReturn_mdy.asp" METHOD="POST" onSubmit="return datacheck();">  
	<input type="hidden" name="tag" value="<%
		if RsBillMailHistory.eof then 
			response.Write "NEW"
		else
			response.write "UPD"
		end if	
	%>"> 
	<input type="hidden" name="BillSN" value="<%=RsBillBase("SN")%>">	
	<input type="hidden" name="BillNo" value="<%=RsBillBase("BillNo")%>">
	<input type="hidden" name="CarNo" value="<%=RsBillBase("CarNo")%>">	

<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle style3">退件管理</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right" class="style3">舉發單號</div></td>
        <td width="10%"><div align="center"><span class="style3"> <%=RsBillBase("BillNo")%> </span></div></td>
        <td width="8%" nowrap bgcolor="#FFFF99"><div align="center"><span class="style3">車號</span></div></td>
        <td width="11%"><div align="center"><span class="style3"> <%=RsBillBase("CarNo")%> </span></div></td>
        <td width="8%" nowrap bgcolor="#FFFF99"><div align="center"><span class="style3">違規日期</span></div></td>
        <td><span class="style3"> <%=gInitDT(RsBillBase("IllegalDate"))%> </span></td>
        </tr>

      <tr>
        <td width="11%" nowrap bgcolor="#FFFF99"><div align="right" class="style3">大宗掛號</div></td>
        <td colspan="5"><div align="left" class="style3">
							<% if not RsBillMailHistory.eof then 
									response.Write RsBillMailHistory("MailNumber")
								end if
							%> 
		</div></td>
      </tr>				

      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right"><span class="style3">第一次郵寄日期</span></div></td>
        <td><span class="style3">
        
		  		  					<% if not RsBillMailHistory.eof then
										  response.Write gInitDT(RsBillMailHistory("MailDate"))
									  end if
									%>
							
          </span></td>
        <td nowrap bgcolor="#FFFF99"><span class="style3">退件日期</span></td>
        <td nowrap><span class="style3">
          <input name="MailReturnDate" type="text" id="MailReturnDate3" size="10" maxlength="8"		  
			  		  					<% if not RsBillMailHistory.eof then 
											response.Write "value=" 
											response.Write gInitDT(RsBillMailHistory("MailReturnDate"))
											end if
										%>	 		  
		  >
          <input type="button" name="datestr" value="..." onclick="OpenWindow('MailReturnDate');">
</span></td>
        <td bgcolor="#FFFF99"><span class="style3">退件原因</span></td>
        <td><p class="style3">
            <select name="ReturnResonID" id="select4" >
                <option  value="" selected>選擇退件原因...</option>
                <%
p = 0
While Not RsReturnReson.Eof
%>
                <option value="<%=RsReturnReson("ID")%>" 
   	 <%if not RsBillMailHistory.eof then
			if RsReturnReson("ID")=RsBillMailHistory("ReturnResonID") then 
				response.write " selected" 
			end if
		end if
	 %>
	
	  ><%=RsReturnReson("Content")%> </option>
                <%
  p = p + 1
  RsReturnReson.MoveNext
Wend
RsReturnReson.MoveFirst
%>
            </select>
            <br>
            <input name="firstisstoreandsend" type="checkbox" id="firstisstoreandsend" value="yes"> 
            第一次寄送含送達證書</p>          </td>
        </tr>
      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right"><span class="style3">第二次郵寄日期</span></div></td>
        <td><span class="style3">
          <% if not RsBillMailHistory.eof then
										  response.Write gInitDT(RsBillMailHistory("StoreAndSendMailDate"))
									  end if
									%>
          </span></td>
        <td nowrap bgcolor="#FFFF99"><span class="style3">退件日期</span></td>
        <td nowrap><span class="style3">
          <input name="StoreAndSendMailReturnDate" type="text" id="StoreAndSendMailReturnDate2" size="10" maxlength="8"		  
			  		  					<% if not RsBillMailHistory.eof then 
											response.Write "value=" 
											response.Write gInitDT(RsBillMailHistory("StoreAndSendMailReturnDate"))
											end if
										%>	 		  
		  >
          <input type="button" name="datestr2" value="..." onclick="OpenWindow('StoreAndSendMailReturnDate');">
</span></td>
        <td bgcolor="#FFFF99"><span class="style3">退件原因</span></td>
        <td><span class="style3">
          <select name="StoreAndSendReturnResonID" id="StoreAndSendReturnResonID" >
            <option  value="" selected>選擇退件原因...</option>
            <%
p = 0
While Not RsReturnReson.Eof
%>
            <option value="<%=RsReturnReson("ID")%>" 
   	 <%if not RsBillMailHistory.eof then
			if RsReturnReson("ID")=RsBillMailHistory("StoreAndSendReturnResonID") then 
				response.write " selected" 
			end if
		end if
	 %>
	
	  ><%=RsReturnReson("Content")%> </option>
            <%
  p = p + 1
  RsReturnReson.MoveNext
Wend
%>
          </select>
        </span></td>
        </tr>
      <tr>
        <td bgcolor="#FFFF99"><div align="right" class="style3">
            <div align="right">修改人員</div>
        </div></td>
        <td colspan="5">

        <%if not RsBillMailHistory.eof then
        	if RsBillMailHistory("ReturnRecordMemberID")<>"" then
                sql="select ChName from MemberData where MemberID =" & RsBillMailHistory("ReturnRecordMemberID")

                set rsMemberData=conn.execute(SQL)
                if not rsMemberData.eof then
                    response.write rsMemberData("ChName")
                end if
                response.write " ( 修改時間 : " & ginitDT( RsBillMailHistory("ReturnRecordDate") ) & ")"
            end if
        end if

        %></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
        <input type="submit" name="Submit423" value="確 定" 
			 <% if Session("Msg") <> "" then 
			 	response.write "disabled" 
				 Session("Msg") = "" 
			 end if %>  
		>
        <span class="style3">
        	<!--
        <input name="wantDCIReturn" type="checkbox" id="wantDCIReturn" value="dcireturn" checked>
        <strong>退件標記紀錄. 單退監理所標記</strong>        
        -->
        <img src="../Image/space.gif" width="20" height="8"></span>        
        <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉">
    </p>    </td>
  </tr>
  <tr>
    <td><p>&nbsp;</p></td></tr>
</table>
</FORM>
</body>
</html>
<!-- #include file="../Common/ClearObject.asp" -->
