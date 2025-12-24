<!-- #include file="../Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
    BIllSN=request("BillSN")
    if BillSN="" then
        Response.write "<script>"
        Response.Write "alert('無舉發單編號帶入！');"
        Response.write "self.close();"
        Response.write "</script>"
    end if
	'
	SQL="select BillSN,OpenGovDate,OpenGovNumber,OpenGovSendDate,OpenGOvRecordMemberID,OpenGovRecordDate from BillMailHistory where BillSN =" & BIllSN
	set RsMailHisotry=conn.execute(SQL)
	if not RsMailHisotry.eof then
       OpenGovDate = ginitdt(RsMailHisotry("OpenGovDate"))
       OpenGovNumber = RsMailHisotry("OpenGovNumber")
       OpenGovSendDate= ginitdt(RsMailHisotry("OpenGovSendDate"))
   	else
        Response.write "<script>"
        Response.Write "alert('無該筆公示送達紀錄. ');"
        Response.write "self.close();"
        Response.write "</script>"
	end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>公示送達</title>
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
.style3 {font-size: 15px}
.style6 {font-size: 15pt}
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


<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="style3">公示送達 </span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">				

      <tr>
        <td width="11%" nowrap bgcolor="#FFFF99"><div align="right"><span class="style3">公告日期</span></div></td>
        <td width="89%" nowrap><span class="style3"><%=OpenGovDate%></span>
          </span></td>
      </tr>
      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right" class="style3">公示文號</div></td>
        <td nowrap><span class="style3"><%=OpenGovNumber%></span></td>
      </tr>
      <tr>
        <td nowrap bgcolor="#FFFF99"><div align="right" class="style3">
          <div align="right">修改人員</div>
        </div></td>
        <td nowrap>        <%
        if not RsMailHisotry.eof then
          if RsMailHisotry("OpenGOvRecordMemberID")<>"" then
			SQL="select ChName from MemberData where MemberID =" & RsMailHisotry("OpenGOvRecordMemberID")
			set rsMemberData=conn.execute(SQL)
            if not rsMemberData.eof then
            	response.write rsMemberData("ChName")
            end if
            response.write ginitDT( RsMailHisotry("OpenGOvRecordDate") )
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
        <span class="style3">        <img src="../Image/space.gif" width="20" height="8"></span>        
        <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉">
    </p>    </td>
  </tr>
  <tr>
    <td><p>&nbsp;</p></td></tr>
</table>

</body>
</html>
<!-- #include file="../Common/ClearObject.asp" -->
