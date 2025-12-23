<!--#include virtual="Traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/OldData.INI"-->
<!--#include virtual="Traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舊資料備註</title>
<%
    function QuotedStr(Str)
        QuotedStr="'"+Str+"'"
    end function
    

    if request("DB_Selt")="Selt" then
        SelectSQL = "select tkt_no from Vil_Rec_Note where tkt_no=" & QuotedStr(trim(request("BillNo")))
        set Rs1 = Server.CreateObject("ADODB.RecordSet")
        Set Rs1 = Conn.Execute(SelectSQL)
        if not Rs1.eof then
            sql="update vil_rec_note set Note=" & QuotedStr(trim(request("content"))) & ",RecordTime=" &_
            "TO_DATE(" & QuotedStr( DateValue(Date()) & " " & Hour(Time()) & ":" & Minute(Time()) & ":" & Second(Time())) &  ",'YYYY/MM/DD/HH24/MI/SS')" & "," &_
            "recordusername=" & QuotedStr(request("User_id")) & " where tkt_no=" & QuotedStr(trim(request("BillNo")))
        else  
            sql = "Insert into Vil_Rec_Note values (" & QuotedStr(trim(request("BillNo"))) & "," & QuotedStr(trim(request("content"))) & "," &_
                   "TO_DATE(" & QuotedStr( DateValue(Date()) & " " & Hour(Time()) & ":" & Minute(Time()) & ":" & Second(Time())) &  ",'YYYY/MM/DD/HH24/MI/SS')" & "," &_
                  QuotedStr(request("User_id")) & ")"
        end if         
        'response.Write sql
        'response.End  
        Conn.Execute(sql)
        response.write "<script>"
        response.write " alert(""儲存完畢"");"
        response.write "window.opener.myForm.submit();"
        response.write "window.close();" 
        response.write "</script>"  
    end if  
 %>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<SCRIPT LANGUAGE=javascript>
<!--
function datacheck()
{
	
  if(document.all.content.value=="")   
  {
    alert('請輸入備註內容');
    return false;  
  }	
  
  addUnitInfo.DB_Selt.value="Selt";
  addUnitInfo.submit();	 
}

  function textCounter(field,countfield,maxlimit) {
      if (field.value.length > maxlimit) {
        field.value = field.value.substring(0,maxlimit);
      }
      else {
        countfield.value = maxlimit - field.value.length;
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
.style3 {font-size: 15px}
-->
</style></head>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<!-- #include virtual="traffic/Common/checkFunc.inc"-->
<body>
<%
if Session("Msg")<>"" then
	 Response.write "<font  color='Red' size='2'>" & Session("Msg") & "</font>"
	 Session("Msg") = ""
end if	
%>	
<FORM NAME="addUnitInfo" METHOD="POST">  	
<table width="100%" height="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle style3">舊資料備註</span></td>
  </tr>
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td bgcolor="#FFFFCC"><div align="right"><span class="style3 style3">單號</span></div></td>
        <td><input name="BillNO" type="text" id="Text1" value=<%=request("BillNo") %>  maxlength="10" style="width: 112px" readonly="readOnly"></td>
      </tr>
      <%
        
        ContentSQL = "select Note from Vil_Rec_Note where tkt_no=" & QuotedStr(trim(request("BillNo")))
        set RS_Content = Server.CreateObject("ADODB.RecordSet")
        Set RS_Content = Conn.Execute(ContentSQL)
        if not RS_Content.eof then
            NoteContent =  trim(RS_Content("Note"))
        end if  
       %>
      <tr>
        <td bgcolor="#FFFFCC" style="height: 36px"><div align="right" class="style3 style3">
            <div align="right" class="style3">
              <div align="right"><font color="red">*</font>備註內容</div>
            </div>
        </div></td>
        <td style="height: 36px">
            <textarea name="content" id="content" cols="50" rows="5" wrap="physical"
                onKeyDown="textCounter(content,remLen,300);" 
                onKeyUp="textCounter(content,remLen,300);"><%=NoteContent %></textarea>
            <br>
            尚能輸入
            <input type="text" name="remLen" size="4" maxlength="3" value="300" readonly>
            個字元    
        </td>
      </tr>
      <tr>
        <td bgcolor="#FFFFCC"><div align="right" class="style3">
            建立人員</div></td>
        <td><%=request("UserID") %></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="35" bgcolor="#FFDD77"><p align="center" class="style1">
        <input type="button"name="Summit" value="確 定" onclick="datacheck();">
        <span class="style3"><img src="space.gif" width="9" height="8"></span>        <input type="button" name="Submit4232" onClick="javascript:window.close();" value="關 閉">
</p>    </td>
  </tr>
  <tr>
    <td>    <p align="center">&nbsp;
      </p>    <p>&nbsp;</p>
    <p>&nbsp;</p></td></tr>
</table>
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="USER_ID" value="<%=request("UserID") %>">
</FORM>
</body>
</html>
<!-- #include virtual="traffic/Common/ClearObject.asp" -->
