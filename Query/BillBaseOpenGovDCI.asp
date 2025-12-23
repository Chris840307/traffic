<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<script language="JavaScript">
	window.focus();
</script>
<style type="text/css">
<!--
.style3 {font-size: 15px}
-->
</style>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<!--#include file="../query/sqlDCIExchangeData.asp"-->
<title>公示送達</title>
<%
Server.ScriptTimeout = 68000
Response.flush
'檢查是否可進入本系統
'AuthorityCheck(253)
sQuerySQL=Session("BillSQLforOpenGovUpload")

strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=Nothing

if trim(request("kinds"))="DB_BilltoPublic" then
	if sQuerySQL="" then
		Response.write "<script>"
		Response.Write "alert('無查詢條件,請重新鍵入');"
		Response.write "self.close();"
		Response.write "</script>"		
	Else
		set rsUPD1=conn.execute(sQuerySQL)		
		if rsUPD1.eof then 
			Response.write "<script>"
			Response.Write "alert('系統找不到該筆資料. 無法完成監理站公示送達註記');"
			Response.write "self.close();"
			Response.write "</script>"					
		else
			strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
			set rsSN=conn.execute(strSN)
			if not rsSN.eof then
				theBatchTime=(year(now)-1911)&"N"&trim(rsSN("SN"))
			end if
			rsSN.close
			set rsSN=nothing
			While Not rsUPD1.Eof
				if request("OpenGovSendDate") <> "" then 
					sOpenGovSendDateSQL= " , OpenGovSendDate=" & funGetDate(gOutDt(request("OpenGovSendDate")),0) 
				end if
				sUpdSQL="update BillMailHistory set OpenGovDate=" & funGetDate(gOutDt(request("OpenGovDate")),0)  & _
					" , OpenGovNumber='" & request("OpenGovNumber") & "'" & _
					" , OpenGovRecordMemberID=" & trim(Session("User_ID")) & _
					" , OpenGovRecordDate=SYSDate " & _
					sOpenGovSendDateSQL   & _
					" Where BillSN=" & rsUPD1("SN")
				Conn.BeginTrans
	
				Conn.Execute(sUpdSQL)
			if sys_City="新北市" then	
				sUpdSQL2="insert into OPENGOVDATA(BILLSN,BILLNO,BATCHNUMBER,OPENYEAR,OPENNO,OPENPAGE,REPORTNO) " &_
					" values("&rsUPD1("SN")&",'','"&theBatchTime&"','"&Trim(request("OPENYEAR"))&"'" &_
					",'"&Trim(request("OPENNO"))&"','"&Trim(request("OPENPAGE"))&"'" &_
					",'"&Trim(request("REPORTNO"))&"'" &_
					")"
				conn.execute sUpdSQL2
			End if
			
				funcPublic conn,trim(rsUPD1("SN")),trim(rsUPD1("BillNo")),trim(rsUPD1("BillTypeID")),trim(rsUPD1("CarNo")),trim(rsUPD1("BillUnitID")),trim(rsUPD1("RecordDate")),trim(rsUPD1("RecordMemberID")),theBatchTime
				
				if err.number = 0 then
				   conn.CommitTrans
				else            
				   conn.RollbackTrans
					Response.write "<script>"
					Response.Write "alert('處理過程發生異常 - 公示送達 ' " & sUDPSQL & "');"
					Response.write "self.close();"
					Response.write "</script>"		
			   
				end if  
			rsUPD1.movenext
			wend
			Response.write "<script>"
			Response.Write "alert('監理站公示送達註記完成，批號："&theBatchTime&"');"
			Response.write "self.close();"
			Response.write "opener.myForm.submit();" 'by kevin
			Response.write "</script>"			
		end if
	end if	

end if
%>

</head>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4"><span class="pagetitle">公示送達</span></td>
			</tr>
			<tr bgcolor="#EBFBE3">
			  <td width="15%" align="left" bgcolor="#FFFFFF"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
                <tr>
                  <td width="11%" nowrap bgcolor="#FFFF99" class="content"><div align="right"><span class="style3">公告日期</span></div></td>
                  <td width="89%" nowrap> <div align="left"><span class="style3">
<input class="btn1" type='text' size='7' id='OpenGovDate' name='OpenGovDate'">                    
<input type="button" name="datestra" value="..." onclick="OpenWindow('OpenGovDate');">
                 </span>  </div>請輸入 公告日期，<br>監理站會固定加上行政程序法規定期限作為 公告生效日</td>
                </tr>
                <tr>
                  <td nowrap bgcolor="#FFFF99" class="content"><div align="right" class="style3">公示文號</div></td>
                  <td nowrap>                    <div align="left">
                      <input name="OpenGovNumber" type="text" class="btn1" id="OpenGovNumber" size="20" maxlength="9">(文號最多為9個英文/數字,<strong>請勿輸入中文</strong>)
                      <span class="style3">
                      <input class="btn1" type='hidden' size='7' id='OpenGovSendDate' name='OpenGovSendDate'>
                      </span></div></td>
                </tr>
<%if sys_City="新北市" then	%>
				<tr>	
                  <td nowrap bgcolor="#FFFF99" class="content" colspan="2"><div align="right" class="style3">
                      <div align="left"><strong>市&nbsp; &nbsp;府&nbsp; &nbsp;公&nbsp; &nbsp;報</strong></div>
                  </div></td>
                </tr>
				<tr>	
                  <td nowrap bgcolor="#FFFF99" class="content"><div align="right" class="style3">
                      <div align="right">出版年度</div>
                  </div></td>
                  <td nowrap>
				   <input class="btn1" type='text' size='7' name='OPENYEAR'>
				  </td>
                </tr>
				<tr>	
                  <td nowrap bgcolor="#FFFF99" class="content"><div align="right" class="style3">
                      <div align="right">期數</div>
                  </div></td>
                  <td nowrap>
				   <input class="btn1" type='text' size='7' name='OPENNO'>
				  </td>
                </tr>
				<tr>	
                  <td nowrap bgcolor="#FFFF99" class="content"><div align="right" class="style3">
                      <div align="right">頁次</div>
                  </div></td>
                  <td nowrap>
				   <input class="btn1" type='text' size='7' name='OPENPAGE'>
				  </td>
                </tr>
				<tr>	
                  <td nowrap bgcolor="#FFFF99" class="content"><div align="right" class="style3">
                      <div align="right">刊載碼</div>
                  </div></td>
                  <td nowrap>
				   <input class="btn1" type='text' size='7' name='REPORTNO'>
				  </td>
                </tr>
<%End If %>
                <tr>	
                  <td nowrap bgcolor="#FFFF99" class="content"><div align="right" class="style3">
                      <div align="right">修改人員</div>
                  </div></td>
                  <td nowrap> <div align="left"><%=Session("Ch_Name")%></div></td>
                </tr>
              </table></td>
		  </tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="clickbutton" value=" 確 定 " onclick="BillOpenGov();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
						if CheckPermission(253,4)=false then
							'response.write "disabled"
						end if
						%>>

				<input type="button" name="closebutton" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				</td>
			</tr>
		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script language="JavaScript">
function BillOpenGov(){
	var NoErr=0;
	for (x = 0;x < myForm.OpenGovNumber.value.length ; x++){
		//alert(myForm.OpenGovNumber.value.substr(x,1).charCodeAt());
		if (myForm.OpenGovNumber.value.substr(x,1).charCodeAt() > 127){
			NoErr=1;
			break;
		}
	}
	if (myForm.OpenGovDate.value==""){
		alert("請先輸入公告日期！");
		result=false;
	}else if (myForm.OpenGovNumber.value==""){
		alert("請先輸入公示文號！");
		result=false;
	}else if (NoErr==1){
		alert("公示文號請勿輸入中文！");
		result=false;
	}else{
		myForm.kinds.value="DB_BilltoPublic";
		myForm.submit();
		myForm.clickbutton.disabled=true;
		myForm.closebutton.disabled=true;
		myForm.clickbutton.value="存檔中,請稍候.....";
	}	
}
</script>
</html>
