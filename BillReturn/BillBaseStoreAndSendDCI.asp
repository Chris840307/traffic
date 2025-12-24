<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
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
<title>寄存送達</title>
<%
'檢查是否可進入本系統
'AuthorityCheck(253)

StoreANdSendMemID=trim(Session("User_ID"))

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

sQuerySQL=Session("BillSQLforStoreAndSendUpload")

if trim(request("kinds"))="DB_BillStoreAndSend" then
	if sQuerySQL="" then
		Response.write "<script>"
		Response.Write "alert('無查詢條件,請重新鍵入');"
		Response.write "self.close();"
		Response.write "</script>"		
	else
		set rsUPD1=conn.execute(sQuerySQL)		
		if rsupd1.eof then
			Response.write "<script>"
			Response.Write "alert('系統找不到該筆資料. 無法完成監理站寄存送達註記');"
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
			'高雄市要寫log到 BillReturnCenterJob
			if sys_City="高雄市" or sys_City="高港局" then
			
				if trim(request("store_BatchNumber"))<>"" then
					strBData="select a.FileName,b.UnitID from Dcilog a,MemberData b " &_
						" where a.Batchnumber='"&trim(request("store_BatchNumber"))&"' " &_
						" and rownum<=1 and a.RecordMemberID=b.MemberID "
					
					set rsBData=conn.execute(strBData)
					if not rsBData.eof then
						strBRCJ="Insert into BillReturnCenterJob(BatchNumber,FileName,UnitID,StoreAndSendDate" &_
							",StoreAndSendNumber,RecordMemberID,RecordDate) " &_
							" values('"&trim(request("store_BatchNumber"))&"','"&trim(rsBData("FileName"))&"'" &_
							",'"&trim(rsBData("UnitID"))&"',"&funGetDate(gOutDt(request("storeAndSendMailDate")),0) &_
							",'"&TRIM(request("StoreAndSendGovNumber"))&"',"&trim(Session("User_ID"))&",sysdate" &_
							")"
							'response.write strBRCJ
							'response.end
						conn.execute strBRCJ
					end if
					rsBData.close
					set rsBData=nothing
				end if
			end if

			While Not rsUPD1.Eof
				if request("StoreAndSendEffectDate") <> "" then 
					sStoreAndSendEffectDateSQL= " , StoreAndSendEffectDate=" & funGetDate(gOutDt(request("StoreAndSendEffectDate")),0) 
				end if
				sUpdSQL="update BillMailHistory set storeAndSendMailDate=" & funGetDate(gOutDt(request("storeAndSendMailDate")),0)  & _
					" , StoreAndSendSendDate="&funGetDate(gOutDt(request("storeAndSendSendDate")),0) &_
					" , StoreAndSendGovNumber='" & request("StoreAndSendGovNumber") & "'" & _
					" , StoreAndSendRecordMemberID=" & trim(Session("User_ID")) & _
					" , StoreAndSendRecordDate=SYSDate " & _				
					sStoreAndSendEffectDateSQL   & _
					" Where BillSN=" & rsUPD1("SN")
				conn.BeginTrans
				conn.Execute(sUpdSQL)
				funcSafeKeep conn,trim(rsUPD1("SN")),trim(rsUPD1("BillNo")),trim(rsUPD1("BillTypeID")),trim(rsUPD1("CarNo")),trim(rsUPD1("BillUnitID")),trim(rsUPD1("RecordDate")),trim(rsUPD1("RecordMemberID")),theBatchTime
	
				if err.number = 0 then
				   conn.CommitTrans
				else
				   conn.RollbackTrans
					Response.write "<script>"
					Response.Write "alert('處理過程發生異常 - 寄存送達 ' " & sUDPSQL & "');"
					Response.write "self.close();"
					Response.write "</script>"		
				   
				end if  
				rsUPD1.movenext						
			wend
			Response.write "<script>"
			Response.Write "alert('監理站寄存送達註記完成，批號："&theBatchTime&"');"
			Response.write "opener.myForm.submit();" 'by kevin
			Response.write "self.close();"
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
				<td colspan="4"><span class="pagetitle">寄存送達</span></td>
			</tr>
			<tr bgcolor="#EBFBE3">
			  <td width="15%" align="left" bgcolor="#FFFFFF"><table width="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
                  <td width="11%" nowrap bgcolor="#FFFFCC"><div align="right" class="content"><span class="style3">郵寄日期</span></div></td>
                  <td width="89%" nowrap> <div align="left"><span class="style3">
                    <input class="btn1" type='text' size='7' id='storeAndSendSendDate' name='storeAndSendSendDate' onblur="getDealLineDate()">
                  </span> </div></td>
                </tr>
                <tr>
                  <td width="11%" nowrap bgcolor="#FFFFCC"><div align="right" class="content"><span class="style3">寄存送達日期生效日</span></div></td>
                  <td width="89%" nowrap> <div align="left"><span class="style3">
                    <input class="btn1" type='text' size='7' id='storeAndSendMailDate' name='storeAndSendMailDate'">
                  </span> (應以開始寄存郵局日為上傳日，無須預加90天)</div></td>
                </tr>

                <tr>
                  <td nowrap bgcolor="#FFFFCC" class="content"><div align="right" class="content"><span class="style3">送達文號</span></div></td>
                  <td nowrap><div align="left">
                      <input name="StoreAndSendGovNumber" type="text" class="btn1" id="StoreAndSendGovNumber" size="9">
                      <span class="style3">
                      <input class="btn1" type='hidden' size='7' id='StoreAndSendEffectDate' name='StoreAndSendEffectDate'>
                      </span></div>
					  (文號最多為9個英文/數字,<strong>請勿輸入中文</strong>)
					</td>
                </tr>
                <tr>
                  <td nowrap bgcolor="#FFFFCC" class="content"><div align="right" class="style3">
                      <div align="right" class="style3">修改人員</div>
                  </div></td>
                  <td nowrap> <div align="left"><%=Session("Ch_Name")%></div></td>
                </tr>
              </table></td>
		  </tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="close" value=" 確 定 " onclick="BillStoreAndSend();" <%
						'1:查詢 ,2:新增 ,3:修改 ,4:刪除
						if CheckPermission(253,4)=false then
							response.write "disabled"
						end if
						%>>
				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
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
function BillStoreAndSend(){
	if (myForm.storeAndSendMailDate.value==""){
		alert("請先輸入寄存送達日期生效日！");
		result=false
	//}else if (myForm.StoreAndSendGovNumber.value==""){
	//	alert("請先輸入送達文號！");
	//	result=false
	}else{
		myForm.kinds.value="DB_BillStoreAndSend";
		myForm.submit();
	}
}
//寄存送達日期=郵寄日期+90
function getDealLineDate(){
	getDealDateValue=0
	myForm.storeAndSendSendDate.value=myForm.storeAndSendSendDate.value.replace(/[^\d]/g,'');
	BFillDateTemp=myForm.storeAndSendSendDate.value;
	if (BFillDateTemp.length >= 6){
		Byear=parseInt(BFillDateTemp.substr(0,BFillDateTemp.length-4))+1911;
		Bmonth=BFillDateTemp.substr(BFillDateTemp.length-4,2);
		Bday=BFillDateTemp.substr(BFillDateTemp.length-2,2);
		var BFillDate=new Date(Byear,Bmonth-1,Bday)
		var DLineDate=new Date()
		DLineDate=DateAdd("d",getDealDateValue,BFillDate);
		Dyear=parseInt(DLineDate.getYear())-1911;
		Dmonth=DLineDate.getMonth()+1;
		Dday=DLineDate.getDate();
		Dyear=Dyear.toString();
		if (Dmonth < 10){
			Dmonth="0"+Dmonth;
		}
		if (Dday < 10){
			Dday="0"+Dday;
		}
		myForm.storeAndSendMailDate.value=Dyear+Dmonth+Dday;
	}
}
</script>
</html>
