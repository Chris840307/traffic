
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/bannernodata.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<%
sqlUnit = "Select UnitName , UnitID from UnitInfo"
set RsUnit=Server.CreateObject("ADODB.RecordSet")
RsUnit.open sqlUnit,Conn,3,3

sql = "Select ChName , MemberID from MemberData where UnitID in('" & Request("UnitID_q") & "')"
set RsUpd1=Server.CreateObject("ADODB.RecordSet")
RsUpd1.open sql,Conn,3,3

if Request("BillStartNumber_q")<>"" and Request("BillEndNumber_q")<>"" then
	BillStartNumber=trim(Request("BillStartNumber_q"))
	BillEndNumber=trim(Request("BillEndNumber_q"))
	for i=1 to len(BillStartNumber)
		if IsNumeric(mid(BillStartNumber,i,1)) then
			Sno=MID(BillStartNumber,1,i-1)
			Tno=MID(BillStartNumber,i,len(BillStartNumber))
			Tno2=MID(BillEndNumber,i,len(BillEndNumber))
			exit for
		end if
	next
end if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>警告單管理</title>
<link rel="stylesheet" href="../Common/css.txt" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style3 {font-size: 15px}
.style5 {
	font-size: 11px;
	color: #666666;
}
-->
</style></head>
<body>
<%
if Session("Msg")<>"" then
	 Response.write "<font  color='Red' size='2'>" & Session("Msg") & "</font>"
	 Session("Msg") = ""
end if	
%> 
<FORM NAME="myForm" ACTION="" METHOD="POST">    
<table width="100%" border="0">
  <tr>
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">警告單管理</span></td>
  </tr>
	<input type="hidden" name="sendType" value="<%=Request("sendType")%>">   
  <tr>
    <td height="26" bgcolor="#CCCCCC"><table width="100%" height="100%"  border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr>
        <td>
		<table border="0">
			<tr><td>
				<span class="font12"><font color="red">領單</font>日期</span>
			</td><td>
        		<%sStartDate=Request("fGetBillDate_q"):sEndDate=Request("tGetBillDate_q")%> 
				<input class="btn1" type='text' size='6' id='fGetBillDate_q' name='fGetBillDate_q' value='<%=sStartDate%>'>
				<input type="button" name="datestra" value="..." onclick="OpenWindow('fGetBillDate_q');">
				~
				<input class="btn1" type='text' size='6' id='tGetBillDate_q' name='tGetBillDate_q' value='<%=sEndDate%>'>
				<input type="button" name="datestrb" value="..." onclick="OpenWindow('tGetBillDate_q');">
			</td><td>
				<span class="font12">現　　任<br>領單單位</span>
			</td><td>
				<%=UnSelectUnitOption("UnitID_q","GetBillMemberID_q")%>
			</td><td>
				  <span class="font12">人員</span>
			</td><td>
				  <%=UnSelectMemberOption("UnitID_q","GetBillMemberID_q")%>
			</td></tr>
			<tr><td>
				<span class="font12"> 警告單號</span>
			</td><td>
				<input name="BillStartNumber_q" type="text" value="<%=Request("BillStartNumber_q")%>" size="8" maxlength="11" class="btn1">
				~      
				<input name="BillEndNumber_q" type="text" value="<%=Request("BillEndNumber_q")%>" size="8" maxlength="11" class="btn1">
			</td><td>
				狀態
			</td><td>
				<Select Name="Sys_Note">
					<option value="">全部</option>
					<option value="0"<%if trim(Request("Sys_Note"))="0" then response.write " selected"%>>領單</option>
					<option value="1"<%if trim(Request("Sys_Note"))="1" then response.write " selected"%>>入庫</option>
					<option value="2"<%if trim(Request("Sys_Note"))="2" then response.write " selected"%>>出庫</option>
				</select>
			</td><td>
				排序
			</td><td>
				<Select Name="Sys_Order">
					<option value="getbilldate desc"<%if trim(Request("Sys_Order"))="getbilldate desc" then response.write " selected"%>>領單日</option>
					<option value="billstartnumber"<%if trim(Request("Sys_Order"))="billstartnumber" then response.write " selected"%>>單號</option>
				</select>
			</td></tr>
			<tr><td>
				<span class="font12">離職(調職)<br>領單單位</span>
			</td><td>
				<%=UnLaverSelectUnitOption("UnitIDLaver_q","GetBillMemberIDLaver_q")%>
			</td><td>
				<span class="font12">人員</span>
			</td><td>
				<%=UnLaverSelectMemberOption("UnitIDLaver_q","GetBillMemberIDLaver_q")%>
			</td><td colspan=2>
				<input type="button" name="Submit" onclick="sendQry();" value="查詢" <%=ReturnPermission(CheckPermission(221,1))%>>
				<input type="button" name="Submit2" value="新增" onclick="openAddGetBill('WarningGetBillAdd.asp?tag=new','AddGetBill')" <%=ReturnPermission(CheckPermission(221,2))%>>

				<input type="button" name="Submit5" value="標示單漏號稽核" onclick="funAuditCar();">
			<!--<input type="button" name="Submit5" value="漏號稽核" onclick="sendLoss();">
				<input type="button" name="Submit5" value="漏號稽核" onclick="funAudit();">
				<input type="button" name="Submit5" value="漏號稽核"  onclick="openAddGetBill('BillAudit.asp','BillAudit');">--> 
			</td></tr>
			</table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
     <td height="26" bgcolor="#FFCC33"><span class="pagetitle">領單紀錄列表(<font size="3" color="red">*新增標示單時請連同100A一併填寫</font>)</span></td>
  </tr>
<%
qryType = 0
strWhere=""
SQL="select " & _
    "gb.GETBILLSN,ut.UNITNAME,md.ChName,gb.GetBillDate,gb.BillStartNumber,gb.BILLENDNUMBER, " & _
    "decode(gb.CounterfoiReturn,0,'使用中',1,'使用完畢') as CounterfoiReturnDesc,gb.BILLIN , gb.note, gb.RecordMemberID,gb.RecordDate " & _
    "from warninggetbillbase gb,memberdata md,unitinfo ut " & _
    "where gb.GetBillMemberID=md.MemberID and gb.RecordStateID <> -1 " & _
    "and md.UnitID=ut.UnitID "

	if Request("GetBillMemberIDLaver_q")<>"" then
		if Request("UnitIDLaver_q")<>"" then
			strWhere = strWhere & "and md.UnitID in('" & Request("UnitIDLaver_q") & "') "
			qryType = 1
		end if
		strWhere = strWhere & "and md.MemberID=" & Request("GetBillMemberIDLaver_q") & " "
		qryType = 1
	else
		if Request("UnitID_q")<>"" then
			strWhere = strWhere & "and md.UnitID in('" & Request("UnitID_q") & "') "
			qryType = 1
		end if
		if Request("GetBillMemberID_q")<>"" then
			strWhere= strWhere& "and md.MemberID=" & Request("GetBillMemberID_q") & " "
			qryType = 1
		end if
	end if

	  if (Request("BillStartNumber_q")<>"" and Request("BillEndNumber_q")<>"") then
	     'SQL = SQL & "and (gb.BILLSTARTNUMBER>='" & Request("BillStartNumber_q") & "' and gb.BILLENDNUMBER <= '" & Request("BillEndNumber_q") & "') "
	     strWhere = strWhere & "and (gb.GETBILLSN IN (select DISTINCT GETBILLSN " & _
	           "from warninggetbilldetail where billno >='"&Tno&"' and billno<= '"&Tno2&"'))"
	     qryType = 1
	  end if	  
	  
	  if (Request("fGetBillDate_q")<>"" and Request("tGetBillDate_q")<>"") then
				' after edit
	  	 fGetBillDate_q=gOutDT(Request("fGetBillDate_q"))&" 0:0:0"
	  	 tGetBillDate_q=gOutDT(Request("tGetBillDate_q")	)&" 23:59:59"										'CStr(fGetBillDate_q)     
	  	 strWhere = strWhere & "and getbilldate between TO_DATE('"&fGetBillDate_q&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&tGetBillDate_q&"','YYYY/MM/DD/HH24/MI/SS')"                           'CStr(tGetBillDate_q)
	     'source
	     'SQL = SQL & "and (to_char(gb.getbilldate,'yyyy/mm/dd') between " & 	funGetDate(goutDT(fGetBillDate_q),0) & " and " & funGetDate(goutDT(tGetBillDate_q),0) & ") "
	     qryType = 1
	  end if

	  if Request("Sys_Note")<>"" then
	     strWhere = strWhere & "and gb.BillIn="&Request("Sys_Note")
	     qryType = 1
	  end if	
	' smith 20080301 改為讀取 當天建檔的資料  避免 使用者 建檔資料的領單日期 不是今天的時候 就不會顯示在列表中
	' 有的人就會以為系統有問題.....
    if qryType = 0 and trim(request("DBState"))="" then
    	  strWhere = strWhere & "AND to_char(gb.getBillDate,'yyyy/mm/dd') = to_char(SYSDATE,'yyyy/mm/dd') "  
    end if 

'response.write SQL
set Rs=Server.CreateObject("ADODB.RecordSet")
rs.cursorlocation = 3

strCnt="select count(*) as cnt from ("&SQL&strWhere&")"
set Dbrs=conn.execute(strCnt)
DBsum=Dbrs("cnt")
Dbrs.close
If not ifnull(request("Sys_Order")) Then
strOrder=",gb."&request("Sys_Order")
End if
rs.open SQL&strWhere &  " order by md.UnitID"&strOrder,Conn,3,3

if Not rs.eof then
%>  
  <tr>
     <td bgcolor="#E0E0E0">
     	   <table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
              <tr bgcolor="#EBFBE3">
                <th width="12%" height="15" nowrap><span class="font12">領單單位</span></th>
                <th width="8%" height="15" nowrap><span class="font12">領單人員</span></th>
                <th width="5%" height="15" nowrap><span class="font12">領單日期</span></th>
                <th width="15%" nowrap><span class="font12">警告單起始碼~單截止碼</span></th>
              <!--  <th width="3%" height="15" nowrap><span class="font12">數量</span></th>-->
                <th width="10%" nowrap ><span class="font12">使用狀態</span></th>
                <th width="14%" nowrap><span class="font12">備註</span></th>
               <!-- <th width="5%" nowrap><span class="style3">漏號</span></th> -->
                <th width="20%" height="15" nowrap><span class="font12">操作</span></th>
              </tr>
	<%             
	if Trim(request("DB_Move"))="" then
		DBcnt=0
	else
		DBcnt=request("DB_Move")
	end if
	if Not rs.eof then rs.move Cint(DBcnt)
	for i=DBcnt+1 to DBcnt+10
		if rs.eof then exit for
	   billstartnumber = rs("billstartnumber")
	   billendnumber = rs("billendnumber")
     startTail = Mid(billstartnumber,5,len(BillStartNumber))
     endTail = Mid(billendnumber,5,len(BillStartNumber))
     numStart = FormatNumber(startTail,0)
     intStart = Int(numStart)
     numEnd = FormatNumber(endTail,0)
     intEnd = Int(numEnd)	
	   billAmount = intEnd - intStart + 1  
	   getbilldateTemp = gInitDT(rs("getbilldate")) '= Right("00" & Year(rs("getbilldate"))-1911, 3) & "-" & Right("0" & Month(rs("getbilldate")), 2) & "-" & Right("0" & Day(rs("getbilldate")), 2)
     detailPara = "WarningGetBillDetail.asp?GETBILLSN=" & rs("GETBILLSN") & "&getbilldate=" & getbilldateTemp & "&chname=" & rs("chname") & "&billstartnumber=" & billstartnumber & "&billendnumber=" & billendnumber & "&qryType=1"

				response.write "<tr bgcolor='#FFFFFF' align='center' "
				lightbarstyle 0 
				response.write ">"
	%>          <td align="right"><span class="font11"><%=rs("unitname")%></span></td>
                <td align="right"><span class="font11"><%=rs("chname")%></span></td>
                <td align="right"><span class="font11"><%=getbilldateTemp%></span></td>
                <td align="right"><span class="font11"><%=billstartnumber%> ~ <%=billendnumber%></span></td>
               <!--<td align="right"><span class="font11"><%=billAmount%></span></td>-->
                <td align="right"><span class="font11"><%=rs("CounterfoiReturnDesc")%></span></td>
                <td ><span class="font11"><%
					if rs("BILLIN")=1 then
						response.write "入庫 .  "&rs("note")
					elseif rs("BILLIN")=2 then
						response.write "出庫 .  "&rs("note")
					else
						response.write rs("note")
					end if%></span></td>
                <!-- <td >　</td>  -->
                <td nowrap>
                    <input type="button" name="Submit43" value="檢視明細" onclick="openAddGetBill('<%=detailPara%>','GetBillDetail');">
				                    
				
						
						<input type="button" name="Submit433" value="轉為他人使用" onclick="openAddGetBill('WarningGetBillChange.asp?tag=change&GETBILLSN=<%=rs("GETBILLSN")%>','ChgGetBill')" >
		
                    <input type="button" name="Submit433" value="修改" onclick="openAddGetBill('WarningGetBillUpdate.asp?tag=upd&GETBILLSN=<%=rs("GETBILLSN")%>','UpdGetBill')" >
                  <input type="button" name="Submit3" value="刪除" onclick="delGetBill('WarningGetBill_mdy.asp?tag=del&GetBillSN=<%=rs("GetBillSN")%>');" >
				
                </td>
              </tr>
	<%              
		rs.movenext
	next%>
         </table>
     </td>
  </tr>
	<tr>
		<td bgcolor="#FFDD77" align="center">
			<img src='../Image/PREVPAGE.GIF' border=0 onclick="funDbMove(-10);">
			<span class="style2"> <%=Cint(DBcnt)/10+1&"/"&fix(Cint(DBsum)/10+0.9)%></span>
			<img src='../Image/NextPage.gif' border=0 onclick="funDbMove(10);">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="button" name="SaveAs" value="轉換成Excel" onclick="funReporExcel();"></td>              
	</tr>  

<% else %>    
  <tr>
  	 <td align="center" >        
	      <center><font  color="Red" size="2">              
	<%              
	Response.Write "目前查無任何資料 ..."              
	%>              
	      </font></center>	      <p>　</p>   	    <p>　</p></td>
  </tr>             
<%              
end if              
rs.close              
set rs = nothing              
%>   
</table>
  <input type="Hidden" name="DBState" value="<%=trim(request("DBState"))%>">
  <input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
	<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
	<input type="Hidden" name="Sys_strWhere" value="<%=strWhere%>">
</FORM>
</body>


</html>
<script language=javascript src='../js/date.js'></script>
<script language=javascript src='../js/form.js'></script>
<script language=javascript src='../js/GetBill.js'></script>
<Script language="JavaScript">
<!--
function funDbMove(MoveCnt){
	if (eval(MoveCnt)>0){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}else{
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}
}
function funAudit(){
	UrlStr="BillBaseAudit.asp";
	newWin(UrlStr,"Audit",900,550,50,10,"yes","yes","yes","no");
}
function funAuditCar(){
	UrlStr="BillBaseAuditCar.asp";
	newWin(UrlStr,"Audit",950,750,50,10,"yes","yes","yes","no");
}
function qryCheck()
{
	var chkResult ;
	var rtnChkBillNum ;

  if ((document.myForm.fGetBillDate_q.value=="") && (document.myForm.tGetBillDate_q.value!="")){
  	alert('請輸入領單起始日期!!');
  	document.myForm.fGetBillDate_q.focus();
  	return false;
  }else if ((document.myForm.fGetBillDate_q.value!="") && (document.myForm.tGetBillDate_q.value=="")){
  	alert('請輸入領單結束日期!!');
  	document.myForm.tGetBillDate_q.focus();
  	return false;    	
  }else if ((document.myForm.fGetBillDate_q.value!="") && (document.myForm.tGetBillDate_q.value!="")){
  	if (document.myForm.fGetBillDate_q.value > document.myForm.tGetBillDate_q.value){
  		alert('領單起始日期不得大於領單結束日期');
  		return false; 
  	}
  }
  
  if ((document.myForm.BillStartNumber_q.value=="") && (document.myForm.BillEndNumber_q.value!="")){
  	alert('請輸入警告單起始碼!!');
  	document.myForm.fGetBillDate_q.focus();
  	return false;
  }  
  if ((document.myForm.BillStartNumber_q.value!="") && (document.myForm.BillEndNumber_q.value=="")){
  	alert('請輸入警告單截止碼!!');
  	document.myForm.fGetBillDate_q.focus();
  	return false;
  }

}

function funReporExcel(){
	UrlStr="saveExcel.asp";
	myForm.action=UrlStr;
	myForm.target="HuaLien";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}

function sendQry(){
	var rtn;
	rtn = qryCheck();
	if (rtn!=false){

     var form_A= document.forms[0];
     form_A.sendType.value="1";
     form_A.action = "WarningGetBill.asp";
	 form_A.DBState.value="1";
     form_A.submit();		
	}
}

function sendLoss(){
	var tmpFlag = "1";
	var form_A= document.forms[0];
	if (form_A.sendType.value=="1"){
		 openAddGetBill('BillAudit.asp','GetBillDetail');
		 form_A.sendType.value="0";
	}else{
		 alert('請先重新進行查詢作業再執行[漏號稽核]!!');
		 form_A.sendType.value="0";
	}
				
}

function delGetBill(param){
	 var rtn;
	 rtn = window.confirm("您確定要刪除此筆資料嗎?");
	 if (rtn!=false){
     openAddGetBill(param,'DelGetBill');	
   }
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	win.focus();
	return win;
}
<%response.write "UnitMan('UnitID_q','GetBillMemberID_q','"&request("GetBillMemberID_q")&"');"%>
<%response.write "UnitLaverMan('UnitIDLaver_q','GetBillMemberIDLaver_q','"&request("GetBillMemberIDLaver_q")&"');"%>
-->
</Script>
<!-- #include file="../Common/ClearObject.asp" -->