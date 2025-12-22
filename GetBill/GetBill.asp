
<!-- #include file="../Common/DbUtil.asp"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!-- #include file="../Common/Util.asp"-->
<!-- #include file="../Common/bannernodata.asp"-->

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<%
strCity="select value from Apconfigure where id=31"
set rsCity=conn.execute(strCity)
sys_City=trim(rsCity("value"))
rsCity.close

sqlUnit = "Select UnitName , UnitID from UnitInfo"
set RsUnit=Server.CreateObject("ADODB.RecordSet")
RsUnit.open sqlUnit,Conn,3,3

sql = "Select ChName , MemberID from MemberData where UnitID in('" & Request("UnitID_q") & "')"
set RsUpd1=Server.CreateObject("ADODB.RecordSet")
RsUpd1.open sql,Conn,3,3

if Request("BillStartNumber_q")<>"" and Request("BillEndNumber_q")<>"" then
	BillStartNumber=trim(Ucase(Request("BillStartNumber_q")))
	BillEndNumber=trim(Ucase(Request("BillEndNumber_q")))
	for i=len(BillStartNumber) to 1 step -1
		if not IsNumeric(mid(BillStartNumber,i,1)) then
			Sno=MID(BillStartNumber,1,i)
			Tno=MID(BillStartNumber,i+1,len(BillStartNumber))
			Tno2=MID(BillEndNumber,i+1,len(BillEndNumber))
			exit for
		end if
	next
end if

If not ifnull(Request("DB_CounterfoiReturn")) Then
	sql = "Update GetBillBase Set COUNTERFOIRETURN=" & trim(Request("DB_CounterfoiReturn")) & ",BillReturnDate="&funGetDate(now,1)&" Where GETBILLSN=" & trim(Request("DB_GetBillSN"))
	Conn.Execute(sql)
End if

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>領單管理</title>
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
    <td height="27" bgcolor="#FFCC33"><span class="pagetitle">領單管理</span></td>
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
				<span class="font12"> 舉發單號</span>
			</td><td>
				<input name="BillStartNumber_q" type="text" value="<%=Request("BillStartNumber_q")%>" size="8" maxlength="9" class="btn1">
				~      
				<input name="BillEndNumber_q" type="text" value="<%=Request("BillEndNumber_q")%>" size="8" maxlength="9" class="btn1">
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
				使用狀況
			</td><td>
				<Select Name="Sys_CounterfoiReturn">
					<option value="">全部</option>
					<option value="0"<%if trim(Request("Sys_CounterfoiReturn"))="0" then response.write " selected"%>>使用中</option>
					<option value="1"<%if trim(Request("Sys_CounterfoiReturn"))="1" then response.write " selected"%>>使用完畢</option>
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
			</td><td>
				排序
			</td>
			<td>
				<Select Name="Sys_Order">
					<option value="getbilldate desc"<%if trim(Request("Sys_Order"))="getbilldate desc" then response.write " selected"%>>領單日</option>
					<option value="billstartnumber"<%if trim(Request("Sys_Order"))="billstartnumber" then response.write " selected"%>>單號</option>
				</select>
			</td></tr>
			<tr>
				<td colspan="6" align="right">
					<input type="button" name="Submit" onclick="sendQry();" value="查詢">
					<input type="button" name="Submit2" value="新增" onclick="openAddGetBill('GetBillAdd.asp?tag=new','AddGetBill');">
					<!--<input type="button" name="Submit5" value="漏號稽核" onclick="sendLoss();"> -->
					
					<!--<input type="button" name="Submit5" value="漏號稽核"  onclick="openAddGetBill('BillAudit.asp','BillAudit');">--> 
				
					<input type="button" name="Submit5" value="漏號稽核" onclick="funAudit();">
					<%If sys_City="台南市" then%>
						<input type="button" name="Submit5" value="庫存檢視" onclick="funViewBill();">
					<%end if%>
				<td>
			</tr>
			</table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
     <td height="26" bgcolor="#FFCC33">
		<span class="pagetitle">領單紀錄列表</span>
	 </td>
  </tr>
<%
qryType = 0
SQL="select " & _
    "gb.GETBILLSN,ut.UNITNAME,md.ChName,md.LoginID,gb.GetBillDate,gb.Billreturndate,gb.BillStartNumber,gb.BILLENDNUMBER, " & _
    "decode(gb.CounterfoiReturn,0,'使用中',1,'使用完畢') as CounterfoiReturnDesc,gb.BILLIN , gb.note, gb.RecordMemberID,gb.RecordDate " & _
    "from getbillbase gb,memberdata md,unitinfo ut " & _
    "where gb.GetBillMemberID=md.MemberID and gb.RecordStateID <> -1 " & _
    "and md.UnitID=ut.UnitID "

	if Request("GetBillMemberIDLaver_q")<>"" then
		if Request("UnitIDLaver_q")<>"" then
			Sys_strWhere = Sys_strWhere & "and md.UnitID in('" & Request("UnitIDLaver_q") & "') "
			qryType = 1
		end if
		Sys_strWhere = Sys_strWhere & "and md.MemberID=" & Request("GetBillMemberIDLaver_q") & " "
		qryType = 1
	else
		if Request("UnitID_q")<>"" then
			if trim(Session("UnitLevelID"))="1" then

				strSQL="select UnitLevelID from UnitInfo where UnitID='"&Request("UnitID_q")&"'"

				set Dbrs=conn.execute(strSQL)
					Db_UnitLevelID=Dbrs("UnitLevelID")
				Dbrs.close

				If Db_UnitLevelID > 2 Then
					Sys_strWhere = Sys_strWhere & "and md.UnitID in('" & Request("UnitID_q") & "') "

				else
					Sys_strWhere = Sys_strWhere & "and md.UnitID in(select UnitID from UnitInfo where UnitTypeID in(select UnitTypeID from UnitInfo where UnitID='" & Request("UnitID_q") & "')) "

				End if 
				
			else
				
				Sys_strWhere = Sys_strWhere & "and md.UnitID in('" & Request("UnitID_q") & "') "
			End if 

			qryType = 1
		end if
		if Request("GetBillMemberID_q")<>"" then
			Sys_strWhere = Sys_strWhere & "and md.MemberID=" & Request("GetBillMemberID_q") & " "
			qryType = 1
		end if
	end if
	if Request("Sys_CounterfoiReturn")<>"" then
		Sys_strWhere = Sys_strWhere & "and gb.CounterfoiReturn=" & Request("Sys_CounterfoiReturn") & " "
	end if

	  if (Request("BillStartNumber_q")<>"" and Request("BillEndNumber_q")<>"") then
	     'SQL = SQL & "and (gb.BILLSTARTNUMBER>='" & Request("BillStartNumber_q") & "' and gb.BILLENDNUMBER <= '" & Request("BillEndNumber_q") & "') "
	     Sys_strWhere = Sys_strWhere & "and (gb.GETBILLSN IN (select DISTINCT GETBILLSN " & _
	           "from getbilldetail where SUBSTR(BillNo,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(BillNo,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"') OR (SUBSTR(gb.BILLSTARTNUMBER,1,"&len(Sno)&")='"&Sno&"' and SUBSTR(gb.BILLSTARTNUMBER,"&len(Sno)+1&") between '"&Tno&"' and '"&Tno2&"'))"
	     qryType = 1
	  end if	  
	  if (Request("fGetBillDate_q")<>"" and Request("tGetBillDate_q")<>"") then
				' after edit
	  	 fGetBillDate_q=gOutDT(Request("fGetBillDate_q"))&" 0:0:0"
	  	 tGetBillDate_q=gOutDT(Request("tGetBillDate_q")	)&" 23:59:59"										'CStr(fGetBillDate_q)     
	  	 Sys_strWhere = Sys_strWhere & "and getbilldate between TO_DATE('"&fGetBillDate_q&"','YYYY/MM/DD/HH24/MI/SS') and TO_DATE('"&tGetBillDate_q&"','YYYY/MM/DD/HH24/MI/SS')"                           'CStr(tGetBillDate_q)
	     'source
	     'SQL = SQL & "and (to_char(gb.getbilldate,'yyyy/mm/dd') between " & 	funGetDate(goutDT(fGetBillDate_q),0) & " and " & funGetDate(goutDT(tGetBillDate_q),0) & ") "
	     qryType = 1
	  end if

	  if Request("Sys_Note")<>"" then
	     Sys_strWhere = Sys_strWhere & "and gb.BillIn="&Request("Sys_Note")
	     qryType = 1
	  end if	
	' smith 20080301 改為讀取 當天建檔的資料  避免 使用者 建檔資料的領單日期 不是今天的時候 就不會顯示在列表中
	' 有的人就會以為系統有問題.....
    if qryType = 0 and trim(request("DBState"))="" then
    	  Sys_strWhere = Sys_strWhere & "AND to_char(gb.getBillDate,'yyyy/mm/dd') = to_char(SYSDATE,'yyyy/mm/dd') "  
    end if 

'Session("ExcelSql") = SQL

SQL=SQL&Sys_strWhere

set Rs=Server.CreateObject("ADODB.RecordSet")
rs.cursorlocation = 3

strCnt="select count(*) as cnt from ("&SQL&")"
set Dbrs=conn.execute(strCnt)
DBsum=Dbrs("cnt")
Dbrs.close
If not ifnull(request("Sys_Order")) Then
strOrder=",gb."&request("Sys_Order")
End if

'response.write SQL &  " order by md.UnitID"&strOrder
'response.end
rs.open SQL &  " order by md.UnitID"&strOrder,Conn,3,3

if Not rs.eof then
%>  
  <tr>
     <td bgcolor="#E0E0E0">
     	   <table width="100%" height="100%" border="0" cellpadding="4" cellspacing="1">
              <tr bgcolor="#EBFBE3">
                <th width="12%" height="15" nowrap><span class="font12">領單單位</span></th>
                <th width="8%" height="15" nowrap><span class="font12">領單人員</span></th>
                <th width="5%" height="15" nowrap><span class="font12">領單日</span></th>
                <th width="5%" height="15" nowrap><span class="font12">繳回日</span></th>
                <th width="15%" nowrap><span class="font12">舉發單起始碼~單截止碼</span></th>
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
     startTail = Mid(billstartnumber,4,6)
     endTail = Mid(billendnumber,4,6)
     numStart = FormatNumber(startTail,0)
     intStart = Int(numStart)
     numEnd = FormatNumber(endTail,0)
     intEnd = Int(numEnd)	
	   billAmount = intEnd - intStart + 1  
	   getbilldateTemp = gInitDT(rs("getbilldate")) '= Right("00" & Year(rs("getbilldate"))-1911, 3) & "-" & Right("0" & Month(rs("getbilldate")), 2) & "-" & Right("0" & Day(rs("getbilldate")), 2)
     detailPara = "GetBillDetail.asp?GETBILLSN=" & rs("GETBILLSN") & "&getbilldate=" & getbilldateTemp & "&chname=" & rs("chname") & "&billstartnumber=" & billstartnumber & "&billendnumber=" & billendnumber & "&qryType=1"
		billreturndatetemp = gInitDT(rs("billreturndate")) '= Right("00" & Year(rs("billreturndate"))-1911, 3) & "-" & Right("0" & Month(rs("billreturndate")), 2) & "-" & Right("0" & Day(rs("billreturndate")), 2)
				response.write "<tr bgcolor='#FFFFFF' align='center' "
				lightbarstyle 0 
				response.write ">"
	%>          <td align="right" nowrap><span class="font11"><%=rs("unitname")%></span></td>
                <td align="right" nowrap><span class="font11"><%=rs("chname")%></span></td>
                <td align="right" nowrap><span class="font11"><%=getbilldateTemp%></span></td>
                <td align="right" nowrap><span class="font11"><%=billreturndatetemp%></span></td>
                <td align="right" nowrap><span class="font11"><%=billstartnumber%> ~ <%=billendnumber%></span></td>
               <!--<td align="right"><span class="font11"><%=billAmount%></span></td>-->
                <td align="right" id="over_<%=rs("GETBILLSN")%>">
					<span class="font11"><%
					response.write rs("CounterfoiReturnDesc")
					 	
					strCity="select value from Apconfigure where id=31"
					set rsCity=conn.execute(strCity)
						sys_City=trim(rsCity("value"))
					rsCity.close
					set rsCity=nothing

					if sys_City="高雄縣" and trim(rs("CounterfoiReturnDesc"))="使用完畢" then
						response.write "繳回"
					end if
					%></span>
				</td>

                <td ><span class="font11"><%
					if rs("BILLIN")=1 then
						response.write "入庫 .  "&rs("note")
					elseif rs("BILLIN")=2 then
						response.write "出庫 .  "&rs("note")
					else
						response.write rs("note")
					end if%></span>
				</td>
                <!-- <td >　</td>  -->
                <td nowrap>
                    <input type="button" name="Submit43" value="檢視明細" onclick="openAddGetBill('<%=detailPara%>','GetBillDetail');">
				    <input type="button" name="Submit433" value="轉為他人使用" onclick="openAddGetBill('GetBillChange.asp?tag=change&GETBILLSN=<%=rs("GETBILLSN")%>','ChgGetBill')" >
<!-- 
					<%If rs("BILLIN")=0 and trim(rs("CounterfoiReturnDesc"))="使用完畢" Then%>
						<span id="btn_<%=rs("GETBILLSN")%>">
							<input type="button" name="Submit433" value="使用中" onclick="funBillOver('<%=rs("GETBILLSN")%>','0');" >
						</span>
					<%elseif rs("BILLIN")=0 then%>
						<span id="btn_<%=rs("GETBILLSN")%>">
							<input type="button" name="Submit433" value="使用完畢" onclick="funBillOver('<%=rs("GETBILLSN")%>','1');" >
						</span>
					<%End if%>
				 -->
                    <input type="button" name="Submit433" value="修改" onclick="newWin('GetBillUpdate.asp?tag=upd&GETBILLSN=<%=rs("GETBILLSN")%>','UpdGetBill',950,750,50,10,'yes','yes','yes','no');">
					<input type="button" name="Submit3" value="刪除" onclick="delGetBill('GetBill_mdy.asp?tag=del&GetBillSN=<%=rs("GetBillSN")%>');" >
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
			<input type="button" name="SaveAs" value="轉換成Excel" onclick="funExportExcel();"></td>              
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
	<input type="Hidden" name="DB_CounterfoiReturn" value="">
	<input type="Hidden" name="DB_GetBillSN" value="">
	<input type="Hidden" name="Sys_strWhere" value="<%=Sys_strWhere%>">
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
function funExportExcel(){

		UrlStr="saveExcel.asp";
		myForm.action=UrlStr;
		myForm.target="saveExcel";
		myForm.submit();
		myForm.action="";
		myForm.target="";
}
function funAudit(){
	UrlStr="BillBaseAudit.asp";
	newWin(UrlStr,"Audit",950,750,50,10,"yes","yes","yes","no");
}
function funAuditCar(){
	UrlStr="BillBaseAuditCar.asp";
	newWin(UrlStr,"Audit",950,750,50,10,"yes","yes","yes","no");
}
function funViewBill(){
	UrlStr="TotalGetBillbase.asp";
	newWin(UrlStr,"Audit",550,450,150,10,"yes","yes","yes","no");
}
function funBillOver(sn,tag){
	runServerScript("chkOverBill.asp?GetBillSN="+sn+"&tag="+tag);
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
  	alert('請輸入舉發單起始碼!!');
  	document.myForm.fGetBillDate_q.focus();
  	return false;
  }  
  if ((document.myForm.BillStartNumber_q.value!="") && (document.myForm.BillEndNumber_q.value=="")){
  	alert('請輸入舉發單截止碼!!');
  	document.myForm.fGetBillDate_q.focus();
  	return false;
  }

  if (document.myForm.BillStartNumber_q.value!=""){
     chkResult = chkBillNumber(document.all.BillStartNumber_q,"[舉發單起始碼] 格式錯誤!!"); 
     if (chkResult != "Y") return false;
  }
  if (document.myForm.BillEndNumber_q.value!=""){
     chkResult = chkBillNumber(document.all.BillEndNumber_q,"[舉發單截止碼] 格式錯誤!!"); 
     if (chkResult != "Y") return false;    
  }  
   
  if (document.myForm.BillStartNumber_q.value!="" && document.myForm.BillEndNumber_q.value!=""){  
     rtnChkBillNum = ValidateBillNumbers (document.myForm.BillStartNumber_q,document.myForm.BillEndNumber_q,'N');
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
	rtn = qryCheck();
	if (rtn!=false){
     var form_A= document.forms[0];
     form_A.sendType.value="1";
     form_A.action = "GetBill.asp";
	 form_A.DBState.value="1";
     form_A.submit();		
	}
}

function funCounter(TypeID,BillSN){
	myForm.DB_CounterfoiReturn.value=TypeID;
	myForm.DB_GetBillSN.value=BillSN;
	sendQry();
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