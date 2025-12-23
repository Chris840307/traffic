<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Login_Check.asp"-->
<%
AuthorityCheck(233)
strSQL="select distinct a.SN,a.BillNo,a.CarNo,a.Rule1,a.IllegalSpeed,a.BillMem1,c.Content,c.value,d.UnitName,e.Chname from BillBase a, BilLBaseDciReturn b, CarSpeed c,UnitInfo d,MemberData e where a.BillNo=b.BillNo and a.CarNo=b.CarNo and b.DciReturnCarType=c.ID and a.RecordMemberID=e.MemberID and a.BillUnitID=d.UnitID and a.RecordStateID=0 and b.ExChangeTypeID='W' and b.Status='Y' and a.IllegalSpeed>c.value and a.SN in(select distinct a.BillSN from (select * from DCILog"&request("strDCISQL")&") a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+)"&trim(request("TempSQL"))&")"
set rsfound=conn.execute(strSQL)

strCnt="select count(*) as cnt from BillBase a, BilLBaseDciReturn b, CarSpeed c,UnitInfo d,MemberData e where a.BillNo=b.BillNo and a.CarNo=b.CarNo and b.DciReturnCarType=c.ID and a.RecordMemberID=e.MemberID and a.BillUnitID=d.UnitID and a.RecordStateID=0 and b.ExChangeTypeID='W' and b.Status='Y' and a.IllegalSpeed>c.value and a.SN in(select distinct a.BillSN from (select * from DCILog"&request("strDCISQL")&") a,MemberData b,(select * from DCIcode where TypeID=2) c,DCIReturnStatus d,(select distinct DCIACTIONID,DCIACTIONNAME from DCIReturnStatus) e,(select * from DciReturnStatus where DciActionID='WE') f,(select * from DciReturnStatus where DciActionID='WE') g where a.ExchangeTypeID=d.DCIActionID(+) and a.DCIReturnStatusID=d.DCIReturn(+) and a.RecordMemberID=b.MemberID(+) and a.BillTypeID=c.ID(+) and a.ExchangeTypeID=e.DCIActionID(+) and a.DCIERRORCARDATA=f.DciReturn(+) and a.DCIERRORIDDATA=g.DciReturn(+)"&trim(request("TempSQL"))&")"
set Dbrs=conn.execute(strCnt)
DBsum=CDbl(Dbrs("cnt"))
Dbrs.close

%>
<HTML>
<HEAD>
<TITLE>稽核特殊車種車速系統</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</HEAD>

<BODY>
<form name=myForm method="post">
<table width="100%" border="0">
<tr>
	<td bgcolor="#FFCC33" height="33">稽核特殊車種車速列表<img src="space.gif" width="15" height="8"><strong>( 查詢 <%=DBsum%> 筆紀錄 )</strong></td>
</tr>
<tr>
	<td bgcolor="#E0E0E0">
		<Div style="overflow:auto;width:100%;height:360px;background:#FFFFFF">
			<table width="100%" bgcolor="#E0E0E0" border="0" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<td>單號</td>
					<td>車號</td>
					<td>車種</td>
					<td>違規法條</td>
					<td>實際車速</td>					
					<td>設定車速</td>
					<td>舉發單位</td>
					<td>舉發員警</td>
					<td>建檔人</td>
					<td>詳細資料</td>
				</tr><%
				if Trim(request("DB_Move"))="" then
					DBcnt=0
				else
					DBcnt=request("DB_Move")
				end if
				if Not rsfound.eof then rsfound.move DBcnt
				for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
					if rsfound.eof then exit for
					response.write "<tr bgcolor='#FFFFFF' align=""right"""
					lightbarstyle 0 
					response.write ">"
					response.write "<td>"&rsfound("BillNo")&"</td>"
					response.write "<td>"&rsfound("CarNo")&"</td>"
					response.write "<td>"&rsfound("Content")&"</td>"
					response.write "<td>"&rsfound("Rule1")&"</td>"
					response.write "<td>"&rsfound("IllegalSpeed")&"</td>"
					response.write "<td>"&rsfound("Value")&"</td>"
					response.write "<td>"&rsfound("UnitName")&"</td>"
					response.write "<td>"&rsfound("BillMem1")&"</td>"
					response.write "<td>"&rsfound("ChName")&"</td>"
					response.write "<td><input type=""button"" name=""Update"" value=""詳細資料"" onclick=""funDataDetail('"&rsfound("SN")&"');""></td>"
					response.write "</tr>"
					rsfound.movenext
				next%>
			</table>
		</Div>
	</td>
</tr>
<tr>
	<td height="30" colspan="10" bgcolor="#FFDD77" align="center">
		<a href="file:///.."></a>
		<input type="button" name="MoveFirst" value="第一頁" onclick="funDbMove(0);">
		<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
		<span class="style2"> <%=fix(CDbl(DBcnt)/(10)+1)&"/"&fix(CDbl(DBsum)/(10)+0.9)%></span>
		<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
		<input type="button" name="MoveDown" value="最後一頁" onclick="funDbMove(999);">
		<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
	</td>
</tr>
</table>
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="TempSQL" value="<%=trim(request("TempSQL"))%>">
</form>
</BODY>
</HTML>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function funDbMove(MoveCnt){
	if (eval(MoveCnt)==0){
		myForm.DB_Move.value="";
		myForm.submit();
	}else if (eval(MoveCnt)==10){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}else if(eval(MoveCnt)==-10){
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt;
			myForm.submit();
		}
	}else if(eval(MoveCnt)==999){
		if (eval(myForm.DB_Cnt.value)%(10)==0){
			myForm.DB_Move.value=(Math.floor(eval(myForm.DB_Cnt.value)/(10))-1)*(10);
		}else{
			myForm.DB_Move.value=Math.floor(eval(myForm.DB_Cnt.value)/(10))*(10);
		}
		myForm.submit();
	}
}
function funDataDetail(SN){
	UrlStr="ViewBillBaseData_Car.asp?BillSn="+SN;
	newWin(UrlStr,"DetailCar",900,550,50,10,"yes","yes","yes","no");
}
function funchgExecel(){
	myForm.action="CarSpeed_Execel.asp";
	myForm.target="DetailCar";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
	winopen=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=yes");
	winopen.focus();
	return win;
}
</script>