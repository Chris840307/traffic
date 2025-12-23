<!--#include virtual="/traffic/Common/Login_Check.asp"--> 
<!--#include virtual="/traffic/Common/db.ini"-->
<!--#include virtual="/traffic/Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<!--#include file="sqlDCIExchangeData.asp"-->
<title>舉發單多筆刪除</title>
<% Server.ScriptTimeout = 800 %>
<%
'檢查是否可進入本系統
'AuthorityCheck(237)

DelMemID=trim(Session("User_ID"))
theBatchNumber=trim(request("BatchNumber"))

if trim(request("kinds"))="DB_BillDel" then
		CaseIN7DayList=""	'列出超過七天的
		BillDelFlag="N"
		strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theBatchTime=(year(now)-1911)&"E"&trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing
		'取批號
		strSN="select DCILOGBATCHNUMBER.nextval as SN from Dual"
		set rsSN=conn.execute(strSN)
		if not rsSN.eof then
			theBatchTime=(year(now)-1911)&"E"&trim(rsSN("SN"))
		end if
		rsSN.close
		set rsSN=nothing

		BillSnArray=split(trim(request("BillNoSelectlist")),",")
		for i=0 to ubound(BillSnArray)
			BillDataTmp=split(trim(BillSnArray(i)),"-")
			NoteTmp=replace(trim(request("DelNote")),",","")
			if trim(BillDataTmp(0))="0" then	'攔停、逕舉
				strBill=""
				CaseInDate=""

				dBillNo=""
				dBillTypeID=""
				dCarNo=""
				dBillUnitID=""
				dBillStatus=""
				dRecordDate=""
				dRecordMemberID=""
				CaseInStatus=0
				strCType="select a.DciCaseInDate,b.BillNo,b.CarNo,b.BillTypeID,b.BillStatus,b.BillUnitID" &_
					",b.RecordDate,b.RecordMemberID from BillBaseDCIReturn a,BillBase b" &_
					" where a.BillNo=b.BillNo and a.CarNo=B.CarNo and b.Sn="&trim(BillDataTmp(1))&_
					" and a.ExchangeTypeID='W' and a.Status in ('Y','S','n')"
				set rsCType=conn.execute(strCType)
				if not rsCType.eof then
					'response.write trim(rsCType("DciCaseInDate"))
					CaseInDate=gOutDT(trim(rsCType("DciCaseInDate")))
					CaseInStatus=1	'入案成功
					dBillNo=trim(rsCType("BillNo"))
					dBillTypeID=trim(rsCType("BillTypeID"))
					dCarNo=trim(rsCType("CarNo"))
					dBillUnitID=trim(rsCType("BillUnitID"))
					dBillStatus=trim(rsCType("BillStatus"))
					dRecordDate=trim(rsCType("RecordDate"))
					dRecordMemberID=trim(rsCType("RecordMemberID"))
				end if
				rsCType.close
				set rsCType=nothing
				'計算入案幾天
				CountCaseIN=0
				if CaseInDate<>"" then
					CountCaseIN=dateDiff("d",CaseInDate,now)
				end if

				if CountCaseIN>7 or CaseInStatus=0 then	'入案超過七天或沒有入案成功
					'response.write CaseInDate&"---"&CountCaseIN&"..."&billnotest&"no log"
					strBData="select b.BillNo,b.CarNo,b.BillTypeID,b.BillStatus,b.BillUnitID" &_
						",b.RecordDate,b.RecordMemberID from BillBase b" &_
						" where b.Sn="&trim(BillDataTmp(1))
					set rsBData=conn.execute(strBData)
					if not rsBData.eof then
						dBillNo=trim(rsBData("BillNo"))
						dBillTypeID=trim(rsBData("BillTypeID"))
						dCarNo=trim(rsBData("CarNo"))
						dBillUnitID=trim(rsBData("BillUnitID"))
						dBillStatus=trim(rsBData("BillStatus"))
						dRecordDate=trim(rsBData("RecordDate"))
						dRecordMemberID=trim(rsBData("RecordMemberID"))
					end if
					rsBData.close
					set rsBData=nothing
					CaseInStatus=0	'不寫DCILOG
					BillDelFlag=funcDCIdel(conn,trim(BillDataTmp(1)),dBillNo,dBillTypeID,dCarNo,dBillUnitID,dBillStatus,dRecordDate,dRecordMemberID,trim(request("DelReason")),NoteTmp,CaseInStatus,theBatchTime)

					if CountCaseIN>7 then
						if CaseIN7DayList="" then
							CaseIN7DayList="下列舉發單入案已超過七天，需發文請監理站刪除：\n"
							CaseIN7DayList=CaseIN7DayList&dBillNo&" \n"
						else
							CaseIN7DayList=CaseIN7DayList&dBillNo&" \n"
						end if
					end if
				else	
					'response.write CaseInDate&"---"&CountCaseIN&"..."&billnotest&" log"
					BillDelFlag=funcDCIdel(conn,trim(BillDataTmp(1)),dBillNo,dBillTypeID,dCarNo,dBillUnitID,dBillStatus,dRecordDate,dRecordMemberID,trim(request("DelReason")),NoteTmp,CaseInStatus,theBatchTime)
				end if

				'寫入LOG
				DeleteReason=""
				if trim(request("DelReason"))<>"" then
					strRea="select ID,Content from DCICode where TypeID=3 and ID='"&trim(request("DelReason"))&"'"
					set rsRea=conn.execute(strRea)
					if not rsRea.eof then
						DeleteReason=trim(rsRea("Content"))
					end if
					rsRea.close
					set rsRea=nothing
				else
					DeleteReason="無"
				end if
				ConnExecute "舉發單刪除 單號:"&dBillNo&" 車號:"&dCarNo&" 原因:"&DeleteReason&","&trim(NoteTmp)&","&CaseInStatus,352
			else	'行人攤販
				strDel="delete from PasserDeleteReason where PasserSN="&trim(BillDataTmp(1))
				conn.execute strDel
				
				strUpd="update PasserBase set BillStatus='6',RecordDate=sysdate,RecordStateID=-1,DelMemberID="&DelMemID&" where SN="&trim(BillDataTmp(1))
				conn.execute strUpd

				strIns="Insert into PasserDeleteReason(PasserSN,DelDate,DelReason,Note)" &_
					" values("&trim(BillDataTmp(1))&",sysdate,'"&trim(request("DelReason"))&"','"&trim(NoteTmp)&"')"
				conn.execute strIns
				
				'寫入LOG
				DeleteReason=""
				strRea="select ID,Content from DCICode where TypeID=3 and ID='"&trim(request("DelReason"))&"'"
				set rsRea=conn.execute(strRea)
				if not rsRea.eof then
					DeleteReason=trim(rsRea("Content"))
				end if
				rsRea.close
				set rsRea=nothing
				'抓單號
				theBillNO=""
				strbillno="select BillNo from PasserBase where SN="&trim(BillDataTmp(1))
				set rsBillno=conn.execute(strbillno)
				if not rsBillno.eof then
					theBillNO=trim(rsBillno("BillNo"))
				end if
				rsBillno.close
				set rsBillno=nothing
				ConnExecute "舉發單刪除 單號:"&theBillNO&" 原因:"&DeleteReason&","&trim(NoteTmp),352
			end if
		next
%>
		<script language="JavaScript">
			alert("刪除完成！！ \n<%=CaseIN7DayList%>");
		</script>
<%
end if
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="4"><strong>舉發單多筆刪除</strong></td>
			</tr>
			<tr>
				<td>
					&nbsp;單號 <input type="text" name="DelBillNo" value="" size="12" maxlength="9">
					<input type="Button" name="DelButton1A" value="加入" onclick="addDelBillNo();">
					<br><br>
					&nbsp;&nbsp;<select name="DelBillList" multiple size="8"></select>
					<input type="Button" name="cDelButton1A" value="取消刪除" onclick="cancelDelBillNo();">
					<br>
					&nbsp;用Ctrl或Shift可以多選
				</td>
			</tr>
			<tr>
				<td>
					刪除原因 
					<select name="DelReason">
						<option value="">請選擇</option>
<%
				strCode="select * from dcicode where typeid=3"
				set rsCode=conn.execute(strCode)
				If Not rsCode.Bof Then rsCode.MoveFirst 
				While Not rsCode.Eof	
%>
						<option value="<%=trim(rsCode("ID"))%>"><%=trim(rsCode("Content"))%></option>
<%
				rsCode.MoveNext
				Wend
				rsCode.close
				set rsCode=nothing
%>
					</select>
					<br>
					備註
					<input type="text" name="DelNote" value="" maxlength="25" size="40">
				</td>
			</tr>
			<tr>
				<td bgcolor="#FFDD77" colspan="4" align="center">
				<input type="button" name="close" value="確定刪除" onclick="BillDel();" >
				<input type="button" name="close" value=" 離 開 " onclick="window.close();">
				<input type="hidden" name="kinds" value="">
				<input type="hidden" name="BillNoSelectlist" value="">
		
				</td>
			</tr>
		</table>		
	</form>
<%
conn.close
set conn=nothing
%>
</body>
<script type="text/javascript" src="../js/date.js"></script>
<script type="text/javascript" src="../js/form.js"></script>
<script language="JavaScript">
function BillDel(){
	var BillNoList;
	for(j=0;j<document.all.DelBillList.length;j++){
		if(j==0){
			BillNoList = document.all.DelBillList.options[j].value ;		  	 
		}else{
			BillNoList = BillNoList + "," + document.all.DelBillList.options[j].value ;		  	 
		}
	}
	myForm.BillNoSelectlist.value=BillNoList;
	
	if (BillNoList=="" || BillNoList==undefined ){
		alert("請輸入舉發單單號！！");
	}else if (myForm.DelReason.value==""){
		alert("請選擇刪除舉發單原因！！");
	}else{
		myForm.kinds.value="DB_BillDel";
		myForm.submit();
	}
}
function addDelBillNo(){
	var DelBillNo=myForm.DelBillNo.value;
	if (DelBillNo==""){
		alert("請輸入舉發單號！！");
	}else{
		runServerScript("getDelBillNoToList.asp?sysBillNo="+DelBillNo);
	}
}
function addDelBillNoToList(OptValue,OptContent,ErrorCode){
	 var opt;
	 
	 obj = document.all.DelBillList ;
	
	if (ErrorCode=="1"){
		alert("找不到此筆單號，請確認 單號是否正確 或 該單號已經刪除！！");
	}else if (ErrorCode=="2"){
		alert("您沒有刪除此筆單號的權限！！");
	}else if (ErrorCode=="3"){
		alert("請在此筆舉發單入案回傳後再做刪除！！");
	}else{
		errFlg = false;	 
		for(i=0;i<document.all.DelBillList.length;i++){
		  oldValue = document.all.DelBillList.options[i].value;
	   	  if (OptValue==oldValue){
	   	  	 alert("該單號已加入過！！");
	   	  	 errFlg = true;	 	  	 
	   	  	 break;
	   	  }
		}  
		if (errFlg==false){
			if (obj.length==0){
				nextIndex = 0;
			}else{
				nextIndex = eval(obj.length) ; 
			}
			opt = new Option(OptContent,OptValue);
			document.all.DelBillList.options[nextIndex] = opt;  
			myForm.DelBillNo.value="";
		}
	}
}
function cancelDelBillNo(){
	obj = document.all.DelBillList ;
	objIndex = obj.selectedIndex;
	while(objIndex != -1){
		if (objIndex != -1) {
			obj.remove(objIndex);
		}
		objIndex = obj.selectedIndex;
	}
}
</script>
</html>
