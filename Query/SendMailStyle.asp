<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<script language=javascript src='../js/form.js'></script>
<script language=javascript>
var bar = new Array();
var bar2 = new Array();
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

sqlUnit = "select * from UnitInfo where unitid in (Select Distinct UnitTypeId From UnitInfo )"
Set RsUnit = Conn.Execute(sqlUnit)     
While Not RsUnit.Eof
	tmpStr = ""
	tmpStr2 = ""
	If sys_City<>"台南市" Then
		if RsUnit("UnitLevelID")="1" then
			sqlUnit = "Select UnitId, UnitName From UnitInfo Where (UnitLevelID=2 and unitid in (Select Distinct UnitTypeId From UnitInfo )) or (UnitTypeID='"&RsUnit("UnitId")&"') "
		else
			sqlUnit = "Select UnitId, UnitName From UnitInfo Where UnitTypeId='" & RsUnit("UnitId") & "' "
		end if
	else
		sqlUnit = "Select UnitId, UnitName From UnitInfo Where UnitTypeId='" & RsUnit("UnitId") & "' "
	End if
	Set RsUnit2 = Conn.Execute(sqlUnit)      
	While Not RsUnit2.Eof
		tmpStr = tmpStr & RsUnit2("UnitId") & ","
		tmpStr2 = tmpStr2 & Trim(RsUnit2("UnitName")) & ","
		RsUnit2.MoveNext
	Wend
	RsUnit2.close
	if tmpStr<>"" then
		tmpStr = Left(tmpStr,Len(tmpStr) - 1)
		tmpStr2 = Left(tmpStr2,Len(tmpStr2) - 1)
		tmpStr = RsUnit("UnitId") & "," & tmpStr
		tmpStr2 = RsUnit("UnitName") & "," & tmpStr2
		Response.Write "bar[""" & RsUnit("UnitId") & """] = """ & tmpStr & """;" & CHR(10)   
		Response.Write "bar2[""" & RsUnit("UnitId") & """] = """ & tmpStr2 & """;" & CHR(10)       
	end if
	RsUnit.MoveNext
Wend
RsUnit.close
%>
function window_onload(){
   document.all.unit.checked = false;
   document.all.unit.value = "n";
   document.all.UnitID_q.disabled = true;
}
</script>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>查詢郵寄未退回清冊</title>
</head>
<body onload="window_onload();">
<form name=myForm method="post">
<table width="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33"><font size=4>查詢未退還清冊</font></td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="1" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr>
									
					<td nowrap valign="top" bgcolor="#FFFFFF">
		
						<b>統計期間</b>
						<input name="chkDate" type="radio" value="0" checked>填單日
						<input name="chkDate" type="radio" value="1">違規日
						<input name="chkDate" type="radio" value="2">建檔日<br>
						<input type='text' size='9' name='startDate_q' value='' maxLength='8'>
						<input name="datestra" type="button" value="..." onclick="OpenWindow('startDate_q');">
						~
						<input type='text' size='9' name='endDate_q' value='' maxLength='8'>
						<input name="datestrb" type="button" value="..." onclick="OpenWindow('endDate_q');">
						
						<br>			<br>	
						<b>舉發單號</b>
						<input type='text' size='10' name='startBillNo_q' value='' maxLength='9' onkeyup="this.value=this.value.toUpperCase()"> ~
						<input type='text' size='10' name='endBillNo_q' value='' maxLength='9' onkeyup="this.value=this.value.toUpperCase()">
						<br>			<br>			
							<input type="radio" name="rdMemberIn" value="0" checked ><b>所有人員</b><br>
							<input type="radio" name="rdMemberIn" value="1"  >選擇建檔人員<input type="button" value="加入" name="btnMemberIn"  onclick="openQryMemberTypeList();">
          					<input type="button" name="btnDelMemberSelect" onClick="MemberSelectremove();" value="刪除"><br>
							<select name="MemberSelect" id="MemberSelect" multiple size=5></select>				
						
					</td>
					<td nowrap bgcolor="#FFFFFF">

						<b>入案批號</b>						<input name="batchnumber" type="text" value="" onkeyup="this.value=this.value.toUpperCase()">		
						<br>	<br>			
						<input name="unit" type="checkbox" value="y" onClick="ctlUnit();">
						<b>單位</b>（不勾選代表統計所有單位）<br>
						<select name="UnitID_q" ><%
							sqlUnit = "select UnitTypeId,unitid,unitname from unitinfo where unitid in (select distinct unittypeid from unitinfo) " & _
							"Union select UnitTypeId,unitid,unitname from unitinfo where unittypeid is null Order By UnitTypeId"
							set RsUnit=Server.CreateObject("ADODB.RecordSet")
							RsUnit.open sqlUnit,Conn,3,3
							While Not RsUnit.Eof%>
								<option value="<%=RsUnit("UnitID")%>" <%if RsUnit("UnitID")=Request("UnitID_q") then response.write " selected" end if%>><%=RsUnit("UnitName")%></option><%
								RsUnit.MoveNext
							Wend%>
						</select>
						<br>
						<INPUT TYPE="button" name="btnAddUnit" value="加入統計" onClick="addUnit();">
						<INPUT TYPE="button" name="btnDelUnit" value="刪除統計" onClick="removeUnit();">		  
						<br>
						<select name="unitSelect" id="unitSelect" multiple size=5></select>
						<br>
						用Ctrl或Shift可以多選
						<br>
		<%
		if sys_City="台中市" or sys_City="南投縣" or sys_City="雲林縣" Or sys_City="基隆市" Or sys_City="苗栗縣" Or sys_City="澎湖縣" Or sys_City="花蓮縣" then
		%>
						<input type="checkbox" name="ReturnDateFlag" value="1"><b>單退_寄存上傳日</b>(二次未退回才有用)
						<br>
						<input type="text" name="ReturnDate1" value="" size='9' maxlength="7">
						<input name="datestrv" type="button" value="..." onclick="OpenWindow('ReturnDate1');">
						~ 
						<input type="text" name="ReturnDate2" value="" size='9' maxlength="7">
						<input name="datestrv" type="button" value="..." onclick="OpenWindow('ReturnDate2');">
		<%end if%>
					</td>

				</tr>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td bgcolor="#FFDD77">
			<input name="btnadd" type="button" value="產生清冊 " onclick="funAdd();"> 
		
		<%
		If sys_City="苗栗縣" Then
		%>
			<input name="btnadd" type="button" value="產生清冊(苗栗) " onclick="funAdd_P_M();"> 
			<input name="btnadd" type="button" value="產生清冊(個資版) " onclick="funAdd_P();"> 
		<%
		End If 
		if sys_City="台中市" or sys_City="南投縣" or sys_City="雲林縣" Or sys_City="基隆市" Or sys_City="苗栗縣" Or sys_City="澎湖縣" or sys_City="花蓮縣" then
		%>
			<input name="btnadd" type="button" value="產生二次郵寄未退回清冊 " onclick="funAdd2();"> 
		<%end If
		If sys_City="苗栗縣" Then
		%>
			<input name="btnadd" type="button" value="產生二次郵寄未退回清冊(苗栗) " onclick="funAdd2_P_M();"> 
			<input name="btnadd" type="button" value="產生二次郵寄未退回清冊(個資版) " onclick="funAdd2_P();"> 
		<%
		End If
		%>

			<input name="btnexit" type="button" value=" 關 閉 " onclick="funExt();">
		<%
		if sys_City<>"花蓮縣" then
		%>	
			<input name="btnexit" type="button" value=" 產生郵局查詢單(A4) " onclick="funMailQry();">
		<%end if%>

		<%
		if sys_City="台中市" or sys_City="南投縣" or sys_City="雲林縣" Or sys_City="基隆市" Or sys_City="苗栗縣" Or sys_City="澎湖縣" or sys_City="花蓮縣" then
		%>
			<input name="btnexit" type="button" value=" 產生二次郵局查詢單(A4) " onclick="funMailQry2();">
		<%end if%>

		<%		if sys_City="花蓮縣" then		%>
			<input name="btnexit" type="button" value=" 產生一次郵局查詢單新版(A4) " onclick="funMailQry3();">
		<%end if%>

		<%		if sys_City="台中市" then		%>
			<input name="btnadd" type="button" value="產生單次未退回清冊 " onclick="funAddSingle();"> 
			<input name="btnadd" type="button" value="產生單次未退回清冊(橫印附條碼版) " onclick="funAddSingleBarCode();"> 
			<input name="btnexit" type="button" value=" 產生單次郵局查詢單(A4) " onclick="funMailQrySingle();">
			<br>
			<input name="btnadd" type="button" value="產生單次未退回清冊(橫印附條碼版)郵局送達I4" onclick="funAddSingleBarCode_I4();"> 
		<%end if%>

				
			<img src="space.gif" width="20" height="5">
		<%if sys_City="南投縣" then%>
			<br>
			<input name="btnexit" type="button" value="未退回件數統計" onclick="funMailNotBackReport();">
		<%end if%>
		</td>
	</tr>
</table>
	<input type="hidden" name="unitSelectlist">
	<input type="hidden" name="MemSelectlist">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
function ctlUnit(){
	if (document.all.unit.checked==true){
		document.all.unit.value="y";
		document.all.UnitID_q.disabled=false;
		document.all.btnAddUnit.disabled = false;  
		document.all.btnDelUnit.disabled = false;  
		document.all.unitSelect.disabled = false; 		
	}else{
		document.all.unit.value="n";
		document.all.UnitID_q.disabled=true;
		document.all.btnAddUnit.disabled = true;  
		document.all.btnDelUnit.disabled = true;  
		document.all.unitSelect.disabled = true; 		
	}
}
function addUnit(){
	 var opt;
	 var oldValue;
	 var errFlg;
	 var tmpAry;
	 var tmpAry2;
	 
   obj = document.all.unitSelect ;
   objUnit = document.all.UnitID_q ;
   objUnitIndex = document.all.UnitID_q.selectedIndex;
   objUnitText = document.all.UnitID_q.options[objUnitIndex].text;
   objUnitValue = document.all.UnitID_q.options[objUnitIndex].value;
   errFlg = false;
   tmpAry = bar[objUnitValue] ;
   tmpAry2 = bar2[objUnitValue];
   
   if (tmpAry != undefined){
      var tmpStr = tmpAry.split(","); 
      var tmpStr2 = tmpAry2.split(","); 
      for( j=0; j<tmpStr.length; j++ ){
      	 objUnitValue = tmpStr[j];
      	 errFlg=false;
	       for(i=0;i<document.all.unitSelect.length;i++){
	       	  oldValue = document.all.unitSelect.options[i].value;	
	       	  if (objUnitValue==oldValue){
	       	  	 errFlg=true;
               break;
	       	  }	           	  
         }  
        
         if(errFlg==false){
         	     nextIndex = eval(obj.length);                   
               opt = new Option(tmpStr2[j],tmpStr[j]);
               document.all.unitSelect.options[nextIndex] = opt;  
         } 	  	 
      }    
   }else{
   	 errFlg = false;	 
	   for(i=0;i<document.all.unitSelect.length;i++){
	   	  oldValue = document.all.unitSelect.options[i].value;
	   	  if (objUnitValue==oldValue){
	   	  	 alert("該單位【" + objUnitText + "】已加入過！！");
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
        opt = new Option(objUnitText,objUnitValue);
        document.all.unitSelect.options[nextIndex] = opt;      
     }      	
   }  
  
}

function removeUnit(){
	obj = document.all.unitSelect ;
	objUnit = document.all.UnitID_q ;
	objUnitIndex = obj.selectedIndex;
	objIndex = obj.selectedIndex;
	while(objIndex != -1){
		if (objIndex != -1) {
			obj.remove(objIndex);
		}
		objIndex = obj.selectedIndex;
	}
}

function funExt() {
	if(confirm("是否關閉維護系統?")){
		window.close();
	}
}

function funAdd(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="")){
		alert("入案批號、統計期間、建檔人員、舉發單號碼請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else{
<%if sys_City="台中市" then%>
			//UrlStr="MailNotBakList_TC.asp";
			UrlStr="MailNotBakList.asp";
<%else%>
			UrlStr="MailNotBakList.asp";
<%end if %>
			myForm.action=UrlStr;			
			myForm.target="MailNotBakList";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	}
}
function funAddSingle(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="")){
		alert("入案批號、統計期間、建檔人員、舉發單號碼請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else{
			UrlStr="MailNotBakListSingle.asp";
			myForm.action=UrlStr;			
			myForm.target="MailNotBakList";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	}
}
function funAddSingleBarCode(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="")){
		alert("入案批號、統計期間、建檔人員、舉發單號碼請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else{
			UrlStr="MailNotBakListSingle_TC.asp";
			myForm.action=UrlStr;			
			myForm.target="MailNotBakList";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	}
}

function funAddSingleBarCode_I4(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="")){
		alert("入案批號、統計期間、建檔人員、舉發單號碼請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else{
			UrlStr="MailNotBakListSingle_TC_I4.asp";
			myForm.action=UrlStr;			
			myForm.target="MailNotBakList";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	}
}

function funAdd_P(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="")){
		alert("入案批號、統計期間、建檔人員、舉發單號碼請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else{
			UrlStr="MailNotBakList_P.asp";
			myForm.action=UrlStr;			
			myForm.target="MailNotBakList_P";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	}
}

function funAdd_P_M(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="")){
		alert("入案批號、統計期間、建檔人員、舉發單號碼請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else{
			UrlStr="MailNotBakList_P_M.asp";
			myForm.action=UrlStr;			
			myForm.target="MailNotBakList_P_M";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	}
}

function funAdd2(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="") && (myForm.ReturnDateFlag.checked!=true)){
		alert("入案批號、統計期間、建檔人員、舉發單號碼、單退_寄存上傳日請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else if (myForm.ReturnDateFlag.checked==true && (myForm.ReturnDate1.value=="" || myForm.ReturnDate2.value=="")){
			alert('單退上傳日格式不正確!!');
		}else{
			UrlStr="SecondMailNotBakList.asp";
			myForm.action=UrlStr;			
			myForm.target="SecondMailNotBakList";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	}
}

function funAdd2_P(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="") && (myForm.ReturnDateFlag.checked!=true)){
		alert("入案批號、統計期間、建檔人員、舉發單號碼、單退_寄存上傳日請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else if (myForm.ReturnDateFlag.checked==true && (myForm.ReturnDate1.value=="" || myForm.ReturnDate2.value=="")){
			alert('單退上傳日格式不正確!!');
		}else{
			UrlStr="SecondMailNotBakList_P.asp";
			myForm.action=UrlStr;			
			myForm.target="SecondMailNotBakList_P";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	}
}

function funAdd2_P_M(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="") && (myForm.ReturnDateFlag.checked!=true)){
		alert("入案批號、統計期間、建檔人員、舉發單號碼、單退_寄存上傳日請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else if (myForm.ReturnDateFlag.checked==true && (myForm.ReturnDate1.value=="" || myForm.ReturnDate2.value=="")){
			alert('單退上傳日格式不正確!!');
		}else{
			UrlStr="SecondMailNotBakList_P_M.asp";
			myForm.action=UrlStr;			
			myForm.target="SecondMailNotBakList_P_M";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			//window.close();
		}
	}
}

function funMailQry(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="")){
		alert("入案批號、統計期間、建檔人員、舉發單號碼請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else{
			UrlStr="FAXQryMail.asp";
			myForm.action=UrlStr;			
			myForm.target="FAXQryMail";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			window.close();
		}
	}
}

function funMailQry2(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="") && (myForm.ReturnDateFlag.checked!=true)){
		alert("入案批號、統計期間、建檔人員、舉發單號碼、單退_寄存上傳日請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else{
			UrlStr="FAXQryMail2.asp";
			myForm.action=UrlStr;			
			myForm.target="FAXQryMail2";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			window.close();
		}
	}
}

function funMailQry3(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="") && (myForm.ReturnDateFlag.checked!=true)){
		alert("入案批號、統計期間、建檔人員、舉發單號碼、單退_寄存上傳日請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else{
			UrlStr="FAXQryMail_Hulien.asp";
			myForm.action=UrlStr;			
			myForm.target="FAXQryMail_Hulien";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			window.close();
		}
	}
}

function funMailQrySingle(){
	var sDate = myForm.startDate_q.value;
	var eDate = myForm.endDate_q.value;
	if(myForm.unit.checked){
		for(j=0;j<document.all.unitSelect.length;j++){
			  if(j==0){
				 unitList = document.all.unitSelect.options[j].value ;		  	 
			  }else{
				 unitList = unitList + "," + document.all.unitSelect.options[j].value ;		  	 
			  }
		}
		myForm.unitSelectlist.value=unitList;
	}
	if(myForm.rdMemberIn(1).checked){
		for(j=0;j<document.all.MemberSelect.length;j++){
			  if(j==0){
				 MemberList = document.all.MemberSelect.options[j].value ;		  	 
			  }else{
				 MemberList = MemberList + "," + document.all.MemberSelect.options[j].value ;		  	 
			  }
		}
		myForm.MemSelectlist.value=MemberList;
	}

	if (((sDate=="") || (eDate=="")) && (myForm.batchnumber.value=="") && (myForm.MemSelectlist.value=="") & (myForm.startBillNo_q.value=="" || myForm.endBillNo_q.value=="") && (myForm.ReturnDateFlag.checked!=true)){
		alert("入案批號、統計期間、建檔人員、舉發單號碼、單退_寄存上傳日請擇一輸入!!");

	}else{
		if (((!dateCheck(sDate))||(!dateCheck(eDate))) && ((sDate!="") && (eDate!=""))){
			alert('統計日期格式不正確!!');
		}else if ((sDate > eDate) && ((sDate!="") && (eDate!=""))){
			alert('統計期間之起始日期不得大於結束日期!!');
		}else if(((myForm.startBillNo_q.value!="") && (myForm.endBillNo_q.value!="")) && ((myForm.startBillNo_q.value.length!=9) || (myForm.endBillNo_q.value.length!=9))){
			alert('舉發單號不得小於九碼!!');
		}else{
			UrlStr="FAXQryMail_Single.asp";
			myForm.action=UrlStr;			
			myForm.target="FAXQryMail_Single";
			myForm.submit();
			myForm.action="";
			myForm.target="";
			window.close();
		}
	}
}

function openQryMemberTypeList(){
	myForm.rdMemberIn(1).checked=true;
	 window.open("../Report/Query_MemID.asp?qryType=1&reportId=REPORTBASE0010","tmpWindow","width=600,height=355,left=0,top=0,resizable=yes,scrollbars=yes");
}

function funMailNotBackReport(){
	window.open("MailNotBackReport.asp","tmpWindow","width=600,height=355,left=0,top=0,resizable=yes,scrollbars=yes");
}

function MemberSelectremove(){
   obj = document.all.MemberSelect;
	objIndex = obj.selectedIndex;
	while(objIndex != -1){
		if (objIndex != -1) {
			obj.remove(objIndex);
		}
		objIndex = obj.selectedIndex;
	}
} 
</script>