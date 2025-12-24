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
<!--#include virtual="/traffic/Common/CssForCaseIn.txt"-->
<title>附加說明</title>
<%
	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=Nothing

	strLaw="select * from Law where ItemID='"&trim(request("RuleID"))&"' and Version='"&trim(request("theRuleVer"))&"'"
	set rsLaw=conn.execute(strLaw)
	if not rsLaw.eof then 
		IllArrayStr=trim(rsLaw("IllegalRule"))
	end if
	rsLaw.close
	set rsLaw=nothing

if trim(request("kinds"))="Del_LawPlus" then
	strDel="Delete from Code where ID="&trim(request("LawPlusID"))&" and TypeID=88"
	conn.execute strDel
end if
if trim(request("kinds"))="Add_LawPlus" then
	strSn="select max(ID) as MaxID from Code"
	set rsSn=conn.execute(strSn)
	if not rsSn.eof then
		MaxID=trim(rsSn("MaxID"))+1
	end if
	rsSn.close
	set rsSn=nothing
	strAdd="insert into code values("&MaxID&",88,'"&trim(request("LawPlus_Value"))&"',"&Trim(session("User_ID"))&",0,0,'0','"&Trim(Session("User_ID"))&"')"
	conn.execute strAdd
end if
%>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<form name="myForm" method="post">  
		<table width='100%' border='1' align="left" cellpadding="1">
			<tr bgcolor="#FFCC33">
				<td colspan="7" >附加說明</td>
			</tr>
			<tr>
				<td colspan="2">
				<input type="button" name="qqq" value="確定" onclick="LawPlus_Click();" <%
	strPlus2="select ID,Content from code where TypeID=88"
	set rsPlus2=conn.execute(strPlus2)
	if rsPlus2.eof then
		response.write "disabled"
	end if
	rsPlus2.close
	set rsPlus2=nothing
				%>>
				<input type="button" name="qqqqq" value="取消附加說明" onclick="LawPlus_Cancel();" <%
	strPlus2="select ID,Content from code where TypeID=88"
	set rsPlus2=conn.execute(strPlus2)
	if rsPlus2.eof then
		response.write "disabled"
	end if
	rsPlus2.close
	set rsPlus2=nothing
				%>>
				<br />
				<input type="text" name="KeyWord" value="<%=Trim(request("KeyWord"))%>">
				<input type="Button" name="QueryWord_B" value="查詢" onclick="QueryKeyWord();">
				<br />
				<font color="red">勾選後，按『確定』按鈕帶入</font>
				</td>
			<%If sys_City="台南市" then%>
				<td>新增人</td>
			<%End if%>
			</tr>
<%
	strSqlAdd=""
	if trim(request("kinds"))="QueryKeyWord" Then
		strSqlAdd=strSqlAdd & " and Content like '%" & trim(request("KeyWord")) & "%'"
	End If 
	If sys_City="基隆市" Or sys_City="台南市" Or sys_City="高雄市" Then
		LawForUser=1
	Else
		LawForUser=0
	End If 
	If LawForUser=1 Then
		strSqlAdd=strSqlAdd & " and ShowOrder=" & Trim(session("User_ID"))
	End If 
	strPlus="select ID,Content,NpaID from code where TypeID=88" & strSqlAdd
	set rsPlus=conn.execute(strPlus)
	If Not rsPlus.Bof Then rsPlus.MoveFirst 
	While Not rsPlus.Eof
%>
			<tr>
				<td width="80%">
				<input type="radio" name="rdLawPlusID" value="<%=trim(rsPlus("Content"))%>"><%=trim(rsPlus("Content"))%>
				</td>
				<td width="10%">
				<input type="button" value="刪除" onclick="Del_LawPlus(<%=trim(rsPlus("ID"))%>);" <%
				If Trim(Session("Credit_ID"))<>"A000000000" And sys_City="高雄市" Then
					response.write "disabled"
				End If
				If LawForUser=1 Then
					If Trim(rsPlus("NpaID"))<>"" And Trim(rsPlus("NpaID"))<>"0" And Trim(Session("Credit_ID"))<>"A000000000" Then
						If Trim(rsPlus("NpaID"))<>Trim(Session("User_ID")) Then
							response.write "disabled"
						End If 
					End If 
				End if
				%>>
				</td>
			<%If LawForUser=1 then%>
				<td>
				<%
				If Trim(rsPlus("NpaID"))<>"" Then
					strM="select * from MemberData where MemberID="&Trim(rsPlus("NpaID"))
					Set rsM=conn.execute(strM)
					If not rsM.eof Then
						response.write Trim(rsM("chName"))
					Else
						response.write "&nbsp;"	
					End If
					rsM.close
					Set rsM=Nothing 
				Else
					response.write "&nbsp;"	
				End If 
				%>
				</td>
			<%End if%>
			</tr>
<%
	rsPlus.MoveNext
	Wend
	rsPlus.close
	set rsPlus=nothing
%>			<tr>
				<td colspan="3">
				<input type="text" value="" name="LawPlus_Value" size="32" maxlength="20">
				<input type="button" value="新增附加說明" onclick="Add_LawPlus();">
				<input type="hidden" value="" name="kinds">
				<input type="hidden" value="" name="LawPlusID">
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
function LawPlus_Click(){
	if (myForm.rdLawPlusID.length>0 ){
		for (i=0; i<myForm.rdLawPlusID.length; i++){
			if (myForm.rdLawPlusID[i].checked==true){
				opener.Layer1.innerHTML="<%=IllArrayStr%>"+"("+myForm.rdLawPlusID[i].value+")";
				opener.myForm.Rule4.value=myForm.rdLawPlusID[i].value;
			}
		}
		window.close();
	}else{
		if (myForm.rdLawPlusID.checked==true){
			opener.Layer1.innerHTML="<%=IllArrayStr%>"+"("+myForm.rdLawPlusID.value+")";
			opener.myForm.Rule4.value=myForm.rdLawPlusID.value;
		}
		window.close();
	}
	//alert(myForm.rdLawPlusID(0).value);
	//myForm.submit();
}
function LawPlus_Cancel(){
	opener.Layer1.innerHTML="<%=IllArrayStr%>";
	opener.myForm.Rule4.value="";
	window.close();
}
function Del_LawPlus(LID){
	myForm.kinds.value="Del_LawPlus";
	myForm.LawPlusID.value=LID;
	myForm.submit();
}
function Add_LawPlus(){
	if (myForm.LawPlus_Value.value==""){
		alert("請輸入附加說明!");
	}else{
		myForm.kinds.value="Add_LawPlus";
		myForm.submit();
	}
}

function QueryKeyWord(){
	if (myForm.KeyWord.value==""){
		alert("請輸入關鍵字!");
	}else{
		myForm.kinds.value="QueryKeyWord";
		myForm.submit();
	}
}
</script>
</html>
