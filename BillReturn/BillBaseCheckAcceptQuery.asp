<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Bannernodata.asp"-->
<%
	DB_Selt=trim(request("DB_Selt"))
	tablename="BillStopCarAccept"

	If trim(Request("DB_state"))="Update" Then
		recordid=gInitDT(Request("DB_RecordDate"))&hour(Request("DB_RecordDate"))&minute(Request("DB_RecordDate"))&second(Request("DB_RecordDate"))

		str_Mem3="Sys_Mem3"&trim(Request("DB_BillUnitID"))&recordid&trim(Request("DB_AcceptDate"))

		str_unit="Sys_"&trim(Request("DB_BillUnitID"))&recordid&trim(Request("DB_AcceptDate"))

		str_memid="Sys_Chk"&trim(Request("DB_BillUnitID"))&recordid&trim(Request("DB_AcceptDate"))


		If (not ifnull(Request(str_Mem3))) and not (ifnull(Request(str_unit))) Then
			strSQL="update "&trim(Request("DB_BillType"))&" set recordstateid=-1,Note='"&trim(Request(str_Mem3))&"' where billno='"&trim(Request(str_unit))&"' and billunitid='"&trim(Request("DB_BillUnitID"))&"' and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)&"  and recorddate="&funGetDate(Request("DB_RecordDate"),1)

			conn.execute(strSQL)

			Response.write "<script>"
			Response.Write "alert('儲存完成！');"
			Response.write "</script>"
		elseIf not ifnull(Request(str_memid)) Then
			strSQL="update "&trim(Request("DB_BillType"))&" set RecordMemberID2="&Session("User_ID")&" where billunitid='"&trim(Request("DB_BillUnitID"))&"' and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)&"  and recorddate="&funGetDate(Request("DB_RecordDate"),1)

			conn.execute(strSQL)

			Response.write "<script>"
			Response.Write "alert('儲存完成！');"
			Response.write "</script>"

		else
			Response.write "<script>"
			Response.Write "alert('退件單號及原因不能空白！');"
			Response.write "</script>"
		end if
	End if
	

	if DB_Selt="Selt" then
		where=""				
		If not ifnull(Request("Sys_Billno")) Then
			where=where&" and BillNo='"&trim(Request("Sys_Billno"))&"'"
		End if

		If not ifnull(Request("Sys_BillUnitID")) Then
			sqlUit="select UnitLevelid from unitinfo where unitid='"&trim(Request("Sys_BillUnitID"))&"'"
			set rsuit=conn.execute(sqlUit)
			If trim(rsuit("UnitLevelid"))="2" Then
				where=where&" and BillUnitID in(select UnitID from UnitInfo where UnitTypeid='"&trim(Request("Sys_BillUnitID"))&"')"
			else
				where=where&" and BillUnitID in('"&trim(Request("Sys_BillUnitID"))&"')"
			End if
			rsuit.close
		End if

		If not ifnull(Request("Sys_BillMem")) Then
			where=where&" and BillMemID1="&trim(Request("Sys_BillMem"))
		End if

		If not ifnull(Request("chkAccept")) Then
			If trim(Request("chkAccept"))="1" Then
				where=where&" and RecordMemberID2 is null"

			elseIf trim(Request("chkAccept"))="2" Then
				where=where&" and RecordMemberID2 is not null"

			End if
		End if		
	end if

	If not ifnull(Request("billtype")) Then
		If trim(Request("billtype"))="2" Then
			tablename="BillRunCarAccept"
		elseif trim(Request("billtype"))="1" Then
			tablename="BillStopCarAccept"
		else
			tablename="BillStopCarAccept"
		end if
	End if

	If trim(Request("chkAccept"))="" Then
		where=where&" and RecordMemberID2 is null"
	End if

	Union=" Union all select 'BillRunCarAccept' BillType,a.AcceptDate,a.BillUnitID,a.recorddate,a.suess,a.delss,b.UnitName,c.chname chname1,d.chname chname2,e.memberid memberid3,e.unitid recordunitid from (select AcceptDate,BillUnitID,recorddate,recordmemberid1,recordmemberid2,recordmemberid3,sum(decode(recordstateid,0,1,0)) suess, sum(decode(recordstateid,-1,1,0)) delss from BillRunCarAccept where recordmemberid1 is not null"&where&" group by AcceptDate,RecordDate,BillUnitID,recordmemberid1,recordmemberid2,recordmemberid3) a,UnitInfo b,memberdata c,memberdata d,memberdata e where a.BillUnitID=b.UnitID and a.recordmemberid1=c.memberid(+) and a.recordmemberid2=d.memberid(+) and a.recordmemberid3=e.memberid(+)"

	If not ifnull(Request("billtype")) Then
		If trim(Request("billtype"))="2" Then
			Union=""

		elseif trim(Request("billtype"))="1" Then
			Union=""

		end if
	End if

	strSQL="select '"&tablename&"' BillType,a.AcceptDate,a.BillUnitID,a.recorddate,a.suess,a.delss,b.UnitName,c.chname chname1,d.chname chname2,e.memberid memberid3,e.unitid recordunitid from (select AcceptDate,BillUnitID,recorddate,recordmemberid1,recordmemberid2,recordmemberid3,sum(decode(recordstateid,0,1,0)) suess, sum(decode(recordstateid,-1,1,0)) delss from "&tablename&" where recordmemberid1 is not null"&where&" group by AcceptDate,RecordDate,BillUnitID,recordmemberid1,recordmemberid2,recordmemberid3) a,UnitInfo b,memberdata c,memberdata d,memberdata e where a.BillUnitID=b.UnitID and a.recordmemberid1=c.memberid(+) and a.recordmemberid2=d.memberid(+) and a.recordmemberid3=e.memberid(+)"&Union

	set rsfild=conn.execute(strSQL&" order by billunitid,AcceptDate,recorddate")
	
	cntSQL="select count(1) cmt from ("&strSQL&")"
	set rscnt=conn.execute(cntSQL)
	DBsum=cdbl(rscnt("cmt"))
	rscnt.close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>點收件查詢系統</title>
<!--#include virtual="traffic/Common/css.txt"-->
<style type="text/css">
.btn3{
   font-family:新細明體;
   background-color:#EEEEEE;
   border-style:solid;
}
</style>
</head>
<body>
<form name="myForm" method="post">
<table width="100%" height="100%" border="0">
	<tr>
		<td bgcolor="#FFCC33" height="33"><b>點收件查詢系統</b>
			<a href="./Upaddress/CheckAccept.doc"><font size="3" color="blue"><u>點收件系統使用說明</u></font></a>
		</td>
	</tr>
	<tr>
		<td bgcolor="#CCCCCC">
			<table border="0" bgcolor="#FFFFFF" width="100%">
				<tr>
					<td>(標示單)單號</td>
					<td>
						<input name="Sys_Billno" class="btn1" type="text" value="<%=request("Sys_Billno")%>" size="12" maxlength="18">
					</td>					
					<td>舉發單位</td>
					<td>
						<%
						UnitName="Sys_BillUnitID"
						MemberName="Sys_BillMem"
						strtmp=""
						strUnitID=""
						strUnitName=""
						strCity="select value from Apconfigure where id=31"
						set rsCity=conn.execute(strCity)
						sys_City=trim(rsCity("value"))
						rsCity.close
						strSQL="select UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"'"
						set rsUnit=conn.execute(strSQL)
						If Not rsUnit.eof Then strUnitName=trim(rsUnit("UnitName"))
						rsUnit.close
						if trim(MemberName)<>"" then
							strtmp="<select name="""&UnitName&""" ID="""&UnitName&""" class=""btn1"" onchange=""UnitMan('"&UnitName&"','"&MemberName&"');"">"
						else
							strtmp="<select name="""&UnitName&""" ID="""&UnitName&""" class=""btn1"">"
						end if
						if trim(Session("UnitLevelID"))="1" or Instr(strUnitName,"交通隊")>0 then
							strSQL="select UnitID,UnitName from UnitInfo order by UnitOrder,UnitTypeID,UnitName"
							strtmp=strtmp+"<option value="""">所有單位</option>"
						elseif trim(Session("UnitLevelID"))="2" or Instr(strUnitName,"分局")>0 then
							strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"' or UnitTypeID='"&Session("Unit_ID")&"' order by UnitOrder,UnitTypeID,UnitName"

							If Instr(strUnitName,"分局")>0 and Instr(strUnitName,"組")>0 and sys_City<>"南投縣" Then
								strSQL="select UnitID,UnitName from UnitInfo where UnitID in(select UnitTypeID from UnitInfo where  UnitID='"&Session("Unit_ID")&"') or UnitTypeID in(select UnitTypeID from UnitInfo where  UnitID='"&Session("Unit_ID")&"') order by UnitOrder,UnitTypeID,UnitName"
							End if
							set rs1=conn.execute(strSQL)
							while Not rs1.eof
								if trim(strUnitID)<>"" then strUnitID=trim(strUnitID)&","
								if trim(strUnitID)="" then
									strUnitID=strUnitID&trim(rs1("UnitID"))
								else
									strUnitID=strUnitID&"'"&trim(rs1("UnitID"))
								end if
								rs1.movenext
								if Not rs1.eof then strUnitID=strUnitID&"'"
							wend
							rs1.close
							strtmp=strtmp+"<option value="""&strUnitID&""">所有單位</option>"
						elseif trim(Session("UnitLevelID"))="3" then
							strSQL="select UnitID,UnitName from UnitInfo where UnitID='"&Session("Unit_ID")&"' order by UnitOrder,UnitTypeID,UnitName"
						end if
						set rs1=conn.execute(strSQL)
						while Not rs1.eof
							strtmp=strtmp+"<option value="""&rs1("UnitID")&""""
							if trim(rs1("UnitID"))=trim(request(UnitName)) then
								strtmp=strtmp+" selected"
							end if
							strtmp=strtmp+">"&rs1("UnitID")&" - "&rs1("UnitName")&"</option>"
							rs1.movenext
						wend
						rs1.close
						strtmp=strtmp+"</select>"
						Response.Write strtmp
						%>
					</td>
					<td>舉發人員</td>
					<td>
						<%=UnSelectMemberOption("Sys_BillUnitID","Sys_BillMem")%>
					</td>
				</tr>
				<tr>
					<td>上傳日期</td>
					<td>
						<input name="Sys_AcceptDate" class="btn1" type="text" value="<%=request("Sys_AcceptDate")%>" size="5" maxlength="7">

						<input class="btn3" style="width:20px;height:20px;font-size:14px;" type="button" name="datestr" value="..." onclick="OpenWindow('Sys_AcceptDate');">
					</td>

					<td>點收類別</td>
					<td>
						<select name="billtype">
							<option value="">全部</option>
							<option value="1"<%If trim(Request("billtype"))="1" Then Response.Write " selected"%>>攔停</option>
							<option value="2"<%If trim(Request("billtype"))="2" Then Response.Write " selected"%>>逕舉</option>
						</select>

						交通隊簽收
						<select name="chkAccept">
							<option value="1"<%If trim(Request("chkAccept"))="!" Then Response.Write " selected"%>>未簽收</option>
							<option value="2"<%If trim(Request("chkAccept"))="2" Then Response.Write " selected"%>>已簽收</option>
							<option value="0"<%If trim(Request("chkAccept"))="0" Then Response.Write " selected"%>>全部</option>
						</select>
					</td>
					
					<td colspan="2">
						<input class="btn3" style="width:40px;height:25px;font-size:14px;" type="submit" name="btnSelt" value="查詢" onClick='funSelt();'>&nbsp;&nbsp;
						<input class="btn3" style="width:40px;height:25px;font-size:14px;" type="button" name="cancel" value="清除" onClick="location='BillBaseCheckAcceptQuery.asp'">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFCC33" height="33">
		 ( 查詢 <%=DBsum%> 筆紀錄 )
		<select name="sys_MoveCnt" onchange="repage();">
				<option value="0"<%if trim(request("sys_MoveCnt"))="0" then response.write " Selected"%>>10</option>
				<option value="10"<%if trim(request("sys_MoveCnt"))="10" then response.write " Selected"%>>20</option>
				<option value="20"<%if trim(request("sys_MoveCnt"))="20" then response.write " Selected"%>>30</option>
				<option value="30"<%if trim(request("sys_MoveCnt"))="30" then response.write " Selected"%>>40</option>
				<option value="40"<%if trim(request("sys_MoveCnt"))="40" then response.write " Selected"%>>50</option>
				<option value="50"<%if trim(request("sys_MoveCnt"))="50" then response.write " Selected"%>>60</option>
				<option value="60"<%if trim(request("sys_MoveCnt"))="60" then response.write " Selected"%>>70</option>
				<option value="70"<%if trim(request("sys_MoveCnt"))="70" then response.write " Selected"%>>80</option>
				<option value="80"<%if trim(request("sys_MoveCnt"))="80" then response.write " Selected"%>>90</option>
				<option value="90"<%if trim(request("sys_MoveCnt"))="90" then response.write " Selected"%>>100</option>
			</select>
		</td>
	</tr>
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" height="100%" border="0" cellpadding="1" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th height="30" nowrap>上傳日期</th>
					<th height="34" nowrap>舉發單位</th>
					<th height="34" nowrap>點收人員</th>
					<th height="34" nowrap>已點收數</th>					
					<th height="34" nowrap>已退件數</th>
					<th height="34" nowrap>操作</th>
				</tr><%
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfild.eof then rsfild.move Cint(DBcnt)
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfild.eof then exit for
						response.write "<tr bgcolor='#FFFFFF' align='center' "
						lightbarstyle 0 
						response.write ">"

						response.write "<td align='center' nowrap>"&gInitDT(rsfild("AcceptDate"))&"</td>"
						response.write "<td align='center' nowrap>"&rsfild("UnitName")&"</td>"

						sys_chname=trim(rsfild("chname1"))
						If not ifnull(rsfild("chname2")) Then
							If not ifnull(sys_chname) Then sys_chname=sys_chname&","
							sys_chname=sys_chname&trim(rsfild("chname2"))
						End if
						
						response.write "<td align='center' nowrap>"&sys_chname&"</td>"						
						response.write "<td align='center' nowrap>"&rsfild("suess")&"</td>"
						response.write "<td align='center' nowrap>"&rsfild("delss")&"</td>"

						recordid=gInitDT(rsfild("recorddate"))&hour(rsfild("recorddate"))&minute(rsfild("recorddate"))&second(rsfild("recorddate"))

						selt_unit="Sys_"&trim(rsfild("billunitid"))&recordid&gInitDT(rsfild("AcceptDate"))
						selt_Mem3="Sys_Mem3"&trim(rsfild("billunitid"))&recordid&gInitDT(rsfild("AcceptDate"))
						chk_memid="Sys_Chk"&trim(rsfild("billunitid"))&recordid&gInitDT(rsfild("AcceptDate"))

						RecordDate=datevalue(rsfild("recorddate"))&" "&hour(rsfild("recorddate"))&":"&minute(rsfild("recorddate"))&":"&second(rsfild("recorddate"))

						response.write "<td align='left' nowrap>"
						If trim(Session("UnitLevelID"))=1 then
							Response.Write "<input class='btn1' type='checkbox' name="""&chk_memid&""" value='1'>簽收　"

							Response.Write "退件單號:"

							Response.Write "<input type=""text"" name="""&selt_unit&""" ID="""&selt_unit&""" class=""btn1"" size=""12"" value="""">"

							Response.Write "退件原因:"

							Response.Write "<input type=""text"" name="""&selt_Mem3&""" ID="""&selt_Mem3&""" class=""btn1"" size=""12"" value="""">"

							Response.Write "<input type=""button"" name=""Update"" value=""確定"" class=""btn3"" style=""width:40px;height:25px;font-size:12px;"" onclick=""funAcceptCreat('"&gInitDT(rsfild("AcceptDate"))&"','"&rsfild("BillUnitID")&"','"&RecordDate&"','"&trim(rsfild("BillType"))&"');"">"

						end if

						Response.Write "&nbsp;<input type=""button"" name=""Update"" value=""詳細"" class=""btn3"" style=""width:40px;height:25px;font-size:12px;"" onclick=""funAcceptLoad('"&gInitDT(rsfild("AcceptDate"))&"','"&rsfild("BillUnitID")&"','"&RecordDate&"','"&trim(rsfild("BillType"))&"');"">"
						

						Response.Write "</td>"
						response.write "</tr>"
						rsfild.movenext
					next%>
			</table>
		</td>
	</tr>
	<tr>
		<td bgcolor="#FFDD77" align="center">

			<input type="button" class="btn3" style="width:60px;height:25px;font-size:14px;" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" class="btn3" style="width:60px;height:25px;font-size:14px;" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
		</td>
	</tr>
</table>
<input type="Hidden" name="DB_Selt" value="<%=DB_Selt%>">
<input type="Hidden" name="DB_state" value="">
<input type="Hidden" name="DB_AcceptDate" value="">
<input type="Hidden" name="DB_BillUnitID" value="">
<input type="Hidden" name="DB_RecordDate" value="">
<input type="Hidden" name="DB_BillType" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
<%="UnitMan('Sys_BillUnitID','Sys_BillMem','"&request("Sys_BillMem")&"');"%>
function ListItemID(chkMemID,UnitListName,MemListName){
	runServerScript("/traffic/Common/ListItem.asp?LoginID="+document.all[chkMemID].value+"&UnitListName="+UnitListName+"&MemListName="+MemListName);
}
function funAcceptCreat(AcceptDate,BillUnitID,RecordDate,BillType){
	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordDate.value=RecordDate;
	myForm.DB_BillType.value=BillType;
	myForm.DB_state.value="Update";
	myForm.submit();
}
function funAcceptLoad(AcceptDate,BillUnitID,RecordDate,BillType){
	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordDate.value=RecordDate;
	if(BillType=='BillRunCarAccept'){
		UrlStr="AcceptRunList.asp";
	}else{
		UrlStr="AcceptStopList.asp";
	}
	
	myForm.action=UrlStr;
	myForm.target="PrintAccept";
	myForm.submit();
	myForm.action="";
	myForm.target="";
}
function funSelt(){
	myForm.DB_Move.value=0;
	myForm.DB_Selt.value="Selt";
	myForm.submit();
}
function funDbMove(MoveCnt){
	if (eval(MoveCnt)==0){
		myForm.DB_Move.value="";
		myForm.submit();
	}else if (eval(MoveCnt)==10){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==-10){
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else if(eval(MoveCnt)==999){
		if (eval(myForm.DB_Cnt.value)%(10+eval(myForm.sys_MoveCnt.value))==0){
			myForm.DB_Move.value=(Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))-1)*(10+eval(myForm.sys_MoveCnt.value));
		}else{
			myForm.DB_Move.value=Math.floor(eval(myForm.DB_Cnt.value)/(10+eval(myForm.sys_MoveCnt.value)))*(10+eval(myForm.sys_MoveCnt.value));
		}
		myForm.submit();
	}
}

function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}
</script>
<%conn.close%>