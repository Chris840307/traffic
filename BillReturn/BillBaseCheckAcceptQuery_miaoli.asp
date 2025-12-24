<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<!--#include virtual="traffic/Common/DB.ini"-->
<!--#include virtual="traffic/Common/AllFunction.inc"-->
<!--#include virtual="traffic/Common/Bannernodata.asp"-->
<%
	DB_Selt=trim(request("DB_Selt"))
	tablename="BillStopCarAccept"

	If trim(Request("DB_state"))="AcceptStopBack" Then

		updstr="Note='"&trim(Request("DB_ObjNote"))&"'"

		If not ifnull(Request("DB_chkState")) Then
			updstr=updstr&",RecordStateid=-1"
		End if

		If not ifnull(Request("DB_RecordMemberID2")) Then			
			chkwhere=" and to_char(RecordDate2,'YYYYMMDDHH24')='"&Request("DB_AcceptDate")&"' and RecordMemberID2="&trim(request("DB_RecordMemberID2"))

		else
			chkwhere=" and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)&" and RecordMemberID1="&trim(request("DB_RecordMemberID1"))
		End if
		
		strSQL="update BillStopCarAccept set "&updstr&" where billunitid='"&trim(Request("DB_BillUnitID"))&"' and BillNo='"&trim(Request("DB_ObjCarNo"))&"'"&chkwhere&" and RecordStateID=0"

		conn.execute(strSQL)

		Response.write "<script>"
		Response.Write "alert('儲存完成！');"
		Response.write "</script>"
	End If 

	If trim(Request("DB_state"))="AcceptDel" Then


		If not ifnull(Request("DB_RecordMemberID2")) Then			
			chkwhere=" and to_char(RecordDate2,'YYYYMMDDHH24')='"&Request("DB_AcceptDate")&"' and RecordMemberID2="&trim(request("DB_RecordMemberID2"))

		else
			chkwhere=" and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)&" and RecordMemberID1="&trim(request("DB_RecordMemberID1"))
		End If 
		
		If trim(Request("DB_BillType")) = "BillRunCarAccept" Then
			strSQL="Delete BillRunCarAccept where billunitid='"&trim(Request("DB_BillUnitID"))&"' "&chkwhere
		else
			strSQL="Delete BillStopCarAccept where billunitid='"&trim(Request("DB_BillUnitID"))&"' "&chkwhere
		end if

		conn.execute(strSQL)

		Response.write "<script>"
		Response.Write "alert('刪除完成！');"
		Response.write "</script>"
	End If 


	If trim(Request("DB_state"))="AcceptRunBack" Then

'		Sys_Rule1=""
'		If not ifnull(trim(Request("DB_ObjRule"))) Then					
'			tmpRule=split(trim(Request("DB_ObjRule")),".")
'
'			If Ubound(tmpRule) >= 0 Then Sys_Rule1=tmpRule(0)
'
'			If Ubound(tmpRule) >= 1 Then
'				Sys_Rule1=Sys_Rule1&tmpRule(1)
'			else
'				Sys_Rule1=Sys_Rule1&"0"
'			end If 
'
'			If Ubound(tmpRule) >= 2 Then
'				Sys_Rule1=Sys_Rule1&right("00"&tmpRule(2),2)
'			else
'				Sys_Rule1=Sys_Rule1&"00"
'			end If 
'
'			Sys_Rule1=Sys_Rule1&"01"
'
'		End if


		updstr="Note='"&trim(Request("DB_ObjNote"))&"'"

		If not ifnull(Request("DB_chkState")) Then
			updstr=updstr&",RecordStateid=-1"
		End if

		If not ifnull(Request("DB_RecordMemberID2")) Then			
			chkwhere=" and to_char(RecordDate2,'YYYYMMDDHH24')='"&Request("DB_AcceptDate")&"' and RecordMemberID2="&trim(request("DB_RecordMemberID2"))

		else
			chkwhere=" and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)&" and RecordMemberID1="&trim(request("DB_RecordMemberID1"))
		End if
		
		strSQL="update BillRunCarAccept set "&updstr&" where billunitid='"&trim(Request("DB_BillUnitID"))&"' and CarNo='"&trim(Request("DB_ObjCarNo"))&"' and to_char(illegaldate,'HH24MI')='"&trim(Request("DB_ObjRule"))&"'"&chkwhere&" and RecordStateID=0"

		conn.execute(strSQL)

		Response.write "<script>"
		Response.Write "alert('儲存完成！');"
		Response.write "</script>"
	End if

	If trim(Request("DB_state"))="PrintAllOver" Then
		updstr=""
		If session("UnitLevelID") = "3" Then
			updstr="recordmemberid2 is null and billunitid in('"&trim(Session("Unit_ID"))&"')"	
		else
			updstr="recordmemberid2 is null and billunitid in(select unitid from unitinfo where unittypeid in(select unittypeid from unitinfo where unitid='"&trim(Session("Unit_ID"))&"'))"
		end If 
		
		If not ifnull(Request("DB_AcceptDate")) Then
			updstr=updstr&" and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)
		End If 

		strSQL="Update "&trim(Request("DB_BillType"))&" set RecordMemberID2="&Session("User_ID")&",RecordDate2=sysdate where "&updstr

		conn.execute(strSQL)

		Response.write "<script>"
		Response.Write "alert('設定完成！');"
		Response.write "</script>"
	End if

	If trim(Request("DB_state"))="Update" Then

		'str_Mem3="Sys_Mem3"&trim(Request("DB_BillUnitID"))&trim(Request("DB_AcceptDate"))

		'str_unit="Sys_"&trim(Request("DB_BillUnitID"))&trim(Request("DB_AcceptDate"))

		'str_memid="Sys_Chk"&trim(Request("DB_BillUnitID"))&trim(Request("DB_AcceptDate"))


		'If (not ifnull(Request(str_Mem3))) and not (ifnull(Request(str_unit))) Then
		'	If trim(Request("DB_BillType")) = "BillRunCarAccept" Then				
		'		strSQL="update "&trim(Request("DB_BillType"))&" set recordstateid=-1,Note='"&trim(Request(str_Mem3))&"' where carno='"&trim(Request(str_unit))&"' and billunitid='"&trim(Request("DB_BillUnitID"))&"' and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)
		'	else
		'		strSQL="update "&trim(Request("DB_BillType"))&" set recordstateid=-1,Note='"&trim(Request(str_Mem3))&"' where billno='"&trim(Request(str_unit))&"' and billunitid='"&trim(Request("DB_BillUnitID"))&"' and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)
		'	end if

			If not ifnull(Request("DB_RecordMemberID2")) Then
				updstr="COMPANYACCEPTDATE="&funGetDate(date,0)&",COMPANYMEMBERID="&Session("User_ID")
				chkwhere=" and to_char(RecordDate2,'YYYYMMDDHH24')='"&Request("DB_AcceptDate")&"' and RecordMemberID2="&trim(request("DB_RecordMemberID2"))

			else
				updstr="RecordMemberID2="&Session("User_ID")&",RecordDate2=sysdate,COMPANYACCEPTDATE="&funGetDate(date,0)&",COMPANYMEMBERID="&Session("User_ID")
				chkwhere=" and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)&" and RecordMemberID1="&trim(request("DB_RecordMemberID1"))
			End if
			
			strSQL="update "&trim(Request("DB_BillType"))&" set "&updstr&" where billunitid='"&trim(Request("DB_BillUnitID"))&"'"&chkwhere

			conn.execute(strSQL)

			Response.write "<script>"
			Response.Write "alert('儲存完成！');"
			Response.write "</script>"
		'end If
		
'		If not ifnull(Request(str_memid)) Then
'			strSQL="update "&trim(Request("DB_BillType"))&" set RecordMemberID3="&Session("User_ID")&" where billunitid='"&trim(Request("DB_BillUnitID"))&"' and AcceptDate="&funGetDate(gOutDT(Request("DB_AcceptDate")),0)

'			conn.execute(strSQL)

'			Response.write "<script>"
'			Response.Write "alert('儲存完成！');"
'			Response.write "</script>"
'		end if
	End if
	

	if DB_Selt="Selt" then
		where=""				
		If not ifnull(Request("Sys_Billno")) Then
			where=where&" and BillNo='"&trim(Request("Sys_Billno"))&"'"
		End If 
		
		If not ifnull(Request("Sys_CarNo")) Then
			where=where&" and CarNo='"&trim(Request("Sys_CarNo"))&"'"
		End If 
		
		If (not ifnull(Request("Sys_AcceptDate"))) and (not ifnull(Request("Sys_AcceptDate2"))) Then
			where=where&" and AcceptDate between "&funGetDate(gOutDT(trim(Request("Sys_AcceptDate"))),0)&" and "&funGetDate(gOutDT(trim(Request("Sys_AcceptDate2"))),0)
		End if

		If not ifnull(Request("Sys_BillUnitID")) Then
			sqlUit="select UnitLevelid from unitinfo where unitid in('"&trim(Request("Sys_BillUnitID"))&"')"

			set rsuit=conn.execute(sqlUit)
			If trim(rsuit("UnitLevelid"))="2" Then
				where=where&" and BillUnitID in(select UnitID from UnitInfo where UnitTypeid in('"&trim(Request("Sys_BillUnitID"))&"'))"
			else
				where=where&" and BillUnitID in('"&trim(Request("Sys_BillUnitID"))&"')"
			End if
			rsuit.close
		End if

		If not ifnull(Request("Sys_BillMem")) Then
			where=where&" and RecordMemberID1="&trim(Request("Sys_BillMem"))
		End If 
		
		If not ifnull(Request("chkAccept")) Then
			If trim(Request("chkAccept"))="1" Then
				where=where&" and RecordMemberID3 is null"

			elseIf trim(Request("chkAccept"))="2" Then
				where=where&" and RecordMemberID3 is not null"

			End if
		End if	

		If not ifnull(Request("chkPrint")) Then
			If trim(Request("chkPrint"))="1" Then
				where=where&" and RecordMemberID2 is null"

			elseIf trim(Request("chkPrint"))="2" Then
				where=where&" and RecordMemberID2 is not null"

			End if
		End if	
		
'		If trim(Session("UnitLevelID"))="1" Then

'			If not ifnull(Request("chkAccept")) Then
'				If trim(Request("chkAccept"))="1" Then
'					where=where&" and RecordMemberID2 is not null and RecordMemberID3 is null"

'				elseIf trim(Request("chkAccept"))="2" Then
'					where=where&" and RecordMemberID2 is not null and RecordMemberID3 is not null"

'				End if
'			End if	

'		else
'			If not ifnull(Request("chkAccept")) Then
'				If trim(Request("chkAccept"))="1" Then
'					where=where&" and RecordMemberID2 is null"

'				elseIf trim(Request("chkAccept"))="2" Then
'					where=where&" and RecordMemberID2 is not null"

'				End if
'			End if	
		
'		End if
			
	end If 
	
	If ifnull(DB_Selt) then
		If session("UnitLevelID") = "3" Then
			where=" and recordmemberid2 is null and billunitid in('"&trim(Session("Unit_ID"))&"')"	
		else
			where=" and recordmemberid2 is null and billunitid in(select unitid from unitinfo where unittypeid in(select unittypeid from unitinfo where unitid in('"&trim(Session("Unit_ID"))&"')))"
		end If 

		where=where&" and AcceptDate between sysdate-5 and sysdate"
	end if

	If not ifnull(Request("billtype")) Then
		If trim(Request("billtype"))="2" Then
			tablename="BillRunCarAccept"
		elseif trim(Request("billtype"))="1" Then
			tablename="BillStopCarAccept"
		else
			tablename="BillStopCarAccept"
		end if
	End If 
	
	Union=" Union all select 'BillRunCarAccept' BillType,a.AcceptDate,a.BillUnitID,a.RecordMemberID1,a.RecordMemberID2,a.COMPANYACCEPTDATE,a.suess,a.delss,b.UnitName,e.chname chname1,c.chname chname1,d.chname chname2 from (select DeCode(RecordMemberID2,null,to_char(AcceptDate,'YYYY/MM/DD'),to_char(RecordDate2,'YYYYMMDDHH24')) AcceptDate,BillUnitID,RecordMemberID1,recordmemberid2,recordmemberid3,COMPANYACCEPTDATE,sum(decode(recordstateid,0,1,0)) suess, sum(decode(recordstateid,-1,1,0)) delss from BillRunCarAccept where recordmemberid1 is not null"&where&" group by DeCode(RecordMemberID2,null,to_char(AcceptDate,'YYYY/MM/DD'),to_char(RecordDate2,'YYYYMMDDHH24')),BillUnitID,RecordMemberID1,recordmemberid2,recordmemberid3,COMPANYACCEPTDATE) a,UnitInfo b,memberdata c,memberdata d,Memberdata e where a.BillUnitID=b.UnitID and a.recordmemberid1=e.memberid and a.recordmemberid2=c.memberid(+) and a.recordmemberid3=d.memberid(+)"

	If not ifnull(Request("billtype")) Then
		If trim(Request("billtype"))="2" Then
			Union=""

		elseif trim(Request("billtype"))="1" Then
			Union=""

		end if
	End if

	strSQL="select '"&tablename&"' BillType,a.AcceptDate,a.BillUnitID,a.RecordMemberID1,a.RecordMemberID2,a.COMPANYACCEPTDATE,a.suess,a.delss,b.UnitName,e.chname chname1,c.chname chname2,d.chname chname3 from (select DeCode(RecordMemberID2,null,to_char(AcceptDate,'YYYY/MM/DD'),to_char(RecordDate2,'YYYYMMDDHH24')) AcceptDate,BillUnitID,RecordMemberID1,recordmemberid2,recordmemberid3,COMPANYACCEPTDATE,sum(decode(recordstateid,0,1,0)) suess, sum(decode(recordstateid,-1,1,0)) delss from "&tablename&" where recordmemberid1 is not null"&where&" group by DeCode(RecordMemberID2,null,to_char(AcceptDate,'YYYY/MM/DD'),to_char(RecordDate2,'YYYYMMDDHH24')),BillUnitID,RecordMemberID1,recordmemberid2,recordmemberid3,COMPANYACCEPTDATE) a,UnitInfo b,memberdata c,memberdata d,Memberdata e where a.BillUnitID=b.UnitID and a.recordmemberid1=e.memberid and a.recordmemberid2=c.memberid(+) and a.recordmemberid3=d.memberid(+)"&Union

	set rsfild=conn.execute(strSQL&" order by billunitid,AcceptDate")

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
					<td>單號</td>
					<td>
						<input name="Sys_Billno" class="btn1" type="text" value="<%=request("Sys_Billno")%>" onkeyup="UpperCase(this);" size="12" maxlength="18">
					</td>
					<td>點收類別</td>
					<td>
						<select name="billtype">
							<option value="">全部</option>
							<option value="1"<%If trim(Request("billtype"))="1" Then Response.Write " selected"%>>攔停</option>
							<option value="2"<%If trim(Request("billtype"))="2" Then Response.Write " selected"%>>逕舉</option>
						</select>
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
					<td>建檔人員</td>
					<td>
						<%=UnSelectMemberOption("Sys_BillUnitID","Sys_BillMem")%>
					</td>
				</tr>
				<tr>
					<td>車號</td>
					<td>
						<input name="Sys_CarNo" class="btn1" type="text" value="<%=request("Sys_CarNo")%>" onkeyup="UpperCase(this);" size="12" maxlength="18">
					</td>
					<td>上傳日期</td>
					<td>
						<input name="Sys_AcceptDate" class="btn1" type="text" value="<%=request("Sys_AcceptDate")%>" size="5" maxlength="7">
						<input class="btn3" style="width:20px;height:20px;font-size:14px;" type="button" name="datestr" value="..." onclick="OpenWindow('Sys_AcceptDate');">
						～
						<input name="Sys_AcceptDate2" class="btn1" type="text" value="<%=request("Sys_AcceptDate2")%>" size="5" maxlength="7">
						<input class="btn3" style="width:20px;height:20px;font-size:14px;" type="button" name="datestr" value="..." onclick="OpenWindow('Sys_AcceptDate2');">
					</td>

					<td>審核狀態</td>
					<td>						
						<select name="chkAccept">
							<option value="99"<%If trim(Request("chkAccept"))="99" Then Response.Write " selected"%>>全部</option>
							<option value="1"<%If trim(Request("chkAccept"))="1" Then Response.Write " selected"%>>未審核</option>
							<option value="2"<%If trim(Request("chkAccept"))="2" Then Response.Write " selected"%>>已審核</option>							
						</select>
					</td>

					<td>列印狀態</td>
					<td>						
						<select name="chkPrint">
							<option value="99"<%If trim(Request("chkPrint"))="99" Then Response.Write " selected"%>>全部</option>
							<option value="1"<%If trim(Request("chkPrint"))="1" Then Response.Write " selected"%>>未列印</option>
							<option value="2"<%If trim(Request("chkPrint"))="2" Then Response.Write " selected"%>>已列印</option>							
						</select>
						<input class="btn3" style="width:40px;height:25px;font-size:14px;" type="submit" name="btnSelt" value="查詢" onClick='funSelt();'>&nbsp;&nbsp;
						<input class="btn3" style="width:40px;height:25px;font-size:14px;" type="button" name="cancel" value="清除" onClick="location='BillBaseCheckAcceptQuery_miaoli.asp'">
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
					<th>類別</th>
					<th>上傳日</th>
					<th>舉發單位</th>
					<th>建檔人</th>
					<th>審核人</th>
					<th>列印人</th>
					<th>已點收</th>					
					<th>已退件</th>
					<th>操作</th>
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

						If trim(rsfild("BillType")) = "BillRunCarAccept" Then
							response.write "<td align='center' nowrap>逕舉</td>"
						else
							response.write "<td align='center' nowrap>攔停</td>"
						end If 
						
						strAcceptDate=""

						If not ifnull(rsfild("recordmemberid2")) Then
							Response.Write "<td width=""95"">"&(left(rsfild("AcceptDate"),4)-1911)&Mid(rsfild("AcceptDate"),5,4)&"</th>"
							strAcceptDate=rsfild("AcceptDate")

						else
							Response.Write "<td width=""95"">"&gInitDT(rsfild("AcceptDate"))&"</th>"
							strAcceptDate=gInitDT(rsfild("AcceptDate"))

						End If  

						response.write "<td align='center' nowrap>"&rsfild("UnitName")&"</td>"
						response.write "<td align='center' nowrap>"&trim(rsfild("chname1"))&"</td>"
						sys_chname=""
'						sys_chname=trim(rsfild("chname2"))
'						If not ifnull(rsfild("chname3")) Then
'							If not ifnull(sys_chname) Then sys_chname=sys_chname&","
							sys_chname=sys_chname&trim(rsfild("chname3"))
'						End if
						
						response.write "<td align='center' nowrap>"&sys_chname&"</td>"
						response.write "<td align='center' nowrap>"&trim(rsfild("chname2"))&"</td>"
						response.write "<td align='center' nowrap>"&rsfild("suess")&"</td>"
						response.write "<td align='center' nowrap>"&rsfild("delss")&"</td>"

						selt_BackChk="selt_BackChk"&trim(rsfild("billunitid"))&strAcceptDate&trim(rsfild("RecordMemberID1"))&trim(rsfild("BillType"))
						selt_BackCarNo="selt_BackCarNo"&trim(rsfild("billunitid"))&strAcceptDate&trim(rsfild("RecordMemberID1"))&trim(rsfild("BillType"))
						selt_BackRule="selt_BackRule"&trim(rsfild("billunitid"))&strAcceptDate&trim(rsfild("RecordMemberID1"))&trim(rsfild("BillType"))
						selt_Note="selt_Note"&trim(rsfild("billunitid"))&strAcceptDate&trim(rsfild("RecordMemberID1"))&trim(rsfild("BillType"))
						chk_memid="selt_Chk"&trim(rsfild("billunitid"))&strAcceptDate&trim(rsfild("RecordMemberID1"))&trim(rsfild("BillType"))

						response.write "<td align='left' nowrap>"

						If trim(Session("UnitLevelID"))=1 and trim(Session("Unit_ID"))="03BA" then							
							
							Response.Write "<input class='btn1' type='checkbox' name="""&chk_memid&""" value='1'>簽收　"
							
							Response.Write "<input type=""button"" name=""Update"" value=""確定"" class=""btn3"" style=""width:40px;height:25px;font-size:12px;"" onclick=""funAcceptCreat('"&strAcceptDate&"','"&rsfild("BillUnitID")&"','"&trim(rsfild("BillType"))&"','"&trim(rsfild("RecordMemberID1"))&"','"&trim(rsfild("recordmemberid2"))&"',myForm."&chk_memid&");"">"							
							
						end If 
						
						Response.Write "&nbsp;<input type=""button"" name=""Update"" value=""詳細"" class=""btn3"" style=""width:40px;height:25px;font-size:12px;"" onclick=""funAcceptLoad('"&strAcceptDate&"','"&rsfild("BillUnitID")&"','"&trim(rsfild("BillType"))&"','"&trim(rsfild("RecordMemberID1"))&"','"&trim(rsfild("RecordMemberID2"))&"');"">"
						
						If cdbl(rsfild("delss")) > 0 Then
							Response.Write "&nbsp;<input type=""button"" name=""Update"" value=""退件清冊"" class=""btn3"" style=""width:80px;height:25px;font-size:12px;"" onclick=""funAcceptBackLoad('"&strAcceptDate&"','"&rsfild("BillUnitID")&"','"&trim(rsfild("BillType"))&"','"&trim(rsfild("RecordMemberID1"))&"','"&trim(rsfild("RecordMemberID2"))&"');"">"
						End If 

						Response.Write "<br>"

						If cdbl(Session("UnitLevelID"))<3 then
							If trim(rsfild("BillType")) = "BillRunCarAccept" Then

								Response.Write "<input class='btn1' type='checkbox' name="""&selt_BackChk&""" value='1'>退件　"

								Response.Write "退件車號:"

								Response.Write "<input type=""text"" name="""&selt_BackCarNo&""" ID="""&selt_BackCarNo&""" class=""btn1"" size=""6"" value="""">"

								Response.Write "違規時間:"

								Response.Write "<input type=""text"" name="""&selt_BackRule&""" ID="""&selt_BackRule&""" class=""btn1"" size=""5"" value="""">"

								Response.Write "<br>"

								Response.Write "原因:"

								Response.Write "<input type=""text"" name="""&selt_Note&""" ID="""&selt_Note&""" class=""btn1"" size=""30"" value="""">"

								Response.Write "<input type=""button"" name=""Update"" value=""確定"" class=""btn3"" style=""width:40px;height:25px;font-size:12px;"" onclick=""funAcceptRunBack('"&strAcceptDate&"','"&rsfild("BillUnitID")&"','"&trim(rsfild("RecordMemberID1"))&"','"&trim(rsfild("recordmemberid2"))&"',myForm."&selt_BackChk&",myForm."&selt_BackCarNo&",myForm."&selt_BackRule&",myForm."&selt_Note&");"">"
							else
								
								Response.Write "<input class='btn1' type='checkbox' name="""&selt_BackChk&""" value='1'>退件　"

								Response.Write "退件單號:"

								Response.Write "<input type=""text"" name="""&selt_BackCarNo&""" ID="""&selt_BackCarNo&""" class=""btn1"" size=""12"" value="""">"

								Response.Write "<br>"

								Response.Write "原因:"

								Response.Write "<input type=""text"" name="""&selt_Note&""" ID="""&selt_Note&""" class=""btn1"" size=""30"" value="""">"

								Response.Write "<input type=""button"" name=""Update"" value=""確定"" class=""btn3"" style=""width:40px;height:25px;font-size:12px;"" onclick=""funAcceptStopBack('"&strAcceptDate&"','"&rsfild("BillUnitID")&"','"&trim(rsfild("RecordMemberID1"))&"','"&trim(rsfild("recordmemberid2"))&"',myForm."&selt_BackChk&",myForm."&selt_BackCarNo&",myForm."&selt_Note&");"">"
							End If 
						end if
						If trim(Session("Credit_ID"))="A000000000" Then
							Response.Write "&nbsp;<input type=""button"" name=""Update"" value=""刪除"" class=""btn3"" style=""width:40px;height:25px;font-size:12px;"" onclick=""funAcceptDel('"&strAcceptDate&"','"&rsfild("BillUnitID")&"','"&trim(rsfild("BillType"))&"','"&trim(rsfild("RecordMemberID1"))&"','"&trim(rsfild("RecordMemberID2"))&"');"">"
						End if 
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
			<input type="button" name="cancel" class="btn3" style="width:100px;height:25px;font-size:16px;" value="轉成Excel"  onclick="funAcceptExcel();">
			<br>
			<input type="button" name="cancel" class="btn3" style="width:100px;height:25px;font-size:16px;" value="合併列印"  onclick="funAcceptAllLoad();"<%
				If trim(Request("chkPrint"))="2" or trim(Request("chkPrint"))="99" or trim(request("Sys_AcceptDate2"))<>"" or trim(request("billtype"))="" then Response.Write " disabled"
			%>>
			<input type="button" name="cancel" class="btn3" style="width:100px;height:25px;font-size:16px;" value="完成列印"  onclick="funPrintAllOver();"<%
				If trim(Request("chkPrint"))="2" or trim(Request("chkPrint"))="99" or trim(request("Sys_AcceptDate2"))<>"" or trim(request("billtype"))="" then Response.Write " disabled"
			%>>
		</td>
	</tr>
</table>
<center>
	<a href="../Report/Report0010.asp"><span class="pagetitle">舉發件數統計表(員警別明細)</span></a>
	<a href="../Report/Report0011.asp"><span class="pagetitle">舉發件數統計表(員警別總計)</span></a>
	<a href="../Report/Report0024.asp"><span class="pagetitle">舉發件數統計表（單位）</span></a>
</center>

<input type="Hidden" name="DB_Selt" value="<%=DB_Selt%>">
<input type="Hidden" name="DB_state" value="">
<input type="Hidden" name="DB_AcceptDate" value="">
<input type="Hidden" name="DB_BillUnitID" value="">
<input type="Hidden" name="DB_RecordMemberID1" value="">
<input type="Hidden" name="DB_RecordMemberID2" value="">
<input type="Hidden" name="DB_BillType" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">

<input type="Hidden" name="DB_ObjCarNo" value="">
<input type="Hidden" name="DB_ObjRule" value="">
<input type="Hidden" name="DB_ObjNote" value="">
<input type="Hidden" name="DB_chkState" value="">
</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">
<%="UnitMan('Sys_BillUnitID','Sys_BillMem','"&request("Sys_BillMem")&"');"%>
function ListItemID(chkMemID,UnitListName,MemListName){
	runServerScript("/traffic/Common/ListItem.asp?LoginID="+document.all[chkMemID].value+"&UnitListName="+UnitListName+"&MemListName="+MemListName);
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


function funAcceptAllLoad(){
	myForm.DB_BillType.value="";
	myForm.DB_AcceptDate.value="";
	myForm.DB_BillUnitID.value="";
	myForm.DB_RecordMemberID1.value="";

	if(myForm.billtype.value=='2'){
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

function funAcceptExcel(){
	if(myForm.DB_Selt.value=='Selt'){
		myForm.action="BillBaseCheckAcceptQueryExcel_miaoli.asp";
		myForm.target="PrintAcceptExcel";
		myForm.submit();
		myForm.action="";
		myForm.target="";
	}else{
		alert("請先進行查詢！！");
	}
}

function funPrintAllOver(){
	myForm.DB_AcceptDate.value=myForm.Sys_AcceptDate.value;
	myForm.DB_BillUnitID.value="";
	myForm.DB_RecordMemberID1.value="";

	if(myForm.billtype.value=="1"){
		myForm.DB_BillType.value="BILLSTOPCARACCEPT";
	}else{
		myForm.DB_BillType.value="BILLRUNCARACCEPT";
	}

	myForm.DB_state.value="PrintAllOver";
	myForm.submit();
}

function funAcceptCreat(AcceptDate,BillUnitID,BillType,RecordMemberID1,RecordMemberID2,object){
	if (object.checked==true){
		myForm.DB_AcceptDate.value=AcceptDate;
		myForm.DB_BillUnitID.value=BillUnitID;
		myForm.DB_BillType.value=BillType;
		myForm.DB_RecordMemberID1.value=RecordMemberID1;
		myForm.DB_RecordMemberID2.value=RecordMemberID2;
		myForm.DB_state.value="Update";
		myForm.submit();
	}
}

function funAcceptRunBack(AcceptDate,BillUnitID,RecordMemberID1,RecordMemberID2,object,ObjCarNo,ObjRule,ObjNote){

	if(ObjCarNo.value!=""&&ObjRule.value!=""&&ObjNote.value!=""){

		myForm.DB_chkState.value="";

		if (object.checked==true){
			myForm.DB_chkState.value="1";
		}

		myForm.DB_AcceptDate.value=AcceptDate;
		myForm.DB_BillUnitID.value=BillUnitID;
		myForm.DB_BillType.value="";
		myForm.DB_RecordMemberID1.value=RecordMemberID1;
		myForm.DB_RecordMemberID2.value=RecordMemberID2;

		myForm.DB_ObjCarNo.value=ObjCarNo.value;
		myForm.DB_ObjRule.value=ObjRule.value;
		myForm.DB_ObjNote.value=ObjNote.value;

		myForm.DB_state.value="AcceptRunBack";
		myForm.submit();
	}
}

function funAcceptStopBack(AcceptDate,BillUnitID,RecordMemberID1,RecordMemberID2,object,ObjCarNo,ObjNote){
	if(ObjCarNo.value!=""&&ObjNote.value!=""){
		myForm.DB_chkState.value="";

		if (object.checked==true){
			myForm.DB_chkState.value="1";
		}

		myForm.DB_AcceptDate.value=AcceptDate;
		myForm.DB_BillUnitID.value=BillUnitID;
		myForm.DB_BillType.value="";
		myForm.DB_RecordMemberID1.value=RecordMemberID1;
		myForm.DB_RecordMemberID2.value=RecordMemberID2;

		myForm.DB_ObjCarNo.value=ObjCarNo.value;
		myForm.DB_ObjRule.value="";
		myForm.DB_ObjNote.value=ObjNote.value;

		myForm.DB_state.value="AcceptStopBack";
		myForm.submit();
	}
}

function funAcceptDel(AcceptDate,BillUnitID,BillType,RecordMemberID1,RecordMemberID2){
	myForm.DB_BillType.value=BillType;
	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordMemberID1.value=RecordMemberID1;
	myForm.DB_RecordMemberID2.value=RecordMemberID2;

	myForm.DB_state.value="AcceptDel";
	myForm.submit();
}

function funAcceptLoad(AcceptDate,BillUnitID,BillType,RecordMemberID1,RecordMemberID2){
	myForm.DB_BillType.value="";
	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordMemberID1.value=RecordMemberID1;
	myForm.DB_RecordMemberID2.value=RecordMemberID2;

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

function funAcceptBackLoad(AcceptDate,BillUnitID,BillType,RecordMemberID1,RecordMemberID2){
	myForm.DB_BillType.value="";
	myForm.DB_AcceptDate.value=AcceptDate;
	myForm.DB_BillUnitID.value=BillUnitID;
	myForm.DB_RecordMemberID1.value=RecordMemberID1;
	myForm.DB_RecordMemberID2.value=RecordMemberID2;

	if(BillType=='BillRunCarAccept'){
		UrlStr="AcceptRunBackList.asp";
	}else{
		UrlStr="AcceptStopBackList.asp";
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

function UpperCase(obj){
	if(obj.value!=obj.value.toUpperCase()){
		obj.value=obj.value.toUpperCase();
	}
}

</script>
<%conn.close%>