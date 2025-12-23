<!--#include virtual="traffic/Common/Login_Check.asp"-->
<!--#include virtual="/traffic/Common/db.ini"-->
<!-- #include file="../Common/AllFunction.inc"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">

<html>
<script language="JavaScript">
	window.focus();
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<!--#include virtual="/traffic/Common/css.txt"-->
<style media=print>
.Noprint{display:none;}
.PageNext{page-break-after: always;}
</style>
<style type="text/css">
<!--
.style1 {
	font-size:11pt; 
	font-weight: bold;
	font-family: "標楷體";
}
.style2 {
	font-size:11pt; 
}
.style3 {
	font-size:11pt; 
	font-weight: bold;
}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:20px;
	font-family: "標楷體";
}
-->
</style>
<title>舉發單綜合查詢</title>
<script type="text/javascript" src="../js/Print.js"></script>
<script type="text/javascript" src="../js/date.js"></script>
<% Server.ScriptTimeout = 800 %>
<%	

	'抓縣市
	strCity="select value from Apconfigure where id=31"
	set rsCity=conn.execute(strCity)
		sys_City=trim(rsCity("value"))
	rsCity.close
	set rsCity=nothing

	strSQLTemp=""
	strQry="[快速查詢]"
	if trim(request("BillNo"))<>"" Then
		If strQry="[快速查詢]" then
			strQry=strQry&"BillNo="&Trim(request("BillNo"))
		Else
			strQry=strQry&",BillNo="&Trim(request("BillNo"))
		End if
		strSQLTemp=strSQLTemp&" and BillNo='"&trim(request("BillNo"))&"'"
	end if
	if trim(request("CarNo"))<>"" Then
		If strQry="[快速查詢]" then
			strQry=strQry&"CarNo="&Trim(request("CarNo"))
		Else
			strQry=strQry&",CarNo="&Trim(request("CarNo"))
		End if
		strSQLTemp=strSQLTemp&" and CarNo='"&trim(request("CarNo"))&"'"
	end if


	strSQL="Select * from BillBase " &_
		" where RecordStateID=0 "&strSQLTemp

%>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%	Cnt=0
	NoCase=0
	tdstopflag=0
	'on error resume next  
	ConnExecute strQry&",原因:"&trim(request("QryReason2")),356
	set rs1=conn.execute(strSQL)
	If Not rs1.Bof Then
		rs1.MoveFirst 
	else
		NoCase=1
		response.write "查無資料!!"
	end if
	While Not rs1.Eof
	if Cnt>0 then
%>
<div class="PageNext"></div>
<%	end if

%>
<form name=myForm method="post">
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td align="center">
				<span class="style6">舉發違反交通管理事件通知單</span>
			</td>
		</tr>
		<tr>
			<td><span class="style2">製表單位：</span><span class="style1"><%
			strUnit="select UnitName from UnitInfo where UnitID='"&trim(session("Unit_ID"))&"'"
			set rsUnit=conn.execute(strUnit)
			if not rsUnit.eof then
				response.write trim(rsUnit("UnitName"))
			end if
			rsUnit.close
			set rsUnit=nothing
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">操作人：</span><span class="style1"><%
			strMem="select ChName from MemberData where MemberID='"&trim(session("User_ID"))&"'"
			set rsMem=conn.execute(strMem)
			if not rsMem.eof then
				response.write trim(rsMem("ChName"))
			end if
			rsMem.close
			set rsMem=nothing
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">製表時間：</span><span class="style3"><%=now%></span></td>
		</tr>
	</table>
	<hr>
	<table width='100%' border='0' cellpadding="2">
		<tr>
			<td width="25%"><span class="style2">告發單號：</span><span class="style1"><%
			if trim(rs1("BillNO"))<>"" and not isnull(rs1("BillNO")) then
				response.write trim(rs1("BillNO"))
			end if
			%></span></td>
			<!-- <td width="27%"><span class="style2">到案處所：</span><span class="style1"><%
				strStation="select * from Station where DciStationID='"&trim(rs1("MemberStation"))&"'"
				set rsStation=conn.execute(strStation)
				if not rsStation.eof then
					response.write trim(rsStation("DCIStationName"))
				end if
				rsStation.close
				set rsStation=nothing
			%></span></td> -->
			<td width="23%"><span class="style2">告發類別：</span><span class="style1"><%
			if trim(rs1("BillTypeID"))="2" then
				response.write "逕舉"
			else
				response.write "攔停"
			end if
			%></span></td>
			<td width="25%"><span class="style2">舉發單狀態：</span><span class="style1"><%
			if trim(rs1("RecordStateID"))="-1" then
				response.write "<font color=""red"">已刪除</font>"
			else
				response.write "正常"
			end if
			'刪除原因
			if trim(rs1("RecordStateID"))="-1" or sys_City="台中市" or trim(Session("Credit_ID"))="A000000000" then
				strDelRea="select b.Content from BillDeleteReason a,DciCode b where a.BillSn="&trim(rs1("Sn"))&" and b.TypeID=3 and a.DelReason=b.ID"
				set rsDelRea=conn.execute(strDelRea)
				if not rsDelRea.eof then
					response.write "<font color=""red"">." & trim(rsDelRea("Content")) & "</font>"
				else
					response.write "&nbsp;"
				end if
				rsDelRea.close
				set rsDelRea=nothing
			end if
			if trim(rs1("RecordStateID"))="-1" and (sys_City="高雄市" or sys_City=ApconfigureCityName) then
				strDelTime="select * from log where typeid=352 and ActionContent like '%單號:"&trim(rs1("BillNo"))&"%' and ActionContent like '%車號:"&trim(rs1("CarNo"))&"%' and rownum<=1 order by ActionDate Desc"
				set rsDelTime=conn.execute(strDelTime)
				if not rsDelTime.eof then
					response.write "<font color=""red"">."&year(rsDelTime("ActionDate"))-1911&"/"&month(rsDelTime("ActionDate"))&"/"&day(rsDelTime("ActionDate"))&" "&hour(rsDelTime("ActionDate"))&":"&minute(rsDelTime("ActionDate"))&"</font>"
				end if
				rsDelTime.close
				set rsDelTime=nothing
			end if
			%></span></td>
		</tr>
		<tr>
			<td colspan="2"><span class="style2">違規時間：</span><span class="style1"><%
			if trim(rs1("IllegalDate"))<>"" and not isnull(rs1("IllegalDate")) then
				response.write gArrDT(trim(rs1("IllegalDate")))&"&nbsp;"
				response.write Right("00"&hour(rs1("IllegalDate")),2)&":"
				response.write Right("00"&minute(rs1("IllegalDate")),2)
			end if		
			%></span></td>
			<td colspan="2"><span class="style2">舉發員警：</span><span class="style1"><%
			if trim(rs1("BillMem1"))<>"" and not isnull(rs1("BillMem1")) then
				response.write trim(rs1("BillMem1"))
				strMem1="select LoginID from MemberData where memberId="&trim(rs1("BillMemID1"))
				set rsMem1=conn.execute(strMem1)
				if not rsMem1.eof then
					response.write "("&trim(rsMem1("LoginID"))&")"
				end if
				rsMem1.close
				set rsMem1=nothing
			end if	
			if trim(rs1("BillMem2"))<>"" and not isnull(rs1("BillMem2")) then
				response.write "/&nbsp;"&trim(rs1("BillMem2"))
				strMem2="select LoginID from MemberData where memberId="&trim(rs1("BillMemID2"))
				set rsMem2=conn.execute(strMem2)
				if not rsMem2.eof then
					response.write "("&trim(rsMem2("LoginID"))&")"
				end if
				rsMem2.close
				set rsMem2=nothing
			end if	
			if trim(rs1("BillMem3"))<>"" and not isnull(rs1("BillMem3")) then
				response.write "/&nbsp;"&trim(rs1("BillMem3"))
				strMem3="select LoginID from MemberData where memberId="&trim(rs1("BillMemID3"))
				set rsMem3=conn.execute(strMem3)
				if not rsMem3.eof then
					response.write "("&trim(rsMem3("LoginID"))&")"
				end if
				rsMem3.close
				set rsMem3=nothing
			end if	
			if trim(rs1("BillMem4"))<>"" and not isnull(rs1("BillMem4")) then
				response.write "/&nbsp;"&trim(rs1("BillMem4"))
				strMem4="select LoginID from MemberData where memberId="&trim(rs1("BillMemID4"))
				set rsMem4=conn.execute(strMem4)
				if not rsMem4.eof then
					response.write "("&trim(rsMem4("LoginID"))&")"
				end if
				rsMem4.close
				set rsMem4=nothing
			end if	
			%></span></td>
		</tr>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule1")),2)="40" or (int(rs1("Rule1"))>4310200 and int(rs1("Rule1"))<4310209) or (int(rs1("Rule1"))>3310100 and int(rs1("Rule1"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule1"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule1"))<>"" and not isnull(rs1("Rule1")) then
					if left(trim(rs1("Rule1")),4)="2110" or left(trim(rs1("Rule1")),4)="2210" or trim(rs1("Rule1"))="4310102" or trim(rs1("Rule1"))="4310103" or trim(rs1("Rule1"))="4310104" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple=" and CarSimpleID in ('3','0')"
						else
							strCarImple=""
						end if
					end if
					strR1="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule1"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple&" order by CarSimpleID Desc"
					set rsR1=conn.execute(strR1)
					if not rsR1.eof then 
						response.write trim(rs1("Rule1"))&" "&trim(rsR1("IllegalRule"))
					end if
					rsR1.close
					set rsR1=nothing

					if trim(rs1("BillTypeID"))="2" and trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
						response.write "&nbsp;"&trim(rs1("Rule4"))
					end if
				end if	
			end If
			%></span></td>
		</tr>
<%if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule2")),2)="40" or (int(rs1("Rule2"))>4310200 and int(rs1("Rule2"))<4310209) or (int(rs1("Rule2"))>3310100 and int(rs1("Rule2"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule2"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule2"))<>"" and not isnull(rs1("Rule2")) then
					if left(trim(rs1("Rule2")),4)="2110" or left(trim(rs1("Rule1")),4)="2210" or trim(rs1("Rule2"))="4310102" or trim(rs1("Rule2"))="4310103" or trim(rs1("Rule2"))="4310104" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple2=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple2=" and CarSimpleID in ('3','0')"
						else
							strCarImple2=""
						end if
					end if
					strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule2"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
					set rsR2=conn.execute(strR2)
					if not rsR2.eof then 
						response.write trim(rs1("Rule2"))&" "&trim(rsR2("IllegalRule"))
					end if
					rsR2.close
					set rsR2=nothing
				end if
			end If
			%></span></td>
		</tr>
<%end if%>
<%if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule3")),2)="40" or (int(rs1("Rule3"))>4310200 and int(rs1("Rule3"))<4310209) or (int(rs1("Rule3"))>3310100 and int(rs1("Rule3"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule3"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule3"))<>"" and not isnull(rs1("Rule3")) then
					if left(trim(rs1("Rule3")),4)="2110" or left(trim(rs1("Rule1")),4)="2210" or trim(rs1("Rule3"))="4310102" or trim(rs1("Rule3"))="4310103" or trim(rs1("Rule3"))="4310104" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple2=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple2=" and CarSimpleID in ('3','0')"
						else
							strCarImple2=""
						end if
					end if
					strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule3"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
					set rsR2=conn.execute(strR2)
					if not rsR2.eof then 
						response.write trim(rs1("Rule3"))&" "&trim(rsR2("IllegalRule"))
					end if
					rsR2.close
					set rsR2=nothing
				end if	
			end If
			%></span></td>
		</tr>
<%end if%>
<%if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) and trim(rs1("BillTypeID"))<>"2" then%>
		<tr>
			<td colspan="5"><span class="style2">違反法條：</span><span class="style1"><%
			if (left(trim(rs1("Rule4")),2)="40" or (int(rs1("Rule4"))>4310200 and int(rs1("Rule4"))<4310209) or (int(rs1("Rule4"))>3310100 and int(rs1("Rule4"))<3310111)) and sys_City="基隆市" then
				response.write trim(rs1("Rule4"))&" "&"該路段限速"&trim(rs1("RuleSpeed"))&"公里、經雷達測速為"&trim(rs1("IllegalSpeed"))&"公里、超速"&cint(rs1("IllegalSpeed"))-cint(rs1("RuleSpeed"))&"公里"
			else
				if trim(rs1("Rule4"))<>"" and not isnull(rs1("Rule4")) then
					if left(trim(rs1("Rule4")),4)="2110" or left(trim(rs1("Rule1")),4)="2210" or trim(rs1("Rule4"))="4310102" or trim(rs1("Rule4"))="4310103" or trim(rs1("Rule4"))="4310104" then
						if trim(rs1("CarSimpleID"))=1 or trim(rs1("CarSimpleID"))=2 then
							strCarImple2=" and CarSimpleID in ('5','0')"
						elseif trim(rs1("CarSimpleID"))=3 or trim(rs1("CarSimpleID"))=4 then
							strCarImple2=" and CarSimpleID in ('3','0')"
						else
							strCarImple2=""
						end if
					end if
					strR2="select IllegalRule,Level1 from Law where ItemID='"&trim(rs1("Rule4"))&"' and Version='"&trim(rs1("RuleVer"))&"'"&strCarImple2&" order by CarSimpleID Desc"
					set rsR2=conn.execute(strR2)
					if not rsR2.eof then 
						response.write trim(rs1("Rule4"))&" "&trim(rsR2("IllegalRule"))
					end if
					rsR2.close
					set rsR2=nothing
				end if	
			end If
			%></span></td>
		</tr>
<%end if%>
		<tr>
			<td colspan="<%
				response.write "3"
			%>"><span class="style2">違規路段：</span><span class="style1"><%
			response.write trim(rs1("IllegalAddressID"))&" "&trim(rs1("IllegalAddress"))
			%></span></td>
			<td><span class="style2">是否郵寄：</span><span class="style1"><%
			if trim(rs1("EquipMentID"))<>"" and not isnull(rs1("EquipMentID")) then
				if trim(rs1("EquipMentID"))="1" then
					response.write "是"
				else
					response.write "否"
				end if
			end if	
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">車號：</span><span class="style1"><font size="4"><%=trim(rs1("CarNo"))%></font></span></td>
		</tr>

		<tr>
			<td><span class="style2">填單日期：</span><span class="style1"><%
			if trim(rs1("BillFillDate"))<>"" and not isnull(rs1("BillFillDate")) then
				response.write gArrDT(trim(rs1("BillFillDate")))
			end if	
			%></span></td>
			<td colspan="3"><span class="style2">舉發單位：</span><span class="style1"><%
			if trim(rs1("BillUnitID"))<>"" and not isnull(rs1("BillUnitID")) then
				response.write trim(rs1("BillUnitID"))&"&nbsp;"
				strBillUnit="select UnitName from UnitInfo where UnitID='"&trim(rs1("BillUnitID"))&"'"
				set rsBillUnit=conn.execute(strBillUnit)
				if not rsBillUnit.eof then
					response.write trim(rsBillUnit("UnitName"))
				end if
				rsBillUnit.close
				set rsBillUnit=nothing
			end if	
			%></span></td>
		</tr>
		<tr>
			<td><span class="style2">到案日期：</span><span class="style1"><%
			if trim(rs1("DealLineDate"))<>"" and not isnull(rs1("DealLineDate")) then
				response.write gArrDT(trim(rs1("DealLineDate")))
			end if	
			%></span></td>
			<td><span class="style2">簡式車種：</span><span class="style1"><%
			if trim(rs1("CarSimpleID"))<>"" and not isnull(rs1("CarSimpleID")) then
				if trim(rs1("CarSimpleID"))="1" then
					response.write "汽車"
				elseif trim(rs1("CarSimpleID"))="2" then
					response.write "拖車"
				elseif trim(rs1("CarSimpleID"))="3" then
					response.write "重機"
				elseif trim(rs1("CarSimpleID"))="4" then
					response.write "輕機"
				elseif trim(rs1("CarSimpleID"))="6" then
					response.write "簡式車種"
				end if
			end if	
			%></span></td>
			<td><span class="style2">建檔日期：</span><span class="style1"><%
			if trim(rs1("RecordDate"))<>"" and not isnull(rs1("RecordDate")) then
				response.write gArrDT(trim(rs1("RecordDate")))
			end if	
			%></span></td>
			<td><span class="style2">操作人員：</span><span class="style1"><%
			strRecMem="select ChName from MemberData where MemberID='"&trim(rs1("RecordMemberID"))&"'"
			set rsRecMem=conn.execute(strRecMem)
			if not rsRecMem.eof then
				response.write trim(rsRecMem("ChName"))
			end if
			rsRecMem.close
			set rsRecMem=nothing
			%></span></td>
		</tr>
		<tr>
			<td>
				<span class="style2">備註：</span><span class="style1"><%=trim(rs1("Note"))%></span>
			</td>
		</tr>
<%
		strDSupd="select * from DCISTATUSUPDATE where Billsn="&Trim(rs1("Sn"))
		Set rsDSupd=conn.execute(strDSupd)
		If Not rsDSupd.eof Then
		%>
				<tr>
				<td colspan="3">
					<span class="style2">強制入案前狀態：</span><span class="style1"><%
				strDS1="select * from Dcireturnstatus where DciActionID='W' " &_
					" and DciReturn='"&Trim(rsDSupd("StatUS"))&"'"
				Set rsDS1=conn.execute(strDS1)
				If Not rsDS1.eof Then
					response.write rsDS1("StatusContent")
				End If
				rsDS1.close
				Set rsDS1=Nothing
				strDS2="select * from Dcireturnstatus where DciActionID='WE' " &_
					" and DciReturn='"&Trim(rsDSupd("DciErrorCarData"))&"'"
				Set rsDS2=conn.execute(strDS2)
				If Not rsDS2.eof Then
					response.write " "&rsDS2("StatusContent")
				End If
				rsDS2.close
				Set rsDS2=Nothing
				response.write " "&rsDSupd("RecordDate")
					%></span>
				</td>
				</tr>
		<%
		End If
		rsDSupd.close
		Set rsDSupd=nothing
		%>			
		
	<%if sys_City="高雄縣" or (sys_City="高雄市" or sys_City=ApconfigureCityName) then%>
		<tr>
			<td colspan="4">
			<span class="style2">代保管物：</span><span class="style1"><%
			strFast="select a.FASTENERTYPEID,b.Content from BILLFASTENERDETAIL a,DciCode b" &_
				" where a.FASTENERTYPEID=b.ID and b.TypeID=6 and a.BillSn="&trim(rs1("SN"))
			set rsFast=conn.execute(strFast)
			If Not rsFast.Bof Then rsFast.MoveFirst 
			While Not rsFast.Eof
					response.write trim(rsFast("FASTENERTYPEID"))&trim(rsFast("Content"))&" "
				rsFast.MoveNext
			Wend
			rsFast.close
			set rsFast=nothing
			%></span>
			</td>
		</tr>
	<%end if%>
	</table>

<%		
		dim fp
			If sys_City="高雄市" Then
				fp="F:\\Image\\BillBaseDetail\\"&rs1("Sn") 
			elseIf sys_City="苗栗縣" Then
				fp="F:\\Image\\BillBaseDetail\\"&rs1("Sn") 
			else
                 fp="d:\\F\\Image\\BillBaseDetail\\"&rs1("Sn") 
			End if
                 set fso=Server.CreateObject("Scripting.FileSystemObject")
				 i=0
	    if (fso.FolderExists(fp))=true then 
			response.write "<br>"
            set fod=fso.GetFolder(fp)
            set fic=fod.Files
            For Each fil In fic
				If fil.Name<>"Thumbs.db" then
        		 i=i+1
				 if i<>1 then response.write ", "
				 %>
                    <a title="開啟影像資料.." onclick="OpenImageWinUserUpload('../../billimage/<%=rs1("Sn")&"/"&fil.Name%>')" <%lightbarstyle 1%>><u><font color="blue">影像 <%=i%></font></u></a>
			    <%
				End If 
            Next
       else
	       response.write "&nbsp;"
       end if
	   if i=0 then response.write "&nbsp;"

		 set fso=nothing
         set fod=nothing
         set fic=Nothing


	rs1.MoveNext
	Wend
	rs1.close
	set rs1=Nothing
	
	
%>
<Div id="Layer111" style="width:1041px; height:24px; ">
  <div align="center">
  <input type="hidden" value="" name="IsShow">
  <input type="button" value="列印" onclick="DP();">
  <br>
   <%if (sys_City<>"高雄市" and sys_City<>ApconfigureCityName) then%>
    (若無列印鈕，可按下滑鼠右鍵選擇列印功能，格式為A4橫印)
	<%end if%>
  </div>
</Div>
<%
conn.close
set conn=nothing
%>
</form>
</body>
<script language="JavaScript">
function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
}
function OpenImageWin(ImgSN,illdate){
	urlstr='../ProsecutionImage/ShowIllImage.asp?ImgSN='+ImgSN+'&illdate='+illdate;
	newWin(urlstr,'MyDetail',1000,600,0,0,"yes","no","yes","no");
}

function showMailHistory(){
	
	myForm.IsShow.value="1";
	myForm.submit();
}

function hiddenMailHistory(){
	
	myForm.IsShow.value="0";
	myForm.submit();
}

function DP(){
<%if (sys_City="高雄市" or sys_City=ApconfigureCityName) and NoCase=0 then%>
	urlstr='BillBaseData_Detail_Print_Set.asp?BillSnTmp=<%=BillSnTmp%>';
	newWin(urlstr,'Billprint',350,400,300,150,"no","no","yes","no");
<%else%>
	window.focus();
	<%if Cnt=1 then%>
	Layer112.style.visibility="hidden";
	<%end if%>
	Layer111.style.visibility="hidden";
	window.print();
	window.close();
<%end if%>
}

function OpenImageWinUserUpload(ImgFileName){
	urlstr=ImgFileName;
	newWin(urlstr,'MyDetail',1000,600,0,0,"yes","no","yes","no");
}
</script>
</html>
