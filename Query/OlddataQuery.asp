<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<%Server.ScriptTimeout=12000%>
<script language="JavaScript">
    var w;
    function openWindow(url, id, args) {
	    w = window.open(url,id,args);
    }
    function sendMsg() {
	    if (w && !w.closed) {
		    w.document.form2.msg.value = document.form1.msg.value;
		    w.focus();
	    } else {
		    alert("目標視窗未開啟！");
	    }
    }
	//window.focus();
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舊資料查詢</title>
<!--#include virtual="Traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/OldData.INI"-->
<!--#include virtual="Traffic/Common/AllFunction.inc"-->
<%
function GetCtoWDate(iDate)
Dim iTemp 
	if trim(iDate)<>"" then
		iTemp=cdbl(left("0"&iDate,3))+1911&"/"&_
		left(right("0"&month(iDate),4),2) &"/"&_
		right("0"&day(iDate),2)
	    GetCtoWDate = iTemp
	else
		GetCtoWDate = "- -"
	end if
End Function

dim Conn2
Set Conn2 = Server.CreateObject("ADODB.Connection")
Provider="Provider=MSDAORA;Data Source=orcl;User Id=traffic;Password=joly902f;"
Conn2.Open Provider
strCity="select value from Apconfigure where id=31"
set rsCity=conn2.execute(strCity)
	sys_City=trim(rsCity("value"))
rsCity.close
set rsCity=nothing
conn2.close
set conn2=nothing

function QuotedStr(Str)
    QuotedStr="'"+Str+"'"
end function

function chkBillType(BillTypeID)
    if trim(BillTypeID) <> "" then
        Select Case  trim(BillTypeID)
            Case "1" chkBillType="攔停"
            Case "2" chkBillType="逕舉"
            Case "3" chkBillType="逕舉手開單" 
            Case "4" chkBillType="拖吊" 
            Case "5" chkBillType="慢車行人"   
            Case "6" chkBillType="肇事"   
            Case "7" chkBillType="掌-攔停"   
            Case "8" chkBillType="掌-行人"   
            Case "9" chkBillType="掌電拖吊"   
            Case "H" chkBillType="人工移送"   
            Case "M" chkBillType="郵寄處理"   
            Case "N" chkBillType="攔停逕行(未開單)"   
            Case "D" chkBillType="註銷"   
            Case "R" chkBillType="單退"   
            Case "V" chkBillType="掌電拖吊(補開單)"   
        end select       
    end if 
end function

function GetCarType(CarTypeID)
    if trim(CarTypeID) <> "" then
        Select Case  trim(CarTypeID)
            Case "1" GetCarType="自大客車"
            Case "2" GetCarType="自大貨車"
            Case "3" GetCarType="自小客(貨)" 
            Case "4" GetCarType="營大客車" 
            Case "5" GetCarType="營大貨車"   
            Case "6" GetCarType="營小貨車"   
            Case "7" GetCarType="營小客車"   
            Case "8" GetCarType="租賃小客"   
            Case "9" GetCarType="遊覽客車"   
            Case "A" GetCarType="營交通車"   
            Case "B" GetCarType="貨櫃曳引"   
            Case "C" GetCarType="自用拖車"   
            Case "D" GetCarType="營業拖車"   
            Case "E" GetCarType="外賓小客"   
            Case "F" GetCarType="外賓大客"   
            Case "H" GetCarType="普通重機"    
            Case "L" GetCarType="輕機"   
            Case "p" GetCarType="併裝車"   
            Case "x" GetCarType="動力機械"      
            Case "Y" GetCarType="租賃小貨車"
            Case "W" GetCarType="自小客"
            Case "V" GetCarType="自小貨"
            Case "G" GetCarType="大型重機250CC"
            Case "Q" GetCarType="大型重機550CC"    
        end select       
    end if 
end function

'組小數點位數
function composeDot(value1,value2)  
    if (trim(value1) <> "") and (trim(value2) <> "") then
        composeDot = value1 & "." & value2
    end if 
    if (trim(value1) <> "") and (trim(value2) = "") then
        composeDot = value1
    end if
end  function

'判斷法條第八碼，將其組合起來
function composeLaw(value1,value2)
    if trim(value2) <> "" then
        composeLaw = value1 & value2
    else
        composeLaw = value1
    end if  
end function


'查詢單位名稱
function QueryUnitName(value)
    UnitSql="Select Acc_NM from accnew where ACC_No=" & QuotedStr(trim(value))
    set UnitRs=conn.execute(UnitSql)
    if  not UnitRs.Eof then
        QueryUnitName = UnitRs("Acc_NM")
    end if      
    UnitRs.close
end function  

'查詢其中一個欄位
function QueryCol(value1,value2,value3)
    if trim(value3) > "" then
        Sqltxt="Select " & value1 & " from " & value2 & " where " & value3
    else
        Sqltxt="Select " & value1 & " from " & value2 
    end if  
    set QueryRs=conn.execute(Sqltxt)
    if  not QueryRs.Eof then
        QueryCol = trim(QueryRs(value1))
    end if      
    QueryRs.close
end function   

function SetDate(tDate)
    if len(tDate)=6 then
        SetDate=mid(tDate,1,2)+1911 &"/"& mid(tDate,3,2)&"/"& mid(tDate,5,2)
    else
        SetDate=""
    end if
end function

 '查詢法條內容
function QueryLawContent(value,RealSpeed,LimitSpeed) 
    LawSql="Select * from rule_n where Rule_C=" & QuotedStr(trim(value))
    set LawRs=conn.execute(LawSql)
    if not LawRs.Eof then
        if (mid(LawRs("Rule_c"),1,3)="293") and (LawRs("A_DESC")="1") then
            QueryLawContent = replace(LawRs("Rule_D"),"重量 噸","重量" & LimitSpeed & "噸")
            QueryLawContent = replace(QueryLawContent,"過磅 噸","過磅" & RealSpeed & "噸")
            QueryLawContent = replace(QueryLawContent,"超載 噸","超載 " & RealSpeed-LimitSpeed & "噸")
        elseif (mid(LawRs("Rule_c"),1,4)="4010") and (LawRs("A_DESC")="1") then
            QueryLawContent = replace(LawRs("Rule_D"),"限速 公里","限速" & LimitSpeed & "公里")
            QueryLawContent = replace(QueryLawContent,"時速 公里","時速" & RealSpeed & "公里")
            QueryLawContent = replace(QueryLawContent,"超速 公里","超速" & RealSpeed-LimitSpeed & "公里")
        else
             QueryLawContent = trim(LawRs("Rule_D"))
        end if 
    end if 
end  function

'組地址字串
function composeAddress(Address,lane,alley,No,Dash,Direct)
    composeAddress=""
    composeAddress=Address
    if  trim(lane) <> "" then
        composeAddress=composeAddress & lane & "巷"
    end if
    if  trim(alley) <> "" then
        composeAddress=composeAddress & alley & "弄"
    end if
    if  trim(No) <> "" then
        composeAddress=composeAddress & No & "號"
    end if
    if  trim(Dash) <> "" then
        composeAddress=composeAddress & "之" & Dash
    end if
    if  trim(Direct) <> "" then
        composeAddress=composeAddress & "往" &Direct & "方向"
    end if 
end function

function GetTime(ttime)
    W=""
    H=""
    H=left(ttime,2)
    N=right(ttime,2)
    if len(ttime)>=4 then
        GetTime=H & ":" & N
    else
        GetTime=""
    end if
end function

'判斷如果是0的話回傳&nbsp;
function ReplaceSpace(value)
    if value="" then
        ReplaceSpace =  "&nbsp;"
    else
        ReplaceSpace = value 
    end if 
end function  

'查詢操作人員名稱
function QueryOperat(value)  
    OperatSql="Select OPName from operat where Operat=" & QuotedStr(trim(value))
    set OperatRs=conn.execute(OperatSql)
    if  not OperatRs.Eof then
        QueryOperat = OperatRs("OPName")
    end if 
    OperatRs.close
end  function 
		
		'法條加入Select
'		response.write trim(request("LawIDList"))
		LawListOption = ""
		LawIDList2 = ""
		If trim(request("LawNameList")) <> "" Then
			array_Law = Split(trim(request("LawNameList")),",")
			For i=0 To UBound(array_Law)
				If LawListOption = "" Then
					LawListOption = "<option value=""" & array_Law(i) & """>" & array_Law(i) & "</option>" & CHR(10)
				Else
					LawListOption = LawListOption & "," & "<option value=""" & array_Law(i) & """>" & array_Law(i) & "</option>" & CHR(10)
				End If 
			Next

			array_LawID = Split(trim(request("LawIDList")),",")
			For i=0 To UBound(array_LawID)
				If LawIDList2 = "" Then
					LawIDList2 = "'" & array_LawID(i) & "'"
				Else
					LawIDList2 = LawIDList2 & "," & "'" & array_LawID(i) & "'"
				End if
			next
		End if

    'Parse UnitList
    UnitIDList = ""
    UnitOption = ""
    if trim(request("UnitList")) <> "" then
        array_Unit=Split(trim(request("UnitList")),",")
        for i=0 to ubound(array_Unit)
            Unit=Split(array_Unit(i),"_")
            if UnitOption = "" then
                UnitOption = "<option value=""" & Unit(0) & """>" & Unit(1) & "</option>" & CHR(10)
            else
                UnitOption = UnitOption & "," & "<option value=""" & Unit(0) & """>" & Unit(1) & "</option>" & CHR(10)
            end if

            if UnitIDList = "" then
                UnitIDList = QuotedStr(Unit(0))
            else
                UnitIDList = UnitIDList & "," & QuotedStr(Unit(0))
            end if
        next
    end if
    
    if request("CloseFlag")="1" then
        if sys_City="台中市" then
            CloseSql="update peo_ALL set cls_dt=" & QuotedStr(Year(now())-1911 &_
                         Right("00" & Month(Now())-1,2) & Day(Now())) &_
                        " where Tkt_no=" & QuotedStr(trim(request("CloseBillNo")))       
            Conn.Execute(CloseSql)
        end if
        CloseSql="update peo_New set cls_dt=" & QuotedStr(Year(now())-1911 &_
                     Right("00" & Month(Now())-1,2) & Day(Now())) &_
                    " where Tkt_no=" & QuotedStr(trim(request("CloseBillNo")))       
        Conn.Execute(CloseSql)
        CloseFlag = "0"
    end if
    
    RecNoCloseSQL=""
    if request("RecNoCloseState")="1" then
        if sys_City="台中市" then
            RecNoCloseSQL="update peo_all set cls_dt=null,cls_no=null,clspay=null where Tkt_no=" & QuotedStr(trim(request("RecNoCloseBillNo")))         
            Conn.Execute(RecNoCloseSQL)
        end if
        RecNoCloseSQL="update peo_new set cls_dt=' ',cls_no=' ',clspay='0' where Tkt_no=" & QuotedStr(trim(request("RecNoCloseBillNo")))         
        Conn.Execute(RecNoCloseSQL)
    end if
    
if request("DB_Selt")="Selt" then   
  DateSQL=""
  if request("IllegalDate")<>"" and request("IllegalDate1")<>""then
		DateSQL=" and a.vil_dt >= " & QuotedStr(trim(request("IllegalDate"))) &" and a.vil_dt <= " & QuotedStr(trim(request("IllegalDate1")))
  end if		

	PayDaySQL = ""
	if request("PayDate")<>"" and request("PayDate1")<>""Then
		PayDaySQL = " and b.cls_dt >= " & QuotedStr(trim(request("PayDate"))) &" and b.cls_dt <= " & QuotedStr(trim(request("PayDate1")))
	End If
		
	BillNoSQL=""
	if request("BillNo")<>"" then
	    BillNoSQL=" and a.tkt_no = " & QuotedStr(trim(request("BillNo")))
	end if
    CarNoSQL=""
	if request("CarNo")<>"" then
	    CarNoSQL=" and a.plt_no = " & QuotedStr(trim(request("CarNo")))
	end if
    DriverIDSQL=""
	if request("DriverID")<>"" then 
	    DriverIDSQL=" and a.id_num = " & QuotedStr(trim(request("DriverID")))
    end if
    BillTypeSQL=""
	if trim(request("BillType"))<>"" and trim(request("BillType"))<>"0"  then 
	    BillTypeSQL=" and a.acc_tp = " & Quotedstr(trim(request("BillType")))
    end if
    BillUnitSQL = ""
    if UnitIDList <> "" then
        BillUnitSQL = " and a.acc_no in ( " & UnitIDList & ")"
    end if
    
  BillCloseSQL=""
	if trim(request("sBillClose")) <> "" and trim(request("sBillClose"))="1"  then 
			BillCloseSQL=" and (a.cls_dt<>'0' and a.cls_dt<>' ') "
	else
			BillCloseSQL=" and a.cls_dt=' '"
	end If
	
	LawSQL = ""

	If LawIDList2 <> "" Then
		LawSQL = " and a.Rule_1 in (" & LawIDList2 & ")"
	End If 
	
	ARVSQL = ""
	If Trim(request("ARVADD")) <>"" Then
		ARVSQL = " and ARVADD=" & QuotedStr(Trim(request("ARVADD")))
  End If
  
    if trim(request("billclosechk"))="on" then
        strSQL="Select * from Vil_rec a,(Select * from peo_New a where 1=1 " & DateSQL & BillCloseSQL & ") b "  &_
        "where a.Tkt_no = b.Tkt_no "  & BillNoSQL & CarNoSQL & DriverIDSQL & BillTypeSQL & BillUnitSQL & ARVSQL & LawSQL
        set rsfound=conn.execute(strSQL)        

        strCnt="Select Count(*) as FieldCount from Vil_rec a,(Select * from peo_New a where 1=1 " & DateSQL  & BillCloseSQL & ") b "  &_
        "where a.Tkt_no = b.Tkt_no "  & BillNoSQL & CarNoSQL & DriverIDSQL & BillTypeSQL & BillUnitSQL & ARVSQL & LawSQL
        set cnt=conn.execute(strCnt)
        DBsum=cnt("FieldCount")
	    set cnt=nothing
    else
    	if trim(request("Paydatechk"))="on" then
				strSQL="Select * from Vil_rec a,(Select * from peo_New a where 1=1 " & DateSQL & ") b "  &_
        "where a.Tkt_no = b.Tkt_no "  & BillNoSQL & CarNoSQL & DriverIDSQL & BillTypeSQL & BillUnitSQL & ARVSQL & LawSQL & PayDaySQL
        '筆數
        strCnt="Select Count(*) as FieldCount from Vil_rec a,(Select * from peo_New a where 1=1 " & DateSQL  & ") b "  &_
        "where a.Tkt_no = b.Tkt_no "  & BillNoSQL & CarNoSQL & DriverIDSQL & BillTypeSQL & BillUnitSQL & ARVSQL & LawSQL & PayDaySQL
    	else
	      'acc_tp,tkt_no,plt_no,vil_dt,driver,vilad1,vil_a1,vil_b1,vil_c1,vil_d1,vil_dr,Rule_1,Car_tp 
		    strSQL="Select a.* from Vil_rec a where 1=1 " & DateSQL & BillNoSQL & CarNoSQL & DriverIDSQL & BillTypeSQL & BillUnitSQL & ARVSQL  & LawSQL
		    '筆數
		    strCnt="Select Count(*) as FieldCount from Vil_rec a where 1=1 " & DateSQL & BillNoSQL & CarNoSQL & DriverIDSQL & BillTypeSQL & BillUnitSQL & ARVSQL & LawSQL
		  end if
		    'response.write strSQL
		    'response.End
	      set rsfound=conn.execute(strSQL)
		    '筆數
	      set cnt=conn.execute(strCnt)
		    DBsum=cnt("FieldCount")
		    set cnt=nothing
	end if

  Session.Contents.Remove("BillSQL")
  Session("BillSQL")=strSQL
end if


%>

<style type="text/css">
<!--
.style5 {
	font-size: 10pt;
}
.style7 {
	font-size: 10pt;
	font-family: "標楷體";}
.style8 {
	font-size: 14pt;
	}
.style6 {
	font-size: 16pt;
	font-weight: bold;
	line-height:22px;
	font-family: "標楷體";
}
.style11 {
	font-size: 10px;
	font-family: "標楷體";
}
.style12 {
	font-size: 20px;
	font-family: "標楷體";
}
.style22 {font-size: 9pt; font-family: "標楷體"; }
-->
</style>
</head>
<body>
<form name=myForm method="post">


<table width="100%" border="0">
	<tr>
		<td bgcolor="#CCCCCC">
			<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
				<tr height="25">
					<td bgcolor="#FFCC33">
					    <b>舊資料查詢</b>
					    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
					    <img src="space.gif" width="20" height="2"> <A HREF="..\舊資料查詢系統.doc"><FONT SIZE="3"><b>!!  第一次使用請看.DOC !! </b> </font></A>
					</td>
				</tr>			
				<tr>
					<td>
						違規日期
						<input name="IllegalDate" type="text" value="<%=request("IllegalDate")%>" size="6" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('IllegalDate');">
						~
						<input name="IllegalDate1" type="text" value="<%=request("IllegalDate1")%>" size="6" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="datestr" value="..." onclick="OpenWindow('IllegalDate1');">
						
						<img src="space.gif" width="20" height="2">						
						<input type="checkbox" name="Paydatechk" <%if trim(request("Paydatechk"))="on" then response.write "checked"%>/>
						繳款日期(<Font Color='Red'><B>打勾只會查慢車行人</B></Font>)
						<input name="PayDate" type="text" value="<%=request("PayDate")%>" size="6" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="Paydatestr" value="..." onclick="OpenWindow('PayDate');">
						~
						<input name="PayDate1" type="text" value="<%=request("PayDate1")%>" size="6" maxlength="7" class="btn1" onKeyup="value=value.replace(/[^\d]/g,'')">
						<input type="button" name="Paydatestr" value="..." onclick="OpenWindow('PayDate1');">
						<br /><br />
						
						違規人證號
						<input name="DriverID" type="text" value="<%=request("DriverID")%>" size="9" maxlength="10" class="btn1" onkeyup="value=value.toUpperCase()">
						<img src="space.gif" width="20" height="2">
						車<img src="space.gif" width="20" height="2">號
						<input name="CarNo" type="text" value="<%=request("CarNo")%>" size="9" maxlength="8" class="btn1" onkeyup="value=value.toUpperCase()">					
						<img src="space.gif" width="20" height="2">
						<b>單<img src="space.gif" width="20" height="2">號</b>
						<input name="BillNo" type="text" value="<%=request("BillNo")%>" size="9" maxlength="9" class="btn1" onkeyup="value=value.toUpperCase()">					
            <br /><br />
						舉發類型
						<select id="BillType" name="BillType" >
								<option value="0" /<%if trim(request("BillType"))="0" then response.write "selected"%>>全部適用</option>
								<option value="1" /<%if trim(request("BillType"))="1" then response.write "selected"%>>攔停</option>
								<option value="2" /<%if trim(request("BillType"))="2" then response.write "selected"%>>逕舉</option>
								<option value="3" /<%if trim(request("BillType"))="3" then response.write "selected"%>>逕舉手開單</option>
								<option value="4" /<%if trim(request("BillType"))="4" then response.write "selected"%>>拖吊</option>
								<option value="5" /<%if trim(request("BillType"))="5" then response.write "selected"%>>慢車行人</option>
								<option value="6" /<%if trim(request("BillType"))="6" then response.write "selected"%>>肇事</option>
						</select>

						<img src="space.gif" width="20" height="2">
						舉發單位
						<select id="BillUnit" name="BillUnit" >
							<%=UnitOption%>
						</select>

						<input type="button" name="btnUnit" value="選擇單位" onclick="opened=openWindow('Olddata_Unit.asp','myWin','scrollbars=yes');">
						<img src="space.gif" width="20" height="2">

						到案地點
						<select id="ARVADD" name="ARVADD" >
							<%
								response.write "<option value="
								if trim(request("ARVADD"))="" then response.write ""
								response.write " >全部適用</option>"
								AddARVSql="Select * from arvadd where item_N=5 and SPRVSN <>'0'"       
								set AddARVRs=Conn.Execute(AddARVSql)
								While Not AddARVRs.Eof
									response.write "<option value=" & AddARVRs("ARVADD") & " "
									If Trim(request("ARVADD"))=AddARVRs("ARVADD") Then response.write "Selected"
									response.write " >" & AddARVRs("ARV_NM") & "</option>"
									AddARVRs.moveNext
								Wend
							%>
						</select>

						&nbsp;
						<input type="checkbox" name="billclosechk" <%if trim(request("billclosechk"))="on" then response.write "checked"%>/>					
						是否結案(69條之後)
						<select id="Select2" name="sBillClose" >
                            <option value="0" /<%if trim(request("sBillClose"))="0" then response.write "selected"%>>否</option>
                            <option value="1" /<%if trim(request("sBillClose"))="1" then response.write "selected"%>>是</option>
                        </select>
					</td>		
				</tr>
				<tr>
					<td>
						法條選擇
						<select name="select3" id="select3">
						<%=LawListOption%>
          </select>
						<img src="space.gif" width="20" height="2">
						<input type="button" name="btnAdd" value="加入法條" onClick="openQryLaw(this.value);">
						<img src="space.gif" width="20" height="2">
						<input type="button" name="btnSelt" value="查詢" onclick="funSelt();">
						<img src="space.gif" width="20" height="2">
						<input type="button" name="cancel" value="清除" onClick="location='OlddataQuery.asp'"> 
					</td>
				</tr>
			</table>
	
	<tr height="30">
		<td bgcolor="#FFCC33" class="style3">
			資料紀錄列表
			<img src="space.gif" width="5" height="8">
			每頁 
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
			筆 <font color="#F90000"><strong>(共 <%=DBsum%> 筆 )</strong></font>
			&nbsp; &nbsp; 
			&nbsp;
			
<!--			<select name="sys_OrderType" onchange="repage();">
'				<option value="2" <%if trim(request("sys_OrderType"))="1" then response.write " Selected"%>>違規日期</option>
'				<option value="3" <%if trim(request("sys_OrderType"))="3" then response.write " Selected"%>>綜合資料號</option>
			</select>
			<select name="sys_OrderType2" onchange="repage();">
				<option value="1" <%if trim(request("sys_OrderType2"))="1" then response.write " Selected"%>>由小至大</option>
				<option value="2" <%if trim(request("sys_OrderType2"))="2" then response.write " Selected"%>>由大至小</option>
			</select>
			排列&nbsp; &nbsp;
-->			
		</td>
	</tr>
	
	<tr>
		<td bgcolor="#E0E0E0">
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th nowrap width='50'>類別</th>
					<th nowrap>舉發單號</th>
					<th nowrap>車號</th>
					<th >違規日</th>
					<th >車種</th>
					<th >駕駛人</th>
					<th >違規地點</th>
					<th >催告日</th>
					<th >裁決日</th>
					<th >移送日</th>
					<th >法條</th>
					<th >操作</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
				<%
                
				if request("DB_Selt")="Selt"  then
					if Trim(request("DB_Move"))="" then
						DBcnt=0
					else
						DBcnt=request("DB_Move")
					end if
					if Not rsfound.eof then rsfound.move DBcnt
					for i=DBcnt+1 to DBcnt+10+request("sys_MoveCnt")
						if rsfound.eof then exit for
						Address=""
						AddressSQL=""
						if trim(rsfound("vilad1")) <> "" then
                            AddressSQL="Select * from addr_c where addr_c= " & QuotedStr(trim(rsfound("vilad1")))
                            set Rs2=conn.execute(AddressSQL)
							if not RS2.eof then
								Address= Rs2("Addr_D")
							end if
                        end if  
                        Rs2.close
                        
						response.write "<tr bgcolor='#FFFFFF' align='center'  height='23'"
						lightbarstyle 0 
						response.write ">"
						response.write "<td width='6%'><font size='2'>" & chkBillType(trim(rsfound("acc_tp"))) & "</font>&nbsp;</td>"
						response.write "<td width='6%'><font size='2'>" & trim(rsfound("tkt_no")) & "</font>&nbsp;</td>"
						response.write "<td width='6%'><font size='2'>" & trim(rsfound("plt_no")) & "</font>&nbsp;</td>"
            response.write "<td width='6%'><font size='2'>" & trim(rsfound("vil_dt")) & "</font></td>"						
						response.write "<td width='6%'><font size='2'>"& GetCarType(trim(rsfound("Car_tp"))) &"</font>&nbsp;</td>"
						response.write "<td width='6%'><font size='2'>" & trim(rsfound("driver")) &  "</font>&nbsp;</td>"					
						response.write "<td width='10%'><font size='2'>" & composeAddress(Address,trim(rsfound("vil_a1")),trim(rsfound("vil_b1")),trim(rsfound("vil_c1")),trim(rsfound("vil_d1")),trim(rsfound("vil_dr"))) &  "</font></td>"					
						if trim(request("billclosechk"))="on" then
						    response.write "<td width='6%'><font size='2'>" & trim(rsfound("HUR_DT")) &  "</font>&nbsp;</td>"
						    response.write "<td width='6%'><font size='2'>" & trim(rsfound("DES_DT")) &  "</font>&nbsp;</td>"
						    response.write "<td width='6%'><font size='2'>" & trim(rsfound("REM_DT")) &  "</font>&nbsp;</td>"					
						    response.write "<td width='6%'><font size='2'>" & trim(rsfound("Rule_1")) &  "</font>&nbsp;</td>"					
						else
						PasserDate = SetDate(QueryCol("cin_dt","peo_rec","tkt_no=" & QuotedStr(trim(rsfound("tkt_no"))))) 
						    response.write "<td width='6%'><font size='2'>"& QueryCol("HUR_DT","peo_New"," tkt_no='"&trim(rsfound("tkt_no")) & "'") &"</font>&nbsp;</td>"

						    'response.write "<td width='6%'><font size='2'>"& QueryCol("DES_DT","peo_New"," tkt_no='"&trim(rsfound("tkt_no")) & "'") &"</font>&nbsp;</td>"

							response.write "<td width='6%'><font size='2'>"&SetDate(QueryCol("cin_dt","peo_rec","tkt_no=" & QuotedStr(trim(rsfound("tkt_no"))))) &"</font>&nbsp;</td>"

						    'response.write "<td width='6%'><font size='2'>"& QueryCol("REM_DT","peo_New"," tkt_no='"&trim(rsfound("tkt_no")) & "'") &"</font>&nbsp;</td>"

							response.write "<td width='6%'><font size='2'>"&SetDate(QueryCol("out_dt","peo_rec","tkt_no=" & QuotedStr(trim(rsfound("tkt_no")))))&"</font>&nbsp;</td>"

						    response.write "<td width='6%'><font size='2'>" & trim(rsfound("Rule_1")) &  "</font>&nbsp;</td>"					
						end if
						
						response.write "<td align='left' >"
            %>	
                
			    <input type="button" name="b1" value="詳細" onclick='window.open("OlddataDetail.asp?BillNo=<%=trim(rsfound("tkt_no"))%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=620,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">
			   <% UserID=trim(Session("User_ID")) %>
			    <input type="button" name="b2" value="備註"  onclick='window.open("olddataNote.asp?<%="UserID="&UserID & "&BillNO=" & trim(rsfound("tkt_no"))%>","OldBaseDetail","left=0,top=0,location=0,width=500,height=275,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;"> 
			    
			    <%
							
			        '違規地點 
			        ILLEGALADDRESS = composeAddress(Address,trim(rsfound("vil_a1")),trim(rsfound("vil_b1")),trim(rsfound("vil_c1")),trim(rsfound("vil_d1")),trim(rsfound("vil_dr")))  
			        '標準及違規車速一
		            RuleSpeed1 = composeDot(trim(rsfound("R1_SB1")),Mid(trim(rsfound("rece01")),1,2)) 
		            IllegalSpeed1 = composeDot(trim(rsfound("R1_SB2")),Mid(trim(rsfound("rece01")),3,2))
		            '標準及違規車速二
		            RuleSpeed2 = composeDot(trim(rsfound("R2_SB1")),Mid(trim(rsfound("rece02")),1,2)) 
		            IllegalSpeed2 = composeDot(trim(rsfound("R2_SB2")),Mid(trim(rsfound("rece02")),3,2))
		            '違規法條一
		            Rule1 = composeLaw(trim(rsfound("Rule_1")),Mid(rsfound("rece03"),3,1))
		            '違規法條二
		            Rule2 = composeLaw(trim(rsfound("Rule_2")),Mid(rsfound("rece03"),4,1))
		            '違規事實一
		            Rule1txt = trim(QueryLawContent(composeLaw(trim(rsfound("Rule_1")),Mid(rsfound("rece03"),3,1)),IllegalSpeed1,RuleSpeed1))
		            '違規事實二
		            Rule2txt = trim(QueryLawContent(composeLaw(trim(rsfound("Rule_2")),Mid(rsfound("rece03"),4,1)),IllegalSpeed2,RuleSpeed2))
		            '監理站名稱
		            StationName = QueryCol("arv_nm","Arvadd"," arvadd=" & QuotedStr(trim(rsfound("arvadd"))))

			        if (trim(rsfound("acc_tp")) <> "5")  then		            
			            'DriverZipName,OwnerZipName由titan抓取  
			            '扣件

			            hold = QueryCol("Hold_D","hold_c","Hold_C=" & QuotedStr(mid(rsfound("hold_c"),1,1)) & " and Hold_C !=" & QuotedStr("0")) & "," & QueryCol("Hold_D","hold_c","Hold_C=" & QuotedStr(mid(rsfound("hold_c"),2,1)) & " and Hold_C !=" & QuotedStr("0"))  & "," & QueryCol("Hold_D","hold_c","Hold_C=" & QuotedStr(mid(rsfound("hold_c"),3,1)) & " and Hold_C !=" & QuotedStr("0"))

			            ManagerLevel = QueryCol("job_nm","accnew","acc_no=" & QuotedStr(trim(rsfound("acc_no"))))               '舉發單位主管職稱
			            Boss = QueryCol("led_nm","accnew","acc_no=" & QuotedStr(trim(rsfound("acc_no"))))               '應到案單位局長		            
			            CarColor = QueryCol("colord","colorc","colorc  =" & QuotedStr(trim(rsfound("colorc"))))
									
			            PrtBillBase = "BillTypeID=" & trim(rsfound("acc_tp")) & "&" & "Driver=" & trim(rsfound("driver")) & "&" & "DriverID=" & trim(rsfound("id_num")) & "&" & "DriverAddress=" & trim(rsfound("d_addr")) & "&" & "DriverZip=" & trim(rsfound("drvzip")) & "&" & "ILLEGALADDRESS=" & ILLEGALADDRESS & "&" & "IllegalSpeed=" & IllegalSpeed & "&" & "RuleSpeed=" & RuleSpeed &_
			                             "&" & "Note=" & "" & "&" & "BillFillDate=" & GetCtoWDate(trim(rsfound("kin_dt"))) & "&" & "DriverHomeZip=" & trim(rsfound("drvzip")) & "&" & "Owner=" & trim(rsfound("ownerx")) & "&" & "OwnerAddress=" & trim(rsfound("addres")) & "&" & "OwnerZip=" & trim(rsfound("vehzip"))  & "&" & "FORFEIT1=" & trim(rsfound("money1")) & "&" & "FORFEIT2=" & trim(rsfound("money2")) & "&" & "DCIReturnStation=" & trim(rsfound("arvadd")) & "&" & "BillNo=" & trim(rsfound("tkt_no")) &_
			                             "&" & "CarNo=" & trim(rsfound("plt_no")) & "&" & "Rule1=" & Rule1 & "&" & "Rule2=" & Rule2 & "&" & "DetailCarType =" & trim(GetCarType(trim(rsfound("car_tp"))))  & "&" & "FillUnitName=" & trim(QueryUnitName(trim(rsfound("acc_no"))))  & "&" & "FillUnitTEL=" & ""  & "&" & "IllegalDate=" & ReplaceSpace(SetDate(rsfound("vil_dt")) & " " & GetTime(rsfound("vil_tm"))) &_
			                             "&" & "DealLineDate=" & trim(rsfound("ARV_DT")) & "&" & "operat=" & trim(rsfound("operat"))  & "&" & "MailDate=" & Now() & "&" & "Hold=" & trim(hold)  & "&" & "CarMark=" & "" & "&" & "CarColor=" & colord & "&" & "ManagerLevel=" & ManagerLevel & "&" & "Boss=" & Boss & "&" & "Rule1txt=" & Rule1txt & "&" & "Rule2txt=" & Rule2txt & "&" & "StationName=" & StationName
			         %> 

			         <%if sys_City="台中縣" then  %>
			            <input type="button" name=<%=PrtBillBase%>
			            value="補印" onclick='window.open("BillPrints_lattice_TaiChung_Mend.asp?<%=PrtBillBase%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=620,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">
			         <%else %>
                        <input type="button" name=<%=PrtBillBase%>
			            value="補印" onclick='window.open("BillPrintsTaiChungCity_a4_Mend.asp?<%=PrtBillBase%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=620,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 40px; height:26px;">			         
			         <%end if %>
			    <%
                    else
			            '裁決書
			            
			            ArrUnitName = QueryCol("arv_nm","arvadd","arvadd=" & QuotedStr(trim(rsfound("arvadd"))))   '應到案單位
			            Boss = QueryCol("led_nm","accnew","acc_no=" & QuotedStr(trim(rsfound("acc_no"))))               '應到案單位局長
			            UnitTEL = QueryCol("acctel","accnew","acc_no=" & QuotedStr(trim(rsfound("acc_no"))))           '應到案單位電話
			            UnitAddress = QueryCol("addres","accnew","acc_no=" & QuotedStr(trim(rsfound("acc_no"))))     '應到案單位地址
			            PasserPay = QueryCol("moneyn","peo_rec","tkt_no=" & QuotedStr(trim(rsfound("tkt_no"))))      '裁罰金額
			            PasserDate = SetDate(QueryCol("cin_dt","peo_rec","tkt_no=" & QuotedStr(trim(rsfound("tkt_no")))))      '裁罰日期
			            CloseDate = ""
			            clsno = ""
									
			            if sys_City="台中市" then
			                CloseDate = QueryCol("cls_dt","(Select cls_dt from peo_ALL where tkt_no=" & QuotedStr(trim(rsfound("tkt_no"))) & "union Select cls_dt from peo_New where tkt_no=" & QuotedStr(trim(rsfound("tkt_no"))) & ")","")      '結案日期
			                clsno = QueryCol("cls_no","(Select cls_no from peo_ALL where tkt_no=" & QuotedStr(trim(rsfound("tkt_no"))) & "union Select cls_no from peo_New where tkt_no=" & QuotedStr(trim(rsfound("tkt_no"))) & ")","")      '結案日期
                  else
                      CloseDate = QueryCol("cls_dt"," peo_New "," tkt_no=" & QuotedStr(trim(rsfound("tkt_no"))))      '結案日期
			                clsno = QueryCol("cls_no"," peo_New "," tkt_no=" & QuotedStr(trim(rsfound("tkt_no"))))      '結案日期
                  end if
			            if trim(PasserDate)="" then
			                PasserDate = DateAdd("d",15,Now())    '裁罰日期    
			            end if
			            operatName=QueryCol("OPName","Operat","OPerat=" & QuotedStr(trim(rsfound("operat"))))   '建檔人姓名 

			            '裁決字號
			            if sys_City="台中縣" then
			                PasserUnitName = "中警交字"
			            else
			                PasserUnitName = "中市警交字"
			            end if 
			            '裁決文號
			            PasserNo = QueryCol("vi_ser","peo_rec","tkt_no=" & QuotedStr(trim(rsfound("tkt_no")))) 
			            '移送日期
			            SendDate = trim(SetDate(QueryCol("out_dt","peo_rec","tkt_no=" & QuotedStr(trim(rsfound("tkt_no"))))))
			            if SendDate ="" then SendDate = now()
			            
			            '裁決
			            PrtPasser = "ArrUnitName=" & ArrUnitName & "&" & "SubBoss=" & "" & "&" & "Boss=" & Boss & "&" & "UnitTEL=" & UnitTEL & "&" & "UnitAccount=" & "" & "&" & "UnitAccountName=" & "" & "&" & "UnitAddress=" & UnitAddress  & "&" & "PasserUnitName=" & PasserUnitName &_
			                            "&" & "PasserNo=" & PasserNo & "&" & "operat=" & trim(rsfound("operat")) & "&" & "PasserDate=" & PasserDate & "&" & "BillNo=" & trim(rsfound("tkt_no")) & "&" & "DriverAddress=" & trim(rsfound("d_addr"))  & "&" & "DriverID=" & trim(rsfound("id_num"))  & "&" & "DriverBirth=" & SetDate(trim(rsfound("birthd"))) & "&" & "IllegalDate=" & ReplaceSpace(SetDate(rsfound("vil_dt")) & " " & GetTime(rsfound("vil_tm"))) &_
			                            "&" & "Boss=" & Boss & "&" & "IllegalAddress=" & ILLEGALADDRESS & "&" & "DealLineDate=" & SetDate(rsfound("ARV_DT")) & "&" & "Rule1=" & Rule1 & "&" & "Rule2=" & Rule2 & "&" & "Rule3=" & "" & "&" & "Rule4=" & "" &_
			                            "&" & "DriverZip=" & trim(rsfound("drvzip")) & "&" & "Driver=" & trim(rsfound("driver")) & "&" & "Rule1txt=" & Rule1txt & "&" & "Rule2txt=" & Rule2txt & "&" & "PasserPay=" & PasserPay
                        '移送
                        PrtSend = "BillNo=" & trim(rsfound("tkt_no")) & "&" & "ArrUnitName=" & ArrUnitName  & "&" & "operatLevel=" & "" & "&" & "operat=" & operatName & "&" & "PostDate=" & Now() & "&" & "UnitTEL=" & UnitTEL &_
                                      "&" & "PasserUnitName=" & PasserUnitName & "&" &  "SendWordNum=" & trim(rsfound("tkt_no")) & "&" & "PasserNo=" & PasserNo & "&" & "Driver=" & trim(rsfound("driver")) & "&" & "DriverBirth=" & SetDate(trim(rsfound("birthd"))) &_
                                      "&" & "DriverID=" & trim(rsfound("id_num")) & "&" & "DriverAddress=" & trim(rsfound("d_addr")) & "&" & "Rule1=" & Rule1 & "&" & "Rule2=" & Rule2 & "&" & "Rule1txt=" & Rule1txt & "&" & "Rule2txt=" & Rule2txt & "&" & "IllegalDate=" & ReplaceSpace(SetDate(rsfound("vil_dt")) & " " & GetTime(rsfound("vil_tm")))  &_
                                      "&" & "PasserPay=" & PasserPay & "&" & "SendDate=" & SendDate  & "&" & "UnitAccount=" & "" & "&" & "UnitAccountName=" & "&" & "DriverZip=" & trim(rsfound("drvzip")) 
                        '催告
                        PrtUrge = "BillNo=" & trim(rsfound("tkt_no")) & "&" & "IllegalDate=" & ReplaceSpace(SetDate(rsfound("vil_dt")) & " " & GetTime(rsfound("vil_tm"))) & "&" & "Driver=" & trim(rsfound("driver")) & "&" & "DriverID=" & trim(rsfound("id_num")) & "&" & "DriverAddress=" & trim(rsfound("d_addr")) &_
                                      "&" & "PasserUnitName=" & PasserUnitName & "&" & "UnitTEL=" & UnitTEL & "&" & "ArrUnitName=" & ArrUnitName & "&" & "ArrUnitID=" & trim(rsfound("arvadd")) & "&" & "Boss=" & Boss & "&" & "BillFillDate=" & trim(rsfound("kin_dt")) & "&" & "Rule1=" & Rule1
			    %>                     
	            <input type="button" name="b6" value="催告書"  onclick='window.open("../PasserBase/PasserUrge_Word_old.asp?<%=PrtUrge%>","OldBaseDetail","left=0,top=0,location=0,width=500,height=275,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 42px; height:26px;" >  
			        <input type="button" name="b4" value="裁決書" onclick='window.open("../Passerbase/PasserJudePrint_label_Mend.asp?<%=PrtPasser%>","OldBaseDetail","left=0,top=0,location=0,width=980,height=620,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 42px; height:26px;">   
			        <input type="button" name="b5" value="移送書"  onclick='window.open("../PasserBase/PaseBillPrit96_not_Mend.asp?<%=PrtSend%>","OldBaseDetail","left=0,top=0,location=0,width=500,height=275,resizable=yes,scrollbars=yes,menubar=yes")' style="font-size: 10pt; width: 42px; height:26px;"> 
			    <%
			        
			        if (isNull(CloseDate)) or (trim(CloseDate)="")  and (trim(CloseDate) <> "0") then		            
			    %>
			        <input type="button" name="btnClose" value="結案註記" onclick="CloseSend('<%=trim(rsfound("tkt_no")) %>');" style="font-size: 10pt; width: 57px; height:26px;">
			        <%
			        else
					'20091027 smith add 收據字號
			            response.Write "<font color=#F90000><b>"+"已結案<font size='2'>(收據號"+clsno +")</font></b></font>"
			        %>
                        <input type="button" name="btnRecNoClose" value="回復成未結案" onclick="RecNoClose('<%=trim(rsfound("tkt_no")) %>');" style="font-size: 10pt; width: 85px; height:26px;">			                                    
                    <%
			        end if
			    %>				
	<%
				    end if
				    
						response.write "</td>"
						response.write "</tr>"
						rsfound.movenext
					next
				end if
				%>
			</table>
		</td>
	</tr>
	<tr>
		<td height="35" bgcolor="#FFFFFF" align="center">
			<input type="button" name="MoveUp" value="上一頁" onclick="funDbMove(-10);">
			<span class="style2"> <%=fix(CDbl(DBcnt)/(10+request("sys_MoveCnt"))+1)&"/"&fix(CDbl(DBsum)/(10+request("sys_MoveCnt"))+0.9)%></span>
			<input type="button" name="MoveDown" value="下一頁" onclick="funDbMove(10);">
			<input type="button" name="btnExecel" value="轉換成Excel" onclick="funchgExecel();">
			<input type="button" name="btnExecel" value="慢車行人繳款清冊" onclick="funchgExecel1();">
		</td>
	</tr>
<!--</table>-->
<input type="Hidden" name="RecNoCloseState" value="">
<input type="Hidden" name="RecNoCloseBillNo" value="">
<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="CloseFlag" value="">
<input type="Hidden" name="CloseBillNo" value="">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="DB_Move" value="<%=DBcnt%>">
<input type="Hidden" name="DB_Cnt" value="<%=DBsum%>">
<input type="Hidden" name="tmpSQL" value="<%=tempSQL%>">
<input type="Hidden" name="UnitList" value=<%=trim(request("UnitList"))  %>>
<input type="Hidden" name="LawIDList" value=<%=trim(request("LawIDList"))  %>>
<input type="Hidden" name="LawNameList" value=<%=trim(request("LawNameList"))  %>>

</form>
</body>
</html>
<script type="text/javascript" src="../js/date.js"></script>
<script language="javascript">

function funDbMove(MoveCnt){
	if (eval(MoveCnt)>0){
		if (eval(myForm.DB_Move.value) < eval(myForm.DB_Cnt.value)-10-eval(myForm.sys_MoveCnt.value)){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt+eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}else{
		if (eval(myForm.DB_Move.value)>0){
			myForm.DB_Move.value=eval(myForm.DB_Move.value)+MoveCnt-eval(myForm.sys_MoveCnt.value);
			myForm.submit();
		}
	}
}
function repage(){
	myForm.DB_Move.value=0;
	myForm.submit();
}

function RecNoClose(BillNo){
    var aaa=confirm("你是否要回復成未結案狀態?")
    if (aaa)
    {
        myForm.RecNoCloseState.value=1;
				myForm.RecNoCloseBillNo.value=BillNo;
	    myForm.submit();
	}
}


function CloseSend(Billno)
{
    myForm.CloseFlag.value=1;
    window.open("OlddataCloseCase.asp?BillNo=" +Billno);
    myForm.submit();   
}

	function funSelt(){
		var error=0;
		var errorString="";

		if(myForm.IllegalDate.value!=""){
			if(!dateCheck(myForm.IllegalDate.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：違規日期輸入不正確!!";
			}
		}

		if(myForm.IllegalDate1.value!=""){
			if(!dateCheck(myForm.IllegalDate1.value)){
				error=error+1;
				errorString=errorString+"\n"+error+"：違規日期輸入不正確!!";
			}
		}

		if (!myForm.Paydatechk.checked)
		{
			if (myForm.IllegalDate.value=="" && myForm.IllegalDate1.value=="" && myForm.DriverID.value==""  && myForm.CarNo.value=="" && myForm.BillNo.value==""  ) {
					error=error+1;
					errorString=errorString+"\n"+error+"：請至少填入一樣資料!!";
			}
		}
		
		
		if (myForm.Paydatechk.checked)
		{
			if (myForm.PayDate.value=="" && myForm.PayDate1.value=="") {
					error=error+1;
					errorString=errorString+"\n"+error+"：請輸入繳款日期!!";
			}
		}
		
		
		if (error>0){
			alert(errorString);
		}else{
			myForm.DB_Move.value=0;
			myForm.DB_Selt.value="Selt";
			myForm.submit();
		}
	}

	function newWin(url,win,w,h,l,t,sBar,mBar,res,full){
		var win=window.open(url,win,"width="+w+",height="+h+",left="+l+",top="+t+",scrollbars="+sBar+",menubar="+mBar+",resizable="+res+",fullscreen="+full+",status=yes,toolbar=no");
		win.focus();
		return win;
	}

  function funchgExecel()
    {
		UrlStr="Olddata_BillList.asp?WorkType=1";
		newWin(UrlStr,"inputWin",790,550,50,10,"yes","yes","yes","no");
	}
	
	function funchgExecel1()
    {
		UrlStr="Olddata_detailList.asp?WorkType=1";
		newWin(UrlStr,"inputWin",790,550,50,10,"yes","yes","yes","no");
	}
	
	function openQryLaw(){
	 window.open("OldQueryLaw.asp?qryType=1&reportId=myForm","tmpWindow","width=600,height=355,left=0,top=0,resizable=yes,scrollbars=yes");
}	
</script>
<%
conn.close
set conn=nothing
%>