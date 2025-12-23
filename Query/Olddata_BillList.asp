<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舊資料查詢</title>
<!--#include virtual="Traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/OldData.INI"-->
<!--#include virtual="Traffic/Common/AllFunction.inc"-->
<%
Server.ScriptTimeout = 7200
dim Conn2
Set Conn2 = Server.CreateObject("ADODB.Connection")
'  smith  這部分要確認台中縣 / 市的 connection 怎麼設定 
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

     
	strSQL=session("BillSQL")
	'response.write strSQL
	'response.End
    set rsfound=conn.execute(strSQL)



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
		<td bgcolor="#E0E0E0">
			<table width="100%" border="1" cellpadding="4" cellspacing="1">
				<tr bgcolor="#EBFBE3" align="center">
					<th nowrap>類別</th>
					<th nowrap>舉發單號</th>
					<th nowrap>車號</th>
					<th >違規日</th>
					<th >車種</th>
					<th >駕駛人</th>
					<th >違規地點</th>
					<th >裁決日</th>
					<th >移送日</th>
					<th >法條</th>
				</tr>
				<tr bgcolor="#FFFFFF" align="center">
				<%
					while not rsfound.eof 

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
						response.write "<td width='7%'>" & chkBillType(trim(rsfound("acc_tp"))) & "&nbsp;</td>"
						response.write "<td width='8%'>" & trim(rsfound("tkt_no")) & "&nbsp;</td>"
						response.write "<td width='8%'>" & trim(rsfound("plt_no")) & "&nbsp;</td>"
                        response.write "<td width='6%'>" & trim(rsfound("vil_dt")) & "</td>"						
						response.write "<td width='13%'>"& GetCarType(trim(rsfound("Car_tp"))) &"&nbsp;</td>"
						response.write "<td width='8%'>" & trim(rsfound("driver")) &  "&nbsp;</td>"					
						response.write "<td width='30%'>" & composeAddress(Address,trim(rsfound("vil_a1")),trim(rsfound("vil_b1")),trim(rsfound("vil_c1")),trim(rsfound("vil_d1")),trim(rsfound("vil_dr"))) &  "&nbsp;</td>"

						response.write "<td width='6%'>" & QueryCol("cin_dt","peo_rec","tkt_no=" & QuotedStr(trim(rsfound("tkt_no")))) & "</td>"
						response.write "<td width='6%'>" & QueryCol("out_dt","peo_rec","tkt_no=" & QuotedStr(trim(rsfound("tkt_no")))) & "</td>"

						response.write "<td width='8%'>" & trim(rsfound("Rule_1")) &  "&nbsp;</td>"					
						response.write "<td align='left' >"
						response.write "</td>"
						response.write "</tr>"
						rsfound.movenext
				    wend
				%>
			</table>
		</td>
	</tr>

<!--</table>-->

<input type="Hidden" name="DB_Selt" value="<%=request("DB_Selt")%>">
<input type="Hidden" name="kinds" value="">
<input type="Hidden" name="tmpSQL" value="<%=tempSQL%>">
<input type="Hidden" name="UnitList" value=<%=trim(request("UnitList"))  %>>

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

		if (myForm.IllegalDate.value=="" && myForm.IllegalDate1.value=="" && myForm.DriverID.value==""  && myForm.CarNo.value=="" && myForm.BillNo.value==""  ) {
				error=error+1;
				errorString=errorString+"\n"+error+"：至少要輸入一項!!";
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
	
</script>
<%
conn.close
set conn=nothing
response.contenttype="application/x-msexcel; charset=MS950" 
Response.AddHeader "Content-Disposition", "filename=舊資料查詢.xls"  
%>