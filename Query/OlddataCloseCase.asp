<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>舊資料結案註記</title>
<!--#include virtual="Traffic/Common/css.txt"-->
<!--#include virtual="traffic/Common/OldData.INI"-->
<!--#include virtual="Traffic/Common/AllFunction.inc"-->
<%
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

function chknullStr(str)
    if isnull(str) then
        chknullStr = ""
    else
        chknullStr = str
    end if
end function

function QuotedStr(Str)
    QuotedStr="'"+Str+"'"
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
        Sqltxt="Select " & trim(value1) & " from " & trim(value2) & " where " & trim(value3)
    else
        Sqltxt="Select " & trim(value1) & " from " & trim(value2)
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
        CloseNoSQL=""
        if trim(request("CloseNo")) <> "" then
            CloseNoSQL=",cls_no=" & QuotedStr(trim(request("CloseNo"))) 
        end If
        if sys_City="台中市" then
			CloseSql="update peo_ALL set cls_dt=" & QuotedStr(Year(now())-1911 &_
			 Right("00" & Month(Now())-1,2) & Day(Now())) &_
			CloseNoSQL &_
			",clspay=" & QuotedStr(request("ClosePay")) &_
			" where Tkt_no=" & QuotedStr(trim(request("CloseBillNo")))       
			Conn.Execute(CloseSql)
		End If
		
		If Trim(QueryCol("tkt_no","peo_New"," tkt_no='" & Trim(request("CloseBillNo")) & "'")) <> "" then
			CloseSql="update peo_New set cls_dt=" & QuotedStr(Year(now())-1911 &_
					  Right("00" & Month(Now())-1,2) & Day(Now())) &_
					  CloseNoSQL &_
					  ",clspay=" & QuotedStr(request("ClosePay")) &_
					  " where Tkt_no=" & QuotedStr(trim(request("CloseBillNo")))         
			Conn.Execute(CloseSql)
        Else
			'我要先撈出Vil_rec資料
			ACC_No=""
			ARV_DT=""
			ADVADD=""
			DRIVER=""
			ID_NUM=""
			MONEY1="0"
			MONEY2="0"
			POLICE=""
			R1_SB1=""
			R1_SB2=""
			R2_SB1=""
			R2_SB2=""
			Rule_1=""
			Rule_2=""
			TKT_no=""
			Vil_DT=""
			Vil_TM=""
			VILAD1=""

			BillBaseSQL="Select ACC_NO,ARV_DT,ARVADD,DRIVER,ID_NUM,Money1,Money2,Police,R1_SB1," &_
			"R1_SB2,R2_SB1,R2_SB2,Rule_1,Rule_2,Tkt_No,Vil_DT,Vil_TM,VILAD1 from vil_rec where tkt_No='" &_
			Trim(request("CloseBillNo")) & "'"
			set BillBaseRs=conn.execute(BillBaseSQL)
			While Not BillBaseRs.eof
				ACC_No = chknullStr(BillBaseRs("ACC_No"))
				ARV_DT = chknullStr(BillBaseRs("ARV_DT"))
				ADVADD = chknullStr(BillBaseRs("ARVADD"))
				DRIVER = chknullStr(BillBaseRs("DRIVER"))
				ID_NUM = chknullStr(BillBaseRs("ID_NUM"))
				MONEY1 = chknullStr(BillBaseRs("MONEY1"))
				MONEY2 = chknullStr(BillBaseRs("MONEY2"))
				POLICE = chknullStr(BillBaseRs("POLICE"))
				R1_SB1 = chknullStr(BillBaseRs("R1_SB1"))
				R1_SB2 = chknullStr(BillBaseRs("R1_SB2"))
				R2_SB1 = chknullStr(BillBaseRs("R2_SB1"))
				R2_SB2 = chknullStr(BillBaseRs("R2_SB2"))
				Rule_1 = chknullStr(BillBaseRs("Rule_1"))
				Rule_2 = chknullStr(BillBaseRs("Rule_2"))
				TKT_no = chknullStr(BillBaseRs("TKT_no"))
				Vil_DT = chknullStr(BillBaseRs("Vil_DT"))
				Vil_TM = chknullStr(BillBaseRs("Vil_TM"))
				VILAD1 = chknullStr(BillBaseRs("VILAD1"))
				BillBaseRs.movenext
			Wend
			BillBaseRs.close
			'新增資料
			AddPeoSQL = "Insert into Peo_new (ACC_NO,ARV_DT,ARVADD,DRIVER,ID_NUM,Money1,Money2,Police,R1_SB1," &_
			"R1_SB2,R2_SB1,R2_SB2,Rule_1,Rule_2,Tkt_No,Vil_DT,Vil_TM,VILAD1,cls_dt,cls_no,clspay) values (" &_
			QuotedStr(ACC_NO) & "," & QuotedStr(ARV_DT) & "," & QuotedStr(ADVADD) & "," & QuotedStr(DRIVER) & "," &_
			QuotedStr(ID_NUM) & "," & Money1 & "," & Money2 & "," & QuotedStr(Police) & "," &_
			R1_SB1 & "," & R1_SB2 & "," & R2_SB1 & "," & R2_SB2 & "," & QuotedStr(Rule_1) & "," &_
			QuotedStr(Rule_2) & "," & QuotedStr(TKT_no) & "," & QuotedStr(Vil_DT) & "," & QuotedStr(Vil_TM) & "," &_
			QuotedStr(VILAD1) & "," & QuotedStr(Year(now())-1911 & Right("00" & Month(Now())-1,2) & Day(Now())) & "," &_
			QuotedStr(trim(request("CloseNo"))) & "," & QuotedStr(request("ClosePay")) &_
			")"
			Conn.Execute(AddPeoSQL)
	    end if  
        
        CloseFlag = "0"
        response.Write "<script language='javascript'>"
        response.Write "alert('結案更新完成');"
        response.Write "window.opener.myForm.submit();"
        response.Write "window.close();"
        response.Write "</script>"
        
    end if

    
	BillNoSQL=""
	if request("BillNo")<>"" then
	    BillNoSQL=" and tkt_no = " & QuotedStr(trim(request("BillNo")))
	end If
	If Trim(QueryCol("tkt_no","peo_New"," tkt_no='" & Trim(request("CloseBillNo")) & "'")) <> "" then
		strSQL="Select * from PEO_NEW where 1=1 " & BillNoSQL
	Else
		strSQL="Select * from Vil_Rec where 1=1 " & BillNoSQL
	End if
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
<table width="100%" border="0" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF" style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid">
	<tr height="25" >
		<td bgcolor="#FFCC33" colspan="2" >
		    <b>舊資料結案註記</b>
		</td>
	</tr>			
	<tr>
	    <td  style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid; width:10%; background-color: #ffff99;">
	        舉發單號
	    </td>
	    <td  style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid">
	        <%=rsfound("tkt_no") %>
	    </td>
	</tr>
	<tr>
	    <td  style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid; width:10%; background-color: #ffff99;">
	        違規人
	    </td>
	    <td  style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid">
	        <%=rsfound("Driver") %>
	    </td>
	</tr>
	<tr>
	    <td  style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid; width:10%; background-color: #ffff99;">
	        違規法條
	    </td>
	    <td  style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid">
	        <% 
	            response.Write rsfound("Rule_1") & " , " &QueryCol("MEY_01","rule_n","Rule_C=" & QuotedStr(rsfound("Rule_1"))) & " , " &_
	            QueryCol("MEY_02","rule_n","Rule_C=" & QuotedStr(rsfound("Rule_1"))) & " , " &_
	            QueryCol("MEY_03","rule_n","Rule_C=" & QuotedStr(rsfound("Rule_1"))) & " , " &_
	            QueryCol("MEY_04","rule_n","Rule_C=" & QuotedStr(rsfound("Rule_1"))) & "<br>" & QueryLawContent(rsfound("Rule_1"),0,0)
	        %>
	    </td>
	</tr>
	<tr>
	    <td  style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid; width:10%; background-color: #ffff99;">
	        繳費金額
	    </td>
	    <td  style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid">
	        <input name="ClsPay" type="text" value="" size="20" maxlength="9" class="btn2">
	    </td>
	</tr>
	<tr>
	    <td  style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid; width:10%; background-color: #ffff99;">
	        收據字號
	    </td>
	    <td  style="border-right: thin solid; border-top: thin solid; border-left: thin solid; border-bottom: thin solid">
	        <input name="CLSNO" type="text" value="" size="20" maxlength="9" class="btn2">
	    </td>
	</tr>
	<tr>
		<td colspan="2">
		    <input type="button" name="btnSelt" value="結案" onclick="CloseSend('<%=trim(rsfound("tkt_no")) %>');">
		    <input type="button" name="btnSelt" value="取消/關閉" onclick="javascript:window.close();">
		</td>					
	</tr>
</table>
<input type="Hidden" name="CloseFlag" value="">
<input type="Hidden" name="CloseBillNo" value="">
<input type="Hidden" name="CloseNo" value="">
<input type="Hidden" name="ClosePay" value="">
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

function CloseSend(Billno)
{
    var aaa=confirm("你是否要結案?")
    if (aaa)
    {
        myForm.CloseFlag.value=1;
        myForm.CloseBillNo.value=Billno;
        myForm.CloseNo.value = myForm.CLSNO.value;
        myForm.ClosePay.value = myForm.ClsPay.value;
        myForm.submit();
    }
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
%>