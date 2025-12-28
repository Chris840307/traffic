<%@ Page Language="VB" AutoEventWireup="false" CodeFile="getRewardList_Should_Unit_Set.aspx.vb" Inherits="ScoreList_getRewardList_Unit_Set" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script language="JavaScript">
	window.focus();
<%
    if ErrorCode<>"" then
        Response.Write("alert(""" & ErrorCode & """);")
    end if
%>
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>每月應領實領報表</title>
</head>
<body style="text-align: center; font-size: 12pt;">
    <form id="form1" runat="server">
    <div>
        <div style="text-align: center">
            <table style="width: 300px" border="1" >
                <tr style="background-color:#FFCC33">
                    <td style="height: 25px" align="center">
                        每月應領實領報表</td>
                </tr>
                <tr>
                    <td align="center" style="height: 35px">
                        <input name="tbDate1" type="text" value="<%=trim(request("tbDate1"))%>" MaxLength="6" onkeyup="value=value.replace(/[^\d]/g,'')" style="width: 70px" />
                        <input id="Button1" type="button" value=".." onclick="OpenSelectDate1('tbDate1')" />&nbsp;
                        至 &nbsp;<input name="tbDate2" type="text" value="<%=trim(request("tbDate2"))%>" MaxLength="6" onkeyup="value=value.replace(/[^\d]/g,'')" style="width: 70px" />
                        <input id="Button2" type="button" value=".." onclick="OpenSelectDate1('tbDate2')" />
                        <br>
                        <input type="radio" name="DateType" value="BillFillDate" checked />填單日期&nbsp;
                        <input type="radio" name="DateType" value="RecordDate" />建檔日期
                    </td>
                </tr>
                <tr style="background-color:#FFCC33">
                    <td style="height: 25px" align="center">
                        統計單位</td>
                </tr>
                <tr>
                    <td style="height: 99px" align="center">
                        <asp:Panel ID="Panel1" runat="server" BorderStyle="Inset" Height="95px" HorizontalAlign="Left"
                            ScrollBars="Vertical" Width="245px" Font-Size="12pt">
<%
    '取得 Web.config 檔的資料連接設定
    Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
    '建立 Connection 物件
    Dim conn As New Data.OracleClient.OracleConnection()
    conn.ConnectionString = setting.ConnectionString
    '開啟資料連接
    conn.Open()
    

    
    Dim IsChecked As String = ""
    Response.Write(IsChecked)
    Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
    Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
    If rdUnit.HasRows Then
        While rdUnit.Read()
           
            If InStr(Trim(Request("sUnitID")), Trim(rdUnit("UnitID"))) <> 0 Then
                IsChecked = " checked"
            Else
                IsChecked = ""
            End If
                
            Response.Write("<input type=""checkbox"" name=""sUnitID"" value=""" & "'" & Trim(rdUnit("UnitID")) & "'" & """ " & IsChecked & " />" & Trim(rdUnit("UnitName")) & "<br />")
        End While
    End If
    rdUnit.Close()
    conn.Close()
%>
                        </asp:Panel>
                        <input id="Button6" style="font-size: 8pt; width: 55px; height: 20px" type="button"
                            value="全部選取" onclick="AllUnit()" />&nbsp;
                        <input id="Button5" style="font-size: 8pt; width: 55px; height: 20px" type="button"
                            value="全部取消" onclick="NoUnit()" />
                        
                        </td>
                </tr>
                <tr style="background-color:#FFCC33; font-size: 12pt;">
                    <td style="height: 25px" align="center">
                        報表種類</td>
                </tr>
                <tr style="font-size: 12pt">
                    <td style="height: 35px" align="center">
                        <input id="btnDirector" style="width: 136px" type="button" value="直接人員結餘款" onclick="return btnDirector_onclick()" /><br />
                        <input id="btnCommon" style="width: 136px" type="button" value="共同人員結餘款" onclick="return btnCommon_onclick()" /><br />
                        <input id="btnItem" style="width: 136px" type="button" value="作業項目結餘" onclick="return btnItem_onclick()" /></td>
                </tr>
                <tr style="background-color:#F0FFFF; font-size: 12pt;">
                    <td align="center" style="height: 35px">
                        &nbsp; &nbsp;
                        <asp:Button ID="Button4" runat="server" Text="離開" OnClientClick="window.close();" Font-Size="12pt" Height="28px" Width="50px" />
                        <input type="hidden" name="AnalyzeType" value="<%=trim(request("AnalyzeType")) %>" />
                        <input type="hidden" name="AnalyzeMoney" value="<%=trim(request("AnalyzeMoney")) %>" />
                        </td>
                        
                </tr>
            </table>
        </div>
    
    </div>
        <span style="font-size: 10pt; color: #0066ff">
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
    //開啟萬年曆視窗
	function OpenSelectDate1(tag){
	    InitDate=eval("form1."+tag).value;
	    window.open("SelectDate.aspx?tag="+tag+"&InitDate="+InitDate,"OpenSelectDate1","width=240,height=240,left=350,top=250,scrollbars=no,menubar=no,resizable=no,fullscreen=no,status=no,toolbar=no");
	}
	function OpenRewardList(){
	    var error=0;
	    var errorString="";
	    var UnitName="";
	    //alert(form1.sCountyOrNpa2.value);
	    CheckFlag1=dateCheck(form1.tbDate1.value);
	    CheckFlag2=dateCheck(form1.tbDate2.value);
	    if (CheckFlag1==false){
	        error=error+1;
		    errorString=error+"：起始日期輸入錯誤!!";
	    }
	    if (CheckFlag2==false){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：結束日期輸入錯誤!!";
	    }
	    if (form1.sUnitID.length > 0){
	        for (i=0; i< form1.sUnitID.length; i++){
	            if(form1.sUnitID[i].checked==true){
	                if(UnitName==""){
	                    UnitName=form1.sUnitID[i].value;
	                }else{
	                    UnitName=UnitName + "," + form1.sUnitID[i].value;
	                }
	            }
	        }
	    }else{
	        if(form1.sUnitID.checked==true){
	            UnitName=form1.sUnitID.value;
	        }
	    }
	    if (UnitName==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計單位!!";
	    }
	    if (error==0){
	        var AnalyzeMoney=form1.AnalyzeMoney.value;
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        var sCountyOrNpa=form1.sCountyOrNpa2.value;
	        var sUnitID=UnitName;
	        var AnalyzeType=form1.AnalyzeType.value;
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
	        window.open("getRewardList_Unit.aspx?AnalyzeMoney="+AnalyzeMoney+"&Date1="+Date1+"&Date2="+Date2+"&sCountyOrNpa="+sCountyOrNpa+"&sUnitID="+sUnitID+"&AnalyzeType="+AnalyzeType+"&DateType="+DateType,"getRewardList_Unit2","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        window.close();
	    }else{
	        alert(errorString);
	    }	
	}
	//開啟Excel清冊
	function OpenRewardList_Excel(){
	    var error=0;
	    var errorString="";
	    var UnitName="";
	    CheckFlag1=dateCheck(form1.tbDate1.value);
	    CheckFlag2=dateCheck(form1.tbDate2.value);
	    if (CheckFlag1==false){
	        error=error+1;
		    errorString=error+"：起始日期輸入錯誤!!";
	    }
	    if (CheckFlag2==false){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：結束日期輸入錯誤!!";
	    }
	    if (form1.sCountyOrNpa2.value==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計標準!!";
	    }
	    if (form1.sUnitID.length > 0){
	        for (i=0; i< form1.sUnitID.length; i++){
	            if(form1.sUnitID[i].checked==true){
	                if(UnitName==""){
	                    UnitName=form1.sUnitID[i].value;
	                }else{
	                    UnitName=UnitName + "," + form1.sUnitID[i].value;
	                }
	            }
	        }
	    }else{
	        if(form1.sUnitID.checked==true){
	            UnitName=form1.sUnitID.value;
	        }
	    }
	    if (UnitName==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計單位!!";
	    }
	    if (error==0){
	        var AnalyzeMoney=form1.AnalyzeMoney.value;
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        var sCountyOrNpa=form1.sCountyOrNpa2.value;
	        var sUnitID=UnitName;
	        var AnalyzeType=form1.AnalyzeType.value;
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
	        window.open("getRewardList_Unit_Excel.aspx?AnalyzeMoney="+AnalyzeMoney+"&Date1="+Date1+"&Date2="+Date2+"&sCountyOrNpa="+sCountyOrNpa+"&sUnitID="+sUnitID+"&AnalyzeType="+AnalyzeType+"&DateType="+DateType,"getRewardList_Unit_Excel2","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        window.close();
	    }else{
	        alert(errorString);
	    }	    
	}
	function AllUnit(){
	    if(form1.sUnitID.length > 0){
	        for (i=0; i< form1.sUnitID.length; i++){
	            form1.sUnitID[i].checked=true;
	        }		        
	    }else{
	        form1.sUnitID.checked=true;
	    }
	}
	function NoUnit(){
	    if(form1.sUnitID.length > 0){
	        for (i=0; i< form1.sUnitID.length; i++){
	            form1.sUnitID[i].checked=false;
	        }		        
	    }else{
	        form1.sUnitID.checked=false;
	    }
	}
function btnDirector_onclick() {
var error=0;
	    var errorString="";
	    var UnitName="";
	    CheckFlag1=dateCheck(form1.tbDate1.value);
	    CheckFlag2=dateCheck(form1.tbDate2.value);
	    if (CheckFlag1==false){
	        error=error+1;
		    errorString=error+"：起始日期輸入錯誤!!";
	    }
	    if (CheckFlag2==false){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：結束日期輸入錯誤!!";
	    }
	    if (form1.sCountyOrNpa2.value==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計標準!!";
	    }
	    if (form1.sUnitID.length > 0){
	        for (i=0; i< form1.sUnitID.length; i++){
	            if(form1.sUnitID[i].checked==true){
	                if(UnitName==""){
	                    UnitName=form1.sUnitID[i].value;
	                }else{
	                    UnitName=UnitName + "," + form1.sUnitID[i].value;
	                }
	            }
	        }
	    }else{
	        if(form1.sUnitID.checked==true){
	            UnitName=form1.sUnitID.value;
	        }
	    }
	    if (UnitName==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計單位!!";
	    }
	    if (error==0){
	        var AnalyzeMoney=form1.AnalyzeMoney.value;
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        var sCountyOrNpa=form1.sCountyOrNpa2.value;
	        var sUnitID=UnitName;
	        var AnalyzeType=form1.AnalyzeType.value;
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
	        window.open("getRewardList_Balance_Director_Excel.aspx?AnalyzeMoney="+AnalyzeMoney+"&Date1="+Date1+"&Date2="+Date2+"&sCountyOrNpa="+sCountyOrNpa+"&sUnitID="+sUnitID+"&AnalyzeType="+AnalyzeType+"&DateType="+DateType,"getRewardList_Unit_Excel2","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        window.close();
	    }else{
	        alert(errorString);
	    }	    
	}
}

function btnCommon_onclick() {
var error=0;
	    var errorString="";
	    var UnitName="";
	    CheckFlag1=dateCheck(form1.tbDate1.value);
	    CheckFlag2=dateCheck(form1.tbDate2.value);
	    if (CheckFlag1==false){
	        error=error+1;
		    errorString=error+"：起始日期輸入錯誤!!";
	    }
	    if (CheckFlag2==false){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：結束日期輸入錯誤!!";
	    }
	    if (form1.sCountyOrNpa2.value==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計標準!!";
	    }
	    if (form1.sUnitID.length > 0){
	        for (i=0; i< form1.sUnitID.length; i++){
	            if(form1.sUnitID[i].checked==true){
	                if(UnitName==""){
	                    UnitName=form1.sUnitID[i].value;
	                }else{
	                    UnitName=UnitName + "," + form1.sUnitID[i].value;
	                }
	            }
	        }
	    }else{
	        if(form1.sUnitID.checked==true){
	            UnitName=form1.sUnitID.value;
	        }
	    }
	    if (UnitName==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計單位!!";
	    }
	    if (error==0){
	        var AnalyzeMoney=form1.AnalyzeMoney.value;
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        var sCountyOrNpa=form1.sCountyOrNpa2.value;
	        var sUnitID=UnitName;
	        var AnalyzeType=form1.AnalyzeType.value;
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
	        window.open("getRewardList_Balance_Common_Excel.aspx?AnalyzeMoney="+AnalyzeMoney+"&Date1="+Date1+"&Date2="+Date2+"&sCountyOrNpa="+sCountyOrNpa+"&sUnitID="+sUnitID+"&AnalyzeType="+AnalyzeType+"&DateType="+DateType,"getRewardList_Unit_Excel2","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        window.close();
	    }else{
	        alert(errorString);
	    }	    
	}
}

function btnItem_onclick() {
var error=0;
	    var errorString="";
	    var UnitName="";
	    CheckFlag1=dateCheck(form1.tbDate1.value);
	    CheckFlag2=dateCheck(form1.tbDate2.value);
	    if (CheckFlag1==false){
	        error=error+1;
		    errorString=error+"：起始日期輸入錯誤!!";
	    }
	    if (CheckFlag2==false){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：結束日期輸入錯誤!!";
	    }
	    if (form1.sCountyOrNpa2.value==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計標準!!";
	    }
	    if (form1.sUnitID.length > 0){
	        for (i=0; i< form1.sUnitID.length; i++){
	            if(form1.sUnitID[i].checked==true){
	                if(UnitName==""){
	                    UnitName=form1.sUnitID[i].value;
	                }else{
	                    UnitName=UnitName + "," + form1.sUnitID[i].value;
	                }
	            }
	        }
	    }else{
	        if(form1.sUnitID.checked==true){
	            UnitName=form1.sUnitID.value;
	        }
	    }
	    if (UnitName==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計單位!!";
	    }
	    if (error==0){
	        var AnalyzeMoney=form1.AnalyzeMoney.value;
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        var sCountyOrNpa=form1.sCountyOrNpa2.value;
	        var sUnitID=UnitName;
	        var AnalyzeType=form1.AnalyzeType.value;
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
	        window.open("getRewardList_Balance_Item_Excel.aspx?AnalyzeMoney="+AnalyzeMoney+"&Date1="+Date1+"&Date2="+Date2+"&sCountyOrNpa="+sCountyOrNpa+"&sUnitID="+sUnitID+"&AnalyzeType="+AnalyzeType+"&DateType="+DateType,"getRewardList_Unit_Excel2","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        window.close();
	    }else{
	        alert(errorString);
	    }	    
	}
}

</script>
</html>