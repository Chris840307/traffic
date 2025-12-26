<%@ Page Language="VB" AutoEventWireup="false" CodeFile="getRewardList_Person_Set.aspx.vb" Inherits="ScoreList_getRewardList_Person_Set" %>

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
     <title>支領獎勵金核發清冊(個人)</title>
</head>
<body style="text-align: center">
<form id="form1" runat="server">
    <div style="text-align: center">
        <table style="width: 400px" border="1" >
            <tr style="background-color:#FFCC33">
                <td align="center" style="height: 25px">
                    獎勵金核發清冊統計日期</td>
            </tr>
            <tr>
                <td align="center" style="height: 40px">
                    <input name="tbDate1" type="text" value="<%=trim(request("tbDate1"))%>" MaxLength="7" onkeyup="value=value.replace(/[^\d]/g,'')" style="width: 70px" />
                    <input id="Button1" type="button" value=".." onclick="OpenSelectDate1('tbDate1')" />&nbsp;
                    至 &nbsp;<input name="tbDate2" value="<%=trim(request("tbDate2"))%>" type="text" MaxLength="7" onkeyup="value=value.replace(/[^\d]/g,'')" style="width: 70px" />
                    <input id="Button2" type="button" value=".." onclick="OpenSelectDate1('tbDate2')" />
                    <br>
                    <input type="radio" name="DateType" value="BillFillDate" checked />填單日期&nbsp;
                    <input type="radio" name="DateType" value="RecordDate" <%if trim(request("DateType"))="RecordDate" then response.write("checked") %>/>建檔日期
                </td>
            </tr>
            <tr style="background-color:#FFCC33">
                <td align="center" style="height: 25px">統計單位</td>
            </tr>
            <tr>
                <td align="center">
                    <asp:RadioButtonList ID="RadioButtonList1" runat="server" Height="20px" RepeatDirection="Horizontal"
                        Width="166px" AutoPostBack="True" >
                        <asp:ListItem Value="0" Selected="True">單位</asp:ListItem>
                        <asp:ListItem Value="1">個人</asp:ListItem>
                    </asp:RadioButtonList><asp:MultiView ID="MultiView1" ActiveViewIndex="0" runat="server" OnLoad="MultiView1_OnLoad">
            <asp:View ID="View1" runat="server">
            <asp:Panel ID="Panel2" runat="server" Height="95px" HorizontalAlign="Left"
                        ScrollBars="Vertical" Width="245px" BorderStyle="Inset" Font-Size="12pt">
                       
<%
    '取得 Web.config 檔的資料連接設定
    Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
    '建立 Connection 物件
    Dim conn As New Data.OracleClient.OracleConnection()
    conn.ConnectionString = setting.ConnectionString
    '開啟資料連接
    conn.Open()
        
    Dim IsChecked As String = ""
    
    Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
    Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
    If rdUnit.HasRows Then
        While rdUnit.Read()

            If InStr(Trim(Request("mUnitID")), Trim(rdUnit("UnitID"))) <> 0 Then
                IsChecked = " checked"
            Else
                IsChecked = ""
            End If
            Response.Write("<input type=""checkbox"" name=""mUnitID"" value=""" & "'" & Trim(rdUnit("UnitID")) & "'" & """ " & IsChecked & " />" & Trim(rdUnit("UnitName")) & "<br />")
            
        End While
    End If
    rdUnit.Close()
    conn.Close()
%>
                    </asp:Panel>
                <input id="Button5" style="font-size: 8pt; width: 55px; height: 20px" type="button"
                    value="全部選取" onclick="AllUnit()" />
                &nbsp;
                <input id="Button1" style="font-size: 8pt; width: 55px; height: 20px" type="button"
                    value="全部取消" onclick="NoUnit()" /></asp:View>
            <asp:View ID="View2" runat="server">
            <br />
            <br />
            單位：<select name="sUnitID" onchange="getMemberID();" style="font-size: 12pt">
                    <option value="" selected="selected">請選擇</option>
<%
    '取得 Web.config 檔的資料連接設定
    Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
    '建立 Connection 物件
    Dim conn As New Data.OracleClient.OracleConnection()
    conn.ConnectionString = setting.ConnectionString
    '開啟資料連接
    conn.Open()
    Dim IsSelect As String = ""
   
    Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
    Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
    If rdUnit.HasRows Then
        While rdUnit.Read()
            If Trim(Request("sUnitID")) = "'" & Trim(rdUnit("UnitID")) & "'" Then
                IsSelect = "selected"
            Else
                IsSelect = ""
            End If
            Response.Write("<option value=""" & "'" & Trim(rdUnit("UnitID")) & "'" & """ " & IsSelect & ">" & Trim(rdUnit("UnitName")) & "</option>")
        End While

    End If
        
%>
                    </select>
              姓名：<select name="sMemID" style="font-size: 12pt">
                    <option value="" selected="selected">全部</option>
<%

    If Trim(Request("sUnitID")) <> "" Then
        Dim strMem = "select * from MemberData where UnitID=" & Trim(Request("sUnitID")) & " order by CHName"
        Dim CmdMem As New Data.OracleClient.OracleCommand(strMem, conn)
        Dim rsMem As Data.OracleClient.OracleDataReader = CmdMem.ExecuteReader()
        If rsMem.HasRows Then
            While rsMem.Read()
                If Trim(Request("sMemID")) = Trim(rsMem("MemberID")) Then
                    IsSelect = "selected"
                Else
                    IsSelect = ""
                End If
                Response.Write("<option value=""" & Trim(rsMem("MemberID")) & """ " & IsSelect & ">" & Trim(rsMem("CHName")) & "</option>")
            End While

        End If
        rsMem.Close()
    End If
    conn.Close()
%>
                    </select>
                    <br />
                    <br />
                    <br />
            </asp:View>
        </asp:MultiView></td>
            </tr>
            <tr style="background-color:#FFCC33">
                <td align="center" style="height: 25px">
                    配分標準</td>
            </tr>
            <tr>
                <td align="center" style="height: 40px">
                    <select name="sCountyOrNpa2" style="font-size: 12pt">
                        <option value="0" <%if trim(request("sCountyOrNpa2"))="0" then response.write("selected") %>>獎勵金</option>
                        <option value="1" <%if trim(request("sCountyOrNpa2"))="1" then response.write("selected") %>>績效</option>
                    </select>
                </td>
            </tr>
            <tr style="background-color:#F0FFFF">
                <td align="center">
                    <input id="Button6" type="button" value="匯出Excel檔"  onclick="OpenRewardList_Excel();" style="width: 90px; height: 30px" />
                    <input id="Button7" type="button" value="列印" onclick="OpenRewardList();"  style="font-size: 12pt; width: 50px; height: 30px" />
                    <asp:Button ID="Button4" runat="server" Text="離開" OnClientClick="window.close();" Font-Size="12pt" Height="30px" Width="48px" />
                    <input type="hidden" name="AnalyzeType" value="<%=trim(request("AnalyzeType")) %>" />
                    <input type="hidden" name="AnalyzeMoney" value="<%=trim(request("AnalyzeMoney")) %>" />
                    <br />
                    <input type="checkbox" name="SaveFlag" value="1" /><span style="font-size: 10pt">儲存為統計月份道路交通安全人員獎勵金分配統計表</span>
                </td>
            </tr>
        </table>
    </div>
    <span style="font-size: 10pt; color: #0066ff">計算獎勵金會需要較多時間等待，此為正常現象 </span>
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">

    //取得單位人員
    function getMemberID(){
        Unit1=form1.sUnitID.value;
        {
	        form1.submit();
	    }
    }
    		
    //開啟萬年曆視窗
	function OpenSelectDate1(tag){
	    InitDate=eval("form1."+tag).value;
	    window.open("SelectDate.aspx?tag="+tag+"&InitDate="+InitDate,"OpenSelectDate1","width=240,height=240,left=350,top=250,scrollbars=no,menubar=no,resizable=no,fullscreen=no,status=no,toolbar=no");
	}
	//開啟Excel清冊
	function OpenRewardList_Excel(){
	    
	    var error=0;
	    var errorString="";
	    var UnitName="";
	    var UnitFlag=<%=UnitFlag %>;
	    var sUnitID="";
	    var sMemID="";
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
	    //alert(UnitFlag);
	    if (UnitFlag=="0"){
	        if (form1.mUnitID.length > 0){
	            for (i=0; i< form1.mUnitID.length; i++){
	                if(form1.mUnitID[i].checked==true){
	                    if(UnitName==""){
	                        UnitName=form1.mUnitID[i].value;
	                    }else{
	                        UnitName=UnitName + "," + form1.mUnitID[i].value;
	                    }
	                }
	                sUnitID=UnitName;
	                sMemID="";
	            }
	        }else{
	            if(form1.mUnitID.checked==true){
	                sUnitID=form1.mUnitID.value;
	                sMemID="";
	            }
	        }
	    }else{
	        sUnitID=form1.sUnitID.value;
	        sMemID=form1.sMemID.value;
	    }
	    if (sUnitID==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計單位!!";
	    }
	    if (error==0){
	        var AnalyzeMoney=form1.AnalyzeMoney.value;
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        var AnalyzeType=form1.AnalyzeType.value;
	        var sCountyOrNpa=form1.sCountyOrNpa2.value;
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
	        if (form1.SaveFlag.checked==true){
	            var SaveFlag="1";
	        }else{
	            var SaveFlag="0";
	        }
	        window.open("getRewardList_Person_Excel.aspx?AnalyzeMoney="+AnalyzeMoney+"&Date1="+Date1+"&Date2="+Date2+"&sCountyOrNpa="+sCountyOrNpa+"&sUnitID="+sUnitID+"&sMemID="+sMemID+"&AnalyzeType="+AnalyzeType+"&DateType="+DateType+"&SaveFlag="+SaveFlag,"getRewardList_Person_Excel","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        window.close();
	    }else{
	        alert(errorString);
	    }	    
	}
	
	function AllUnit(){
	    for (i=0; i< form1.mUnitID.length; i++){
	        form1.mUnitID[i].checked=true;
	    }	
	}
	function NoUnit(){
	    for (i=0; i< form1.mUnitID.length; i++){
	        form1.mUnitID[i].checked=false;
	    }	
	}
	
	//開啟清冊
	function OpenRewardList(){
	    
	    var error=0;
	    var errorString="";
	    var UnitName="";
	    var UnitFlag=<%=UnitFlag %>;
	    var sUnitID="";
	    var sMemID="";
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
	    //alert(UnitFlag);
	    if (UnitFlag=="0"){
	        if (form1.mUnitID.length > 0){
	            for (i=0; i< form1.mUnitID.length; i++){
	                if(form1.mUnitID[i].checked==true){
	                    if(UnitName==""){
	                        UnitName=form1.mUnitID[i].value;
	                    }else{
	                        UnitName=UnitName + "," + form1.mUnitID[i].value;
	                    }
	                }
	                sUnitID=UnitName;
	                sMemID="";
	            }
	        }else{
	            if(form1.mUnitID.checked==true){
	                sUnitID=form1.mUnitID.value;
	                sMemID="";
	            }
	        }
	    }else{
	        sUnitID=form1.sUnitID.value;
	        sMemID=form1.sMemID.value;
	    }
	    if (sUnitID==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請選擇統計單位!!";
	    }
	    if (error==0){
	        var AnalyzeMoney=form1.AnalyzeMoney.value;
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        var AnalyzeType=form1.AnalyzeType.value;
	        var sCountyOrNpa=form1.sCountyOrNpa2.value;
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
	        if (form1.SaveFlag.checked==true){
	            var SaveFlag="1";
	        }else{
	            var SaveFlag="0";
	        }
	        window.open("getRewardList_Person.aspx?AnalyzeMoney="+AnalyzeMoney+"&Date1="+Date1+"&Date2="+Date2+"&sCountyOrNpa="+sCountyOrNpa+"&sUnitID="+sUnitID+"&sMemID="+sMemID+"&AnalyzeType="+AnalyzeType+"&DateType="+DateType+"&SaveFlag="+SaveFlag,"getRewardList_Person3","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        window.close();
	    }else{
	        alert(errorString);
	    }	    
	}
	
	function AllUnit(){
	    if(form1.mUnitID.length > 0){
	        for (i=0; i< form1.mUnitID.length; i++){
	            form1.mUnitID[i].checked=true;
	        }		        
	    }else{
	        form1.mUnitID.checked=true;
	    }
	}
	function NoUnit(){
	    if(form1.mUnitID.length > 0){
	        for (i=0; i< form1.mUnitID.length; i++){
	            form1.mUnitID[i].checked=false;
	        }		        
	    }else{
	        form1.mUnitID.checked=false;
	    }
	}
</script>
</html>
