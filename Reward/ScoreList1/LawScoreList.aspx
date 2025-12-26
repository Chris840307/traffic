<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%  
    LoginCheck()
%>
<script runat="server">
    Public AllReward As Integer = 0
    Public sys_City As String = ""
    
    Sub LoginCheck()
        If (Request.Cookies("UserFunction") IsNot Nothing) Then
            Dim FuncCookie As HttpCookie = Request.Cookies("UserFunction")
            If Trim(FuncCookie.Values("FuncID")) = "" Then
                Response.Redirect("/traffic/Reward/Login.aspx?ErrMsg=1")
            End If
        Else
            Response.Redirect("/traffic/Reward/Login.aspx?ErrMsg=1")
        End If
    End Sub
</script>
<script language="JavaScript">
	window.focus();
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>積效獎勵金統計報表1</title>
</head>
<body style="font-size: 12pt">
    <form id="form1" runat="server">
    <center>
        <div style="width: 317px; height: 363px; text-align: left; background-color: transparent; border-top-width: thick; border-left-width: thick; border-bottom-width: thick; border-right-width: thick;">
            <span style="font-family: 新細明體"><strong>
                <br />
                <span style="font-size: 14pt">
        獎勵金總額：</span></strong></span>
            <input Name="RewardTotal" value="<% 
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
            
        Dim strCity = "select Value from ApConfigure where ID=31"
        Dim CmdCity As New Data.OracleClient.OracleCommand(strCity, conn)
        Dim rdCity As Data.OracleClient.OracleDataReader = CmdCity.ExecuteReader()
        If rdCity.HasRows Then
            rdCity.Read()
            sys_City = Trim(rdCity("Value"))
        End If
        rdCity.Close()
    
        Dim strRewardValue = "select Value from ApConfigure where ID=46"
        Dim CmdRewardValue As New Data.OracleClient.OracleCommand(strRewardValue, conn)
        Dim rdRewardValue As Data.OracleClient.OracleDataReader = CmdRewardValue.ExecuteReader()
        If rdRewardValue.HasRows Then
            rdRewardValue.Read()

            AllReward=Trim(rdRewardValue("Value"))
            response.write(AllReward)
        End If
        rdRewardValue.Close()
        
            %>" style="width: 90px" onkeyup="value=value.replace(/[^\d]/g,'')" type="text" />
            <span style="font-family: 新細明體">元</span>
        &nbsp;<input type="button" value="確定" onclick="OpenRewardSet();" style="width: 50px; height: 25px; font-size: 12pt;" />
        <br />
            <hr />
            <table style="width: 280px">
                <tr>
                    <td style="width: 105px">
                        <span style="font-family: 新細明體">共同人員 28 %</span></td>
                    <td style="width: 126px" align="right">
                        <input name="divMoney1a" type="text" value="<%=Format(Decimal.Round(AllReward * 0.28), "##,##0")%>" size="10" />
                        &nbsp;元
                        
                    </td>
                </tr>
                <tr>
                    <td style="width: 105px">
                        <span style="font-family: 新細明體">直接人員 72 %</span></td>
                    <td align="right" style="width: 126px">
                        <input name="divMoney2a" type="text" value="<%=Format(Decimal.Round(AllReward * 0.72), "##,##0")%>" size="10" />
                        &nbsp;元
                        
                    </td>
                </tr>
            </table>
            <hr id="HR3" />
            <input type="button" value="直接人員每點點數金額" onclick="getDirectRewardPoint();" style="width: 294px; height: 30px; font-size: 11pt;" /><br />
            <input type="button" value="獎勵金發放一覽表" onclick="getRewardList_Total();" style="width: 294px; height: 30px; font-size: 11pt;" /><br />
            <hr id="HR2" />
            <strong>計算方式<br />
                <input name="radioAnalyzeType" checked="checked" type="radio" value="0" /></strong>統計所有單位<br />
            <input name="radioAnalyzeType" type="radio" value="1" <%
            Dim UserCookie As HttpCookie = Request.Cookies("RewardUser")
            Dim AnalyzeUnitID As String = ""
            AnalyzeUnitID = Trim(UserCookie.Values("UnitID"))
    
            dim TrafficUnitID as string
            Dim strTrafficUnitID = "select Value from ApConfigure where ID=49"
            Dim CmdTrafficUnitID As New Data.OracleClient.OracleCommand(strTrafficUnitID, conn)
            Dim rdTrafficUnitID As Data.OracleClient.OracleDataReader = CmdTrafficUnitID.ExecuteReader()
            If rdTrafficUnitID.HasRows Then
                rdTrafficUnitID.Read()
                TrafficUnitID=Trim(rdTrafficUnitID("Value"))
            End If
            rdTrafficUnitID.Close()
            
           
            if AnalyzeUnitID<>TrafficUnitID then
                response.write("disabled")
            end if
             %>/>僅統計交通隊<br />
            <input name="radioAnalyzeType" type="radio" value="2" <%                  
            if AnalyzeUnitID=TrafficUnitID then
                response.write("disabled")
            end if
             %>/>僅統計分局及其管轄派出所<br />
            <hr id="HR1" />
            <span style="font-size: 12pt"><strong><span style="font-family: 新細明體">直接人員獎勵金計算</span><br />
            </strong></span><span style="font-family: 新細明體">獎勵金不超過個人薪資百分之</span>
            <input name="PayPercent" type="text" value="<%
           
        Dim strPayValue = "select Value from ApConfigure where ID=47"
        Dim CmdPayValue As New Data.OracleClient.OracleCommand(strPayValue, conn)
        Dim rdPayValue As Data.OracleClient.OracleDataReader = CmdPayValue.ExecuteReader()
        If rdPayValue.HasRows Then
            rdPayValue.Read()
            response.write(Trim(rdPayValue("Value")))
        End If
        rdPayValue.Close()
        
        conn.Close() 
            
            
             %>" style="width: 31px" />
            <input id="Button1" style="width: 50px; height: 25px; font-size: 12pt;" type="button" value="確定" onclick="OpenPaySet();" /><br />
            <br />
        <input type="button" value="支領獎勵金核發清冊（單位）" onclick="getRewardList_Unit();" style="width: 300px; height: 30px; font-size: 11pt;" />&nbsp;
            <br />
             <input type="button" value="支領獎勵金核發清冊（單位）分局別" onclick="getRewardList_Unit_SubUnit();" style="width: 300px; height: 30px; font-size: 11pt;" />&nbsp;
        <input id="Button2" type="button" value="支領獎勵金核發清冊（個人）" onclick="getRewardList_Person();" style="width: 300px; height: 30px; font-size: 11pt;" />
            <br />
        <input id="Button6" type="button" value="支領獎勵金核發清冊（個人）分局別" onclick="getRewardList_Person_SubUnit();" style="width: 300px; height: 30px; font-size: 11pt;" />
            <br />
            <hr />
            &nbsp;<strong>共同人員獎勵金計算<br />
                <input id="Button3" style="font-size: 11pt; width: 300px; height: 30px" type="button"
                    value="共同人員分配比率設定" onclick="Percent_Set();" /><br />
                <%--<input id="Button4" style="font-size: 11pt; width: 300px; height: 30px" type="button"
                    value="共同人員獎勵金核發清冊(列印)" onclick="getCommonScore_Print();" />--%>
                <%--<input id="Button5" style="font-size: 11pt; width: 300px; height: 30px" type="button"
                    value="共同人員獎勵金核發清冊(匯出Excel檔)" onclick="getCommonScore_Excel();" /> --%>
                <input id="Button17" style="font-size: 11pt; width: 300px; height: 30px" type="button"
                    value="共同人員獎勵金印領清冊" onclick="getCommonScore_Person();" />    
                </strong>
                <br /><hr />
            <input type="button" value="獎勵金結餘報表" onclick="getRewardList_Should();" style="width: 299px; height: 30px; font-size: 11pt;" /><br />
           <%-- <hr />
            
                <input id="Button4" style="font-size: 11pt; width: 300px; height: 30px" type="button"
                    value="處理道路交通安全人員獎勵金分配統計表" onclick="getRewardAnalyze1();" />  --%>
                </div>
    </center>
    </form>
</body>
<script language="JavaScript">
//宣告ajax物件，ie與其他瀏覽器宣告類別不同，故須做判斷處理
    var AjaxObj = false;
    if (window.XMLHttpRequest) { // Mozilla, Safari,...
        AjaxObj = new ActiveXObject("Msxml2.XMLHTTP");
    } else if (window.ActiveXObject) { // IE
        try {
            AjaxObj = new ActiveXObject("Msxml2.XMLHTTP");
        } catch (e) {
            try {
                AjaxObj = new ActiveXObject("Microsoft.XMLHTTP");
            } catch (e) {}
        }
    }
    //儲存獎勵金金額
    function OpenRewardSet(){
        RewardTotal=form1.RewardTotal.value;

        if (RewardTotal==""){
            alert("請輸入獎勵金總額!!");
        }else{
	        AjaxObj.Open("POST","RewardTotalSet.aspx",true);
            AjaxObj.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
	        AjaxObj.send("RewardTotal=" + RewardTotal);
	        AjaxObj.onreadystatechange=ServerUpdate;
	    }
    }
    function ServerUpdate()
    {
	    if (AjaxObj.readystate==4 || AjaxObj.readystate=='complete')
	    {
	        strMoney=AjaxObj.responseText;
	        mArray=strMoney.split("@@");
	        form1.divMoney1a.value = mArray[0];
	        form1.divMoney2a.value = mArray[1];
		    //document.getElementById('divMoney1').innerHTML = mArray[0] + " 元";
		    //document.getElementById('divMoney2').innerHTML = mArray[1] + " 元";
		    //alert(AjaxObj.responsetext);
	    }
    }
    //儲存個人薪資百分比
    function OpenPaySet(){
        PayPercent=form1.PayPercent.value;

        if (PayPercent==""){
            alert("請輸入獎勵金總額!!");
        }else{
	        AjaxObj.Open("POST","PayPercentSet.aspx",true);
            AjaxObj.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
	        AjaxObj.send("PayPercent=" + PayPercent);
	        AjaxObj.onreadystatechange=PayPercentUpdate;
	    }
    }
    function PayPercentUpdate()
    {
	    if (AjaxObj.readystate==4 || AjaxObj.readystate=='complete')
	    {
	        alert(AjaxObj.responsetext);
	    }
    }
    //單位清冊
    function getRewardList_Unit(){
        var AnalyzeType=0;
        if (form1.radioAnalyzeType(0).checked == true){
            AnalyzeType=0;
        }else if (form1.radioAnalyzeType(1).checked == true){
            AnalyzeType=1;
        }else if (form1.radioAnalyzeType(2).checked == true){
            AnalyzeType=2;
        }
        var AnalyzeMoney=0;
        AnalyzeMoney=form1.divMoney2a.value;
        window.open("getRewardList_Unit_Set.aspx?AnalyzeType="+AnalyzeType+"&AnalyzeMoney="+AnalyzeMoney,"getRewardList_Unit","width=420,height=500,left=250,top=100,scrollbars=no,menubar=no,resizable=no,fullscreen=no,status=no,toolbar=no");
    }
    //個人清冊
    function getRewardList_Person(){
        var AnalyzeType=0;
        if (form1.radioAnalyzeType(0).checked == true){
            AnalyzeType=0;
        }else if (form1.radioAnalyzeType(1).checked == true){
            AnalyzeType=1;
        }else if (form1.radioAnalyzeType(2).checked == true){
            AnalyzeType=2;
        }
        var AnalyzeMoney=0;
        AnalyzeMoney=form1.divMoney2a.value;
        window.open("getRewardList_Person_Set.aspx?AnalyzeType="+AnalyzeType+"&AnalyzeMoney="+AnalyzeMoney,"getRewardList_Person","width=420,height=480,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
    }
    //個人清冊分局別
    function getRewardList_Person_SubUnit(){
        var AnalyzeType=0;
        if (form1.radioAnalyzeType(0).checked == true){
            AnalyzeType=0;
        }else if (form1.radioAnalyzeType(1).checked == true){
            AnalyzeType=1;
        }else if (form1.radioAnalyzeType(2).checked == true){
            AnalyzeType=2;
        }
        var AnalyzeMoney=0;
        AnalyzeMoney=form1.divMoney2a.value;
        window.open("getRewardList_Person_SubUnit_Set.aspx?AnalyzeType="+AnalyzeType+"&AnalyzeMoney="+AnalyzeMoney,"getRewardList_Person","width=420,height=480,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
    }
    //共同人員分配比率設定
    function Percent_Set(){
   <%
  if sys_City="台中縣" then
   %>
        window.open("../CommonPercent/CommonPercent_Set_TC.aspx","CommonPercent_Set1","width=800,height=600,left=10,top=10,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
   <%
  else
  %>
        window.open("../CommonPercent/CommonPercent_Set.aspx","CommonPercent_Set1","width=800,height=600,left=10,top=10,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");  
  <%
  end if
  
   %>
    }
    //共同人員統計 列印
    function getCommonScore_Print(){
        window.open("../CommonPercent/getCommonScore_Print.aspx","getCommonScore_Print1","width=800,height=600,left=10,top=10,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
    }
    //共同人員統計 excel
    function getCommonScore_Excel(){
        var AnalyzeType=0;
        AnalyzeMoney=form1.RewardTotal.value;
        window.open("../CommonPercent/getCommonScore_DateSet.aspx?AnalyzeMoney="+AnalyzeMoney,"getCommonScore_Excel1","width=420,height=400,left=250,top=100,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
    }
    //處理道路交通安全人員獎勵金分配統計表
    function getRewardAnalyze1(){
        window.open("../CommonPercent/RewardAnalyze_1_Set.aspx","getRewardAnalyze_1_Set1","width=520,height=300,left=100,top=100,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
    }
    
    //獎勵金發放一覽表
    function getRewardList_Total(){
        var AnalyzeType=0;
        AnalyzeMoney=form1.RewardTotal.value;
  <%
  if sys_City="台中縣" then
   %>
        window.open("../CommonPercent/getRewardList_Total_Set_TC.aspx?AnalyzeMoney="+AnalyzeMoney,"getRewardList_Total_Set1","width=520,height=350,left=250,top=100,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
  <%
  else
  %>
        window.open("../CommonPercent/getRewardList_Total_Set.aspx?AnalyzeMoney="+AnalyzeMoney,"getRewardList_Total_Set1","width=520,height=350,left=250,top=100,scrollbars=yes,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
  <%
  end if
  
   %>
    }
    
    //直接人員每點點數金額
    function getDirectRewardPoint(){
        var AnalyzeType=0;
        if (form1.radioAnalyzeType(0).checked == true){
            AnalyzeType=0;
        }else if (form1.radioAnalyzeType(1).checked == true){
            AnalyzeType=1;
        }else if (form1.radioAnalyzeType(2).checked == true){
            AnalyzeType=2;
        }
        var AnalyzeMoney=0;
        AnalyzeMoney=form1.divMoney2a.value;
        AllAnalyzeMoney=form1.RewardTotal.value;
        window.open("DirectMemReward_Set.aspx?AllAnalyzeMoney="+AllAnalyzeMoney+"&AnalyzeType="+AnalyzeType+"&AnalyzeMoney="+AnalyzeMoney,"getDirectRewardPoint_set","width=420,height=480,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
    }
    
    //結餘
    function getRewardList_Should(){
        window.open("../CommonPercent/getRewardList_Should_Unit_Set.aspx","getDirectRewardPoint_set","width=420,height=480,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
    }
    
    //共同人員獎勵金印領清冊
    function getCommonScore_Person(){
        var AnalyzeType=0;
        if (form1.radioAnalyzeType(0).checked == true){
            AnalyzeType=0;
        }else if (form1.radioAnalyzeType(1).checked == true){
            AnalyzeType=1;
        }else if (form1.radioAnalyzeType(2).checked == true){
            AnalyzeType=2;
        }
        var AnalyzeMoney=0;
        AnalyzeMoney=form1.divMoney1a.value;
        AllAnalyzeMoney=form1.RewardTotal.value;
        window.open("getCommonPercent_Person_Set.aspx?AllAnalyzeMoney="+AllAnalyzeMoney+"&AnalyzeType="+AnalyzeType+"&AnalyzeMoney="+AnalyzeMoney,"getDirectRewardPoint_set","width=420,height=480,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
    }
    
    //
    function getRewardList_Unit_SubUnit(){
            var AnalyzeType=0;
        if (form1.radioAnalyzeType(0).checked == true){
            AnalyzeType=0;
        }else if (form1.radioAnalyzeType(1).checked == true){
            AnalyzeType=1;
        }else if (form1.radioAnalyzeType(2).checked == true){
            AnalyzeType=2;
        }
        var AnalyzeMoney=0;
        AnalyzeMoney=form1.divMoney1a.value;
        AllAnalyzeMoney=form1.RewardTotal.value;
        window.open("getRewardList_Unit_SubUnit_Set.aspx?AllAnalyzeMoney="+AllAnalyzeMoney+"&AnalyzeType="+AnalyzeType+"&AnalyzeMoney="+AnalyzeMoney,"getDirectRewardPoint_set","width=420,height=480,left=250,top=100,scrollbars=no,menubar=no,resizable=yes,fullscreen=no,status=no,toolbar=no");
    }
</script>
</html>
