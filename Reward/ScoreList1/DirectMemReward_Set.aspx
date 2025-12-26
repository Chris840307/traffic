<%@ Page Language="VB" AutoEventWireup="false" CodeFile="DirectMemReward_Set.aspx.vb" Inherits="ScoreList1_DirectMemReward_Set" %>

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
<head id="Head1" runat="server">
     <title>直接人員每點點數金額日期設定</title>
</head>
<body style="text-align: center">
<form id="form1" runat="server">
    <div style="text-align: center">
        <table style="width: 400px" border="1" >
            <tr style="background-color:#FFCC33">
                <td align="center" style="height: 25px">
                    直接人員每點點數金額統計日期</td>
            </tr>
            <tr>
                <td align="center" style="height: 40px">
                    <input name="tbDate1" type="text" value="<%=trim(request("tbDate1"))%>" MaxLength="2" onkeyup="value=value.replace(/[^\d]/g,'')" style="width: 68px; font-size: 14pt; height: 19px;" />
                    &nbsp;年 &nbsp;<input name="tbDate2" value="<%=trim(request("tbDate2"))%>" type="text" MaxLength="2" onkeyup="value=value.replace(/[^\d]/g,'')" style="width: 54px; font-size: 14pt; height: 19px;" />
                    月<br>
                    <input type="radio" name="DateType" value="BillFillDate" <%
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
                
                conn.close()
                
                if sys_City<>"台中縣" then
                    response.write("checked")
                end if                
                     %> />填單日期&nbsp;
                    <input type="radio" name="DateType" value="RecordDate" <%
                 if sys_City="台中縣" then
                     response.write("checked")
                 end if
                    %>/>建檔日期
                </td>
            </tr>
            <tr style="background-color:#FFCC33">
                <td align="center" style="height: 25px">
                    配分標準
                </td>
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
                    <input id="Button7" type="button" value="列印" onclick="OpenRewardList();"  style="font-size: 12pt; width: 50px; height: 30px" />
                    <asp:Button ID="Button4" runat="server" Text="離開" OnClientClick="window.close();" Font-Size="12pt" Height="30px" Width="48px" />
                    <input type="hidden" name="AnalyzeType" value="<%=trim(request("AnalyzeType")) %>" />
                    <input type="hidden" name="AnalyzeMoney" value="<%=trim(request("AnalyzeMoney")) %>" />
                    <input type="hidden" name="AllAnalyzeMoney" value="<%=trim(request("AllAnalyzeMoney")) %>" />
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
    		
    //開啟萬年曆視窗
	function OpenSelectDate1(tag){
	    InitDate=eval("form1."+tag).value;
	    window.open("SelectDate.aspx?tag="+tag+"&InitDate="+InitDate,"OpenSelectDate1","width=240,height=240,left=350,top=250,scrollbars=no,menubar=no,resizable=no,fullscreen=no,status=no,toolbar=no");
	}
	
	//開啟清冊
	function OpenRewardList(){
	    
	    var error=0;
	    var errorString="";
	    var UnitName="";

	    CheckFlag1=form1.tbDate1.value;
	    CheckFlag2=form1.tbDate2.value;
	    if (CheckFlag1==""){
	        error=error+1;
		    errorString=error+"：請輸入統計年份!!";
	    }
	    if (CheckFlag2==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：請輸入統計月份!!";
	    }
	    if (error==0){
	        var AnalyzeMoney=form1.AnalyzeMoney.value;
	        var AllAnalyzeMoney=form1.AllAnalyzeMoney.value;
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        var AnalyzeType=form1.AnalyzeType.value;
	        var sCountyOrNpa=form1.sCountyOrNpa2.value;
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }

	        window.open("DirectMemReward.aspx?AllAnalyzeMoney="+AllAnalyzeMoney+"&AnalyzeMoney="+AnalyzeMoney+"&Date1="+Date1+"&Date2="+Date2+"&sCountyOrNpa="+sCountyOrNpa+"&AnalyzeType="+AnalyzeType+"&DateType="+DateType,"getRewardList_Point","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        window.close();
	    }else{
	        alert(errorString);
	    }	    
	}

</script>
</html>
