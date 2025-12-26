<%@ Page Language="VB" AutoEventWireup="false" CodeFile="getCommonScore_DateSet.aspx.vb" Inherits="CommonPercent_getCommonScore_DateSet" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%  
    Server.ScriptTimeout = 86400
    Response.Flush()
    
%>
<script language="JavaScript">
	window.focus();
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>共同人員獎勵金統計</title>
</head>
<body>
    <form id="form1" runat="server">
        <table align="center" border="1" style="position: relative; top: 20px">
            <tr>
                <td style="width: 371px; background-color: #FFCC33" align="center" >
                    <strong>共同人員獎勵金統計日期</strong></td>
            </tr>
            <tr>
                <td align="center" style="width: 371px">
                    <input name="tbDate1" type="text" value="<%=trim(request("tbDate1"))%>" MaxLength="6" onkeyup="value=value.replace(/[^\d]/g,'')" style="width: 70px" />
                    <input id="Button1" type="button" value=".." onclick="OpenSelectDate1('tbDate1')" />&nbsp;
                    至 &nbsp;<input name="tbDate2" value="<%=trim(request("tbDate2"))%>" type="text" MaxLength="6" onkeyup="value=value.replace(/[^\d]/g,'')" style="width: 70px" />
                    <input id="Button2" type="button" value=".." onclick="OpenSelectDate1('tbDate2')" />
                    <br>
                    <input type="radio" name="DateType" value="BillFillDate" checked />填單日期&nbsp;
                    <input type="radio" name="DateType" value="RecordDate" />建檔日期
                </td>
            </tr>

            <tr>
                <td align="center" style="width: 371px; background-color: #ccffff">
                <input type="button" value="匯出Excel檔" onclick="OpenRewardList_Excel();" style="font-size: 12pt; width: 102px; height: 30px" />
                <input type="button" value="離開" onclick="window.close();" style="font-size: 12pt; width: 50px; height: 30px" />
                <input type="hidden" name="AnalyzeMoney" value="<%=trim(request("AnalyzeMoney")) %>" />
                <br>
                <span style="font-size: 10pt; color: #0066ff">計算獎勵金會需要較多時間等待，此為正常現象</span>
                </td>
            </tr>
        </table>
        
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
    //開啟萬年曆視窗
	function OpenSelectDate1(tag){
	    InitDate=eval("form1."+tag).value;
	    window.open("../ScoreList2/SelectDate.aspx?tag="+tag+"&InitDate="+InitDate,"OpenSelectDate1","width=240,height=240,left=350,top=250,scrollbars=no,menubar=no,resizable=no,fullscreen=no,status=no,toolbar=no");
	}
	function OpenRewardList_Excel(){
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
	    if (error==0){
	        var AnalyzeMoney=0;
            AnalyzeMoney=form1.AnalyzeMoney.value;
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
            window.open("../CommonPercent/getCommonScore_Excel.aspx?AnalyzeMoney="+AnalyzeMoney+"&Date1="+Date1+"&Date2="+Date2+"&DateType="+DateType,"getCommonScore_Print15","width=800,height=600,left=10,top=10,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        window.close();
	    }else{
	        alert(errorString);
	    }	
    }
</script>
</html>
