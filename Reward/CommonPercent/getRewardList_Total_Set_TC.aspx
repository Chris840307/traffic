<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">

</script>
<script language="JavaScript">
	window.focus();
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>處理道路交通安全人員獎勵金發放一覽表</title>
</head>
<body>
    <form id="form1" runat="server">
        <table align="center" border="1" style="position: relative; top: 20px">
            <tr>
                <td style="width: 471px; background-color: #FFCC33" align="center" >
                    <strong>處理道路交通安全人員獎勵金發放一覽表</strong></td>
            </tr>
            <tr>
                <td align="center" style="width: 471px">
                    <input name="tbYear1" type="text" value="<%=trim(request("tbYear1"))%>" MaxLength="3" onkeyup="value=value.replace(/[^\d]/g,'')" style="width: 50px; font-size: 14pt; height: 25px;" />年
                    <input name="tbMonth1" type="text" value="<%=trim(request("tbMonth1"))%>" MaxLength="2" onkeyup="value=value.replace(/[^\d]/g,'')" style="width: 50px; font-size: 14pt; height: 25px;" />月
                    <br>
                    <input type="radio" name="DateType" value="BillFillDate" checked />填單日期&nbsp;
                    <input type="radio" name="DateType" value="RecordDate" />建檔日期
                </td>
            </tr>
            <tr style="background-color:#FFCC33; font-size: 12pt;">
                <td style="height: 25px" align="center">
                    配分標準</td>
            </tr>
            <tr style="font-size: 12pt">
                <td style="height: 35px" align="center">
                    <select name="sCountyOrNpa2" style="font-size: 12pt">
                        <option value="0" <%if trim(request("sCountyOrNpa2"))="0" then response.write("selected") %>>獎勵金</option>
                        <option value="1" <%if trim(request("sCountyOrNpa2"))="1" then response.write("selected") %>>績效</option>
                    </select>
                </td>
            </tr>
            <tr>
                <td align="center" style="width: 471px; background-color: #ccffff">
                <input type="button" value="匯出Excel檔" onclick="OpenRewardList_Excel();" style="font-size: 12pt; width: 102px; height: 30px" />
                <input type="button" value="離開" onclick="window.close();" style="font-size: 12pt; width: 50px; height: 30px" />
                <input type="hidden" name="AnalyzeMoney" value="<%=trim(request("AnalyzeMoney")) %>" />
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
	    CheckFlag1=form1.tbMonth1.value;
	    if (form1.tbYear1.value==""){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：統計年份未輸入!!";
	    }
	    if (CheckFlag1<1 || CheckFlag1>12){
	        error=error+1;
		    errorString=errorString+"\n"+error+"：統計月份輸入錯誤!!";
	    }

	    if (error==0){
	        var AnalyzeMoney=form1.AnalyzeMoney.value;
	        var Year1=form1.tbYear1.value;
	        var Month1=form1.tbMonth1.value;
	        var sCountyOrNpa=form1.sCountyOrNpa2.value;
            if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
            window.open("../CommonPercent/getRewardList_Total_TC.aspx?AnalyzeMoney="+AnalyzeMoney+"&Year1="+Year1+"&Month1="+Month1+"&DateType="+DateType+"&sCountyOrNpa="+sCountyOrNpa,"getRewardList_Totala","width=800,height=600,left=10,top=10,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        //window.close();
	    }else{
	        alert(errorString);
	    }	
    }
</script>
</html>
