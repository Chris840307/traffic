<%@ Page Language="VB" AutoEventWireup="false" CodeFile="getScoreList_SubUnit_set.aspx.vb" Inherits="ScoreList_getScoreList_SubUnit_set" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%  
    Server.ScriptTimeout = 86400
    Response.Flush()
%>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>法條別配分統計1</title>
</head>
<body>
    <form id="form1" runat="server">
        <table align="center" border="1" style="position: relative; top: 20px">
            <tr>
                <td style="width: 371px; background-color: #FFCC33" align="center" >
                    <strong>法條別配分統計日期</strong></td>
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
                <td style="width: 371px; background-color: #FFCC33" align="center" >
                    <strong>統計法條</strong></td>
            </tr>
            <tr>
                <td align="center" style="width: 371px">
                    <asp:Panel ID="Panel3" runat="server" Height="50px" HorizontalAlign="Left" Width="245px">
                    <input type="radio" name="RadioLaw" value="0" checked/>全部
                    <br>
                    <input type="radio" name="RadioLaw" value="1" />攔停、逕舉(1~68條)
                    <br>
                    <input type="radio" name="RadioLaw" value="2" />慢車行人道路障礙(69條之後)
                    </asp:Panel>
                </td>
            </tr>
            <tr>
                <td style="width: 371px; background-color: #FFCC33" align="center" >
                    <strong>配分標準</strong></td>
            </tr>
            <tr>
                <td align="center">
                    <select id="Select1" name="theCountyOrNpa">
                        <option value="1">績效</option>
                        <option value="0">獎勵金</option>
                    </select>
                </td>
            </tr>
            <tr>
                <td style="width: 371px; background-color: #FFCC33" align="center" >
                    <strong>統計單位</strong></td>
            </tr>
            <tr>
                <td align="center">
                    <input type="radio" name="AnalyzeType" value="0" checked />單位別
                    <input type="radio" name="AnalyzeType" value="1" />員警別
                </td>
            </tr>
            <tr>
                <td align="center" style="width: 371px">
                   <%-- <input id="Radio1" value="0" name="UnitFlag" type="radio" onclick="ReportFalg()" checked />單位<br />--%>
                    <span style="color: #0066ff">請選擇統計單位或輸入員警臂章號碼<br />
                    </span>
                    <asp:Panel ID="Panel1" runat="server" BorderStyle="Inset" Font-Size="12pt" Height="95px"
                        ScrollBars="Vertical" Width="245px" HorizontalAlign="Left">
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
                    <input id="Button3" style="width: 62px; height: 21px" type="button" value="全部選取" onclick="AllUnit();" />
                    <input id="Button4" style="width: 62px; height: 21px" type="button" value="全部取消" onclick="NoUnit();" /><br />
                    <br />
                    <asp:Panel ID="Panel2" runat="server" Height="30px" HorizontalAlign="Left" Width="265px">
                    臂章號碼：<input id="Text1" name="LoginID" style="width: 80px" type="text" onkeyup="QueryMemName();" />
                    <div id="Layer1" style="position:absolute ; width:110px; height:25px; z-index:0;  border: 1px none #000000; "></div>
                    <br><span style="font-size: 10pt; color: #0066ff">(計算單一員警請輸入員警個人代號)</span>
                    </asp:Panel>
                </td>
            </tr>
            <tr>
                <td align="center" style="width: 371px; background-color: #ccffff">
                <input type="button" value="匯出Excel檔" onclick="OpenRewardList_Excel();" style="font-size: 12pt; width: 102px; height: 30px" />
                <input type="button" value="列印" onclick="OpenRewardList();" style="font-size: 12pt; width: 50px; height: 30px" />
                <input type="button" value="離開" onclick="window.close();" style="font-size: 12pt; width: 50px; height: 30px" />
                <br>
                <span style="font-size: 10pt; color: #0066ff">計算獎勵金會需要較多時間等待，此為正常現象</span>
                </td>
            </tr>
        </table>
        
    </form>
</body>
<script type="text/javascript" src="../form.js"></script>
<script language="JavaScript">
    //宣告ajax物件，ie與其他瀏覽器宣告類別不同，故須做判斷處理
    var AjaxObj = false;
    if (window.XMLHttpRequest) { // Mozilla, Safari,...
        AjaxObj = new XMLHttpRequest();
        if (AjaxObj.overrideMimeType) {
            AjaxObj.overrideMimeType('text/xml');
        }
    } else if (window.ActiveXObject) { // IE
        try {
            AjaxObj = new ActiveXObject("Msxml2.XMLHTTP");
        } catch (e) {
            try {
                AjaxObj = new ActiveXObject("Microsoft.XMLHTTP");
            } catch (e) {}
        }
    }
    //抓員警姓名
    function QueryMemName(){
        MLoginID=form1.LoginID.value;
        
        if (MLoginID.length > 2){
            AjaxObj.Open("POST","QueryMemName.aspx",true);
            AjaxObj.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
	        AjaxObj.send("MLoginID=" + MLoginID);
	        AjaxObj.onreadystatechange=ServerUpdate;
	    }
    }
    function ServerUpdate()
    {
	    if (AjaxObj.readystate==4 || AjaxObj.readystate=='complete')
	    {
	        strName=AjaxObj.responseText;
		    document.getElementById('Layer1').innerHTML = strName;
		    //alert(AjaxObj.responsetext);
	    }
    }
    
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
	    if (form1.mUnitID.length > 0){
	        for (i=0; i< form1.mUnitID.length; i++){
	            if(form1.mUnitID[i].checked==true){
	                if(UnitName==""){
	                    UnitName=form1.mUnitID[i].value;
	                }else{
	                    UnitName=UnitName + "," + form1.mUnitID[i].value;
	                }
	            }
	        }
	    }else{
            if(form1.mUnitID.checked==true){
                if(UnitName==""){
                    UnitName=form1.mUnitID.value;
                }
            }
	    }

	    if (form1.AnalyzeType(0).checked==true){
	        if (UnitName==""){
	            error=error+1;
		        errorString=errorString+"\n"+error+"：請選擇統計單位!!";
	        }
        }else{
	        if (UnitName=="" && form1.LoginID.value==""){
	            error=error+1;
		        errorString=errorString+"\n"+error+"：請選擇統計單位!!";
	        }
	    }

	    if (error==0){
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        var sUnitID=UnitName;
	        var MemLoginID=form1.LoginID.value;
	        if (form1.RadioLaw(0).checked){
	            var LawRange="0";
	        }else if(form1.RadioLaw(1).checked){
	            var LawRange="1";
	        }else{
	        	var LawRange="2";
	        }
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
	        var theCountyOrNpa=form1.theCountyOrNpa.value;
	        if (form1.AnalyzeType(0).checked==true){
        	    window.open("getScoreList_SubUnit_Unit.aspx?Date1="+Date1+"&Date2="+Date2+"&sUnitID="+sUnitID+"&MemLoginID="+MemLoginID+"&LawRange="+LawRange+"&DateType="+DateType+"&theCountyOrNpa="+theCountyOrNpa,"getRewardList_Unit4","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        }else{
	            window.open("getScoreList_SubUnit.aspx?Date1="+Date1+"&Date2="+Date2+"&sUnitID="+sUnitID+"&MemLoginID="+MemLoginID+"&LawRange="+LawRange+"&DateType="+DateType+"&theCountyOrNpa="+theCountyOrNpa,"getRewardList_Unit3","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        }
	        //window.close();
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
        if (form1.mUnitID.length > 0){
	        for (i=0; i< form1.mUnitID.length; i++){
	            if(form1.mUnitID[i].checked==true){
	                if(UnitName==""){
	                    UnitName=form1.mUnitID[i].value;
	                }else{
	                    UnitName=UnitName + "," + form1.mUnitID[i].value;
	                }
	            }
	        }
	    }else{
            if(form1.mUnitID.checked==true){
                if(UnitName==""){
                    UnitName=form1.mUnitID.value;
                }
            }
	    }
	    if (form1.AnalyzeType(0).checked==true){
	        if (UnitName==""){
	            error=error+1;
		        errorString=errorString+"\n"+error+"：請選擇統計單位1!!";
	        }
        }else{
	        if (UnitName=="" && form1.LoginID.value==""){
	            error=error+1;
		        errorString=errorString+"\n"+error+"：請選擇統計單位!!";
	        }
	    }
	    if (error==0){
	        var Date1=form1.tbDate1.value;
	        var Date2=form1.tbDate2.value;
	        var sUnitID=UnitName;
	        var MemLoginID=form1.LoginID.value;
	        if (form1.RadioLaw(0).checked){
	            var LawRange="0";
	        }else if(form1.RadioLaw(1).checked){
	            var LawRange="1";
	        }else{
	        	var LawRange="2";
	        }
	        if (form1.DateType(0).checked==true){
	            var DateType=form1.DateType(0).value;
	        }else{
	            var DateType=form1.DateType(1).value;
	        }
	        var theCountyOrNpa=form1.theCountyOrNpa.value;
	        if (form1.AnalyzeType(0).checked==true){
	        	window.open("getScoreList_SubUnit_Unit_Excel.aspx?Date1="+Date1+"&Date2="+Date2+"&sUnitID="+sUnitID+"&MemLoginID="+MemLoginID+"&LawRange="+LawRange+"&DateType="+DateType+"&theCountyOrNpa="+theCountyOrNpa,"getRewardList_Unit_Excel4","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        }else{
	            window.open("getScoreList_SubUnit_Excel.aspx?Date1="+Date1+"&Date2="+Date2+"&sUnitID="+sUnitID+"&MemLoginID="+MemLoginID+"&LawRange="+LawRange+"&DateType="+DateType+"&theCountyOrNpa="+theCountyOrNpa,"getRewardList_Unit_Excel3","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
	        }
	        //window.close();
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
	
//	function ReportFalg(){
//	    if (form1.UnitFlag[0].checked==true){
//	        for (i=0; i< form1.mUnitID.length; i++){
//	            form1.mUnitID[i].disabled=false;
//	        }
//	        form1.LoginID.disabled=true;
//	    }else if (form1.UnitFlag[1].checked==true){
//	        for (i=0; i< form1.mUnitID.length; i++){
//	            form1.mUnitID[i].disabled=true;
//	        }
//	        form1.LoginID.disabled=false;
//	    }
//	}
</script>
</html>
