<%@ Page Language="VB" AutoEventWireup="false" Debug="true" CodeFile="LawScore.aspx.vb" Inherits="LawScore_LawScore" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script language="JavaScript">
	window.focus();
</script>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>法條配分設定</title>
</head>
<body>
    <form id="form1" runat="server">
     <table width="985" border="0" align="left" style="background-color:#E0E0E0">
        <tr style="background-color:#1BF5FF">
<%
    If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
        Response.Write("<td align=""left"" colspan=""12"">")
    Else
        Response.Write("<td align=""left"" colspan=""11"">")
    End If
 %>
        
            <span style="font-size: 14pt"><strong>法條配分設定</strong></span> &nbsp;
            <asp:LinkButton ID="LinkButton1" runat="server" Font-Size="11pt" PostBackUrl="~/LawScore/LawScoreSetAll.aspx">統一設定所有法條配分</asp:LinkButton></td>
        </tr>
        <tr style="background-color:#FFFFFF">
<%
    If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
        Response.Write("<td align=""left"" colspan=""12"">")
    Else
        Response.Write("<td align=""left"" colspan=""11"">")
    End If
 %>
        
        法條代碼
            <asp:TextBox ID="LawID" runat="server" Width="455px" onBlur="value=checkValue(value);"></asp:TextBox>
            &nbsp; 
        版本
            <asp:TextBox ID="LawVer" runat="server" Width="42px" onKeyUp="value=value.replace(/[^\d]/g,'')" MaxLength="3"></asp:TextBox>
            &nbsp; &nbsp;
        配分標準
            <select name="sCountyOrNpa">
                <option value="0" <%if trim(request("sCountyOrNpa"))="0" then response.write("selected")%>>獎勵金</option>
                <option value="1" <%if trim(request("sCountyOrNpa"))="1" then response.write("selected")%>>績效</option>
                <option value="n" <%if trim(request("sCountyOrNpa"))="n" then response.write("selected")%>>全部</option>
            </select>
            &nbsp; &nbsp;<asp:Button ID="Button1" runat="server" Text="查詢" Width="61px" OnClick="LawView" Font-Size="12pt" />&nbsp;<br />
            <span style="font-size: 11pt; color: #0000cc"><img src="../image/space.gif" height="8" style="width: 76px">( 可用逗號間格輸入多筆法條，例如：1210101,3310101,.......
                )</span></td>
        </tr>
        <tr style="background-color:#1BF5FF">
<%
    If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
        Response.Write("<td align=""left"" colspan=""12"">")
    Else
        Response.Write("<td align=""left"" colspan=""11"">")
    End If
 %>
        法條列表
        <select name="PageCount">
            <option value="10" <%if trim(request("PageCount"))="10" then response.write("selected") %>>10</option>
            <option value="20" <%if trim(request("PageCount"))="20" then response.write("selected") %>>20</option>
            <option value="30" <%if trim(request("PageCount"))="30" then response.write("selected") %>>30</option>
            <option value="40" <%if trim(request("PageCount"))="40" then response.write("selected") %>>40</option>
            <option value="50" <%if trim(request("PageCount"))="50" then response.write("selected") %>>50</option>
            <option value="60" <%if trim(request("PageCount"))="60" then response.write("selected") %>>60</option>
            <option value="70" <%if trim(request("PageCount"))="70" then response.write("selected") %>>70</option>
            <option value="80" <%if trim(request("PageCount"))="80" then response.write("selected") %>>80</option>
            <option value="90" <%if trim(request("PageCount"))="90" then response.write("selected") %>>90</option>
            <option value="100" <%if trim(request("PageCount"))="100" then response.write("selected") %>>100</option>
        </select>
        </td>
        </tr>
        <tr style="background-color:#EBFBE3">
        <td align="center" style="width: 70px">法條代碼</td>
        <td align="center" style="width: 355px">法條名稱</td>
        <td align="center" style="width: 75px">車種</td>
        <td align="center" style="width: 40px">版本</td>
        <td align="center" style="width: 55px">配分標準</td>
        <td align="center" style="width: 65px">攔停配分</td>
        <td align="center" style="width: 65px">逕舉配分</td>
<%

    
    If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
        Response.Write("<td align=""center"" style=""width: 65px"">拖吊配分</td>")
    End If
 %>
        <td align="center" style="width: 65px">A1 配分</td>
        <td align="center" style="width: 65px">A2 配分</td>
        <td align="center" style="width: 65px">A3 配分</td>
        <td align="center" style="width: 65px">操作</td>
        </tr>
<%
    Response.Write(str1)
 %>
        <tr style="background-color:#1BF5FF">
<%
    If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
        Response.Write("<td align=""center"" colspan=""12"">")
    Else
        Response.Write("<td align=""center"" colspan=""11"">")
    End If
 %>
            <asp:Button ID="btBack" runat="server" Text="上一頁" OnClick="Page_Back" />&nbsp;
            <asp:Label ID="LbPageNum" runat="server" Height="22px" Width="30px" Visible="False">0</asp:Label>&nbsp;<asp:Label
                ID="LbPageNumTotal" runat="server" Height="22px" Width="115px"></asp:Label>
            <asp:Button ID="btNext" runat="server" Text="下一頁" OnClick="Page_Next" />
            <input name="btExcel" type="button" value="匯出Excel檔" onclick="OpenExcel();" /></td>
        </tr>
     </table>
    </form>
    <asp:literal id="Literal1" runat="server"></asp:literal>
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
    //限制法條輸入格式
    function checkValue(v){
	    return ((v.replace(/^[^\d]+|[^\d,]|,+$/g,'')).replace(/,+/g,',')).replace(/,+$/g,'');
    }
    //法條積分更新
    function ScoreUpdate(LawID,LawVer,CorN,CarSimple,TagOrder){
<%
    If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
%>
        tag1="form1.BillType1" + TagOrder;
        tag2="form1.BillType2" + TagOrder;
        tag3="form1.A1" + TagOrder;
        tag4="form1.A2" + TagOrder;
        tag5="form1.A3" + TagOrder;
        tag6="form1.Other1" + TagOrder;
        if (eval(tag1).value=="" || eval(tag2).value=="" || eval(tag3).value=="" || eval(tag4).value=="" || eval(tag5).value=="" || eval(tag6).value==""){
            alert("請將此法條所有欄位輸入!!");
        }else{
	        AjaxObj.Open("POST","LawScoreSet.aspx",true);
            AjaxObj.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
	        AjaxObj.send("LawID=" + LawID + "&LawVer=" + LawVer + "&CorN=" + CorN + "&CarSimple=" + CarSimple + "&Type1=" + eval(tag1).value + "&Type2=" + eval(tag2).value + "&A1=" + eval(tag3).value + "&A2=" + eval(tag4).value + "&A3=" + eval(tag5).value + "&Other1=" + eval(tag6).value);
	        AjaxObj.onreadystatechange=ServerUpdate;
	    }
<%
    Else
%>
        tag1="form1.BillType1" + TagOrder;
        tag2="form1.BillType2" + TagOrder;
        tag3="form1.A1" + TagOrder;
        tag4="form1.A2" + TagOrder;
        tag5="form1.A3" + TagOrder;
        if (eval(tag1).value=="" || eval(tag2).value=="" || eval(tag3).value=="" || eval(tag4).value=="" || eval(tag5).value==""){
            alert("請將此法條所有欄位輸入!!");
        }else{
	        AjaxObj.Open("POST","LawScoreSet.aspx",true);
            AjaxObj.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
	        AjaxObj.send("LawID=" + LawID + "&LawVer=" + LawVer + "&CorN=" + CorN + "&CarSimple=" + CarSimple + "&Type1=" + eval(tag1).value + "&Type2=" + eval(tag2).value + "&A1=" + eval(tag3).value + "&A2=" + eval(tag4).value + "&A3=" + eval(tag5).value);
	        AjaxObj.onreadystatechange=ServerUpdate;
	    }
<%
    End If
 %>

    }
    		
    function ServerUpdate()
    {
	    if (AjaxObj.readystate==4 || AjaxObj.readystate=='complete')
	    {
		    //document.getElementById('nameList').innerHTML =AjaxObj.responseText;
		    alert(AjaxObj.responsetext);
	    }
    }
    
    function OpenExcel(){
    <%
        if LawID.Text<>"" then
            response.write("sLawID=""" & LawID.Text & """;")
        else
            response.write("sLawID="""";")
        end if
        if LawVer.Text<>"" then
            response.write("sLawVer=""" & LawVer.Text & """;")
        else
            response.write("sLawVer="""";")
        end if
        if trim(request("sCountyOrNpa"))<>"" then
            response.write("sCountyOrNpa=""" & trim(request("sCountyOrNpa")) & """;")
        else
            response.write("sCountyOrNpa="""";")
        end if
    %>
        window.open("LawScore_Excel.aspx?sLawID="+sLawID+"&sLawVer="+sLawVer+"&sCountyOrNpa="+sCountyOrNpa,"LawScore_Excel1","width=800,height=600,left=0,top=0,scrollbars=yes,menubar=yes,resizable=yes,fullscreen=no,status=no,toolbar=no");
    }
</script>
</html>
