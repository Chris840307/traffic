<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CommonPercent_Set.aspx.vb" Inherits="ScoreList_CommonPercent_Set" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script language="JavaScript">
	window.focus();
</script>

<asp:literal runat="server" id="literal1"></asp:literal>
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>共同人員分配比率設定</title>
</head>
<body>
    <form id="form1" runat="server">
    <table border="1" align="center"> 
        <tr>
            <td colspan="3" style="background-color: #FFCC66">
                <strong>共同人員獎勵金分配比率設定</strong></td>
        </tr>
        <tr>
            <td colspan="3" style="height: 63px">
                分配類別：<asp:DropDownList ID="DropDownList1" runat="server" AutoPostBack="True">
                    <asp:ListItem Value="1">業務規劃機關之主官、副主官、承辦單位主管及相關人員  </asp:ListItem>
                    <asp:ListItem Value="2">勤務監督執行機關之主官、副主官承辦單位主管及相關人員</asp:ListItem>
                    <asp:ListItem Value="3">勤務執行機構之主官、副主官承辦單位主管及相關人員</asp:ListItem>
                    <asp:ListItem Value="4">負責交通安全工作之督導考核、資訊、後勤、人事、主計、秘書、出納等相關作業人員</asp:ListItem>
                </asp:DropDownList><br />
                分配人員/單位：
                <asp:TextBox ID="TbUnit" runat="server" Width="252px"></asp:TextBox><span style="font-size: 10pt">
                </span>&nbsp;
                分配比例：
                <asp:TextBox ID="TbPercent" runat="server" Width="43px"></asp:TextBox>
                % &nbsp;&nbsp;&nbsp;
                <br>
                單位：<asp:DropDownList ID="DropDownList2" runat="server" DataSourceID="SqlDataSource1"
                    DataTextField="UNITNAME" DataValueField="UNITID" AutoPostBack="True">
                </asp:DropDownList><asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:DB_Orcl %>"
                    ProviderName="<%$ ConnectionStrings:DB_Orcl.ProviderName %>" >
                </asp:SqlDataSource>
                &nbsp; &nbsp; &nbsp; &nbsp; 
                <asp:Button ID="Button1" runat="server" Font-Size="10pt" Height="28px" Text="新增分配單位"
                    Width="87px" />
                <asp:Button ID="Button2" runat="server" Font-Size="10pt" Height="28px" Text="離開"
                    Width="87px" OnClientClick="window.close();" />
            </td>
        </tr>

        <tr>
            <td colspan="2" style="height: 30px; background-color: #ccffcc;">
                業務規劃機關之主官、副主官、承辦單位主管及相關人員
            </td>
            <td>
                <input type="text" name="GroupPercent1" value="<%
    '取得 Web.config 檔的資料連接設定
    Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
    '建立 Connection 物件
    Dim conn As New Data.OracleClient.OracleConnection()
    conn.ConnectionString = setting.ConnectionString
    '開啟資料連接
    conn.Open()
    
    dim GroupPercentTmp1,GroupPercentTmp2,GroupPercentTmp3,GroupPercentTmp4 as integer
    dim strGroupPercent as string
    strGroupPercent="select CommonShareUnit,SharePercent from CommonShareReward where ShareGroupID=0"
    Dim CmdGroupPercent As New Data.OracleClient.OracleCommand(strGroupPercent, conn)
    Dim rdGroupPercent As Data.OracleClient.OracleDataReader = CmdGroupPercent.ExecuteReader()
    If rdGroupPercent.HasRows Then
        While rdGroupPercent.Read()
            if trim(rdGroupPercent("CommonShareUnit"))="1" then
                GroupPercentTmp1=rdGroupPercent("SharePercent")*100
            elseif trim(rdGroupPercent("CommonShareUnit"))="2" then
                GroupPercentTmp2=rdGroupPercent("SharePercent")*100
            elseif trim(rdGroupPercent("CommonShareUnit"))="3" then
                GroupPercentTmp3=rdGroupPercent("SharePercent")*100
            elseif trim(rdGroupPercent("CommonShareUnit"))="4" then
                GroupPercentTmp4=rdGroupPercent("SharePercent")*100
            end if
        End While
    End If
    rdGroupPercent.Close()
    
    response.write(GroupPercentTmp1)
                 %>" size="5" />
                ％
               <input type="button" name="buttonA" value="修改" onclick="UpdateGroupPercent(1);" <%=strDisable %>/>
            </td>
        </tr>
        <tr>
            <td style="width: 350px; background-color: #FFFF99">分配人員/單位</td>
            <td style="width: 210px; background-color: #FFFF99">分配比例%</td>
            <td style="width: 130px; background-color: #FFFF99">操作</td>
        </tr>
<%

    Dim ShareUnit1 As String
    Dim SharePercent1, ShareSN1 As Decimal
    Dim strUnit1 = "select * from CommonShareReward where ShareGroupID=1 order by sn"
    Dim CmdUnit1 As New Data.OracleClient.OracleCommand(strUnit1, conn)
    Dim rdUnit1 As Data.OracleClient.OracleDataReader = CmdUnit1.ExecuteReader()
    If rdUnit1.HasRows Then
        While rdUnit1.Read()
            Response.Write("<tr>")
            ShareUnit1 = rdUnit1("CommonShareUnit")
            SharePercent1 = rdUnit1("SharePercent") * 100
            ShareSN1 = rdUnit1("SN")
 
            
            Response.Write("<td>" & ShareUnit1 & rdUnit1("ChName") & "</td>")
            Response.Write("<td><input type=""text"" style=""width: 150px"" value=""" & SharePercent1 & """ name=""SharePercent" & ShareSN1 & """></td>")
            Response.Write("<td>")
            Response.Write("<input type=""button"" value=""修改"" onclick=""UpdatePercent(" & ShareSN1 & ");"" " & strDisable & "/>")
            Response.Write("<input type=""button"" value=""刪除"" onclick=""if(confirm('是否確定要刪除')){DeletePercent(" & ShareSN1 & ")}"" " & strDisable & "/>")
            Response.Write("</td>")
            
            Response.Write("</tr>")
        End While
    End If
    rdUnit1.Close()
    
    
%>
        <tr>
            <td colspan="2" style="height: 30px; background-color: #ccffcc;">
                勤務監督執行機關之主官、副主官承辦單位主管及相關人員
            </td>
            <td>
                <input type="text" name="GroupPercent2" value="<%=GroupPercentTmp2%>" size="5" />
                ％
               <input type="button" name="buttonB" value="修改" onclick="UpdateGroupPercent(2);" <%=strDisable %>/>
            </td>
        </tr>
        <tr>
            <td style="width: 350px; background-color: #FFFF99">分配人員/單位</td>
            <td style="width: 210px; background-color: #FFFF99">分配比例%</td>
            <td style="width: 130px; background-color: #FFFF99">操作</td>
        </tr>
<%

    Dim ShareUnit2 As String
    Dim SharePercent2, ShareSN2 As Decimal
    Dim strUnit2 = "select * from CommonShareReward where ShareGroupID=2 and UnitID='" & Trim(DropDownList2.SelectedValue) & "' order by sn"
    Dim CmdUnit2 As New Data.OracleClient.OracleCommand(strUnit2, conn)
    Dim rdUnit2 As Data.OracleClient.OracleDataReader = CmdUnit2.ExecuteReader()
    If rdUnit2.HasRows Then
        While rdUnit2.Read()
            Response.Write("<tr>")
            ShareUnit2 = rdUnit2("CommonShareUnit")
            SharePercent2 = rdUnit2("SharePercent") * 100
            ShareSN2 = rdUnit2("SN")
 
            
            Response.Write("<td>" & ShareUnit2 & rdUnit2("ChName") & "</td>")
            Response.Write("<td><input type=""text"" style=""width: 150px"" value=""" & SharePercent2 & """ name=""SharePercent" & ShareSN2 & """></td>")
            Response.Write("<td>")
            Response.Write("<input type=""button"" value=""修改"" onclick=""UpdatePercent(" & ShareSN2 & ");"" />")
            Response.Write("<input type=""button"" value=""刪除"" onclick=""if(confirm('是否確定要刪除')){DeletePercent(" & ShareSN2 & ")}"" />")
            Response.Write("</td>")
            
            Response.Write("</tr>")
        End While
    End If
    rdUnit2.Close()
    
    
%>
        <tr>
            <td colspan="2" style="height: 30px; background-color: #ccffcc;">
                勤務執行機構之主官、副主官承辦單位主管及相關人員
            </td>
            <td>
                <input type="text" name="GroupPercent3" value="<%=GroupPercentTmp3%>" size="5" />
                ％
               <input type="button" name="buttonC" value="修改" onclick="UpdateGroupPercent(3);" <%=strDisable %>/>
            </td>
        </tr>
        <tr>
            <td style="width: 350px; background-color: #FFFF99">分配人員/單位</td>
            <td style="width: 210px; background-color: #FFFF99">分配比例%</td>
            <td style="width: 130px; background-color: #FFFF99">操作</td>
        </tr>
<%

    Dim ShareUnit3 As String
    Dim SharePercent3, ShareSN3 As Decimal
    Dim strUnit3 = "select * from CommonShareReward where ShareGroupID=3 and UnitID='" & Trim(DropDownList2.SelectedValue) & "' order by sn"
    Dim CmdUnit3 As New Data.OracleClient.OracleCommand(strUnit3, conn)
    Dim rdUnit3 As Data.OracleClient.OracleDataReader = CmdUnit3.ExecuteReader()
    If rdUnit3.HasRows Then
        While rdUnit3.Read()
            Response.Write("<tr>")
            ShareUnit3 = rdUnit3("CommonShareUnit")
            SharePercent3 = rdUnit3("SharePercent") * 100
            ShareSN3 = rdUnit3("SN")
 
            
            Response.Write("<td>" & ShareUnit3 & rdUnit3("ChName") & "</td>")
            Response.Write("<td><input type=""text"" style=""width: 150px"" value=""" & SharePercent3 & """ name=""SharePercent" & ShareSN3 & """></td>")
            Response.Write("<td>")
            Response.Write("<input type=""button"" value=""修改"" onclick=""UpdatePercent(" & ShareSN3 & ");"" />")
            Response.Write("<input type=""button"" value=""刪除"" onclick=""if(confirm('是否確定要刪除')){DeletePercent(" & ShareSN3 & ")}"" />")
            Response.Write("</td>")
            
            Response.Write("</tr>")
        End While
    End If
    rdUnit3.Close()
    
    
%>
        <tr>
            <td colspan="2" style="height: 30px; background-color: #ccffcc;">
                負責交通安全工作之督導考核、資訊、後勤、人事、主計、秘書、出納等相關作業人員
            </td>
            <td>
                <input type="text" name="GroupPercent4" value="<%=GroupPercentTmp4%>" size="5" />
                ％
               <input type="button" name="buttonD" value="修改" onclick="UpdateGroupPercent(4);" <%=strDisable %>/>
            </td>
        </tr>
        <tr>
            <td style="width: 350px; background-color: #FFFF99">分配人員/單位</td>
            <td style="width: 210px; background-color: #FFFF99">分配比例%</td>
            <td style="width: 130px; background-color: #FFFF99">操作</td>
        </tr>        
<%

    
    Dim ShareUnit4 As String
    Dim SharePercent4, ShareSN4 As Decimal
    Dim strUnit4 = "select * from CommonShareReward where ShareGroupID=4 order by sn"
    Dim CmdUnit4 As New Data.OracleClient.OracleCommand(strUnit4, conn)
    Dim rdUnit4 As Data.OracleClient.OracleDataReader = CmdUnit4.ExecuteReader()
    If rdUnit4.HasRows Then
        While rdUnit4.Read()
            Response.Write("<tr>")
            ShareUnit4 = rdUnit4("CommonShareUnit")
            SharePercent4 = rdUnit4("SharePercent") * 100
            ShareSN4 = rdUnit4("SN")
 
            
            Response.Write("<td>" & ShareUnit4 & rdUnit4("ChName") & "</td>")
            Response.Write("<td><input type=""text"" style=""width: 150px"" value=""" & SharePercent4 & """ name=""SharePercent" & ShareSN4 & """></td>")
            Response.Write("<td>")
            Response.Write("<input type=""button"" value=""修改"" onclick=""UpdatePercent(" & ShareSN4 & ");"" " & strDisable & "/>")
            Response.Write("<input type=""button"" value=""刪除"" onclick=""if(confirm('是否確定要刪除')){DeletePercent(" & ShareSN4 & ")}"" " & strDisable & "/>")
            Response.Write("</td>")
            
            Response.Write("</tr>")
        End While
    End If
    rdUnit4.Close()
    
    
%>
<%

    conn.Close()
 %>       
    </table>
    </form>
</body>
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
    //群組配分更新
    function UpdateGroupPercent(GroupID){

        if (GroupID==1){
            GroupPercent=form1.GroupPercent1.value;
        }else if (GroupID==2){
            GroupPercent=form1.GroupPercent2.value;
        }else if (GroupID==3){
            GroupPercent=form1.GroupPercent3.value;
        }else{
            GroupPercent=form1.GroupPercent4.value;
        }
        AjaxObj.Open("POST","Update_CommonGroupPercent.aspx",true);
        AjaxObj.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        AjaxObj.send("SharePercent=" + GroupPercent + "&GroupID=" + GroupID);
        AjaxObj.onreadystatechange=ServerUpdate;
    }		
    
    //配分更新
    function UpdatePercent(SN){
        tag1="form1.SharePercent" + SN;

        if (eval(tag1).value==""){
            alert("請輸入分配比例!!");
        }else{
	        AjaxObj.Open("POST","Update_CommonPercent.aspx",true);
            AjaxObj.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
	        AjaxObj.send("SharePercent=" + eval(tag1).value + "&SN=" + SN + "&ActionType=1");
	        AjaxObj.onreadystatechange=ServerUpdate;
	    }
    }		
    function ServerUpdate()
    {
	    if (AjaxObj.readystate==4 || AjaxObj.readystate=='complete')
	    {
		    //document.getElementById('nameList').innerHTML =AjaxObj.responseText;
		    alert(AjaxObj.responsetext);
	    }
    }

    
    //群組刪除
    function DeletePercent(SN){

	    AjaxObj.Open("POST","Update_CommonPercent.aspx",true);
        AjaxObj.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded');
        AjaxObj.send("SN=" + SN + "&ActionType=2");
	    AjaxObj.onreadystatechange=ServerDelete;
    }		
    function ServerDelete()
    {
	    if (AjaxObj.readystate==4 || AjaxObj.readystate=='complete')
	    {
		    //document.getElementById('nameList').innerHTML =AjaxObj.responseText;
		    //alert(AjaxObj.responsetext);
		    form1.submit();
	    }
    }

</script>
</html>
