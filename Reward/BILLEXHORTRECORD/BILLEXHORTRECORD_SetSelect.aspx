<%@ Page Language="VB" AutoEventWireup="false" CodeFile="BILLEXHORTRECORD_SetSelect.aspx.vb" ResponseEncoding = "big5" Inherits="BILLEXHORTRECORD_SetSelect" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>交通違規勸導績效統計表</title>
</head>
<script language="javascript" type="text/javascript">

document.onkeydown = checkKey;

function checkKey(oEvent){
  var oEvent = (oEvent)? oEvent : event;
  var oTarget =(oEvent.target)? oEvent.target : oEvent.srcElement;
  if(oEvent.keyCode==13) 
    oEvent.keyCode = 9;
}

</script>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:ScriptManager ID="ScriptManager1" runat="server" EnableScriptGlobalization="true"
EnableScriptLocalization="true"/>
        &nbsp;</div>
        &nbsp;
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
        <table style="width: 696px; height: 40px" border="1">
            <tr>
                <td>
                    <asp:Panel ID="Panel3" runat="server" BackColor="Orange" Height="10px">
                        統計時間</asp:Panel>
                </td>
                <td colspan="2">
                    <asp:Panel ID="Panel4" runat="server" BackColor="Orange" Height="10px">
                        統計單位</asp:Panel>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:Label ID="Label1" runat="server" Text="違規起迄日期" Width="109px"></asp:Label><asp:TextBox ID="txtDate1" runat="server" Width="67px"></asp:TextBox>
                    <asp:Label ID="Label2" runat="server" Text="～"></asp:Label>
                    <asp:TextBox ID="txtDate2" runat="server" Width="72px"></asp:TextBox>
        <cc1:CalendarExtender ID="CalendarExtender1" runat="server" Format="yyyy/MM/dd" PopupButtonID="txtDate1"
            TargetControlID="txtDate1">
        </cc1:CalendarExtender>
        <cc1:CalendarExtender ID="CalendarExtender2" runat="server" Format="yyyy/MM/dd" PopupButtonID="txtDate2"
            TargetControlID="txtDate2">
        </cc1:CalendarExtender>
        </td>
                <td colspan="2">
                    <asp:CheckBox ID="cbxUnit" runat="server" Text="單位" AutoPostBack="True" /><br />
                    <asp:DropDownList ID="DDLUnit" runat="server" DataSourceID="DSUnit" DataTextField="UNITNAME"
                        DataValueField="UNITID" Enabled="False">
                    </asp:DropDownList><asp:SqlDataSource ID="DSUnit" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                        ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand="select UNITID,UNITTYPEID,UNITNAME,UNITLEVELID from unitinfo&#13;&#10;ORDER BY SHOWORDER">
                    </asp:SqlDataSource>
                </td>
            </tr>
        </table>
            </ContentTemplate>
        </asp:UpdatePanel>
        <table style="width: 570px; height: 24px">
            <tr>
                <td style="height: 48px" align="center">
                    <asp:Panel ID="Panel2" runat="server" BackColor="Orange" Height="10px">
                        <asp:Button ID="btnExcel" runat="server" Text="產出報表(輸出格式 Excel )" /></asp:Panel>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
