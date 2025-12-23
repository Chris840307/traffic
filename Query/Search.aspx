<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Search.aspx.cs" Inherits="Search" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>申請資料查詢</title>
    <link href="Css/Css.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        .auto-style1 {
            height: 28px;
        }

        .bar {
            background-color: white;
        }

            .bar li {
                height: 25px;
                float: left;
                background-color: white;
                display: block;
                white-space: nowrap;
                padding-left: 20px;
                padding-bottom: 20px;
            }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <asp:HiddenField ID="OutPutID" runat="server" Value="GetTACSearch" />
        <table class="title0" width="1024px" border="0">
            <tr>
                <td class="title0">申請資料查詢
                </td>
            </tr>
            <tr>
                <td style="background-color: white;">
                    <ul class="bar">
                        <li>線上申請日期：
                                    <asp:TextBox ID="tbx_CreDt1" CssClass="btn1" Width="100" MaxLength="7" onkeyup="chknumber(this);" runat="server" />
                            <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="Image/date.jpg"
                                OnClientClick="OpenWindow('tbx_CreDt1'); return false;" />
                            至　
                                    <asp:TextBox ID="tbx_CreDt2" CssClass="btn1" Width="100" MaxLength="7" onkeyup="chknumber(this);" runat="server" />
                            <asp:ImageButton ID="ImageButton4" runat="server" ImageUrl="Image/date.jpg"
                                OnClientClick="OpenWindow('tbx_CreDt2'); return false;" />
                        </li>
                        <li>當事人姓名：
                                    <asp:TextBox ID="tbx_ContactName" CssClass="btn1" ToolTip="可用(*)模糊查詢" Width="100" MaxLength="30" runat="server" />
                        </li>
                        <li>聯絡電話：
                                    <asp:TextBox ID="tbx_ContactTel" CssClass="btn1" Width="100" MaxLength="30" runat="server" />
                        </li>
                        <li>匯款帳戶：
                                    <asp:TextBox ID="tbx_BankAccount" CssClass="btn1" Width="100" MaxLength="30" runat="server" />
                        </li>
                        <li>匯款進度：
                                    <asp:DropDownList ID="ddl_UserData" runat="server">
                                        <asp:ListItem Value="" Text="請選擇" />
                                        <asp:ListItem Value="1" Text="已匯款" />
                                        <asp:ListItem Value="2" Text="未匯款" />
                                    </asp:DropDownList>
                        </li>
                        <li>
                            <asp:Button ID="btn_Excel" CssClass="btn3" runat="server" Text="匯出Excel" OnClick="btn_Excel_Click" />
                            <asp:Button ID="btn_Search" CssClass="btn3" runat="server" Text="查詢" OnClick="btn_Search_Click" />
                        </li>
                    </ul>
                </td>
            </tr>
            <tr>
                <td>
                    <asp:GridView ID="gdv_Search" runat="server" CssClass="title2" DataSourceID="DbSearch" AutoGenerateColumns="False" Width="1024px"
                        DataKeyNames="ApplyNo" RowStyle-BackColor="#EEEEFF" AllowPaging="True" AllowSorting="True"
                        OnDataBound="gdv_Search_DataBound"
                        OnRowCommand="gdv_Search_RowCommand" OnRowDataBound="gdv_Search_RowDataBound">
                        <PagerSettings Mode="NumericFirstLast" />
                        <RowStyle Wrap="false" />
                        <Columns>
                            <asp:BoundField DataField="ApplyNo" HeaderText="申請編號" SortExpression="ApplyNo" />
                            <asp:BoundField DataField="ContactName" HeaderText="申請人<br/>姓名" HtmlEncode="false" SortExpression="ContactName" />
                            <asp:BoundField DataField="ContactAddress" HeaderText="申請人地址" HtmlEncode="false" SortExpression="ContactAddress" />
                            <asp:BoundField DataField="ContactTel" HeaderText="申請人電話" HtmlEncode="false" SortExpression="ContactTel" />
                            <asp:BoundField DataField="ContactRelations" HeaderText="與當事人關係" HtmlEncode="false" SortExpression="ContactRelations" />
                            <asp:BoundField DataField="CaseFlag" HeaderText="傷亡情形" SortExpression="CaseFlag" />
                            <asp:CheckBoxField DataField="IsJusticiability" HeaderText="司法審理中" SortExpression="IsJusticiability" />
                            <asp:BoundField DataField="BankAccount" HeaderText="匯款帳戶" HtmlEncode="false" SortExpression="BankAccount" />
                            <asp:BoundField DataField="UserData" HeaderText="匯款狀態" SortExpression="UserData" />
                            <asp:CheckBoxField DataField="Adopt" HeaderText="是否<br/>通過" SortExpression="Adopt" />
                            <asp:TemplateField HeaderText="申請日期" SortExpression="CreDt">
                                <ItemTemplate>
                                    <asp:Label ID="Label4" runat="server" Text='<%# DateTime.Parse( Eval("CreDt","")).ToStrTaiwanCalendar("yyy/MM/dd") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="事故分析<br/>研判表" SortExpression="ContactDoc1">
                                <ItemTemplate>
                                    <asp:Button ID="Button1" runat="server" Text="下載" CommandName="ContactDoc1" CommandArgument='<%# Bind("ContactDoc1") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="身分證<br/>正反面影本1" SortExpression="ContactDoc2">
                                <ItemTemplate>
                                    <asp:Button ID="Button2" runat="server" Text="下載" CommandName="ContactDoc2" CommandArgument='<%# Bind("ContactDoc2") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField HeaderText="身分證<br/>正反面影本2" SortExpression="ContactDoc3">
                                <ItemTemplate>
                                    <asp:Button ID="Button3" runat="server" Text="下載" CommandName="ContactDoc3" CommandArgument='<%# Bind("ContactDoc3") %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                            <asp:TemplateField ShowHeader="false">
                                <ItemTemplate>
                                    <!--<%# Session["SearchRow"+ ((GridViewRow) Container).RowIndex] = GetDataItem() %>-->
                                    <asp:Button ID="Download" runat="server" Text="下載鑑定申請表" CommandName="Download" />
                                    <br />
                                    <asp:Button ID="Adopt" runat="server" Text="通過" CommandName="Adopt" Visible='<%# !"true".Equals(Eval("Adopt",""),StringComparison.CurrentCultureIgnoreCase) %>' />
                                    <asp:Button ID="NotAdopt" runat="server" Text="不通過" CommandName="NotAdopt" Visible='<%# "true".Equals(Eval("Adopt",""),StringComparison.CurrentCultureIgnoreCase) %>' />
                                </ItemTemplate>
                            </asp:TemplateField>
                        </Columns>
                        <RowStyle BackColor="#EEEEFF"></RowStyle>
                        <PagerTemplate>
                            <table>
                                <tr>
                                    <td></td>
                                    <td>
                                        <asp:Button ID="btn_Pre" runat="server" CssClass="btn3" Text="上一頁" CommandName="Prev" />
                                        <asp:Label ID="lbl_PageCnt" runat="server" Text="共幾頁" Font-Size="Small" />
                                        <asp:Button ID="btn_Next" runat="server" CssClass="btn3" Text="下一頁" CommandName="Next" />
                                        <asp:Button ID="btn_Goto" CssClass="btn3" runat="server" Text="跳至" CommandName="Goto" />
                                        <asp:TextBox ID="tbx_Goto" CssClass="btn1" Width="40" runat="server" />
                                        <asp:Button ID="btn" CssClass="btn3" runat="server" Text="每頁筆數" CommandName="PageSize" />
                                        <asp:TextBox ID="tbx_PageSize" CssClass="btn1" Width="40" runat="server" />
                                    </td>
                                    <td></td>
                                </tr>
                            </table>
                        </PagerTemplate>
                    </asp:GridView>
                    <asp:ObjectDataSource ID="DbSearch" runat="server" SelectMethod="GetTACSearch" TypeName="TcpWebUsing" />
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
<script src="Scripts/data.js" type="text/javascript"></script>
