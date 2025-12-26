<%@ Page Language="VB" AutoEventWireup="true" EnableEventValidation = "false" CodeFile="BILLEXHORTRECORD.aspx.vb"  Inherits="_Default" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>勸導單管理</title>
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


<body onload="document.form1.txtNo.focus()">
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server"  EnableScriptGlobalization="true"
EnableScriptLocalization="true"/>
        <table style="width: 937px; height: 389px">
            <tr>
                <td colspan="3" style="width: 735px; height: 1px">
                    <asp:Panel ID="Panel3" runat="server" BackColor="#FFCC33" Height="10px" Width="872px" Font-Bold="True">
                        勸導單管理</asp:Panel>
                    <asp:Panel ID="Panel2" runat="server" BackColor="#FFFFC0" BorderColor="#000040" BorderWidth="0px"
                        Height="50px" HorizontalAlign="Left" Width="872px">
                        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                            <ContentTemplate>
                        <table style="width: 398px">
                            <tr>
                                <td style="height: 55px; width: 948px;">
                        <table style="width: 789px">
                            <tr>
                                <td style="width: 66px; height: 24px">
                                    <asp:Label ID="Label1" runat="server" Text="編號"></asp:Label></td>
                                <td colspan="2" style="height: 24px">
                                    <asp:TextBox ID="txtNo" runat="server" Width="213px" AutoPostBack="True"></asp:TextBox></td>
                                <td style="width: 93px; height: 24px">
                                    <asp:Label ID="Label2" runat="server" Text="身份證字號" Width="84px"></asp:Label></td>
                                <td style="width: 152px; height: 24px">
                                    <asp:TextBox ID="txtID" runat="server" Width="96px" AutoPostBack="True"></asp:TextBox></td>
                                <td style="width: 87px; height: 24px">
                                    <asp:Label ID="Label3" runat="server" Text="車號號碼"></asp:Label></td>
                                <td style="height: 24px" colspan="2">
                                    <asp:TextBox ID="txtCarNo" runat="server" Width="93px" AutoPostBack="True"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td style="width: 66px; height: 5px;">
                                    <asp:Label ID="Label5" runat="server" Text="違規日"></asp:Label></td>
                                <td style="width: 116px; height: 5px;">
                                    <asp:TextBox ID="txtIllDate" runat="server" Width="93px" AutoPostBack="True"></asp:TextBox>～</td>
                                <td style="height: 5px;">
                                    <asp:TextBox ID="txtIllDate2" runat="server" Width="93px" AutoPostBack="True"></asp:TextBox></td>
                                <td style="width: 93px; height: 5px;" valign="middle">
                                    <asp:Label ID="Label6" runat="server" Text="勸導單位"></asp:Label></td>
                                <td style="width: 152px; height: 5px;">
                                    <asp:DropDownList ID="DDLUnit" runat="server" AppendDataBoundItems="True" AutoPostBack="True"
                                        DataSourceID="DSUnit" DataTextField="UNITNAME" DataValueField="UNITID">
                                        <asp:ListItem Selected="True">所有單位</asp:ListItem>
                                    </asp:DropDownList></td>
                                <td style="width: 87px; height: 5px;">
                                    <asp:Label ID="Label4" runat="server" Text="填單人"></asp:Label></td>
                                <td colspan="2" style="height: 5px">
                                    <asp:DropDownList ID="DDLFillMemID" runat="server" AppendDataBoundItems="True" DataSourceID="DSMemID"
                                        DataTextField="CHNAME" DataValueField="MEMBERID" EnableViewState="False">
                                        <asp:ListItem>所有人員</asp:ListItem>
                                    </asp:DropDownList></td>
                            </tr>
                        </table>
                    <asp:SqlDataSource ID="DSMemID" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                        ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand="SELECT MEMBERID, CHNAME FROM MEMBERDATA WHERE (UNITID = :UnitID)  and recordstateid=0 order by chname">
                        <SelectParameters>
                            <asp:ControlParameter ControlID="DDLUnit" Name="UnitID" PropertyName="SelectedValue" />
                        </SelectParameters>
                    </asp:SqlDataSource>
                    <asp:SqlDataSource ID="DSUnit" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                        ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand='SELECT "UNITID", "UNITTYPEID", "UNITNAME", "UNITLEVELID" FROM "UNITINFO"&#13;&#10;'>
                    </asp:SqlDataSource>
                    <cc1:calendarextender id="CalendarExtender1" runat="server" format="yyyy/MM/dd" popupbuttonid="txtIllDate"
                        targetcontrolid="txtIllDate"></cc1:calendarextender>
                    <cc1:calendarextender id="Calendarextender2" runat="server" format="yyyy/MM/dd" popupbuttonid="txtIllDate2"
                        targetcontrolid="txtIllDate2">
                        </cc1:CalendarExtender>
                                </td>
                                <td style="height: 55px; width: 547px;" valign="middle">
                                </td>
                            </tr>
                        </table>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                        <table style="width: 222px">
                            <tr>
                                <td colspan="1" style="width: 468px; height: 26px">
                                </td>
                                <td colspan="1" style="width: 57px; height: 26px;">
                                    <asp:Button ID="btnClear" runat="server" Text="清除" OnClick="btnClear_Click" /></td>
                                <td colspan="3" style="width: 57px; height: 26px;">
                                    <asp:Button ID="btnQry" runat="server" Text="查詢" /></td>
                                <td colspan="1" style="width: 176px; height: 26px">
                                    <asp:Button ID="Button2" runat="server" Text="新增" /></td>
                            </tr>
                            <tr>
                            </tr>
                        </table>
                    </asp:Panel>
                </td>
            </tr>
            <tr>
                <td colspan="3" style="width: 735px; height: 13px">
                    <asp:SqlDataSource ID="SQLDSOrcl" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                        DeleteCommand="Update BILLEXHORTRECORD Set RECORDSTATEID = 1 where NO= :NO "
                        ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand='SELECT * FROM "BILLEXHORTRECORD" where RECORDSTATEID =0 order by no'
                        UpdateCommand="Update BILLEXHORTRECORD Set  ILLEGALADDRESS= :ILLEGALADDRESS , USERNAME= :USERNAME , USERID= :USERID , CARNO= :CARNO  where NO= :NO ">
                        <DeleteParameters>
                            <asp:ControlParameter ControlID="GridView1" Name="NO" PropertyName="SelectedValue" />
                        </DeleteParameters>
                        <UpdateParameters>
                            <asp:ControlParameter ControlID="GridView1" Name="ILLEGALADDRESS" PropertyName="SelectedValue" />
                            <asp:ControlParameter ControlID="GridView1" Name="USERNAME" PropertyName="SelectedValue" />
                            <asp:ControlParameter ControlID="GridView1" Name="USERID" PropertyName="SelectedValue" />
                            <asp:ControlParameter ControlID="GridView1" Name="CARNO" PropertyName="SelectedValue" />
                            <asp:ControlParameter ControlID="GridView1" Name="NO" PropertyName="SelectedValue" />
                        </UpdateParameters>
                    </asp:SqlDataSource>
                <asp:Panel ID="Panel1" runat="server" BackColor="#FFCC33" BorderColor="#000040" BorderWidth="0px"
                        Height="10px" Width="872px">
                        <b>勸導單記錄列表</b> &nbsp; &nbsp;每頁
                        <asp:DropDownList ID="DropDownList1" runat="server" AppendDataBoundItems="True">
                            <asp:ListItem>10</asp:ListItem>
                            <asp:ListItem>20</asp:ListItem>
                            <asp:ListItem>30</asp:ListItem>
                            <asp:ListItem>40</asp:ListItem>
                            <asp:ListItem>50</asp:ListItem>
                            <asp:ListItem>60</asp:ListItem>
                            <asp:ListItem>70</asp:ListItem>
                            <asp:ListItem>80</asp:ListItem>
                            <asp:ListItem>90</asp:ListItem>
                            <asp:ListItem>100</asp:ListItem>
                        </asp:DropDownList>筆 &nbsp; &nbsp; 第<asp:Label ID="Label8" runat="server"></asp:Label>/<asp:Label
                            ID="Label9" runat="server"></asp:Label>頁，共<asp:Label ID="Label7" runat="server"></asp:Label>筆
                        &nbsp;&nbsp;&nbsp;</asp:Panel>
                    <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
                        BackColor="#FFFFC0" DataKeyNames="NO" DataSourceID="SQLDSOrcl"
                        Height="249px" Width="872px">
                        <PagerTemplate>
                            <br />
                        </PagerTemplate>
                        <Columns>
                            <asp:BoundField DataField="NO" HeaderText="編號" SortExpression="NO">
                                <HeaderStyle BackColor="#EBFBE3" />
                                <ItemStyle Width="100px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ILLEGALDATETIME" HeaderText="違規時間" ReadOnly="True" SortExpression="ILLEGALDATETIME">
                                <HeaderStyle BackColor="#EBFBE3" />
                                <ItemStyle Width="180px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="ILLEGALADDRESS" HeaderText="違規地點" SortExpression="ILLEGALADDRESS">
                                <HeaderStyle BackColor="#EBFBE3" />
                            </asp:BoundField>
                            <asp:BoundField DataField="USERNAME" HeaderText="姓名" SortExpression="USERNAME">
                                <HeaderStyle BackColor="#EBFBE3" />
                                <ItemStyle Width="100px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="USERID" HeaderText="身份證字號" SortExpression="USERID">
                                <HeaderStyle BackColor="#EBFBE3" />
                                <ItemStyle Width="100px" />
                            </asp:BoundField>
                            <asp:BoundField DataField="CARNO" HeaderText="車牌號碼" SortExpression="CARNO">
                                <HeaderStyle BackColor="#EBFBE3" />
                                <ItemStyle Width="100px" />
                            </asp:BoundField>
                            
                            <asp:TemplateField HeaderText="功能按鈕" ShowHeader="False">
                                <EditItemTemplate>
                                    <asp:Button ID="Button1" runat="server" CausesValidation="True" CommandName="Update"
                                        Text="更新" />
                                    <asp:Button ID="Button2" runat="server" CausesValidation="False" CommandName="Cancel"
                                        Text="取消" />
                                </EditItemTemplate>
                                <ItemTemplate>
                                    <asp:Button ID="btnEdit" runat="server" CausesValidation="False" CommandArgument='<%# Eval("NO") %>'
                                        CommandName="EditData" OnClick="btnEdit_Click" Text="修改" />
                                    <asp:Button ID="btnDel" runat="server" CommandName="Delete"
                                        OnClientClick='javascript:return confirm("是否刪除？")' Text="刪除" OnClick="btnDel_Click" />
                                </ItemTemplate>
                                
                                <ItemStyle HorizontalAlign="Center" Width="100px" />
                                <HeaderStyle BackColor="#EBFBE3" />
                                <FooterStyle BackColor="#FFE0C0" />
                            </asp:TemplateField>
                            <asp:BoundField DataField="RECORDMEMBERID" DataFormatString="{0:c}">
                                <ControlStyle Width="0px" />
                                <HeaderStyle BackColor="#C0FFC0" />
                            </asp:BoundField>
                        </Columns>
                        <EmptyDataTemplate>
                            目前無資料
                        </EmptyDataTemplate>
                        <FooterStyle BackColor="White" />
                    </asp:GridView>
                </td>
            </tr>
            <tr>
                <td colspan="3" style="width: 735px; height: 8px">
                    <asp:Panel ID="Panel4" runat="server" BackColor="#FFDD77" BorderColor="White" BorderWidth="1px"
                                Height="30px" Width="872px" HorizontalAlign="Center">
                                <asp:Button ID="btnFirst" runat="server" Text="第一頁" />
                                <asp:Button ID="btnPre" runat="server" Text="上一頁" />
                                <asp:Button ID="btnNext" runat="server" Text="下一頁" />
                                <asp:Button ID="btnLast" runat="server" Text="最末頁" />
                                <asp:Button ID="btnExcel" runat="server" Text="轉出Excel" Height="24px" /></asp:Panel>
                </td>
            </tr>
        </table>
        
        <br />
        <div>
        </div>
    </form>
</body>
</html>
