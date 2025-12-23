<%@ Page Language="VB" AutoEventWireup="false" CodeFile="CheckDocSignData.aspx.vb" Inherits="CheckDocSignData" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>刑案承辦人簽收</title>
    <style type="text/css">
        .style2 {
            background-color: #FFFF99;
        }

        #Button2 {
            height: 28px;
            width: 54px;
        }
        .auto-style1 {
            height: 50px;
        }
        .auto-style2 {
            width: 99px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <table border="1" width="100%">
            <tr>
                <td colspan="6"
                    style="background-color: #FFCC66; font-family: 新細明體; font-size: large; font-weight: bold" class="auto-style1">刑案承辦人簽收
                    <asp:Label ID="Label1" runat="server" Text="編號  " Font-Bold="True" Font-Size="Medium"></asp:Label>
                    <asp:Label ID="lblProject_ID" runat="server" Font-Size="Medium" ForeColor="#CC3300"></asp:Label>
                        &nbsp;&nbsp;&nbsp;
                    <asp:Label ID="Label2" runat="server" Text="案由  " Font-Bold="True" Font-Size="Medium"></asp:Label>
                    <asp:Label ID="lblProject_Name" runat="server" Font-Size="Medium" ForeColor="#CC3300"></asp:Label>
                </td>
            <tr>
                <td valign="top" >
                    <table>
                        <td style="background-color: #FFFFCC;" class="auto-style2">
                    <asp:Label ID="Label8" runat="server" Text="刑案承辦人" Font-Size="Medium" ></asp:Label>
                            </td>
                        <td>
                    &nbsp;<asp:TextBox ID="txtSendPoliceName" runat="server" Font-Size="Medium" ></asp:TextBox>
                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                            <asp:Button ID="btnEdit0" runat="server" Text="簽名" Font-Size="Medium" Height="30px" Width="80px" />
                            </td>
                            <td  rowspan="6">
                    <asp:Image ID="Image1" runat="server" Height="169px" Width="265px" />
                                </td>
                    <tr>
                        <td style="background-color: #FFFFCC;" class="auto-style2">
                    <asp:Label ID="Label9" runat="server" Text="簽收日期" Font-Size="Medium" ></asp:Label>
                            </td>
                        <td>
                    &nbsp;<asp:TextBox ID="txtSendPoliceSignDate" runat="server" Font-Size="Medium"  MaxLength="7" onkeyup="chknumber(this);"></asp:TextBox>
                    <asp:ImageButton ID="ImageButton2" runat="server" ImageUrl="../Image/date.jpg" 
                                 OnClientClick="OpenWindow('txtSendPoliceSignDate'); return false;" />
                     <asp:CustomValidator ID="SendPoliceSignDate" runat="server" 
                    ControlToValidate="txtSendPoliceSignDate" EnableClientScript="False"  ForeColor="Red"
                    ErrorMessage="＊日期格式錯誤" SetFocusOnError="True"></asp:CustomValidator>
                    <asp:TextBox ID="txtSendPoliceSignFileName" runat="server" Font-Size="Medium" Visible="False" Width="18px" ></asp:TextBox>
                    </td>
                        <tr>
                            <td style="background-color: #FFFFCC;" class="auto-style2">
                    <asp:Label ID="Label10" runat="server" Text="簽收時間" Font-Size="Medium" ></asp:Label>
                                </td>
                            <td>
                    &nbsp;<asp:TextBox ID="txtSendPoliceSignTime" runat="server" Font-Size="Medium" MaxLength="4" ></asp:TextBox>
                    
                     <asp:CustomValidator ID="SendPoliceSignTime" runat="server" 
                    ControlToValidate="txtSendPoliceSignTime" EnableClientScript="False"  ForeColor="Red"
                    ErrorMessage="＊時間格式錯誤" SetFocusOnError="True"></asp:CustomValidator>
                    </td>
                            </tr>
                        <td style="background-color: #FFFFCC;" class="auto-style2">
                    <asp:Label ID="Label11" runat="server" Text="備註" Font-Size="Medium" ></asp:Label>
                            </td>
                        <td>
                    &nbsp;<asp:TextBox ID="txtReMark" runat="server" Font-Size="Medium" Height="98px" Rows="5" TextMode="MultiLine" Width="528px"></asp:TextBox>
                    </td>
                            </table>
                </td>
                <tr>
            <td colspan="6"
                style="background-color: #FFCC66; text-align: center;">
                    <asp:Button ID="btnSave" runat="server" Text="儲存" Font-Size="Medium" />
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <asp:Button ID="btnClose" runat="server" Text="關閉" Font-Size="Medium" />
            </td>
        </tr>
        </table>
    </form>
    <script src="../Scripts/data.js" type="text/javascript"></script>    
</body>


</html>
