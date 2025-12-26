<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Login.aspx.vb" Inherits="_Default" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.OracleClient" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    <title>績效獎勵金試算系統</title>
</head>
<body>
    <form id="form1" runat="server">
    <div id="Layer1" style="position:absolute ; width:266px; height:160px; z-index:0; border: 1px none #000000; left: 356px; top: 270px;">
        <table width='300' border='0' cellpadding="1">
            <tr>
                <td align="center" style="width: 40% ;height: 30px">身分證帳號</td>
                <td align="left" style="width: 60%">
                    <input type="text" name="UserID" onkeyup="value=value.toUpperCase()" style="width: 127px; font-size: 14pt; height: 25px;" maxlength="10"/>
                </td>
            </tr>
            <tr>
                <td align="center" style="height: 30px">使用者密碼</td>
                <td align="left">
                    <input type="password" name="UserPW" style="width: 127px; font-size: 14pt; height: 25px;" />
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" style="height: 30px">
                    <asp:Button ID="Button1" runat="server" Text="登入" Font-Size="Medium" Height="31px" Width="54px" onclick="User_Check"/>
                </td>
            </tr>
            <tr>
                <td align="center" colspan="2" style="height: 30px">
                    <asp:Label ID="ErrorMsg" runat="server" ForeColor="Red" Height="22px" Width="277px">
                    <%
                        If Trim(Request("ErrMsg")) = "1" Then
                            Response.Write("請先登入本系統")
                        End If
                    %></asp:Label></td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
