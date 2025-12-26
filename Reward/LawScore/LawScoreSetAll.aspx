<%@ Page Language="VB" AutoEventWireup="false" CodeFile="LawScoreSetAll.aspx.vb" Inherits="LawScore_LawScoreSetAll" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>法條配分統一設定</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <table border="1" width="500" align="center" style="margin-top: 100px">
            <tr style="background-color:#FFCC33">
                <td align="left" colspan="2" style="height: 35px">
                    <span style="font-size: 14pt"><strong>
                    法條配分設定</strong></span>
                </td>
            </tr>
            <tr>
                <td style="width: 160px; height: 35px; background-color:#FFFFCC" align="center">配分標準</td>
                <td style="height: 35px; width: 328px;">
                            <select name="sCountyOrNpa1">
                                <option value="0" <%if trim(request("sCountyOrNpa1"))="0" then response.write("selected")%>>獎勵金</option>
                                <option value="1" <%if trim(request("sCountyOrNpa1"))="1" then response.write("selected")%>>績效</option>
                                <option value="n" <%if trim(request("sCountyOrNpa1"))="n" then response.write("selected")%>>全部</option>
                            </select></td>
             </tr>
             <tr>
                <td style="width: 160px; height: 50px; background-color:#FFFFCC" align="center">法條類別</td>
                <td style="height: 50px; width: 328px;">
                    <asp:Panel ID="Panel1" runat="server" Height="80px" Width="125px" BorderStyle="Inset">
                        <input type="checkbox" name="SecoreType" value="1" <%
                        if InStr(Trim(Request("SecoreType")), "1") <> 0 Then
                            response.write("checked")
                        end if
                        %> />攔停<br />
                        <input type="checkbox" name="SecoreType" value="2" <%
                        if InStr(Trim(Request("SecoreType")), "2") <> 0 Then
                            response.write("checked")
                        end if
                        %> />逕舉<br />
<%
    If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Then
%>
                        <input type="checkbox" name="SecoreType" value="6" <%
                        if InStr(Trim(Request("SecoreType")), "6") <> 0 Then
                            response.write("checked")
                        end if
                        %> />拖吊<br />
<%        
    End If
%>               
                        <input type="checkbox" name="SecoreType" value="3" <%
                        if InStr(Trim(Request("SecoreType")), "3") <> 0 Then
                            response.write("checked")
                        end if
                        %> />A1<br />
                        <input type="checkbox" name="SecoreType" value="4" <%
                        if InStr(Trim(Request("SecoreType")), "4") <> 0 Then
                            response.write("checked")
                        end if
                        %> />A2<br />
                        <input type="checkbox" name="SecoreType" value="5" <%
                        if InStr(Trim(Request("SecoreType")), "5") <> 0 Then
                            response.write("checked")
                        end if
                        %> />A3
                    </asp:Panel>
                 </td>
             </tr>
             <tr>
                <td align="center" style="width: 160px; height: 56px; background-color: #ffffcc">
                    法條範圍</td>
                <td style="width: 328px; height: 56px">
                    <input name="LawRange" type="radio" value="1" <%if trim(request("LawRange"))="1" or trim(request("LawRange"))="" then response.write("checked")%>/>全部<br />
                    <input name="LawRange" type="radio" value="2" <%if trim(request("LawRange"))="2" then response.write("checked")%>/>12~68條<br />
                    <input name="LawRange" type="radio" value="3" <%if trim(request("LawRange"))="3" then response.write("checked")%>/>69條以後</td>
            </tr>
            <tr>
                <td style="width: 160px; height: 35px; background-color:#FFFFCC" align="center">配分</td>
                <td style="height: 35px; width: 328px;">
                            <asp:TextBox ID="tScore" runat="server" Width="35px"></asp:TextBox>分</td>
             </tr>
            <tr style="background-color:#FFCC33">
                <td align="center" colspan="2" style="height: 35px">
                    <asp:Button ID="BtUpdate" runat="server" Text="確定" OnClick="LawScoreUpdate" Font-Size="12pt" Height="30px" Width="55px" />
                    <asp:Button ID="BtClose" runat="server" Text="離開" PostBackUrl="~/LawScore/LawScore.aspx" Font-Size="12pt" Height="30px" Width="55px" />
                </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
