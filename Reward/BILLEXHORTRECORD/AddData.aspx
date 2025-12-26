<%@ Page Language="VB" AutoEventWireup="false" CodeFile="AddData.aspx.vb" Inherits="AddData" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>交通違規勸導單</title>
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


<body bgcolor="#ffffcc" onload="document.form1.txtNo.focus()"> 
    <form id="form1" runat="server">
    <div>
        <div>
            <asp:ScriptManager ID="ScriptManager1" runat="server">
            </asp:ScriptManager>
            <table border="1" style="width: 678px">
                <tr>
                    <td align="center" colspan="4" style="height: 22px">
                        &nbsp;<asp:Label ID="Label2" runat="server" Font-Size="Large" Text="交通違規勸導單"></asp:Label>&nbsp;
                        (日期格式：95 07 07 &nbsp;時間格式：23 00(24小時制))<br />
                        編號：<asp:TextBox ID="txtNo" runat="server" MaxLength="20"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator14" runat="server" ControlToValidate="txtNo"
                            Display="Dynamic" ErrorMessage="編號未輸入"></asp:RequiredFieldValidator>
                        &nbsp; &nbsp;
                    </td>
                </tr>
                <tr style="font-size: 12pt">
                    <td colspan="2" style="width: 783px; height: 28px">
                        &nbsp;<asp:RadioButton ID="RBUserName" runat="server" Checked="True" GroupName="UserName"
                            Text="駕駛人" />
                        <asp:RadioButton ID="RBUserName2" runat="server" GroupName="UserName" Text="行為人" />
                        <asp:TextBox ID="txtUserName" runat="server" MaxLength="30" Width="146px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ControlToValidate="txtUserName"
                            Display="Dynamic" ErrorMessage="駕駛人或行為人未輸入"></asp:RequiredFieldValidator></td>
                    <td colspan="2" style="width: 296px; color: #000000; height: 28px">
                        &nbsp;身份證統一編號：<asp:TextBox ID="txtID" runat="server" MaxLength="20" Width="119px"></asp:TextBox><asp:RequiredFieldValidator
                            ID="RequiredFieldValidator2" runat="server" ControlToValidate="txtID" Display="Dynamic"
                            ErrorMessage="身份證號碼未輸入"></asp:RequiredFieldValidator></td>
                </tr>
                <tr style="font-size: 12pt; color: #000000">
                    <td colspan="2" style="width: 783px; height: 34px">
                        出生日期：<asp:TextBox ID="txtBirthYear" runat="server" MaxLength="3" Width="47px"></asp:TextBox>年<asp:TextBox
                            ID="txtBirthMonth" runat="server" MaxLength="2" Width="47px"></asp:TextBox>月<asp:TextBox
                                ID="txtBirthDay" runat="server" MaxLength="2" Width="51px"></asp:TextBox>日<asp:RangeValidator
                                    ID="RangeValidator1" runat="server" ControlToValidate="txtBirthYear" Display="Dynamic"
                                    ErrorMessage="年輸入錯誤" MaximumValue="200" MinimumValue="0" Type="Integer"></asp:RangeValidator>
                        <asp:RangeValidator ID="RangeValidator2" runat="server" ControlToValidate="txtBirthMonth"
                            Display="Dynamic" ErrorMessage="月輸入錯誤" MaximumValue="12" MinimumValue="1" Type="Integer"></asp:RangeValidator>
                        <asp:RangeValidator ID="RangeValidator3" runat="server" ControlToValidate="txtBirthDay"
                            Display="Dynamic" ErrorMessage="日輸入錯誤" MaximumValue="31" MinimumValue="1" Type="Integer"></asp:RangeValidator>&nbsp;
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator7" runat="server" ControlToValidate="txtBirthMonth"
                            Display="Dynamic" ErrorMessage="月未輸入"></asp:RequiredFieldValidator>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator8" runat="server" ControlToValidate="txtBirthDay"
                            Display="Dynamic" ErrorMessage="日未輸入"></asp:RequiredFieldValidator>&nbsp;
                        <asp:RegularExpressionValidator ID="RegularExpressionValidator2" runat="server" ControlToValidate="txtBirthMonth"
                            Display="Dynamic" ErrorMessage="請參考日期格式" ValidationExpression="\d{2}"></asp:RegularExpressionValidator>
                        <asp:RegularExpressionValidator ID="RegularExpressionValidator3" runat="server" ControlToValidate="txtBirthDay"
                            Display="Dynamic" ErrorMessage="請參考日期格式" ValidationExpression="\d{2}"></asp:RegularExpressionValidator></td>
                    <td colspan="2" style="width: 296px; color: #000000; height: 34px">
                        &nbsp;車牌號碼：<asp:TextBox ID="txtCar_No" runat="server" MaxLength="10" Width="168px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ControlToValidate="txtCar_No"
                            Display="Dynamic" ErrorMessage="車牌號碼未輸入"></asp:RequiredFieldValidator></td>
                </tr>
                <tr style="font-size: 12pt; color: #000000">
                    <td colspan="4" style="height: 28px">
                        &nbsp;<asp:RadioButton ID="RBUserAddress" runat="server" Checked="True" GroupName="UserAddress"
                            Text="駕駛人" /><asp:RadioButton ID="RBUserAddress2" runat="server" GroupName="UserAddress"
                                Text="行為人" />
                        &nbsp; &nbsp; &nbsp;地址：<asp:TextBox ID="txtAddress" runat="server" MaxLength="100"
                            Width="410px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator4" runat="server" ControlToValidate="txtAddress"
                            Display="Dynamic" ErrorMessage="地址未輸入"></asp:RequiredFieldValidator></td>
                </tr>
                <tr style="font-size: 12pt">
                    <td colspan="2" style="width: 783px; height: 6px">
                        違規時間：<asp:TextBox ID="txtIllYear" runat="server" MaxLength="3" Width="42px"></asp:TextBox>年<asp:TextBox
                            ID="txtIllMonth" runat="server" MaxLength="2" Width="26px"></asp:TextBox>月<asp:TextBox
                                ID="txtIllDay" runat="server" MaxLength="2" Width="34px"></asp:TextBox>日<asp:TextBox
                                    ID="txtIllHour" runat="server" MaxLength="2" Width="32px"></asp:TextBox>時<asp:TextBox
                                        ID="txtIllMin" runat="server" MaxLength="2" Width="40px"></asp:TextBox>分<asp:RangeValidator
                                            ID="RangeValidator4" runat="server" ControlToValidate="txtIllYear" Display="Dynamic"
                                            ErrorMessage="年輸入錯誤" MaximumValue="200" MinimumValue="0" Type="Integer"></asp:RangeValidator>
                        <asp:RangeValidator ID="RangeValidator5" runat="server" ControlToValidate="txtIllMonth"
                            Display="Dynamic" ErrorMessage="月輸入錯誤" MaximumValue="12" MinimumValue="1" Type="Integer"></asp:RangeValidator>
                        <asp:RangeValidator ID="RangeValidator6" runat="server" ControlToValidate="txtIllDay"
                            Display="Dynamic" ErrorMessage="日輸入錯誤" MaximumValue="31" MinimumValue="1" Type="Integer"></asp:RangeValidator>
                        <asp:RangeValidator ID="RangeValidator7" runat="server" ControlToValidate="txtIllHour"
                            Display="Dynamic" ErrorMessage="時輸入錯誤" MaximumValue="23" MinimumValue="0" Type="Integer"></asp:RangeValidator>
                        <asp:RangeValidator ID="RangeValidator8" runat="server" ControlToValidate="txtIllHour"
                            Display="Dynamic" ErrorMessage="分輸入錯誤" MaximumValue="59" MinimumValue="0" Type="Integer"></asp:RangeValidator>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator9" runat="server" ControlToValidate="txtIllYear"
                            Display="Dynamic" ErrorMessage="年未輸入"></asp:RequiredFieldValidator>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator10" runat="server" ControlToValidate="txtIllMonth"
                            Display="Dynamic" ErrorMessage="月未輸入"></asp:RequiredFieldValidator>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator11" runat="server" ControlToValidate="txtIllDay"
                            Display="Dynamic" ErrorMessage="日未輸入"></asp:RequiredFieldValidator>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator12" runat="server" ControlToValidate="txtIllHour"
                            Display="Dynamic" ErrorMessage="時未輸入"></asp:RequiredFieldValidator>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator13" runat="server" ControlToValidate="txtIllDay"
                            Display="Dynamic" ErrorMessage="分未輸入"></asp:RequiredFieldValidator>
                        <asp:RegularExpressionValidator ID="RegularExpressionValidator4" runat="server" ControlToValidate="txtIllDay"
                            Display="Dynamic" ErrorMessage="請參考日期格式" ValidationExpression="\d{2}"></asp:RegularExpressionValidator>
                        <asp:RegularExpressionValidator ID="RegularExpressionValidator5" runat="server" ControlToValidate="txtIllHour"
                            Display="Dynamic" ErrorMessage="請參考日期格式" ValidationExpression="\d{2}"></asp:RegularExpressionValidator>
                        <asp:RegularExpressionValidator ID="RegularExpressionValidator6" runat="server" ControlToValidate="txtIllMin"
                            Display="Dynamic" ErrorMessage="請參考日期格式" ValidationExpression="\d{2}"></asp:RegularExpressionValidator>
                        <asp:RegularExpressionValidator ID="RegularExpressionValidator7" runat="server" ControlToValidate="txtIllMonth"
                            Display="Dynamic" ErrorMessage="請參考日期格式" ValidationExpression="\d{2}"></asp:RegularExpressionValidator></td>
                    <td colspan="2" style="width: 296px; height: 6px">
                        &nbsp;違規地點：<asp:TextBox ID="txtIllAddress" runat="server" MaxLength="100" Width="165px"></asp:TextBox>
                        <asp:RequiredFieldValidator ID="RequiredFieldValidator5" runat="server" ControlToValidate="txtIllAddress"
                            Display="Dynamic" ErrorMessage="違規地點未輸入"></asp:RequiredFieldValidator></td>
                </tr>
                <tr style="font-size: 12pt">
                    <td align="center" colspan="4" style="height: 299px">
                        &nbsp; &nbsp;&nbsp;<br />
                        <table border="0" style="width: 614px; height: 115px">
                            <tr>
                                <td align="left" style="width: 293px; height: 9px">
                                    違規事實(請勾選)：</td>
                                <td align="left" colspan="2" style="width: 311px; height: 9px">
                                </td>
                            </tr>
                            <tr>
                                <td align="left" style="width: 293px; height: 9px">
                                    <asp:CheckBox ID="cbxIllTrue1" runat="server" Text="1.未帶駕照(經查證領有駕照)。" /></td>
                                <td align="left" colspan="2" style="width: 311px; height: 9px">
                                    <asp:CheckBox ID="cbxIllTrue8" runat="server" Text="8.機車附載人員或物品未依規定者。" /></td>
                            </tr>
                            <tr>
                                <td align="left" style="width: 293px; height: 9px">
                                    <asp:CheckBox ID="cbxIllTrue2" runat="server" Text="2.未帶行駕(經查證領有行照)。" /></td>
                                <td align="left" colspan="2" style="width: 311px; height: 9px">
                                    <asp:CheckBox ID="cbxIllTrue9" runat="server" Text="9.號誌燈變換，車前輪未進入停止線。" /></td>
                            </tr>
                            <tr>
                                <td align="left" style="width: 293px; height: 9px">
                                    <asp:CheckBox ID="cbxIllTrue3" runat="server" Text="3.號牌污穢(責令當場改正)。" /></td>
                                <td align="left" colspan="2" style="width: 311px; height: 9px">
                                    <asp:CheckBox ID="cbxIllTrue10" runat="server" Text="10.號誌燈變換，車前輪未進入機車停" /></td>
                            </tr>
                            <tr>
                                <td align="left" style="width: 293px">
                                    <asp:CheckBox ID="cbxIllTrue4" runat="server" Text="4.亂鳴喇叭(當場勸戒)。" /></td>
                                <td align="left" colspan="2" style="width: 311px">
                                    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                                    <asp:Label ID="Label1" runat="server" Text="等區。"></asp:Label></td>
                            </tr>
                            <tr>
                                <td align="left" style="width: 293px">
                                    <asp:CheckBox ID="cbxIllTrue5" runat="server" Text="5.超載10%以下" /></td>
                                <td align="left" colspan="2" style="width: 311px">
                                    <asp:CheckBox ID="cbxIllTrue11" runat="server" Text="11.在道路堆積、置放、設置或拋擲足" /></td>
                            </tr>
                            <tr>
                                <td align="left" style="width: 293px; height: 4px">
                                    <asp:CheckBox ID="cbxIllTrue6" runat="server" Text="6.酒測值逾0.02毫克以下。" /></td>
                                <td align="left" colspan="2" style="width: 311px; height: 4px">
                                    &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
                                    <asp:Label ID="Label3" runat="server" Text="以妨礙交通之物。"></asp:Label></td>
                            </tr>
                            <tr>
                                <td align="left" style="width: 293px">
                                    <asp:CheckBox ID="cbxIllTrue7" runat="server" Text="7.大型車右轉未先駛入外側車道。" /></td>
                                <td align="left" colspan="2" style="width: 311px">
                                </td>
                            </tr>
                            <tr>
                                <td align="left" colspan="3" style="height: 16px">
                                    臨時停車及其他違規（詳載違規事實）勸導理由：</td>
                            </tr>
                            <tr>
                                <td align="left" colspan="3">
                                    <asp:TextBox ID="txtOTHERILLEGALITEM" runat="server" Height="52px" MaxLength="100"
                                        TextMode="MultiLine" Width="601px"></asp:TextBox></td>
                            </tr>
                        </table>
                        &nbsp; &nbsp;&nbsp;</td>
                </tr>
                <tr style="font-size: 12pt">
                    <td colspan="2" style="width: 783px; height: 39px">
                        &nbsp;&nbsp; 勸導單位<asp:DropDownList ID="DDLUnit" runat="server" AutoPostBack="True"
                            DataSourceID="DSUnit" DataTextField="UNITNAME" DataValueField="UNITID" OnSelectedIndexChanged="DDLUnit_SelectedIndexChanged">
                        </asp:DropDownList><asp:SqlDataSource ID="DSUnit" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand='SELECT "UNITID", "UNITNAME", "UNITTYPEID", "UNITLEVELID" FROM "UNITINFO"&#13;&#10;'>
                        </asp:SqlDataSource>
                    </td>
                    <td colspan="2" rowspan="1" style="width: 296px">
                        填單人職名章<asp:DropDownList ID="DDLFillMemID" runat="server" DataSourceID="DSFillMemID"
                            DataTextField="CHNAME" DataValueField="MEMBERID">
                        </asp:DropDownList>
                        <asp:TextBox ID="txtMemID" runat="server" AutoPostBack="True" MaxLength="4" OnTextChanged="txtMemID_TextChanged"
                            Width="52px"></asp:TextBox>
                        <asp:SqlDataSource ID="DSFillMemID" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString %>"
                            ProviderName="<%$ ConnectionStrings:ConnectionString.ProviderName %>" SelectCommand='SELECT "MEMBERID", "CHNAME" FROM "MEMBERDATA"&#13;&#10;where UnitID=:UnitID and recordstateid=0 order by chname'>
                            <SelectParameters>
                                <asp:ControlParameter ControlID="DDLUnit" Name="UnitID" PropertyName="SelectedValue" />
                            </SelectParameters>
                        </asp:SqlDataSource>
                        &nbsp;
                        </td>
                </tr>
            </table>
            &nbsp; &nbsp;
        </div>
        <center>
            &nbsp;<asp:Button ID="btnSave" runat="server" Text="儲存" Width="118px" />
            &nbsp; &nbsp;&nbsp;
            <asp:SqlDataSource ID="SqlDataSource1" runat="server"></asp:SqlDataSource>
        </center>
    
    </div>
    </form>
</body>
</html>
